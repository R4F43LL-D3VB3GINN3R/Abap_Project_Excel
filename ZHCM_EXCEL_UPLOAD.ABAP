REPORT zhcm_excel_upload.

"-----------------------------------------------------------------------------

DATA: gt_file_data TYPE TABLE OF string, "tabela de strings para armazenar dados lidos
      ls_data      TYPE string.

DATA: v_path      TYPE string,
      lt_bin_data TYPE TABLE OF x255, "tabela binária para dados do arquivo
      lv_xstr     TYPE xstring.       "variável para armazenar o conteúdo em XSTRING

"-----------------------------------------------------------------------------

"tabela interna de exibicao de dados da timesheet
TYPES: BEGIN OF ty_timesheet,
         num       TYPE string,
         nome      TYPE string,
         equipa    TYPE string,
         cntr_cust TYPE string,
         dia       TYPE string,
         pep       TYPE string,
         horas     TYPE string,
       END OF ty_timesheet.

DATA: it_timesheet TYPE TABLE OF ty_timesheet,
      ls_timesheet TYPE ty_timesheet.

"-----------------------------------------------------------------------------

"index da sheet
DATA: lv_index TYPE i.
lv_index = 2.

"flag da sheet
DATA: flag_next_sheet TYPE flag.
flag_next_sheet = abap_false.

"index do dia trabalhado
DATA: index_day TYPE i.
index_day = 11.

"index da hora trabalhada no dia
DATA: index_hour TYPE i.
index_hour = 139.

"coordenada da celula
DATA: lv_coord     TYPE string,
      lv_coord_num TYPE i,
      lv_str_coord TYPE string.
lv_coord = 'A'.
lv_coord_num = 10.

"-----------------------------------------------------------------------------

SELECTION-SCREEN: SKIP 1.
SELECTION-SCREEN: BEGIN OF BLOCK a1 WITH FRAME TITLE TEXT-001.
PARAMETERS: p_path TYPE ibipparms-path.
SELECTION-SCREEN: END OF BLOCK a1.
SELECTION-SCREEN: SKIP 1.

"-----------------------------------------------------------------------------

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_path. "insere o ícone de procura
  "chama a função para abrir a caixa de seleção do arquivo.
  CALL FUNCTION 'F4_FILENAME'
    EXPORTING
      program_name  = syst-cprog
      dynpro_number = syst-dynnr
      field_name    = 'P_PATH'
    IMPORTING
      file_name     = p_path.

  "-----------------------------------------------------------------------------

START-OF-SELECTION.
  PERFORM get_file.
  PERFORM display_table.

END-OF-SELECTION.

FORM get_file.

  v_path = p_path. "define o caminho do arquivo

  "carrega o arquivo em uma tabela binária
  cl_gui_frontend_services=>gui_upload(
    EXPORTING
      filename             = v_path         " Nome do arquivo
      filetype             = 'BIN'          " Tipo de arquivo como binário
    CHANGING
      data_tab             = lt_bin_data    " Tabela binária para dados
    EXCEPTIONS
      file_open_error      = 1
      file_read_error      = 2
      no_batch             = 3
      gui_refuse_filetransfer = 4
      invalid_type         = 5
      no_authority         = 6
      unknown_error        = 7
      bad_data_format      = 8
      header_not_allowed   = 9
      separator_not_allowed = 10
      header_too_long      = 11
      unknown_dp_error     = 12
      access_denied        = 13
      dp_out_of_memory     = 14
      disk_full            = 15
      dp_timeout           = 16
      not_supported_by_gui = 17
      error_no_gui         = 18
      OTHERS               = 19
    ).

  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
      WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    RETURN.
  ENDIF.

  "converte a tabela binária para XSTRING
  CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
    EXPORTING
      input_length = lines( lt_bin_data ) * 255   " tamanho total dos dados
    IMPORTING
      buffer       = lv_xstr
    TABLES
      binary_tab   = lt_bin_data
    EXCEPTIONS
      failed       = 1
      OTHERS       = 2.

ENDFORM.

FORM display_table.

  CLEAR: gt_file_data.

  DATA(lo_reader) = NEW zcl_excel_reader_2007( ).
  DATA(lo_excel)  = lo_reader->zif_excel_reader~load( i_excel2007 = lv_xstr ).  "passa o XSTRING carregado

  DATA(i) = 2.

  "itera por todas as sheets do excel, seja ela quantas houverem
  WHILE i <= lo_excel->get_worksheets_size( ).

    "começa a partir da segunda sheet, sendo a primeira a exibicao de dados gerais
    DATA(lo_worksheet) = lo_excel->get_worksheet_by_index( i ).

    "tratamos cada sheet individualmente e só passamos para a próxima quando a flag for ativa
    WHILE flag_next_sheet EQ abap_false.

      "-----------------------------------------------------
      "           cabeçalho de colaboradores
      "-----------------------------------------------------

      READ TABLE lo_worksheet->sheet_content REFERENCE INTO DATA(cell) INDEX lv_index. "B2

      "numero do colaborador
      ls_timesheet-num = cell->cell_value.

      ADD 2 TO lv_index.

      READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_index. "B3

      "nome do colaborador
      ls_timesheet-nome = cell->cell_value.

      ADD 2 TO lv_index.

      READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_index. "B4

      "equipa do colaborador
      ls_timesheet-equipa = cell->cell_value.

      ADD 2 TO lv_index.

      READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_index. "B5

      "centro de custo do colaborador
      IF cell->cell_value IS INITIAL.
        cell->cell_value = '---------'.
        ls_timesheet-cntr_cust = cell->cell_value.
      ELSE.
        ls_timesheet-cntr_cust = cell->cell_value.
      ENDIF.

      "-----------------------------------------------------
      "                calendario - peps
      "-----------------------------------------------------

      "define a coordenada da celula
      lv_str_coord = lv_coord_num.
      CONCATENATE lv_coord lv_str_coord INTO lv_str_coord. "A10 ou "B10
      CONDENSE lv_str_coord NO-GAPS.

      CLEAR cell->cell_value. "limpa a celula

      "itera pelos 6 projetos --> A10 - A11 - A12 - A13 - A14 - A15 ou B10 - B11 - B12 - B13 - B14 - B15
      DO 6 TIMES.

        "le a celula de acordo com a coordenada
        READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell WITH KEY cell_coords = lv_str_coord. "A10

        "verifica se o colaborador trabalhou em algum projeto
        IF cell->cell_value NE 'Selecione'.
          ls_timesheet-pep = cell->cell_value.
        "se nao trabalhou, procuramos algum motivo de ausencia
        ELSE.
          lv_coord = 'B'.
          CLEAR lv_str_coord.
          lv_str_coord = lv_coord_num.
          CONCATENATE lv_coord lv_str_coord INTO lv_str_coord. "A10 ou "B10
          CONDENSE lv_str_coord NO-GAPS.

          "le a celula de acordo com a coordenada
          READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell WITH KEY cell_coords = lv_str_coord. "B10

          IF cell->cell_value NE 'Selecione'.
            ls_timesheet-pep = cell->cell_value.
          ENDIF.
        ENDIF.

        "se houve algum projeto ou motivo de ausencia
        IF ls_timesheet-PEP is not initial.

          "itera pelo calendário
          DO 30 TIMES.

            "verifica quantas horas foram trabalhadas
            READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX index_hour. "E10-AI10"..."E11-AI11..E12-AI12"..."E13-AI13..E14-AI14"..."E15-AI15..

            "se houver hora trabalhada no dia.
            IF cell->cell_value NE '0' AND cell->cell_value IS NOT INITIAL.
              ls_timesheet-horas = cell->cell_value.

              "pega o dia relacionado a hora de trabalho.
              READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX index_day. "E6-AI6"

              ls_timesheet-dia = cell->cell_value.
              APPEND ls_timesheet TO it_timesheet. "insere a linha na tabela interna

            ENDIF.

            ADD 1 TO index_hour. "proxima hora
            ADD 1 TO index_day.  "proximo dia

          ENDDO.

          "redefine o index de hora e dia
          index_hour = 139. "E10
          index_day = 11.   "E11

        ENDIF.

        "redefine a coordenada para o proximo projeto
        lv_coord = 'A'.
        ADD 1 TO lv_coord_num.
        lv_str_coord = lv_coord_num.
        CONCATENATE lv_coord lv_str_coord INTO lv_str_coord.
        CONDENSE lv_str_coord NO-GAPS.

        "redefine index de horarios conforme coordenadas
        CASE lv_str_coord.
          WHEN 'A10'.
            index_hour = 139. "E10
          WHEN 'A11'.
            index_hour = 171. "E11
          WHEN 'A12'.
            index_hour = 203. "E12
          WHEN 'A13'.
            index_hour = 235. "E13
          WHEN 'A14'.
            index_hour = 267. "E14
          WHEN 'A15'.
            index_hour = 299. "E15
        ENDCASE.

      ENDDO.

    "redefine o index de hora e dia
    index_hour = 139.
    index_day = 11.

    CLEAR ls_timesheet.

    "----------------------------------------------------------------------

    flag_next_sheet = abap_true.         "pode passar para a proxima sheet
    lv_index = 2.                        "reinicia o index
    CLEAR: lv_str_coord.                 "limpa a coordenada

  ENDWHILE.

  "passa para a próxima sheet
  ADD 1 TO i.
  flag_next_sheet = abap_false.

  lv_coord_num = 10. "redefine a linha da coordenada

ENDWHILE.

cl_demo_output=>display( it_timesheet ). "display da tabela interna

ENDFORM.
