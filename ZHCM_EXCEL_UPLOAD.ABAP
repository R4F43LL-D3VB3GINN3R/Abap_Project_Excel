REPORT zhcm_excel_upload.

DATA: gt_file_data TYPE TABLE OF string, " Tabela de strings para armazenar dados lidos
      ls_data      TYPE string.

DATA: v_path      TYPE string,
      lt_bin_data TYPE TABLE OF x255, " Tabela binária para dados do arquivo
      lv_xstr     TYPE xstring.       " Variável para armazenar o conteúdo em XSTRING

TYPES: BEGIN OF ty_timesheet,
         num       TYPE string,
         nome      TYPE string,
         equipa    TYPE string,
         cntr_cust TYPE string,
         dia1      TYPE string,
       END OF ty_timesheet.

DATA: it_timesheet TYPE TABLE OF ty_timesheet,
      ls_timesheet TYPE ty_timesheet.

DATA: count_columns TYPE i.
count_columns = 5.

"-----------------------------------------------------------------------------

SELECTION-SCREEN: SKIP 1.
SELECTION-SCREEN: BEGIN OF BLOCK a1 WITH FRAME TITLE TEXT-001.
PARAMETERS: p_path TYPE ibipparms-path.
SELECTION-SCREEN: END OF BLOCK a1.
SELECTION-SCREEN: SKIP 1.

"-----------------------------------------------------------------------------

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_path. " Insere o ícone de procura
  " Chama a função para abrir a caixa de seleção do arquivo.
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

  v_path = p_path. " Define o caminho do arquivo

  " Carrega o arquivo em uma tabela binária
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

  " Converte a tabela binária para XSTRING
  CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
    EXPORTING
      input_length = lines( lt_bin_data ) * 255   " Tamanho total dos dados
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
  DATA(lo_excel)  = lo_reader->zif_excel_reader~load( i_excel2007 = lv_xstr ).  " Passa o XSTRING carregado

  DATA(i) = 2.

  WHILE i <= lo_excel->get_worksheets_size( ).

    DATA(lo_worksheet) = lo_excel->get_worksheet_by_index( i ). "controlar o index até ter toda a informacao - tentar rodar pela sheet duas ou mais vezes.

    "itera sobre cada célula da planilha
    LOOP AT lo_worksheet->sheet_content REFERENCE INTO DATA(cell).

      "-----------------------------------------------------
      "           cabeçalho de colaboradores
      "-----------------------------------------------------

      IF cell->cell_row EQ '1' AND cell->cell_column EQ '2'.
        ls_timesheet-num = cell->cell_value.
      ENDIF.

      IF cell->cell_row EQ '2' AND cell->cell_column EQ '2'.
        ls_timesheet-nome = cell->cell_value.
      ENDIF.

      IF cell->cell_row EQ '3' AND cell->cell_column EQ '2'.
        ls_timesheet-equipa = cell->cell_value.
      ENDIF.

      IF cell->cell_row EQ '4' AND cell->cell_column EQ '2'.
        IF cell->cell_value IS INITIAL.
          cell->cell_value = '---------'.
          ls_timesheet-cntr_cust = cell->cell_value.
        ELSE.
          ls_timesheet-cntr_cust = cell->cell_value.
        ENDIF.
      ENDIF.

      "-----------------------------------------------------
      "                calendario - peps
      "-----------------------------------------------------
      "inserir pep + quantidade de tempo trabalhado
      
      "TENHO DE VERIFICAR SE O PEP FOR INICIAL E CORRER A LINHA PARA BUSCAR O TEMPO TRABALHADO
      "
      IF cell->cell_row EQ '6' AND cell->cell_column EQ count_columns.
        ls_timesheet-dia1 = cell->cell_value.
        APPEND ls_timesheet TO it_timesheet.
        ADD 1 TO count_columns.
      ENDIF.

    ENDLOOP.

    ADD 1 TO i.
    count_columns = 5.

  ENDWHILE.

  cl_demo_output=>display( it_timesheet ).

ENDFORM.
