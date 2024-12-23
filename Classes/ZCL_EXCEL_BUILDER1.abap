CLASS zcl_excel_builder2 DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    
    TYPES:
      "types de informacoes de colaboradores
      BEGIN OF wa_col,
        pernr TYPE pa0001-pernr, "Número Pessoal
        sname TYPE pa0002-cname, "Nome
        vdsk1 TYPE pa0001-vdsk1, "Chave de Organizacao
        kostl TYPE pa0001-kostl, "Centro de Custo
      END OF wa_col .
    TYPES:
      "types de ausencias de precensas
      BEGIN OF wa_pre_aus,
        subty TYPE awart,  "t554s-subty Tipos de presença e ausência
        atext TYPE abwtxt, "t554t-atext Textos de ausência e presença
      END OF wa_pre_aus .
    TYPES:
      "linha unica para guardar ausencia e presença
      BEGIN OF wa_line_preaus,
        line TYPE string,
      END OF wa_line_preaus .
    TYPES:
      "types para projetos abertos
      BEGIN OF wa_project,
        objnr TYPE j_objnr,  "Nº objeto
        pspid TYPE ps_pspid, "Definição do projeto
        post1 TYPE ps_post1, "PS: descrição breve (1ª linha de texto)
      END OF wa_project .
    TYPES:
      "linha unica para guardar projetos
      BEGIN OF wa_line_projects,
        line TYPE string,
      END OF wa_line_projects .

    "data do mes
    DATA: gv_datemonth TYPE sy-datum.

    "work schedule do colaborador
    DATA:
      tb_psp TYPE STANDARD TABLE OF ptpsp,
      wa_psp TYPE ptpsp.

    DATA:
      "informacoes dos colaboradores
      it_colaboradores TYPE TABLE OF wa_col .
    DATA ls_colaborador TYPE wa_col .
    DATA tt_colaboradores TYPE zcol_tt .
    DATA st_colaborador TYPE zcol_st .

    DATA:
      "informacoes de ausencia e presenca
      it_aus_pre TYPE TABLE OF wa_pre_aus .
    DATA ls_aus_pre TYPE wa_pre_aus .
    DATA:
      "linha de ausencia e presenca concatenada
      it_line_preaus TYPE TABLE OF wa_line_preaus .
    DATA ls_line_preaus TYPE wa_line_preaus .

    "celula de horas trabalhadas e planeadas
    DATA total_planeadas TYPE string .
    DATA total_trabalhadas TYPE string .

    DATA:
      "tabela e estrutura de projetos abertos
      it_projetos TYPE TABLE OF wa_project .
    DATA ls_projetos TYPE wa_project .
    DATA:
      "tabela de linha concatenada de projetos
      it_linha_projetos TYPE TABLE OF wa_line_projects .
    DATA ls_linha_projeto TYPE wa_line_projects .

    DATA e_result TYPE zrla_result .

    "objetos de construcao de arquivos excel
    DATA o_xl TYPE REF TO zcl_excel .
    DATA lo_worksheet TYPE REF TO zcl_excel_worksheet .
    "objetos de componentes do excel
    DATA lo_column TYPE REF TO zcl_excel_column .
    DATA lo_data_validation TYPE REF TO zcl_excel_data_validation .
    DATA lo_data_validation2 TYPE REF TO zcl_excel_data_validation .
    DATA lo_range TYPE REF TO zcl_excel_range .
    DATA o_converter TYPE REF TO zcl_excel_converter .
    DATA lo_style TYPE REF TO zcl_excel_style .
    DATA o_border_dark TYPE REF TO zcl_excel_style_border .
    DATA o_border_light TYPE REF TO zcl_excel_style_border .
    DATA tp_style_bold_center_guid TYPE zexcel_cell_style .
    DATA tp_style_bold_center_guid2 TYPE zexcel_cell_style .
    DATA ol_hyperlink TYPE REF TO zcl_excel_hyperlink .

    "tabela binária para dados do arquivo
    DATA: lt_bin_data TYPE TABLE OF x255,
          lv_xstr     TYPE xstring. "variável para armazenar o conteúdo em XSTRING

    "dados gerais da timesheet em excel file
    TYPES: BEGIN OF ty_timesheet,
             num       TYPE pa0001-pernr,
             nome      TYPE pa0002-cname,
             equipa    TYPE pa0001-vdsk1,
             cntr_cust TYPE pa0001-kostl,
             dia       TYPE sy-datum,
             pep       TYPE char100,
             auspres   TYPE char100,
             hora      TYPE catshours,
             row       TYPE string,
             info      TYPE string,
           END OF ty_timesheet.

    DATA: it_timesheet TYPE TABLE OF ty_timesheet,
          ls_timesheet TYPE ty_timesheet.

    "dados dos colaboradores em excel file
    TYPES: BEGIN OF ty_employee,
             num       TYPE string,
             nome      TYPE string,
             equipa    TYPE string,
             cntr_cust TYPE string,
           END OF ty_employee.

    DATA: it_employee TYPE TABLE OF ty_employee,
          ls_employee TYPE ty_employee.

    "dados dos peps em excel file
    TYPES: BEGIN OF ty_peps,
             num  TYPE string,
             dia  TYPE string,
             pep  TYPE string,
             hora TYPE catshours,
             row  TYPE string,
           END OF ty_peps.

    DATA: it_peps TYPE TABLE OF ty_peps,
          ls_peps TYPE ty_peps.

    "dados de ausencia e presenca em excel file
    TYPES: BEGIN OF ty_auspres,
             num     TYPE string,
             dia     TYPE string,
             auspres TYPE string,
             hora    TYPE string,
             row     TYPE string,
           END OF ty_auspres.

    DATA: it_auspres TYPE TABLE OF ty_auspres,
          ls_auspres TYPE ty_auspres.

    DATA: tt_alv TYPE ztshralv_tt, "tabela para alv
          st_alv TYPE ztshralv_st. "linha para alv

    METHODS get_data
      IMPORTING
        !colaboradores TYPE zcol_tt.
    METHODS download_xls
      EXPORTING
        result TYPE zrla_result.
    METHODS display_fast_excel
      IMPORTING
        !i_table_content TYPE REF TO data
        !i_table_name    TYPE string .
    METHODS get_date
      IMPORTING date TYPE sy-datum.
    METHODS upload_timesheet
      IMPORTING
        str_path_file TYPE string.
    METHODS get_timesheet_datafile
      EXPORTING
        table_timesheet TYPE ztshralv_tt
        result          TYPE zrla_result.
    METHODS insert_bapicats
      IMPORTING
        table_timesheet TYPE ztshralv_tt
      EXPORTING
        result          TYPE zrla_result.
  PROTECTED SECTION.
  PRIVATE SECTION.

    METHODS convert_xstring .
    METHODS set_database .
    METHODS append_extension
      IMPORTING
        !old_extension TYPE string
      EXPORTING
        !new_extension TYPE string .
    METHODS get_file_directory
      IMPORTING
        !filename  TYPE string
      EXPORTING
        !full_path TYPE string .
    METHODS set_style .
    METHODS set_sheets .
    METHODS generate_calendar .
    METHODS convert_excel_column
      IMPORTING
        column_int    TYPE i
      EXPORTING
        column_string TYPE string.
    METHODS get_auspres.
    METHODS get_projects.
    METHODS get_work_schedule
      IMPORTING pernr TYPE p_pernr.
    METHODS set_rangemonthdate
      EXPORTING
        begin_date TYPE sy-datum
        end_date   TYPE sy-datum.
    METHODS get_employee_datafile.
    METHODS get_month_datafile.
    METHODS get_peps_datafile.
    METHODS get_auspres_datafile.
    METHODS set_workschedule_datafile.
    METHODS get_wronglines_datafile.
    METHODS validation_datafile
      IMPORTING table_timesheet        TYPE ztshralv_tt
      EXPORTING table_timesheet_output TYPE ztshralv_tt.
    METHODS set_coordenates
      IMPORTING
        letter_coord   TYPE string
      EXPORTING
        string_coord_a TYPE string
        string_coord_b TYPE string
        index_coord    TYPE i.
    METHODS switch_coordenates
      IMPORTING
        coordenate     TYPE string
      EXPORTING
        string_coord_a TYPE string
        string_coord_b TYPE string
        index_coord    TYPE i.
ENDCLASS.



CLASS ZCL_EXCEL_BUILDER2 IMPLEMENTATION.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->APPEND_EXTENSION
* +-------------------------------------------------------------------------------------------------+
* | [--->] OLD_EXTENSION                  TYPE        STRING
* | [<---] NEW_EXTENSION                  TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
    METHOD append_extension.

      "-------------------------------------------------------
      "info: concatena a extensao do arquivo ao path principal
      "
      "data de alteracao: 09.11.2024
      "alteracao: criacao do método
      "criado por: rafael albuquerque
      "-------------------------------------------------------

      CONCATENATE old_extension 'xlsx' INTO new_extension SEPARATED BY '.'.

    ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->CONVERT_EXCEL_COLUMN
* +-------------------------------------------------------------------------------------------------+
* | [--->] COLUMN_INT                     TYPE        I
* | [<---] COLUMN_STRING                  TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
    METHOD convert_excel_column.

      "----------------------------------------------------------------------------------------------
      "info: através de um numero, retorna o valor em string da referida coluna em excel
      "este metodo é usado para complementar a formula da soma das horas no metodo generate calendar
      "assim a formula é atualizada a cada iteracao
      "
      "data de alteracao: 09.11.2024
      "alteracao: criacao do método
      "criado por: rafael albuquerque
      "----------------------------------------------------------------------------------------------

      "variavel recebe o parametro de entrada
      DATA: lv_column_int TYPE i.
      lv_column_int = column_int.

      "verifica se o numero é positivo
      IF column_int GT 0.

        DO.
          DATA(lv_mod) = ( lv_column_int - 1 ) MOD 26.         "divide o numero da coluna pela quantidade de letras do alfabeto - 1
          DATA(lv_div) = lv_column_int DIV 26.                 "divide o numero da coluna pela quantidade de letras do alfabeto
          lv_column_int = lv_div.                              "o numero recebe a quantidade da divisao
          column_string = sy-abcde+lv_mod(1) && column_string. "string recebe os caracteres referidos
          IF lv_column_int <= 0.
            EXIT.
          ENDIF.
        ENDDO.

      ENDIF.

    ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->CONVERT_XSTRING
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
    METHOD convert_xstring.

      "----------------------------------------------------------------------------------------------
      "info: converte a tabela de colaboradores em xstring
      "a tabela posteriormente é usada para preencher o excel file
      "
      "data de alteracao: 09.11.2024
      "alteracao: criacao do método
      "criado por: rafael albuquerque
      "----------------------------------------------------------------------------------------------

      DATA: lx_error      TYPE REF TO cx_root,  "define uma referência para exceções
            lv_error_text TYPE string.          "define uma variável para o texto do erro

      TRY.
          "cria o objeto para o conversor
          CREATE OBJECT o_converter.

          "converte os dados para o formato Excel
          o_converter->convert(
            EXPORTING
              it_table = me->it_colaboradores
            CHANGING
              co_excel = me->o_xl
          ).

          "verificação de erros na conversão
          IF sy-subrc NE 0.
            MESSAGE 'Não foi possível converter os dados para xstring' TYPE 'S' DISPLAY LIKE 'E'.
            RETURN.
          ENDIF.

        CATCH cx_root INTO lx_error.
          lv_error_text = lx_error->if_message~get_text( ).
          MESSAGE lv_error_text TYPE 'S' DISPLAY LIKE 'E'.
          RETURN.
      ENDTRY.


    ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->DISPLAY_FAST_EXCEL
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_TABLE_CONTENT                TYPE REF TO DATA
* | [--->] I_TABLE_NAME                   TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
    METHOD display_fast_excel.

      "----------------------------------------------------------------------------------------------
      "info: converte uma tabela da base de dados em arquivo excel
      "
      "data de alteracao: 09.11.2024
      "alteracao: criacao do método
      "criado por: rafael albuquerque
      "----------------------------------------------------------------------------------------------

      "-------------------------------------------------------------------------------
      "                           recebe uma tabela generica
      "-------------------------------------------------------------------------------

      "tipo de dados generico
      DATA: lr_table TYPE REF TO data.

      "instanciar esse tipo de dados em runtime para ser uma tabela do tipo (i_table_name)
      CREATE DATA lr_table TYPE TABLE OF (i_table_name).

      "preencher a tabela do método com o conteudo que vem no parametro
      lr_table = i_table_content.

      "como foi criada por referência ao tipo genérico "data" não dá para aceder diretamente
      "usar field symbol e apontar o conteudo da tabela (->*) para o field symbol
      FIELD-SYMBOLS: <lit_table> TYPE ANY TABLE.
      ASSIGN lr_table->* TO <lit_table>.

      CREATE OBJECT o_xl. "cria objeto excel
      CREATE OBJECT o_converter.

      "-------------------------------------------------------------------------------
      "                             converte para xstring
      "-------------------------------------------------------------------------------

      DATA: lx_error      TYPE REF TO cx_root,       "define uma referência para exceções
            lv_error_text TYPE string.          "define uma variável para o texto do erro

      TRY.
          "converte os dados para o formato Excel
          o_converter->convert(
            EXPORTING
              it_table      = <lit_table>
            CHANGING
              co_excel      = me->o_xl
          ).

          " Verificação de erros na conversão
          IF sy-subrc NE 0.
            MESSAGE 'Não foi possível converter os dados para xstring' TYPE 'S' DISPLAY LIKE 'E'.
            RETURN.
          ENDIF.

        CATCH cx_root INTO lx_error.
          lv_error_text = lx_error->if_message~get_text( ).
          MESSAGE lv_error_text TYPE 'S' DISPLAY LIKE 'E'.
          RETURN.
      ENDTRY.

      "cria um worksheet
      DATA(o_xl_ws) = o_xl->get_active_worksheet( ).
      lo_worksheet = o_xl_ws.

      "-------------------------------------------------------------------------------
      "                    conta a quantidade de colunas da tabela
      "-------------------------------------------------------------------------------

      DATA: lo_table_descr  TYPE REF TO cl_abap_tabledescr,
            lo_struct_descr TYPE REF TO cl_abap_structdescr.

      lo_table_descr ?= cl_abap_tabledescr=>describe_by_data( p_data = <lit_table> ).
      lo_struct_descr ?= lo_table_descr->get_table_line_type( ).

      DATA(lv_number_of_columns) = lines( lo_struct_descr->components ).

      "-------------------------------------------------------------------------------
      "                         setup das colunas - largura
      "-------------------------------------------------------------------------------

      me->set_style( ). "insere o estilo na coluna

      "contador de colunas
      DATA: count_columns TYPE i.
      count_columns = 1. "começa pela primeira

      "conta até a quantidade de colunas da tabela
      DO lv_number_of_columns TIMES.
        TRY.
            lo_column = lo_worksheet->get_column( ip_column = count_columns ).  "pega a coluna relativo ao index
            lo_column->set_width( ip_width = 30 ).                              "define o tamanho da coluna
          CATCH zcx_excel INTO DATA(lx_excel).
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        ADD 1 TO count_columns. "incrementa o contador de colunas
      ENDDO.

      count_columns = 1. "reseta o contador

      "titulo do worksheet principal
      DATA(worksheet_title) = CONV zexcel_sheet_title( |{ i_table_name }| ).

      TRY.
          lo_worksheet->set_title( ip_title = worksheet_title ).
        CATCH zcx_excel INTO DATA(lx_excel2).
          MESSAGE lx_excel2->get_text( ) TYPE 'E'.
      ENDTRY.

      "-------------------------------------------------------------------------------
      "                            caminho para o arquivo
      "-------------------------------------------------------------------------------

      "tratamento de nome e extensão do arquivo
      DATA full_path TYPE string.
      DATA namefile TYPE string.

      namefile = 'file'.

      "metodo que salva nome e diretorio
      me->get_file_directory(
        EXPORTING
          filename  = namefile
        IMPORTING
          full_path = full_path
      ).

      "se o download for cancelado...
      IF full_path IS INITIAL.
        MESSAGE 'O download foi cancelado pelo usuário.' TYPE 'S' DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.

      "-------------------------------------------------------------------------------
      "                             escritor para arquivo
      "-------------------------------------------------------------------------------

      "inicia o escritor do arquivo
      DATA(o_xlwriter)  = CAST zif_excel_writer( NEW zcl_excel_writer_2007( ) ).

      TRY.
          DATA(lv_xl_xdata) = o_xlwriter->write_file( o_xl ).
        CATCH zcx_excel INTO lx_excel.
          MESSAGE lx_excel->get_text( ) TYPE 'E'.
      ENDTRY.

      DATA(it_raw_data) = cl_bcs_convert=>xstring_to_solix( EXPORTING iv_xstring = lv_xl_xdata ).

      "-------------------------------------------------------------------------------
      "                            download do arquivo Excel
      "-------------------------------------------------------------------------------

      TRY.
          cl_gui_frontend_services=>gui_download(
            EXPORTING
              filename     = full_path
              filetype     = 'BIN'
              bin_filesize = xstrlen( lv_xl_xdata )
            CHANGING
              data_tab     = it_raw_data
          ).
        CATCH cx_root INTO lx_error.
          lv_error_text = lx_error->if_message~get_text( ).
          MESSAGE lv_error_text TYPE 'S' DISPLAY LIKE 'E'.
          RETURN.
      ENDTRY.

      "-------------------------------------------------------------------------------
      "                               tratamento de erros
      "-------------------------------------------------------------------------------

      IF sy-subrc NE 0.
        MESSAGE 'Não foi possível realizar o download do arquivo' TYPE 'S' DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.

    ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->DOWNLOAD_XLS
* +-------------------------------------------------------------------------------------------------+
* | [<---] RESULT                         TYPE        ZRLA_RESULT
* +--------------------------------------------------------------------------------------</SIGNATURE>
    METHOD download_xls.

      "----------------------------------------------------------------------------------------------
      "info: realiza download do arquivo excel file
      "
      "data de alteracao: 15.11.2024
      "alteracao: inclusao do parametro result para retorno de operacoes
      "criado por: rafael albuquerque
      "----------------------------------------------------------------------------------------------

      "tratamento de nome e extensão do arquivo
      DATA full_path TYPE string.
      DATA namefile TYPE string.

      namefile = 'file'. "nome default para arquivo

      "metodo que salva nome e diretorio
      me->get_file_directory(
        EXPORTING
          filename  = namefile
        IMPORTING
          full_path = full_path
      ).

      "se o download for cancelado...
      IF full_path IS INITIAL.
        MESSAGE 'O download foi cancelado pelo usuário.' TYPE 'S' DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.

      "----------------------------------------------------------------

      CREATE OBJECT o_xl. "cria objeto excel

      "insere o estilo
      me->set_style( ).
      "converte dados para xstring
      me->convert_xstring( ).
      "insere o worksheet com a tabela completa
      me->set_database(  ).
      "insere worksheets com cada linha da tabela individualmente
      me->set_sheets( ).

      "----------------------------------------------------------------

      "inicia o escritor do arquivo
      DATA(o_xlwriter) = CAST zif_excel_writer( NEW zcl_excel_writer_2007( ) ).

      TRY.
          DATA(lv_xl_xdata) = o_xlwriter->write_file( o_xl ).
        CATCH zcx_excel INTO DATA(lx_excel).
          MESSAGE lx_excel->get_text( ) TYPE 'E'.
      ENDTRY.

      DATA(it_raw_data) = cl_bcs_convert=>xstring_to_solix( EXPORTING iv_xstring = lv_xl_xdata ).

      "----------------------------------------------------------------

      "download do arquivo Excel
      TRY.
          cl_gui_frontend_services=>gui_download(
            EXPORTING
              filename     = full_path
              filetype     = 'BIN'
              bin_filesize = xstrlen( lv_xl_xdata )
            CHANGING
              data_tab     = it_raw_data
          ).
        CATCH cx_root INTO DATA(ex_txt).
          WRITE: / ex_txt->get_text( ).
      ENDTRY.

      "----------------------------------------------------------------

      "tratamento de erros
      IF sy-subrc NE 0.
        result-rc = sy-subrc.
        result-message = 'Não foi possível realizar o download do arquivo'.
        RETURN.
      ENDIF.

    ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GENERATE_CALENDAR
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
    METHOD generate_calendar.

      "----------------------------------------------------------------------------------------------
      "info: gera um calendario para cada colaborador de acordo com o mes requerido
      "
      "data de alteracao: 15.11.2024
      "alteracao: documentacao e testes falhos de insercao de numeros inteiros em ip_values das cells
      "criado por: rafael albuquerque
      "----------------------------------------------------------------------------------------------

      "metodo para capturar a data enviada pelo programa
      me->get_date( date = gv_datemonth ).

      "verificacao para envio de data
      IF gv_datemonth IS INITIAL.
        MESSAGE | 'Para impressão do calendário é preciso a data' | TYPE 'S' DISPLAY LIKE 'E'.
      ENDIF.

      "buscando a quantidade de dias no mes
      DATA: lv_date           TYPE /osp/dt_date, "data enviada
            lv_countdays      TYPE /osp/dt_day,  "dias do mes recebidos
            lv_countdays2     TYPE i,            "dias do mes em inteiro
            lv_counterdays    TYPE i,            "contador de dias
            lv_newdate        TYPE sy-datum,     "nova data formatada
            lv_stringdaydate  TYPE string,       "dia formatado
            lv_day            TYPE i,            "dia em inteiro
            lv_strday         TYPE string,       "dia em string
            lv_counterployees TYPE i,            "index da tabela de horarios de trabalho
            lv_stringhour     TYPE char5. "horario na celula em decimais

      DATA: horas_planeadas TYPE p DECIMALS 2.
      horas_planeadas = '8.00'.
      DATA: horas_planeadas2 TYPE string.

      "letra da coluna para formula para calculos de horas de trabalhos diarios
      DATA: lv_lettercollum TYPE string.

      "formula para dias trabalhados
      DATA: form_dia_trab TYPE string.

      "formula para somar horas planeadas e trabalhadas
      total_planeadas   = '=SUM(E7:AI7)'.   "formula para somar horas a trabalhar
      total_trabalhadas = '=SUM(E10:AI15)'. "formula para somar horas trabalhadas

      "-------------------------------------------

      lv_date = gv_datemonth. "recebe a data enviada pelo programa

      "funcao retorna a quantidade de dias do mes
      CALL FUNCTION '/OSP/GET_DAYS_IN_MONTH'
        EXPORTING
          iv_date = lv_date
        IMPORTING
          ev_days = lv_countdays.

      lv_countdays2 = lv_countdays. "casting int
      lv_counterdays = 5.           "inicia o contador como cinco para contar a partir da 5th coluna

      "-------------------------------------------
      "  formatacao da data para começo do mês
      "-------------------------------------------

      "reseta a data
      lv_newdate = lv_date+0(6). "recebe ano + mes
      lv_strday = '01'.          "sempre começamos pelo primeiro dia do mes
      "junta ano + mes e primeiro dia do mes
      CONCATENATE lv_newdate lv_strday INTO lv_newdate.

      "---------------------------------------------------------------------------------------------
      "                                  work schedule do cats
      "---------------------------------------------------------------------------------------------

      lv_counterployees = 1. "inicia o contador de index da tabela horarios

      "pega os horarios de cada funcionario por index de tabela
      READ TABLE me->it_colaboradores INTO me->ls_colaborador INDEX lv_counterployees.
      me->get_work_schedule( pernr = me->ls_colaborador-pernr ). "metodo para buscar work schedule

      "repete a quantidade de dias que tem o mes
      DO lv_countdays TIMES.

        "funcao retorna a data formatada [ numdia + nomediasemana ]
        CALL FUNCTION 'ZWEEKDATE'
          EXPORTING
            date           = lv_newdate
          IMPORTING
            format_daydate = lv_stringdaydate
            e_result       = e_result.

        IF sy-subrc EQ 0.

          "verifica se é sábado ou domingo para nao contabilizar as horas.
          IF lv_stringdaydate CS 'Sábado' OR lv_stringdaydate CS 'Domingo'.
            horas_planeadas = '0'.
          ELSE.

            "busca o a quantidade de horarios diários do colaborador
            READ TABLE me->tb_psp INTO me->wa_psp INDEX lv_counterployees.

            "verifica se é feriado
            CASE me->wa_psp-ftkla.
              WHEN 1.
                horas_planeadas = '0'.
              WHEN 0.
                horas_planeadas = me->wa_psp-stdaz.
            ENDCASE.

            CLEAR me->wa_psp.

          ENDIF.

          ADD 1 TO lv_counterployees. "passa para o proximo colaborador na verificacao do workschedule

          "------------------------------------------------------------------------------------------------------------------------------------------------------------------
          "                                            impressao do calendario - "dias do calendario 01-31
          "------------------------------------------------------------------------------------------------------------------------------------------------------------------

          "lv_stringdaydate = retorno da funcao zweekdate
          "horas_planeadas  = retorno do workschedule do cats do colaborador

          TRY.
              lo_worksheet->set_cell( ip_row = 6 ip_column = lv_counterdays ip_value = lv_stringdaydate ip_style = tp_style_bold_center_guid ).  "cabeçalho do calendário
              lo_worksheet->set_cell( ip_row = 7 ip_column = lv_counterdays ip_value = horas_planeadas  ip_style = tp_style_bold_center_guid2 ). "horas planeadas
            CATCH zcx_excel INTO DATA(lx_excel).
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          CLEAR lv_lettercollum. "limpa a letra da coluna para evitar concatenacoes

          "-----------------------------------------------------------------------------------------------------

          "converte o numero da coluna em string da coluna
          TRY.
              me->convert_excel_column(
                EXPORTING
                  column_int    = lv_counterdays   "numero da coluna
                IMPORTING
                  column_string = lv_lettercollum  "coluna do excel
              ).
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          "celula que recebe a formula da soma de horas trabalhadas no dia
          "a formula é atualizada conforme o calendario é construído
          CLEAR form_dia_trab.
          form_dia_trab = '=SUM(' && lv_lettercollum && '10:' && lv_lettercollum && '15)'. "atualiza a formula

          "------------------------------------------------------------------------------------------------------

          TRY.
              lo_worksheet->set_cell( ip_row = 8 ip_column = lv_counterdays ip_value = '0,0' ip_formula = form_dia_trab ip_style = tp_style_bold_center_guid2 ). "horas trabalhadas
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          "valor do horario dos tempos gastos em projetos
          lv_stringhour = '0,0'. "é obrigatório que se envie uma string

          "cabeçalho de tempo trabalhado ou ausentado
          TRY.
              lo_worksheet->set_cell( ip_row = 9 ip_column = lv_counterdays ip_value = 'Tempo' ip_style = tp_style_bold_center_guid ). "horas trabalhadas
              "colunas de tempo de trabalho
              lo_worksheet->set_cell( ip_row = 10 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ip_conv_exit_length = abap_true ).
              lo_worksheet->set_cell( ip_row = 11 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ip_conv_exit_length = abap_true ).
              lo_worksheet->set_cell( ip_row = 12 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ip_conv_exit_length = abap_true ).
              lo_worksheet->set_cell( ip_row = 13 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ip_conv_exit_length = abap_true ).
              lo_worksheet->set_cell( ip_row = 14 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ip_conv_exit_length = abap_true ).
              lo_worksheet->set_cell( ip_row = 15 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ip_conv_exit_length = abap_true ).
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          "configuracoes das colunas
          TRY.
              "setup da coluna para cada celula criada
              lo_column = lo_worksheet->get_column( ip_column = lv_counterdays ).
              lo_column->set_width( ip_width = 25 ). " Define o tamanho da coluna
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          ADD 1 TO lv_counterdays. "incrementa o contador para a proxima coluna

          "-----------------------------------------------------------------------------------------
          "                                  atualizacao da data
          "-----------------------------------------------------------------------------------------

          lv_day = lv_strday. "casting int
          ADD 1 TO lv_day.    "incrementa o dia
          lv_strday = lv_day. "casting string
          "se nao passamos dos 10 primeiros dias do mês
          IF lv_day LT 10.
            CONCATENATE '0' lv_strday INTO lv_strday. "adiciona o 0 na frente do numero
          ENDIF.

          CLEAR lv_newdate.                                 "limpa a variavel
          lv_newdate = lv_date+0(6).                        "busca novamente ano e mes
          CONCATENATE lv_newdate lv_strday INTO lv_newdate. "redefine a data para o dia seguinte.

          "-----------------------------------------------------------------------------------------

        ENDIF.

      ENDDO.

      "---------------------------------------------------------------------------------------------------------------------------------------------------------------
      "                                                 impressao do complemento do calendario - caso necessite
      "---------------------------------------------------------------------------------------------------------------------------------------------------------------

      "verifica quanto falta para 31 dias para completar o calendario
      IF lv_countdays LT 31.
        "enquanto o calendario nao estiver completo...
        WHILE lv_countdays LT 31.

          "cria a celula
          TRY.
              lo_worksheet->set_cell( ip_row = 6 ip_column = lv_counterdays ip_value = 'XXXXXXX'  ip_style = tp_style_bold_center_guid ).  "cabeçalho do calendário
              lo_worksheet->set_cell( ip_row = 7 ip_column = lv_counterdays ip_value = '0,0'      ip_style = tp_style_bold_center_guid2 ). "horas planeadas
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          TRY.
              lo_worksheet->set_cell( ip_row = 8 ip_column = lv_counterdays ip_value = '0,0'  ip_style = tp_style_bold_center_guid2 ). "horas trabalhadas
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          "cabeçalho de tempo trabalhado ou ausentado
          TRY.
              lo_worksheet->set_cell( ip_row = 9 ip_column = lv_counterdays ip_value = 'Tempo' ip_style = tp_style_bold_center_guid ). "horas trabalhadas
              "colunas de tempo de trabalho
              lo_worksheet->set_cell( ip_row = 10 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ).
              lo_worksheet->set_cell( ip_row = 11 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ).
              lo_worksheet->set_cell( ip_row = 12 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ).
              lo_worksheet->set_cell( ip_row = 13 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ).
              lo_worksheet->set_cell( ip_row = 14 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ).
              lo_worksheet->set_cell( ip_row = 15 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ).
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          TRY.
              "setup da coluna para cada celula criada
              lo_column = lo_worksheet->get_column( ip_column = lv_counterdays ).
              lo_column->set_width( ip_width = 25 ). " Define o tamanho da coluna
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          ADD 1 TO lv_countdays.
          ADD 1 TO lv_counterdays.

        ENDWHILE.

      ENDIF.

      "---------------------------------------------------------------------------------------------------------------------------------------------------------------

      "resets e refreshs
      lv_counterdays = 5. "reseta o contador para a 5th coluna
      lv_counterployees = 1. "reseta o contador de horarios de trabalho
      CLEAR: lv_day, lv_strday. "limpa os contadores de dias em string e int.
      REFRESH me->tb_psp.

    ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_AUSPRES
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
    METHOD get_auspres.

      "----------------------------------------------------------------------------------------------
      "info: recebe ausencias e presencas da base de dados
      "
      "data de alteracao: 15.11.2024
      "alteracao: modificacao de verificacao de retorno de dados
      "criado por: rafael albuquerque
      "----------------------------------------------------------------------------------------------

      "---------------------------------------------------
      "     criação do range de datas para consulta
      "---------------------------------------------------

      DATA: begin_month TYPE begda,
            end_month   TYPE endda.

      "recebe o range do inicio e o final do mes
      me->set_rangemonthdate(
        IMPORTING
          begin_date = begin_month
          end_date   = end_month
      ).

      "---------------------------------------------------
      " consulta para obter textos de ausencia e presenca
      "---------------------------------------------------

      SELECT t554s~subty,               "tipo de ausência e presenca
             t554t~atext                "Texto descritivo
        FROM t554s                      "Da tabela do de Tipos de presença e ausência
        INNER JOIN t554t                "Junta da tabela de Textos de ausência e presença
        ON t554s~moabw = t554t~moabw    "Juntas por chave de agrupamento em RH
        INTO TABLE @me->it_aus_pre
        WHERE t554s~moabw EQ 19
        AND   t554t~moabw EQ 19
        AND   t554t~sprsl EQ @sy-langu      "Onde o idioma for aquele do sistema
        AND   t554t~atext NE ''             "O texto não esteja vazio
        AND   t554s~endda GT @end_month     "E a data fim seja maior do que a data final do mes
        AND   t554s~begda LT @begin_month   "E a data inicio maior que a data final do mes
        AND   t554s~subty EQ t554t~awart.   "Onde o tipo de ausencia e presenca é igual ao Texto de ausência e presença

        "----------------------------------------------------
        "      verifica se houveram dados retornados
        "----------------------------------------------------

        IF me->it_aus_pre IS INITIAL.
          MESSAGE | Não foram achados motivos de ausência e presença na base de dados | TYPE 'S' DISPLAY LIKE 'E'.
          RETURN.
        ENDIF.

        "----------------------------------------------------
        "formacao da linha de textos para ausencia e presenca
        "----------------------------------------------------

        DATA stringline TYPE string.

        "itera sobre a tabela de textos concatenando Tipos de presença e ausência com os Textos de ausência e presença
        LOOP AT me->it_aus_pre INTO me->ls_aus_pre.
          stringline = me->ls_aus_pre-subty. "casting do numero
          CONCATENATE stringline me->ls_aus_pre-atext INTO me->ls_line_preaus-line SEPARATED BY ' - '.
          APPEND me->ls_line_preaus TO me->it_line_preaus.
          CLEAR stringline.
        ENDLOOP.

      ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_AUSPRES_DATAFILE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
      METHOD get_auspres_datafile.

        "----------------------------------------------------------------------------------------------
        "info: recebe ausencias e presencas do arquivo excel
        "
        "data de alteracao: 10.11.2024
        "alteracao: criacao do método
        "criado por: rafael albuquerque
        "----------------------------------------------------------------------------------------------

        "----------------------------------------------------
        "    verificacao de dados essenciais para consulta
        "----------------------------------------------------

        IF me->lv_xstr IS INITIAL.          "tabela convertida em xstring
          RETURN.
        ELSEIF me->gv_datemonth IS INITIAL. "data do mes
          RETURN.
        ELSEIF me->it_employee IS INITIAL.  "tabela de colaboradores
          RETURN.
        ENDIF.

        "----------------------------------------------------
        "             tratmento das coordenadas
        "----------------------------------------------------

        DATA: lv_index TYPE i.
        lv_index = 2.

        "coordenada da celula
        DATA: lv_coord     TYPE string,
              lv_coord_num TYPE i,
              lv_str_coord TYPE string.

        DATA: lv_hour_index TYPE i.

        "metodo para envio de coordenadas
        me->set_coordenates(
          EXPORTING
            letter_coord   = 'B'
          IMPORTING
*        string_coord_a =
            string_coord_b = lv_str_coord
            index_coord    = lv_hour_index
        ).

        "----------------------------------------------------
        "                 leitura do arquivo
        "----------------------------------------------------

        "flag da sheet
        DATA: flag_next_sheet TYPE flag.
        flag_next_sheet = abap_false.

        "leitor do arquivo
        DATA(lo_reader) = NEW zcl_excel_reader_2007( ).
        DATA(lo_excel)  = lo_reader->zif_excel_reader~load( i_excel2007 = me->lv_xstr ).  "passa o XSTRING carregado

        DATA(i) = 2. "primeira pagina de colaboradores

        "itera por todas as sheets do excel, seja ela quantas houverem
        WHILE i <= lo_excel->get_worksheets_size( ).

          "começa a partir da segunda sheet, sendo a primeira a exibicao de dados gerais
          DATA(lo_worksheet) = lo_excel->get_worksheet_by_index( i ).

          CLEAR me->ls_auspres.

          "define a coordenada da celula ao inicio da sheet
          lv_str_coord = lv_coord_num.
          CONCATENATE lv_coord lv_str_coord INTO lv_str_coord. "B10
          CONDENSE lv_str_coord NO-GAPS.

          "pega primeiramente o numero do colaborador
          READ TABLE lo_worksheet->sheet_content REFERENCE INTO DATA(cell) INDEX lv_index. "B2

          "numero do colaborador
          me->ls_auspres-num = cell->cell_value.

          "----------------------------------------------------
          "     itera sobre as seis ausencias e presencas
          "----------------------------------------------------

          DO 6 TIMES.

            "procura se motivos de ausencia e presenca
            READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell WITH KEY cell_coords = lv_str_coord. "B10...

            "se houver ausencia ou presenca disponivel
            IF cell->cell_value NE 'Selecione'.
              ls_auspres-auspres = cell->cell_value. "recebe o nome do pep

              "itera sobre os 31 dias do mes -- valor fixo
              DO 31 TIMES.

                "verifica as horas de ausencia e presenca
                READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_hour_index. "E10

                "se houver hora...
                IF cell->cell_value NE '0' AND cell->cell_value NE '0,0'.

                  me->ls_auspres-dia  = gv_datemonth.     "recebe o dia do mes
                  me->ls_auspres-hora = cell->cell_value. "recebe a hora trabalhada
                  me->ls_auspres-row  = lv_coord_num.     "recebe a linha do projeto
                  APPEND ls_auspres TO it_auspres.           "insere a tabela de peps
                  CLEAR: ls_auspres-dia, ls_auspres-hora.
                  CLEAR: cell->cell_value.
                ENDIF.

                ADD 1 TO lv_hour_index. "incrementa para a proxima hora
                ADD 1 TO gv_datemonth.  "incrementa para o proximo dia

              ENDDO.

              me->get_month_datafile( ). "reseta data do mes

            ENDIF.

            "---------------------------------------------------------------------
            " redefine a coordenada para o proximo motivo de ausencia ou presenca
            "---------------------------------------------------------------------

            lv_coord = 'B'.
            ADD 1 TO lv_coord_num.
            lv_str_coord = lv_coord_num.
            CONCATENATE lv_coord lv_str_coord INTO lv_str_coord.
            CONDENSE lv_str_coord NO-GAPS.

            "redefine o index de horarios conforme coordenada
            CASE lv_str_coord.
              WHEN 'B10'.
                lv_hour_index = 144.
              WHEN 'B11'.
                lv_hour_index = 177.
              WHEN 'B12'.
                lv_hour_index = 210.
              WHEN 'B13'.
                lv_hour_index = 243.
              WHEN 'B14'.
                lv_hour_index = 276.
              WHEN 'B15'.
                lv_hour_index = 309.
            ENDCASE.

            CLEAR: ls_auspres-dia, ls_auspres-hora, ls_auspres-auspres, ls_auspres-row.

          ENDDO.

          "-------------------------------------------
          "    redefine dados para proxima sheet.
          "-------------------------------------------

          lv_hour_index = 144. "index de ausencias e presencas

          me->get_month_datafile( ). "reseta data do mes

          "passa para a próxima sheet
          ADD 1 TO i.
          lv_index = 2.

          CLEAR: lv_str_coord. "limpa a coordenada
          lv_coord_num = 10. "redefine a linha da coordenada

          CLEAR ls_auspres. "limpa a estrutura para a proxima sheet

        ENDWHILE.

      ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->GET_DATA
* +-------------------------------------------------------------------------------------------------+
* | [--->] COLABORADORES                  TYPE        ZCOL_TT
* +--------------------------------------------------------------------------------------</SIGNATURE>
      METHOD get_data.

        "----------------------------------------------------------------------------------------------
        "info: recebe os dados dos colaboradores da base de dados
        "
        "data de alteracao: 14.11.2024
        "alteracao: criacao do método
        "criado por: rafael albuquerque
        "----------------------------------------------------------------------------------------------

        "mockdata
*    me->it_colaboradores = VALUE #( ( pernr = '1'  sname = 'Colaborador A' vdsk1 = 'PT01'  kostl = '001'  )
*                                    ( pernr = '2'  sname = 'Colaborador B' vdsk1 = 'PT02'  kostl = '002'  )
*                                    ( pernr = '3'  sname = 'Colaborador C' vdsk1 = 'PT03'  kostl = '003'  )
*                                    ( pernr = '4'  sname = 'Colaborador D' vdsk1 = 'PT04'  kostl = '004'  )
*                                    ( pernr = '5'  sname = 'Colaborador E' vdsk1 = 'PT05'  kostl = '005'  )
*                                    ( pernr = '6'  sname = 'Colaborador F' vdsk1 = 'PT06'  kostl = '006'  )
*                                    ( pernr = '7'  sname = 'Colaborador G' vdsk1 = 'PT07'  kostl = '007'  )
*                                    ( pernr = '8'  sname = 'Colaborador H' vdsk1 = 'PT08'  kostl = '008'  )
*                                    ( pernr = '9'  sname = 'Colaborador I' vdsk1 = 'PT09'  kostl = '009'  )
*                                    ( pernr = '10' sname = 'Colaborador J' vdsk1 = 'PT010' kostl = '0010' ) ).

        "recebe a tabela de colaboradores vindo da aplicacao principal
        me->it_colaboradores = colaboradores. "recebe uma tabela interna e preenche o atributo de classe

        "verifica se algum dado foi enviado
        IF colaboradores IS INITIAL.
          MESSAGE | Não foi possível receber os dados da base de dados | TYPE 'S' DISPLAY LIKE 'E'.
        ENDIF.

        "------------------------------------------------------------
        me->get_auspres( ).  "recebe os dados de ausencia e presenca
        me->get_projects( ). "consulta para obter textos de projetos
        "------------------------------------------------------------

      ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->GET_DATE
* +-------------------------------------------------------------------------------------------------+
* | [--->] DATE                           TYPE        SY-DATUM
* +--------------------------------------------------------------------------------------</SIGNATURE>
      METHOD get_date.

        "----------------------------------------------------------------------------------------------
        "info: atributo de classe recebe data atual do sistema
        "
        "data de alteracao: 09.11.2024
        "alteracao: criacao do método
        "criado por: rafael albuquerque
        "----------------------------------------------------------------------------------------------

        "se a data nao for enviada...
        "envia a data atual do sistema.
        IF date IS INITIAL.
          me->gv_datemonth = sy-datum.
        ELSE.
          me->gv_datemonth = date.
        ENDIF.

      ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_EMPLOYEE_DATAFILE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
      METHOD get_employee_datafile.

        "----------------------------------------------------------------------------------------------
        "info: recebe os dados dos colaboradores do arquivo excel e guarda numa tabela interna
        "
        "data de alteracao: 11.11.2024
        "alteracao: criacao do método
        "criado por: rafael albuquerque
        "----------------------------------------------------------------------------------------------

        IF lv_xstr IS INITIAL.
          RETURN.
        ENDIF.

        DATA: lv_index TYPE i.
        lv_index = 2.

        CLEAR me->it_employee.

        DATA(lo_reader) = NEW zcl_excel_reader_2007( ).
        DATA(lo_excel)  = lo_reader->zif_excel_reader~load( i_excel2007 = me->lv_xstr ).  "passa o XSTRING carregado

        DATA(i) = 2.

        "itera por todas as sheets do excel, seja ela quantas houverem
        WHILE i <= lo_excel->get_worksheets_size( ).

          "começa a partir da segunda sheet, sendo a primeira a exibicao de dados gerais
          DATA(lo_worksheet) = lo_excel->get_worksheet_by_index( i ).

          CLEAR me->ls_employee.

          "-----------------------------------------------------
          "           cabeçalho de colaboradores
          "-----------------------------------------------------

          READ TABLE lo_worksheet->sheet_content REFERENCE INTO DATA(cell) INDEX lv_index. "B2

          "numero do colaborador
          me->ls_employee-num = cell->cell_value.

          ADD 2 TO lv_index.

          READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_index. "B3

          "nome do colaborador
          me->ls_employee-nome = cell->cell_value.

          ADD 2 TO lv_index.

          READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_index. "B4

          "equipa do colaborador
          me->ls_employee-equipa = cell->cell_value.

          ADD 2 TO lv_index.

          READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_index. "B5

          "centro de custo do colaborador
          me->ls_employee-cntr_cust = cell->cell_value.

          APPEND ls_employee TO it_employee.

          "passa para a próxima sheet
          ADD 1 TO i.
          lv_index = 2.

        ENDWHILE.

      ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_FILE_DIRECTORY
* +-------------------------------------------------------------------------------------------------+
* | [--->] FILENAME                       TYPE        STRING
* | [<---] FULL_PATH                      TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
      METHOD get_file_directory.

        "----------------------------------------------------------------------------------------------
        "info: recebe o diretorio do arquivo
        "
        "data de alteracao: 09.11.2024
        "alteracao: criacao do método
        "criado por: rafael albuquerque
        "----------------------------------------------------------------------------------------------

        DATA: namefile  TYPE string, "nome do arquivo
              directory TYPE string, "diretorio
              fullpath  TYPE string. "caminho completo

        namefile = 'file'.

        "adiciona a extensão '.xlsx' ao nome do arquivo
        me->append_extension(
          EXPORTING
            old_extension = namefile
          IMPORTING
            new_extension = namefile
        ).

        "diálogo para selecionar diretorio e nome do arquivo
        CALL METHOD cl_gui_frontend_services=>file_save_dialog
          EXPORTING
            default_extension = 'xlsx'
            default_file_name = namefile
          CHANGING
            filename          = namefile
            path              = directory
            fullpath          = fullpath
          EXCEPTIONS
            OTHERS            = 1.

        "se o user nao cancelar a operacao...
        IF sy-subrc = 0.
          CONCATENATE directory namefile INTO fullpath SEPARATED BY '\'. "cria diretorio completo do arquivo
        ELSE.
          CLEAR fullpath. "limpa o caminho
        ENDIF.

        full_path = fullpath. "retorna caminho completo do arquivo

      ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_MONTH_DATAFILE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
      METHOD get_month_datafile.

        "----------------------------------------------------------------------------------------------
        "info: recebe o mes do arquivo excel referente a celula B6
        "
        "data de alteracao: 11.11.2024
        "alteracao: criacao do método
        "criado por: rafael albuquerque
        "----------------------------------------------------------------------------------------------

        IF lv_xstr IS INITIAL.
          RETURN.
        ENDIF.

        DATA: integerdate TYPE i. "recebe o numeiro inteiro da data do arquivo excel

        CLEAR me->gv_datemonth.

        DATA(lo_reader) = NEW zcl_excel_reader_2007( ).
        DATA(lo_excel)  = lo_reader->zif_excel_reader~load( i_excel2007 = me->lv_xstr ).  "passa o XSTRING carregado

        "começa a partir da segunda sheet, sendo a primeira a exibicao de dados gerais
        DATA(lo_worksheet) = lo_excel->get_worksheet_by_index( 2 ).

        "lê diretamente a celula onde está a data
        READ TABLE lo_worksheet->sheet_content REFERENCE INTO DATA(cell) INDEX 10. "B6

        "numero do colaborador
        integerdate = cell->cell_value.                    "recebe o valor inteiro
        integerdate = integerdate - 2.                     "remove os dois dias da data
        me->gv_datemonth = '19000101'.                     "recebe data default do sistema
        me->gv_datemonth = me->gv_datemonth + integerdate. "soma com a quantidade do numero inteiro

      ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_PEPS_DATAFILE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
      METHOD get_peps_datafile.

        "----------------------------------------------------------------------------------------------
        "info: recebe os peps dos colaboradores no excel file
        "
        "data de alteracao: 11.11.2024
        "alteracao: criacao do método
        "criado por: rafael albuquerque
        "----------------------------------------------------------------------------------------------

        IF me->lv_xstr IS INITIAL.
          RETURN.
        ELSEIF me->gv_datemonth IS INITIAL.
          RETURN.
        ELSEIF me->it_employee IS INITIAL.
          RETURN.
        ENDIF.

        DATA: lv_index TYPE i.
        lv_index = 2.

        "coordenada da celula
        DATA: lv_coord     TYPE string,
              lv_coord_num TYPE i,
              lv_str_coord TYPE string.

        lv_coord = 'A'.
        lv_coord_num = 10.

        DATA: lv_hour_index TYPE i.
        lv_hour_index = 144.

        "flag da sheet
        DATA: flag_next_sheet TYPE flag.
        flag_next_sheet = abap_false.

        DATA(lo_reader) = NEW zcl_excel_reader_2007( ).
        DATA(lo_excel)  = lo_reader->zif_excel_reader~load( i_excel2007 = me->lv_xstr ).  "passa o XSTRING carregado

        DATA(i) = 2.

        "itera por todas as sheets do excel, seja ela quantas houverem
        WHILE i <= lo_excel->get_worksheets_size( ).

          "começa a partir da segunda sheet, sendo a primeira a exibicao de dados gerais
          DATA(lo_worksheet) = lo_excel->get_worksheet_by_index( i ).

          CLEAR me->ls_peps.

          "define a coordenada da celula ao inicio da sheet
          lv_str_coord = lv_coord_num.
          CONCATENATE lv_coord lv_str_coord INTO lv_str_coord. "A10
          CONDENSE lv_str_coord NO-GAPS.

          "pega primeiramente o numero do colaborador
          READ TABLE lo_worksheet->sheet_content REFERENCE INTO DATA(cell) INDEX lv_index. "B2

          "numero do colaborador
          me->ls_peps-num = cell->cell_value.

          "itera sobre os seis projetos
          DO 6 TIMES.

            "procura se há peps ativas
            READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell WITH KEY cell_coords = lv_str_coord. "A10...

            "se houver pep disponivel
            IF cell->cell_value NE 'Selecione'.
              ls_peps-pep = cell->cell_value. "recebe o nome do pep

              "itera sobre os 31 dias do mes -- valor fixo
              DO 31 TIMES.

                "verifica as horas trabalhadas
                READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_hour_index. "E10

                "se houver hora trabalhada...
                IF cell->cell_value NE '0' AND cell->cell_value NE '0,0'.

                  me->ls_peps-dia  = gv_datemonth.     "recebe o dia do mes
                  me->ls_peps-hora = cell->cell_value. "recebe a hora trabalhada
                  me->ls_peps-row  = lv_coord_num.     "recebe a linha do projeto
                  APPEND ls_peps TO it_peps.           "insere a tabela de peps
                  CLEAR: ls_peps-dia, ls_peps-hora.
                  CLEAR: cell->cell_value.
                ENDIF.

                ADD 1 TO lv_hour_index. "incrementa para a proxima hora
                ADD 1 TO gv_datemonth.  "incrementa para o proximo dia

              ENDDO.

              me->get_month_datafile( ). "reseta data do mes

            ENDIF.

            "redefine a coordenada para o proximo projeto
            lv_coord = 'A'.
            ADD 1 TO lv_coord_num.
            lv_str_coord = lv_coord_num.
            CONCATENATE lv_coord lv_str_coord INTO lv_str_coord.
            CONDENSE lv_str_coord NO-GAPS.

            "redefine o index de horarios conforme coordenada
            CASE lv_str_coord.
              WHEN 'A10'.
                lv_hour_index = 144.
              WHEN 'A11'.
                lv_hour_index = 177.
              WHEN 'A12'.
                lv_hour_index = 210.
              WHEN 'A13'.
                lv_hour_index = 243.
              WHEN 'A14'.
                lv_hour_index = 276.
              WHEN 'A15'.
                lv_hour_index = 309.
            ENDCASE.

            CLEAR: ls_peps-dia, ls_peps-hora, ls_peps-pep, ls_peps-row.

          ENDDO.

          "-------------------------------------------
          "    redefine dados para proxima sheet.
          "-------------------------------------------

          lv_hour_index = 144. "index de horas trabalhadas

          me->get_month_datafile( ). "reseta data do mes

          "passa para a próxima sheet
          ADD 1 TO i.
          lv_index = 2.

          CLEAR: lv_str_coord. "limpa a coordenada
          lv_coord_num = 10. "redefine a linha da coordenada

          CLEAR ls_peps. "limpa a estrutura para a proxima sheet

        ENDWHILE.

      ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_PROJECTS
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
      METHOD get_projects.

        "----------------------------------------------------------------------------------------------
        "info: recebe os projetos ativos da base de dados
        "
        "data de alteracao: 09.11.2024
        "alteracao: criacao do método
        "criado por: rafael albuquerque
        "----------------------------------------------------------------------------------------------

        DATA stringline TYPE string.

        SELECT proj~objnr, "Nº objeto
               proj~pspid, "Definição do projeto
               proj~post1  "PS: descrição breve (1ª linha de texto)
          FROM proj AS proj
          INNER JOIN jest AS jest
          ON proj~objnr = jest~objnr
          INTO TABLE @it_projetos
          WHERE jest~inact EQ ''
          AND jest~stat EQ 'I0002'.

          "formacao da linha de textos para projetos
          "---------------------------------------------

          "itera sobre a tabela de textos concatenando o numero dos projetos à descricao dos projetos
          LOOP AT me->it_projetos INTO me->ls_projetos.
            stringline = me->ls_projetos-objnr.
            CONCATENATE stringline me->ls_projetos-post1 me->ls_projetos-pspid INTO me->ls_linha_projeto-line SEPARATED BY ' - '.
            APPEND me->ls_linha_projeto TO me->it_linha_projetos.
          ENDLOOP.

          CLEAR stringline.

        ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->GET_TIMESHEET_DATAFILE
* +-------------------------------------------------------------------------------------------------+
* | [<---] TABLE_TIMESHEET                TYPE        ZTSHRALV_TT
* | [<---] RESULT                         TYPE        ZRLA_RESULT
* +--------------------------------------------------------------------------------------</SIGNATURE>
        METHOD get_timesheet_datafile.

          "-------------------------------------------------------
          "info: preenche a tabela que vai ser exibida no alv
          "
          "data de alteracao: 09.11.2024
          "alteracao: criacao do método
          "criado por: rafael albuquerque
          "-------------------------------------------------------

          "verifica se o atributo está preenchido
          IF me->it_timesheet IS INITIAL.
            result-rc = 1.
            result-message = | Não foi possível receber os dados da Timesheet. |.
            RETURN.
          ENDIF.

          "limpa possiveis dados
          REFRESH table_timesheet.
          REFRESH me->tt_alv.

          DATA: ts TYPE ztshralv_st. "linha da timesheet.

          "passa os dados do atributo para a tabela interna
          "sao todos os dados coletados do arquivo e com verificacoes para celulas vazias
          LOOP AT me->it_timesheet INTO me->ls_timesheet.
            ts-num       = me->ls_timesheet-num.
            ts-nome      = me->ls_timesheet-nome.
            ts-equipa    = me->ls_timesheet-equipa.
            ts-cntr_cust = me->ls_timesheet-cntr_cust.
            ts-dia       = me->ls_timesheet-dia.
            ts-pep       = me->ls_timesheet-pep.
            ts-auspres   = me->ls_timesheet-auspres.

            "-------------------------------------------------------------------------------
            "verificacoes para casos de celulas erradas em projetos de ausencias e presencas

            "tratamento das horas do projeto em caso de palavras
            DATA: lv_hours TYPE string.
            lv_hours = me->ls_timesheet-hora.
            TRANSLATE lv_hours TO UPPER CASE.

            "se houver letra ao invés de um número
            IF lv_hours CA sy-abcde.
              me->ls_timesheet-hora = '0'.
              me->ls_timesheet-info = '@05@' && 'Caractere Inválido' .
            ELSEIF me->ls_timesheet-hora IS INITIAL.
              me->ls_timesheet-hora = '0'.
              me->ls_timesheet-info = '@05@' && 'Célula Vazia' .
            ENDIF.

            "-------------------------------------------------------------------------------

            ts-hora      = me->ls_timesheet-hora.
            ts-info      = me->ls_timesheet-info.

            APPEND ts TO table_timesheet.
            CLEAR me->ls_timesheet.
            CLEAR ts.
          ENDLOOP.

          "-------------------------------------------------------------------------------
          "verificacoes para casos de celulas erradas sem projetos de ausencias e presencas

          me->get_wronglines_datafile( ). "metodo que procura erros no excel que nao estejam vinculados a projetos

          "passa linhas de erros para a tabela de saida
          IF me->tt_alv IS NOT INITIAL.
            LOOP AT me->tt_alv INTO me->st_alv.
              APPEND me->st_alv TO table_timesheet.
            ENDLOOP.
            REFRESH me->tt_alv.
          ENDIF.

          "-------------------------------------------------------------------------------

          "ordenacao da tabela de saida
          SORT table_timesheet BY dia nome ASCENDING. "ordena por dia e nome.

          "validacao das linhas da timesheet (pernr, equipa, centro de custo, projeto e ausencias e presencas
          me->validation_datafile(
            EXPORTING
              table_timesheet        = table_timesheet
            IMPORTING
              table_timesheet_output = table_timesheet
          ).

        ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_WORK_SCHEDULE
* +-------------------------------------------------------------------------------------------------+
* | [--->] PERNR                          TYPE        P_PERNR
* +--------------------------------------------------------------------------------------</SIGNATURE>
        METHOD get_work_schedule.

          "----------------------------------------------------------------------------------------------
          "info: recebe os horarios de trabalho do colaborador enviados da base de dados
          "
          "data de alteracao: 11.11.2024
          "alteracao: criacao do método
          "criado por: rafael albuquerque
          "----------------------------------------------------------------------------------------------

          DATA: begda TYPE begda,
                endda TYPE endda.

          "recebe range do mes enviado
          me->set_rangemonthdate(
            IMPORTING
              begin_date = begda
              end_date   = endda
          ).

          "retorna a tabela com todas as horas de trabalho do mes do funcionario
          CALL FUNCTION 'HR_PERSONAL_WORK_SCHEDULE'
            EXPORTING
              pernr             = pernr
              begda             = begda
              endda             = endda
              switch_activ      = '1'
              i0001_i0007_error = '0'
              read_cluster      = ''
            TABLES
              perws             = me->tb_psp.

        ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_WRONGLINES_DATAFILE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
        METHOD get_wronglines_datafile.

          "----------------------------------------------------------------------------------------------
          "info: procura linhas mal preenchidas no documento
          "
          "data de alteracao: 11.11.2024
          "alteracao: criacao do método
          "criado por: rafael albuquerque
          "----------------------------------------------------------------------------------------------

          IF me->lv_xstr IS INITIAL.
            RETURN.
          ELSEIF me->gv_datemonth IS INITIAL.
            RETURN.
          ELSEIF me->it_employee IS INITIAL.
            RETURN.
          ENDIF.

          REFRESH tt_alv.

          DATA: lv_index TYPE i.
          lv_index = 2.

          "coordenada da celula
          DATA: lv_coord      TYPE string,
                lv_coord2     TYPE string,
                lv_coord_num  TYPE i,
                lv_str_coord  TYPE string,
                lv_str_coord2 TYPE string.

          "setup das coordenadas
          lv_coord = 'A'.
          lv_coord_num = 10.
          lv_coord2 = 'B'.

          DATA: lv_hour_index TYPE i.
          lv_hour_index = 144.

          "flag da sheet
          DATA: flag_next_sheet TYPE flag.
          flag_next_sheet = abap_false.

          DATA(lo_reader) = NEW zcl_excel_reader_2007( ).
          DATA(lo_excel)  = lo_reader->zif_excel_reader~load( i_excel2007 = me->lv_xstr ).  "passa o XSTRING carregado

          DATA(i) = 2.

          "itera por todas as sheets do excel, seja ela quantas houverem
          WHILE i <= lo_excel->get_worksheets_size( ).

            "começa a partir da segunda sheet, sendo a primeira a exibicao de dados gerais
            DATA(lo_worksheet) = lo_excel->get_worksheet_by_index( i ).

            "define a coordenada da celula ao inicio da sheet
            lv_str_coord = lv_coord_num.
            lv_str_coord2 = lv_coord_num.
            CONCATENATE lv_coord lv_str_coord INTO lv_str_coord. "A10
            CONDENSE lv_str_coord NO-GAPS.
            CONCATENATE lv_coord2 lv_str_coord2 INTO lv_str_coord2. "A10
            CONDENSE lv_str_coord2 NO-GAPS.

            "pega as informacoes do colaborador
            READ TABLE lo_worksheet->sheet_content REFERENCE INTO DATA(cell) INDEX lv_index. "B2

            "numero do colaborador
            me->st_alv-num = cell->cell_value.

            ADD 2 TO lv_index.

            READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_index. "B2

            me->st_alv-nome = cell->cell_value.

            ADD 2 TO lv_index.

            READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_index. "B2

            me->st_alv-equipa = cell->cell_value.

            ADD 2 TO lv_index.

            READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_index. "B2

            me->st_alv-cntr_cust = cell->cell_value.

            "itera sobre os seis projetos tanto na coluna A quanto na B
            DO 6 TIMES.

              "procura se há peps ativas
              READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell WITH KEY cell_coords = lv_str_coord. "A10...
              READ TABLE lo_worksheet->sheet_content REFERENCE INTO DATA(cell2) WITH KEY cell_coords = lv_str_coord2. "B10...

              "se nao houverem peps selecionadas
              IF cell->cell_value EQ 'Selecione' AND cell2->cell_value EQ 'Selecione'.

                "itera sobre os 31 dias do mes -- valor fixo
                DO 31 TIMES.

                  "verifica as horas trabalhadas
                  READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_hour_index. "E10

                  "caso seja uma letra, passa o valor da celula para maiuscula por precaucao
                  TRANSLATE cell->cell_value TO UPPER CASE.

                  "substitui pontos por virgulas
                  REPLACE
                    ALL OCCURRENCES OF '.'
                    IN cell->cell_value WITH ','.

                  "se houver algum ponto encontrado
                  IF sy-subrc EQ 0.
                    cell->cell_value = 'A'.
                  ENDIF.

                  "se houver hora trabalhada em projeto vazio...
                  IF cell->cell_value NE '0' AND cell->cell_value NE '0,0'.

                    me->st_alv-dia = gv_datemonth.     "recebe o dia do mes

                    "se for uma letra ou palavra...
                    IF cell->cell_value CA sy-abcde.
                      cell->cell_value = '0'.
                      me->st_alv-info = '@05@' && 'Caractere Inválido' .

                      "se a celula nao estiver preenchida
                    ELSEIF cell->cell_value IS INITIAL.
                      cell->cell_value = '0'.
                      me->st_alv-info = '@05@' && 'Célula Vazia' .

                      "se houver hora nao vinculada a projeto / ausencia e presenca
                    ELSE.
                      me->st_alv-hora = cell->cell_value. "recebe a hora trabalhada
                      me->st_alv-info = '@05@' && 'Horário Inválido' .
                    ENDIF.

                    APPEND st_alv TO tt_alv.

                    CLEAR: me->st_alv-dia, me->st_alv-hora.
                    CLEAR: cell->cell_value.
                    CLEAR: cell2->cell_value.

                  ELSEIF cell->cell_value IS INITIAL.

                    me->st_alv-dia = gv_datemonth.     "recebe o dia do mes
                    me->st_alv-hora = '0'. "recebe a hora trabalhada
                    me->st_alv-info = '@05@' && 'Célula Vazia' .
                    APPEND st_alv TO tt_alv.

                    CLEAR: me->st_alv-dia, me->st_alv-hora.
                    CLEAR: cell->cell_value.
                    CLEAR: cell2->cell_value.

                  ENDIF.

                  "substitui pontos por virgulas
                  REPLACE
                    ALL OCCURRENCES OF '.'
                    IN cell->cell_value WITH ','.

                  "se houver algum ponto encontrado
                  IF sy-subrc EQ 0.
                    me->st_alv-hora = '0'.
                  ENDIF.

                  ADD 1 TO lv_hour_index. "incrementa para a proxima hora
                  ADD 1 TO gv_datemonth.  "incrementa para o proximo dia

                ENDDO.

                me->get_month_datafile( ). "reseta data do mes

              ENDIF.

              me->get_month_datafile( ). "reseta data do mes

              "redefine a coordenada para o proximo projeto
              lv_coord = 'A'.
              ADD 1 TO lv_coord_num.
              lv_str_coord = lv_coord_num.
              CONCATENATE lv_coord lv_str_coord INTO lv_str_coord.
              CONDENSE lv_str_coord NO-GAPS.

              "redefine a coordenada para o proximo projeto
              lv_coord2 = 'B'.
              lv_str_coord2 = lv_coord_num.
              CONCATENATE lv_coord2 lv_str_coord2 INTO lv_str_coord2.
              CONDENSE lv_str_coord2 NO-GAPS.

              "redefine o index de horarios conforme coordenada
              CASE lv_str_coord.
                WHEN 'A10'.
                  lv_hour_index = 144.
                WHEN 'A11'.
                  lv_hour_index = 177.
                WHEN 'A12'.
                  lv_hour_index = 210.
                WHEN 'A13'.
                  lv_hour_index = 243.
                WHEN 'A14'.
                  lv_hour_index = 276.
                WHEN 'A15'.
                  lv_hour_index = 309.
              ENDCASE.

              CLEAR: st_alv-dia, st_alv-hora, st_alv-pep.

            ENDDO.

            "-------------------------------------------
            "    redefine dados para proxima sheet.
            "-------------------------------------------

            lv_hour_index = 144. "index de horas trabalhadas

            me->get_month_datafile( ). "reseta data do mes

            "passa para a próxima sheet
            ADD 1 TO i.
            lv_index = 2.

            CLEAR: lv_str_coord. "limpa a coordenada
            CLEAR: lv_str_coord2. "limpa a coordenada
            lv_coord_num = 10. "redefine a linha da coordenada

            CLEAR st_alv. "limpa a estrutura para a proxima sheet

          ENDWHILE.

        ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->INSERT_BAPICATS
* +-------------------------------------------------------------------------------------------------+
* | [--->] TABLE_TIMESHEET                TYPE        ZTSHRALV_TT
* | [<---] RESULT                         TYPE        ZRLA_RESULT
* +--------------------------------------------------------------------------------------</SIGNATURE>
        METHOD insert_bapicats.

          "-------------------------------------------------------
          "info: insere os dados validados na alv no catsdb
          "
          "data de alteracao: 13.11.2024
          "alteracao: criacao do método
          "criado por: rafael albuquerque
          "-------------------------------------------------------

          IF table_timesheet IS INITIAL.
            result-rc = 4.
            result-message = | Não há dados para serem inseridos no cats |.
          ENDIF.

          "Inserir um novo registro em cats
          DATA: it_bapicats1 TYPE TABLE OF bapicats1,
                ls_bapicats1 TYPE bapicats1.

          "tabela e estrutura para retorno de dados
          DATA: it_return TYPE TABLE OF bapiret2,
                ls_return TYPE bapiret2.

          "passa os dados para a estrutura bapicast e insere a tabela interna.
          LOOP AT table_timesheet INTO DATA(ls_timesheet).
            ls_bapicats1-workdate       = ls_timesheet-dia.
            ls_bapicats1-employeenumber = ls_timesheet-num.
            ls_bapicats1-send_cctr      = ls_timesheet-cntr_cust.
            ls_bapicats1-wbs_element    = ls_timesheet-pep.
            ls_bapicats1-catshours      = ls_timesheet-hora.
            APPEND ls_bapicats1 TO it_bapicats1.
            CLEAR ls_bapicats1.
            CLEAR ls_timesheet.
          ENDLOOP.

          IF it_bapicats1 IS NOT INITIAL.

            CALL FUNCTION 'BAPI_CATIMESHEETMGR_INSERT'
              EXPORTING
                testrun        = 'X'
              TABLES
                catsrecords_in = it_bapicats1
                return         = it_return.  " Tabela de mensagens de retorno

            result-rc = 0.

            LOOP AT it_return INTO ls_return.
              IF ls_return-type = 'E'.  " Tipo 'E' indica erro
                result-rc = 4.
                result-message = ls_return-message.
                EXIT.  " Interrompe a execução se houver erros
              ENDIF.
            ENDLOOP.

            IF result-rc EQ 0.
              result-message = | Timesheet submetida com sucesso. |.
            ENDIF.

          ENDIF.

        ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->SET_COORDENATES
* +-------------------------------------------------------------------------------------------------+
* | [--->] LETTER_COORD                   TYPE        STRING
* | [<---] STRING_COORD_A                 TYPE        STRING
* | [<---] STRING_COORD_B                 TYPE        STRING
* | [<---] INDEX_COORD                    TYPE        I
* +--------------------------------------------------------------------------------------</SIGNATURE>
        METHOD set_coordenates.

          "-------------------------------------------------------
          "info: envio da coordenada da coluna A ou B
          "
          "data de alteracao: 15.11.2024
          "alteracao: criacao do método
          "criado por: rafael albuquerque
          "-------------------------------------------------------

          "tratamento da coordenada enviada
          IF letter_coord IS NOT INITIAL.
            DATA: letter_coord2 TYPE string.
            letter_coord2 = letter_coord.
            TRANSLATE letter_coord2 TO UPPER CASE.
          ENDIF.

          "verifica qual coluna
          CASE letter_coord2.
            WHEN 'A'.
              string_coord_a = 'A10'.
              index_coord = 144.
            WHEN 'B'.
              string_coord_b = 'B10'.
              index_coord = 144.
            WHEN OTHERS.
              MESSAGE | Escolha uma coordenada entre A ou B | TYPE 'S' DISPLAY LIKE 'E'.
          ENDCASE.

        ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->SET_DATABASE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
        METHOD set_database.

          "----------------------------------------------------------------------------------------------
          "info: insere as informacoes gerais dos colaboradores na primeira sheet do excel file
          "
          "data de alteracao: 09.11.2024
          "alteracao: criacao do método
          "criado por: rafael albuquerque
          "ultima alteracao: insercao de campo selecione em validacoes de projetos ausencias e presencas
          "----------------------------------------------------------------------------------------------

          DATA(o_xl_ws) = o_xl->get_active_worksheet( ).
          lo_worksheet = o_xl_ws.

          "insere titulo na worksheet.
          DATA: lv_title TYPE zexcel_sheet_title. "titulo de worksheets
          lv_title = 'Colaboradores'.

          TRY.
              lo_worksheet->set_title( ip_title = lv_title ).
            CATCH zcx_excel INTO DATA(lx_excel).
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          DATA: it_stringtable TYPE TABLE OF string, "tabela da dropdown validation
                ls_stringtable TYPE string.

          "index para correr as linhas
          DATA: lv_index TYPE i.
          lv_index = 2.

          "cabeçalho da tabela
          TRY.
              lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = 'Número'          ip_style = tp_style_bold_center_guid ).
              lo_worksheet->set_cell( ip_row = 1 ip_column = 'B' ip_value = 'Colaborador'     ip_style = tp_style_bold_center_guid ).
              lo_worksheet->set_cell( ip_row = 1 ip_column = 'C' ip_value = 'Equipa'          ip_style = tp_style_bold_center_guid ).
              lo_worksheet->set_cell( ip_row = 1 ip_column = 'D' ip_value = 'Centro de Custo' ip_style = tp_style_bold_center_guid ).
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.


          "linhas da tabela
          LOOP AT me->it_colaboradores INTO me->ls_colaborador.

*      =HIPERLIGAÇÃO("#'  00000001 - Colaborador A'!B2"; "Colaborador A") "TRABALHAR COM ISTO!!!!

            DATA: link TYPE string.
            link = | '=HIPERLINK({ me->ls_colaborador-pernr } - { me->ls_colaborador-sname } A!B2)' |.

            TRY.
                lo_worksheet->set_cell( ip_row = lv_index ip_column = 'A' ip_value = ls_colaborador-pernr ip_style = tp_style_bold_center_guid2  ).
                lo_worksheet->set_cell( ip_row = lv_index ip_column = 'B' ip_value = ls_colaborador-sname ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = lv_index ip_column = 'C' ip_value = ls_colaborador-vdsk1 ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = lv_index ip_column = 'D' ip_value = ls_colaborador-kostl ip_style = tp_style_bold_center_guid2 ).
              CATCH zcx_excel INTO lx_excel.
                MESSAGE lx_excel->get_text( ) TYPE 'E'.
            ENDTRY.

            ADD 1 TO lv_index. "incrementa o contador

            "preenche a tabela da dropdown.
            CONCATENATE ls_colaborador-pernr '-' ls_colaborador-sname INTO ls_stringtable SEPARATED BY space.
            APPEND ls_stringtable TO it_stringtable.
            CLEAR: ls_stringtable, me->ls_colaborador.
          ENDLOOP.

          lv_index = 2. "reseta o contador

          "----------------------------------------------------------------------------

          "começa a escrever a tabela da dropdown da lista de colaboradores
          TRY.
              lo_worksheet->set_cell( ip_row = 1 ip_column = 'Z' ip_value = 'Lista de Colaboradores' ip_style = tp_style_bold_center_guid ).
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          LOOP AT it_stringtable INTO ls_stringtable.

            TRY.
                lo_worksheet->set_cell( ip_row = lv_index ip_column = 'Z' ip_value = ls_stringtable  ip_style = tp_style_bold_center_guid2 ).
              CATCH zcx_excel INTO lx_excel.
                MESSAGE lx_excel->get_text( ) TYPE 'E'.
            ENDTRY.

            ADD 1 TO lv_index. "incrementa o contador
          ENDLOOP.

          lv_index = 3. "reseta o contador

          "----------------------------------------------------------------------------

          "começa a escrever a tabela da dropdown de ausencias e presencas

          TRY.
              lo_worksheet->set_cell( ip_row = 1 ip_column = 'AA' ip_value = 'Ausências / Presenças' ip_style = tp_style_bold_center_guid ).
              lo_worksheet->set_cell( ip_row = 2 ip_column = 'AA' ip_value = 'Selecione'             ip_style = tp_style_bold_center_guid ).
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          LOOP AT me->it_line_preaus INTO me->ls_line_preaus.

            TRY.
                lo_worksheet->set_cell( ip_row = lv_index ip_column = 'AA' ip_value = me->ls_line_preaus-line  ip_style = tp_style_bold_center_guid2 ).
              CATCH zcx_excel INTO lx_excel.
                MESSAGE lx_excel->get_text( ) TYPE 'E'.
            ENDTRY.

            ADD 1 TO lv_index. "incrementa o contador
          ENDLOOP.

          lv_index = 3. "reseta o contador

          "----------------------------------------------------------------------------

          "começa a escrever a tabela da dropdown de projetos

          TRY.
              lo_worksheet->set_cell( ip_row = 1 ip_column = 'AB' ip_value = 'Lista de Projetos' ip_style = tp_style_bold_center_guid ).
              lo_worksheet->set_cell( ip_row = 2 ip_column = 'AB' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid ).
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

          LOOP AT me->it_linha_projetos INTO me->ls_linha_projeto.

            TRY.
                lo_worksheet->set_cell( ip_row = lv_index ip_column = 'AB' ip_value = me->ls_linha_projeto-line  ip_style = tp_style_bold_center_guid2 ).
              CATCH zcx_excel INTO lx_excel.
                MESSAGE lx_excel->get_text( ) TYPE 'E'.
            ENDTRY.

            ADD 1 TO lv_index. "incrementa o contador

          ENDLOOP.

          lv_index = 2. "reseta o contador

          "----------------------------------------------------------------------------

          "setup das colunas

          TRY.
              lo_column = lo_worksheet->get_column( ip_column = 'A' ).
              lo_column->set_width( ip_width = 30 ).
              lo_column = lo_worksheet->get_column( ip_column = 'B' ).
              lo_column->set_width( ip_width = 30 ).
              lo_column = lo_worksheet->get_column( ip_column = 'C' ).
              lo_column->set_width( ip_width = 30 ).
              lo_column = lo_worksheet->get_column( ip_column = 'D' ).
              lo_column->set_width( ip_width = 30 ).
              lo_column = lo_worksheet->get_column( ip_column = 'Z' ).
              lo_column->set_width( ip_width = 40 ).
              lo_column = lo_worksheet->get_column( ip_column = 'AA' ).
              lo_column->set_width( ip_width = 50 ).
              lo_column = lo_worksheet->get_column( ip_column = 'AB' ).
              lo_column->set_width( ip_width = 50 ).
            CATCH zcx_excel INTO lx_excel.
              MESSAGE lx_excel->get_text( ) TYPE 'E'.
          ENDTRY.

        ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->SET_RANGEMONTHDATE
* +-------------------------------------------------------------------------------------------------+
* | [<---] BEGIN_DATE                     TYPE        SY-DATUM
* | [<---] END_DATE                       TYPE        SY-DATUM
* +--------------------------------------------------------------------------------------</SIGNATURE>
        METHOD set_rangemonthdate.

          "----------------------------------------------------------------------------------------------
          "info: cria uma range de datas: começo e final de mes
          "
          "data de alteracao: 09.11.2024
          "alteracao: criacao do método
          "criado por: rafael albuquerque
          "----------------------------------------------------------------------------------------------

          "tratamento de data final do mes
          "-------------------------------------------

          DATA: lv_date      TYPE /osp/dt_date, "data enviada
                lv_countdays TYPE /osp/dt_day.  "dias do mes recebidos

          DATA: lv_str_countdays TYPE string. "dias do mes em string

          lv_date = me->gv_datemonth. "data recebe a data enviada pelo programa

          "funcao retorna a quantidade de dias do mes
          CALL FUNCTION '/OSP/GET_DAYS_IN_MONTH'
            EXPORTING
              iv_date = lv_date
            IMPORTING
              ev_days = lv_countdays.

          lv_str_countdays = lv_countdays. "casting >> int to str

          end_date = lv_date+0(6). "recebe o ano e o mês
          begin_date = lv_date+0(6). "recebe o ano e o mês
          CONCATENATE end_date lv_str_countdays INTO end_date. "concatena ano / mes e quantidade de dias do mes
          CONCATENATE begin_date '01' INTO begin_date. "concatena o mês para a data inicial.

        ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->SET_SHEETS
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
        METHOD set_sheets.

          "----------------------------------------------------------------------------------------------
          "info: insere paginacao no excel para cada colaborador
          "
          "data de alteracao: 09.11.2024
          "alteracao: criacao do método
          "criado por: rafael albuquerque
          "----------------------------------------------------------------------------------------------

          DATA: lv_title TYPE zexcel_sheet_title.

          "------------------------------------------------------------------------------------------------------------------------------------------
          "------------------------------------------------------------------------------------------------------------------------------------------

          TRY.
              LOOP AT me->it_colaboradores INTO me->ls_colaborador.

                lv_title = | { me->ls_colaborador-pernr } - { me->ls_colaborador-sname }|. "titulo da sheet recebe o id + nome do colaborador

                "criando uma nova worksheet
                lo_worksheet = o_xl->add_new_worksheet( ).
                lo_worksheet->set_title( ip_title = | { lv_title } | ).

                "nomes dos campos do cabeçalho
                lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = 'N.Mecan:'         ip_style = tp_style_bold_center_guid ).
                lo_worksheet->set_cell( ip_row = 2 ip_column = 'A' ip_value = 'Nome:'            ip_style = tp_style_bold_center_guid ).
                lo_worksheet->set_cell( ip_row = 3 ip_column = 'A' ip_value = 'Equipa:'          ip_style = tp_style_bold_center_guid ).
                lo_worksheet->set_cell( ip_row = 4 ip_column = 'A' ip_value = 'Centro de Custo:' ip_style = tp_style_bold_center_guid ).

                "nomes das linhas do cabeçalho
                lo_worksheet->set_cell( ip_row = 1 ip_column = 'B' ip_value = me->ls_colaborador-pernr ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 2 ip_column = 'B' ip_value = me->ls_colaborador-sname ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 3 ip_column = 'B' ip_value = me->ls_colaborador-vdsk1 ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 4 ip_column = 'B' ip_value = me->ls_colaborador-kostl ip_style = tp_style_bold_center_guid2 ).

                "------------------------------------------------------------------------------------------------------------------------------------------
                "------------------------------------------------------------------------------------------------------------------------------------------

                "calendários do excel

                me->generate_calendar( ). "gerador do calendario do excel.

                "------------------------------------------------------------------------------------------------------------------------------------------
                "------------------------------------------------------------------------------------------------------------------------------------------

                "cabeçalho das horas trabalhadas e planeadas
                lo_worksheet->set_cell( ip_row = 6 ip_column = 'A' ip_value = 'Dia / Mês'         ip_style = tp_style_bold_center_guid ).
                lo_worksheet->set_cell( ip_row = 6 ip_column = 'B' ip_value = me->gv_datemonth    ip_style = tp_style_bold_center_guid ).
                lo_worksheet->set_cell( ip_row = 7 ip_column = 'A' ip_value = 'Horas Planeadas'   ip_style = tp_style_bold_center_guid ).
                lo_worksheet->set_cell( ip_row = 8 ip_column = 'A' ip_value = 'Horas Trabalhadas' ip_style = tp_style_bold_center_guid ).

                "totais de horas trabalhadas e planeadas
                lo_worksheet->set_cell( ip_row = 6  ip_column = 'D' ip_value = 'Totais'                                ip_style = tp_style_bold_center_guid ).
                lo_worksheet->set_cell( ip_row = 7  ip_column = 'D' ip_value = ''       ip_formula = total_planeadas   ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 8  ip_column = 'D' ip_value = ''       ip_formula = total_trabalhadas ip_style = tp_style_bold_center_guid2 ).

                "------------------------------------------------------------------------------------------------------------------------------------------
                "------------------------------------------------------------------------------------------------------------------------------------------

                "horas planeadas e trabalhadas
                lo_worksheet->set_cell( ip_row = 7 ip_column = 'A' ip_value = 'Horas Planeadas'   ip_style = tp_style_bold_center_guid ).
                lo_worksheet->set_cell( ip_row = 8 ip_column = 'A' ip_value = 'Horas Trabalhadas' ip_style = tp_style_bold_center_guid ).

                "------------------------------------------------------------------------------------------------------------------------------------------
                "------------------------------------------------------------------------------------------------------------------------------------------

                "cabeçalho do pep
                lo_worksheet->set_cell( ip_row = 9  ip_column = 'A' ip_value = 'PEP'                 ip_style = tp_style_bold_center_guid  ).
                lo_worksheet->set_cell( ip_row = 9  ip_column = 'B' ip_value = 'Ausência / Presença' ip_style = tp_style_bold_center_guid  ).

                lo_worksheet->set_cell( ip_row = 10 ip_column = 'A' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 11 ip_column = 'A' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 12 ip_column = 'A' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 13 ip_column = 'A' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 14 ip_column = 'A' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 15 ip_column = 'A' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid2 ).

                lo_worksheet->set_cell( ip_row = 10 ip_column = 'B' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 11 ip_column = 'B' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 12 ip_column = 'B' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 13 ip_column = 'B' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 14 ip_column = 'B' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid2 ).
                lo_worksheet->set_cell( ip_row = 15 ip_column = 'B' ip_value = 'Selecione' ip_style = tp_style_bold_center_guid2 ).

                "setup da primeira coluna
                lo_column = lo_worksheet->get_column( ip_column = 'A' ).
                lo_column->set_width( ip_width = 30 ).
                lo_column = lo_worksheet->get_column( ip_column = 'B' ).
                lo_column->set_width( ip_width = 50 ).
                lo_column = lo_worksheet->get_column( ip_column = 'D' ).
                lo_column->set_width( ip_width = 20 ).

                "range de busca para a dropdown de ausencias e presencas
                DATA(lo_range) = o_xl->add_new_range( ).
                lo_range->name = 'AusenciasPresencas'. "nome do range
                lo_range->set_value(
                  ip_sheet_name   = 'Colaboradores' "sheet escolhida
                  ip_start_column = 'AA'
                  ip_start_row    = 2
                  ip_stop_column  = 'AA'
                  ip_stop_row     = lines( me->it_line_preaus ) + 1 "limite do range
                ).

                "range de busca para a dropdown de peps
                lo_range = o_xl->add_new_range( ).
                lo_range->name = 'PEPS'. "nome do range
                lo_range->set_value(
                  ip_sheet_name   = 'Colaboradores' "sheet escolhida
                  ip_start_column = 'AB'
                  ip_start_row    = 2
                  ip_stop_column  = 'AB'
                  ip_stop_row     = lines( me->it_linha_projetos ) + 1 "limite do range
                ).

                "contador para a quantidade de celulas de validacao
                DATA: counter_listboxes TYPE i.
                counter_listboxes = 10.

                DO 6 TIMES.

                  "validacao do range da dropdown
                  lo_data_validation              = lo_worksheet->add_new_data_validation( ).
                  lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
                  lo_data_validation->formula1    = 'AusenciasPresencas'. "nome do range
                  lo_data_validation->cell_row    = counter_listboxes.
                  lo_data_validation->cell_column = 'B'.
                  lo_data_validation->allowblank  = abap_true.

                  "validacao do range da dropdown
                  lo_data_validation2              = lo_worksheet->add_new_data_validation( ).
                  lo_data_validation2->type        = zcl_excel_data_validation=>c_type_list.
                  lo_data_validation2->formula1    = 'PEPS'. "nome do range
                  lo_data_validation2->cell_row    = counter_listboxes.
                  lo_data_validation2->cell_column = 'A'.
                  lo_data_validation2->allowblank  = abap_true.

                  ADD 1 TO counter_listboxes.

                ENDDO.

                "------------------------------------------------------------------------------------------------------------------------------------------
                "------------------------------------------------------------------------------------------------------------------------------------------

                CLEAR: me->ls_colaborador, counter_listboxes.

              ENDLOOP.

            CATCH cx_root INTO DATA(lx_error).
              MESSAGE lx_error->get_text( ) TYPE 'E'.
          ENDTRY.

        ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->SET_STYLE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
        METHOD set_style.

          "----------------------------------------------------------------------------------------------
          "info: criacao de estilos para celulas
          "
          "data de alteracao: 09.11.2024
          "alteracao: criacao do método
          "criado por: rafael albuquerque
          "----------------------------------------------------------------------------------------------

          "cria objetos das bordas
          CREATE OBJECT o_border_dark.
          o_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
          o_border_dark->border_style = zcl_excel_style_border=>c_border_thin.
          CREATE OBJECT o_border_light.
          o_border_light->border_color-rgb = zcl_excel_style_color=>c_gray.
          o_border_light->border_style = zcl_excel_style_border=>c_border_thin.

          "monta o primeiro estilo para a coluna A da paginacao
          CREATE OBJECT me->lo_style.
          lo_style                         = o_xl->add_new_style( ).
          lo_style->font->bold             = abap_true.
          lo_style->font->italic           = abap_false.
          lo_style->font->name             = zcl_excel_style_font=>c_name_arial.
          lo_style->font->scheme           = zcl_excel_style_font=>c_scheme_none.
          lo_style->font->size             = 12.
          lo_style->font->color-rgb        = zcl_excel_style_color=>c_white.
          lo_style->alignment->horizontal  = zcl_excel_style_alignment=>c_horizontal_center.
          lo_style->alignment->horizontal  = zcl_excel_style_alignment=>c_horizontal_center.
          lo_style->borders->allborders    = o_border_light.
          lo_style->fill->filltype         = zcl_excel_style_fill=>c_fill_solid.
          lo_style->fill->fgcolor-rgb      = zcl_excel_style_color=>c_black.
          tp_style_bold_center_guid        = lo_style->get_guid( ). "nao esquecer

          "monta o primeiro estilo para a coluna B da paginacao
          lo_style                         = o_xl->add_new_style( ).
          lo_style->font->bold             = abap_false.
          lo_style->font->italic           = abap_false.
          lo_style->font->name             = zcl_excel_style_font=>c_name_arial.
          lo_style->font->scheme           = zcl_excel_style_font=>c_scheme_none.
          lo_style->font->size             = 12.
          lo_style->font->color-rgb        = zcl_excel_style_color=>c_black.
          lo_style->alignment->horizontal  = zcl_excel_style_alignment=>c_horizontal_center.
          lo_style->alignment->horizontal  = zcl_excel_style_alignment=>c_horizontal_center.
          lo_style->borders->allborders    = o_border_dark.
          tp_style_bold_center_guid2       = lo_style->get_guid( ). "nao esquecer

          "é possível montar vários estilos guid e usar de forma como convém

        ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->SET_WORKSCHEDULE_DATAFILE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
        METHOD set_workschedule_datafile.

          "----------------------------------------------------------------------------------------------
          "info: insere todos os dados do excel numa tabela interna de saida
          "esta tabela interna sera apresentada num report alv
          "
          "data de alteracao: 11.11.2024
          "alteracao: implementacao de ausencias e presenças e documentacao do método
          "criado por: rafael albuquerque
          "----------------------------------------------------------------------------------------------

          "verifica se há colaborador
          IF me->it_employee IS INITIAL.
            MESSAGE | Não há colaboradores disponíveis | TYPE 'S' DISPLAY LIKE 'E'.
          ENDIF.

          CLEAR me->it_timesheet. "

          "index de cada colaborador
          DATA: lv_indexemployee TYPE i.
          lv_indexemployee = 1.

          "flag para interromper o ciclo apos iterar por cada colaborador
          DATA: lv_stopwhile TYPE flag.
          lv_stopwhile = abap_false.

          "---------------------------------------------------------------------------------------
          "TRANSFERÊNCIA DE PROJETOS - AUSENCIAS E PRESENCAS
          "---------------------------------------------------------------------------------------

          "enquanto a flag estiver inativa
          WHILE lv_stopwhile EQ abap_false.

            "interrompe o ciclo depois de contar todos os colaboradores
            IF lv_indexemployee GT lines( me->it_employee ).
              lv_stopwhile = abap_true.
            ENDIF.

            "limpa todas as estruturas
            CLEAR me->ls_employee.
            CLEAR me->ls_peps.
            CLEAR me->ls_timesheet.

            "le cada colaborador a partir do index
            READ TABLE me->it_employee INTO me->ls_employee INDEX lv_indexemployee.

            "quantidade de dias maximos de um mes
            DO 31 TIMES.

              "itera sobre os projetos do colaborador
              LOOP AT me->it_peps INTO me->ls_peps WHERE dia = me->gv_datemonth AND num = me->ls_employee-num.

                MOVE-CORRESPONDING me->ls_employee TO me->ls_timesheet. "preenche a estrutura com os dados do colaborador
                MOVE-CORRESPONDING me->ls_peps TO me->ls_timesheet.     "insere as informacoes do projeto na linha
                me->ls_timesheet-info = '@04@' &&  ' OK' .
                CLEAR: me->ls_timesheet-auspres.                        "limpeza de dados indesejáveis - importante
                APPEND me->ls_timesheet TO me->it_timesheet.            "insere a linha na tabela de output

                "limpa as linhas
                CLEAR me->ls_peps.
                CLEAR me->ls_timesheet.
              ENDLOOP.

              "itera sobre os motivos de ausencia e presenca do colaborador
              LOOP AT me->it_auspres INTO me->ls_auspres WHERE dia = me->gv_datemonth AND num = me->ls_employee-num.

                MOVE-CORRESPONDING me->ls_employee TO me->ls_timesheet. "preenche o campo de ausencia e presenca
                MOVE-CORRESPONDING me->ls_auspres  TO me->ls_timesheet. "insere a ausencia e presenca na estrutura
                CLEAR: me->ls_timesheet-pep.                            "limpa o pep, caso haja
                me->ls_timesheet-info = '@04@' && 'OK' .
                APPEND me->ls_timesheet TO me->it_timesheet.            "insere novo registro

                "limpa as linhas
                CLEAR me->ls_auspres.
                CLEAR me->ls_timesheet.

              ENDLOOP.

              ADD 1 TO me->gv_datemonth. "cada iteracaoq do ciclo ''DO'', incrementamos para o proximo dia.

            ENDDO.

            ADD 1 TO lv_indexemployee. "passa para o proximo colaborador
            me->get_month_datafile( ). "reinicia a data do mês

          ENDWHILE.

          lv_stopwhile = abap_false. "redefine a flag
          lv_indexemployee = 1.      "redefine o index para o primeiro colaborador

          "depois de coletada toda verificacao, é preciso tratar das possibilidades de haverem peps e motivos de ausencia e presenca na mesma linha

          "---------------------------------------------------------------------------------------
          "VERIFICACAO DE AUSENCIAS E PRESENCAS LIGADAS A PROJETOS ATIVOS E POSSIBILIDADES DE HORAS EXTRAS
          "---------------------------------------------------------------------------------------

          DATA: row3 TYPE me->ty_timesheet. "linha a ser inserida
          DATA: it_timesheet2 TYPE TABLE OF me->ty_timesheet. "tabela para receber dados tratados
          it_timesheet2 = me->it_timesheet. "recebe o conteudo original

          "enquanto a flag estiver inativa
          WHILE lv_stopwhile EQ abap_false.

            "interrompe o ciclo depois de contar todos os colaboradores
            IF lv_indexemployee GT lines( me->it_employee ).
              lv_stopwhile = abap_true.
            ENDIF.

            "limpa todas as estruturas
            CLEAR me->ls_employee.
            CLEAR me->ls_peps.
            CLEAR me->ls_timesheet.

            "le cada colaborador a partir do index
            READ TABLE me->it_employee INTO me->ls_employee INDEX lv_indexemployee.

            "itera sobre cada linha da timesheet
            LOOP AT me->it_timesheet INTO DATA(row1) WHERE num EQ ls_employee-num.

              "cada linha é verificada aqui a procura de projetos existentes com ausencias e presencas
              LOOP AT me->it_timesheet INTO DATA(row2) WHERE num EQ ls_employee-num.

                "cada linha da timesheet é comparada com as outras linhas
                "e vemos se existem linhas com as mesmas caracteristicas contendo ausencias e presencas

                IF row1-row EQ row2-row           "verifica coordenadas
                  AND row1-pep IS NOT INITIAL     "verifica se há pep na linha inicial
                  AND row2-auspres IS NOT INITIAL "verifica se há ausencias e presencas na linha de comparacao
                  AND row2-dia EQ row1-dia.       "verifica os dias

                  row3 = row1.                 "linha de apoio recebe os dados da linha original
                  row3-auspres = row2-auspres. "e recebe o motivo de ausencia e presenca

                  "remove a linha que tem apenas o pep sem a ausencia e a linha que tem a ausencia sem o pep
                  DELETE it_timesheet2 WHERE auspres EQ row2-auspres AND row EQ row1-row AND dia EQ row1-dia.
                  DELETE it_timesheet2 WHERE pep EQ row1-pep AND row EQ row1-row AND dia EQ row2-dia.

                  row3-info      = '@1A@' && 'OBS'.
                  APPEND row3 TO it_timesheet2. "insere a nova linha

                  CLEAR: row3, row2, row1.
                  EXIT.
                ENDIF.

              ENDLOOP.

            ENDLOOP.

            ADD 1 TO lv_indexemployee. "passa para o proximo colaborador

          ENDWHILE.

          lv_stopwhile = abap_false. "redefine a flag
          lv_indexemployee = 1.      "redefine o index para o primeiro colaborador

          "limpa a timesheet e recebe a tabela tratada.
          REFRESH me->it_timesheet.
          me->it_timesheet = it_timesheet2.

        ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->SWITCH_COORDENATES
* +-------------------------------------------------------------------------------------------------+
* | [--->] COORDENATE                     TYPE        STRING
* | [<---] STRING_COORD_A                 TYPE        STRING
* | [<---] STRING_COORD_B                 TYPE        STRING
* | [<---] INDEX_COORD                    TYPE        I
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD switch_coordenates.

    IF coordenate IS NOT INITIAL.

      DATA: coordenate2 TYPE string.
      coordenate2 = coordenate.

      CASE coordenate2.
        WHEN 'A10'.
          string_coord_b = 'A11'.
          index_coord    = 177.
        WHEN 'A11'.
          string_coord_b = 'A12'.
          index_coord    = 210.
        WHEN 'A12'.
          string_coord_b = 'A13'.
          index_coord    = 243.
        WHEN 'A13'.
          string_coord_b = 'A14'.
          index_coord    = 276.
        WHEN 'A14'.
          string_coord_b = 'A15'.
          index_coord    = 309.
        WHEN 'A15'.
          string_coord_b = 'A10'.
          index_coord    = 144.
        WHEN 'B10'.
          string_coord_b = 'B11'.
          index_coord    = 177.
        WHEN 'B11'.
          string_coord_b = 'B12'.
          index_coord    = 210.
        WHEN 'B12'.
          string_coord_b = 'B13'.
          index_coord    = 243.
        WHEN 'B13'.
          string_coord_b = 'B14'.
          index_coord    = 276.
        WHEN 'B14'.
          string_coord_b = 'B15'.
          index_coord    = 309.
        WHEN 'B15'.
          string_coord_b = 'B10'.
          index_coord    = 144.
        WHEN OTHERS.
          MESSAGE | Envie uma coordenada válida entre A10-A15 ou B10-B15 | TYPE 'S' DISPLAY LIKE 'E'.
        ENDCASE.

      ENDIF.

    ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->UPLOAD_TIMESHEET
* +-------------------------------------------------------------------------------------------------+
* | [--->] STR_PATH_FILE                  TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
        METHOD upload_timesheet.

          "----------------------------------------------------------------------------------------------
          "info: carrega o arquivo excel para o programa e invoca todos os metodos envolvidos na leitura
          "do arquivo
          "
          "data de alteracao: 09.11.2024
          "alteracao: criacao do método
          "criado por: rafael albuquerque
          "----------------------------------------------------------------------------------------------

          IF str_path_file IS INITIAL.
            MESSAGE | O Arquivo precisa de um caminho. | TYPE 'S' DISPLAY LIKE 'E'.
            RETURN.
          ELSE.

            "carrega o arquivo em uma tabela binária
            cl_gui_frontend_services=>gui_upload(
              EXPORTING
                filename                = str_path_file  " Nome do arquivo
                filetype                = 'BIN'          " Tipo de arquivo como binário
              CHANGING
                data_tab                = lt_bin_data    " Tabela binária para dados
              EXCEPTIONS
                file_open_error         = 1
                file_read_error         = 2
                no_batch                = 3
                gui_refuse_filetransfer = 4
                invalid_type            = 5
                no_authority            = 6
                unknown_error           = 7
                bad_data_format         = 8
                header_not_allowed      = 9
                separator_not_allowed   = 10
                header_too_long         = 11
                unknown_dp_error        = 12
                access_denied           = 13
                dp_out_of_memory        = 14
                disk_full               = 15
                dp_timeout              = 16
                not_supported_by_gui    = 17
                error_no_gui            = 18
                OTHERS                  = 19
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

          ENDIF.

          "----------------------------------------------------------------------------------------
          "         metodos envolvidos no upload para programama ZHCM_EXCEL_UPLOAD.
          "----------------------------------------------------------------------------------------
          me->get_employee_datafile( ).     "recebe os colaboradores do arquivo
          me->get_month_datafile( ).        "recebe o mês do arquivo
          me->get_peps_datafile( ).         "recebe os projetos dos colaboradores
          me->get_auspres_datafile( ).      "recebe os motivos de ausencia e presenca dos colaboradores
          me->set_workschedule_datafile( ). "recebe todos os dados dos colaboradores no excel
          "----------------------------------------------------------------------------------------

        ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->VALIDATION_DATAFILE
* +-------------------------------------------------------------------------------------------------+
* | [--->] TABLE_TIMESHEET                TYPE        ZTSHRALV_TT
* | [<---] TABLE_TIMESHEET_OUTPUT         TYPE        ZTSHRALV_TT
* +--------------------------------------------------------------------------------------</SIGNATURE>
        METHOD validation_datafile.

          "----------------------------------------------------------------------------------------------
          "info: validacao final para os dados dos colaboradores da timesheet
          "a validacao é a comparacao do que está no documento com que o que existe na base de dados
          "
          "data de alteracao: 13.11.2024
          "alteracao: criacao do método
          "criado por: rafael albuquerque
          "----------------------------------------------------------------------------------------------

          IF table_timesheet IS INITIAL.
            MESSAGE | Não há dados a serem verificados | TYPE 'S' DISPLAY LIKE 'E'.
            RETURN.
          ENDIF.

          DATA: table_timesheet2 TYPE ztshralv_tt.
          table_timesheet2 = table_timesheet.
          DATA: lv_endda TYPE endda.
          lv_endda = '99991231'.

          "recebimento dos dados para validacao
          "-----------------------------------------------------------------------------------------
          me->get_projects( ). "recebe os projetos coletados da base de dados
          me->get_auspres( ).  "recebe os motivos de ausencia e presenca coletados da base de dados
          SELECT pernr, vdsk1  FROM pa0001 INTO TABLE @DATA(it_equipas_2).   "recebe as equipas
            SELECT pernr, kostl  FROM pa0001 INTO TABLE @DATA(it_centrocusto). "recebe os centros de custo
              "-----------------------------------------------------------------------------------------

              "itera sobre a tabela da timesheet
              LOOP AT table_timesheet2 INTO DATA(ls_timesheet).

                "-------------------------------------------------------------------------------
                "                        verifica o numero do colaborador
                "-------------------------------------------------------------------------------

                DATA: lv_name TYPE emnam.

                "funcao que retorna o nome do colaborador baseado no numero recebido
                CALL FUNCTION 'HR_TMW_GET_EMPLOYEE_NAME'
                  EXPORTING
                    person_no = ls_timesheet-num
                  IMPORTING
                    edit_name = lv_name.

                "se o nome foi achado
                IF lv_name IS INITIAL.
                  ls_timesheet-info = me->st_alv-info = '@05@' && 'Número de Colaborador não existe.' .
                  MODIFY table_timesheet2 FROM ls_timesheet.
                ELSE.

                  "-------------------------------------------------------------------------------
                  "                        verifica a equipa do colaborador
                  "-------------------------------------------------------------------------------

                  READ TABLE it_equipas_2 INTO DATA(ls_equipa_2) WITH KEY pernr = ls_timesheet-num vdsk1 = ls_timesheet-equipa.

                  IF sy-subrc NE 0.
                    ls_timesheet-info      = me->st_alv-info = '@05@' && 'Equipa não existe' .
                    MODIFY table_timesheet2 FROM ls_timesheet.
                  ELSE.

                    "-------------------------------------------------------------------------------
                    "                       verifica se o projeto existe
                    "-------------------------------------------------------------------------------

                    IF ls_timesheet-pep IS NOT INITIAL.
                      READ TABLE me->it_linha_projetos INTO DATA(ls_projetos) WITH KEY line = ls_timesheet-pep.
                      IF sy-subrc NE 0.
                        ls_timesheet-info = me->st_alv-info = '@05@' && 'Projeto Inexistente' .
                        MODIFY table_timesheet2 FROM ls_timesheet.
                      ENDIF.
                    ENDIF.

                    "-------------------------------------------------------------------------------
                    "             verifica se o motivo de ausencia e presenca existe
                    "-------------------------------------------------------------------------------

                    IF ls_timesheet-auspres IS NOT INITIAL.
                      READ TABLE me->it_line_preaus INTO DATA(ls_auspres) WITH KEY line = ls_timesheet-auspres.
                      IF sy-subrc NE 0.
                        ls_timesheet-info = me->st_alv-info = '@05@' && 'Código de Ausência Inexistente' .
                        MODIFY table_timesheet2 FROM ls_timesheet.
                      ENDIF.
                    ENDIF.

                    "-------------------------------------------------------------------------------
                    "                          verifica o centro de custo
                    "-------------------------------------------------------------------------------

                    IF ls_timesheet-cntr_cust IS NOT INITIAL.

                      DATA: lv_centro TYPE kostl.

                      "funcao para converter o centro de custo com o 0 a esquerda
                      CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
                        EXPORTING
                          input  = ls_timesheet-cntr_cust
                        IMPORTING
                          output = lv_centro.

                      READ TABLE it_centrocusto INTO DATA(ls_centrocusto) WITH KEY kostl = lv_centro.
                      IF sy-subrc NE 0.
                        ls_timesheet-info      = me->st_alv-info = '@05@' && 'Centro de Custo não Existe' .
                        MODIFY table_timesheet2 FROM ls_timesheet.
                      ENDIF.
                    ENDIF.

                  ENDIF.
                ENDIF.
              ENDLOOP.

              table_timesheet_output = table_timesheet2.

            ENDMETHOD.
ENDCLASS.
