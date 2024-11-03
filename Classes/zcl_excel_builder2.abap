class ZCL_EXCEL_BUILDER2 definition
  public
  final
  create public  .

public section.

  types:
    "types de informacoes de colaboradores
    BEGIN OF wa_col,
             pernr TYPE pa0001-pernr, "Número Pessoal
             sname TYPE pa0002-cname, "Nome
             vdsk1 TYPE pa0001-vdsk1, "Chave de Organizacao
             kostl TYPE pa0001-kostl, "Centro de Custo
           END OF wa_col .
  types:
    "types de ausencias de precensas
    BEGIN OF wa_pre_aus,
             subty TYPE awart,  "t554s-subty Tipos de presença e ausência
             atext TYPE abwtxt, "t554t-atext Textos de ausência e presença
           END OF wa_pre_aus .
  types:
    "linha unica para guardar ausencia e presença
    BEGIN OF wa_line_preaus,
             line TYPE string,
           END OF wa_line_preaus .
  types:
    "types para projetos abertos
    BEGIN OF wa_project,
             objnr TYPE j_objnr,
             pspid TYPE ps_pspid,
             post1 TYPE ps_post1,
           END OF wa_project .
  types:
    "linha unica para guardar projetos
    BEGIN OF wa_line_projects,
             line TYPE string,
           END OF wa_line_projects .

  data: gv_datemonth type sy-datum.

  data:
    "informacoes dos colaboradores
    it_colaboradores TYPE TABLE OF wa_col .
  data LS_COLABORADOR type WA_COL .
  data TT_COLABORADORES type ZCOL_TT .
  data ST_COLABORADOR type ZCOL_ST .
  data:
    "informacoes de ausencia e presenca
    it_aus_pre TYPE TABLE OF wa_pre_aus .
  data LS_AUS_PRE type WA_PRE_AUS .
  data:
    "linha de ausencia e presenca concatenada
    it_line_preaus TYPE TABLE OF wa_line_preaus .
  data LS_LINE_PREAUS type WA_LINE_PREAUS .
    "celula de horas trabalhadas e planeadas
  data TOTAL_PLANEADAS type STRING .
  data TOTAL_TRABALHADAS type STRING .
  data:
    "tabela e estrutura de projetos abertos
    it_projetos TYPE TABLE OF wa_project .
  data LS_PROJETOS type WA_PROJECT .
  data:
    "tabela de linha concatenada de projetos
    it_linha_projetos TYPE TABLE OF wa_line_projects .
  data LS_LINHA_PROJETO type WA_LINE_PROJECTS .
  data E_RESULT type ZRLA_RESULT .
    "objetos de construcao de arquivos excel
  data O_XL type ref to ZCL_EXCEL .
  data LO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET .
    "objetos de componentes do excel
  data LO_COLUMN type ref to ZCL_EXCEL_COLUMN .
  data LO_DATA_VALIDATION type ref to ZCL_EXCEL_DATA_VALIDATION .
  data LO_DATA_VALIDATION2 type ref to ZCL_EXCEL_DATA_VALIDATION .
  data LO_RANGE type ref to ZCL_EXCEL_RANGE .
  data O_CONVERTER type ref to ZCL_EXCEL_CONVERTER .
  data LO_STYLE type ref to ZCL_EXCEL_STYLE .
  data O_BORDER_DARK type ref to ZCL_EXCEL_STYLE_BORDER .
  data O_BORDER_LIGHT type ref to ZCL_EXCEL_STYLE_BORDER .
  data TP_STYLE_BOLD_CENTER_GUID type ZEXCEL_CELL_STYLE .
  data TP_STYLE_BOLD_CENTER_GUID2 type ZEXCEL_CELL_STYLE .
  data OL_HYPERLINK type ref to ZCL_EXCEL_HYPERLINK .

  methods GET_DATA
    importing
      !COLABORADORES type ZCOL_TT.
  methods DOWNLOAD_XLS .
  methods DISPLAY_FAST_EXCEL
    importing
      !I_TABLE_CONTENT type ref to DATA
      !I_TABLE_NAME type STRING .
  methods get_date
    importing date type sy-datum.
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
ENDCLASS.



CLASS ZCL_EXCEL_BUILDER2 IMPLEMENTATION.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->APPEND_EXTENSION
* +-------------------------------------------------------------------------------------------------+
* | [--->] OLD_EXTENSION                  TYPE        STRING
* | [<---] NEW_EXTENSION                  TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD append_extension.

    CONCATENATE old_extension 'xlsx' INTO new_extension SEPARATED BY '.'.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->CONVERT_EXCEL_COLUMN
* +-------------------------------------------------------------------------------------------------+
* | [--->] COLUMN_INT                     TYPE        I
* | [<---] COLUMN_STRING                  TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD convert_excel_column.

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

    "-------------------------------------------------------------------------------
    "recebe uma tabela generica
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
    "converte para xstring
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
    "conta a quantidade de colunas da tabela
    "-------------------------------------------------------------------------------

    DATA: lo_table_descr  TYPE REF TO cl_abap_tabledescr,
          lo_struct_descr TYPE REF TO cl_abap_structdescr.

    lo_table_descr ?= cl_abap_tabledescr=>describe_by_data( p_data = <lit_table> ).
    lo_struct_descr ?= lo_table_descr->get_table_line_type( ).

    DATA(lv_number_of_columns) = lines( lo_struct_descr->components ).

    "-------------------------------------------------------------------------------
    "setup das colunas - largura
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
    "caminho para o arquivo
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
    "escritor para arquivo
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
    "download do arquivo Excel
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
    "tratamento de erros
    "-------------------------------------------------------------------------------

    IF sy-subrc NE 0.
      MESSAGE 'Não foi possível realizar o download do arquivo' TYPE 'S' DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->DOWNLOAD_XLS
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD download_xls.

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
      MESSAGE 'Não foi possível realizar o download do arquivo' TYPE 'S' DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GENERATE_CALENDAR
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD generate_calendar.

    "metodo para capturar a data enviada pelo programa
    me->get_date( date = gv_datemonth ).

    "verificacao para envio de data
    if gv_datemonth is initial.
      message | 'Para impressão do calendário é preciso a data' | type 'S' display like 'E'.
    endif.

    "buscando a quantidade de dias no mes
    DATA: lv_date          TYPE /osp/dt_date, "data enviada
          lv_countdays     TYPE /osp/dt_day,  "dias do mes recebidos
          lv_countdays2    TYPE i,            "dias do mes em inteiro
          lv_counterdays   TYPE i,            "contador de dias
          lv_newdate       TYPE sy-datum,     "nova data formatada
          lv_stringdaydate TYPE string,       "dia formatado
          lv_day           TYPE i,            "dia em inteiro
          lv_strday        TYPE string.       "dia em string

    "rever as horas trabalhadas conforme consulta - aguardar info adicional
    DATA: horas_planeadas TYPE p DECIMALS 2.
    horas_planeadas = '8.00'.

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

    "reseta a data
    lv_newdate = lv_date+0(6). "recebe ano + mes
    lv_strday = '01'.          "sempre começamos pelo primeiro dia do mes

    "junta ano + mes e primeiro dia do mes
    CONCATENATE lv_newdate lv_strday INTO lv_newdate.

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
          horas_planeadas = '8'.
        ENDIF.

        "cria a celula
        TRY.
            lo_worksheet->set_cell( ip_row = 6 ip_column = lv_counterdays ip_value = lv_stringdaydate ip_style = tp_style_bold_center_guid ).  "cabeçalho do calendário
            lo_worksheet->set_cell( ip_row = 7 ip_column = lv_counterdays ip_value = horas_planeadas  ip_style = tp_style_bold_center_guid2 ). "horas planeadas
          CATCH zcx_excel INTO DATA(lx_excel).
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        CLEAR lv_lettercollum. "limpa a letra da coluna para evitar concatenacoes

        "converte o numero da coluna em string da coluna
        TRY.
            me->convert_excel_column(
              EXPORTING
                column_int    = lv_counterdays
              IMPORTING
                column_string = lv_lettercollum
            ).
          CATCH zcx_excel INTO lx_excel.
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        "celula que recebe a formula da soma de horas trabalhadas no dia
        CLEAR form_dia_trab.
        form_dia_trab = '=SUM(' && lv_lettercollum && '10:' && lv_lettercollum && '15)'. "atualiza a formula

        TRY.
            lo_worksheet->set_cell( ip_row = 8 ip_column = lv_counterdays ip_value = '' ip_formula = form_dia_trab ip_style = tp_style_bold_center_guid2 ). "horas trabalhadas
          CATCH zcx_excel INTO lx_excel.
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        "cabeçalho de tempo trabalhado ou ausentado
        TRY.
            lo_worksheet->set_cell( ip_row = 9 ip_column = lv_counterdays ip_value = 'Tempo' ip_style = tp_style_bold_center_guid ). "horas trabalhadas
            "colunas de tempo de trabalho
            lo_worksheet->set_cell( ip_row = 10 ip_column = lv_counterdays ip_value = 0 ip_style = tp_style_bold_center_guid2 ).
            lo_worksheet->set_cell( ip_row = 11 ip_column = lv_counterdays ip_value = 0 ip_style = tp_style_bold_center_guid2 ).
            lo_worksheet->set_cell( ip_row = 12 ip_column = lv_counterdays ip_value = 0 ip_style = tp_style_bold_center_guid2 ).
            lo_worksheet->set_cell( ip_row = 13 ip_column = lv_counterdays ip_value = 0 ip_style = tp_style_bold_center_guid2 ).
            lo_worksheet->set_cell( ip_row = 14 ip_column = lv_counterdays ip_value = 0 ip_style = tp_style_bold_center_guid2 ).
            lo_worksheet->set_cell( ip_row = 15 ip_column = lv_counterdays ip_value = 0 ip_style = tp_style_bold_center_guid2 ).
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

        ADD 1 TO lv_counterdays. "incrementa o contador para a proxima coluna

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

      ENDIF.

    ENDDO.

    lv_counterdays = 5. "reseta o contador para a 5th coluna
    CLEAR: lv_day, lv_strday. "limpa os contadores de dias em string e int.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_AUSPRES
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD get_auspres.

    "consulta para obter textos de ausencia e presenca
    "-------------------------------------------

    SELECT t554s~subty,               "tipo de ausência e presenca
           t554t~atext                "Texto descritivo
      FROM t554s                      "Da tabela do de Tipos de presença e ausência
      INNER JOIN t554t                "Junta da tabela de Textos de ausência e presença
      ON t554s~moabw = t554t~moabw    "Juntas por chave de agrupamento em RH
      INTO TABLE @me->it_aus_pre
      WHERE t554s~moabw EQ 19
      AND   t554t~moabw EQ 19
      AND   t554t~sprsl EQ @sy-langu  "Onde o idioma for aquele do sistema
      AND   t554t~atext NE ''         "O texto não esteja vazio
      and   t554s~endda GT @sy-datum. "E a data fim seja maior do que a data final

    "formacao da linha de textos para ausencia e presenca
    "----------------------------------------------------

    DATA stringline TYPE string.

    "itera sobre a tabela de textos concatenando Tipos de presença e ausência com os Textos de ausência e presença
    LOOP AT me->it_aus_pre INTO me->ls_aus_pre.
      stringline = me->ls_aus_pre-subty. "casting do numero
      CONCATENATE stringline me->ls_aus_pre-atext INTO me->ls_line_preaus-line SEPARATED BY ' - '.
      APPEND me->ls_line_preaus TO me->it_line_preaus.
    ENDLOOP.

    "verifica se algum dado foi enviado
    IF me->it_line_preaus IS INITIAL.
      MESSAGE | Não foi possível receber os dados da base de dados | TYPE 'S' DISPLAY LIKE 'E'.
    ENDIF.

    CLEAR stringline.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->GET_DATA
* +-------------------------------------------------------------------------------------------------+
* | [--->] COLABORADORES                  TYPE        ZCOL_TT
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD get_data.

    DATA stringline TYPE string.

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

    me->it_colaboradores = colaboradores. "recebe uma tabela interna e preenche o atributo de classe

    "verifica se algum dado foi enviado
    IF colaboradores IS INITIAL.
      MESSAGE | Não foi possível receber os dados da base de dados | TYPE 'S' DISPLAY LIKE 'E'.
    ENDIF.

    "consulta para obter textos de ausencia e presenca
    "-------------------------------------------

    me->get_auspres( ).

    "consulta para obter textos de projetos
    "---------------------------------------

    me->get_projects( ).


  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->GET_DATE
* +-------------------------------------------------------------------------------------------------+
* | [--->] DATE                           TYPE        SY-DATUM
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method get_date.

    me->gv_datemonth = date.

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_FILE_DIRECTORY
* +-------------------------------------------------------------------------------------------------+
* | [--->] FILENAME                       TYPE        STRING
* | [<---] FULL_PATH                      TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD get_file_directory.

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
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_PROJECTS
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD get_projects.

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
* | Instance Private Method ZCL_EXCEL_BUILDER2->SET_DATABASE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD set_database.

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

      data: link type string.
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

    lv_index = 2. "reseta o contador

    "----------------------------------------------------------------------------

    "começa a escrever a tabela da dropdown de ausencias e presencas

    TRY.
        lo_worksheet->set_cell( ip_row = 1 ip_column = 'AA' ip_value = 'Ausências / Presenças' ip_style = tp_style_bold_center_guid ).
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

    lv_index = 2. "reseta o contador

    "----------------------------------------------------------------------------

    "começa a escrever a tabela da dropdown de projetos

    TRY.
        lo_worksheet->set_cell( ip_row = 1 ip_column = 'AB' ip_value = 'Lista de Projetos' ip_style = tp_style_bold_center_guid ).
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

    lv_index = 2. "reseta o contador

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->SET_SHEETS
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD set_sheets.

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
            ip_stop_row     = lines( me->it_aus_pre ) + 1 "limite do range
          ).

          "range de busca para a dropdown de peps
          lo_range = o_xl->add_new_range( ).
          lo_range->name = 'PEPS'. "nome do range
          lo_range->set_value(
            ip_sheet_name   = 'Colaboradores' "sheet escolhida
            ip_start_column = 'AB'
            ip_start_row    = 2
            ip_stop_column  = 'AB'
            ip_stop_row     = lines( me->it_colaboradores ) + 1 "limite do range
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
ENDCLASS.
