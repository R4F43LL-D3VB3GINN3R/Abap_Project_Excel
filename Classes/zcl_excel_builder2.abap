class zcl_excel_builder2 definition
  public
  final
  create public .

  public section.

    types: begin of wa_col,
             pernr type pa0001-pernr, "Número Pessoal
             sname type pa0002-cname, "Nome
             vdsk1 type pa0001-vdsk1, "Chave de Organizacao
             kostl type pa0001-kostl, "Centro de Custo
           end of wa_col.
    data:
      it_colaboradores type table of wa_col,
      ls_colaborador   type wa_col.

    data: tt_colaboradores type zcol_tt,
          st_colaborador   type zcol_st.

    data: total_planeadas type string,
          total_trabalhadas type string.

    data e_result type zrla_result .

    data: o_xl               type ref to zcl_excel,
          lo_worksheet       type ref to zcl_excel_worksheet,
          lo_column          type ref to zcl_excel_column,
          lo_data_validation type ref to zcl_excel_data_validation,
          lo_range           type ref to zcl_excel_range,
          o_converter        type ref to zcl_excel_converter.

    data: lo_style                   type ref to zcl_excel_style,
          o_border_dark              type ref to zcl_excel_style_border,
          o_border_light             type ref to zcl_excel_style_border,
          tp_style_bold_center_guid  type zexcel_cell_style,
          tp_style_bold_center_guid2 type zexcel_cell_style.

    methods get_data
      exporting
        !colaboradores type zcol_tt
        !e_result      type zrla_result .
    methods download_xls .
    methods display_fast_excel
      importing
        !i_table_content type ref to data
        !i_table_name    type string .

  protected section.
  private section.
    methods convert_xstring .
    methods set_database .
    methods append_extension
      importing
        !old_extension type string
      exporting
        !new_extension type string .
    methods get_file_directory
      importing
        !filename  type string
      exporting
        !full_path type string .
    methods set_style .
    methods set_sheets .
    methods generate_calendar.

ENDCLASS.



CLASS ZCL_EXCEL_BUILDER2 IMPLEMENTATION.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->APPEND_EXTENSION
* +-------------------------------------------------------------------------------------------------+
* | [--->] OLD_EXTENSION                  TYPE        STRING
* | [<---] NEW_EXTENSION                  TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method append_extension.

    concatenate old_extension 'xlsx' into new_extension separated by '.'.

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->CONVERT_XSTRING
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method convert_xstring.

    data: lx_error      type ref to cx_root,  "define uma referência para exceções
          lv_error_text type string.          "define uma variável para o texto do erro

    try.
        "cria o objeto para o conversor
        create object o_converter.

        "converte os dados para o formato Excel
        o_converter->convert(
          exporting
            it_table = me->it_colaboradores
          changing
            co_excel = me->o_xl
        ).

        "verificação de erros na conversão
        if sy-subrc ne 0.
          message 'Não foi possível converter os dados para xstring' type 'S' display like 'E'.
          return.
        endif.

      catch cx_root into lx_error.
        lv_error_text = lx_error->if_message~get_text( ).
        message lv_error_text type 'S' display like 'E'.
        return.
    endtry.


  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->DISPLAY_FAST_EXCEL
* +-------------------------------------------------------------------------------------------------+
* | [--->] I_TABLE_CONTENT                TYPE REF TO DATA
* | [--->] I_TABLE_NAME                   TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method display_fast_excel.

    "-------------------------------------------------------------------------------
    "recebe uma tabela generica
    "-------------------------------------------------------------------------------

    " Tipo de dados generico
    data: lr_table type ref to data.

    " Instanciar esse tipo de dados em runtime para ser uma tabela do tipo (i_table_name)
    create data lr_table type table of (i_table_name).

    " Preencher a tabela do método com o conteudo que vem no parametro
    lr_table = i_table_content.

    " Como foi criada por referência ao tipo genérico "data" não dá para aceder diretamente
    " Usar field symbol e apontar o conteudo da tabela (->*) para o field symbol
    field-symbols: <lit_table> type any table.
    assign lr_table->* to <lit_table>.

    create object o_xl. "cria objeto excel
    create object o_converter.

    "-------------------------------------------------------------------------------
    "converte para xstring
    "-------------------------------------------------------------------------------

    data: lx_error      type ref to cx_root,       "define uma referência para exceções
          lv_error_text type string.          "define uma variável para o texto do erro

    try.
        "converte os dados para o formato Excel
        o_converter->convert(
          exporting
            it_table      = <lit_table>
          changing
            co_excel      = me->o_xl
        ).

        " Verificação de erros na conversão
        if sy-subrc ne 0.
          message 'Não foi possível converter os dados para xstring' type 'S' display like 'E'.
          return.
        endif.

      catch cx_root into lx_error.
        lv_error_text = lx_error->if_message~get_text( ).
        message lv_error_text type 'S' display like 'E'.
        return.
    endtry.

    "cria um worksheet
    data(o_xl_ws) = o_xl->get_active_worksheet( ).
    lo_worksheet = o_xl_ws.

    "-------------------------------------------------------------------------------
    "conta a quantidade de colunas da tabela
    "-------------------------------------------------------------------------------

    data: lo_table_descr  type ref to cl_abap_tabledescr,
          lo_struct_descr type ref to cl_abap_structdescr.

*     Use RTTI services to describe table variable
    lo_table_descr ?= cl_abap_tabledescr=>describe_by_data( p_data = <lit_table> ).
*     Use RTTI services to describe table structure
    lo_struct_descr ?= lo_table_descr->get_table_line_type( ).

*     Count number of columns in structure
    data(lv_number_of_columns) = lines( lo_struct_descr->components ).

    "-------------------------------------------------------------------------------
    "setup das colunas - largura
    "-------------------------------------------------------------------------------

    me->set_style( ). "insere o estilo na coluna

    "contador de colunas
    data: count_columns type i.
    count_columns = 1. "começa pela primeira

    "conta até a quantidade de colunas da tabela
    do lv_number_of_columns times.
      lo_column = lo_worksheet->get_column( ip_column = count_columns ).                 "pega a coluna relativo ao index
*      lo_column->set_column_style_by_guid( ip_style_guid = tp_style_bold_center_guid2 ). "insere o estilo na coluna
      lo_column->set_width( ip_width = 30 ).                                             "insere o tamanho da coluna
      add 1 to count_columns.
    enddo.

    count_columns = 1. "reseta o contador

    "titulo do worksheet principal
    data(worksheet_title) = conv zexcel_sheet_title( |{ i_table_name }| ).
    lo_worksheet->set_title( ip_title = worksheet_title ).

    "-------------------------------------------------------------------------------
    "caminho para o arquivo
    "-------------------------------------------------------------------------------

    "tratamento de nome e extensão do arquivo
    data full_path type string.
    data namefile type string.

    namefile = 'file'.

    "metodo que salva nome e diretorio
    me->get_file_directory(
      exporting
        filename  = namefile
      importing
        full_path = full_path
    ).

    "se o download for cancelado...
    if full_path is initial.
      message 'O download foi cancelado pelo usuário.' type 'S' display like 'E'.
      return.
    endif.

    "-------------------------------------------------------------------------------
    "escritor para arquivo
    "-------------------------------------------------------------------------------

    "inicia o escritor do arquivo
    data(o_xlwriter)  = cast zif_excel_writer( new zcl_excel_writer_2007( ) ).
    data(lv_xl_xdata) = o_xlwriter->write_file( o_xl ).
    data(it_raw_data) = cl_bcs_convert=>xstring_to_solix( exporting iv_xstring = lv_xl_xdata ).

    "-------------------------------------------------------------------------------
    "download do arquivo Excel
    "-------------------------------------------------------------------------------

    try.
        cl_gui_frontend_services=>gui_download(
          exporting
            filename     = full_path
            filetype     = 'BIN'
            bin_filesize = xstrlen( lv_xl_xdata )
          changing
            data_tab     = it_raw_data
        ).
      catch cx_root into lx_error.
        lv_error_text = lx_error->if_message~get_text( ).
        message lv_error_text type 'S' display like 'E'.
        return.
    endtry.

    "-------------------------------------------------------------------------------
    "tratamento de erros
    "-------------------------------------------------------------------------------

    if sy-subrc ne 0.
      message 'Não foi possível realizar o download do arquivo' type 'S' display like 'E'.
      return.
    endif.

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->DOWNLOAD_XLS
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method download_xls.

    "tratamento de nome e extensão do arquivo
    data full_path type string.
    data namefile type string.

    namefile = 'file'.

    "metodo que salva nome e diretorio
    me->get_file_directory(
      exporting
        filename  = namefile
      importing
        full_path = full_path
    ).

    "se o download for cancelado...
    if full_path is initial.
      message 'O download foi cancelado pelo usuário.' type 'S' display like 'E'.
      return.
    endif.

    "----------------------------------------------------------------

    create object o_xl. "cria objeto excel

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
    data(o_xlwriter) = cast zif_excel_writer( new zcl_excel_writer_2007( ) ).
    data(lv_xl_xdata) = o_xlwriter->write_file( o_xl ).
    data(it_raw_data) = cl_bcs_convert=>xstring_to_solix( exporting iv_xstring = lv_xl_xdata ).

    "----------------------------------------------------------------

    "download do arquivo Excel
    try.
        cl_gui_frontend_services=>gui_download(
          exporting
            filename     = full_path
            filetype     = 'BIN'
            bin_filesize = xstrlen( lv_xl_xdata )
          changing
            data_tab     = it_raw_data
        ).
      catch cx_root into data(ex_txt).
        write: / ex_txt->get_text( ).
    endtry.

    "----------------------------------------------------------------

    "tratamento de erros
    if sy-subrc ne 0.
      message 'Não foi possível realizar o download do arquivo' type 'S' display like 'E'.
      return.
    endif.

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GENERATE_CALENDAR
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method generate_calendar.

    "buscando a quantidade de dias no mes
    data: lv_date          type /osp/dt_date, "data enviada
          lv_countdays     type /osp/dt_day,  "dias do mes recebidos
          lv_countdays2    type i,            "dias do mes em inteiro
          lv_counterdays   type i,            "contador de dias
          lv_newdate       type sy-datum,     "nova data formatada
          lv_stringdaydate type string,       "dia formatado
          lv_day           type i,            "dia em inteiro
          lv_strday        type string.       "dia em string

    lv_date = sy-datum. "recebe a data atual

    "funcao retorna a quantidade de dias do mes
    call function '/OSP/GET_DAYS_IN_MONTH'
      exporting
        iv_date = lv_date
      importing
        ev_days = lv_countdays.

    lv_countdays2 = lv_countdays. "casting int
    lv_counterdays = 5.           "inicia o contador como cinco para contar a partir da 5th coluna

    "-------------------------------------------

    "reseta a data
    lv_newdate = lv_date+0(6). "recebe ano + mes
    lv_strday = '01'.          "sempre começamos pelo primeiro dia do mes

    "junta ano + mes e primeiro dia do mes
    concatenate lv_newdate lv_strday into lv_newdate.

    "formula para somar horas planeadas e trabalhadas
    total_planeadas   = '=SUM(E7:AI7)'. "formula para somar horas a trabalhar
    total_trabalhadas = '=SUM(E8:AI8)'. "formula para somar horas a trabalhar

    "rever as horas trabalhadas conforme consulta - aguardar info adicional
    data: horas_planeadas type p decimals 2.
    horas_planeadas = '8'.

    "repete a quantidade de dias que tem o mes
    do lv_countdays times.

      "funcao retorna a data formatada [ numdia + nomediasemana ]
      call function 'ZWEEKDATE'
        exporting
          date           = lv_newdate
        importing
          format_daydate = lv_stringdaydate
          e_result       = e_result.

      if sy-subrc eq 0.

        "verifica se é sábado ou domingo para nao contabilizar as horas.
        if lv_stringdaydate cs 'Sábado' or lv_stringdaydate cs 'Domingo'.
          horas_planeadas = '0'.
        else.
          horas_planeadas = '8'.
        endif.

        "cria a celula
        lo_worksheet->set_cell( ip_row = 6 ip_column = lv_counterdays ip_value = lv_stringdaydate ip_style = tp_style_bold_center_guid  ). "cabeçalho do calendário
        lo_worksheet->set_cell( ip_row = 7 ip_column = lv_counterdays ip_value = horas_planeadas  ip_style = tp_style_bold_center_guid2 ). "horas planeadas
        lo_worksheet->set_cell( ip_row = 8 ip_column = lv_counterdays ip_value = ''               ip_style = tp_style_bold_center_guid2 ). "horas trabalhadas

        "cabeçalho de ausencia de trabalho
        lo_worksheet->set_cell( ip_row = 9 ip_column = lv_counterdays ip_value = 'Tempo' ip_style = tp_style_bold_center_guid ). "horas trabalhadas
        "coluna de tempo trabalhado
        lo_worksheet->set_cell( ip_row = 10 ip_column = lv_counterdays ip_value = 0 ip_style = tp_style_bold_center_guid2 ). "horas trabalhadas
        lo_worksheet->set_cell( ip_row = 11 ip_column = lv_counterdays ip_value = 0 ip_style = tp_style_bold_center_guid2 ). "horas trabalhadas
        lo_worksheet->set_cell( ip_row = 12 ip_column = lv_counterdays ip_value = 0 ip_style = tp_style_bold_center_guid2 ). "horas trabalhadas

        "setup da coluna para cada celula criada
        lo_column = lo_worksheet->get_column( ip_column = lv_counterdays ).
        lo_column->set_width( ip_width = 20 ).

        add 1 to lv_counterdays. "incrementa o contador para a proxima coluna

        lv_day = lv_strday. "casting int
        add 1 to lv_day.    "incrementa o dia
        lv_strday = lv_day. "casting string

        "se nao passamos dos 10 primeiros dias do mês
        if lv_day lt 10.
          concatenate '0' lv_strday into lv_strday. "adiciona o 0 na frente do numero
        endif.

        clear lv_newdate.                                 "limpa a variavel
        lv_newdate = lv_date+0(6).                        "busca novamente ano e mes
        concatenate lv_newdate lv_strday into lv_newdate. "redefine a data para o dia seguinte.

      endif.

    enddo.

    lv_counterdays = 5. "reseta o contador para a 5th coluna
    clear: lv_day, lv_strday. "limpa os contadores de dias em string e int.

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER2->GET_DATA
* +-------------------------------------------------------------------------------------------------+
* | [<---] COLABORADORES                  TYPE        ZCOL_TT
* | [<---] E_RESULT                       TYPE        ZRLA_RESULT
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method get_data.

    me->it_colaboradores = value #( ( pernr = '1'  sname = 'Colaborador A' vdsk1 = 'PT01'  kostl = '001'  )
                                    ( pernr = '2'  sname = 'Colaborador B' vdsk1 = 'PT02'  kostl = '002'  )
                                    ( pernr = '3'  sname = 'Colaborador C' vdsk1 = 'PT03'  kostl = '003'  )
                                    ( pernr = '4'  sname = 'Colaborador D' vdsk1 = 'PT04'  kostl = '004'  )
                                    ( pernr = '5'  sname = 'Colaborador E' vdsk1 = 'PT05'  kostl = '005'  )
                                    ( pernr = '6'  sname = 'Colaborador F' vdsk1 = 'PT06'  kostl = '006'  )
                                    ( pernr = '7'  sname = 'Colaborador G' vdsk1 = 'PT07'  kostl = '007'  )
                                    ( pernr = '8'  sname = 'Colaborador H' vdsk1 = 'PT08'  kostl = '008'  )
                                    ( pernr = '9'  sname = 'Colaborador I' vdsk1 = 'PT09'  kostl = '009'  )
                                    ( pernr = '10' sname = 'Colaborador J' vdsk1 = 'PT010' kostl = '0010' ) ).

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->GET_FILE_DIRECTORY
* +-------------------------------------------------------------------------------------------------+
* | [--->] FILENAME                       TYPE        STRING
* | [<---] FULL_PATH                      TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method get_file_directory.

    data: namefile  type string, "nome do arquivo
          directory type string, "diretorio
          fullpath  type string. "caminho completo

    namefile = 'file'.

    "adiciona a extensão '.xlsx' ao nome do arquivo
    me->append_extension(
      exporting
        old_extension = namefile
      importing
        new_extension = namefile
    ).

    "diálogo para selecionar diretorio e nome do arquivo
    call method cl_gui_frontend_services=>file_save_dialog
      exporting
        default_extension = 'xlsx'
        default_file_name = namefile
      changing
        filename          = namefile
        path              = directory
        fullpath          = fullpath
      exceptions
        others            = 1.

    "se o user nao cancelar a operacao...
    if sy-subrc = 0.
      concatenate directory namefile into fullpath separated by '\'. "cria diretorio completo do arquivo
    else.
      clear fullpath. "limpa o caminho
    endif.

    full_path = fullpath. "retorna caminho completo do arquivo

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->SET_DATABASE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method set_database.

    data(o_xl_ws) = o_xl->get_active_worksheet( ).
    lo_worksheet = o_xl_ws.

    "insere titulo na worksheet.
    data: lv_title type zexcel_sheet_title. "titulo de worksheets
    lv_title = 'Colaboradores'.
    lo_worksheet->set_title( ip_title = lv_title ).

    data: it_stringtable type table of string, "tabela da dropdown validation
          ls_stringtable type string.

    "index para correr as linhas
    data: lv_index type i.
    lv_index = 2.

    "cabeçalho da tabela
    lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = 'Número'          ip_style = tp_style_bold_center_guid ).
    lo_worksheet->set_cell( ip_row = 1 ip_column = 'B' ip_value = 'Colaborador'     ip_style = tp_style_bold_center_guid ).
    lo_worksheet->set_cell( ip_row = 1 ip_column = 'C' ip_value = 'Equipa'          ip_style = tp_style_bold_center_guid ).
    lo_worksheet->set_cell( ip_row = 1 ip_column = 'D' ip_value = 'Centro de Custo' ip_style = tp_style_bold_center_guid ).

    "linhas da tabela
    loop at me->it_colaboradores into me->ls_colaborador.
      lo_worksheet->set_cell( ip_row = lv_index ip_column = 'A' ip_value = ls_colaborador-pernr ip_style = tp_style_bold_center_guid2 ).
      lo_worksheet->set_cell( ip_row = lv_index ip_column = 'B' ip_value = ls_colaborador-sname ip_style = tp_style_bold_center_guid2 ).
      lo_worksheet->set_cell( ip_row = lv_index ip_column = 'C' ip_value = ls_colaborador-vdsk1 ip_style = tp_style_bold_center_guid2 ).
      lo_worksheet->set_cell( ip_row = lv_index ip_column = 'D' ip_value = ls_colaborador-kostl ip_style = tp_style_bold_center_guid2 ).
      add 1 to lv_index. "incrementa o contador

      "preenche a tabela da dropdown.
      concatenate ls_colaborador-pernr '-' ls_colaborador-sname into ls_stringtable separated by space.
      append ls_stringtable to it_stringtable.
      clear: ls_stringtable, me->ls_colaborador.
    endloop.

    lv_index = 2. "reseta o contador

    "começa a escrever a tabela da dropdown
    lo_worksheet->set_cell( ip_row = 1 ip_column = 'Z' ip_value = 'Lista de Colaboradores' ip_style = tp_style_bold_center_guid ).
    loop at it_stringtable into ls_stringtable.
      lo_worksheet->set_cell( ip_row = lv_index ip_column = 'Z' ip_value = ls_stringtable  ip_style = tp_style_bold_center_guid2 ).
      add 1 to lv_index. "incrementa o contador
    endloop.

    "setup das colunas
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

    lv_index = 2. "reseta o contador

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->SET_SHEETS
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method set_sheets.

    data: lv_title type zexcel_sheet_title.
    lv_title = 'Colaboradores'.

    "formulas usadas
    data: lv_formula_pernr type zexcel_cell_formula,
          lv_formula_sname type zexcel_cell_formula,
          lv_formula_vdsk1 type zexcel_cell_formula,
          lv_formula_kostl type zexcel_cell_formula.
    lv_formula_pernr = '=VLOOKUP(A10,Colaboradores!A2:B12,1)'. "procura por id
    lv_formula_sname = '=VLOOKUP(A10,Colaboradores!A2:B12,2)'. "procura por nome
    lv_formula_vdsk1 = '=VLOOKUP(A10,Colaboradores!A2:C12,3)'. "procura por equipa
    lv_formula_kostl = '=VLOOKUP(A10,Colaboradores!A2:D12,4)'. "procura por centro de custos

    "criando uma nova worksheet
    lo_worksheet = o_xl->add_new_worksheet( ).
    lo_worksheet->set_title( ip_title = | { lv_title } | ).

    "------------------------------------------------------------------------------------------------------------------------------------------
    "------------------------------------------------------------------------------------------------------------------------------------------

    "nomes dos campos do cabeçalho
    lo_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = 'N.Mecan:'         ip_style = tp_style_bold_center_guid ).
    lo_worksheet->set_cell( ip_row = 2 ip_column = 'A' ip_value = 'Nome:'            ip_style = tp_style_bold_center_guid ).
    lo_worksheet->set_cell( ip_row = 3 ip_column = 'A' ip_value = 'Equipa:'          ip_style = tp_style_bold_center_guid ).
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'A' ip_value = 'Centro de Custo:' ip_style = tp_style_bold_center_guid ).

    "nomes das linhas do cabeçalho
    lo_worksheet->set_cell( ip_row = 1 ip_column = 'B' ip_value = '' ip_style = tp_style_bold_center_guid2 ip_formula = lv_formula_pernr ).
    lo_worksheet->set_cell( ip_row = 2 ip_column = 'B' ip_value = '' ip_style = tp_style_bold_center_guid2 ip_formula = lv_formula_sname ).
    lo_worksheet->set_cell( ip_row = 3 ip_column = 'B' ip_value = '' ip_style = tp_style_bold_center_guid2 ip_formula = lv_formula_vdsk1 ).
    lo_worksheet->set_cell( ip_row = 4 ip_column = 'B' ip_value = '' ip_style = tp_style_bold_center_guid2 ip_formula = lv_formula_kostl ).

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
    lo_worksheet->set_cell( ip_row = 6 ip_column = 'D' ip_value = 'Totais'                                ip_style = tp_style_bold_center_guid ).
    lo_worksheet->set_cell( ip_row = 7 ip_column = 'D' ip_value = ''       ip_formula = total_planeadas   ip_style = tp_style_bold_center_guid2 ).
    lo_worksheet->set_cell( ip_row = 8 ip_column = 'D' ip_value = ''       ip_formula = total_trabalhadas ip_style = tp_style_bold_center_guid2 ).

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
    lo_worksheet->set_cell( ip_row = 10 ip_column = 'A' ip_value = ''                    ip_style = tp_style_bold_center_guid2 ).

    lo_worksheet->set_cell( ip_row = 11 ip_column = 'A' ip_value = 'PEP 1' ip_style = tp_style_bold_center_guid ).
    lo_worksheet->set_cell( ip_row = 12 ip_column = 'A' ip_value = 'PEP 1' ip_style = tp_style_bold_center_guid ).
    lo_worksheet->set_cell( ip_row = 13 ip_column = 'A' ip_value = 'PEP 1' ip_style = tp_style_bold_center_guid ).

    lo_worksheet->set_cell( ip_row = 11 ip_column = 'B' ip_value = '880 - Horas Noturnas' ip_style = tp_style_bold_center_guid2 ).
    lo_worksheet->set_cell( ip_row = 12 ip_column = 'B' ip_value = '880 - Horas Noturnas' ip_style = tp_style_bold_center_guid2 ).
    lo_worksheet->set_cell( ip_row = 13 ip_column = 'B' ip_value = '880 - Horas Noturnas' ip_style = tp_style_bold_center_guid2 ).

    "------------------------------------------------------------------------------------------------------------------------------------------
    "------------------------------------------------------------------------------------------------------------------------------------------

    "setup da primeira coluna
    lo_column = lo_worksheet->get_column( ip_column = 'A' ).
    lo_column->set_width( ip_width = 30 ).
    lo_column = lo_worksheet->get_column( ip_column = 'B' ).
    lo_column->set_width( ip_width = 50 ).
    lo_column = lo_worksheet->get_column( ip_column = 'D' ).
    lo_column->set_width( ip_width = 20 ).

    "range de busca para a dropdown
    data(lo_range) = o_xl->add_new_range( ).
    lo_range->name = 'CollaboratorNumbers'. "nome do range
    lo_range->set_value(
      ip_sheet_name   = lv_title "sheet escolhida
      ip_start_column = 'Z'
      ip_start_row    = 2
      ip_stop_column  = 'Z'
      ip_stop_row     = lines( me->it_colaboradores ) + 1 "limite do range
    ).

    "validacao do range da dropdown
    lo_data_validation              = lo_worksheet->add_new_data_validation( ).
    lo_data_validation->type        = zcl_excel_data_validation=>c_type_list.
    lo_data_validation->formula1    = 'CollaboratorNumbers'. "nome do range
    lo_data_validation->cell_row    = 10.
    lo_data_validation->cell_column = 'A'.
    lo_data_validation->allowblank  = abap_true.

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER2->SET_STYLE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method set_style.

    "cria objetos das bordas
    create object o_border_dark.
    o_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
    o_border_dark->border_style = zcl_excel_style_border=>c_border_thin.
    create object o_border_light.
    o_border_light->border_color-rgb = zcl_excel_style_color=>c_gray.
    o_border_light->border_style = zcl_excel_style_border=>c_border_thin.

    "monta o primeiro estilo para a coluna A da paginacao
    create object me->lo_style.
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

  endmethod.
ENDCLASS.
