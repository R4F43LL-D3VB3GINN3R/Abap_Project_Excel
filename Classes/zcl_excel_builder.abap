class zcl_excel_builder definition
  public
  final
  create public .

  public section.

    interfaces zif_excel_book_properties .
    interfaces zif_excel_book_protection .
    interfaces zif_excel_book_vba_project .

    types:
      begin of wa_materials,
        matnr type mara-matnr,
        maktx type makt-maktx,
        bwkey type mbew-bwkey,
        lbkum type mbew-lbkum,
        salk3 type mbew-salk3,
      end of wa_materials .

    data: o_xl         type ref to zcl_excel,
          lo_worksheet type ref to zcl_excel_worksheet,
          lo_hyperlink type ref to zcl_excel_hyperlink,
          lo_column    type ref to zcl_excel_column,
          o_converter  type ref to zcl_excel_converter.

    data: guid     type zexcel_cell_style,
          lo_style type ref to zcl_excel_style.

    data:
      wt_materials type table of wa_materials,
      ws_materials type wa_materials,
      e_result     type zrla_result.

    data: o_border_dark              type ref to zcl_excel_style_border,
          o_border_light             type ref to zcl_excel_style_border,
          tp_style_bold_center_guid  type zexcel_cell_style,
          tp_style_bold_center_guid2 type zexcel_cell_style.

    methods get_materials
      importing
        !matnr      type mara-matnr optional
        !bwkey      type mbew-bwkey
        !low_ersda  type mara-ersda
        !high_ersda type mara-ersda
      exporting
        !materials  type zmat_tt
        !e_result   type zrla_result .
    methods download_xls .
    methods set_sheets.
  protected section.

  private section.

    methods convert_xstring .
    methods set_columns .
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
endclass.



class zcl_excel_builder implementation.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER->APPEND_EXTENSION
* +-------------------------------------------------------------------------------------------------+
* | [--->] OLD_EXTENSION                  TYPE        STRING
* | [<---] NEW_EXTENSION                  TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method append_extension.

    concatenate old_extension 'xlsx' into new_extension separated by '.'.

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER->CONVERT_XSTRING
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method convert_xstring.

    create object o_converter.

    "converte os dados para o formato Excel
    o_converter->convert(
      exporting
        it_table      = me->wt_materials
      changing
        co_excel      = me->o_xl
    ).

    "tratamento de erros
    if sy-subrc ne 0.
      message 'Não foi possível converter os dados para xstring' type 'S' display like 'E'.
      return.
    endif.

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER->DOWNLOAD_XLS
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

    "cria objeto excel
    create object o_xl.

    "insere o estilo
    me->set_style( ).
    "converte dados para xstring
    me->convert_xstring( ).
    "insere paginacoes
    me->set_sheets( ).

    "cria um worksheet
    data(o_xl_ws) = o_xl->get_active_worksheet( ).
    lo_worksheet = o_xl_ws.

    "setup da primeira sheet com visao geral da tabela inteira
    me->set_columns(  ).

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

    " Tratamento de erros.
    if sy-subrc ne 0.
      message 'Não foi possível realizar o download do arquivo' type 'S' display like 'E'.
      return.
    endif.

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER->GET_FILE_DIRECTORY
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
* | Instance Public Method ZCL_EXCEL_BUILDER->GET_MATERIALS
* +-------------------------------------------------------------------------------------------------+
* | [--->] MATNR                          TYPE        MARA-MATNR(optional)
* | [--->] BWKEY                          TYPE        MBEW-BWKEY
* | [--->] LOW_ERSDA                      TYPE        MARA-ERSDA
* | [--->] HIGH_ERSDA                     TYPE        MARA-ERSDA
* | [<---] MATERIALS                      TYPE        ZMAT_TT
* | [<---] E_RESULT                       TYPE        ZRLA_RESULT
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method get_materials.

    "verifica se o numero do material foi enviado
    "se nao for, busca todos os materiais relacionados a consulta
    if matnr is not initial.

      select mara~matnr,
             makt~maktx,
             mbew~bwkey,
             mbew~lbkum,
             mbew~salk3
      from mara
      inner join makt on makt~matnr = mara~matnr
      inner join mbew on mbew~matnr = mara~matnr
      into corresponding fields of table @me->wt_materials
      where mara~lvorm ne 'X'
      and mara~matnr eq @matnr
      and mbew~bwkey eq @bwkey
      and mara~ersda ge @low_ersda
      and mara~ersda le @high_ersda.

    else.

      select mara~matnr,
             makt~maktx,
             mbew~bwkey,
             mbew~lbkum,
             mbew~salk3
      from mara
      inner join makt on makt~matnr = mara~matnr
      inner join mbew on mbew~matnr = mara~matnr
      into corresponding fields of table @me->wt_materials
      where mara~lvorm ne 'X'
      and mbew~bwkey eq @bwkey
      and mara~ersda ge @low_ersda
      and mara~ersda le @high_ersda.

    endif.

    materials = me->wt_materials. "atributo de classe recebe resultado da consulta e envia por parametro

    "retorno da consulta
    if materials is initial.
      e_result-rc = sy-subrc.
      e_result-message = 'Não foram retornados dados da consulta'.
    else.
      e_result-rc = sy-subrc.
      sort me->wt_materials by matnr ascending.
    endif.

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER->SET_COLUMNS
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method set_columns.

    "titulo do worksheet
    data(worksheet_title) = conv zexcel_sheet_title( |Materiais| ).
    lo_worksheet->set_title( ip_title = worksheet_title ).

    lo_column = lo_worksheet->get_column( ip_column = 'A' ).
    lo_column->set_width( ip_width = 20 ).

    lo_column = lo_worksheet->get_column( ip_column = 'B' ).
    lo_column->set_width( ip_width = 20 ).
    lo_column->set_column_style_by_guid( ip_style_guid = guid ).

    lo_column = lo_worksheet->get_column( ip_column = 'C' ).
    lo_column->set_width( ip_width = 20 ).
    lo_column->set_column_style_by_guid( ip_style_guid = guid ).

    lo_column = lo_worksheet->get_column( ip_column = 'D' ).
    lo_column->set_width( ip_width = 20 ).
    lo_column->set_column_style_by_guid( ip_style_guid = guid ).

    lo_column = lo_worksheet->get_column( ip_column = 'E' ).
    lo_column->set_width( ip_width = 20 ).
    lo_column->set_column_style_by_guid( ip_style_guid = guid ).

    lo_column = lo_worksheet->get_column( ip_column = 'F' ).
    lo_column->set_width( ip_width = 20 ).
    lo_column->set_column_style_by_guid( ip_style_guid = guid ).

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER->SET_SHEETS
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method set_sheets.

    data: unit_price type string value '=B4 / B5'. "define uma formula excel

    "itera sobre a tabela principal e monta as celulas do excel
    loop at me->wt_materials into ws_materials.

      "adiciona um novo worksheet para cada iteracao
      data(lo_new_worksheet) = o_xl->add_new_worksheet( ).

      "titulo do worksheet
      data(worksheet_title) = conv zexcel_sheet_title( |Material_{ ws_materials-matnr }| ).
      lo_new_worksheet->set_title( ip_title = worksheet_title ).

      if tp_style_bold_center_guid is not initial.

        "tratamento da formula para campos com valores zero
        if ws_materials-lbkum eq 0 or ws_materials-salk3 eq 0.
          unit_price = '0'.
        else.
          unit_price = '=ROUND(B4 / B5, 2)'. "resultado da operacao de divisao com duas casas decimais
        endif.

        "construcao da primeira coluna
        lo_new_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = 'Nº Material' ip_style = tp_style_bold_center_guid ). " Número do material
        lo_new_worksheet->set_cell( ip_row = 2 ip_column = 'A' ip_value = 'Descrição'   ip_style = tp_style_bold_center_guid ). " Descrição do material
        lo_new_worksheet->set_cell( ip_row = 3 ip_column = 'A' ip_value = 'Área'        ip_style = tp_style_bold_center_guid ). " Chave de avaliação
        lo_new_worksheet->set_cell( ip_row = 4 ip_column = 'A' ip_value = 'Stock'       ip_style = tp_style_bold_center_guid ). " Estoque
        lo_new_worksheet->set_cell( ip_row = 5 ip_column = 'A' ip_value = 'Total'       ip_style = tp_style_bold_center_guid ). " Saldo contábil
        lo_new_worksheet->set_cell( ip_row = 6 ip_column = 'A' ip_value = 'Unidade'     ip_style = tp_style_bold_center_guid ). " Preço Unidade

        "construcao da segunda coluna
        lo_new_worksheet->set_cell( ip_row = 1 ip_column = 'B' ip_value   = ws_materials-matnr ip_style = tp_style_bold_center_guid2 ). " Número do material
        lo_new_worksheet->set_cell( ip_row = 2 ip_column = 'B' ip_value   = ws_materials-maktx ip_style = tp_style_bold_center_guid2 ). " Descrição do material
        lo_new_worksheet->set_cell( ip_row = 3 ip_column = 'B' ip_value   = ws_materials-bwkey ip_style = tp_style_bold_center_guid2 ). " Chave de avaliação
        lo_new_worksheet->set_cell( ip_row = 4 ip_column = 'B' ip_value   = ws_materials-lbkum ip_style = tp_style_bold_center_guid2 ). " Estoque
        lo_new_worksheet->set_cell( ip_row = 5 ip_column = 'B' ip_value   = ws_materials-salk3 ip_style = tp_style_bold_center_guid2 ). " Saldo contábil
        lo_new_worksheet->set_cell( ip_row = 6 ip_column = 'B' ip_formula = unit_price         ip_style = tp_style_bold_center_guid2 ). " Preço Unidade

        "setup da primeira coluna
        lo_column = lo_new_worksheet->get_column( ip_column = 'A' ).
        lo_column->set_width( ip_width = 20 ).

        "setup da segunda coluna
        lo_column = lo_new_worksheet->get_column( ip_column = 'B' ).
        lo_column->set_width( ip_width = 20 ).

      endif.

    endloop.

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER->SET_STYLE
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

    "monta o primeiro estilo
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

    "monta o segundo estilo
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


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER->ZIF_EXCEL_BOOK_PROTECTION~INITIALIZE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method zif_excel_book_protection~initialize.
    " Método para inicializar as configurações de proteção das planilhas Excel.
    " Esse método pode ser utilizado para definir as configurações de proteção,
    " como senhas ou restrições de edição, antes de aplicar a proteção nas planilhas.
  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER->ZIF_EXCEL_BOOK_VBA_PROJECT~SET_CODENAME
* +-------------------------------------------------------------------------------------------------+
* | [--->] IP_CODENAME                    TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method zif_excel_book_vba_project~set_codename.
    " Método para definir o *codename* de um objeto no projeto VBA do documento Excel.
    " O *codename* é um identificador que pode ser utilizado para referenciar objetos
    " como planilhas ou módulos de forma programática no VBA.
  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER->ZIF_EXCEL_BOOK_VBA_PROJECT~SET_CODENAME_PR
* +-------------------------------------------------------------------------------------------------+
* | [--->] IP_CODENAME_PR                 TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method zif_excel_book_vba_project~set_codename_pr.
    " Método para definir o *codename* de um projeto ou módulo específico no VBA.
    " Esse método pode ser utilizado para atualizar o *codename* de um elemento do projeto
    " VBA, permitindo referenciá-lo programaticamente com um novo nome.
  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZCL_EXCEL_BUILDER->ZIF_EXCEL_BOOK_VBA_PROJECT~SET_VBAPROJECT
* +-------------------------------------------------------------------------------------------------+
* | [--->] IP_VBAPROJECT                  TYPE        XSTRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method zif_excel_book_vba_project~set_vbaproject.
    " Método para inserir ou modificar um projeto VBA no documento Excel.
    " Esse método deve aceitar um projeto VBA na forma de um XSTRING e realizar a
    " inserção ou atualização do projeto dentro do arquivo Excel, permitindo a execução
    " de código VBA associado.
  endmethod.
endclass.
