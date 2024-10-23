class ZCL_EXCEL_BUILDER definition
  public
  final
  create public .

public section.

  interfaces ZIF_EXCEL_BOOK_PROPERTIES .
  interfaces ZIF_EXCEL_BOOK_PROTECTION .
  interfaces ZIF_EXCEL_BOOK_VBA_PROJECT .

  types:
    begin of wa_materials,
        matnr type mara-matnr,
        maktx type makt-maktx,
        bwkey type mbew-bwkey,
        lbkum type mbew-lbkum,
        salk3 type mbew-salk3,
      end of wa_materials .

  data O_XL type ref to ZCL_EXCEL .
  data LO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET .
  data LO_HYPERLINK type ref to ZCL_EXCEL_HYPERLINK .
  data LO_COLUMN type ref to ZCL_EXCEL_COLUMN .
  data O_CONVERTER type ref to ZCL_EXCEL_CONVERTER .    "classe para manipulacao de excel

  data GUID type ZEXCEL_CELL_STYLE .
  data lo_style type ref to ZCL_EXCEL_STYLE .

  data:
    wt_materials type table of wa_materials .
  data WS_MATERIALS type WA_MATERIALS .
  data E_RESULT type ZRLA_RESULT .

  methods GET_MATERIALS
    importing
      !MATNR type MARA-MATNR optional
      !BWKEY type MBEW-BWKEY
      !LOW_ERSDA type MARA-ERSDA
      !HIGH_ERSDA type MARA-ERSDA
    exporting
      !MATERIALS type ZMAT_TT
      !E_RESULT type ZRLA_RESULT .
  methods DOWNLOAD_XLS .
  protected section.

private section.

  methods CONVERT_XSTRING .
  methods SET_COLUMNS .
  methods APPEND_EXTENSION                   "prepara caminho para o arquivo com extensao
    importing
      !OLD_EXTENSION type STRING
    exporting
      !NEW_EXTENSION type STRING .
  methods GET_FILE_DIRECTORY                 "prepara o caminho para o arquivo
    importing
      !FILENAME type STRING
    exporting
      !FULL_PATH type STRING .
  methods set_style.
ENDCLASS.



CLASS ZCL_EXCEL_BUILDER IMPLEMENTATION.


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

    " Converte os dados para o formato Excel
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

    "----------------------------------------------------------------

    "----------------------------------------------------------------
    " tratamento de nome e extensão do arquivo
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

    "converte dados para xstring
    me->convert_xstring( ).
    "insere o estilo
    me->set_style( ).

    data(o_xl_ws) = o_xl->get_active_worksheet( ).
    lo_worksheet = o_xl_ws.

    "itera sobre a tabela principal e monta as celulas do excel
    loop at me->wt_materials into ws_materials.

      "adiciona um novo worksheet para cada iteracao
      data(lo_new_worksheet) = o_xl->add_new_worksheet( ).

      "titulo do worksheet
      data(worksheet_title) = conv ZEXCEL_SHEET_TITLE( |Material_{ ws_materials-matnr }| ).
      lo_new_worksheet->set_title( ip_title = worksheet_title ).

      lo_new_worksheet->get_style_cond( ip_guid = guid ).
      lo_new_worksheet->set_cell( ip_row = 1 ip_column = 'A' ip_value = 'Nº Material' ip_style = guid ). " Número do material ip_style =
      lo_new_worksheet->set_cell( ip_row = 2 ip_column = 'A' ip_value = 'Descrição'   ip_style = guid ). " Descrição do material
      lo_new_worksheet->set_cell( ip_row = 3 ip_column = 'A' ip_value = 'Área'        ip_style = guid ). " Chave de avaliação
      lo_new_worksheet->set_cell( ip_row = 4 ip_column = 'A' ip_value = 'Stock'       ip_style = guid ). " Estoque
      lo_new_worksheet->set_cell( ip_row = 5 ip_column = 'A' ip_value = 'Total'       ip_style = guid ). " Saldo contábil
      lo_new_worksheet->set_cell( ip_row = 6 ip_column = 'A' ip_value = 'Unidade'     ip_style = guid ). " Preço Unidade

      lo_column = lo_new_worksheet->get_column( ip_column = 1 ).
      lo_column->set_width( ip_width = 20 ).

      lo_column = lo_new_worksheet->get_column( ip_column = 2 ).
      lo_column->set_width( ip_width = 20 ).

      lo_new_worksheet->set_cell( ip_row = 1 ip_column = 'B' ip_value   = ws_materials-matnr ip_style = guid ). " Número do material
      lo_new_worksheet->set_cell( ip_row = 2 ip_column = 'B' ip_value   = ws_materials-maktx ip_style = guid ). " Descrição do material
      lo_new_worksheet->set_cell( ip_row = 3 ip_column = 'B' ip_value   = ws_materials-bwkey ip_style = guid ). " Chave de avaliação
      lo_new_worksheet->set_cell( ip_row = 4 ip_column = 'B' ip_value   = ws_materials-lbkum ip_style = guid ). " Estoque
      lo_new_worksheet->set_cell( ip_row = 5 ip_column = 'B' ip_value   = ws_materials-salk3 ip_style = guid ). " Saldo contábil
      lo_new_worksheet->set_cell( ip_row = 6 ip_column = 'B' ip_formula = '=B4 / B5'         ip_style = guid ). " Saldo contábil

      lo_column = lo_new_worksheet->get_column( ip_column = 1 ).
      lo_column->set_width( ip_width = 20 ).

    endloop.

    me->set_columns(  ).

    "----------------------------------------------------------------

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

    data: namefile  type string,
          directory type string,
          fullpath  type string.

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

**    insere a formula da celula na tabela
*    loop at wt_materials into ws_materials.
*      ws_materials-valor_unitario = '=C7 / C8'.
*      modify wt_materials from ws_materials.
*    endloop.

    materials = me->wt_materials. "tabela recebe objeto de classe.

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

    lo_column = lo_worksheet->get_column( ip_column = 'A' ). "Ajusta a coluna A
    lo_column->set_width( ip_width = 20 ). "Define a largura da coluna A como 200 unidades

    lo_column = lo_worksheet->get_column( ip_column = 'B' ). "Ajusta a coluna A
    lo_column->set_width( ip_width = 20 ). "Define a largura da coluna A como 200 unidades
    lo_column->set_column_style_by_guid( ip_style_guid = guid ).

    lo_column = lo_worksheet->get_column( ip_column = 'C' ). "Ajusta a coluna A
    lo_column->set_width( ip_width = 20 ). "Define a largura da coluna A como 200 unidades
    lo_column->set_column_style_by_guid( ip_style_guid = guid ).

    lo_column = lo_worksheet->get_column( ip_column = 'D' ). "Ajusta a coluna A
    lo_column->set_width( ip_width = 20 ). "Define a largura da coluna A como 200 unidades
    lo_column->set_column_style_by_guid( ip_style_guid = guid ).

    lo_column = lo_worksheet->get_column( ip_column = 'E' ). "Ajusta a coluna A
    lo_column->set_width( ip_width = 20 ). "Define a largura da coluna A como 200 unidades
    lo_column->set_column_style_by_guid( ip_style_guid = guid ).

    lo_column = lo_worksheet->get_column( ip_column = 'F' ). "Ajusta a coluna A
    lo_column->set_width( ip_width = 20 ). "Define a largura da coluna A como 200 unidades
    lo_column->set_column_style_by_guid( ip_style_guid = guid ).

  endmethod.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZCL_EXCEL_BUILDER->SET_STYLE
* +-------------------------------------------------------------------------------------------------+
* +--------------------------------------------------------------------------------------</SIGNATURE>
  method set_style.

      create object me->lo_style.

      me->lo_style->font->name = 'Arial'.            " Definir a fonte como Arial
      me->lo_style->font->size = 12.                 " Definir o tamanho da fonte
      me->lo_style->font->bold = abap_true.          " Definir a fonte como negrito
      me->lo_style->font->italic = abap_true.        " Sem itálico
      me->lo_style->font->color = 'FF0000'.          " Definir a cor da fonte como vermelha

      me->lo_style->fill->filltype = 'solid'.        " Preenchimento sólido
      me->lo_style->fill->fgcolor = 'FFFF00'.        " Cor de preenchimento amarelo
      me->lo_style->fill->bgcolor = '000000'.        " Cor de fundo preta

      data c_border_medium type zexcel_border value 'medium'. "#EC NOTEXT.

*     lo_style->borders->allborders = C_BORDER_MEDIUM.

      me->lo_style->alignment->horizontal = 'center'. " Alinhamento centralizado horizontalmente
      me->lo_style->alignment->vertical   = 'center'.  " Alinhamento centralizado verticalmente

      me->lo_style->number_format->format_code = '#,##0.00 [$R$-416];[RED]-#,##0.00 [$R$-416]'.

*      lo_style->protection->locked = abap_true.   " Bloquear a célula

      guid = lo_style->get_guid( ).

      me->o_xl->add_new_style(
          ip_guid = me->guid
      ).

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
ENDCLASS.
