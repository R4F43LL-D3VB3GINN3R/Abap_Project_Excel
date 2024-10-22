class zcl_excel_builder definition
  public
  final
  create public .

  public section.

    interfaces zif_excel_book_properties .
    interfaces zif_excel_book_protection .
    interfaces zif_excel_book_vba_project .

    data: o_xl         type ref to zcl_excel,           "classe para manipulacao de excel
          lo_worksheet type ref to zcl_excel_worksheet,
          lo_hyperlink type ref to zcl_excel_hyperlink,
          lo_column    type ref to zcl_excel_column,
          o_converter  type ref to zcl_excel_converter.

    data: ip_guid     type zexcel_cell_style,
          io_clone_of type ref to zcl_excel_style.

    types:
      begin of wa_materials,
        matnr type mara-matnr,
        maktx type makt-maktx,
        bwkey type mbew-bwkey,
        lbkum type mbew-lbkum,
        salk3 type mbew-salk3,
      end of wa_materials .
    data:
      wt_materials type table of wa_materials .
    data ws_materials type wa_materials .
    data e_result type zrla_result .

    methods get_materials
      importing
        !matnr      type mara-matnr optional
        !bwkey      type mbew-bwkey
        !low_ersda  type mara-ersda
        !high_ersda type mara-ersda
      exporting
        !materials  type zmat_tt
        !e_result   type zrla_result.

    methods:
      download_xls.     "realiza download do arquivo excel

  protected section.

  private section.

    methods:
      convert_xstring, "converte tabela interna para xstring
      set_columns,     "configura as colunas da planilha

      append_extension                       "prepara caminho para o arquivo com extensao
        importing old_extension type string
        exporting new_extension type string,
      get_file_directory                     "prepara o caminho para o arquivo
        importing
          filename  type string
        exporting
          full_path type string.

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

    "----------------------------------------------------------------
    "converte dados para xstring
    me->convert_xstring( ).

    data(o_xl_ws) = o_xl->get_active_worksheet( ).
    lo_worksheet = o_xl_ws.

    data: lo_style type ref to zcl_excel_style,
          lv_guid  type zexcel_cell_style.

* Criar objeto de estilo
    create object lo_style.

* Estilizar a fonte
    lo_style->font->name = 'Arial'.            " Definir a fonte como Arial
    lo_style->font->size = 12.                 " Definir o tamanho da fonte
    lo_style->font->bold = abap_true.          " Definir a fonte como negrito
    lo_style->font->italic = abap_false.       " Sem itálico
    lo_style->font->color = 'FF0000'.          " Definir a cor da fonte como vermelha

* Estilizar o preenchimento
    lo_style->fill->filltype = 'solid'.        " Preenchimento sólido
    lo_style->fill->fgcolor = 'FFFF00'.        " Cor de preenchimento amarelo
    lo_style->fill->bgcolor = '000000'.        " Cor de fundo preta

    data c_border_medium type zexcel_border value 'medium'. "#EC NOTEXT.

* Estilizar as bordas
*    lo_style->borders->allborders = C_BORDER_MEDIUM.

* Estilizar o alinhamento
    lo_style->alignment->horizontal = 'center'. " Alinhamento centralizado horizontalmente
    lo_style->alignment->vertical = 'center'.   " Alinhamento centralizado verticalmente

* Definir o formato numérico (exemplo para formato de moeda)
    lo_style->number_format->format_code = '#,##0.00 [$R$-416];[RED]-#,##0.00 [$R$-416]'.

* Definir proteção (exemplo de célula bloqueada)
    lo_style->protection->locked = abap_true.   " Bloquear a célula

* Agora obtenha o GUID estilizado
    lv_guid = lo_style->get_guid( ).

* Exibir o GUID estilizado
    write: / 'O GUID do estilo configurado é: ', lv_guid.

    " Adaptação: Criar uma aba para cada linha da tabela
    loop at me->wt_materials into ws_materials.

      " Cria uma nova aba para cada material
      data(lo_new_worksheet) = o_xl->add_new_worksheet( ).

      " Define o título da aba com o número do material para garantir que seja único
      data(worksheet_title) = conv string( |Material_{ ws_materials-matnr }| ).

      " Define os dados da aba com os valores da linha correspondente

      lo_new_worksheet->set_cell( ip_row = 1 ip_column = 1 ip_value = 'Nº Material' ip_style = lv_guid ). " Número do material ip_style =
      lo_new_worksheet->set_cell( ip_row = 2 ip_column = 1 ip_value = 'Descrição' ip_style = lv_guid ). " Descrição do material
      lo_new_worksheet->set_cell( ip_row = 3 ip_column = 1 ip_value = 'Área' ip_style = lv_guid ). " Chave de avaliação
      lo_new_worksheet->set_cell( ip_row = 4 ip_column = 1 ip_value = 'Stock' ip_style = lv_guid ). " Estoque
      lo_new_worksheet->set_cell( ip_row = 5 ip_column = 1 ip_value = 'Total' ip_style = lv_guid ). " Saldo contábil

      lo_column = lo_new_worksheet->get_column( ip_column = 1 ). "Ajusta a coluna A
      lo_column->set_width( ip_width = 20 ). "Define a largura da coluna A como 200 unidades

      lo_column = lo_new_worksheet->get_column( ip_column = 2 ). "Ajusta a coluna A
      lo_column->set_width( ip_width = 20 ). "Define a largura da coluna A como 200 unidades

      lo_new_worksheet->set_cell( ip_row = 1 ip_column = 2 ip_value = ws_materials-matnr ip_style = lv_guid ). " Número do material
      lo_new_worksheet->set_cell( ip_row = 2 ip_column = 2 ip_value = ws_materials-maktx ip_style = lv_guid ). " Descrição do material
      lo_new_worksheet->set_cell( ip_row = 3 ip_column = 2 ip_value = ws_materials-bwkey ip_style = lv_guid ). " Chave de avaliação
      lo_new_worksheet->set_cell( ip_row = 4 ip_column = 2 ip_value = ws_materials-lbkum ip_style = lv_guid ). " Estoque
      lo_new_worksheet->set_cell( ip_row = 5 ip_column = 2 ip_value = ws_materials-salk3 ip_style = lv_guid ). " Saldo contábil

      lo_column = lo_new_worksheet->get_column( ip_column = 1 ). "Ajusta a coluna A
      lo_column->set_width( ip_width = 20 ). "Define a largura da coluna A como 200 unidades

    endloop.

**    *Ajustar as colunas da aba (pode reutilizar o método 'set_columns' se necessário)
    me->set_columns(  ).

    "----------------------------------------------------------------
    " Escrita final do arquivo Excel com todas as abas
    data(o_xlwriter) = cast zif_excel_writer( new zcl_excel_writer_2007( ) ).

    data(lv_xl_xdata) = o_xlwriter->write_file( o_xl ).

    data(it_raw_data) = cl_bcs_convert=>xstring_to_solix( exporting iv_xstring = lv_xl_xdata ).

    "----------------------------------------------------------------
    " Download do arquivo Excel
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
