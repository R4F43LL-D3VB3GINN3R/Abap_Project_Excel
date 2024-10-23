*&---------------------------------------------------------------------*
*& Include          ZRLA_ROT
*&---------------------------------------------------------------------*

start-of-selection.

  create object ol_excel. "instancia da classe excel
  perform get_data.       "consulta sql
  perform display_alv.    "renderiza alv

form get_data.

  "metodo que retorna uma tabela de materiais
  ol_excel->get_materials(
    exporting
      matnr      = p_matnr      "numero do material
      bwkey      = p_bwkey      "area
      low_ersda  = s_ersda-low  "data inicial
      high_ersda = s_ersda-high "data final
    importing
      materials  = it_mat       "tabela de materiais
      e_result   = result       "resultado da operacao
  ).

  if result-rc ne 0.
    message result-message type 'E'.
    return.
  endif.

endform.

form build_alv_columns.

  try.
      "funcoes
      lo_alv_functions = lo_alv->get_functions( ).
      lo_alv_functions->set_all( abap_true ).

      "opcoes de display
      lo_alv_display = lo_alv->get_display_settings( ).
      lo_alv_display->set_striped_pattern( cl_salv_display_settings=>true ).
      lo_alv_display->set_list_header( 'Listagem Materiais' ).

      "configurando os nomes das colunas
      lo_alv_columns = lo_alv->get_columns( ).

      "por preferencia, os nomes serao alterados, centralizados e sempre
      "lidos na forma mais extensa possivel e com medidas de largura fixas

      lo_alv_column = lo_alv_columns->get_column( 'MATNR' ).
      lo_alv_column->set_long_text( 'Nº Material' ).
      lo_alv_column->set_fixed_header_text( 'L' ).
      lo_alv_column->set_medium_text( '' ).
      lo_alv_column->set_short_text( '' ).
      lo_alv_column->set_output_length('10').
      lo_alv_column->set_optimized( 'X' ).
      lo_alv_column->set_alignment(
      value = if_salv_c_alignment=>centered
      ).

      lo_alv_column = lo_alv_columns->get_column( 'MAKTX' ).
      lo_alv_column->set_long_text( 'Descrição' ).
      lo_alv_column->set_medium_text( '' ).
      lo_alv_column->set_short_text( '' ).
      lo_alv_column->set_output_length('30').
      lo_alv_column->set_optimized( 'X' ).
      lo_alv_column->set_alignment(
      value = if_salv_c_alignment=>centered
      ).

      lo_alv_column = lo_alv_columns->get_column( 'BWKEY' ).
      lo_alv_column->set_long_text( 'Área' ).
      lo_alv_column->set_medium_text( '' ).
      lo_alv_column->set_short_text( '' ).
      lo_alv_column->set_output_length('10').
      lo_alv_column->set_optimized( 'X' ).
      lo_alv_column->set_alignment(
      value = if_salv_c_alignment=>centered
      ).

      lo_alv_column ?= lo_alv_columns->get_column( 'LBKUM' ).
      lo_alv_column->set_long_text( 'Stock' ).
      lo_alv_column->set_medium_text( '' ).
      lo_alv_column->set_short_text( '' ).
      lo_alv_column->set_output_length('10').
      lo_alv_column->set_optimized( abap_true ).
      lo_alv_column->set_alignment(
      value = if_salv_c_alignment=>centered
      ).

      lo_alv_column ?= lo_alv_columns->get_column( 'SALK3' ).
      lo_alv_column->set_long_text( 'Total' ).
      lo_alv_column->set_medium_text( '' ).
      lo_alv_column->set_short_text( '' ).
      lo_alv_column->set_output_length('10').
      lo_alv_column->set_optimized( abap_true ).
      lo_alv_column->set_alignment(
      value = if_salv_c_alignment=>centered
      ).

    catch cx_root into data(lx_not_found).
      message lx_not_found->get_text( ) type 'E' display like 'E'.
  endtry.

endform.

form display_alv.

  try.
      cl_salv_table=>factory(
      importing
        r_salv_table   = lo_alv
      changing
        t_table        = it_mat
      ).
    catch cx_salv_msg.
  endtry.

  perform build_alv_columns. "formata as colunas do alv

  lo_alv->set_screen_status(
    exporting
      report        = sy-repid
      pfstatus      = 'ZSTATUS'
      set_functions = cl_salv_table=>c_functions_all
  ).

  lo_events = lo_alv->get_event( ). "objeto de evento recebe o evento da classe

  set handler zcl_event_handler=>added_function for lo_events. "envia o evento para o metodo estático da classe

  lo_alv->display( ). "renderiza o alv

endform.
