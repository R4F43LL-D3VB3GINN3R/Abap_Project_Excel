*&---------------------------------------------------------------------*
*&  Include           ZROT
*&---------------------------------------------------------------------*

START-OF-SELECTION.

  IF ol_excel IS INITIAL.
    CREATE OBJECT ol_excel.
  ENDIF.


  "a cada iteracao deste ciclo ele busca todas as informacoes de um colaborador

GET pernr.

  "itera sobre todas as informacoes deste colaborador
  LOOP AT p0001 INTO DATA(ls_p0001).
    "passa todas as informacoes do mesmo funcionario para a tabela interna
    MOVE-CORRESPONDING ls_p0001 TO ls_colaborador.
    APPEND ls_colaborador TO it_colaboradores.
  ENDLOOP.

END-OF-SELECTION.

  "---------------------------------------------------
  "---------------------------------------------------
  "---------------------------------------------------

  PERFORM order_data.    "trata os dados coletados
  PERFORM verify_date.   "verifica se o radiobutton foi ativado
  PERFORM get_employees. "envia os dados dos colaboradores
  PERFORM display_alv.   "apresenta o alv

  "---------------------------------------------------
  "---------------------------------------------------
  "---------------------------------------------------

FORM order_data.

  "index temporario
  DATA: index TYPE i.
  index = 1.

  "ordena a tabela interna por numero e ultima data inicial
  SORT it_colaboradores BY pernr ASCENDING begda ASCENDING.

  "-------------------------------------------------------

* "passa os dados das linhas com as ultimas datas
  LOOP AT it_colaboradores INTO ls_colaborador.

    "pega a primeira linha
    READ TABLE it_colaboradores INTO DATA(ls_colab2) INDEX index.

    ADD 1 TO index. "incrementa o index

    "pega a segunda linha
    READ TABLE it_colaboradores INTO DATA(ls_colab3) INDEX index.

    "verifica entre as duas linhas se houve uma mudanca de colaborador
    "como a tabela esta ordenada em ascendente o ultimo registro do colaborador sempre é o que queremos
    "ou seja...a ultima data em que mudou o centro de custo
    IF ls_colab2-pernr NE ls_colab3-pernr.
      MOVE-CORRESPONDING ls_colab2 TO ls_colaboradores2.
      APPEND ls_colaboradores2 TO it_colaboradores2.
    ENDIF.

  ENDLOOP.

  "-------------------------------------------------------

ENDFORM.

FORM verify_date.

  "verifica se o radiobutton foi acionado.
  IF pnptimr2 EQ 'X'.
    "invoca o metodo para enviar a data
    ol_excel->get_date( date = lv_date ).
  else.
    "invoca o metodo para enviar a data
    ol_excel->get_date( date = sy-datum ).
  ENDIF.

ENDFORM.

FORM get_employees.

  "metodo para enviar os dados dos colaboradores
  ol_excel->get_data(
    EXPORTING
      colaboradores = it_colaboradores2
  ).

ENDFORM.

FORM display_alv.

  TRY.
      cl_salv_table=>factory(
      IMPORTING
        r_salv_table   = lo_alv
      CHANGING
        t_table        = it_colaboradores2
      ).
    CATCH cx_salv_msg.
  ENDTRY.

  PERFORM build_alv_columns. "formata as colunas do alv

  lo_alv->set_screen_status(
    EXPORTING
      report        = sy-repid
      pfstatus      = 'ZSTATUS'
      set_functions = cl_salv_table=>c_functions_all
  ).

  lo_events = lo_alv->get_event( ). "objeto de evento recebe o evento da classe

  SET HANDLER zcl_event_handler=>added_function FOR lo_events. "envia o evento para o metodo estático da classe

  lo_alv->display( ). "renderiza o alv

ENDFORM.

FORM build_alv_columns.

  TRY.
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

      lo_alv_column = lo_alv_columns->get_column( 'PERNR' ).
      lo_alv_column->set_long_text( 'Nº Pessoal' ).
      lo_alv_column->set_fixed_header_text( 'L' ).
      lo_alv_column->set_medium_text( '' ).
      lo_alv_column->set_short_text( '' ).
      lo_alv_column->set_output_length('50').
      lo_alv_column->set_optimized( 'X' ).
      lo_alv_column->set_alignment(
      value = if_salv_c_alignment=>centered
      ).

      lo_alv_column = lo_alv_columns->get_column( 'SNAME' ).
      lo_alv_column->set_long_text( 'Nome' ).
      lo_alv_column->set_medium_text( '' ).
      lo_alv_column->set_short_text( '' ).
      lo_alv_column->set_output_length('50').
      lo_alv_column->set_optimized( 'X' ).
      lo_alv_column->set_alignment(
      value = if_salv_c_alignment=>centered
      ).

      lo_alv_column = lo_alv_columns->get_column( 'VDSK1' ).
      lo_alv_column->set_long_text( 'Chave de Organizacao' ).
      lo_alv_column->set_medium_text( '' ).
      lo_alv_column->set_short_text( '' ).
      lo_alv_column->set_output_length('50').
      lo_alv_column->set_optimized( 'X' ).
      lo_alv_column->set_alignment(
      value = if_salv_c_alignment=>centered
      ).

      lo_alv_column ?= lo_alv_columns->get_column( 'KOSTL' ).
      lo_alv_column->set_long_text( 'Centro de Custo' ).
      lo_alv_column->set_medium_text( '' ).
      lo_alv_column->set_short_text( '' ).
      lo_alv_column->set_output_length('50').
      lo_alv_column->set_optimized( abap_true ).
      lo_alv_column->set_alignment(
      value = if_salv_c_alignment=>centered
      ).

    CATCH cx_root INTO DATA(lx_not_found).
      MESSAGE lx_not_found->get_text( ) TYPE 'E' DISPLAY LIKE 'E'.
  ENDTRY.

ENDFORM.
