*&---------------------------------------------------------------------*
*&  Include           ZTOP
*&---------------------------------------------------------------------*

"----------------------------------------------

TABLES: pernr,  "Dados Mestres RH
        pa0001, "Numero Pessoal,
        pa0002, "Dados Pessoais
        t554s,  "Tipos de Presença e Ausência
        t554t,  "Textos de Ausência e Presença
        t526.   "Tabela de Encarregados

"----------------------------------------------

INFOTYPES: 0001, "Atribuição Organizacional
           0002, "Dados Pessoais
           0003.

"----------------------------------------------

"estrutura modelo para receber os dados do colaborador
TYPES: BEGIN OF wa_col,
         pernr TYPE p_pernr,   "Número Pessoal
         sname TYPE pad_cname, "Nome
         vdsk1 TYPE vdsk1,     "Chave de Organizacao
         kostl TYPE kostl,     "Centro de Custo
         begda TYPE begda,     "Data de Início
         endda TYPE endda,     "Data Final
       END OF wa_col .

"tabela interna + estrutura
DATA: it_colaboradores TYPE TABLE OF wa_col,
      ls_colaborador   TYPE wa_col.

"tabela de colaboradores sem range de datas :: classe
DATA: it_colaboradores2 TYPE zcol_tt,
      ls_colaboradores2 TYPE zcol_st.

"----------------------------------------------

"instancia do arquivo excel
DATA: ol_excel TYPE REF TO zcl_excel_builder2.

"----------------------------------------------

"instancias alv
DATA: lo_alv           TYPE REF TO cl_salv_table,
      lo_alv_functions TYPE REF TO cl_salv_functions,
      lo_alv_columns   TYPE REF TO cl_salv_columns,
      lo_alv_column    TYPE REF TO cl_salv_column,
      lo_alv_display   TYPE REF TO cl_salv_display_settings.

"classe de eventos
DATA: lo_events TYPE REF TO cl_salv_events_table.

"----------------------------------------------

"data enviada
DATA: lv_date TYPE sy-datum.

"----------------------------------------------

"lista de meses
DATA: it_list_months TYPE vrm_values,
      ls_list_months LIKE LINE OF it_list_months.

"tabelas e estruturas de dropdown
TYPES: BEGIN OF wa_months,
         month_name TYPE string,
         month_num  TYPE string,
       END OF wa_months.

DATA: it_months TYPE TABLE OF wa_months,
      ls_months TYPE wa_months.

"----------------------------------------------

"estrutura para resultado de operacoes
data: e_result type zrla_result.

"----------------------------------------------

"types de encarregados
types: begin of wa_mangr,
  sachx type t526-sachx,
  sachn type t526-sachn,
end of wa_mangr.

"tabela de encarregados.
data: it_managers type table of t526,
      ls_managers type t526.

"dropdown de encarregados
DATA: it_list_mngr TYPE vrm_values,
      ls_list_mngr LIKE LINE OF it_list_months.

*&---------------------------------------------------------------------*
*&  Include           ZROT
*&---------------------------------------------------------------------*

INITIALIZATION.

  PERFORM fill_dropdown_months.   "dropdown de meses
  PERFORM get_managers.           "consulta por encarregador de colaboradores
  PERFORM fill_dropdown_managers. "dropdown de encarregados.

  "--------------------------------------------------------------
  "bloco anexo de tela de selecao pnp

  SELECTION-SCREEN: BEGIN OF BLOCK a1 WITH FRAME TITLE TEXT-001.
  SELECTION-SCREEN: SKIP 1.
  PARAMETERS: p_month TYPE char12 AS LISTBOX VISIBLE LENGTH 12,
              p_year  TYPE char4  DEFAULT 'YYYY',
              p_mangr TYPE char20 AS LISTBOX VISIBLE LENGTH 20.
  SELECTION-SCREEN: SKIP 1.
  SELECTION-SCREEN: END OF BLOCK a1.

  "--------------------------------------------------------------

START-OF-SELECTION.

  "verifica se existe um opjeto da classe excel
  IF ol_excel IS INITIAL.
    CREATE OBJECT ol_excel.
  ENDIF.

GET pernr. "a cada iteracao deste ciclo ele busca todas as informacoes de um colaborador

  "itera sobre todas as informacoes deste colaborador
  LOOP AT p0001 INTO DATA(ls_p0001) WHERE sachz EQ p_mangr AND endda GT lv_date.
    "passa todas as informacoes do mesmo funcionario para a tabela interna
    MOVE-CORRESPONDING ls_p0001 TO ls_colaborador.
    APPEND ls_colaborador TO it_colaboradores.
  ENDLOOP.

END-OF-SELECTION.

  "---------------------------------------------------
  "---------------------------------------------------
  "---------------------------------------------------

  PERFORM verify_entries.     "tratamento de campos

  "verifica erros nos dados de inicializacao
  IF e_result-rc EQ 1.
    MESSAGE e_result-message TYPE 'S' DISPLAY LIKE 'E'.
    RETURN.
  ENDIF.

  PERFORM format_date.        "tratamento da data.
  PERFORM order_data.         "trata os dados coletados do colaborador
  PERFORM get_employees.      "envia os dados dos colaboradores
  PERFORM display_alv.        "apresenta o alv

  "---------------------------------------------------
  "---------------------------------------------------
  "---------------------------------------------------

FORM fill_dropdown_months.

  CLEAR: it_months, it_list_months.

  "preenche a tabela de meses
  it_months = VALUE #( ( month_name = 'Janeiro'   month_num = '01' )
                       ( month_name = 'Fevereiro' month_num = '02' )
                       ( month_name = 'Março'     month_num = '03' )
                       ( month_name = 'Abril'     month_num = '04' )
                       ( month_name = 'Maio'      month_num = '05' )
                       ( month_name = 'Junho'     month_num = '06' )
                       ( month_name = 'Julho'     month_num = '07' )
                       ( month_name = 'Agosto'    month_num = '08' )
                       ( month_name = 'Setembro'  month_num = '09' )
                       ( month_name = 'Outubro'   month_num = '10' )
                       ( month_name = 'Novembro'  month_num = '11' )
                       ( month_name = 'Dezembro'  month_num = '12' ) ).

  "preenche a dropdown de meses
  LOOP AT it_months INTO ls_months.
    CLEAR ls_list_months.
    ls_list_months-key  = ls_months-month_num.
    ls_list_months-text = ls_months-month_name.
    APPEND ls_list_months TO it_list_months.
    CLEAR ls_months.
  ENDLOOP.

  "chama a funcao que preenche a dropdown
  CALL FUNCTION 'VRM_SET_VALUES'
    EXPORTING
      id     = 'P_MONTH'
      values = it_list_months.

ENDFORM.

FORM get_managers.

  CLEAR it_managers.

  SELECT sachx
         sachn
    FROM t526
    INTO CORRESPONDING FIELDS OF TABLE it_managers
    WHERE werks EQ 'S718'.

  IF it_managers IS INITIAL.
    e_result-rc = 1.
    e_result-message = | Não foi possível achar Encarregados no sistema |.
  ENDIF.

ENDFORM.

FORM fill_dropdown_managers.

  CLEAR it_list_mngr.

  "passa os encarregados para a tabela de dropdown
  LOOP AT it_managers INTO ls_managers.
    ls_list_mngr-key = ls_managers-sachx.
    ls_list_mngr-text = ls_managers-sachn.
    CLEAR ls_managers.
    APPEND ls_list_mngr TO it_list_mngr.
  ENDLOOP.

  "chama a funcao que preenche a dropdown
  CALL FUNCTION 'VRM_SET_VALUES'
    EXPORTING
      id     = 'P_MANGR'
      values = it_list_mngr.

ENDFORM.

FORM verify_entries.

  "verificacao do mes
  IF p_month IS INITIAL.
    e_result-rc = 1.
    e_result-message = | Preencha o campo mês para gerar a Timesheet |.
  ENDIF.

  "verificao do ano
  IF p_year EQ 'YYYY'.
    e_result-rc = 1.
    e_result-message = | Preencha o campo ano devidamente com uma data válida |.
  ELSEIF p_year CA sy-abcde.
    e_result-rc = 1.
    e_result-message = | Preencha o ano com valores válidos |.
  ENDIF.

  "verifica o encarregado
  if p_mangr is initial.
    e_result-rc = 1.
    e_result-message = | Escolha o Encarregado da Timesheet |.
  endif.

ENDFORM.

FORM format_date.

  CLEAR lv_date. "limpa a data

  CONCATENATE p_year p_month '01' INTO lv_date. "ano + mes + dia 01

  "verifica se a data existe e envia
  IF lv_date IS NOT INITIAL.
    ol_excel->get_date( date = lv_date ). "invoca o metodo para enviar a data
  ENDIF.

ENDFORM.

FORM order_data.

  "index temporario
  DATA: index TYPE i.
  index = 1.

  "ordena a tabela interna por numero e ultima data inicial
  SORT it_colaboradores BY pernr ASCENDING begda ASCENDING.

  "-------------------------------------------------------

* "passa os dados das linhas com as ultimas datas
  LOOP AT it_colaboradores INTO ls_colaborador.

    MOVE-CORRESPONDING ls_colaborador TO ls_colaboradores2.
    APPEND ls_colaboradores2 TO it_colaboradores2.

  ENDLOOP.

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

*&---------------------------------------------------------------------*
*&  Include           ZALV
*&---------------------------------------------------------------------*

CLASS zcl_event_handler DEFINITION.

  PUBLIC SECTION.

    CLASS-METHODS:
      "método para implementar botoes na toolbar do alv
      added_function FOR EVENT added_function OF cl_salv_events_table
        IMPORTING
          e_salv_function.

ENDCLASS.

CLASS zcl_event_handler IMPLEMENTATION.

  METHOD added_function.

    CASE sy-ucomm.
      WHEN 'EXPORT'.
        "download do arquivo excel
        ol_excel->download_xls( ).
    ENDCASE.

  ENDMETHOD.

ENDCLASS.
