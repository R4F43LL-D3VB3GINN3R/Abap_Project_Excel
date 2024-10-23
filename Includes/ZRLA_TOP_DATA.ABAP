*&---------------------------------------------------------------------*
*& Include          ZRLA_TOP_DATA
*&---------------------------------------------------------------------*

tables: mara, makt, mbew.

"tela de selecao
*&---------------------------------------------------------------------*

selection-screen: begin of block a1 with frame.
parameters: p_matnr type matnr,
            p_bwkey type mbew-bwkey default 'PT02'.
select-options: s_ersda for mara-ersda.
selection-screen: end of block a1.

selection-screen: begin of block b2 with frame title text-002.
selection-screen: begin of line.
selection-screen: comment 1(30) text-003.
selection-screen: end of line.
selection-screen: end of block b2.

selection-screen: begin of line.
selection-screen: comment 1(30) text-004.
selection-screen: end of line.

"variaveis - estruturas - tabelas
*&---------------------------------------------------------------------*

data: it_mat type zmat_tt, "tabela de materiais
      ls_mat type zmat_st. "estrutura de materiais

"search help
*&---------------------------------------------------------------------*

types: begin of wa_id_mat,
         matnr type mara-matnr,
       end of wa_id_mat.

data: it_mat_id type table of wa_id_mat. "tabela de ids de materiais

data: it_return type table of ddshretval,
      ls_return type ddshretval.

"classes
*&---------------------------------------------------------------------*

  data: ol_excel type ref to zcl_excel_builder.

"alv container
*&---------------------------------------------------------------------*

"instancias alv
data: lo_alv           type ref to cl_salv_table,
      lo_alv_functions type ref to cl_salv_functions,
      lo_alv_columns   type ref to cl_salv_columns,
      lo_alv_column    type ref to cl_salv_column,
      lo_alv_display   type ref to cl_salv_display_settings.

"classe de eventos
data: lo_events type ref to cl_salv_events_table.

"outros
*&---------------------------------------------------------------------*

data: result type zrla_result. "estrutura para resultado de operacoes

"dados inicializados como...
*&---------------------------------------------------------------------*

*adicionar ajuda de pesquisa (F4) ao select-options
at selection-screen on value-request for p_matnr.

  if it_mat_id is not initial.

  call function 'F4IF_INT_TABLE_VALUE_REQUEST'
    exporting
      retfield        = 'MATNR'
      value_org       = 'S'
    tables
      value_tab       = it_mat_id
      return_tab      = it_return
    exceptions
      parameter_error = 1
      no_values_found = 2
      others          = 3.

  endif.

  if sy-subrc eq 0.

    "leitura dos dados para funcao enquanto menu de selecao
    read table it_return into ls_return index 1.
    p_matnr = ls_return-fieldval.

  else.
    return.
  endif.

initialization.

  perform get_ids. "preenche a tabela de s4 com os ids.

  s_ersda-low = '20190423'.
  s_ersda-high = '20241018 '.
  append s_ersda.

form get_ids.

  "Preenche a tabela it_mat_id com valores válidos
  select matnr
    from mara
    into table it_mat_id.

endform.
