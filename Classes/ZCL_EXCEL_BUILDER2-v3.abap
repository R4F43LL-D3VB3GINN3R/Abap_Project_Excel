  METHOD set_database.

    "----------------------------------------------------------------------------------------------
    "info: insere as informacoes gerais dos colaboradores na primeira sheet do excel file
    "
    "data de alteracao: 09.11.2024
    "alteracao: criacao do método
    "criado por: rafael albuquerque
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
