  METHOD validation_datafile.

    IF table_timesheet IS INITIAL.
      MESSAGE | Não há dados a serem verificados | TYPE 'S' DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    DATA: flag_projects TYPE flag.
    flag_projects = abap_false.
    DATA: flag_auspres TYPE flag.
    flag_auspres = abap_false.
    DATA: table_timesheet2 TYPE ztshralv_tt.
    table_timesheet2 = table_timesheet.
    data: centro_custo type i.

    me->get_projects( ). "
    me->get_auspres( ).
    
    "itera sobre a tabela da timesheet
    LOOP AT table_timesheet2 INTO DATA(ls_timesheet).
      flag_projects = abap_false.
      flag_auspres  = abap_false.

      "-------------------------------------------------------------------------------
      "                        verifica o numero do colaborador
      "-------------------------------------------------------------------------------
      SELECT SINGLE pernr FROM pa0001 INTO @DATA(lv_pernr) WHERE pernr EQ @ls_timesheet-num.
      IF sy-subrc NE 0.
        ls_timesheet-validacao = icon_red_light.
        ls_timesheet-info      = me->st_alv-info = '@05@' && 'Colaborador não existe' .
        MODIFY table_timesheet2 FROM ls_timesheet.
      ELSE.

        "-------------------------------------------------------------------------------
        "                        verifica a equipa do colaborador
        "-------------------------------------------------------------------------------

        SELECT SINGLE pernr FROM pa0001 INTO @DATA(lv_pernr2) WHERE pernr EQ @ls_timesheet-num AND vdsk1 EQ @ls_timesheet-equipa.
        IF sy-subrc NE 0.
          ls_timesheet-validacao = icon_red_light.
          ls_timesheet-info      = me->st_alv-info = '@05@' && 'Colaborador não pertence a Equipa' .
          MODIFY table_timesheet2 FROM ls_timesheet.
        ELSE.

            "-------------------------------------------------------------------------------
            "                       verifica se o projeto existe
            "-------------------------------------------------------------------------------

            IF ls_timesheet-pep IS NOT INITIAL.
              LOOP AT me->it_linha_projetos INTO DATA(ls_projetos).
                IF ls_timesheet-pep EQ ls_projetos-line.
                  flag_projects = abap_true.
                  EXIT.
                ENDIF.
              ENDLOOP.
            IF flag_projects NE abap_true.
              ls_timesheet-validacao = icon_red_light.
              ls_timesheet-info      = me->st_alv-info = '@05@' && 'Projeto Inexistente' .
              MODIFY table_timesheet2 FROM ls_timesheet.
            ELSE.
            ENDIF.

            "-------------------------------------------------------------------------------
            "             verifica se o motivo de ausencia e presenca existe
            "-------------------------------------------------------------------------------

            IF ls_timesheet-auspres IS NOT INITIAL.
              LOOP AT me->it_line_preaus INTO DATA(ls_auspres).
                IF ls_timesheet-auspres EQ ls_auspres-line.
                  flag_auspres = abap_true.
                  EXIT.
                ENDIF.
              ENDLOOP.
              IF flag_auspres NE abap_true.
                ls_timesheet-validacao = icon_red_light.
                ls_timesheet-info      = me->st_alv-info = '@05@' && 'Código de Ausência Inexistente' .
                MODIFY table_timesheet2 FROM ls_timesheet.
              ENDIF.
            ENDIF.

            "-------------------------------------------------------------------------------
            "                          verifica o centro de custo
            "-------------------------------------------------------------------------------

            IF ls_timesheet-cntr_cust IS NOT INITIAL.
              centro_custo = ls_timesheet-cntr_cust.
              SELECT SINGLE pernr FROM pa0001 INTO @DATA(lv_pernr3) WHERE pernr EQ @ls_timesheet-num AND kostl EQ @centro_custo.
              IF sy-subrc NE 0.
                ls_timesheet-validacao = icon_red_light.
                ls_timesheet-info      = me->st_alv-info = '@05@' && 'Centro de Custo não Existe' .
                MODIFY table_timesheet2 FROM ls_timesheet.
              ENDIF.
            ENDIF.

          ENDIF.
        ENDIF.
      ENDIF.
    ENDLOOP.

    table_timesheet_output = table_timesheet2.

  ENDMETHOD.
