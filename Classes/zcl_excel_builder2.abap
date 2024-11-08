  METHOD generate_calendar.

    "metodo para capturar a data enviada pelo programa
    me->get_date( date = gv_datemonth ).

    "verificacao para envio de data
    IF gv_datemonth IS INITIAL.
      MESSAGE | 'Para impressão do calendário é preciso a data' | TYPE 'S' DISPLAY LIKE 'E'.
    ENDIF.

    "buscando a quantidade de dias no mes
    DATA: lv_date           TYPE /osp/dt_date, "data enviada
          lv_countdays      TYPE /osp/dt_day,  "dias do mes recebidos
          lv_countdays2     TYPE i,            "dias do mes em inteiro
          lv_counterdays    TYPE i,            "contador de dias
          lv_newdate        TYPE sy-datum,     "nova data formatada
          lv_stringdaydate  TYPE string,       "dia formatado
          lv_day            TYPE i,            "dia em inteiro
          lv_strday         TYPE string,       "dia em string
          lv_counterployees TYPE i,            "index da tabela de horarios de trabalho
          lv_stringhour     type char5. "horario na celula em decimais

    "rever as horas trabalhadas conforme consulta - aguardar info adicional
    DATA: horas_planeadas type p decimals 2.
    horas_planeadas = '8.00'.
    data: horas_planeadas2 type string.

    "letra da coluna para formula para calculos de horas de trabalhos diarios
    DATA: lv_lettercollum TYPE string.

    "formula para dias trabalhados
    DATA: form_dia_trab TYPE string.

    "valor do horario dosq tempos gastos em projetos
    lv_stringhour = '0,0'.

    "formula para somar horas planeadas e trabalhadas
    total_planeadas   = '=SUM(E7:AI7)'.   "formula para somar horas a trabalhar
    total_trabalhadas = '=SUM(E10:AI15)'. "formula para somar horas trabalhadas

    "-------------------------------------------

    lv_date = gv_datemonth. "recebe a data enviada pelo programa

    "funcao retorna a quantidade de dias do mes
    CALL FUNCTION '/OSP/GET_DAYS_IN_MONTH'
      EXPORTING
        iv_date = lv_date
      IMPORTING
        ev_days = lv_countdays.

    lv_countdays2 = lv_countdays. "casting int
    lv_counterdays = 5.           "inicia o contador como cinco para contar a partir da 5th coluna

    "-------------------------------------------

    "reseta a data
    lv_newdate = lv_date+0(6). "recebe ano + mes
    lv_strday = '01'.          "sempre começamos pelo primeiro dia do mes

    "junta ano + mes e primeiro dia do mes
    CONCATENATE lv_newdate lv_strday INTO lv_newdate.

    "-------------------------------------------

    lv_counterployees = 1. "inicia o contador de index da tabela horarios

    "pega os horarios de cada funcionario por index de tabela
    READ TABLE me->it_colaboradores INTO me->ls_colaborador INDEX lv_counterployees.
    me->get_work_schedule( pernr = me->ls_colaborador-pernr ). "metodo para buscar work schedule

    "-------------------------------------------

    "repete a quantidade de dias que tem o mes
    DO lv_countdays TIMES.

      "funcao retorna a data formatada [ numdia + nomediasemana ]
      CALL FUNCTION 'ZWEEKDATE'
        EXPORTING
          date           = lv_newdate
        IMPORTING
          format_daydate = lv_stringdaydate
          e_result       = e_result.

      IF sy-subrc EQ 0.

        "verifica se é sábado ou domingo para nao contabilizar as horas.
        IF lv_stringdaydate CS 'Sábado' OR lv_stringdaydate CS 'Domingo'.
          horas_planeadas = '0'.
        ELSE.

          "busca o a quantidade de horarios diários do colaborador
          READ TABLE me->tb_psp INTO me->wa_psp INDEX lv_counterployees.

          "verifica se é feriado
          CASE me->wa_psp-ftkla.
            WHEN 1.
              horas_planeadas = '0'.
            WHEN 0.
              horas_planeadas = me->wa_psp-stdaz.
          ENDCASE.

          CLEAR me->wa_psp.

        ENDIF.

        ADD 1 TO lv_counterployees. "incrementa o contador de horarios

        "cria a celula
        TRY.
            lo_worksheet->set_cell( ip_row = 6 ip_column = lv_counterdays ip_value = lv_stringdaydate ip_style = tp_style_bold_center_guid ).  "cabeçalho do calendário
            lo_worksheet->set_cell( ip_row = 7 ip_column = lv_counterdays ip_value = horas_planeadas  ip_style = tp_style_bold_center_guid2 ). "horas planeadas
          CATCH zcx_excel INTO DATA(lx_excel).
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        CLEAR lv_lettercollum. "limpa a letra da coluna para evitar concatenacoes

        "converte o numero da coluna em string da coluna
        TRY.
            me->convert_excel_column(
              EXPORTING
                column_int    = lv_counterdays
              IMPORTING
                column_string = lv_lettercollum
            ).
          CATCH zcx_excel INTO lx_excel.
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        "celula que recebe a formula da soma de horas trabalhadas no dia
        CLEAR form_dia_trab.
        form_dia_trab = '=SUM(' && lv_lettercollum && '10:' && lv_lettercollum && '15)'. "atualiza a formula

        TRY.
            lo_worksheet->set_cell( ip_row = 8 ip_column = lv_counterdays ip_value = '0,0' ip_formula = form_dia_trab ip_style = tp_style_bold_center_guid2 ). "horas trabalhadas
          CATCH zcx_excel INTO lx_excel.
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        "cabeçalho de tempo trabalhado ou ausentado
        TRY.
            lo_worksheet->set_cell( ip_row = 9 ip_column = lv_counterdays ip_value = 'Tempo' ip_style = tp_style_bold_center_guid ). "horas trabalhadas
            "colunas de tempo de trabalho
            lo_worksheet->set_cell( ip_row = 10 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 IP_CONV_EXIT_LENGTH = abap_true ).
            lo_worksheet->set_cell( ip_row = 11 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 IP_CONV_EXIT_LENGTH = abap_true ).
            lo_worksheet->set_cell( ip_row = 12 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 IP_CONV_EXIT_LENGTH = abap_true ).
            lo_worksheet->set_cell( ip_row = 13 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 IP_CONV_EXIT_LENGTH = abap_true ).
            lo_worksheet->set_cell( ip_row = 14 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 IP_CONV_EXIT_LENGTH = abap_true ).
            lo_worksheet->set_cell( ip_row = 15 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 IP_CONV_EXIT_LENGTH = abap_true ).
          CATCH zcx_excel INTO lx_excel.
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        TRY.
            "setup da coluna para cada celula criada
            lo_column = lo_worksheet->get_column( ip_column = lv_counterdays ).
            lo_column->set_width( ip_width = 25 ). " Define o tamanho da coluna
          CATCH zcx_excel INTO lx_excel.
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        ADD 1 TO lv_counterdays. "incrementa o contador para a proxima coluna

        lv_day = lv_strday. "casting int
        ADD 1 TO lv_day.    "incrementa o dia
        lv_strday = lv_day. "casting string

        "se nao passamos dos 10 primeiros dias do mês
        IF lv_day LT 10.
          CONCATENATE '0' lv_strday INTO lv_strday. "adiciona o 0 na frente do numero
        ENDIF.

        CLEAR lv_newdate.                                 "limpa a variavel
        lv_newdate = lv_date+0(6).                        "busca novamente ano e mes
        CONCATENATE lv_newdate lv_strday INTO lv_newdate. "redefine a data para o dia seguinte.

      ENDIF.

    ENDDO.

    "----------------------------------------------------------------------------------------------------

    "verifica quanto falta para 31 dias para completar o calendario
    IF lv_countdays LT 31.
      "enquanto o calendario nao estiver completo...
      WHILE lv_countdays LT 31.

        "cria a celula
        TRY.
            lo_worksheet->set_cell( ip_row = 6 ip_column = lv_counterdays ip_value = 'XXXXXXX'  ip_style = tp_style_bold_center_guid ).  "cabeçalho do calendário
            lo_worksheet->set_cell( ip_row = 7 ip_column = lv_counterdays ip_value = '0,0'        ip_style = tp_style_bold_center_guid2 ). "horas planeadas
          CATCH zcx_excel INTO lx_excel.
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        TRY.
            lo_worksheet->set_cell( ip_row = 8 ip_column = lv_counterdays ip_value = '0,0'  ip_style = tp_style_bold_center_guid2 ). "horas trabalhadas
          CATCH zcx_excel INTO lx_excel.
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        "cabeçalho de tempo trabalhado ou ausentado
        TRY.
            lo_worksheet->set_cell( ip_row = 9 ip_column = lv_counterdays ip_value = 'Tempo' ip_style = tp_style_bold_center_guid ). "horas trabalhadas
            "colunas de tempo de trabalho
            lo_worksheet->set_cell( ip_row = 10 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ).
            lo_worksheet->set_cell( ip_row = 11 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ).
            lo_worksheet->set_cell( ip_row = 12 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ).
            lo_worksheet->set_cell( ip_row = 13 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ).
            lo_worksheet->set_cell( ip_row = 14 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ).
            lo_worksheet->set_cell( ip_row = 15 ip_column = lv_counterdays ip_value = lv_stringhour ip_style = tp_style_bold_center_guid2 ).
          CATCH zcx_excel INTO lx_excel.
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        TRY.
            "setup da coluna para cada celula criada
            lo_column = lo_worksheet->get_column( ip_column = lv_counterdays ).
            lo_column->set_width( ip_width = 25 ). " Define o tamanho da coluna
          CATCH zcx_excel INTO lx_excel.
            MESSAGE lx_excel->get_text( ) TYPE 'E'.
        ENDTRY.

        add 1 to lv_countdays.
        add 1 to lv_counterdays.

      ENDWHILE.

    ENDIF.

    lv_counterdays = 5. "reseta o contador para a 5th coluna
    lv_counterployees = 1. "reseta o contador de horarios de trabalho
    CLEAR: lv_day, lv_strday. "limpa os contadores de dias em string e int.

    REFRESH me->tb_psp.


  ENDMETHOD.
