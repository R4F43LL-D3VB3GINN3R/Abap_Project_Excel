  METHOD get_peps_datafile.

    "----------------------------------------------------------------------------------------------
    "info: recebe os peps dos colaboradores no excel file
    "
    "data de alteracao: 11.11.2024
    "alteracao: criacao do método
    "criado por: rafael albuquerque
    "----------------------------------------------------------------------------------------------

    "----------------------------------------------------
    "    verificacao de dados essenciais para consulta
    "----------------------------------------------------

    IF me->lv_xstr IS INITIAL.
      RETURN.
    ELSEIF me->gv_datemonth IS INITIAL.
      RETURN.
    ELSEIF me->it_employee IS INITIAL.
      RETURN.
    ENDIF.

    "----------------------------------------------------
    "             tratmento das coordenadas
    "----------------------------------------------------

    "coordenada da celula
    DATA: lv_str_coord TYPE string. "Ex: A10
    DATA: lv_hour_index TYPE i.     "Ex: 144
    DATA: lv_coord_num TYPE i.      "Ex: 10

    "metodo para envio de coordenadas
    me->set_coordenates(
      EXPORTING
        letter_coord   = 'A'
      IMPORTING
        string_coord_a = lv_str_coord  "coluna / linha
*        string_coord_b =
        index_coord    = lv_hour_index "index da hora trabalhada
        numrow         = lv_coord_num  "numero da linha
    ).

    "----------------------------------------------------
    "                 leitura do arquivo
    "----------------------------------------------------

    "leitor do arquivo
    DATA(lo_reader) = NEW zcl_excel_reader_2007( ).
    DATA(lo_excel)  = lo_reader->zif_excel_reader~load( i_excel2007 = me->lv_xstr ).  "passa o XSTRING carregado
    DATA(i) = 2.

    "itera por todas as sheets do excel, seja ela quantas houverem
    WHILE i <= lo_excel->get_worksheets_size( ).

      "começa a partir da segunda sheet, sendo a primeira a exibicao de dados gerais
      DATA(lo_worksheet) = lo_excel->get_worksheet_by_index( i ).

      CLEAR me->ls_peps.

      "pega primeiramente o numero do colaborador
      READ TABLE lo_worksheet->sheet_content REFERENCE INTO DATA(cell) INDEX 2. "B2

      "numero do colaborador
      me->ls_peps-num = cell->cell_value.

      "----------------------------------------------------
      "          itera sobre os seis projetos
      "----------------------------------------------------

      "itera sobre os seis projetos
      DO 6 TIMES.

        "procura se há peps ativas
        READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell WITH KEY cell_coords = lv_str_coord. "A10...

        "se houver pep disponivel
        IF cell->cell_value NE 'Selecione'.
          ls_peps-pep = cell->cell_value. "recebe o nome do pep

          "itera sobre os 31 dias do mes -- valor fixo
          DO 31 TIMES.

            "verifica as horas trabalhadas
            READ TABLE lo_worksheet->sheet_content REFERENCE INTO cell INDEX lv_hour_index. "E10

            "se houver hora trabalhada...
            IF cell->cell_value NE '0' AND cell->cell_value NE '0,0'.

              me->ls_peps-dia  = gv_datemonth.     "recebe o dia do mes
              me->ls_peps-hora = cell->cell_value. "recebe a hora trabalhada
              me->ls_peps-row  = lv_coord_num.     "recebe a linha do projeto
              APPEND ls_peps TO it_peps.           "insere a tabela de peps
              CLEAR: ls_peps-dia, ls_peps-hora, ls_peps-row.
              CLEAR: cell->cell_value.
            ENDIF.

            ADD 1 TO lv_hour_index. "incrementa para a proxima hora
            ADD 1 TO gv_datemonth.  "incrementa para o proximo dia

          ENDDO.

          me->get_month_datafile( ). "reseta data do mes

        ENDIF.

        "---------------------------------------------------------------------
        "            redefine a coordenada para o proximo pep
        "---------------------------------------------------------------------

        "método para troca de coordenadas
        me->switch_coordenates(
          EXPORTING
            coordenate     = lv_str_coord "coordenada enviada
          IMPORTING
           string_coord_a = lv_str_coord  "nova coordenada
*            string_coord_b = lv_str_coord  "nova coordenada
            index_coord    = lv_hour_index "index da celula
            numrow         = lv_coord_num  "numero da linha
        ).

      ENDDO.

      "-------------------------------------------
      "    redefine dados para proxima sheet.
      "-------------------------------------------

      me->get_month_datafile( ). "reseta data do mes

      "passa para a próxima sheet
      ADD 1 TO i.

      CLEAR ls_peps. "limpa a estrutura para a proxima sheet

    ENDWHILE.

  ENDMETHOD.
