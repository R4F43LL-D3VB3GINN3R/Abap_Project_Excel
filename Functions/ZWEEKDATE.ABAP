function zweekdate.
*"----------------------------------------------------------------------
*"*"Interface local:
*"  IMPORTING
*"     REFERENCE(DATE) TYPE  SY-DATUM
*"  EXPORTING
*"     REFERENCE(FORMAT_DAYDATE) TYPE  STRING
*"     REFERENCE(E_RESULT) TYPE  ZRLA_RESULT
*"----------------------------------------------------------------------

  data: lv_date      type scal-date,      "data atual
        lv_day       type scal-indicator, "dia da semana em numero
        lv_month     type string,         "dia do mês.
        lv_stringday type string.         "dia do mês

  lv_date = date. "data atual recebe parametro de envio

  call function 'DATE_COMPUTE_DAY'
    exporting
      date = lv_date
    importing
      day  = lv_day.

  if sy-subrc ne 0.
    e_result-rc = sy-subrc.
    e_result-message = | Não foi possível estabelecer uma data com formatação original. |.
    return.
  endif.

  format_daydate = lv_day. "cast string
  lv_month = lv_date+6(2).  "recebe o dia do mes

  "pegando o dia da semana.
  case format_daydate.
    when 1.
      format_daydate = 'Segunda-Feira'.
    when 2.
      format_daydate = 'Terça-Feira'.
    when 3.
      format_daydate = 'Quarta-Feira'.
    when 4.
      format_daydate = 'Quinta-Feira'.
    when 5.
      format_daydate = 'Sexta-Feira'.
    when 6.
      format_daydate = 'Sábado'.
    when 7.
      format_daydate = 'Domingo'.
  endcase.

  concatenate lv_month format_daydate into format_daydate separated by '-'. "concatena dia da semana e dia do mes

endfunction.
