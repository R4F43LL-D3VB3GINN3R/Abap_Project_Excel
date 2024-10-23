*&---------------------------------------------------------------------*
*& Include          ZRLA_ALV2
*&---------------------------------------------------------------------*

*&---------------------------------------------------------------------*
*& Include          ZRLA_EVENT_ALV
*&---------------------------------------------------------------------*
class zcl_event_handler definition.

  public section.

    class-methods:
      "método para implementar botoes na toolbar do alv
      added_function for event added_function of cl_salv_events_table
        importing
          e_salv_function.

endclass.

class zcl_event_handler implementation.

  method added_function.

    case sy-ucomm.
      when 'EXPORT'.
        ol_excel->download_xls( ).
    endcase.

  endmethod.

endclass.
