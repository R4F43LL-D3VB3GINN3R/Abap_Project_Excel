DATA: lit_t001 TYPE TABLE OF sflight.

DATA: lr_data_table TYPE REF TO data.

data: ol_excel type ref to zcl_excel_builder.

SELECT * FROM SFLIGHT INTO TABLE lit_t001.

GET REFERENCE OF lit_t001 INTO lr_data_table.

zcl_demo_any_table=>display_any_table(
  EXPORTING
    i_table_content = lr_data_table
    i_table_name    = 'T001'
).
