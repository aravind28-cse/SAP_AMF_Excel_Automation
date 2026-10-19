# SAP ABAP Excel Automation --- Complete PROC Include

This file puts the `ZAMF_EXCEL_AUTOMATION_PROC` include in one
continuous code block, as requested.

> **Important:** This is a consolidated implementation template, not yet
> guaranteed activation-ready for every SAP release. Verify the
> `CL_FDT_XL_SPREADSHEET` API in SE24. It writes Excel-compatible
> SpreadsheetML XML with an `.xls` extension, not a native `.xlsx`
> workbook. The TOP include must declare all methods and attributes used
> below.

## `ZAMF_EXCEL_AUTOMATION_PROC`

``` abap
*&---------------------------------------------------------------------*
*& Include          ZAMF_EXCEL_AUTOMATION_PROC
*&---------------------------------------------------------------------*

START-OF-SELECTION.

  DATA(lo_app) = NEW lcl_amf_automation( ).
  lo_app->run( ).


CLASS lcl_amf_automation IMPLEMENTATION.

  METHOD run.

    mv_input_file = CONV string( p_infile ).
    mv_release    = CONV i( p_relver ).

    TRY.
        read_excel( ).
        process_data( ).
        select_output_file( ).

        IF mv_output_file IS NOT INITIAL.
          download_excel( ).
          MESSAGE 'Excel processing completed.' TYPE 'S'.
        ELSE.
          MESSAGE 'Output file selection was cancelled.' TYPE 'S'
            DISPLAY LIKE 'W'.
        ENDIF.

      CATCH cx_root INTO DATA(lx_error).
        MESSAGE lx_error->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
    ENDTRY.

  ENDMETHOD.


  METHOD read_excel.

    DATA:
      lt_binary  TYPE solix_tab,
      lv_length  TYPE i,
      lv_xstring TYPE xstring,
      lo_excel   TYPE REF TO cl_fdt_xl_spreadsheet.

    FIELD-SYMBOLS <lt_sheet> TYPE STANDARD TABLE.

    cl_gui_frontend_services=>gui_upload(
      EXPORTING
        filename   = mv_input_file
        filetype   = 'BIN'
      IMPORTING
        filelength = lv_length
      CHANGING
        data_tab   = lt_binary
      EXCEPTIONS
        OTHERS     = 1 ).

    IF sy-subrc <> 0.
      MESSAGE 'Unable to read the input Excel file.' TYPE 'E'.
    ENDIF.

    CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
      EXPORTING
        input_length = lv_length
      IMPORTING
        buffer       = lv_xstring
      TABLES
        binary_tab   = lt_binary
      EXCEPTIONS
        OTHERS       = 1.

    IF sy-subrc <> 0.
      MESSAGE 'Unable to convert the Excel file to XSTRING.' TYPE 'E'.
    ENDIF.

    lo_excel = NEW cl_fdt_xl_spreadsheet(
      document_name = mv_input_file
      xdocument     = lv_xstring ).

    lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
      IMPORTING
        worksheet_names = mt_sheet_names ).

    IF mt_sheet_names IS INITIAL.
      MESSAGE 'No worksheet was found in the input workbook.' TYPE 'E'.
    ENDIF.

    mr_excel_data =
      lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet(
        worksheet_name = mt_sheet_names[ 1 ] ).

    ASSIGN mr_excel_data->* TO <lt_excel>.

    IF sy-subrc <> 0 OR <lt_excel> IS INITIAL.
      MESSAGE 'The first worksheet is empty or could not be read.' TYPE 'E'.
    ENDIF.

  ENDMETHOD.


  METHOD get_cell.

    FIELD-SYMBOLS <lv_value> TYPE any.

    ASSIGN COMPONENT iv_column OF STRUCTURE is_row TO <lv_value>.

    IF sy-subrc = 0.
      rv_value = |{ <lv_value> }|.
      CONDENSE rv_value.
    ELSE.
      CLEAR rv_value.
    ENDIF.

  ENDMETHOD.


  METHOD process_data.

    DATA:
      lv_col          TYPE i,
      lv_header       TYPE string,
      lv_data_type    TYPE string,
      lv_enh_type     TYPE string,
      lv_enh_name     TYPE string,
      lv_scope        TYPE string,
      lv_append_name  TYPE string,
      lv_column_count TYPE i,
      ls_result       TYPE ty_result.

    FIELD-SYMBOLS:
      <ls_header> TYPE any,
      <ls_row>    TYPE any,
      <lv_comp>   TYPE any.

    CLEAR:
      mv_type_col,
      mv_enh_type_col,
      mv_enh_name_col,
      mv_scope_col,
      mv_append_col,
      mt_results.

    ASSIGN <lt_excel>[ 1 ] TO <ls_header>.
    IF sy-subrc <> 0.
      MESSAGE 'The workbook does not contain a header row.' TYPE 'E'.
    ENDIF.

    DO.
      ASSIGN COMPONENT sy-index OF STRUCTURE <ls_header> TO <lv_comp>.
      IF sy-subrc <> 0.
        EXIT.
      ENDIF.
      lv_column_count = sy-index.
    ENDDO.

    DO lv_column_count TIMES.
      lv_col = sy-index.
      lv_header = get_cell(
        is_row    = <ls_header>
        iv_column = lv_col ).

      CONDENSE lv_header.
      TRANSLATE lv_header TO UPPER CASE.

      CASE lv_header.
        WHEN 'DATA TYPE'.
          mv_type_col = lv_col.
        WHEN 'ENHANCEMENT TYPE'.
          mv_enh_type_col = lv_col.
        WHEN 'ENHANCEMENT NAME'.
          mv_enh_name_col = lv_col.
        WHEN 'AMF SCOPE'.
          mv_scope_col = lv_col.
        WHEN 'PROPOSED APPEND NAME'.
          mv_append_col = lv_col.
      ENDCASE.
    ENDDO.

    IF mv_type_col = 0
      OR mv_enh_type_col = 0
      OR mv_enh_name_col = 0.
      MESSAGE
        'Required headers are missing: Data Type, Enhancement Type, Enhancement Name.'
        TYPE 'E'.
    ENDIF.

    LOOP AT <lt_excel> ASSIGNING <ls_row> FROM 2.

      lv_data_type = get_cell(
        is_row    = <ls_row>
        iv_column = mv_type_col ).

      lv_enh_type = get_cell(
        is_row    = <ls_row>
        iv_column = mv_enh_type_col ).

      lv_enh_name = get_cell(
        is_row    = <ls_row>
        iv_column = mv_enh_name_col ).

      TRANSLATE lv_data_type TO UPPER CASE.
      TRANSLATE lv_enh_type TO UPPER CASE.

      CLEAR lv_scope.
      IF mv_scope_col > 0.
        lv_scope = get_cell(
          is_row    = <ls_row>
          iv_column = mv_scope_col ).
      ENDIF.

      IF mv_release <= 2023.
        CASE lv_data_type.
          WHEN 'QUAN' OR 'CURR' OR 'CUKY' OR 'UNIT'.
            lv_scope = 'OUT'.
        ENDCASE.
      ENDIF.

      CLEAR lv_append_name.
      IF mv_append_col > 0.
        lv_append_name = get_cell(
          is_row    = <ls_row>
          iv_column = mv_append_col ).
      ENDIF.

      CASE lv_enh_type.
        WHEN 'APPEND_STRUCTURE'.
          IF lv_enh_name IS NOT INITIAL.
            lv_append_name = |{ lv_enh_name }_AMF|.
          ENDIF.
        WHEN 'INCLUDE_STRUCTURE'.
          IF lv_enh_name IS NOT INITIAL.
            lv_append_name = |Z{ lv_enh_name }_AMF|.
          ENDIF.
        WHEN 'CUSTOMIZING_INCLUDE'.
          "No naming rule supplied; preserve any existing value.
      ENDCASE.

      ls_result-amf_scope       = lv_scope.
      ls_result-proposed_append = lv_append_name.
      APPEND ls_result TO mt_results.

    ENDLOOP.

  ENDMETHOD.


  METHOD select_output_file.

    DATA:
      lv_filename TYPE string,
      lv_path     TYPE string,
      lv_fullpath TYPE string.

    cl_gui_frontend_services=>file_save_dialog(
      EXPORTING
        window_title      = 'Save Processed Excel File'
        default_extension = 'xls'
        default_file_name = 'AMF_Output.xls'
        file_filter       = 'Excel XML (*.xls)|*.xls|'
      CHANGING
        filename          = lv_filename
        path              = lv_path
        fullpath          = lv_fullpath
      EXCEPTIONS
        OTHERS            = 1 ).

    IF sy-subrc <> 0 OR lv_fullpath IS INITIAL.
      CLEAR mv_output_file.
      RETURN.
    ENDIF.

    mv_output_file = lv_fullpath.

  ENDMETHOD.


  METHOD escape_xml.

    rv_value = iv_value.

    REPLACE ALL OCCURRENCES OF '&' IN rv_value WITH '&amp;'.
    REPLACE ALL OCCURRENCES OF '<' IN rv_value WITH '&lt;'.
    REPLACE ALL OCCURRENCES OF '>' IN rv_value WITH '&gt;'.
    REPLACE ALL OCCURRENCES OF '"' IN rv_value WITH '&quot;'.
    REPLACE ALL OCCURRENCES OF '''' IN rv_value WITH '&apos;'.

  ENDMETHOD.


  METHOD download_excel.

    DATA:
      lt_lines       TYPE STANDARD TABLE OF string WITH EMPTY KEY,
      lv_value       TYPE string,
      lv_col         TYPE i,
      lv_output_cols TYPE i,
      lv_header      TYPE string.

    FIELD-SYMBOLS:
      <ls_row>    TYPE any,
      <ls_header> TYPE any,
      <lv_test>   TYPE any.

    ASSIGN <lt_excel>[ 1 ] TO <ls_header>.
    IF sy-subrc <> 0.
      MESSAGE 'No header row available for output.' TYPE 'E'.
    ENDIF.

    DO.
      ASSIGN COMPONENT sy-index OF STRUCTURE <ls_header> TO <lv_test>.
      IF sy-subrc <> 0.
        EXIT.
      ENDIF.
      lv_output_cols = sy-index.
    ENDDO.

    DATA(lv_add_scope)  = xsdbool( mv_scope_col = 0 ).
    DATA(lv_add_append) = xsdbool( mv_append_col = 0 ).
    DATA(lv_original_cols) = lv_output_cols.

    IF lv_add_scope = abap_true.
      lv_output_cols = lv_output_cols + 1.
    ENDIF.
    IF lv_add_append = abap_true.
      lv_output_cols = lv_output_cols + 1.
    ENDIF.

    APPEND '<?xml version="1.0"?>' TO lt_lines.
    APPEND '<?mso-application progid="Excel.Sheet"?>' TO lt_lines.
    APPEND
      '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"'
      && ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
      TO lt_lines.
    APPEND '<Worksheet ss:Name="AMF Output"><Table>' TO lt_lines.

    "Header row
    APPEND '<Row>' TO lt_lines.
    DO lv_output_cols TIMES.
      lv_col = sy-index.
      CLEAR lv_header.

      IF lv_col <= lv_original_cols.
        lv_header = get_cell(
          is_row    = <ls_header>
          iv_column = lv_col ).
      ELSEIF lv_add_scope = abap_true
          AND lv_col = lv_original_cols + 1.
        lv_header = 'AMF Scope'.
      ELSE.
        lv_header = 'Proposed Append Name'.
      ENDIF.

      lv_value = escape_xml( lv_header ).
      APPEND |<Cell><Data ss:Type="String">{ lv_value }</Data></Cell>| TO lt_lines.
    ENDDO.
    APPEND '</Row>' TO lt_lines.

    "Data rows
    LOOP AT <lt_excel> ASSIGNING <ls_row> FROM 2.

      READ TABLE mt_results INDEX sy-tabix - 1 INTO DATA(ls_result).
      IF sy-subrc <> 0.
        CLEAR ls_result.
      ENDIF.

      APPEND '<Row>' TO lt_lines.

      DO lv_output_cols TIMES.
        lv_col = sy-index.
        CLEAR lv_value.

        IF lv_col <= lv_original_cols.
          IF mv_scope_col > 0 AND lv_col = mv_scope_col.
            lv_value = ls_result-amf_scope.
          ELSEIF mv_append_col > 0 AND lv_col = mv_append_col.
            lv_value = ls_result-proposed_append.
          ELSE.
            lv_value = get_cell(
              is_row    = <ls_row>
              iv_column = lv_col ).
          ENDIF.
        ELSEIF lv_add_scope = abap_true
            AND lv_col = lv_original_cols + 1.
          lv_value = ls_result-amf_scope.
        ELSE.
          lv_value = ls_result-proposed_append.
        ENDIF.

        lv_value = escape_xml( lv_value ).
        APPEND |<Cell><Data ss:Type="String">{ lv_value }</Data></Cell>| TO lt_lines.
      ENDDO.

      APPEND '</Row>' TO lt_lines.
    ENDLOOP.

    APPEND '</Table></Worksheet></Workbook>' TO lt_lines.

    cl_gui_frontend_services=>gui_download(
      EXPORTING
        filename = mv_output_file
        filetype = 'ASC'
      CHANGING
        data_tab = lt_lines
      EXCEPTIONS
        OTHERS   = 1 ).

    IF sy-subrc <> 0.
      MESSAGE 'Unable to download the output file.' TYPE 'E'.
    ENDIF.

  ENDMETHOD.

ENDCLASS.
```

## Important notes

1.  The TOP include must declare the methods `run`, `read_excel`,
    `process_data`, `select_output_file`, `download_excel`, `get_cell`,
    and `escape_xml`.
2.  It must declare the column-index attributes and `mt_results`, whose
    row type has `amf_scope` and `proposed_append`.
3.  This template assumes `AMF Scope` and `Proposed Append Name` are
    either already in the input or are appended at the end. If both
    already exist, the existing columns are updated.
4.  The `CUSTOMIZING_INCLUDE` naming rule is still unspecified; the code
    preserves its existing value if present.
5.  The XML output is Excel-compatible `.xls`. For true `.xlsx`, use a
    supported XLSX writer in your SAP system.
