# SAP ABAP Excel Automation -- AMF Scope and Proposed Append Name

## 1. Objective

Build a classic SAP GUI executable report that:

1.  Reads a user-selected Excel workbook containing custom-field
    metadata.
2.  Applies the AMF Scope rule based on the selected SAP S/4HANA
    release.
3.  Derives a Proposed Append Name from Enhancement Type and Enhancement
    Name.
4.  Writes all original input columns plus the required output columns
    to an Excel-compatible output file.

The proposed report uses three includes and local classes:

-   `ZAMF_EXCEL_AUTOMATION_TOP` --- types, data, and local class
    declaration.
-   `ZAMF_EXCEL_AUTOMATION_SEL` --- selection screen and input file
    picker.
-   `ZAMF_EXCEL_AUTOMATION_PROC` --- report execution and class
    implementation.

> **Implementation status:** This document describes the design and
> provides the key ABAP building blocks. The sample snippets are not yet
> a fully integrated, activation-ready report. In particular, the
> dynamic output-table construction and XML row generation must be
> completed against the exact Excel columns you provide. Check the API
> signatures in SE24 for your SAP release before activation.

## 2. Business rules

  ---------------------------------------------------------------------------
  Condition               Output column           Expected behavior
  ----------------------- ----------------------- ---------------------------
  Selected release is     AMF Scope               Set to `OUT`
  2023 or earlier and                             
  Data Type is `QUAN`                             

  Selected release is     AMF Scope               Set to `OUT`
  2023 or earlier and                             
  Data Type is `CURR`,                            
  `CUKY`, or `UNIT`                               

  Selected release is     AMF Scope               Do not apply the
  later than 2023                                 release-dependent rule

  Enhancement Type is     Proposed Append Name    `<Enhancement Name>_AMF`
  `APPEND_STRUCTURE`                              

  Enhancement Type is     Proposed Append Name    `Z<Enhancement Name>_AMF`
  `INCLUDE_STRUCTURE`                             

  Enhancement Type is     Proposed Append Name    Rule not supplied yet
  `CUSTOMIZING_INCLUDE`                           
  ---------------------------------------------------------------------------

Assumptions:

-   The first worksheet contains the data.
-   The first row contains column headers.
-   The relevant technical data types appear as names such as `QUAN`,
    `CURR`, `CUKY`, and `UNIT`.
-   Existing AMF Scope values are preserved unless the rule sets the
    value to `OUT`.
-   Input is `.xlsx`. The legacy `.xls` format needs a different reader
    path.

## 3. Create the release domain and data element

In SE11, create the following DDIC objects.

### Domain: `ZDM_S4_RELVER`

-   Data type: `CHAR`
-   Length: `4`
-   Fixed values: add the releases supported in your landscape, for
    example `1909`, `2020`, `2021`, `2022`, `2023`, and `2025`.

### Data element: `ZDE_S4_RELVER`

-   Domain: `ZDM_S4_RELVER`
-   Short description: `S/4HANA Release Version`

Activate both objects.

The report converts the selected release to an integer for the `<= 2023`
comparison. Ensure that every fixed value is a four-character numeric
release value.

## 4. Create the report and includes

Create executable program `ZAMF_EXCEL_AUTOMATION` in SE38 or ADT.

``` abap
REPORT zamf_excel_automation.

INCLUDE zamf_excel_automation_top.
INCLUDE zamf_excel_automation_sel.
INCLUDE zamf_excel_automation_proc.
```

## 5. TOP include --- `ZAMF_EXCEL_AUTOMATION_TOP`

The class separates report execution, reading, business processing, and
downloading.

``` abap
CLASS lcl_amf_automation DEFINITION FINAL.

  PUBLIC SECTION.
    METHODS run.

  PRIVATE SECTION.

    TYPES:
      ty_string_tab TYPE STANDARD TABLE OF string WITH EMPTY KEY,
      BEGIN OF ty_result,
        amf_scope       TYPE string,
        proposed_append TYPE string,
      END OF ty_result,
      ty_result_tab TYPE STANDARD TABLE OF ty_result WITH EMPTY KEY.

    DATA:
      mv_input_file  TYPE string,
      mv_output_file TYPE string,
      mv_release     TYPE i,
      mr_excel_data  TYPE REF TO data,
      mt_sheet_names TYPE STANDARD TABLE OF string WITH EMPTY KEY,
      mt_results     TYPE ty_result_tab,
      mt_output      TYPE ty_string_tab,
      mv_type_col       TYPE i,
      mv_enh_type_col   TYPE i,
      mv_enh_name_col   TYPE i,
      mv_scope_col      TYPE i,
      mv_append_col     TYPE i.

    FIELD-SYMBOLS:
      <lt_excel> TYPE STANDARD TABLE,
      <ls_excel> TYPE any,
      <lv_cell>  TYPE any.

    METHODS:
      select_output_file,
      read_excel,
      process_data,
      download_excel,
      get_cell
        IMPORTING
          is_row    TYPE any
          iv_column TYPE i
        RETURNING VALUE(rv_value) TYPE string,
      escape_xml
        IMPORTING iv_value TYPE string
        RETURNING VALUE(rv_value) TYPE string.

ENDCLASS.
```

## 6. SEL include --- `ZAMF_EXCEL_AUTOMATION_SEL`

``` abap
PARAMETERS:
  p_relver TYPE zde_s4_relver OBLIGATORY DEFAULT '2023',
  p_infile TYPE rlgrap-filename OBLIGATORY.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_infile.

  DATA:
    lt_files TYPE filetable,
    lv_rc    TYPE i.

  cl_gui_frontend_services=>file_open_dialog(
    EXPORTING
      window_title      = 'Select Input Excel File'
      default_extension = 'xlsx'
      file_filter       = 'Excel Files (*.xlsx)|*.xlsx|'
    CHANGING
      file_table        = lt_files
      rc                = lv_rc
    EXCEPTIONS
      OTHERS            = 1 ).

  IF sy-subrc = 0 AND lv_rc > 0.
    p_infile = lt_files[ 1 ]-filename.
  ENDIF.
```

The release parameter uses the data element created above. The file
picker allows the user to browse for a local `.xlsx` workbook.

## 7. PROC include --- `ZAMF_EXCEL_AUTOMATION_PROC`

### 7.1 Report execution

``` abap
START-OF-SELECTION.

  DATA(lo_app) = NEW lcl_amf_automation( ).
  lo_app->run( ).
```

### 7.2 Main method

``` abap
CLASS lcl_amf_automation IMPLEMENTATION.

  METHOD run.

    mv_input_file = p_infile.
    mv_release    = CONV i( p_relver ).

    TRY.
        read_excel( ).
        process_data( ).
        select_output_file( ).

        IF mv_output_file IS NOT INITIAL.
          download_excel( ).
          MESSAGE 'Excel processing completed' TYPE 'S'.
        ENDIF.

      CATCH cx_root INTO DATA(lx_error).
        MESSAGE lx_error->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
    ENDTRY.

  ENDMETHOD.
```

### 7.3 Read an XLSX workbook

This example uses `CL_FDT_XL_SPREADSHEET` to read the first worksheet.
Availability and method signatures can vary by SAP release; verify in
SE24.

``` abap
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
      MESSAGE 'Unable to read the input Excel file' TYPE 'E'.
    ENDIF.

    CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
      EXPORTING
        input_length = lv_length
      IMPORTING
        buffer       = lv_xstring
      TABLES
        binary_tab   = lt_binary
      EXCEPTIONS
        OTHERS      = 1.

    IF sy-subrc <> 0.
      MESSAGE 'Unable to convert Excel binary data' TYPE 'E'.
    ENDIF.

    lo_excel = NEW cl_fdt_xl_spreadsheet(
      document_name = mv_input_file
      xdocument     = lv_xstring ).

    lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
      IMPORTING worksheet_names = mt_sheet_names ).

    IF mt_sheet_names IS INITIAL.
      MESSAGE 'No worksheet found in Excel file' TYPE 'E'.
    ENDIF.

    mr_excel_data =
      lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet(
        worksheet_name = mt_sheet_names[ 1 ] ).

    ASSIGN mr_excel_data->* TO <lt_excel>.

    IF sy-subrc <> 0 OR <lt_excel> IS INITIAL.
      MESSAGE 'Excel worksheet has no data' TYPE 'E'.
    ENDIF.

  ENDMETHOD.
```

### 7.4 Get a cell value by column index

``` abap
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
```

### 7.5 Business processing logic

The logic should discover header positions from the first row, using
case-insensitive comparisons and trimming whitespace. The input columns
required are:

-   `Data Type`
-   `Enhancement Type`
-   `Enhancement Name`

The output columns are:

-   `AMF Scope`
-   `Proposed Append Name`

Core rule logic:

``` abap
"Convert header values to uppercase before comparing.
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
```

For each data row, beginning at row 2:

``` abap
"Rule 1: apply only for S/4HANA releases <= 2023.
IF mv_release <= 2023.
  CASE lv_data_type.
    WHEN 'QUAN' OR 'CURR' OR 'CUKY' OR 'UNIT'.
      lv_amf_scope = 'OUT'.
  ENDCASE.
ENDIF.

"Rule 2: proposed append name.
CLEAR lv_append_name.

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
    "No naming rule supplied yet.
    CLEAR lv_append_name.
ENDCASE.
```

**Important:** these are the business-rule building blocks, not a
complete standalone `process_data` method. Store the result for every
row in `mt_results` and use it while constructing the output workbook.
Do not assume a dynamic table can be extended by assigning a component
index that does not exist.

### 7.6 XML escaping helper

For an XML-based Excel-compatible output, escape reserved characters.

``` abap
  METHOD escape_xml.

    rv_value = iv_value.

    REPLACE ALL OCCURRENCES OF '&'
      IN rv_value WITH '&amp;'.
    REPLACE ALL OCCURRENCES OF '<'
      IN rv_value WITH '&lt;'.
    REPLACE ALL OCCURRENCES OF '>'
      IN rv_value WITH '&gt;'.
    REPLACE ALL OCCURRENCES OF '"'
      IN rv_value WITH '&quot;'.
    REPLACE ALL OCCURRENCES OF ''''
      IN rv_value WITH '&apos;'.

  ENDMETHOD.
```

### 7.7 Select output file

This example uses an Excel-compatible XML workbook with an `.xls`
extension. It is **not** a native `.xlsx` file.

``` abap
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
```

### 7.8 Output writer outline

The writer must:

1.  Read each original header.
2.  Keep all original input columns in their original order.
3.  Add `AMF Scope` only if it does not already exist.
4.  Add `Proposed Append Name` only if it does not already exist.
5.  For each data row, write original values except where a business
    rule updates the output value.
6.  XML-escape each cell.
7.  Download the generated lines using
    `CL_GUI_FRONTEND_SERVICES=>GUI_DOWNLOAD`.

The XML workbook envelope is:

``` abap
APPEND '<?xml version="1.0"?>' TO lt_lines.
APPEND '<?mso-application progid="Excel.Sheet"?>' TO lt_lines.
APPEND
  '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"'
  && ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
  TO lt_lines.
APPEND '<Worksheet ss:Name="AMF Output"><Table>' TO lt_lines.

"Generate header row and every data row here.
"Each cell:
"<Cell><Data ss:Type="String">escaped value</Data></Cell>

APPEND '</Table></Worksheet></Workbook>' TO lt_lines.

cl_gui_frontend_services=>gui_download(
  EXPORTING
    filename = mv_output_file
    filetype = 'ASC'
  CHANGING
    data_tab = lt_lines
  EXCEPTIONS
    OTHERS = 1 ).
```

The row-generation logic must be completed and tested with the final
list of output columns. If native `.xlsx` output is mandatory, use an
approved XLSX writer available in your SAP system rather than writing
SpreadsheetML XML with an `.xlsx` extension.

## 8. Test cases

  ----------------------------------------------------------------------------------------
  Release        Data Type      Enhancement Type        Enhancement    Expected result
                                                        Name           
  -------------- -------------- ----------------------- -------------- -------------------
  2023           `QUAN`         `APPEND_STRUCTURE`      `ZST_MARA`     Scope `OUT`; append
                                                                       `ZST_MARA_AMF`

  2022           `CURR`         `APPEND_STRUCTURE`      `ZST_BSEG`     Scope `OUT`; append
                                                                       `ZST_BSEG_AMF`

  2023           `UNIT`         `INCLUDE_STRUCTURE`     `CI_EKKODB`    Scope `OUT`; append
                                                                       `ZCI_EKKODB_AMF`

  2025           `QUAN`         `APPEND_STRUCTURE`      `ZST_MARA`     Do not apply the
                                                                       release-dependent
                                                                       `OUT` rule

  2023           `CHAR`         `INCLUDE_STRUCTURE`     `CI_CUSTOM`    Append
                                                                       `ZCI_CUSTOM_AMF`

  2023           `CHAR`         `CUSTOMIZING_INCLUDE`   `CI_CUSTOM`    Pending naming rule
  ----------------------------------------------------------------------------------------

## 9. Production considerations

-   Confirm whether nonmatching `AMF Scope` values should be preserved
    or cleared.
-   Confirm the naming rule for `CUSTOMIZING_INCLUDE`.
-   Confirm all output-only columns and their calculation rules.
-   Validate required headers and report clear errors when they are
    missing.
-   Handle blank rows, duplicate headers, special characters, and large
    files.
-   Test on the target SAP release because `CL_FDT_XL_SPREADSHEET`
    signatures and availability vary.
-   `GUI_UPLOAD`, file dialogs, and `GUI_DOWNLOAD` require SAP GUI
    frontend access, so this report is intended for foreground
    execution.

## 10. Details needed to finalize the report

To turn these building blocks into a complete, activation-ready
implementation, provide:

1.  The complete output Excel column list, marking which columns are
    copied and which are populated by logic.
2.  The naming rule for `CUSTOMIZING_INCLUDE`.
3.  Confirmation that `.xlsx` output is mandatory, or whether an
    Excel-compatible XML `.xls` file is acceptable.
4.  A sample input workbook with anonymized data, if possible.
