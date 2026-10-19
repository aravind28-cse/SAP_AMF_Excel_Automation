# ZAMF_EXCEL_AUTOMATION: Merged SAP GUI Report (AMF Scope and Proposed Append Name)

This version keeps the structure of your simpler .md file: one main class, three includes (TOP, SEL, PROC), header-driven column lookup, and all input columns carried into the output in their original order. It adds the robust parts of the earlier solution: native `.xlsx` output, an exception class, validation, and warnings.

> **Not tested on a live SAP system.** Verify the signatures of `CL_FDT_XL_SPREADSHEET`, `CL_BCS_CONVERT` and `CL_SALV_TABLE=>TO_XML` in SE24 for your release before activation. Requires NetWeaver 7.40 SP08 or higher (inline declarations, `VALUE`, `NEW`, `COND`, `xsdbool`).

## 1. What was merged

| Topic | New .md | Earlier solution | Merged choice |
|---|---|---|---|
| Columns | Reads headers, keeps all input columns, adds output columns only if missing | Fixed assumed column layout | **New .md approach.** No assumed layout needed |
| Header lookup | By name, case-insensitive | By name via a mapping table | By name, ignoring case, spaces and underscores |
| Output file | XML `.xls` (not native Excel), row writer left incomplete | Native `.xlsx` via SALV | **Native `.xlsx`**: dynamic table built with RTTS, exported with SALV `to_xml`. No XML escaping needed |
| Class design | One class | Five classes | **One class** plus a small exception class |
| Errors | `MESSAGE TYPE 'E'` inside methods, `CATCH cx_root` | Exception class `lcx_amf` | `lcx_amf` raised everywhere, handled in `run`, `cx_root` as fallback |
| Output location | Save dialog after processing | Selection-screen parameter with F4 | Selection-screen parameter with F4 (fails early) |
| Existing values | Preserved unless a rule applies | Not considered | Preserved: `OUT` is set only by Rule 1, and a name is written only when Rule 2 produces one |
| Extras | Test cases | Required-column check, name-length warning, `.xlsx` and same-file checks | All kept |

Issues in the new .md that were fixed:

- `FIELD-SYMBOLS` declared inside the class definition is not allowed in ABAP. They are now local to methods.
- The file-open F4 ignored a cancelled dialog.
- The business-rule snippets were not wired together. They are now integrated into `process_data`.

## 2. Dictionary objects (SE11)

| Object | Name | Details |
|---|---|---|
| Domain | `ZDM_S4_RELVER` | Data type `CHAR`, length 4. Fixed values: 1909, 2020, 2021, 2022, 2023, 2024, 2025 (add the ones your landscape supports, all four digits numeric) |
| Data element | `ZDE_S4_RELVER` | Domain `ZDM_S4_RELVER`, short description "S/4HANA Release Version" |

Activate both.

## 3. Program, includes and text elements (SE38)

Create executable program `ZAMF_EXCEL_AUTOMATION`, then create the three includes by double-clicking each INCLUDE statement.

Text elements:

- `TEXT-001`: "S/4HANA Release"
- `TEXT-002`: "Files"
- Selection texts: `P_RELVER` = "S/4HANA release version", `P_INFILE` = "Input Excel file (.xlsx)", `P_OUTFL` = "Output Excel file (.xlsx)"

## 4. Main program: `ZAMF_EXCEL_AUTOMATION`

```abap
REPORT zamf_excel_automation.

INCLUDE zamf_excel_automation_top.
INCLUDE zamf_excel_automation_sel.
INCLUDE zamf_excel_automation_proc.
```

## 5. TOP include: `ZAMF_EXCEL_AUTOMATION_TOP`

```abap
*&---------------------------------------------------------------------*
*& Include ZAMF_EXCEL_AUTOMATION_TOP
*&---------------------------------------------------------------------*

*&---------------------------------------------------------------------*
*& Exception class
*&---------------------------------------------------------------------*
CLASS lcx_amf DEFINITION FINAL INHERITING FROM cx_static_check.
  PUBLIC SECTION.
    METHODS constructor IMPORTING iv_text TYPE string.
    METHODS get_text REDEFINITION.
  PRIVATE SECTION.
    DATA mv_text TYPE string.
ENDCLASS.

*&---------------------------------------------------------------------*
*& Main class
*&---------------------------------------------------------------------*
CLASS lcl_amf_automation DEFINITION FINAL.

  PUBLIC SECTION.
    CLASS-METHODS:
      f4_input_file  CHANGING cv_file TYPE rlgrap-filename,
      f4_output_file CHANGING cv_file TYPE rlgrap-filename,
      validate_files IMPORTING iv_input  TYPE rlgrap-filename
                               iv_output TYPE rlgrap-filename.

    METHODS:
      constructor IMPORTING iv_release TYPE zde_s4_relver
                            iv_input   TYPE rlgrap-filename
                            iv_output  TYPE rlgrap-filename,
      run.

  PRIVATE SECTION.
    TYPES:
      ty_cells TYPE STANDARD TABLE OF string WITH EMPTY KEY,   " one row
      tt_grid  TYPE STANDARD TABLE OF ty_cells WITH EMPTY KEY. " all rows

    CONSTANTS:
      c_cutoff_release    TYPE i      VALUE 2023,
      c_out               TYPE string VALUE `OUT`,
      c_append_structure  TYPE string VALUE `APPEND_STRUCTURE`,
      c_include_structure TYPE string VALUE `INCLUDE_STRUCTURE`,
      c_prefix            TYPE string VALUE `Z`,
      c_suffix            TYPE string VALUE `_AMF`,
      c_max_name_len      TYPE i      VALUE 30,
      c_scope_header      TYPE string VALUE `AMF Scope`,
      c_append_header     TYPE string VALUE `Proposed Append Name`.

    DATA:
      mv_release      TYPE i,
      mv_input        TYPE string,
      mv_output       TYPE string,
      mt_header       TYPE ty_cells,   " header texts (input + added columns)
      mt_rows         TYPE tt_grid,    " data rows, same width as header
      mv_type_col     TYPE i,          " Data Type
      mv_enh_type_col TYPE i,          " Enhancement Type
      mv_enh_name_col TYPE i,          " Enhancement Name
      mv_scope_col    TYPE i,          " AMF Scope
      mv_append_col   TYPE i,          " Proposed Append Name
      mv_long_names   TYPE i.          " proposed names > 30 characters

    METHODS:
      " ---- reading ----
      read_excel,
      upload_file        RETURNING VALUE(rv_xstring) TYPE xstring
                         RAISING   lcx_amf,
      get_sheet          IMPORTING iv_xstring     TYPE xstring
                         RETURNING VALUE(rr_data) TYPE REF TO data
                         RAISING   lcx_amf,
      load_grid          IMPORTING ir_data TYPE REF TO data
                         RAISING   lcx_amf,
      row_to_cells       IMPORTING is_row          TYPE any
                         RETURNING VALUE(rt_cells) TYPE ty_cells,
      " ---- columns ----
      find_columns       RAISING lcx_amf,
      add_output_columns,
      " ---- business rules ----
      process_data,
      is_qty_or_curr     IMPORTING iv_type          TYPE string
                         RETURNING VALUE(rv_result) TYPE abap_bool,
      build_append_name  IMPORTING iv_enh_type    TYPE string
                                   iv_enh_name    TYPE string
                         RETURNING VALUE(rv_name) TYPE string,
      " ---- helpers ----
      get_cell           IMPORTING it_cells        TYPE ty_cells
                                   iv_col          TYPE i
                         RETURNING VALUE(rv_value) TYPE string,
      set_cell           IMPORTING iv_col   TYPE i
                                   iv_value TYPE string
                         CHANGING  ct_cells TYPE ty_cells,
      header_key         IMPORTING iv_text       TYPE string
                         RETURNING VALUE(rv_key) TYPE string,
      enh_type_key       IMPORTING iv_text       TYPE string
                         RETURNING VALUE(rv_key) TYPE string,
      " ---- writing ----
      build_xlsx         RETURNING VALUE(rv_xlsx) TYPE xstring
                         RAISING   lcx_amf,
      download_file      IMPORTING iv_xlsx TYPE xstring
                         RAISING   lcx_amf.

ENDCLASS.
```

## 6. SEL include: `ZAMF_EXCEL_AUTOMATION_SEL`

```abap
*&---------------------------------------------------------------------*
*& Include ZAMF_EXCEL_AUTOMATION_SEL
*&---------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
  PARAMETERS p_relver TYPE zde_s4_relver AS LISTBOX VISIBLE LENGTH 12
                      OBLIGATORY DEFAULT '2023'.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-002.
  PARAMETERS: p_infile TYPE rlgrap-filename OBLIGATORY,   " input .xlsx
              p_outfl  TYPE rlgrap-filename OBLIGATORY.   " output .xlsx
SELECTION-SCREEN END OF BLOCK b2.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_infile.
  lcl_amf_automation=>f4_input_file( CHANGING cv_file = p_infile ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_outfl.
  lcl_amf_automation=>f4_output_file( CHANGING cv_file = p_outfl ).

AT SELECTION-SCREEN.
  lcl_amf_automation=>validate_files( iv_input  = p_infile
                                      iv_output = p_outfl ).
```

## 7. PROC include: `ZAMF_EXCEL_AUTOMATION_PROC`

```abap
*&---------------------------------------------------------------------*
*& Include ZAMF_EXCEL_AUTOMATION_PROC
*&---------------------------------------------------------------------*

*&---------------------------------------------------------------------*
*& lcx_amf
*&---------------------------------------------------------------------*
CLASS lcx_amf IMPLEMENTATION.
  METHOD constructor.
    super->constructor( ).
    mv_text = iv_text.
  ENDMETHOD.

  METHOD get_text.
    result = mv_text.
  ENDMETHOD.
ENDCLASS.

*&---------------------------------------------------------------------*
*& lcl_amf_automation
*&---------------------------------------------------------------------*
CLASS lcl_amf_automation IMPLEMENTATION.

  "---------------------------------------------------------------
  " Selection-screen helpers (static)
  "---------------------------------------------------------------
  METHOD f4_input_file.
    DATA: lt_files TYPE filetable,
          lv_rc    TYPE i,
          lv_act   TYPE i.

    cl_gui_frontend_services=>file_open_dialog(
      EXPORTING window_title      = 'Select input Excel file'
                default_extension = 'xlsx'
                file_filter       = 'Excel Files (*.xlsx)|*.xlsx|'
      CHANGING  file_table        = lt_files
                rc                = lv_rc
                user_action       = lv_act
      EXCEPTIONS OTHERS           = 1 ).

    IF sy-subrc = 0
       AND lv_act = cl_gui_frontend_services=>action_ok
       AND lv_rc > 0.
      cv_file = lt_files[ 1 ]-filename.
    ENDIF.
  ENDMETHOD.

  METHOD f4_output_file.
    DATA: lv_name TYPE string,
          lv_path TYPE string,
          lv_full TYPE string,
          lv_act  TYPE i.

    cl_gui_frontend_services=>file_save_dialog(
      EXPORTING window_title      = 'Save processed Excel file'
                default_extension = 'xlsx'
                default_file_name = 'AMF_Output.xlsx'
                file_filter       = 'Excel Files (*.xlsx)|*.xlsx|'
      CHANGING  filename          = lv_name
                path              = lv_path
                fullpath          = lv_full
                user_action       = lv_act
      EXCEPTIONS OTHERS           = 1 ).

    IF sy-subrc = 0
       AND lv_act = cl_gui_frontend_services=>action_ok
       AND lv_full IS NOT INITIAL.
      cv_file = lv_full.
    ENDIF.
  ENDMETHOD.

  METHOD validate_files.
    IF NOT matches( val = to_lower( iv_input )  regex = '.*\.xlsx' )
    OR NOT matches( val = to_lower( iv_output ) regex = '.*\.xlsx' ).
      MESSAGE 'Input and output files must be .xlsx files' TYPE 'E'.
    ENDIF.
    IF to_lower( iv_input ) = to_lower( iv_output ).
      MESSAGE 'Output file must be different from the input file' TYPE 'E'.
    ENDIF.
  ENDMETHOD.

  "---------------------------------------------------------------
  " Main flow
  "---------------------------------------------------------------
  METHOD constructor.
    mv_release = CONV i( iv_release ).
    mv_input   = iv_input.
    mv_output  = iv_output.
  ENDMETHOD.

  METHOD run.
    TRY.
        read_excel( ).            " 1. Excel -> mt_header / mt_rows
        find_columns( ).          " 2. locate required columns
        add_output_columns( ).    " 3. add AMF Scope / Proposed Append Name if missing
        process_data( ).          " 4. apply business rules
        download_file( build_xlsx( ) ).   " 5. write .xlsx

        IF mv_long_names > 0.
          MESSAGE |Output saved ({ lines( mt_rows ) } rows). | &&
                  |{ mv_long_names } proposed name(s) exceed 30 characters - please review.|
                  TYPE 'S' DISPLAY LIKE 'W'.
        ELSE.
          MESSAGE |Output saved: { lines( mt_rows ) } rows processed.| TYPE 'S'.
        ENDIF.

      CATCH lcx_amf INTO DATA(lx_amf).
        MESSAGE lx_amf->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
      CATCH cx_root INTO DATA(lx_root).
        MESSAGE lx_root->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
    ENDTRY.
  ENDMETHOD.

  "---------------------------------------------------------------
  " Reading
  "---------------------------------------------------------------
  METHOD read_excel.
    load_grid( get_sheet( upload_file( ) ) ).
  ENDMETHOD.

  METHOD upload_file.
    DATA: lt_bin TYPE solix_tab,
          lv_len TYPE i.

    cl_gui_frontend_services=>gui_upload(
      EXPORTING filename   = mv_input
                filetype   = 'BIN'
      IMPORTING filelength = lv_len
      CHANGING  data_tab   = lt_bin
      EXCEPTIONS OTHERS    = 1 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION NEW lcx_amf( iv_text = 'Unable to read the input Excel file' ).
    ENDIF.

    rv_xstring = cl_bcs_convert=>solix_to_xstring( it_solix = lt_bin
                                                   iv_size  = lv_len ).
  ENDMETHOD.

  METHOD get_sheet.
    DATA lt_sheets TYPE if_fdt_doc_spreadsheet=>t_worksheet_names.

    TRY.
        DATA(lo_excel) = NEW cl_fdt_xl_spreadsheet( document_name = mv_input
                                                    xdocument     = iv_xstring ).
        lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
          IMPORTING worksheet_names = lt_sheets ).
        IF lt_sheets IS INITIAL.
          RAISE EXCEPTION NEW lcx_amf( iv_text = 'No worksheet found in the Excel file' ).
        ENDIF.
        rr_data = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet(
                    lt_sheets[ 1 ] ).                  " first worksheet
      CATCH cx_fdt_excel_core INTO DATA(lx_excel).
        RAISE EXCEPTION NEW lcx_amf( iv_text = lx_excel->get_text( ) ).
    ENDTRY.
  ENDMETHOD.

  METHOD load_grid.
    FIELD-SYMBOLS <lt_sheet> TYPE STANDARD TABLE.
    DATA lv_cols TYPE i.

    ASSIGN ir_data->* TO <lt_sheet>.
    IF sy-subrc <> 0 OR lines( <lt_sheet> ) < 2.
      RAISE EXCEPTION NEW lcx_amf( iv_text = 'The Excel worksheet has no data rows' ).
    ENDIF.

    LOOP AT <lt_sheet> ASSIGNING FIELD-SYMBOL(<ls_row>).
      DATA(lv_row)   = sy-tabix.
      DATA(lt_cells) = row_to_cells( <ls_row> ).

      IF lv_row = 1.                                   " header row
        mt_header = lt_cells.
        WHILE lines( mt_header ) > 0
          AND mt_header[ lines( mt_header ) ] IS INITIAL.   " drop trailing blanks
          DELETE mt_header INDEX lines( mt_header ).
        ENDWHILE.
        IF mt_header IS INITIAL.
          RAISE EXCEPTION NEW lcx_amf( iv_text = 'The header row of the Excel file is empty' ).
        ENDIF.
        lv_cols = lines( mt_header ).
        CONTINUE.
      ENDIF.

      " skip completely blank rows
      DATA(lv_filled) = abap_false.
      LOOP AT lt_cells TRANSPORTING NO FIELDS WHERE table_line IS NOT INITIAL.
        lv_filled = abap_true.
        EXIT.
      ENDLOOP.
      IF lv_filled = abap_false.
        CONTINUE.
      ENDIF.

      " make every row exactly as wide as the header
      WHILE lines( lt_cells ) < lv_cols.
        APPEND `` TO lt_cells.
      ENDWHILE.
      IF lines( lt_cells ) > lv_cols.
        DELETE lt_cells FROM lv_cols + 1.
      ENDIF.

      APPEND lt_cells TO mt_rows.
    ENDLOOP.

    IF mt_rows IS INITIAL.
      RAISE EXCEPTION NEW lcx_amf( iv_text = 'No usable data rows found in the Excel file' ).
    ENDIF.
  ENDMETHOD.

  METHOD row_to_cells.
    FIELD-SYMBOLS <lv_cell> TYPE any.

    DO.
      ASSIGN COMPONENT sy-index OF STRUCTURE is_row TO <lv_cell>.
      IF sy-subrc <> 0.
        EXIT.
      ENDIF.
      APPEND condense( |{ <lv_cell> }| ) TO rt_cells.
    ENDDO.
  ENDMETHOD.

  "---------------------------------------------------------------
  " Columns
  "---------------------------------------------------------------
  METHOD find_columns.
    LOOP AT mt_header INTO DATA(lv_head).
      DATA(lv_idx) = sy-tabix.
      CASE header_key( lv_head ).
        WHEN 'DATATYPE'.
          IF mv_type_col = 0.     mv_type_col     = lv_idx. ENDIF.
        WHEN 'ENHANCEMENTTYPE'.
          IF mv_enh_type_col = 0. mv_enh_type_col = lv_idx. ENDIF.
        WHEN 'ENHANCEMENTNAME'.
          IF mv_enh_name_col = 0. mv_enh_name_col = lv_idx. ENDIF.
        WHEN 'AMFSCOPE'.
          IF mv_scope_col = 0.    mv_scope_col    = lv_idx. ENDIF.
        WHEN 'PROPOSEDAPPENDNAME'.
          IF mv_append_col = 0.   mv_append_col   = lv_idx. ENDIF.
      ENDCASE.
    ENDLOOP.

    IF mv_type_col = 0.
      RAISE EXCEPTION NEW lcx_amf( iv_text = 'Required column missing in Excel header: Data Type' ).
    ENDIF.
    IF mv_enh_type_col = 0.
      RAISE EXCEPTION NEW lcx_amf( iv_text = 'Required column missing in Excel header: Enhancement Type' ).
    ENDIF.
    IF mv_enh_name_col = 0.
      RAISE EXCEPTION NEW lcx_amf( iv_text = 'Required column missing in Excel header: Enhancement Name' ).
    ENDIF.
  ENDMETHOD.

  METHOD add_output_columns.
    IF mv_scope_col = 0.
      APPEND c_scope_header TO mt_header.
      mv_scope_col = lines( mt_header ).
      LOOP AT mt_rows ASSIGNING FIELD-SYMBOL(<lt_row_s>).
        APPEND `` TO <lt_row_s>.
      ENDLOOP.
    ENDIF.

    IF mv_append_col = 0.
      APPEND c_append_header TO mt_header.
      mv_append_col = lines( mt_header ).
      LOOP AT mt_rows ASSIGNING FIELD-SYMBOL(<lt_row_a>).
        APPEND `` TO <lt_row_a>.
      ENDLOOP.
    ENDIF.
  ENDMETHOD.

  "---------------------------------------------------------------
  " Business rules
  "---------------------------------------------------------------
  METHOD process_data.
    LOOP AT mt_rows ASSIGNING FIELD-SYMBOL(<lt_cells>).

      " Rule 1: AMF Scope = OUT for quantity/currency types, release <= 2023
      IF mv_release <= c_cutoff_release
         AND is_qty_or_curr( get_cell( it_cells = <lt_cells>
                                       iv_col   = mv_type_col ) ) = abap_true.
        set_cell( EXPORTING iv_col   = mv_scope_col
                            iv_value = c_out
                  CHANGING  ct_cells = <lt_cells> ).
      ENDIF.

      " Rule 2: Proposed Append Name from Enhancement Type / Name
      DATA(lv_name) = build_append_name(
                        iv_enh_type = get_cell( it_cells = <lt_cells>
                                                iv_col   = mv_enh_type_col )
                        iv_enh_name = get_cell( it_cells = <lt_cells>
                                                iv_col   = mv_enh_name_col ) ).
      IF lv_name IS NOT INITIAL.
        IF strlen( lv_name ) > c_max_name_len.        " DDIC names: max 30 chars
          mv_long_names = mv_long_names + 1.
        ENDIF.
        set_cell( EXPORTING iv_col   = mv_append_col
                            iv_value = lv_name
                  CHANGING  ct_cells = <lt_cells> ).
      ENDIF.

    ENDLOOP.
  ENDMETHOD.

  METHOD is_qty_or_curr.
    DATA(lv_type) = to_upper( condense( iv_type ) ).
    rv_result = xsdbool( lv_type = 'QUAN' OR lv_type = 'CURR'
                      OR lv_type = 'CUKY' OR lv_type = 'UNIT' ).
  ENDMETHOD.

  METHOD build_append_name.
    DATA(lv_name) = condense( iv_enh_name ).
    IF lv_name IS INITIAL.
      RETURN.
    ENDIF.

    CASE enh_type_key( iv_enh_type ).
      WHEN c_append_structure.                         " ZST_MARA  -> ZST_MARA_AMF
        rv_name = |{ lv_name }{ c_suffix }|.
      WHEN c_include_structure.                        " CI_EKKODB -> ZCI_EKKODB_AMF
        rv_name = |{ c_prefix }{ lv_name }{ c_suffix }|.
      WHEN OTHERS.
        " CUSTOMIZING_INCLUDE: naming rule not supplied yet -> left blank.
        " To treat it like INCLUDE_STRUCTURE use:
        "   WHEN c_include_structure OR 'CUSTOMIZING_INCLUDE'.
    ENDCASE.
  ENDMETHOD.

  "---------------------------------------------------------------
  " Helpers
  "---------------------------------------------------------------
  METHOD get_cell.
    rv_value = VALUE #( it_cells[ iv_col ] OPTIONAL ).
  ENDMETHOD.

  METHOD set_cell.
    MODIFY ct_cells FROM iv_value INDEX iv_col.
  ENDMETHOD.

  METHOD header_key.
    " 'Data Type' / 'DATA_TYPE' / 'data type' -> 'DATATYPE'
    rv_key = to_upper( iv_text ).
    REPLACE ALL OCCURRENCES OF REGEX '[ _]' IN rv_key WITH ''.
  ENDMETHOD.

  METHOD enh_type_key.
    " 'APPEND STRUCTURE' / 'append_structure' -> 'APPEND_STRUCTURE'
    rv_key = to_upper( condense( iv_text ) ).
    REPLACE ALL OCCURRENCES OF ` ` IN rv_key WITH `_`.
  ENDMETHOD.

  "---------------------------------------------------------------
  " Writing: dynamic table (RTTS) -> SALV -> native .xlsx
  "---------------------------------------------------------------
  METHOD build_xlsx.
    DATA: lt_comp  TYPE cl_abap_structdescr=>component_table,
          lr_table TYPE REF TO data.
    FIELD-SYMBOLS: <lt_out>  TYPE STANDARD TABLE,
                   <lv_cell> TYPE any.

    " 1. dynamic structure: one string component per Excel column
    DO lines( mt_header ) TIMES.
      APPEND VALUE #( name = |C{ sy-index WIDTH = 3 ALIGN = RIGHT PAD = '0' }|
                      type = cl_abap_elemdescr=>get_string( ) ) TO lt_comp.
    ENDDO.

    TRY.
        DATA(lo_struct) = cl_abap_structdescr=>create( lt_comp ).
        DATA(lo_tab)    = cl_abap_tabledescr=>create( lo_struct ).
      CATCH cx_sy_struct_creation cx_sy_table_creation INTO DATA(lx_rtts).
        RAISE EXCEPTION NEW lcx_amf( iv_text = lx_rtts->get_text( ) ).
    ENDTRY.

    CREATE DATA lr_table TYPE HANDLE lo_tab.
    ASSIGN lr_table->* TO <lt_out>.

    " 2. fill it from the processed grid
    LOOP AT mt_rows INTO DATA(lt_cells).
      APPEND INITIAL LINE TO <lt_out> ASSIGNING FIELD-SYMBOL(<ls_out>).
      LOOP AT lt_cells INTO DATA(lv_value).
        DATA(lv_idx) = sy-tabix.
        ASSIGN COMPONENT lv_idx OF STRUCTURE <ls_out> TO <lv_cell>.
        IF sy-subrc = 0.
          <lv_cell> = lv_value.
        ENDIF.
      ENDLOOP.
    ENDLOOP.

    " 3. SALV with the Excel header texts, export as .xlsx
    TRY.
        cl_salv_table=>factory( IMPORTING r_salv_table = DATA(lo_alv)
                                CHANGING  t_table      = <lt_out> ).
        DATA(lo_cols) = lo_alv->get_columns( ).
        lo_cols->set_optimize( abap_true ).

        LOOP AT mt_header INTO DATA(lv_head).
          DATA(lv_col)     = sy-tabix.
          DATA(lv_colname) = |C{ lv_col WIDTH = 3 ALIGN = RIGHT PAD = '0' }|.
          DATA(lv_text)    = COND string( WHEN lv_head IS INITIAL
                                          THEN |Column { lv_col }|
                                          ELSE lv_head ).
          DATA(lo_col) = lo_cols->get_column( CONV #( lv_colname ) ).
          lo_col->set_long_text(   CONV #( lv_text ) ).   " cut to 40 chars
          lo_col->set_medium_text( CONV #( lv_text ) ).   " cut to 20 chars
          lo_col->set_short_text(  CONV #( lv_text ) ).   " cut to 10 chars
        ENDLOOP.

        rv_xlsx = lo_alv->to_xml( if_salv_bs_xml=>c_type_xlsx ).
      CATCH cx_salv_msg cx_salv_not_found INTO DATA(lx_salv).
        RAISE EXCEPTION NEW lcx_amf( iv_text = lx_salv->get_text( ) ).
    ENDTRY.
  ENDMETHOD.

  METHOD download_file.
    DATA(lt_solix) = cl_bcs_convert=>xstring_to_solix( iv_xstring = iv_xlsx ).

    cl_gui_frontend_services=>gui_download(
      EXPORTING bin_filesize = xstrlen( iv_xlsx )
                filename     = mv_output
                filetype     = 'BIN'
      CHANGING  data_tab     = lt_solix
      EXCEPTIONS OTHERS      = 1 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION NEW lcx_amf( iv_text = 'Unable to save the output Excel file' ).
    ENDIF.
  ENDMETHOD.

ENDCLASS.

*&---------------------------------------------------------------------*
*& Report execution
*&---------------------------------------------------------------------*
START-OF-SELECTION.
  DATA(go_app) = NEW lcl_amf_automation( iv_release = p_relver
                                         iv_input   = p_infile
                                         iv_output  = p_outfl ).
  go_app->run( ).
```

## 8. How it works

1. The `.xlsx` is uploaded and parsed in memory (no OLE, Excel not needed). Row 1 is the header and the first worksheet is used.
2. The header names are matched ignoring case, spaces and underscores, so "Data Type", "DATA_TYPE" and "data type" all work. The three required columns are Data Type, Enhancement Type and Enhancement Name. The report stops with a clear message if one is missing.
3. All input columns are kept in their original order. `AMF Scope` and `Proposed Append Name` are added at the end only if the input does not already have them.
4. Rule 1 sets `AMF Scope` to `OUT` when the selected release is <= 2023 and Data Type is QUAN, CURR, CUKY or UNIT. For later releases nothing is set. Existing values are never cleared.
5. Rule 2 writes `<name>_AMF` for APPEND_STRUCTURE and `Z<name>_AMF` for INCLUDE_STRUCTURE. A value typed as "APPEND STRUCTURE" with a space also works.
6. The result is exported as a native `.xlsx` through a dynamically built table and SALV.

## 9. Test cases

| Release | Data Type | Enhancement Type | Enhancement Name | Expected result |
|---|---|---|---|---|
| 2023 | QUAN | APPEND_STRUCTURE | ZST_MARA | Scope `OUT`; append `ZST_MARA_AMF` |
| 2022 | CURR | APPEND_STRUCTURE | ZST_BSEG | Scope `OUT`; append `ZST_BSEG_AMF` |
| 2023 | UNIT | INCLUDE_STRUCTURE | CI_EKKODB | Scope `OUT`; append `ZCI_EKKODB_AMF` |
| 2025 | QUAN | APPEND_STRUCTURE | ZST_MARA | Scope blank; append `ZST_MARA_AMF` |
| 2023 | CHAR | INCLUDE_STRUCTURE | CI_CUSTOM | Scope blank; append `ZCI_CUSTOM_AMF` |
| 2023 | CHAR | CUSTOMIZING_INCLUDE | CI_CUSTOM | Scope blank; append blank (rule pending) |

## 10. Open points

- **CUSTOMIZING_INCLUDE:** no naming rule was given, so the proposed name stays blank. See the comment in `build_append_name` to change this.
- **Output columns:** no assumed layout is needed any more, but if your output has extra columns beyond AMF Scope and Proposed Append Name, send the list and their rules.
- **Data Type values:** the match is exact (e.g. `QUAN`). If the file holds values like `QUAN(13,3)`, tell me and I'll switch to a prefix match.
- **Header length:** headers longer than 40 characters are cut in the output because of the SALV limit.
- **Name length:** DDIC names are limited to 30 characters. Longer proposed names are written as they are, and the final message warns with a count.
- **Format and mode:** only `.xlsx` is supported (not `.xls`), and the report must run in the foreground because it uses frontend upload and download.
