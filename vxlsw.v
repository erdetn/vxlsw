// Author: Erdet Nasufi 

module vxlsw

#flag -I /usr/local/include 
#flag -L /usr/local/lib 
#flag -l xlsxwriter

#include <xlsxwriter.h>


//// Workbook ////
struct C.lxw_workbook {
}

pub struct Workbook {
    filename_str string
mut:
    ptr &C.lxw_workbook
}

fn C.workbook_new(filename &char) &C.lxw_workbook
pub fn new_workbook(filename string) Workbook {
    return Workbook {
        filename_str: filename
        ptr: C.workbook_new(filename.str)
    }
}

fn C.workbook_add_worksheet(wb &C.lxw_workbook, sheet_name &char) &C.lxw_worksheet
pub fn (wb Workbook)add_sheet(sheet_name string) Worksheet {
    return Worksheet {
        ptr: C.workbook_add_worksheet(wb.ptr, sheet_name.str)
    }
}

fn C.workbook_close(wb &C.lxw_workbook) int
pub fn (wb Workbook)close() ReturnCode {
    return unsafe {
        ReturnCode(C.workbook_close(wb.ptr))
    }
}

//// Worksheet ////
struct C.lxw_worksheet {
}

pub struct Worksheet {
    ptr &C.lxw_worksheet
}

// TODO: need to fix error return type
fn C.worksheet_write_string(sheet &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, text &char, format &C.lxw_format) int
pub fn (ws Worksheet)write_string(row u32, col u16, text string, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_string(ws.ptr, C.lxw_row_t(row), C.lxw_col_t(col), text.str, format.ptr))
    }
}

struct C.lxw_format {
}

pub struct Format {
    ptr &C.lxw_format = voidptr(0)
}

/// Error ///
pub enum ReturnCode as int {
    no_error = C.LXW_NO_ERROR
    memeory_malloc_failed = C.LXW_ERROR_MEMORY_MALLOC_FAILED
    creating_xlsx_failed = C.LXW_ERROR_CREATING_XLSX_FILE
    creating_tmpfile = C.LXW_ERROR_CREATING_TMPFILE
    reading_tmpfile = C.LXW_ERROR_READING_TMPFILE
    zip_file_operation = C.LXW_ERROR_ZIP_FILE_OPERATION
    zip_parameter_error = C.LXW_ERROR_ZIP_PARAMETER_ERROR
    zip_bad_file = C.LXW_ERROR_ZIP_BAD_ZIP_FILE
    zip_internal_error = C.LXW_ERROR_ZIP_INTERNAL_ERROR
    zip_file_add = C.LXW_ERROR_ZIP_FILE_ADD
    zip_close = C.LXW_ERROR_ZIP_CLOSE
    feature_not_suppoerted = C.LXW_ERROR_FEATURE_NOT_SUPPORTED
    null_parameter_ignored = C.LXW_ERROR_NULL_PARAMETER_IGNORED
    parameter_validation = C.LXW_ERROR_PARAMETER_VALIDATION
    sheetname_length_exceeded = C.LXW_ERROR_SHEETNAME_LENGTH_EXCEEDED
    invalid_sheetname_character = C.LXW_ERROR_INVALID_SHEETNAME_CHARACTER
    sheetname_strat_end_apostrope = C.LXW_ERROR_SHEETNAME_START_END_APOSTROPHE
    sheetname_already_used = C.LXW_ERROR_SHEETNAME_ALREADY_USED
    string_length32_exceeded = C.LXW_ERROR_32_STRING_LENGTH_EXCEEDED
    string_length128_exceeded = C.LXW_ERROR_128_STRING_LENGTH_EXCEEDED
    string_length256_exceeded = C.LXW_ERROR_255_STRING_LENGTH_EXCEEDED
    max_string_length_exceeded = C.LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED
    shared_string_index_not_found = C.LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND
    worksheet_index_out_of_range = C.LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE
    worksheet_url_length_exceeded = C.LXW_ERROR_WORKSHEET_MAX_URL_LENGTH_EXCEEDED
    worksheet_max_number_urls_exceeded = C.LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED
    image_dimensions = C.LXW_ERROR_IMAGE_DIMENSIONS
}


