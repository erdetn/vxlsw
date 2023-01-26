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

fn C.lxw_format_new() &C.lxw_format 
pub fn new_format() Format {
    return Format {
        ptr: C.lxw_format_new()
    }
}

fn C.lxw_format_free(format &C.lxw_format)
pub fn (mut f Format)deallocate() {
    C.lxw_format_free(f.ptr)
}

fn C.format_set_font_name(format &C.lxw_format, font_name &char)
pub fn (f Format)set_font_name(font_name string) {
    C.format_set_font_name(f.ptr, &char(font_name.str))
}

fn C.format_set_font_size(format &C.lxw_format, size f64)
pub fn (f Format)set_font_size(size u8) {
    C.format_set_font_size(f.ptr, f64(size))
}

pub enum FormatUnderlines as u32 {
    @none = C.LXW_UNDERLINE_NONE
    single = C.LXW_UNDERLINE_SINGLE
    double = C.LXW_UNDERLINE_DOUBLE
    single_accounting = C.LXW_UNDERLINE_SINGLE_ACCOUNTING
    double_accounting = C.LXW_UNDERLINE_DOUBLE_ACCOUNTING
}

pub enum FormatScrips as u32 {
    superscript = C.LXW_FONT_SUPERSCRIPT
    subscript = C.LXW_FONT_SUBSCRIPT
}

pub enum FormatAlignments as u32 {
    @none = C.LXW_ALIGN_NONE
    left = C.LXW_ALIGN_LEFT
    center = C.LXW_ALIGN_CENTER
    right = C.LXW_ALIGN_RIGHT
    fill = C.LXW_ALIGN_FILL
    justify = C.LXW_ALIGN_JUSTIFY
    across = C.LXW_ALIGN_CENTER_ACROSS
    distributed = C.LXW_ALIGN_DISTRIBUTED
    vertical_top = C.LXW_ALIGN_VERTICAL_TOP
    vertical_bottom = C.LXW_ALIGN_VERTICAL_BOTTOM
    vertical_center = C.LXW_ALIGN_VERTICAL_CENTER
    vertical_justify = C.LXW_ALIGN_VERTICAL_JUSTIFY
    vertical_distributed = C.LXW_ALIGN_VERTICAL_DISTRIBUTED
}

pub enum FormatDiagonalTypes as u32 {
    border_up = C.LXW_DIAGONAL_BORDER_UP
    border_down = C.LXW_DIAGONAL_BORDER_DOWN
    border_up_down = C.LXW_DIAGONAL_BORDER_UP_DOWN
}

pub enum FormatPatterns as u32 {
    @none = C.LXW_PATTERN_NONE
    solid = C.LXW_PATTERN_SOLID
    medium_gray = C.LXW_PATTERN_MEDIUM_GRAY
    dark_gray = C.LXW_PATTERN_DARK_GRAY
    light_gray = C.LXW_PATTERN_LIGHT_GRAY
    dark_horizontal = C.LXW_PATTERN_DARK_HORIZONTAL
    dark_vertical = C.LXW_PATTERN_DARK_VERTICAL
    dark_down = C.LXW_PATTERN_DARK_DOWN
    dark_up = C.LXW_PATTERN_DARK_UP
    dark_grid = C.LXW_PATTERN_DARK_GRID
    dark_trellis = C.LXW_PATTERN_DARK_TRELLIS
    light_horizontal = C.LXW_PATTERN_LIGHT_HORIZONTAL
    light_vertical = C.LXW_PATTERN_LIGHT_VERTICAL
    light_down = C.LXW_PATTERN_LIGHT_DOWN
    ligh_up = C.LXW_PATTERN_LIGHT_UP
    light_grid = C.LXW_PATTERN_LIGHT_GRID
    light_trellis = C.LXW_PATTERN_LIGHT_TRELLIS
    gray_125 = C.LXW_PATTERN_GRAY_125
    gray_0625 = C.LXW_PATTERN_GRAY_0625
}

pub enum FormatBorders as u32 {
    @none = C.LXW_BORDER_NONE
    thin = C.LXW_BORDER_THIN
    medium = C.LXW_BORDER_MEDIUM
    dashed = C.LXW_BORDER_DASHED
    dotted = C.LXW_BORDER_DOTTED
    thick = C.LXW_BORDER_THICK
    double = C.LXW_BORDER_DOUBLE
    hair = C.LXW_BORDER_HAIR
    medium_dashed = C.LXW_BORDER_MEDIUM_DASHED
    dash_dot = C.LXW_BORDER_DASH_DOT
    medium_dash_dot = C.LXW_BORDER_MEDIUM_DASH_DOT
    dash_dot_dot = C.LXW_BORDER_DASH_DOT_DOT
    medium_dash_dot_dot = C.LXW_BORDER_MEDIUM_DASH_DOT_DOT
    slant_dash_dot = C.LXW_BORDER_SLANT_DASH_DOT
}

// lxw_color_t = u32

// DefinedColors
pub enum DefinedColors as u32 {
    black = C.LXW_COLOR_BLACK
    blue = C.LXW_COLOR_BLUE
    brown = C.LXW_COLOR_BROWN
    cyan = C.LXW_COLOR_CYAN
    gray = C.LXW_COLOR_GRAY
    green = C.LXW_COLOR_GREEN
    lime = C.LXW_COLOR_LIME
    magneta = C.LXW_COLOR_MAGENTA
    navy = C.LXW_COLOR_NAVY
    orange = C.LXW_COLOR_ORANGE
    pink = C.LXW_COLOR_PINK
    purple = C.LXW_COLOR_PURPLE
    red = C.LXW_COLOR_RED
    silver = C.LXW_COLOR_SILVER
    white = C.LXW_COLOR_WHITE
    yellow = C.LXW_COLOR_YELLOW
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


