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

// ChartSheet //
struct C.lxw_chartsheet {
}
pub struct ChartSheet {
    ptr &C.lxw_chartsheet
}

fn C.workbook_add_chartsheet(workbook &C.lxw_workbook, sheetname &char) &C.lxw_chartsheet
pub fn (wb Workbook)add_chartsheet(sheet_name string) ChartSheet {
    return ChartSheet {
        ptr: C.workbook_add_chartsheet(wb.ptr, &char(sheet_name.str))
    }
}


// Chart // 
pub enum ChartType as u8 {
    area = C.LXW_CHART_AREA
    area_stacked = C.LXW_CHART_AREA_STACKED
    area_stacked_percent = C.LXW_CHART_AREA_STACKED_PERCENT
    area_bar = C.LXW_CHART_BAR
    area_bar_stacked = C.LXW_CHART_BAR_STACKED
    area_bar_stacked_percent = C.LXW_CHART_BAR_STACKED_PERCENT
    column = C.LXW_CHART_COLUMN
    column_stacked = C.LXW_CHART_COLUMN_STACKED
    column_stacked_percent = C.LXW_CHART_COLUMN_STACKED_PERCENT
    doughnut = C.LXW_CHART_DOUGHNUT
    line = C.LXW_CHART_LINE
    line_stacked = C.LXW_CHART_LINE_STACKED
    line_stacked_percent = C.LXW_CHART_LINE_STACKED_PERCENT
    pie = C.LXW_CHART_PIE
    scatter = C.LXW_CHART_SCATTER
    scatter_straight = C.LXW_CHART_SCATTER_STRAIGHT
    scatter_straight_with_markers = C.LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS
    scatter_smooth = C.LXW_CHART_SCATTER_SMOOTH
    scatter_smooth_with_markers = C.LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS
    radar = C.LXW_CHART_RADAR
    radar_with_markers = C.LXW_CHART_RADAR_WITH_MARKERS
    radar_filled = C.LXW_CHART_RADAR_FILLED
}

struct C.lxw_chart {
}

pub struct Chart {
    ptr C.lxw_chart 
}

fn C.workbook_add_chart(wb &C.lxw_workbook, chart_type u8) &C.lxw_chart 
pub fn (wb Workbook)add_chart(chart_type ChartType) Chart {
    return unsafe { Chart {
        ptr: C.workbook_add_chart(wb.ptr, u8(chart_type))
    }}
}

// DocProperties //
struct C.lxw_doc_properties {
    title &char 
    subject &char 
    author &char 
    manager &char 
    company &char 
    category &char 
    keywords &char 
    comments &char 
    status &char 
    hyperlink_base &char 
    created &C.time_t 
}
type CLxwDocProperties = C.lxw_doc_properties

pub struct DocProperties {
    title string
    subject string 
    author string 
    manager string 
    company string 
    category string 
    keywords string 
    comments string 
    status string 
    hyperlink_base string
}

fn C.workbook_set_properties(wb &C.lxw_workbook, dp &C.lxw_doc_properties) C.lxw_error 
pub fn (wb Workbook)set_properties(doc_properties DocProperties) ReturnCode {
    unsafe {
        dp := CLxwDocProperties {
            title: &char(doc_properties.title.str)
            subject: &char(doc_properties.subject.str)
            author: &char(doc_properties.author.str)
            manager: &char(doc_properties.manager.str)
            company: &char(doc_properties.company.str)
            category: &char(doc_properties.category.str)
            keywords: &char(doc_properties.category.str)
            comments: &char(doc_properties.comments.str)
            status: &char(doc_properties.status.str)
            hyperlink_base: &char(doc_properties.hyperlink_base.str)
        }
        rerror := C.workbook_set_properties(wb.ptr, &dp)
        return ReturnCode(rerror)
    }
}

fn C.workbook_set_custom_property_string(workbook &C.lxw_workbook , name &char, value &char) C.lxw_error 
pub fn (wb Workbook)set_custom_property_string(key string, value string) ReturnCode {
    unsafe {
        return ReturnCode(C.workbook_set_custom_property_string(wb.ptr, &char(key.str), &char(value.str)))
    }
}

fn C.workbook_set_custom_property_number(workbook &C.lxw_workbook , name &char, value f64) C.lxw_error 
pub fn (wb Workbook)set_custom_property_number(key string, value f64) ReturnCode {
    unsafe {
        return ReturnCode(C.workbook_set_custom_property_number(wb.ptr, &char(key.str), value))
    }
}

fn C.workbook_set_custom_property_integer(workbook &C.lxw_workbook, name &char, value i32) C.lxw_error 
pub fn (wb Workbook)set_custom_property_integer(key string, value i32) ReturnCode {
    unsafe {
        return ReturnCode(C.workbook_set_custom_property_integer(wb.ptr, &char(key.str), value))
    }
}

fn C.workbook_set_custom_property_boolean(workbook &C.lxw_workbook, name &char, value u8) C.lxw_error 
pub fn (wb Workbook)set_custom_property_boolean(key string, value bool) ReturnCode {
    value_u8 := if value {1} else {0}
    unsafe {
        return ReturnCode(C.workbook_set_custom_property_boolean(wb.ptr, &char(key.str), value_u8))
    }
}

fn C.workbook_define_name(workbook &C.lxw_workbook, name &char, formula &char) C.lxw_error 
pub fn (wb Workbook)define_name(name string, formula string) ReturnCode {
    unsafe {
        return ReturnCode(C.workbook_define_name(wb.ptr, &char(name.str), &char(formula.str)))
    }
}

fn C.lxw_workbook_free(wb &C.lxw_workbook)
pub fn (wb Workbook)deallocate() {
    C.lxw_workbook_free(wb.ptr)
}

// Format //
struct C.lxw_format {
}

pub struct Format {
mut:
    ptr &C.lxw_format 
}

fn C.workbook_add_format(wb &C.lxw_workbook) &C.lxw_format
pub fn (wb Workbook)add_format() Format {
    return Format {
        ptr: C.workbook_add_format(wb.ptr)
    }
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

fn C.format_set_font_color(format &C.lxw_format, color C.lxw_color_t)
pub fn (f Format)set_font_color(color DefinedColors) {
    unsafe {
        C.format_set_font_color(f.ptr, C.lxw_color_t(color))
    }
}

fn C.format_set_bold(format &C.lxw_format)
pub fn (f Format)set_bold() {
    C.format_set_bold(f.ptr)
}

fn C.format_set_italic(format &C.lxw_format)
pub fn (f Format)set_italic() {
    C.format_set_italic(f.ptr)
}

fn C.format_set_underline(format &C.lxw_format, style u8)
pub fn (f Format)set_underline(style FormatUnderlines) {
    C.format_set_underline(f.ptr, u8(style))
}

fn C.format_set_font_strikeout(format &C.lxw_format)
pub fn (f Format)set_font_strikeout() {
    C.format_set_font_strikeout(f.ptr)
}

fn C.format_set_font_script(format &C.lxw_format, style u8)
pub fn (f Format)set_font_script(style FormatScrips) {
    C.format_set_font_script(f.ptr, u8(style))
}

fn C.format_set_num_format(format &C.lxw_format, format &char)
pub fn (f Format)set_number_format(number_format string) {
    C.format_set_num_format(f.ptr, &char(number_format.str))
}

/* 
 *   | Index | Index | Format String                                        |
 *   | ----- | ----- | ---------------------------------------------------- |
 *   | 0     | 0x00  | `General`                                            |
 *   | 1     | 0x01  | `0`                                                  |
 *   | 2     | 0x02  | `0.00`                                               |
 *   | 3     | 0x03  | `#,##0`                                              |
 *   | 4     | 0x04  | `#,##0.00`                                           |
 *   | 5     | 0x05  | `($#,##0_)($#,##0)`                                  |
 *   | 6     | 0x06  | `($#,##0_)[Red]($#,##0)`                             |
 *   | 7     | 0x07  | `($#,##0.00_)($#,##0.00)`                            |
 *   | 8     | 0x08  | `($#,##0.00_)[Red]($#,##0.00)`                       |
 *   | 9     | 0x09  | `0%`                                                 |
 *   | 10    | 0x0a  | `0.00%`                                              |
 *   | 11    | 0x0b  | `0.00E+00`                                           |
 *   | 12    | 0x0c  | `# ?/?`                                              |
 *   | 13    | 0x0d  | `# ??/??`                                            |
 *   | 14    | 0x0e  | `m/d/yy`                                             |
 *   | 15    | 0x0f  | `d-mmm-yy`                                           |
 *   | 16    | 0x10  | `d-mmm`                                              |
 *   | 17    | 0x11  | `mmm-yy`                                             |
 *   | 18    | 0x12  | `h:mm AM/PM`                                         |
 *   | 19    | 0x13  | `h:mm:ss AM/PM`                                      |
 *   | 20    | 0x14  | `h:mm`                                               |
 *   | 21    | 0x15  | `h:mm:ss`                                            |
 *   | 22    | 0x16  | `m/d/yy h:mm`                                        |
 *   | ...   | ...   | ...                                                  |
 *   | 37    | 0x25  | `(#,##0_)(#,##0)`                                    |
 *   | 38    | 0x26  | `(#,##0_)[Red](#,##0)`                               |
 *   | 39    | 0x27  | `(#,##0.00_)(#,##0.00)`                              |
 *   | 40    | 0x28  | `(#,##0.00_)[Red](#,##0.00)`                         |
 *   | 41    | 0x29  | `_(* #,##0_)_(* (#,##0)_(* "-"_)_(@_)`               |
 *   | 42    | 0x2a  | `_($* #,##0_)_($* (#,##0)_($* "-"_)_(@_)`            |
 *   | 43    | 0x2b  | `_(* #,##0.00_)_(* (#,##0.00)_(* "-"??_)_(@_)`       |
 *   | 44    | 0x2c  | `_($* #,##0.00_)_($* (#,##0.00)_($* "-"??_)_(@_)`    |
 *   | 45    | 0x2d  | `mm:ss`                                              |
 *   | 46    | 0x2e  | `[h]:mm:ss`                                          |
 *   | 47    | 0x2f  | `mm:ss.0`                                            |
 *   | 48    | 0x30  | `##0.0E+0`                                           |
 *   | 49    | 0x31  | `@`                                                  |
 *   */

fn C.format_set_num_format_index(format &C.lxw_format, index u8)
pub fn (f Format)set_number_format_index(index u8) {
    C.format_set_num_format_index(f.ptr, index)
}

fn C.format_set_unlocked(format &C.lxw_format)
pub fn (f Format)set_unlocked() {
    C.format_set_unlocked(f.ptr)
}

fn C.format_set_hidden(format &C.lxw_format)
pub fn (f Format)set_hidden() {
    C.format_set_hidden(f.ptr)
}

fn C.format_set_align(format &C.lxw_format, alignment u8)
pub fn (f Format)set_align(alignment FormatAlignments) {
    C.format_set_align(f.ptr, u8(alignment))
}

fn C.format_set_text_wrap(format &C.lxw_format)
pub fn (f Format)set_text_wrap() {
    C.format_set_text_wrap(f.ptr)
}

fn C.format_set_rotation(format &C.lxw_format, angle i16)
pub fn (f Format)set_rotation(angle i16) {
    if (angle < -360) || (angle > 360) {
        return 
    }
    C.format_set_rotation(f.ptr, angle)
}

fn C.format_set_indent(format &C.lxw_format, level u8)
pub fn (f Format)set_indent(level u8) {
    C.format_set_indent(f.ptr, level)
}

fn C.format_set_shrink(format &C.lxw_format)
pub fn (f Format)set_shrink() {
    C.format_set_shrink(f.ptr)
}

fn C.format_set_pattern(format &C.lxw_format, index u8)
pub fn (f Format)set_pattern(pattern FormatPatterns) {
    C.format_set_pattern(f.ptr, u8(pattern))
}

fn C.format_set_bg_color(format &C.lxw_format, color C.lxw_color_t)
pub fn (f Format)set_background(color DefinedColors) {
    unsafe {
        C.format_set_bg_color(f.ptr, C.lxw_color_t(color))
    }
}

fn C.format_set_fg_color(format &C.lxw_format, color C.lxw_color_t)
pub fn (f Format)set_foreground(color DefinedColors) {
    unsafe {
        C.format_set_fg_color(f.ptr, C.lxw_color_t(color))
    }
}

pub enum BorderLine {
    all
    top
    bottom
    left
    right
}

fn C.format_set_border(format &C.lxw_format, style u8)
fn C.format_set_top(fmormat &C.lxw_format, style u8)
fn C.format_set_bottom(format &C.lxw_format, style u8)
fn C.format_set_left(format &C.lxw_format, style u8)
fn C.format_set_right(format &C.lxw_format, style u8)
fn C.format_set_border_color(format &C.lxw_format, color C.lxw_color_t)
fn C.format_set_top_color(format &C.lxw_format, color C.lxw_color_t)
fn C.format_set_bottom_color(format &C.lxw_format, color C.lxw_color_t)
fn C.format_set_left_color(format &C.lxw_format, color C.lxw_color_t)
fn C.format_set_right_color(format &C.lxw_format, color C.lxw_color_t)

pub fn (f Format)set_border(border BorderLine, style FormatBorders, color DefinedColors) {
    match border {
        .all {
            unsafe {
                C.format_set_border(f.ptr, u8(style))
                C.format_set_border_color(f.ptr, C.lxw_color_t(color))
            }
        }
        .top {
            unsafe {
                C.format_set_top(f.ptr, u8(style))
                C.format_set_top_color(f.ptr, C.lxw_color_t(color))
            }
        }
        .bottom {
            unsafe {
                C.format_set_bottom(f.ptr, u8(style))
                C.format_set_bottom_color(f.ptr, C.lxw_color_t(color))
            }
        }
        .left {
            unsafe {
                C.format_set_left(f.ptr, u8(style))
                C.format_set_left_color(f.ptr, C.lxw_color_t(color))
            }
        }
        .right {
            unsafe {
                C.format_set_right(f.ptr, u8(style))
                C.format_set_right_color(f.ptr, C.lxw_color_t(color))
            }
        }
    }
}

fn C.format_set_diag_type(format &C.lxw_format, dt u8)
pub fn (f Format)set_diagonal_type(diag_type FormatDiagonalTypes) {
    unsafe {
        C.format_set_diag_type(f.ptr, u8(diag_type))
    }
}

fn C.format_set_diag_border(format &C.lxw_format, style u8)
fn C.format_set_diag_color(format &C.lxw_format, color &C.lxw_color_t)
pub fn (f Format)set_diagonal_style(border FormatBorders, color DefinedColors) {
    unsafe {
        C.format_set_diag_border(f.ptr, u8(border))
        C.format_set_diag_color(f.ptr, C.lxw_color_t(color))
    }
}

fn C.format_set_quote_prefix(format &C.lxw_format)
pub fn (f Format)set_quote_prefix() {
    C.format_set_quote_prefix(f.ptr)
}

fn C.format_set_font_outline(format &C.lxw_format)
pub fn (f Format)set_font_outline() {
    C.format_set_font_outline(f.ptr)
}

fn C.format_set_font_shadow(format &C.lxw_format)
pub fn (f Format)set_font_shadow() {
    C.format_set_font_shadow(f.ptr)
}

fn C.format_set_font_family(format &C.lxw_format, value u8)
pub fn (f Format)set_font_family(value u8) {
    C.format_set_font_family(f.ptr, value)
}

fn C.format_set_font_charset(format &C.lxw_format, value u8)
pub fn (f Format)set_font_charset(value u8) {
    C.format_set_font_charset(f.ptr, value)
}

fn C.format_set_font_scheme(format &C.lxw_format, font_scheme &char)
pub fn (f Format)set_font_scheme(font_scheme string) {
    C.format_set_font_scheme(f.ptr, &char(font_scheme.str))
}

fn C.format_set_font_condense(format &C.lxw_format)
pub fn (f Format)set_font_condense() {
    C.format_set_font_condense(f.ptr)
}

fn C.format_set_font_extend(format &C.lxw_format)
pub fn (f Format)set_font_extend() {
    C.format_set_font_extend(f.ptr)
}

fn C.format_set_reading_order(format &C.lxw_format, value u8)
pub fn (f Format)set_reading_order(value u8) {
    C.format_set_reading_order(f.ptr, value)
}

fn C.format_set_theme(format &C.lxw_format, value u8)
pub fn (f Format)set_theme(theme_index u8) {
    C.format_set_theme(f.ptr, theme_index)
}

fn C.format_set_hyperlink(format &C.lxw_format)
pub fn (f Format)set_hyperlink() {
    C.format_set_hyperlink(f.ptr)
}

fn C.format_set_color_indexed(format &C.lxw_format, value u8)
pub fn (f Format)set_color_indexed(index u8) {
    C.format_set_color_indexed(f.ptr, index)
}

fn C.format_set_font_only(format &C.lxw_format)
pub fn (f Format)set_font_only() {
    C.format_set_font_only(f.ptr)
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


