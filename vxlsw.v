// Copyright (c) 2023 Erdet Nasufi

module vxlsw

#flag -I /usr/local/include/ 
#flag -L /usr/local/lib/
#flag -l xlsxwriter

#include <xlsxwriter.h>

// Workbook 
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

// Worksheet 
struct C.lxw_worksheet {
}

pub struct Worksheet {
    ptr &C.lxw_worksheet
}

fn C.worksheet_write_string(sheet &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, text &char, format &C.lxw_format) C.lxw_error
pub fn (ws Worksheet)write_string(row u32, col u16, text string, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_string(ws.ptr, C.lxw_row_t(row), C.lxw_col_t(col), &char(text.str), format.ptr))
    }
}

fn C.worksheet_write_number(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, number f64, format &C.lxw_format) C.lxw_error
pub fn (ws Worksheet)write_number(row u32, col u16, number f64, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_number(ws.ptr, C.lxw_row_t(row), C.lxw_col_t(col), number, format.ptr))
    }
}

fn C.worksheet_write_formula(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, formula &char, format &C.lxw_format) C.lxw_error 
pub fn (ws Worksheet)write_formula(row u32, col u16, formula string, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_formula(ws.ptr, C.lxw_row_t(row), C.lxw_col_t(col), &char(formula.str), format.ptr))
    }
}

fn C.worksheet_write_array_formula(ws &C.lxw_worksheet, first_row C.lxw_row_t, first_col C.lxw_col_t, last_row C.lxw_row_t, last_col C.lxw_col_t, formula &char, format &C.lxw_format) C.lxw_error 
pub fn (ws Worksheet)write_array_formula(first_row u32, first_col u16, last_row u32, last_col u16, formula string, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_array_formula(ws.ptr, 
                    C.lxw_row_t(first_row), C.lxw_col_t(first_col),
                    C.lxw_row_t(last_row), C.lxw_col_t(last_col),
                    &char(formula.str),
                    format.ptr))
        }
}

fn C.worksheet_write_dynamic_array_formula(ws &C.lxw_worksheet, first_row C.lxw_row_t, first_col C.lxw_col_t, last_row C.lxw_row_t, last_col C.lxw_col_t, formula &char, format &C.lxw_format) C.lxw_error
pub fn (ws Worksheet)write_dynamic_array_formula(first_row u32, first_col u16, last_row u32, last_col u16, formula string, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_dynamic_array_formula(ws.ptr, 
            C.lxw_row_t(first_row), C.lxw_col_t(first_col), 
            C.lxw_row_t(last_row), C.lxw_col_t(last_col),
            &char(formula.str),
            format.ptr))
    }
}

fn C.worksheet_write_dynamic_formula(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, formula &char, format &C.lxw_format) C.lxw_error
pub fn (ws Worksheet)write_dynamic_formula(row u32, col u16, formula string, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_dynamic_formula(ws.ptr, C.lxw_row_t(row), C.lxw_col_t(col), &char(formula.str), format.ptr))
    }
}

fn C.worksheet_write_array_formula_num(ws &C.lxw_worksheet, first_row C.lxw_row_t, first_col C.lxw_col_t, last_row C.lxw_row_t, last_col C.lxw_col_t, formula &char, format &C.lxw_format, result f64) C.lxw_error
pub fn (ws Worksheet)write_array_formula_num(first_row u32, first_col u16, last_row u32, last_col u32, formula string, result f64, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_array_formula_num(ws.ptr,
                    C.lxw_row_t(first_row), C.lxw_col_t(first_col),
                    C.lxw_row_t(last_row), C.lxw_col_t(last_col),
                    &char(formula.str), format.ptr, result))        
    }
}

fn C.worksheet_write_dynamic_array_formula_num(ws &C.lxw_worksheet, first_row C.lxw_row_t, first_col C.lxw_col_t, last_row C.lxw_row_t, last_col C.lxw_col_t, formula &char, format &C.lxw_format, result f64) C.lxw_error
pub fn (ws Worksheet)write_dynamic_array_formula_num(first_row u32, first_col u16, last_row u32, last_col u16, formula string, format Format, result f64) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_dynamic_array_formula_num(ws.ptr,
                    C.lxw_row_t(first_row), C.lxw_col_t(first_col),
                    C.lxw_row_t(last_row), C.lxw_col_t(last_col),
                    &char(formula.str), format.ptr, result))
    }
}

fn C.worksheet_write_dynamic_formula_num(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, formula &char, format &C.lxw_format, result f64) C.lxw_error
pub fn (ws Worksheet)write_dynamic_formula_num(row u32, col u16, formula string, format Format, result f64) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_dynamic_formula_num(ws.ptr, 
                    C.lxw_row_t(row), C.lxw_col_t(col),
                    &char(formula.str), format.ptr, result))
    }
}

// TODO 
fn C.worksheet_write_datetime(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, datetime &C.lxw_datetime, format &C.lxw_format) C.lxw_error
fn C.worksheet_write_unixtime(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, unixtime i64, format &C.lxw_format) C.lxw_error

fn C.worksheet_write_url_opt(ws &C.lxw_worksheet, row_num C.lxw_row_t, col_num C.lxw_col_t, url &char, format &C.lxw_format, text &char, tooltip &char) C.lxw_error
pub fn (ws Worksheet)write_url(row u32, col u16, url string, text string, tooltip string, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_url_opt(ws.ptr, 
                    C.lxw_row_t(row), C.lxw_col_t(col),
                    &char(url.str), 
                    format.ptr,
                    &char(text.str), 
                    &char(tooltip.str)))
    }
}

fn C.worksheet_write_boolean(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, value int, format &C.lxw_format) C.lxw_error
pub fn (ws Worksheet)write_boolean(row u32, col u16, state bool, format Format) ReturnCode {
    state_int := if state {1} else {0}
    return unsafe {
        ReturnCode(C.worksheet_write_boolean(ws.ptr, C.lxw_row_t(row), C.lxw_col_t(col), state_int, format.ptr))
    }
}

fn C.worksheet_write_blank(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, format &C.lxw_format) C.lxw_error 
pub fn (ws Worksheet)write_blank(row u32, col u16, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_blank(ws.ptr, C.lxw_row_t(row), C.lxw_col_t(col), format.ptr))
    }
}

fn C.worksheet_set_column(ws &C.lxw_worksheet, first_col C.lxw_col_t, last_col C.lxw_col_t, width f64, format &C.lxw_format) C.lxw_error 
pub fn (ws Worksheet)set_column(first_col u16, last_col u16, width f64, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_set_column(ws.ptr, C.lxw_col_t(first_col), C.lxw_col_t(last_col), width, format.ptr))
    }
}

fn C.worksheet_set_row(ws &C.lxw_worksheet, row C.lxw_row_t, height f64, format &C.lxw_format) C.lxw_error 
pub fn (ws Worksheet)set_row(row u32, height f64, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_set_row(ws.ptr, C.lxw_row_t(row), height, format.ptr))
    }
}

fn C.worksheet_write_formula_num(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, formula &char, format &C.lxw_format, result f64) C.lxw_error 
pub fn (ws Worksheet)write_formula_num(row u32, col u16, formula string, result f64, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_formula_num(ws.ptr, C.lxw_row_t(row), C.lxw_col_t(col), &char(formula.str), format.ptr, result))
    }
}

fn C.worksheet_write_formula_str(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, formula &char, format &C.lxw_format, result &char) C.lxw_error 
pub fn (ws Worksheet)write_formula_string(row u32, col u16, formula string, format Format, result string) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_write_formula_str(ws.ptr, C.lxw_row_t(row), C.lxw_col_t(col), &char(formula.str), format.ptr, &char(result.str)))
    }
}

// TODO 
fn C.worksheet_write_rich_string(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, rich_string &&char, format &C.lxw_format) C.lxw_error

// TODO 
fn C.worksheet_set_column_pixels_opt(ws &C.lxw_worksheet, first_col C.lxw_col_t, last_col C.lxw_col_t, pixels u32, formats &C.lxw_format, options &C.lxw_row_col_options) C.lxw_error

fn C.worksheet_insert_image(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, filename &char) C.lxw_error 
pub fn (ws Worksheet)insert_image(row u32, col u16, filename string) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_insert_image(ws.ptr, C.lxw_row_t(row), C.lxw_col_t(col), &char(filename.str)))
    }
}

// TODO 
fn C.worksheet_insert_image_opt(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, filename &char, options &C.lxw_image_options) C.lxw_error 

// TODO 
fn C.worksheet_insert_image_buffer(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, image_buffer &u8, image_size C.size_t) C.lxw_error 

fn C.worksheet_insert_image_buffer_opt(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, image_buffer &u8, image_size C.size_t, options &C.lxw_image_options) C.lxw_error 

fn C.worksheet_set_background(ws &C.lxw_worksheet, filename &char) C.lxw_error 
pub fn (ws Worksheet)set_background(filename string) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_set_background(ws.ptr, &char(filename.str)))
    }
}

// TODO 
fn C.worksheet_set_background_buffer(ws &C.lxw_worksheet, buffer_image &u8, image_size C.size_t) C.lxw_error 

fn C.worksheet_insert_chart(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, chart &C.lxw_chart) C.lxw_error 
pub fn (ws Worksheet)insert_chart(row u32, col u16, chart Chart) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_insert_chart(ws.ptr, C.lxw_row_t(row), C.lxw_col_t(col), chart.ptr))
    }
}

// TODO 
fn C.worksheet_insert_chart_opt(ws &C.lxw_worksheet, row C.lxw_row_t, col C.lxw_col_t, chart &C.lxw_chart, user_options &C.lxw_chart_options) C.lxw_error 

fn C.worksheet_merge_range(ws &C.lxw_worksheet, first_row C.lxw_row_t, first_col C.lxw_col_t, last_row C.lxw_row_t, last_col C.lxw_col_t, text &char, fomrat &C.lxw_format) C.lxw_error 
pub fn (ws Worksheet)merge_range(first_row u32, first_col u16, last_row u32, last_col u16, text string, format Format) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_merge_range(ws.ptr, 
                                C.lxw_row_t(first_row), C.lxw_col_t(first_col),
                                C.lxw_row_t(last_row), C.lxw_col_t(last_col),
                                &char(text.str), format.ptr))
    }
}

// TODO 
fn C.worksheet_autofilter(ws &C.lxw_worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col) C.lxw_error 

fn C.worksheet_filter_column(ws &C.lxw_worksheet, col C.lxw_col_t, rule &C.lxw_filter_rule) C.lxw_error 

fn C.worksheet_set_selection(ws &C.lxw_worksheet, first_row C.lxw_row_t, first_col C.lxw_col_t, last_row C.lxw_row_t, last_col C.lxw_col_t)
pub fn (ws Worksheet)set_selection(first_row u32, first_col u16, last_row u32, last_col u16) {
    unsafe {
        C.worksheet_set_selection(ws.ptr,
                                  C.lxw_row_t(first_row), C.lxw_col_t(first_col),
                                  C.lxw_row_t(last_row), C.lxw_col_t(last_col))
    }
}

fn C.worksheet_set_top_left_cell(ws &C.lxw_worksheet, row C.lxw_row_t, col &C.lxw_col_t)
pub fn (ws Worksheet)set_top_left_cell(row u32, col u16) {
    unsafe {
        C.worksheet_set_top_left_cell(ws.ptr, C.lxw_row_t(row), C.lxw_col_t(col))
    }
}

fn C.worksheet_set_landscape(ws &C.lxw_worksheet)
pub fn (ws Worksheet)set_landscape() {
    C.worksheet_set_landscape(ws.ptr)
}

fn C.worksheet_set_portrait(ws &C.lxw_worksheet)
pub fn (ws Worksheet)set_portrait() {
    C.worksheet_set_portrait(ws.ptr)
}

fn C.worksheet_set_page_view(ws &C.lxw_worksheet)
pub fn (ws Worksheet)set_page_view() {
    C.worksheet_set_page_view(ws.ptr)
}

// TODO: paper_type as enum 
fn C.worksheet_set_paper(ws &C.lxw_worksheet, paper_type u8)
pub fn (ws Worksheet)set_paper(paper_type u8) {
    C.worksheet_set_paper(ws.ptr, paper_type)
}

fn C.worksheet_set_margins(ws &C.lxw_worksheet, left f64, right f64, top f64, bottom f64)
pub fn (ws Worksheet)set_margines(left f64, right f64, top f64, bottom f64) {
    C.worksheet_set_margins(ws.ptr, left, right, top, bottom)
}

fn C.worksheet_set_header(ws &C.lxw_worksheet, text &char) C.lxw_error
pub fn (ws Worksheet)set_header(text string) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_set_header(ws.ptr, &char(text.str)))
    }
}

fn C.worksheet_set_footer(ws &C.lxw_worksheet, text &char) C.lxw_error 
pub fn (ws Worksheet)set_footer(text string) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_set_footer(ws.ptr, &char(text.str)))
    }
}

// TODO 
fn C.worksheet_set_header_opt(ws &C.lxw_worksheet, text &char, options &C.lxw_header_footer_options) C.lxw_error 
fn C.worksheet_set_footer_opt(ws &C.lxw_worksheet, text &char, options &C.lxw_header_footer_options) C.lxw_error

fn C.worksheet_print_across(ws &C.lxw_worksheet)
pub fn (ws Worksheet)print_across() {
    C.worksheet_print_across(ws.ptr)
}

fn C.worksheet_set_zoom(ws &C.lxw_worksheet, scale u16)
pub fn (ws Worksheet)set_zoom(scale u16) {
    C.worksheet_set_zoom(ws.ptr, scale)
}

fn C.worksheet_gridlines(ws &C.lxw_worksheet, option u8)
pub fn (ws Worksheet)gridlines(option u8) {
    C.worksheet_gridlines(ws.ptr, option)
}

fn C.worksheet_center_horizontally(ws &C.lxw_worksheet)
pub fn (ws Worksheet)center_horizontally() {
    C.worksheet_center_horizontally(ws.ptr)
}

fn C.worksheet_center_vertically(ws &C.lxw_worksheet)
pub fn (ws Worksheet)center_vertically() {
    C.worksheet_center_vertically(ws.ptr)
}

fn C.worksheet_print_row_col_headers(ws &C.lxw_worksheet)
pub fn (ws Worksheet)print_row_col_headers() {
    C.worksheet_print_row_col_headers(ws.ptr)
}

fn C.worksheet_repeat_rows(ws &C.lxw_worksheet, first_row C.lxw_row_t, last_row C.lxw_row_t) C.lxw_error
pub fn (ws Worksheet)repeat_rows(first_row u32, last_row u32) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_repeat_rows(ws.ptr, C.lxw_row_t(first_row), C.lxw_row_t(last_row)))
    }
}

fn C.worksheet_repeat_columns(ws &C.lxw_worksheet, first_col C.lxw_col_t, last_col C.lxw_col_t) C.lxw_error
pub fn (ws Worksheet)repeat_column(first_col u16, last_col u16) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_repeat_columns(ws.ptr, C.lxw_col_t(first_col), C.lxw_col_t(last_col)))
    }
}

fn C.worksheet_print_area(ws &C.lxw_worksheet, first_row C.lxw_row_t, first_col C.lxw_col_t, last_row C.lxw_row_t, last_col C.lxw_col_t) C.lxw_error
pub fn (ws Worksheet)print_area(first_row u32, first_col u16, last_row u32, last_col u16) ReturnCode {
    return unsafe {
        ReturnCode(C.worksheet_print_area(ws.ptr, 
            C.lxw_row_t(first_row), C.lxw_col_t(first_col),
            C.lxw_row_t(last_row), C.lxw_col_t(last_col)))
    }
}

fn C.worksheet_fit_to_pages(ws &C.lxw_worksheet, width u16, height u16)
pub fn (ws Worksheet)fit_to_pages(width u16, height u16) {
    C.worksheet_fit_to_pages(ws.ptr, width, height)
}

fn C.worksheet_set_start_page(ws &C.lxw_worksheet, start_page u16)
pub fn (ws Worksheet)set_start_page(start_page u16) {
    C.worksheet_set_start_page(ws.ptr, start_page)
}

fn C.worksheet_set_print_scale(ws &C.lxw_worksheet, scale u16)
pub fn (ws Worksheet)set_print_scale(scale u16) {
    C.worksheet_set_print_scale(ws.ptr, scale)
}

fn C.worksheet_print_black_and_white(ws &C.lxw_worksheet)
pub fn (ws Worksheet)print_black_and_white() {
    C.worksheet_print_black_and_white(ws.ptr)
}

fn C.worksheet_right_to_left(ws &C.lxw_worksheet)
pub fn (ws Worksheet)right_to_left() {
    C.worksheet_right_to_left(ws.ptr)
}

fn C.worksheet_hide_zero(ws &C.lxw_worksheet)
pub fn (ws Worksheet)hide_zero() {
    C.worksheet_hide_zero(ws.ptr)
}

fn C.worksheet_set_tab_color(ws &C.lxw_worksheet, color C.lxw_color_t)
pub fn (ws Worksheet)set_tab_color(color DefinedColors) {
    C.worksheet_set_tab_color(ws.ptr, C.lxw_color_t(color))
}

// TODO
fn C.worksheet_protect(ws &C.lxw_worksheet, password &char, options &C.lxw_protection)

fn C.worksheet_outline_settings(ws &C.lxw_worksheet, visible u8, symbols_below u8, symbols_right u8, auto_style u8)
pub fn (ws Worksheet)outline_settings(visible u8, symbols_below u8, symbols_right u8, auto_style u8) {
    C.worksheet_outline_settings(ws.ptr, visible, symbols_below, symbols_right, auto_style)
}

fn C.worksheet_set_default_row(ws &C.lxw_worksheet, height f64, hide_unused_rows u8)
pub fn (ws Worksheet)set_default_row(height f64, hide_unused_rows u8) {
    C.worksheet_set_default_row(ws.ptr, height, hide_unused_rows)
}

fn C.worksheet_show_comments(ws &C.lxw_worksheet)
pub fn (ws Worksheet)show_comments() {
    C.worksheet_show_comments(ws.ptr)
}

fn C.worksheet_set_comments_author(ws &C.lxw_worksheet, author &char)
pub fn (ws Worksheet)set_comments_author(author string) {
    C.worksheet_set_comments_author(ws.ptr, &char(author.str))
}

fn C.lxw_worksheet_write_sheet_views(ws &C.lxw_worksheet)
pub fn (ws Worksheet)write_sheet_views() {
    C.lxw_worksheet_write_sheet_views(ws.ptr)
}

fn C.lxw_worksheet_write_page_margins(ws &C.lxw_worksheet)
pub fn (ws Worksheet)write_page_margins() {
    C.lxw_worksheet_write_page_margins(ws.ptr)
}

fn C.lxw_worksheet_write_drawings(ws &C.lxw_worksheet)
pub fn (ws Worksheet)write_drawings() {
    C.lxw_worksheet_write_drawings(ws.ptr)
}

// TODO
fn C.lxw_worksheet_write_sheet_protection(ws &C.lxw_worksheet, protect &C.lxw_protection_obj)

fn C.lxw_worksheet_write_sheet_pr(ws &C.lxw_worksheet)
pub fn (ws Worksheet)write_sheet_pr() {
    C.lxw_worksheet_write_sheet_pr(ws.ptr)
}

fn C.lxw_worksheet_write_page_setup(ws &C.lxw_worksheet)
pub fn (ws Worksheet)write_page_setup() {
    C.lxw_worksheet_write_page_setup(ws.ptr)
}

fn C.lxw_worksheet_write_header_footer(ws &C.lxw_worksheet)
pub fn (ws Worksheet)write_header_footer() {
    C.lxw_worksheet_write_header_footer(ws.ptr)
}

pub enum Gridlines { // lxw_gridlines
    hide_all_gridlines = C.LXW_HIDE_ALL_GRIDLINES
    show_screen_gridlines = C.LXW_SHOW_SCREEN_GRIDLINES
    show_print_gridlines = C.LXW_SHOW_PRINT_GRIDLINES
    show_all_gridlines = C.LXW_SHOW_ALL_GRIDLINES
}

pub enum ValidationBoolean { // lxw_validation_boolean
    default = C.LXW_VALIDATION_DEFAULT
    off = C.LXW_VALIDATION_OFF
    on = C.LXW_VALIDATION_ON
}

pub enum ValidationTypes { // lxw_validation_types
    @none = C.LXW_VALIDATION_TYPE_NONE
    integer = C.LXW_VALIDATION_TYPE_INTEGER
    integer_formula = C.LXW_VALIDATION_TYPE_INTEGER_FORMULA
    decimal = C.LXW_VALIDATION_TYPE_DECIMAL
    decimal_formula = C.LXW_VALIDATION_TYPE_DECIMAL_FORMULA
    list = C.LXW_VALIDATION_TYPE_LIST
    formula = C.LXW_VALIDATION_TYPE_LIST_FORMULA
    date = C.LXW_VALIDATION_TYPE_DATE
    date_formula = C.LXW_VALIDATION_TYPE_DATE_FORMULA
    date_number = C.LXW_VALIDATION_TYPE_DATE_NUMBER
    time = C.LXW_VALIDATION_TYPE_TIME
    time_formula = C.LXW_VALIDATION_TYPE_TIME_FORMULA
    time_number = C.LXW_VALIDATION_TYPE_TIME_NUMBER
    type_length = C.LXW_VALIDATION_TYPE_LENGTH
    length_formula = C.LXW_VALIDATION_TYPE_LENGTH_FORMULA
    type_custom_formula = C.LXW_VALIDATION_TYPE_CUSTOM_FORMULA
    type_any = C.LXW_VALIDATION_TYPE_ANY
}

pub enum ValidationCriteria { // lxw_validation_criteria 
    @none = C.LXW_VALIDATION_CRITERIA_NONE
    between = C.LXW_VALIDATION_CRITERIA_BETWEEN
    not_between = C.LXW_VALIDATION_CRITERIA_NOT_BETWEEN
    equal_to = C.LXW_VALIDATION_CRITERIA_EQUAL_TO
    not_equal_to = C.LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO
    greater_than = C.LXW_VALIDATION_CRITERIA_GREATER_THAN
    less_than = C.LXW_VALIDATION_CRITERIA_LESS_THAN
    greater_than_or_equal_to = C.LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO
    less_than_or_equal_to = C.LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO
}

pub enum ValidationErrorTypes { // lxw_validation_error_types 
    stop = C.LXW_VALIDATION_ERROR_TYPE_STOP
    warning = C.LXW_VALIDATION_ERROR_TYPE_WARNING
    information = C.LXW_VALIDATION_ERROR_TYPE_INFORMATION
}

pub enum CommentDisplayTypes { // lxw_comment_display_types 
    default = C.LXW_COMMENT_DISPLAY_DEFAULT
    hidden = C.LXW_COMMENT_DISPLAY_HIDDEN
    visible = C.LXW_COMMENT_DISPLAY_VISIBLE
}

pub enum ConditionalFormatTypes { // lxw_conditional_format_types 
    @none = C.LXW_CONDITIONAL_TYPE_NONE
    cell = C.LXW_CONDITIONAL_TYPE_CELL
    text = C.LXW_CONDITIONAL_TYPE_TEXT
    time_period = C.LXW_CONDITIONAL_TYPE_TIME_PERIOD
    average = C.LXW_CONDITIONAL_TYPE_AVERAGE
    duplicate = C.LXW_CONDITIONAL_TYPE_DUPLICATE
    unique = C.LXW_CONDITIONAL_TYPE_UNIQUE
    top = C.LXW_CONDITIONAL_TYPE_TOP
    bottom = C.LXW_CONDITIONAL_TYPE_BOTTOM
    blanks = C.LXW_CONDITIONAL_TYPE_BLANKS
    no_blanks = C.LXW_CONDITIONAL_TYPE_NO_BLANKS
    errors = C.LXW_CONDITIONAL_TYPE_ERRORS
    no_errors = C.LXW_CONDITIONAL_TYPE_NO_ERRORS
    type_formula = C.LXW_CONDITIONAL_TYPE_FORMULA
    color_scale_2 = C.LXW_CONDITIONAL_2_COLOR_SCALE
    color_scale_3 = C.LXW_CONDITIONAL_3_COLOR_SCALE
    data_bar = C.LXW_CONDITIONAL_DATA_BAR
    icon_sets = C.LXW_CONDITIONAL_TYPE_ICON_SETS
    type_last = C.LXW_CONDITIONAL_TYPE_LAST
}

pub enum ConditionalCriteria { // lxw_conditional_criteria 
    @none = C.LXW_CONDITIONAL_CRITERIA_NONE
    equal_to = C.LXW_CONDITIONAL_CRITERIA_EQUAL_TO
    not_equal_to = C.LXW_CONDITIONAL_CRITERIA_NOT_EQUAL_TO
    greater_than = C.LXW_CONDITIONAL_CRITERIA_GREATER_THAN
    less_than = C.LXW_CONDITIONAL_CRITERIA_LESS_THAN
    greater_than_or_equal_to = C.LXW_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO
    less_than_or_equal_to = C.LXW_CONDITIONAL_CRITERIA_LESS_THAN_OR_EQUAL_TO
    between = C.LXW_CONDITIONAL_CRITERIA_BETWEEN
    not_between = C.LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN
    text_containing = C.LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING
    text_not_containig = C.LXW_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING
    text_begins_with = C.LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH
    text_ends_with = C.LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH
    time_period_yesterday = C.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY
    time_period_today = C.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY
    time_period_tomorrow = C.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW
    time_period_last_7_days = C.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS
    time_period_last_week = C.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK
    time_period_this_week = C.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK
    time_period_next_week = C.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK
    time_period_last_month = C.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH
    time_period_this_month = C.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH
    time_period_next_month = C.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH
    average_above = C.LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE
    average_below = C.LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW
    average_above_or_equal = C.LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL
    average_below_or_equal = C.LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL
    average_1std_dev_above = C.LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_ABOVE
    average_1std_dev_below = C.LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW
    average_2std_dev_above = C.LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_ABOVE
    average_2std_dev_below = C.LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW
    average_3std_dev_above = C.LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_ABOVE
    average_3std_dev_below = C.LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW
    top_or_bottom_percent = C.LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT
}

pub enum ConditionalFormat_rule_types { // lxw_conditional_format_rule_types
    @none = C.LXW_CONDITIONAL_RULE_TYPE_NONE
    minimum = C.LXW_CONDITIONAL_RULE_TYPE_MINIMUM
    number = C.LXW_CONDITIONAL_RULE_TYPE_NUMBER
    percent = C.LXW_CONDITIONAL_RULE_TYPE_PERCENT
    percentile = C.LXW_CONDITIONAL_RULE_TYPE_PERCENTILE
    formula = C.LXW_CONDITIONAL_RULE_TYPE_FORMULA
    maximum = C.LXW_CONDITIONAL_RULE_TYPE_MAXIMUM
    auto_min = C.LXW_CONDITIONAL_RULE_TYPE_AUTO_MIN
    auto_max = C.LXW_CONDITIONAL_RULE_TYPE_AUTO_MAX
}

pub enum ConditionalFormatBarDirection { // lxw_conditional_format_bar_direction
    context = C.LXW_CONDITIONAL_BAR_DIRECTION_CONTEXT
    right_to_left = C.LXW_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT
    left_to_right = C.LXW_CONDITIONAL_BAR_DIRECTION_LEFT_TO_RIGHT
}

pub enum ConditionaBarAxisPosition { // lxw_conditional_bar_axis_position 
    automatic = C.LXW_CONDITIONAL_BAR_AXIS_AUTOMATIC
    midpoint = C.LXW_CONDITIONAL_BAR_AXIS_MIDPOINT
    @none = C.LXW_CONDITIONAL_BAR_AXIS_NONE
}

pub enum ConditionalIconTypes { // lxw_conditional_icon_types 
    type_3_arrows_colored = C.LXW_CONDITIONAL_ICONS_3_ARROWS_COLORED
    type_3_arrows_gray = C.LXW_CONDITIONAL_ICONS_3_ARROWS_GRAY
    type_3_flags = C.LXW_CONDITIONAL_ICONS_3_FLAGS
    type_3_traffic_lights_unrimmed = C.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED
    type_3_traffic_lights_rimmed = C.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_RIMMED
    type_3_signs = C.LXW_CONDITIONAL_ICONS_3_SIGNS
    type_3_symbol_circled = C.LXW_CONDITIONAL_ICONS_3_SYMBOLS_CIRCLED
    type_3_symbols_uncircled = C.LXW_CONDITIONAL_ICONS_3_SYMBOLS_UNCIRCLED
    type_4_arrows_colored = C.LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED
    type_4_arrows_gray = C.LXW_CONDITIONAL_ICONS_4_ARROWS_GRAY
    type_4_red_to_black = C.LXW_CONDITIONAL_ICONS_4_RED_TO_BLACK
    type_4_rtings = C.LXW_CONDITIONAL_ICONS_4_RATINGS
    type_4_traffic_lights = C.LXW_CONDITIONAL_ICONS_4_TRAFFIC_LIGHTS
    type_5_arrows_colored = C.LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED
    type_5_arrows_gray = C.LXW_CONDITIONAL_ICONS_5_ARROWS_GRAY
    type_5_ratings = C.LXW_CONDITIONAL_ICONS_5_RATINGS
    type_5_quarters = C.LXW_CONDITIONAL_ICONS_5_QUARTERS
}

pub enum TableStyleType { // lxw_table_style_type 
    default = C.LXW_TABLE_STYLE_TYPE_DEFAULT
    light = C.LXW_TABLE_STYLE_TYPE_LIGHT
    medium = C.LXW_TABLE_STYLE_TYPE_MEDIUM
    dark = C.LXW_TABLE_STYLE_TYPE_DARK
}

pub enum TableTotalFunctions { // lxw_table_total_functions 
    @none = C.LXW_TABLE_FUNCTION_NONE
    average = C.LXW_TABLE_FUNCTION_AVERAGE 
    count_nums = C.LXW_TABLE_FUNCTION_COUNT_NUMS 
    count = C.LXW_TABLE_FUNCTION_COUNT 
    max = C.LXW_TABLE_FUNCTION_MAX 
    min = C.LXW_TABLE_FUNCTION_MIN 
    std_dev = C.LXW_TABLE_FUNCTION_STD_DEV 
    sum = C.LXW_TABLE_FUNCTION_SUM 
    var = C.LXW_TABLE_FUNCTION_VAR 
}

pub enum FilterCriteria { // lxw_filter_criteria 
    @none = C.LXW_FILTER_CRITERIA_NONE
    equal_to = C.LXW_FILTER_CRITERIA_EQUAL_TO
    not_equal_to = C.LXW_FILTER_CRITERIA_NOT_EQUAL_TO
    greater_than = C.LXW_FILTER_CRITERIA_GREATER_THAN 
    less_than = C.LXW_FILTER_CRITERIA_LESS_THAN 
    greater_than_or_equal_to = C.LXW_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO
    less_than_or_equal_to = C.LXW_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO
    blanks = C.LXW_FILTER_CRITERIA_BLANKS
    non_blanks = C.LXW_FILTER_CRITERIA_NON_BLANKS
}

pub enum FilterOperator { // lxw_filter_operator
    and = C.LXW_FILTER_AND
    @or = C.LXW_FILTER_OR
}

pub enum FilterTypes { // lxw_filter_type 
    type_none = C.LXW_FILTER_TYPE_NONE
    type_single = C.LXW_FILTER_TYPE_SINGLE
    type_and = C.LXW_FILTER_TYPE_AND
    type_or = C.LXW_FILTER_TYPE_OR
    type_string_list = C.LXW_FILTER_TYPE_STRING_LIST
}

pub enum ObjectPosition { // lxw_object_position
    default = C.LXW_OBJECT_POSITION_DEFAULT
    move_and_size = C.LXW_OBJECT_MOVE_AND_SIZE
    move_dont_size = C.LXW_OBJECT_MOVE_DONT_SIZE
    dont_move_dont_size = C.LXW_OBJECT_DONT_MOVE_DONT_SIZE
    move_and_size_after = C.LXW_OBJECT_MOVE_AND_SIZE_AFTER
}

pub enum IgnoreErrors { // lxw_ignore_errors
    number_stored_as_text = C.LXW_IGNORE_NUMBER_STORED_AS_TEXT
    eval_error = C.LXW_IGNORE_EVAL_ERROR
    formula_differs = C.LXW_IGNORE_FORMULA_DIFFERS
    formula_range = C.LXW_IGNORE_FORMULA_RANGE
    formula_unlocked = C.LXW_IGNORE_FORMULA_UNLOCKED
    empty_cell_reference = C.LXW_IGNORE_EMPTY_CELL_REFERENCE
    list_data_validation = C.LXW_IGNORE_LIST_DATA_VALIDATION
    calculated_column = C.LXW_IGNORE_CALCULATED_COLUMN
    two_digit_text_year = C.LXW_IGNORE_TWO_DIGIT_TEXT_YEAR
    last_option = C.LXW_IGNORE_LAST_OPTION
}

pub enum CellTypes { // cell_types
    number_cell = C.NUMBER_CELL
    string_cell = C.STRING_CELL
    inline_string_cell = C.INLINE_STRING_CELL
    inline_rich_string_cell = C.INLINE_RICH_STRING_CELL
    formula_cell = C.FORMULA_CELL
    array_formula_cell = C.ARRAY_FORMULA_CELL
    dynamic_array_formula_cell = C.DYNAMIC_ARRAY_FORMULA_CELL
    blank_cell = C.BLANK_CELL
    boolean_cell = C.BOOLEAN_CELL
    comment = C.COMMENT
    hyperlink_url = C.HYPERLINK_URL
    hyperlink_interlal = C.HYPERLINK_INTERNAL
    hyperlink_external = C.HYPERLINK_EXTERNAL
}

pub enum PaneTypes { // pane_types
    no_panes = C.NO_PANES
    freeze_panes = C.FREEZE_PANES
    split_panes = C.SPLIT_PANES
    freeze_split_panes = C.FREEZE_SPLIT_PANES
}

pub enum ImagePosition { // lxw_image_position
    header_left = C.HEADER_LEFT
    header_center = C.HEADER_CENTER
    header_right = C.HEADER_RIGHT
    footer_left = C.FOOTER_LEFT
    footer_center = C.FOOTER_CENTER
    footer_right = C.FOOTER_RIGHT
}

struct C.lxw_table_rows {}
struct C.lxw_row_col_options {}
struct C.lxw_col_options {}
struct C.lxw_merged_range {}
struct C.lxw_repeat_rows {}
struct C.lxw_repeat_cols {}
struct C.lxw_print_area {}
struct C.lxw_autofilter {}
struct C.lxw_panes {}
struct C.lxw_selection {}
struct C.lxw_data_validation {}
struct C.lxw_data_val_obj {}
struct C.lxw_conditional_format {}
struct C.lxw_cond_format_obj {}
struct C.lxw_cond_format_hash_element {}
struct C.lxw_table_column {}
struct C.lxw_table_options {}
struct C.lxw_table_obj {}
struct C.lxw_filter_rule {}
struct C.lxw_filter_rule_obj {}
struct C.lxw_image_options {}
struct C.lxw_chart_options {}
struct C.lxw_object_properties {}
struct C.lxw_comment_options {}
struct C.lxw_button_options {}
struct C.lxw_vml_obj {}
struct C.lxw_header_footer_options {}
struct C.lxw_protection {}
struct C.lxw_protection_obj {}
struct C.lxw_rich_string_tuple {}
struct C.lxw_worksheet_init_data {}
struct C.lxw_row {}
struct C.lxw_cell {}
struct C.lxw_drawing_rel_id {}

// Chartsheet 
struct C.lxw_chartsheet {
}
pub struct Chartsheet {
    ptr &C.lxw_chartsheet
}

fn C.workbook_add_chartsheet(workbook &C.lxw_workbook, sheetname &char) &C.lxw_chartsheet
pub fn (wb Workbook)add_chartsheet(sheet_name string) Chartsheet {
    return Chartsheet {
        ptr: C.workbook_add_chartsheet(wb.ptr, &char(sheet_name.str))
    }
}

fn C.workbook_get_chartsheet_by_name(workbook &C.lxw_workbook, name &char) &C.lxw_chartsheet
pub fn (wb Workbook)get_chartsheet_by_name(name string) Chartsheet {
    return Chartsheet {
        ptr: C.workbook_get_chartsheet_by_name(wb.ptr, &char(name.str))
    }
}

fn C.chartsheet_set_chart(chartsheet &C.lxw_chartsheet, chart &C.lxw_chart) C.lxw_error
pub fn (cs Chartsheet)set_chart(chart Chart) ReturnCode {
    return unsafe {
        ReturnCode(C.chartsheet_set_chart(cs.ptr, chart.ptr))
    }
}

// TODO
fn C.chartsheet_set_chart_opt(chartsheet &C.lxw_chartsheet, chart &C.lxw_chart, user_options &C.lxw_chart_options) C.lxw_error

fn C.chartsheet_activate(chartsheet &C.lxw_chartsheet)
pub fn (cs Chartsheet)activate() {
    C.chartsheet_activate(cs.ptr)
}

fn C.chartsheet_select(chartsheet &C.lxw_chartsheet)
pub fn (cs Chartsheet)highlight() {
    C.chartsheet_select(cs.ptr)
}

fn C.chartsheet_hide(chartsheet &C.lxw_chartsheet)
pub fn (cs Chartsheet)hide() {
    C.chartsheet_hide(cs.ptr)
}

fn C.chartsheet_set_first_sheet(chartsheet &C.lxw_chartsheet)
pub fn (cs Chartsheet)set_first_sheet() {
    C.chartsheet_set_first_sheet(cs.ptr)
}

fn C.chartsheet_set_tab_color(chartsheet &C.lxw_chartsheet, color C.lxw_color_t)
pub fn (cs Chartsheet)set_tab_color(color DefinedColors) {
    unsafe {
        C.chartsheet_set_tab_color(cs.ptr, C.lxw_color_t(color))
    }
}

// TODO 
fn C.chartsheet_protect(chartsheet &C.lxw_chartsheet, password &char, options &C.lxw_protection)

fn C.chartsheet_set_zoom(chartsheet &C.lxw_chartsheet, scale u16)
pub fn (cs Chartsheet)set_zoom(scale u16) {
    C.chartsheet_set_zoom(cs.ptr, scale)
}

fn C.chartsheet_set_landscape(chartsheet &C.lxw_chartsheet)
pub fn (cs Chartsheet)set_landscape() {
    C.chartsheet_set_landscape(cs.ptr)
}

fn C.chartsheet_set_portrait(chartsheet &C.lxw_chartsheet)
pub fn (cs Chartsheet)set_portrait() {
    C.chartsheet_set_portrait(cs.ptr)
}

fn C.chartsheet_set_paper(chartsheeet &C.lxw_chartsheet, paper_type u8)
pub fn (cs Chartsheet)set_paper(paper_type u8) {
    C.chartsheet_set_paper(cs.ptr, paper_type)
}

fn C.chartsheet_set_margins(chartsheet &C.lxw_chartsheet, left f64, right f64, top f64, bottom f64)
pub fn (cs Chartsheet)set_margines(left f64, right f64, top f64, bottom f64) {
    C.chartsheet_set_margins(cs.ptr, left, right, top, bottom)
}

fn C.chartsheet_set_header(chartsheet &C.lxw_chartsheet, text &char) C.lxw_error 
pub fn (cs Chartsheet)set_header(text string) ReturnCode {
    return unsafe {
        ReturnCode(C.chartsheet_set_header(cs.ptr, &char(text.str)))
    }
}

fn C.chartsheet_set_footer(chartsheet &C.lxw_chartsheet, text &char) C.lxw_error 
pub fn (cs Chartsheet)set_footer(text string) ReturnCode {
    return unsafe {
        ReturnCode(C.chartsheet_set_footer(cs.ptr, &char(text.str)))
    }
}

// TODO 
fn C.chartsheet_set_header_opt(chartsheet &C.lxw_chartsheet, text &char, options &C.lxw_header_footer_options) C.lxw_error
fn C.chartsheet_set_footer_opt(chartsheet &C.lxw_chartsheet, text &char, options &C.lxw_header_footer_options) C.lxw_error 

fn C.lxw_chartsheet_new(init_data &C.lxw_worksheet_init_data) &C.lxw_chartsheet

fn C.lxw_chartsheet_free(chartsheet &C.lxw_chartsheet)
pub fn (cs Chartsheet)deallocate() {
    C.lxw_chartsheet_free(cs.ptr)
}


// Chart
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


pub enum ChartLegendPosition as u8 { // lxw_chart_legend_position
    @none = C.LXW_CHART_LEGEND_NONE
    right = C.LXW_CHART_LEGEND_RIGHT
    left = C.LXW_CHART_LEGEND_LEFT
    top = C.LXW_CHART_LEGEND_TOP
    bottom = C.LXW_CHART_LEGEND_BOTTOM
    top_right = C.LXW_CHART_LEGEND_TOP_RIGHT
    overlay_right = C.LXW_CHART_LEGEND_OVERLAY_RIGHT
    overlay_left = C.LXW_CHART_LEGEND_OVERLAY_LEFT
    overlay_top_right = C.LXW_CHART_LEGEND_OVERLAY_TOP_RIGHT
}

pub enum ChartLineDashType as u8 { // lxw_chart_line_dash_type 
    solid = C.LXW_CHART_LINE_DASH_SOLID
    round_dot = C.LXW_CHART_LINE_DASH_ROUND_DOT
    sqaure_dot = C.LXW_CHART_LINE_DASH_SQUARE_DOT
    dash = C.LXW_CHART_LINE_DASH_DASH
    dash_dot = C.LXW_CHART_LINE_DASH_DASH_DOT
    long_dash = C.LXW_CHART_LINE_DASH_LONG_DASH
    long_dash_dot = C.LXW_CHART_LINE_DASH_LONG_DASH_DOT
    long_dash_dot_dot = C.LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT
    line_dash_dot = C.LXW_CHART_LINE_DASH_DOT
    system_dash_dot = C.LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT
    system_dash_dot_dot = C.LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT_DOT
}

pub enum ChartMarkerType as u8 { // lxw_chart_marker_type 
    automatic = C.LXW_CHART_MARKER_AUTOMATIC
    @none = C.LXW_CHART_MARKER_NONE
    sqaure = C.LXW_CHART_MARKER_SQUARE
    diamond = C.LXW_CHART_MARKER_DIAMOND
    triangle = C.LXW_CHART_MARKER_TRIANGLE
    marker_x = C.LXW_CHART_MARKER_X
    star = C.LXW_CHART_MARKER_STAR
    short_dash = C.LXW_CHART_MARKER_SHORT_DASH
    long_dash = C.LXW_CHART_MARKER_LONG_DASH
    circle = C.LXW_CHART_MARKER_CIRCLE
    plus = C.LXW_CHART_MARKER_PLUS
}

pub enum ChartPatternType as u8 { // lxw_chart_pattern_type 
    @none = C.LXW_CHART_PATTERN_NONE
    percent_5 = C.LXW_CHART_PATTERN_PERCENT_5
    percent_10 = C.LXW_CHART_PATTERN_PERCENT_10
    percent_20 = C.LXW_CHART_PATTERN_PERCENT_20
    percent_25 = C.LXW_CHART_PATTERN_PERCENT_25
    percent_30 = C.LXW_CHART_PATTERN_PERCENT_30
    percent_40 = C.LXW_CHART_PATTERN_PERCENT_40
    percent_50 = C.LXW_CHART_PATTERN_PERCENT_50
    percent_60 = C.LXW_CHART_PATTERN_PERCENT_60
    percent_70 = C.LXW_CHART_PATTERN_PERCENT_70
    percent_75 = C.LXW_CHART_PATTERN_PERCENT_75
    percent_80 = C.LXW_CHART_PATTERN_PERCENT_80
    percent_90 = C.LXW_CHART_PATTERN_PERCENT_90
    light_downward_diagonal = C.LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL
    light_upward_diagonal = C.LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL
    dark_downward_diagonal = C.LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL
    dark_upward_diagonal = C.LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL
    wide_downwrd_diagonal = C.LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL
    wide_upward_diagonal = C.LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL
    light_vertical = C.LXW_CHART_PATTERN_LIGHT_VERTICAL
    light_horizontal = C.LXW_CHART_PATTERN_LIGHT_HORIZONTAL
    narrow_vertical = C.LXW_CHART_PATTERN_NARROW_VERTICAL
    narrow_horizontal = C.LXW_CHART_PATTERN_NARROW_HORIZONTAL
    dark_vertical = C.LXW_CHART_PATTERN_DARK_VERTICAL
    dark_horizontal = C.LXW_CHART_PATTERN_DARK_HORIZONTAL
    dashed_downward_diagonal = C.LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL
    dahsed_upward_diagonal = C.LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL
    dashed_horizontal = C.LXW_CHART_PATTERN_DASHED_HORIZONTAL
    dashed_vertical = C.LXW_CHART_PATTERN_DASHED_VERTICAL
    small_confetti = C.LXW_CHART_PATTERN_SMALL_CONFETTI
    large_confetti = C.LXW_CHART_PATTERN_LARGE_CONFETTI
    zigzag = C.LXW_CHART_PATTERN_ZIGZAG
    wave = C.LXW_CHART_PATTERN_WAVE
    diagonal_brick = C.LXW_CHART_PATTERN_DIAGONAL_BRICK
    horizontal_brick = C.LXW_CHART_PATTERN_HORIZONTAL_BRICK
    weave = C.LXW_CHART_PATTERN_WEAVE
    plaid = C.LXW_CHART_PATTERN_PLAID
    divot = C.LXW_CHART_PATTERN_DIVOT
    dotted_grid = C.LXW_CHART_PATTERN_DOTTED_GRID
    dotted_diamond = C.LXW_CHART_PATTERN_DOTTED_DIAMOND
    shingle = C.LXW_CHART_PATTERN_SHINGLE
    trellis = C.LXW_CHART_PATTERN_TRELLIS
    sphere = C.LXW_CHART_PATTERN_SPHERE
    small_grid = C.LXW_CHART_PATTERN_SMALL_GRID
    large_grid = C.LXW_CHART_PATTERN_LARGE_GRID
    small_check = C.LXW_CHART_PATTERN_SMALL_CHECK
    large_check = C.LXW_CHART_PATTERN_LARGE_CHECK
    outlined_diamong = C.LXW_CHART_PATTERN_OUTLINED_DIAMOND
    solid_diamond = C.LXW_CHART_PATTERN_SOLID_DIAMOND
}

pub enum ChartLabelPosition as u8 { // lxw_chart_label_position 
    default = C.LXW_CHART_LABEL_POSITION_DEFAULT
    center = C.LXW_CHART_LABEL_POSITION_CENTER
    right = C.LXW_CHART_LABEL_POSITION_RIGHT
    left = C.LXW_CHART_LABEL_POSITION_LEFT
    above = C.LXW_CHART_LABEL_POSITION_ABOVE
    below = C.LXW_CHART_LABEL_POSITION_BELOW
    inside_base = C.LXW_CHART_LABEL_POSITION_INSIDE_BASE
    inside_end = C.LXW_CHART_LABEL_POSITION_INSIDE_END
    outside_end = C.LXW_CHART_LABEL_POSITION_OUTSIDE_END
    best_fit = C.LXW_CHART_LABEL_POSITION_BEST_FIT
}

pub enum CharLabelSeparator as u8 { // lxw_chart_label_separator 
    comma = C.LXW_CHART_LABEL_SEPARATOR_COMMA
    semicolon = C.LXW_CHART_LABEL_SEPARATOR_SEMICOLON
    period = C.LXW_CHART_LABEL_SEPARATOR_PERIOD
    newline = C.LXW_CHART_LABEL_SEPARATOR_NEWLINE
    space = C.LXW_CHART_LABEL_SEPARATOR_SPACE
}

pub enum ChartAxisType as u8 { // lxw_chart_axis_type 
    type_x = C.LXW_CHART_AXIS_TYPE_X
    type_y = C.LXW_CHART_AXIS_TYPE_Y
}

pub enum ChartSubtype as u8 { // lxw_chart_subtype 
    @none = C.LXW_CHART_SUBTYPE_NONE
    stacked = C.LXW_CHART_SUBTYPE_STACKED
    stacked_percent = C.LXW_CHART_SUBTYPE_STACKED_PERCENT
}

pub enum ChartGrouping as u8 { // lxw_chart_grouping
    clustered = C.LXW_GROUPING_CLUSTERED
    standrd = C.LXW_GROUPING_STANDARD
    percent_stacked = C.LXW_GROUPING_PERCENTSTACKED
    stacked = C.LXW_GROUPING_STACKED
}

pub enum ChartAxisTickPosition as u8 { // lxw_chart_axis_tick_position 
    default = C.LXW_CHART_AXIS_POSITION_DEFAULT
    on_tick = C.LXW_CHART_AXIS_POSITION_ON_TICK
    between = C.LXW_CHART_AXIS_POSITION_BETWEEN
}

pub enum ChartAxisLabelPosition as u8 { // lxw_chart_axis_label_position 
    next_to = C.LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO
    @none = C.LXW_CHART_AXIS_LABEL_POSITION_NONE
}

pub enum ChartAxisLabelAlignment as u8 { // lxw_chart_axis_label_alignment 
    center = C.LXW_CHART_AXIS_LABEL_ALIGN_CENTER
    left = C.LXW_CHART_AXIS_LABEL_ALIGN_LEFT
    right = C.LXW_CHART_AXIS_LABEL_ALIGN_RIGHT
}

pub enum ChartAxisDisplayUnit as u8 { // lxw_chart_axis_display_unit 
    @none = C.LXW_CHART_AXIS_UNITS_NONE
    hundreds = C.LXW_CHART_AXIS_UNITS_HUNDREDS
    thousands = C.LXW_CHART_AXIS_UNITS_THOUSANDS
    ten_thousands = C.LXW_CHART_AXIS_UNITS_TEN_THOUSANDS
    hundredthousends = C.LXW_CHART_AXIS_UNITS_HUNDRED_THOUSANDS
    millions = C.LXW_CHART_AXIS_UNITS_MILLIONS
    ten_millions = C.LXW_CHART_AXIS_UNITS_TEN_MILLIONS
    hundred_millions = C.LXW_CHART_AXIS_UNITS_HUNDRED_MILLIONS
    billions = C.LXW_CHART_AXIS_UNITS_BILLIONS
    trillions = C.LXW_CHART_AXIS_UNITS_TRILLIONS
}

pub enum ChartAxisTickMark as u8 { // lxw_chart_axis_tick_mark 
    default = C.LXW_CHART_AXIS_TICK_MARK_DEFAULT
    @none = C.LXW_CHART_AXIS_TICK_MARK_NONE
    inside = C.LXW_CHART_AXIS_TICK_MARK_INSIDE
    outside = C.LXW_CHART_AXIS_TICK_MARK_OUTSIDE
    crossing = C.LXW_CHART_AXIS_TICK_MARK_CROSSING
}

struct C.lxw_chart {
}

pub struct Chart {
    ptr &C.lxw_chart 
}

fn C.workbook_add_chart(wb &C.lxw_workbook, chart_type u8) &C.lxw_chart 
pub fn (wb Workbook)add_chart(chart_type ChartType) Chart {
    return unsafe { Chart {
        ptr: C.workbook_add_chart(wb.ptr, u8(chart_type))
    }}
}

// DocProperties
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

// Format
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

// Check number_format.md for more details abut index and format strings
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

// Error 
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

fn C.lxw_strerror(error_num C.lxw_error) &char 
pub fn (rc ReturnCode)message() string {
    return unsafe {
        cstring_to_vstring(C.lxw_strerror(C.lxw_error(rc)))
    }
}

fn C.lxw_version_id() u16
pub fn version_id() u16 {
    return C.lxw_version_id()
}

fn C.lxw_version() &char
pub fn version() string {
    return unsafe {
        cstring_to_vstring(C.lxw_version())
    }
}

// DateTime
struct C.lxw_datetime {
    /** Year     : 1900 - 9999 */
    year int
    /** Month    : 1 - 12 */
    month int
    /** Day      : 1 - 31 */
    day int
    /** Hour     : 0 - 23 */
    hour int
    /** Minute   : 0 - 59 */
    min int
    /** Seconds  : 0 - 59.999 */
    sec f64 
}

pub type DateTime = C.lxw_datetime

