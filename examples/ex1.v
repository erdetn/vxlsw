module main

import vxlsw

fn main() {
    mut book := vxlsw.new_workbook('ex1_workbook.xls')
    mut sheet1 := book.add_sheet('sheet1')
    mut format := book.add_format()

    format.set_bold()
    sheet1.write_string(0,0, '<0,0>', format)

    format.set_italic()
    sheet1.write_string(0,1, '<0,1>', format)

    book.close()

}
