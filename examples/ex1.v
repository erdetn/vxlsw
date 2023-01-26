module main

import vxlsw

fn main() {
    mut book := vxlsw.new_workbook('ex1_workbook.xls')
    mut sheet1 := book.add_sheet('sheet1')
    mut format := vxlsw.Format{}

    sheet1.write_string(0,0, '<0,0>', format)
    sheet1.write_string(0,1, '<0,1>', format)

    book.close()

}
