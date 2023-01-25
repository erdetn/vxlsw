module main

import vxlsw

fn main() {
    mut book := vxlsw.new_workbook('ex1_workbook.xls')
    mut sheet1 := book.add_sheet('sheet1')

    sheet1.write_string(0,0, '<0,0>')
    sheet1.write_string(0,1, '<0,1>')

    book.close()

}
