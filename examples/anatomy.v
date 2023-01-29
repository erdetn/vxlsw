module main

import vxlsw

fn main() {
    mut workbook := vxlsw.new_workbook('anatomy.xlsx')
    mut worksheet1 := workbook.add_sheet('Demo')
    mut worksheet2 := workbook.add_sheet('')
    
    mut format1 := workbook.add_format()
    mut format2 := workbook.add_format()

    format1.set_bold()
    format2.set_number_format('$#,##0.00')

    worksheet1.set_column(0, 0, 20, vxlsw.Format{voidptr(0)}) 
    worksheet1.write_string(0, 0, 'Peach', vxlsw.Format{voidptr(0)})
    worksheet1.write_string(1, 0, 'Plum',  vxlsw.Format{voidptr(0)})
    worksheet1.write_string(2, 0, 'Pear', format1)
    worksheet1.write_string(3, 0, 'Persimmon', format1)

    worksheet1.write_number(5, 0, 123, vxlsw.Format{voidptr(0)})
    worksheet1.write_number(6, 0, 4567.555, format2)
    worksheet2.write_string(0, 0, 'Some text...', format1)

    rc := workbook.close()
    
    if rc != vxlsw.ReturnCode.no_error {
        println('Error closing workbook:\n${rc.message()}')
    } 
}

