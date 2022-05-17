import itertools
import win32com.client as win_cl


def formatting_file(ab_path, wb_name, wb_name_ab_path, file1, file2, file3):
    excel_macro = win_cl.DispatchEx("Excel.application")
    workbook4 = excel_macro.Workbooks.Open(Filename=ab_path)
    excel_macro.Visible = 1
    for k, l in itertools.zip_longest(wb_name, wb_name_ab_path):
        workbook = excel_macro.Workbooks.Open(Filename=l)
        excel_macro.Application.Run("%s!test1" % file1, k)
        if k == file2:
            excel_macro.Application.Run("%s!test2" % file1, k)
        elif k == file3:
            excel_macro.Application.Run("%s!test3" % file1, k)
        print("Formatting done for ", l)
        workbook.Save()
        workbook.Close()
    # workbook4.Save()
    workbook4.Close()
    excel_macro.Application.Quit()
