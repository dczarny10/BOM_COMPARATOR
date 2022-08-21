import openpyxl, re, csv, os
import win32com.client as win32
import tempfile


def load_files(paths):
    pattern = "(\d+-?\w?\d+-?\d*)-(.+)"
    data_SAP, results_PLM, products_list, data = [], [], [], []
    for file in paths:
        if file.lower().endswith("xls"):
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(file)
            excel.DisplayAlerts = False
            excel.Visible = False
            wb.SaveAs(tempfile.gettempdir()+"BOM_COMPARATOR_TEMP.xlsx", FileFormat=51)
            wb.Close()
            excel.Application.Quit()
            wb_SAP = openpyxl.load_workbook(file+"x")
            ws_SAP = wb_SAP.active
            ws_SAP.delete_cols(1, 1)
            ws_SAP.delete_rows(1, 1)
            ws_SAP.delete_rows(2, 1)
            ws_SAP.delete_cols(5, 10)
            for q in ws_SAP.iter_rows(values_only=True):
                data_SAP.append(tuple(map(str, (str(q[0])+"_SAP", q[1], q[2], q[3]))))
                if not str(q[0])+"_SAP" in products_list:
                    products_list.append(str(q[0])+"_SAP")
            os.remove(tempfile.gettempdir()+"BOM_COMPARATOR_TEMP.xlsx")
            continue
        wb_PLM = openpyxl.load_workbook(file)
        ws_PLM = wb_PLM.active
        data_PLM = list(ws_PLM.iter_rows(values_only=True))
        #products = ["-".join(i.split("-")[:-2]) for i in data_PLM[0][3:]]
        products = [i for i in data_PLM[0][3:]]
        data_PLM = data_PLM[1:]
        for count, i in enumerate(products):
            for j in data_PLM:
                if j[3+count] != "0":
                    r = re.findall(pattern, j[0])
                    results_PLM.append((i, r[0][0], j[3 + count].split(".")[0], r[0][1]))
                    if not i in products_list:
                        products_list.append(i)

    return  products_list,  results_PLM + data_SAP
