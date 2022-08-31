import openpyxl, re, csv, os
import win32com.client as win32
import tempfile


def load_files(paths):
    pattern = "(\d+-?\w?\d+-?\d*)-(.+)"
    data = {}

    for file in paths:
        if file.lower().endswith("xls"):
            #excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel = win32.Dispatch('Excel.Application')
            wb = excel.Workbooks.Open(file)
            excel.DisplayAlerts = False
            excel.Visible = False
            wb.SaveAs(tempfile.gettempdir()+"BOM_COMPARATOR_TEMP.xlsx", FileFormat=51)
            wb.Close()
            excel.Application.Quit()
            wb_SAP = openpyxl.load_workbook(tempfile.gettempdir()+"BOM_COMPARATOR_TEMP.xlsx")
            ws_SAP = wb_SAP.active
            ws_SAP.delete_cols(1, 1)
            ws_SAP.delete_rows(1, 3)
            ws_SAP.delete_cols(5, 10)
            for q in ws_SAP.iter_rows(values_only=True):
                product_name = str(q[0])+"_SAP"
                data.setdefault(product_name, [])
                index = -1
                for i, obj in enumerate(data[product_name]):
                    if obj[0] == q[1]:
                        index = i
                        break
                if index == -1:
                    data[product_name].append(tuple(map(str, (q[1], q[2], q[3]))))
                else:
                    data[product_name][index] = (data[product_name][index][0], str(int(data[product_name][index][1])+int(q[2])), data[product_name][index][2])
            os.remove(tempfile.gettempdir()+"BOM_COMPARATOR_TEMP.xlsx")
            continue
        wb_PLM = openpyxl.load_workbook(file)
        ws_PLM = wb_PLM.active
        data_PLM = list(ws_PLM.iter_rows(values_only=True)) #create a list of all rows in excel file
        products = [i for i in data_PLM[0][3:]] #create products numbers list
        data_PLM = data_PLM[1:] #remove headlines in excel file
        for count, i in enumerate(products):
            data.setdefault(i, [])
            for j in data_PLM:
                if j[3+count] != "0": #if part belong to this product
                    r = re.findall(pattern, j[0]) #find part number #python regular expression because part numbers are sadly not standarized now
                    index = -1
                    for k, obj in enumerate(data[i]): #check if part is already under this product BOM
                        if obj[0] == r[0][0]:
                            index = k
                            break
                    if index == -1:
                        data[i].append((r[0][0], j[3 + count].split(".")[0], r[0][1]))
                    else:
                        data[i][index] = (data[i][index][0], str(int(data[i][index][1]) + int(j[3 + count].split(".")[0])), data[i][index][2])

    return data
