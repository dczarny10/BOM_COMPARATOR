import openpyxl, re, csv, os
import win32com.client as win32
import tempfile


def load_files(paths):
    pattern = "(\d+-?\w?\d+-?\d*)-(.+)"
    data_SAP, results_PLM, products_list = [], [], []
    data = {}

    for file in paths:
        if file.lower().endswith("xls"):
            excel = win32.gencache.EnsureDispatch('Excel.Application')
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
                index = -1
                for i, obj in enumerate(data_SAP):
                    if obj[0]== (str(q[0])+"_SAP"):
                        if obj[1] == q[1]:
                            index = i
                            break
                if index == -1:
                    data_SAP.append(tuple(map(str, (str(q[0])+"_SAP", q[1], q[2], q[3]))))
                else:
                    data_SAP[index] = (data_SAP[index][0], data_SAP[index][1], str(int(data_SAP[index][2])+int(q[2])), data_SAP[index][3])
                if not str(q[0])+"_SAP" in products_list:
                    products_list.append(str(q[0])+"_SAP")
            os.remove(tempfile.gettempdir()+"BOM_COMPARATOR_TEMP.xlsx")
            continue
        wb_PLM = openpyxl.load_workbook(file)
        ws_PLM = wb_PLM.active
        data_PLM = list(ws_PLM.iter_rows(values_only=True)) #create a list of all rows in excel file
        #products = ["-".join(i.split("-")[:-2]) for i in data_PLM[0][3:]]
        products = [i for i in data_PLM[0][3:]] #create products numbers list
        data_PLM = data_PLM[1:] #remove headlines in excel file
        for count, i in enumerate(products):
            data.setdefault(i, [])
            for j in data_PLM:
                if j[3+count] != "0": #if part belong to this product
                    r = re.findall(pattern, j[0]) #find part number #python regular expression because part numbers are sadly not standarized now
                    index = -1
                    for k, obj in enumerate(results_PLM): #check if part is already under this product BOM
                        if obj[0] == i:
                            if obj[1] == r[0][0]:
                                index = k
                                break
                    if index == -1:
                        data[i].append((r[0][0], j[3 + count].split(".")[0], r[0][1]))
                    else:
                        results_PLM[index] = (results_PLM[index][0], results_PLM[index][1], str(int(results_PLM[index][2]) + int(j[3 + count].split(".")[0])), results_PLM[index][3])
                        ### data[i]=

                    if not i in products_list:
                        products_list.append(i)

    return data
