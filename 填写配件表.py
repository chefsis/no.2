import xlutils.copy
import xlrd
import random
sheetName = "Sheet1"

filePath = r"C:\Users\Administrator\Desktop\111.xls"

new_filePath = rf"C:\Users\Administrator\Desktop\333.xls"

def write_excel(sheetName,filePath,new_filePath):
    book = xlrd.open_workbook(filePath,formatting_info = True) # 读取文件
    new_book = xlutils.copy.copy(book)
    old_sheet = book.sheet_by_name(sheetName)
    sheet = new_book.get_sheet(0)

    rowIdex = [] # 存放单价和费率的位置
    for i in ["单价","费率"]:
        rowIdex.append(old_sheet.row_values(0).index(i))
    col_len = len(old_sheet.col_values(0)) # 存放物资编码的索引值

    # excel里插入数据
    for one in range(col_len):
        if one != 0:
            for one1 in rowIdex:
                if one1 == rowIdex[0]:
                    sheet.write(one,one1,round(random.uniform(0,100),2)) # 输入单价的值
                else:
                    sheet.write(one, one1, random.choice(["3%","5%","6%","9%","12%"]))  # 输入费率的值
    new_book.save(new_filePath) # 保存一个新的文件名

write_excel(filePath = filePath,new_filePath = new_filePath,sheetName = sheetName)