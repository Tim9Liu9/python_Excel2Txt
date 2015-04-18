__author__ = 'timliu'

import time
import xlrd

def excel2Txt():
    book = xlrd.open_workbook("渠道标签分类表.xlsx")
    print("The number of worksheets is", book.nsheets)
    for sheet in range(book.nsheets):
        sheetName = book.sheet_names()[sheet];
        # 存储的文件名，例如：01标签说明.txt、02预装.txt、10论坛.txt
        txtFileName = ''
        if sheet < 10:
            txtFileName = '0' + str(sheet ) + sheetName + ".txt"
        else:
            txtFileName = str(sheet)  + sheetName + ".txt"
        print(txtFileName)

        # 写文件，w+:追加写文本文件
        file_object = open(txtFileName, 'w+')
        #把str写到文件中，F.write(str)并不会在str后加上一个换行符
        file_object.write('<!-- =====  ' + sheetName + '  ==== -->\n')

        mSheet = book.sheet_by_index(sheet)
        for rx in range(mSheet.nrows):
            # 如果是第1行，说明是列标题，不写入文件：
            if rx == 0:
                continue

            # 取第1列的“渠道标签”， 第3列的“渠道名称”
            # print(mSheet.row(rx))
            # round(2.655, 2) int(2.5)  float('%.0f' % float(jdIDFloat))
            jdIDFloat = mSheet.row_values(rx)[0]
            jdIDStr = 0
            # 先判断是否是浮点型：
            if type(jdIDFloat) == float :
                jdIDInt = int(jdIDFloat)
                # 考虑 5、6为的话，后面加：0，补足7位：
                if jdIDInt <100000 :
                    jdIDStr = str(jdIDInt) + "00"
                    print("5位：" + str(jdIDInt))
                elif jdIDInt < 1000000 :
                    jdIDStr = str(jdIDInt) + "0"
                    print("6位：" + str(jdIDInt))
                else:
                    jdIDStr = str(jdIDInt)

            else:
                jdIDStr = jdIDFloat

            jdName = mSheet.row_values(rx)[2]
            # print(str(jdIDStr) + '-->' + jdName)
            file_object.write('<channel value=“' + 'zr' + jdIDStr + '” />     <!-- ' + jdName + ' -->' + '\n')


        file_object.close()

        #
        # print(mSheet)

    # print("Worksheet name(s):", book.sheet_names())
    # sh = book.sheet_by_index(0)
    # print('-->',sh.name, sh.nrows, sh.ncols)
    # print("Cell D30 is", sh.cell_value(rowx=1, colx=3))
    # for rx in range(sh.nrows):
    #     print(sh.row(rx))



if __name__ == '__main__':

    startTimes = time.time()
    excel2Txt()
    endTimes = time.time()
    times = endTimes - startTimes
    print('共耗时：' + repr(times) + '秒')