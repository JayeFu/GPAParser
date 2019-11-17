from bs4 import BeautifulSoup
import xlsxwriter


def main():
    filename = "清华大学学生课程学习记录表.html"
    with open(filename, 'r') as f:
        html_text = f.read()
        
    soup = BeautifulSoup(html_text, 'html.parser')
    table = soup.find_all('table')
    gpa = table[4]
    
    outName = input("请指定输出文件名")
    
    if not outName:
        outName = "GPA.xlsx"
    
    print("输出文件名指定为{}".format(outName))

    f = xlsxwriter.Workbook(outName)
    gradableSheet = f.add_worksheet("noPWEX")
    allgradeSheet = f.add_worksheet("all")

    credits = 0
    GPS = 0.0

    gradeDatas = list()
    nongradeDatas = list()
    
    rowMarker = 1

    for row in gpa.find_all('tr')[1:]:
        columns = row.find_all('td')
        cond = columns[3].get_text().strip()
        if cond!='P' and cond!='W' and cond!='EX':
            dataList = list()
            for column in columns:
                dataList.append(column.get_text().strip())
            dataList.append("=C{}*E{}".format(rowMarker, rowMarker))
            credit = int(dataList[2])
            grade = float(dataList[4])
            credits += credit
            GPS += credit * grade
            dataList[2] = credit
            dataList[4] = grade
            gradeDatas.append(dataList)
            rowMarker += 1
        else:
            dataList = list()
            for column in columns:
                dataList.append(column.get_text().strip())
            credit = int(dataList[2])
            dataList[2] = credit
            nongradeDatas.append(dataList)
        

    rowNum = 0

    for dataList in gradeDatas:
        for i in range(len(dataList)):
            gradableSheet.write(rowNum, i, dataList[i])
            allgradeSheet.write(rowNum, i , dataList[i])
        rowNum += 1
        
    gradableSheet.write_formula(rowNum, 2, "=SUM(C{}:C{})".format(1, rowNum))
    gradableSheet.write_formula(rowNum, 6, "=SUM(G{}:G{})".format(1, rowNum))
    gradableSheet.write_formula(rowNum+1, 6, "=G{}/C{}".format(rowNum+1, rowNum+1))
    
    for dataList in nongradeDatas:
        for i in range(len(dataList)):
            allgradeSheet.write(rowNum, i , dataList[i])
        rowNum += 1
    allgradeSheet.write_formula(rowNum, 2, "=SUM(C{}:C{})".format(1, rowNum))
    
    f.close()

    print("GPS={:.1f}".format(GPS))
    print("credits={}".format(credits))
    print("GPA=GPS/credits={:.2f}".format(GPS/credits))
    
if __name__ == "__main__":
    main()