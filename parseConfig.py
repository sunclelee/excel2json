# -*- coding: utf-8 -*-

import xlrd,json
import codecs

global sheet,worksheetNum
global MergePoints,merges
global text
global headerRow,manualHeaderrow

class node:
    def __init__(self, des, x, y):
        self._des = des
        self._coordinate = [x, y]
        self._children = []
    def getDes(self):
        return self._des
    def getCoordinate(self):
        return self._coordinate
    def getChildren(self):
        return self._children
    def addChild(self, node):
        self._children.append(node)

class tree:
    def __init__(self):
        self._head = node("head", -1, -1)
    def addChild(self, node):
        self._head.addChild(node)
    def getHead(self):
        return self._head
 
def generateTree(rowMax, colMax):
    structTree = tree()
    row = 0
    for col in range(0, colMax):
        endCol = col
        treeAddNodes(structTree, row, col, endCol)
    return structTree

def treeAddNodes(father, row, col, endCol):
    global sheet,headerRow
    global MergePoints
    point = [row, col]
    height,isMergeOnCols,width = isPointInMergePoints(point)
#    print('height',height,'isMergeOnCols',isMergeOnCols,'point',[point,endCol])
    if not isMergeOnCols: #单列的单元格也需要深入多级
        if col < endCol:
            # 遍历叶子的邻接点
            newNode = addNodes(father, row, col)
            if row + height < headerRow:
                treeAddNodes(newNode,row + height,col,col)
            treeAddNodes(father, row, col + 1, endCol)                      
            
        elif col <= n_cols :
            newNode = addNodes(father, row, col)            
#            print('newNode',newNode._des,newNode._coordinate)
            if row + height < headerRow:  
                treeAddNodes(newNode,row + height,col,col) 
            return
    else:
        if not isHeadOfMergePoints(point):
            if col < endCol:
                # 合并单元格的非头部结点只遍历不解析
                treeAddNodes(father, row, col + 1, endCol)
        else:
            # update endCol
            newEndCol = getMergePointsEndCol(point)
            newNode = addNodes(father, row, col)
            # 合并单元格的头部结点子节点遍历
            if row + height < headerRow:
                treeAddNodes(newNode, row + height, col, newEndCol)
            # 合并单元格的头部结点的邻接点遍历
            treeAddNodes(father, row, col + 1, endCol)

def addNodes(father, row, col):
    global sheet
    des = sheet.cell_value(row, col)
    newNode = node(des, row, col)
    father.addChild(newNode)
    return newNode

def generateMergePoints():
    points = []
    for (x, xMax, y, yMax) in merges:
        for i in range(x,xMax):
            for j in range(y,yMax):
                points.append([i, j])
    return points

def isPointInMergePoints(point):
    #判断是否横向合并单元格，并返回高度和宽度
    global merges
    height = 1
    width = 1
    isMergeOnCols = False
    for (x,xMax,y,yMax) in merges:
        if ( x <= point[0] < xMax) and ( y <= point[1] < yMax):
            height = xMax - x
            width = yMax - y
            if yMax - y > 1:
                isMergeOnCols = True
            break
    return height,isMergeOnCols,width

def isHeadOfMergePoints(point):
    global sheet
    if isPointInMergePoints(point) and sheet.cell(point[0], point[1]).ctype != 0:
        return True
    return False

def getStartRow():
    global MergePoints
    global headerRow
    rowMin = 0
    for [row, col] in MergePoints:
        if row > rowMin:
            rowMin = row
    # 标题栏行数就是数据栏的起始行标
    if manualHeaderrow == 0:
        headerRow = rowMin + 1
        return headerRow
    else:
        headerRow = manualHeaderrow
        return manualHeaderrow
    

# 获取当前单元格所在合并单元格的结束列坐标
def getMergePointsEndCol(point):
    global sheet
    for [row, rowMax, col, colMax] in sheet.merged_cells:
        if point[0] == row and point[1] == col:
            return colMax - 1
# 递归法把TREE转为json字符串
def traversingByTree(node, row_values):
    global sheet
    global text
    if node.getChildren() == None:
        return
    else :
        for child in node.getChildren():
            if child.getChildren():  #如果有子节点就递归
                text += '"' + str(child.getDes()).strip().replace(' ','') + '"' + " : {"
                traversingByTree(child, row_values)
            else :  #如果没有子节点就拼接键值和键名
                [row, col] = child.getCoordinate()
                height,isMergeOnCols,width = isPointInMergePoints([row, col])
                if not isMergeOnCols:
#                if row_values[col] != "": # 如果对应的列没有数据,就不记录
#                    text += '"' + str(child.getDes()) + '"' + " : " + '"' + str(row_values[col]).strip() + '"' + ","
                    text += '"' + str(child.getDes()).strip().replace(' ','') + '"' + " : " + '"' + str(row_values[col]).strip().replace(' ','') + '"' + ","
                else:
                    value = ''
                    for i in range(0,width):
                        value += str(row_values[col + i]).strip()
                    text += '"' + str(child.getDes()).strip().replace(' ','') + '"' + " : " + '"' + value + '"' + ","
        text = text + "}, "
def process_excel(src):
    global sheet,worksheetNum
    global MergePoints,merges
    global text,n_cols,headerRow

    data = open_excel(src)
    sheet = data.sheets()[worksheetNum]
    n_rows = sheet.nrows #总行数
    n_cols = sheet.ncols #总列数
    print('总行数',n_rows,'总列数',n_cols)
    print('合并单元格',sheet.merged_cells)
    merges = sheet.merged_cells
    MergePoints = generateMergePoints()
    # 生成树状解析目录
    rowMin = getStartRow()
    print('标题栏行标',headerRow)
    tree = generateTree(n_rows, n_cols)    
    #多行数据分行显示
    for row in range(rowMin, n_rows):
        text = text + "{"
        traversingByTree(tree.getHead(), sheet.row_values(row))
        if row < n_rows - 1:
            text = text + "\n"

def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))

# 此处做一些配置
configDic = {
    "sample.xlsx": "output_4.json"
}
manualHeaderrow = 0   # 手动配置标题行数
worksheetNum = 4    # 工作表编号，从0开始

def main():
    global text
    for key in configDic:
        text = "["
        process_excel(key)
        text += "]"
        file = codecs.open(configDic[key], "w", "utf-8")
#        dict = eval(text)       
        # 对text做json格式化
        result = text.replace(',}','}').replace(', }','}').replace(', ]',']').replace('\n','')  
        result = json.dumps(json.loads(result,strict=False),ensure_ascii=False,indent=4)
        file.write(result)
        file.close()
        print("------->" + configDic[key] + " :导出成功")

if __name__ == '__main__':
    main()