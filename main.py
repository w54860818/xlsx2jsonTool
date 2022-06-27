import json
import os

import xlrd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

jsonList = []
elementName = []


def writeToJson(dictionPath, fileName):
    dirPath = os.path.join(dictionPath, "json_save")
    filePath = os.path.join(dirPath, fileName + ".json")
    if not os.path.isdir(dirPath):
        os.mkdir(dirPath)
    if os.path.isfile(filePath):
        os.remove(filePath)
    file = open(filePath, 'w', encoding="utf-8")
    json.dump(jsonList, file, ensure_ascii=False, indent=1)


def readFromExcel(openPath):
    file = xlrd.open_workbook(openPath)
    sheet = file.sheet_by_index(0)
    rows = sheet.nrows
    cols = sheet.ncols
    if rows > 0:
        elementName = sheet.row_values(0)
        for i in range(1, rows):
            row = sheet.row_values(i)
            element = {}
            for j in range(cols):
                if type(row[j]) == float:
                    row[j] = int(row[j])
                element[elementName[j]] = row[j]
            jsonList.append(element)


def openFile():
    originDir = loadDefaultPath()
    openPath = filedialog.askopenfilename(initialdir=originDir, title="请选择xlsx文件", filetypes=[("xlsx文件", "*.xlsx")])

    return openPath


def convertFile():
    openPath = openFile()
    if openPath != "":
        saveDefaultPath(openPath)
        (dictionPath, fileName) = os.path.split(openPath)
        (fileName, extension) = os.path.splitext(fileName)
        readFromExcel(openPath)
        writeToJson(dictionPath, fileName)
        os.startfile(os.path.join(dictionPath, "json_save"))
        messagebox.showinfo(title="成功！",  message=openPath + " 转换成功")


def convertDir():
    originDir = loadDefaultPath()
    openDirPath = filedialog.askdirectory(initialdir=originDir, title="请选择xlsx所在文件夹")
    if openDirPath == "":
        return
    saveDefaultPath(openDirPath)
    for fileName in os.listdir(openDirPath):
        if fileName != "":
            (pre, extension) = os.path.splitext(fileName)
            if extension == ".xlsx":
                openPath = os.path.join(openDirPath, fileName)
                readFromExcel(openPath)
                writeToJson(openDirPath, pre)
    savePath = os.path.join(openDirPath, "json_save")
    if os.path.isdir(savePath):
        os.startfile(os.path.join(openDirPath, "json_save"))
        messagebox.showinfo(title="成功！",  message=openDirPath + " 转换成功")
    else:
        messagebox.showinfo(title="失败！", message="请检查文件夹下是否有xlsx文件")


def saveDefaultPath(path):
    (defaultPath, fileName) = os.path.split(path)
    saveDir = {"defaultPath": defaultPath}
    file = open("defaultPath.json", 'w', encoding="utf-8")
    json.dump(saveDir, file, ensure_ascii=False, indent=1)


def loadDefaultPath():
    if os.path.isfile("defaultPath.json"):
        saveDir = json.load(open("defaultPath.json", 'r', encoding='utf-8'))
        return saveDir['defaultPath']
    else:
        return "C:"


def main():

    root_window = tk.Tk()
    root_window.title("Excel转json小程序")

    text = tk.Label(root_window, text="选择xlsx文件，自动转换成json， \r保存在同一目录下json_save文件夹中")
    text.pack()

    filebutton = tk.Button(root_window, text="打开需要转换的xlsx文件", command=convertFile)
    filebutton.pack()

    text_dir = tk.Label(root_window, text="选择文件夹，将文件夹下的所有xlsx文件转化为json文件，\r保存在该文件夹下json_save文件夹中")
    text_dir.pack()

    dirbutton = tk.Button(root_window, text="打开需要转换的文件夹", command=convertDir)
    dirbutton.pack()

    root_window.mainloop()


if __name__ == '__main__':
    main()

