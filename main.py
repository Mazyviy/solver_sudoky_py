# -*- coding: utf-8 -*-

import os
import cv2
import pytesseract as pytesseract
from PIL import Image
import datetime
import win32api, win32con, time
import copy

pytesseract.pytesseract.tesseract_cmd = r"E:\tesseract\tesseract.exe"

def readImage():
    board = []
    imageOrigin = Image.open('origin.png')
    imageResize = imageOrigin.resize((523, 930))
    imageResize.save('origin.png')

    widthAndHeight = 52
    i1 = 0
    y = 135
    x = 26

    # цикл for
    while y < 603:
        board.append([])
        while x < 494:
            image_cut = imageResize.crop((x+5, y+5, x + widthAndHeight-5, y + widthAndHeight-5))
            image_cut.save('number.png')
            imageTesseract = cv2.imread('number.png')
            gray = cv2.cvtColor(imageTesseract, cv2.COLOR_BGR2GRAY)
            imageTesseract = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
            number = pytesseract.image_to_string(imageTesseract, lang='eng', \
                     config='--psm 10 --oem 3 -c tessedit_char_whitelist=0123456789')
            os.remove('number.png')

            if number == '':
                board[i1].append(0)
            else:
                board[i1].append(int(number))

            x = x + widthAndHeight
        x = 26
        i1 = i1 + 1
        y = y + widthAndHeight

    return board


def findNextCellToFill(grid, i, j):
    for x in range(i, 9):
        for y in range(j, 9):
            if grid[x][y] == 0:
                return x, y
    for x in range(0, 9):
        for y in range(0, 9):
            if grid[x][y] == 0:
                return x, y
    return -1, -1


def isValid(grid, i, j, e):
    rowOk = all([e != grid[i][x] for x in range(9)])
    if rowOk:
        columnOk = all([e != grid[x][j] for x in range(9)])
        if columnOk:
            # finding the top left x,y co-ordinates of the section containing the i,j cell
            secTopX, secTopY = 3 * (i // 3), 3 * (j // 3)  # floored quotient should be used here.
            for x in range(secTopX, secTopX + 3):
                for y in range(secTopY, secTopY + 3):
                    if grid[x][y] == e:
                        return False
            return True
    return False


def solveSudoku(grid, i=0, j=0):
    i, j = findNextCellToFill(grid, i, j)
    if i == -1:
        return True
    for e in range(1, 10):
        if isValid(grid, i, j, e):
            grid[i][j] = e
            if solveSudoku(grid, i, j):
                return True
            # Undo the current cell for backtracking
            grid[i][j] = 0
    return False


print("window screen in 4 sec")
time.sleep(3)
win32api.keybd_event(win32con.VK_CONTROL, 0x1E, 0, 0);
win32api.keybd_event(win32con.VK_SNAPSHOT, 0x1E, 0, 0);
win32api.keybd_event(win32con.VK_SNAPSHOT, 0x1E, win32con.KEYEVENTF_KEYUP, 0);
win32api.keybd_event(win32con.VK_CONTROL, 0x1E, win32con.KEYEVENTF_KEYUP, 0);
time.sleep(0.4)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

time.sleep(0.5)
today = datetime.datetime.today()
os.system('copy D:\\Documents\\ShareX\\Screenshots\\2022-09\\' +
          today.strftime("%Y.%m.%d_%H-%M") +
          '.png ' + os.getcwd())

if os.path.exists(os.getcwd()+'\\'+ 'origin.png'):
    os.remove(os.getcwd()+'\\'+ 'origin.png')

os.system('rename ' + os.getcwd()+'\\'+today.strftime("%Y.%m.%d_%H-%M") + '.png origin.png')

time.sleep(0.5)

board = readImage()
boardZero = copy.deepcopy(board)
for i in range(9):
    print(board[i])

solveSudoku(board)

print("Solver:")
for i in range(9):
    print(board[i])

print("Window click")
time.sleep(1)

cursor_first = win32api.GetCursorPos()
key = {1: win32con.VK_NUMPAD1, 2: win32con.VK_NUMPAD2, 3: win32con.VK_NUMPAD3,
       4: win32con.VK_NUMPAD4, 5: win32con.VK_NUMPAD5, 6: win32con.VK_NUMPAD6,
       7: win32con.VK_NUMPAD7, 8: win32con.VK_NUMPAD8, 9: win32con.VK_NUMPAD9}

i1 = 0
j1 = 0
x = 0
y = 0
w = 52
h = 52

while y < 468:
    while x < 468:
        if boardZero[i1][j1] == 0:
            win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE,
                                 int(65535 * (cursor_first[0] + x) / win32api.GetSystemMetrics(win32con.SM_CXSCREEN)),
                                 int(65535 * (cursor_first[1] + y) / win32api.GetSystemMetrics(win32con.SM_CYSCREEN)),
                                 0, 0)

            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

            win32api.keybd_event(key[board[i1][j1]], 0x1E, 0, 0);
            win32api.keybd_event(key[board[i1][j1]], 0x1E, win32con.KEYEVENTF_KEYUP, 0);
            time.sleep(0.5)

        x = x + 2
        x = x + w
        j1 = j1 + 1

    x = 0
    j1 = 0
    i1 = i1 + 1
    y = y + 2
    y = y + h
