#coding=UTF-8
import xlrd
import tkinter.filedialog
from tkinter import*
def main():
    print("\tPlease press enter and choose data path:")
    input()

    InputPath=tkinter.filedialog.askopenfilename(filetypes=[("Excel表格","xls"),("Excel表格","xlsx")])

    while InputPath=="":
        tkinter._exit()
        exit()

    print(InputPath)
    if (InputPath[len(InputPath)-1]=='s')|( InputPath[len(InputPath)-1]=='x'):
       Workbook=xlrd.open_workbook(InputPath)
       table=Workbook.sheets()[0]
    else:
        print('Fatal error!\a')
        exit()
    print('Enter the start position to read and the number of columns to read\nExample:1 1 3 means read start at column 1 row 1 and read to column 3')
    raw=input()
    error=0
    nums=raw.split(' ',-1)
    for num in nums:
        if num.isdigit()==0:
            error=1
    if len(nums)!=3:
        error=1
    while error==1:
        error=0
        print('Wrong Input!\a')
        print( 'Enter the start position to read and the number of columns to read\nExample:1 2 3 means read start at column 1, row 2 and read to column 4')
        raw = input()
        nums=raw.split(' ',-1)
        for num in nums:
            if num.isdigit()==0:
                error=1
        if len(nums)!=3:
            error=1
    srtCol=int(nums[0])
    srtRow=int(nums[1])
    readLen=int(nums[2])
    if readLen==0:
        print('Read length can not be 0!')
        exit()

    for i in range(srtCol,srtCol+readLen-1):
        if not(int(table.cell(i, srtRow).value) >= 0 & int(table.cell(i, srtRow).value) <= 150):
            print('Bad data!\a')
            exit()

    #if
    '''
       tmp1=tmp2=0;
        if int(table.cell(0,0).value)>=0&int(table.cell(0,0).value)<=150:
             cell_a1 = table.cell(0, 0).value
             tmp1=tmp2=0
        elif int(table.cell(0,1).value)>=0&int(table.cell(0,1).value)<=150:
             cell_a1 = table.cell(0, 1).value
             tmp1=0;tmp2=1
        elif int(table.cell(1,0).value)>=0&int(table.cell(1,0).value)<=150:
            cell_a1 = table.cell(1, 0).value
            tmp1=1;tmp2=0
        elif int(table.cell(1,1).value)>=0&int(table.cell(1,1).value)<=150:
            cell_a1 = table.cell(1,1).value
            tmp1=tmp2=1
        else:
            print('Bad Data!\nNo score found!\a')
            exit()
    '''
    print('\tExcel data Loaded')
    temp='0'
    while temp=='0':
        print('Input 1 to get score by number\nInput 2 to sort by score')
        temp=input()
        if temp=='1':   #Get Score By Number
            print('Please input number to get score')
            tmp=input()
            ColList=table.col_values(srtCol-1)
            counter=0
            for i in ColList:
                if str(i).split('.',-1)[0]==str(tmp):
                    break
                counter+=1
            print(table.row_values(1))
            print(table.row_values(counter))
            break


        elif temp=='2':
            temp='2'
        else:
            temp='0'





    print('End')

main()


