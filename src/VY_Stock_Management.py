"""
# Copyright 2020 by Vihangam Yoga Karnataka.
# All rights reserved.
# This file is part of the Vihangan Yoga Operations of Ashram Management Software Package(VYOAM),
# and is released under the "VY License Agreement". Please see the LICENSE
# file that should have been included as part of this package.
# Vihangan Yoga Operations  of Ashram Management Software
# File Name : VY_Stock_Management.py
# Developer : Sant Anurag Deo
# Version : 2.0
"""

"""
# Import all necessary packages both system , and software defined
"""
from app_defines import *
from babel.numbers import *
from init_database import *
from account_statement import *
from split_donation import *
from stock_info import *
from member_donation import *
from loading_animation import *
from pledgeaccount_statement import *
from import_database import *
import ctypes
# import for non commercial data operations
from non_commercialedit import *
from monetarydonation_statement import *
from stocksales_statement import *
from gaushala_account_statement import *

"""
# Class definition starts here
"""


class Library:

    def validate_itemId_Excel(self, itemId, local_centerText):
        bIdExist = False
        print("validate_itemId_Excel--> Start for Item: ", itemId)

        # To open the workbook
        # workbook object is created
        subdir_commercialstock = "..\\Library_Stock\\" + local_centerText.get() + "\\Commercial_Stock"
        filename = subdir_commercialstock + "\\Commercial_Stock.xlsx"
        wb_obj = openpyxl.load_workbook(filename)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(filename)
        print(" Total records for books: ", total_records)
        for iLoop in range(2, total_records + 2):
            cell_obj = sheet_obj.cell(row=iLoop, column=2)
            if cell_obj.value == itemId:
                bIdExist = True
                print("Item is found !!!")
        if not bIdExist:
            print("No id found !!!")
        return bIdExist

    def validate_splitcandidate_id(self, splitId):
        bIdExist = False
        print("validate_splitcandidate_id--> Start for Item: ", splitId)

        # To open the workbook
        # workbook object is created
        filename = self.obj_initDatabase.get_splittransaction_database_name()
        wb_obj = openpyxl.load_workbook(filename)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(filename)
        print(" Total records for split transactions: ", total_records)
        for iLoop in range(2, total_records + 1):
            cell_obj = sheet_obj.cell(row=iLoop, column=10)
            if cell_obj.value == splitId:
                bIdExist = True

        return bIdExist

    def validate_advanceId_Excel(self, itemId):
        bIdExist = False
        print("validate_advanceId_Excel--> Start for Item: ", itemId)

        # To open the workbook
        # workbook object is created
        filename = self.obj_initDatabase.get_advance_database_name()
        wb_obj = openpyxl.load_workbook(filename)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(filename)
        print(" Total records for books: ", total_records)
        for iLoop in range(2, total_records + 1):
            cell_obj = sheet_obj.cell(row=iLoop, column=11)
            if cell_obj.value == itemId:
                bIdExist = True

        return bIdExist

    def prepare_dateFromString(self, dateStr):
        # print("Received str for date conversion : ", dateStr)

        new_date = dateStr.split('-')
        new_Day = new_date[0]
        new_Month = new_date[1]
        new_Year = new_date[2]

        date_final = date(int(new_Year), int(new_Month), int(new_Day))
        return date_final

    def isleapYear(self, year):
        if year % 400 == 0:
            return True
        elif year % 100 == 0:
            return False
        elif year % 4 == 0:
            return True
        else:
            return False

    def calculateNoOfDaysInYear(self, year):
        print("calculateNoOfDaysInYear->Received year:", str(year))
        noOfDays = 365
        new_year = int(year)
        if self.isleapYear(new_year):
            noOfDays = 366
        print("calculateNoOfDaysInYear ---end")
        return noOfDays

    def calculateNoOfDaysInMonth(self, month_name, year):
        if month_name == "January":
            month = 1;
        elif month_name == "February":
            month = 2;
        elif month_name == "March":
            month = 3;
        elif month_name == "April":
            month = 4;
        elif month_name == "May":
            month = 5;
        elif month_name == "June":
            month = 6;
        elif month_name == "July":
            month = 7;
        elif month_name == "August":
            month = 8;
        elif month_name == "September":
            month = 9;
        elif month_name == "October":
            month = 10;
        elif month_name == "November":
            month = 11;
        elif month_name == "December":
            month = 12;
        else:
            print("calculateNoOfDaysInMonth--> Invalid month")

        print("Received year:", str(year))
        new_year = int(year)
        print("calculateNoOfDaysInMonth ---start")
        return monthrange(new_year, month)[1], month

    def getFromAndToDates_Account_Statement(self, month, year, noOfDays):
        fromDate_Month = month
        fromDate_Year = year

        fromDate = date(int(fromDate_Year), int(fromDate_Month), 1)
        toDate = fromDate + timedelta(noOfDays - 1)

        print("getFromAndToDates_Account_Statement : ", fromDate, toDate)
        return fromDate, toDate

    def fetchMonthName(self, monthNumber):
        print("fetchMonthName for :", monthNumber)
        if monthNumber == 1:
            month_name = "January"
        elif monthNumber == 2:
            month_name = "February"
        elif monthNumber == 3:
            month_name = "March"
        elif monthNumber == 4:
            month_name = "April"
        elif monthNumber == 5:
            month_name = "May"
        elif monthNumber == 6:
            month_name = "June"
        elif monthNumber == 7:
            month_name = "July"
        elif monthNumber == 8:
            month_name = "August"
        elif monthNumber == 9:
            month_name = "September"
        elif monthNumber == 10:
            month_name = "October"
        elif monthNumber == 11:
            month_name = "November"
        elif monthNumber == 12:
            month_name = "December"
        else:
            print("Invalid number")
        return month_name

    def calculate_dayDifference(self, borrowDate, returnDate):
        print("borrowDate: ", borrowDate, " returnDate :", returnDate)

        borrow_time = borrowDate.split('-')
        borrowDay = borrow_time[2]
        borrowMonth = borrow_time[1]
        borrowYear = borrow_time[0]

        borrow_date = date(int(borrowYear), int(borrowMonth), int(borrowDay))

        return_time = returnDate.split('-')
        returnDay = return_time[2]
        returnMonth = return_time[1]
        returnYear = return_time[0]

        return_date = date(int(returnYear), int(returnMonth), int(returnDay))

        delta = return_date - borrow_date
        return delta.days

    def calculateTotalBorrowFee_Excel(self, bookName):
        borrow_fee = 0
        print("calculateTotalBorrowFee--> Start for Item: ", bookName)

        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(PATH_STOCK)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(PATH_STOCK)
        print(" calculateTotalBorrowFee_Excel-->Total records for books: ", total_records)
        for iLoop in range(2, total_records + 1):
            cell_obj = sheet_obj.cell(row=iLoop, column=3)
            if cell_obj.value == bookName:
                borrow_fee = sheet_obj.cell(row=iLoop, column=6).value
                break;
        return borrow_fee  # return the borrow fees

    def calculate_TotalBorrowRecord_Excel(self, memberId):
        filename = "..\\Member_Data\\" + memberId + ".xlsx"
        if not os.path.isfile(filename):
            return 0
        print("calculate_TotalBorrowRecord_Excel--> Start for Member: ", memberId)

        # Total borrow records is equal to valid rows in member files
        wb_member = openpyxl.load_workbook(filename)
        wb_sheet = wb_member.active
        borrow_count = 0
        total_records = self.totalrecords_excelDataBase(filename)
        for iLoop in range(1, total_records + 1):
            print("Entering inside loop")
            print("cell value :", wb_sheet.cell(row=iLoop + 1, column=7).value)
            if wb_sheet.cell(row=iLoop + 1, column=7).value is None:
                print("found one eligible")
                borrow_count = borrow_count + 1

        print(" Total borrowed items -->> ", borrow_count)
        return borrow_count, total_records

    def searchStockItemRecords(self, itemId_Text, itemid_labelText,
                               itemname_labelText, authordetails_labelText,
                               unitprice_labelText, quantitydetails_labelText, searchinfo_label, local_centerText):
        print("searchStockItemRecords->>Entry")
        bValidItem = False
        bStockAvailable = False
        # To open the workbook
        # workbook object is created
        subdir_commercialstock = "..\\Library_Stock\\" + local_centerText.get() + "\\Commercial_Stock"
        filename = subdir_commercialstock + "\\Commercial_Stock.xlsx"
        wb_obj = openpyxl.load_workbook(filename)
        print("File for search : ", filename)
        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        itemid = "CI-" + itemId_Text.get()
        total_records = self.totalrecords_excelDataBase(filename)
        print(" Total records for books: ", total_records)
        for iLoop in range(2, total_records + 2):
            cell_book_id = sheet_obj.cell(row=iLoop, column=2)
            if cell_book_id.value == itemid:
                bValidItem = True
                if int(sheet_obj.cell(row=iLoop, column=7).value) > 0:  # if stock quantity  > 0
                    print("Item has been found")
                    itemid_labelText['text'] = sheet_obj.cell(row=iLoop, column=2).value
                    itemname_labelText['text'] = sheet_obj.cell(row=iLoop, column=3).value
                    authordetails_labelText['text'] = sheet_obj.cell(row=iLoop, column=4).value
                    unitprice_labelText['text'] = sheet_obj.cell(row=iLoop, column=5).value
                    quantitydetails_labelText['text'] = str(sheet_obj.cell(row=iLoop, column=7).value)
                    bStockAvailable = True
                    searchinfo_label.configure(text="Item found !!", fg='green')
                else:
                    searchinfo_label.configure(text="No Stock Available !!", fg='red')
                break
            else:
                pass

        if not bValidItem:
            searchinfo_label.configure(text="Item Not Found !!", fg='red')
        print("searchStockItemRecords->>End")

    def searchAdvanceIdExcel(self, itemId_Text,
                             advanceId_labelText,
                             issuedTo_labelText,
                             dateofissue_labelText,
                             issuedName_labelText,
                             issuedAmt_labelText,
                             amtReturnText,
                             description_labelText,
                             searchinfo_label, savebtn):
        print("searchAdvanceIdExcel->>Entry")
        bValidItem = False
        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(self.obj_initDatabase.get_advance_database_name())
        amtReturnText.configure(state=NORMAL)
        amtReturnText.delete(0, END)
        savebtn.configure(state=NORMAL, bg="light cyan")
        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(self.obj_initDatabase.get_advance_database_name())
        print(" Total records for advances is: ", total_records)
        for iLoop in range(1, total_records + 1):
            print("Entering loop")
            cell_book_id = sheet_obj.cell(row=iLoop + 1, column=11)
            print("sheet value :", cell_book_id.value, " Advance Id :", itemId_Text.get())
            if str(cell_book_id.value) == itemId_Text.get():
                bValidItem = True
                print("Item has been found")
                advanceId_labelText['text'] = sheet_obj.cell(row=iLoop + 1, column=11).value
                issuedTo_labelText['text'] = sheet_obj.cell(row=iLoop + 1, column=6).value
                issuedName_labelText['text'] = sheet_obj.cell(row=iLoop + 1, column=7).value
                dateofissue_labelText['text'] = sheet_obj.cell(row=iLoop + 1, column=2).value
                print("Balance :", sheet_obj.cell(row=iLoop + 1, column=4).value)
                if sheet_obj.cell(row=iLoop + 1, column=4).value == "0":
                    amtReturnText.insert(0, "Advance Settled")
                    amtReturnText.configure(state=DISABLED)
                    savebtn.configure(state=DISABLED, bg="light grey")
                issuedAmt_labelText['text'] = sheet_obj.cell(row=iLoop + 1, column=3).value
                description_labelText['text'] = sheet_obj.cell(row=iLoop + 1, column=5).value
                searchinfo_label.configure(text="Advance Id found !!", fg='green')
                break
            else:
                pass
        if not bValidItem:
            searchinfo_label.configure(text="No such Advance issued !!", fg='red')
            self.clearAdvanceReturnForm(advanceId_labelText,
                                        issuedTo_labelText,
                                        dateofissue_labelText,
                                        issuedName_labelText,
                                        issuedAmt_labelText,
                                        amtReturnText,
                                        description_labelText,
                                        searchinfo_label)
        print("searchAdvanceIdExcel->>End")

    def clearAdvanceReturnForm(self, advanceId_labelText,
                               issuedTo_labelText,
                               dateofissue_labelText,
                               issuedName_labelText,
                               issuedAmt_labelText,
                               amtReturnText,
                               description_labelText,
                               searchinfo_label):

        amtReturnText.configure(state=NORMAL)
        amtReturnText.delete(0, END)
        advanceId_labelText['text'] = ""
        issuedTo_labelText['text'] = ""
        dateofissue_labelText['text'] = ""
        issuedName_labelText['text'] = ""
        issuedAmt_labelText['text'] = ""
        description_labelText['text'] = ""
        searchinfo_label['text'] = "Please enter issues Advance Id!!"

    def search_bookInfo_Excel(self, book_id, book_name, bookDetailsLabel, printbtn):
        print("search_bookInfo_Excel->>Entry")
        bValidBookName = True
        bValidItem = True

        # keep the Print button disabled , unless valid record is displayed
        printbtn.configure(state=DISABLED, bg='light grey')
        if book_id.get() == "" and book_name.get() == "":
            bookDetailsLabel['text'] = "\tPlease provide at least book id or name "
        else:
            if book_id.get() != "":
                print("Book id :", book_id.get())
                bValidItem = self.validate_itemId_Excel(book_id.get())
                if bValidItem is False and book_name.get() == "":
                    bookDetailsLabel['text'] = "\t\tInvalid ID, Try again !!"
                print("bValidItem :", bValidItem)
            if book_name.get() != "":
                bValidBookName = self.validate_bookName_Excel(book_name.get())
                if bValidBookName == False and book_id.get() == "":
                    bookDetailsLabel['text'] = "\tBook not found, Try again !!"
                print("bValidBookName :", bValidBookName)
            print("bValidItem :", bValidItem, "bValidBookName :", bValidBookName)
            if bValidItem is False and bValidBookName is False:
                bookDetailsLabel['text'] = "\tNo matching book is found !!"
            elif bValidItem is True or bValidBookName is True:

                # To open the workbook
                # workbook object is created
                wb_obj = openpyxl.load_workbook(PATH_STOCK)

                # Get workbook active sheet object
                # from the active attribute
                sheet_obj = wb_obj.active
                total_records = self.totalrecords_excelDataBase(PATH_STOCK)
                print(" Total records for books: ", total_records)
                for iLoop in range(2, total_records + 1):
                    cell_book_id = sheet_obj.cell(row=iLoop, column=2)
                    cell_book_name = sheet_obj.cell(row=iLoop, column=3)
                    if cell_book_id.value == book_id.get() or cell_book_name.value == book_name.get():
                        print("Condition is true")
                        displayText = "Book Details :" + \
                                      "\n\n\t Name : " + cell_book_name.value + \
                                      "\n\n\t Author : " + sheet_obj.cell(row=iLoop, column=4).value + \
                                      "\n\n\t Book Id :" + cell_book_id.value + \
                                      "\n\n\t Price :" + sheet_obj.cell(row=iLoop, column=5).value + \
                                      "\n\n\t Borrow Fee :" + sheet_obj.cell(row=iLoop, column=6).value + \
                                      "\n\n\t Present Stock :" + str(sheet_obj.cell(row=iLoop, column=7).value)
                        bookDetailsLabel['text'] = displayText

                        # prepare the data for print
                        recordList = []
                        for iLoopList in range(0, 6):
                            recordList.append(sheet_obj.cell(row=iLoop, column=iLoopList + 2).value)
                            print("Book Details [", iLoopList, "] :", recordList[iLoopList])

                        print_result = partial(self.printBookDetailsOnDefaultPrinter, recordList)
                        printbtn.configure(state=NORMAL, bg='light cyan', command=print_result)
                        break
            else:
                pass
        print("search_bookInfo_Excel->>End")

    def printBookDetailsOnDefaultPrinter(self, recordList):
        string_1 = "\n\n\n\n\n\t\t\t---------------------------------------"
        string_2 = "\n\t\t\t\t\tBook Details "
        string_3 = "\n\t\t\t---------------------------------------"
        string_4 = "\n\n\t\t\t\tBook Id  :" + recordList[0]
        string_5 = "\n\n\t\t\t\tBook Name : " + recordList[1]
        string_6 = "\n\n\t\t\t\tAuthor : " + recordList[2]
        string_7 = "\n\n\t\t\t\tPrice :" + recordList[3]
        string_8 = "\n\n\t\t\t\tBorrow Fee  :" + recordList[4]
        string_9 = "\n\n\t\t\t\tStock Quantity : " + recordList[5]
        string_10 = "\n\t\t\t---------------------------------------"
        file_name = "..\\Library_Stock\\" + recordList[0] + ".txt"
        member = open(file_name, 'w')
        member.write(
            string_1 + string_2 + string_3 + string_4 + string_5 + string_6 + string_7 + string_8 + string_9 + string_10)
        os.startfile(file_name)
        member.close()

    def search_borrowRecords_Excel(self, bookReturn_window, member_Id, cal):
        dateTimeObj = cal.get_date()
        dateOfReturn = dateTimeObj.strftime("%Y-%m-%d")
        memberId = member_Id.get()
        self.middleFrame_bookdisplay.destroy()
        self.middleFrame_bookdisplay = Frame(bookReturn_window, name="borrow_detail_frame", width=200, height=100, bd=8,
                                             relief='ridge')
        self.middleFrame_bookdisplay.grid(row=2, column=0, padx=10, pady=10, sticky=W)
        bMemberExists = self.validate_memberlibraryID_Excel(memberId, 1)
        if not bMemberExists:
            messagebox.showwarning("Member Id Error ", "Oops !!! Member Id is not valid  ....")
            return
        else:
            file_name = "..\\Member_Data\\" + memberId + ".xlsx"
            if not os.path.isfile(file_name):
                messagebox.showerror("Nothing Borrowed", "No Borrow record  for this id ....")
                return
            totalBorrowedQuantity, total_records = self.calculate_TotalBorrowRecord_Excel(memberId)

            print("calculate_TotalBorrowRecord_Excel->totalBorrowedQuantity:", totalBorrowedQuantity)
            self.bookId_Dict = {}
            self.label_identities = []
            dict_index = 1
            wb_obj = openpyxl.load_workbook(file_name)
            sheet_obj = wb_obj.active

            label_detail1 = Label(self.middleFrame_bookdisplay, text="Book Id", width=20, height=1, anchor='center',
                                  justify=CENTER,
                                  font=('arial narrow', 12, 'bold'),
                                  bg='light yellow')

            label_detail2 = Label(self.middleFrame_bookdisplay, text="Book Name", width=20, height=1, anchor='center',
                                  justify=CENTER,
                                  font=('arial narrow', 12, 'bold'),
                                  bg='light yellow')

            label_detail3 = Label(self.middleFrame_bookdisplay, text="Date of Borrow", width=20, height=1,
                                  anchor='center',
                                  justify=CENTER,
                                  font=('arial narrow', 12, 'bold'),
                                  bg='light yellow')

            label_detail4 = Label(self.middleFrame_bookdisplay, text="Fee-Incured(Rs.)", width=20, height=1,
                                  anchor='center',
                                  justify=CENTER,
                                  font=('arial narrow', 12, 'bold'),
                                  bg='light yellow')

            label_detail1.grid(row=0, column=1, padx=2, pady=5)
            label_detail2.grid(row=0, column=2, padx=2, pady=5)
            label_detail3.grid(row=0, column=3, padx=2, pady=5)
            label_detail4.grid(row=0, column=4, padx=2, pady=5)

            if totalBorrowedQuantity > 0:
                for row_index in range(0, total_records):
                    if sheet_obj.cell(row=row_index + 2, column=7).value is None:
                        days = self.calculate_dayDifference(sheet_obj.cell(row=row_index + 2, column=6).value,
                                                            dateOfReturn)
                        print("Total borrowed duration :", days)
                        # calculate total borrow fees
                        borrow_fee = 0
                        actual_fee = 0
                        bk_name = sheet_obj.cell(row=row_index + 2, column=5).value
                        borrow_fee = self.calculateTotalBorrowFee_Excel(bk_name)
                        # calculate fine for extra borrowed days
                        if days > 7:
                            fine_amount = (days - 7) * LATE_PAYMENT_FEE
                            actual_fee = 7 * int(borrow_fee)
                        else:
                            fine_amount = 0
                            actual_fee = days * int(borrow_fee)

                        total_fee = int(fine_amount) + actual_fee
                        print("Fine amount = ", fine_amount, " actual_fee = ", actual_fee, " total_fee :", total_fee)
                        book_id = sheet_obj.cell(row=row_index + 2, column=4).value
                        self.bookId_Dict[dict_index] = book_id

                        for column_index in range(1, 5):
                            label_detail = Label(self.middleFrame_bookdisplay, text="No Data", width=20, height=1,
                                                 anchor='center', justify=LEFT,
                                                 font=('arial narrow', 12, 'normal'),
                                                 bg='light yellow')
                            label_detail.grid(row=dict_index, column=column_index, padx=2, pady=5, sticky=W)
                            label_detail['text'] = sheet_obj.cell(row=row_index + 2, column=column_index + 3).value
                            print("Label name :", label_detail)
                            self.label_identities.append(label_detail)
                            if column_index == 4:
                                label_detail['text'] = total_fee
                            if column_index == 2:
                                # underlines all the characters in the label text-*
                                f = font.Font(label_detail, label_detail.cget("font"))
                                f.configure(underline=True)
                                label_detail.configure(font=f)

                                # label text can be clicked and ,and return can be issued from there
                                result = partial(self.return_book_Excel, bookReturn_window, member_Id, 'y',
                                                 dateOfReturn, self.middleFrame_bookdisplay, label_detail)
                                label_detail.bind("<Button-1>", result)

                        dict_index = dict_index + 1
            else:
                messagebox.showwarning("No record found !", " No Borrow record in present")

    def myfunction(self, mycanvas, event):
        mycanvas.configure(scrollregion=mycanvas.bbox("all"), width=770, height=407)

    def print_excel_sheet(self, pathToPrint):
        os.startfile(pathToPrint, 'print')
        print("File is sent for printing to default printer !!!")

    def retrieve_MemberRecords_Excel(self, memberId, memtype, search_criteria):
        print("retrieve_MemberRecords->Start")
        if search_criteria == SEARCH_BY_MEMBERID:
            bMemberExists = self.validate_memberlibraryID_Excel(memberId, memtype)
        elif search_criteria == SEARCH_BY_CONTACTNO:
            bMemberExists = self.validate_contactNo_Excel(memberId, memtype)
        else:
            pass
        print("Member Exists : ", bMemberExists)
        recordList = []
        if not bMemberExists:
            if search_criteria == SEARCH_BY_MEMBERID:
                messagebox.showwarning("Member Id Error ", "Oops !!! Member Id is doesnot exists  ....")
            elif search_criteria == SEARCH_BY_CONTACTNO:
                messagebox.showwarning("Member Id Error ", "Oops !!! Contact Number is not valid  ....")
            else:
                pass
        else:
            if memtype == 1:
                file_name = PATH_MEMBER
            if memtype == 2:
                file_name = PATH_STAFF
            # Fail-safe protection  - if database is deleted anonmously at back end while reaching here
            if not os.path.isfile(file_name):
                messagebox.showerror("Database error", "No Members available ....")
                return
            # To open the workbook
            # workbook object is created
            wb_obj = openpyxl.load_workbook(file_name)
            print(" Data extraction logic will be excuted for : ", memtype, " memberid :", memberId,
                  " and search criteria as:", search_criteria)
            # Get workbook active sheet object
            # from the active attribute
            sheet_obj = wb_obj.active
            total_records = self.totalrecords_excelDataBase(file_name)
            if search_criteria == SEARCH_BY_MEMBERID:
                for iLoop in range(1, total_records + 1):
                    cell_obj = sheet_obj.cell(row=iLoop + 1, column=2)
                    if cell_obj.value == memberId:
                        for iColumn in range(2, MAX_RECORD_ENTRY + 1):
                            cell_value = sheet_obj.cell(row=iLoop + 1, column=iColumn).value
                            print("record[", iColumn, "] :", cell_value)
                            recordList.append(cell_value)
                        break
            elif search_criteria == SEARCH_BY_CONTACTNO:
                print("Executing search criteria :", search_criteria)
                for iLoop in range(1, total_records + 1):
                    cell_obj = sheet_obj.cell(row=iLoop + 1, column=13)
                    if cell_obj.value == memberId:
                        for iColumn in range(2, MAX_RECORD_ENTRY + 1):
                            cell_value = sheet_obj.cell(row=iLoop + 1, column=iColumn).value
                            print("record[", iColumn, "] :", cell_value)
                            recordList.append(cell_value)
                        break
            else:
                print("Please specify search_criteria")

        print("retrieve_MemberRecords->End")
        return recordList

    def retrieve_CommercialItemRecords_Excel(self, itemid, local_centerText):
        print("retrieve_CommercialItemRecords_Excel->Start")
        recordList = []
        subdir = "..\\Config\\Center"
        subdir_commercialstock = "..\\Library_Stock\\" + local_centerText.get() + "\\Commercial_Stock"
        file_name = subdir_commercialstock + "\\Commercial_Stock.xlsx"
        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(file_name)
        print(" Data extraction logic will be executed for : ", itemid)
        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(file_name)

        for iLoop in range(1, total_records + 1):
            cell_obj = sheet_obj.cell(row=iLoop + 1, column=2)
            if cell_obj.value == itemid:
                for iColumn in range(2, 11):
                    cell_value = sheet_obj.cell(row=iLoop + 1, column=iColumn).value
                    print("record[", iColumn, "] :", cell_value)
                    recordList.append(cell_value)
                break
        print("retrieve_CommercialItemRecords_Excel->End")
        return recordList

    def assignDataForDisplay(self, display_dataWindow, memberType, member_Id,
                             memberIdText,
                             memberNameText,
                             memberFatherText,
                             memberMotherText,
                             memberDOBText,
                             memberGenderText,
                             memberContactNoText,
                             memberCityText,
                             memberStateText,
                             memberAddressText,
                             memberNationalityText,
                             memberCountryText,
                             memberPinCodeText,
                             memberIdTypeText,
                             memberEmailText,
                             canvasMem,
                             canvasMemID,
                             search_criteria, print_button):
        print("assignDataForDisplay for :", member_Id, "search criteria :", search_criteria, "member type :",
              memberType)
        member_data = self.retrieve_MemberRecords_Excel(member_Id.get(), memberType, search_criteria)
        if len(member_data) > 0:
            print_button.configure(state=NORMAL, bg='light cyan')
            print_result = partial(self.generate_MemberDetails_Form, member_Id)
            print_button.configure(command=print_result)
            memberIdText['text'] = member_data[0]
            memberNameText['text'] = member_data[2]
            memberFatherText['text'] = member_data[3]
            memberMotherText['text'] = member_data[4]
            memberDOBText['text'] = member_data[5]
            memberGenderText['text'] = member_data[6]
            memberContactNoText['text'] = member_data[11]
            memberCityText['text'] = member_data[8]
            memberStateText['text'] = member_data[9]
            memberAddressText['text'] = member_data[7]
            memberNationalityText['text'] = member_data[13]
            memberCountryText['text'] = member_data[12]
            memberPinCodeText['text'] = member_data[10]
            memberEmailText['text'] = member_data[14]
            strip_idIDTypeText = member_data[17].strip("\n")
            memberIdTypeText['text'] = strip_idIDTypeText
            myPhotoimage = ImageTk.PhotoImage(Image.open(member_data[15]).resize((150, 150)))
            canvasMem.create_image(0, 0, anchor=NW, image=myPhotoimage)
            print("member_data[15] :", member_data[15])

            myIdimage = ImageTk.PhotoImage(Image.open(member_data[16]).resize((150, 150)))
            canvasMemID.create_image(0, 0, anchor=NW, image=myIdimage)
            mainloop()

    def assignDataForDisplay_editMemberInfo(self, display_dataWindow, memberType, member_Id,
                                            memberContactNoText,
                                            memberCityText,
                                            memberStateText,
                                            memberAddressText,
                                            memberCountryText,
                                            memberPinCodeText,
                                            memberEmailText,
                                            professionText,
                                            designation_varaible,
                                            akshyaAvailable_varaible,
                                            akshayboxnoText,
                                            isPatrikaSubsc_varaible,
                                            search_criteria, infoLabel, save_button):
        print("assignDataForDisplay for :", member_Id, "search criteria :", search_criteria, "member type :",
              memberType)
        bMemberIdValid = self.validate_memberlibraryID_Excel(member_Id.get(), 1)
        save_button.configure(bg='light grey', state=DISABLED)
        if bMemberIdValid == True:
            infoLabel.configure(fg='green', text="Please press Save button to modify member records")
            member_data = self.retrieve_MemberRecords_Excel(member_Id.get(), memberType, search_criteria)
            if len(member_data) > 0:
                save_button.configure(state=NORMAL, bg='light cyan')
                memberContactNoText.delete(0, END)
                memberContactNoText.insert(0, member_data[11])
                memberCityText.delete(0, END)
                memberCityText.insert(0, member_data[8])
                memberStateText.delete(0, END)
                memberStateText.insert(0, member_data[9])
                memberCountryText.delete(0, END)
                memberCountryText.insert(0, member_data[12])
                memberAddressText.delete(0, END)
                memberAddressText.insert(0, member_data[7])
                memberPinCodeText.delete(0, END)
                memberPinCodeText.insert(0, member_data[10])
                memberEmailText.delete(0, END)
                memberEmailText.insert(0, member_data[14])

                professionText.delete(0, END)
                professionText.insert(0, member_data[19])
                akshayboxnoText.delete(0, END)
                akshayboxnoText.insert(0, member_data[24])
                designation_varaible.set(member_data[22])
                akshyaAvailable_varaible.set(member_data[23])
                isPatrikaSubsc_varaible.set(member_data[25])

        else:
            infoText = "Invalid Member Id :" + member_Id.get()
            infoLabel.configure(fg='red', text=infoText)
            memberContactNoText.delete(0, END)
            memberContactNoText.configure(fg='black')
            memberCityText.delete(0, END)
            memberCityText.configure(fg='black')
            memberStateText.delete(0, END)
            memberStateText.configure(fg='black')
            memberCountryText.delete(0, END)
            memberCountryText.configure(fg='black')
            memberAddressText.delete(0, END)
            memberAddressText.configure(fg='black')
            memberPinCodeText.delete(0, END)
            memberPinCodeText.configure(fg='black')
            memberEmailText.delete(0, END)
            memberEmailText.configure(fg='black')
            akshayboxnoText.delete(0, END)
            akshayboxnoText.configure(fg='black')
        # mainloop()

    def assignDataForDisplay_editCommercialItemInfo(self, display_dataWindow, item_Id,
                                                    itemname_Text,
                                                    itemauthor_Text,
                                                    unitpriceText,
                                                    itemQuantity_Text,
                                                    borrowFee_Text,
                                                    rackno_Text,
                                                    infoLabel, local_centerText):
        print("assignDataForDisplay_editCommercialItemInfo for :", item_Id.get())
        itemdId_str = "CI-" + item_Id.get()
        bItemValid = self.validate_itemId_Excel(itemdId_str, local_centerText)
        self.print_button.configure(bg='light grey', state=DISABLED)
        if bItemValid:
            infoLabel.configure(fg='green', text="Please press Save button to modify item records")
            item_data = self.retrieve_CommercialItemRecords_Excel(itemdId_str, local_centerText)
            if len(item_data) > 0:
                itemname_Text.delete(0, END)
                itemname_Text.insert(0, item_data[1])
                itemauthor_Text.set(item_data[2])
                unitpriceText.delete(0, END)
                unitpriceText.insert(0, item_data[3])
                borrowFee_Text.delete(0, END)
                borrowFee_Text.insert(0, item_data[4])
                itemQuantity_Text.delete(0, END)
                itemQuantity_Text.insert(0, item_data[5])
                rackno_Text.delete(0, END)
                rackno_Text.insert(0, item_data[6])
                self.print_button.configure(bg='light cyan', state=NORMAL)
        else:
            infoText = "Invalid Item Id :" + itemdId_str
            infoLabel.configure(fg='red', text=infoText)
            itemname_Text.delete(0, END)
            itemname_Text.configure(fg='black')
            itemauthor_Text.set("")
            unitpriceText.delete(0, END)
            borrowFee_Text.delete(0, END)
            rackno_Text.delete(0, END)
            rackno_Text.configure(fg='black')
            itemQuantity_Text.delete(0, END)

    def saveModifiedMemberRecords(self, member_Id,
                                  memberContactNoText,
                                  memberCityText,
                                  memberStateText,
                                  memberAddressText,
                                  memberCountryText,
                                  memberPinCodeText,
                                  memberEmailText,
                                  professionText,
                                  designation_varaible,
                                  akshyaAvailable_varaible,
                                  akshayboxnoText,
                                  isPatrikaSubsc_varaible,
                                  search_criteria, infoLabel):

        wb_obj = openpyxl.load_workbook(PATH_MEMBER)
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(PATH_MEMBER)
        for iLoop in range(1, total_records + 1):
            # if member id matches
            # over-write the respective records
            # save the file
            print("database id ->", sheet_obj.cell(row=iLoop + 1, column=2).value, " member id :", member_Id.get())
            if sheet_obj.cell(row=iLoop + 1, column=2).value == member_Id.get():
                print("condition is true")
                sheet_obj.cell(row=iLoop + 1, column=9).value = memberAddressText.get()
                sheet_obj.cell(row=iLoop + 1, column=11).value = memberStateText.get()
                sheet_obj.cell(row=iLoop + 1, column=12).value = memberPinCodeText.get()
                sheet_obj.cell(row=iLoop + 1, column=10).value = memberCityText.get()
                sheet_obj.cell(row=iLoop + 1, column=13).value = memberContactNoText.get()
                sheet_obj.cell(row=iLoop + 1, column=16).value = memberEmailText.get()
                sheet_obj.cell(row=iLoop + 1, column=14).value = memberCountryText.get()

                sheet_obj.cell(row=iLoop + 1, column=21).value = professionText.get()
                sheet_obj.cell(row=iLoop + 1, column=24).value = designation_varaible.get()
                sheet_obj.cell(row=iLoop + 1, column=25).value = akshyaAvailable_varaible.get()
                sheet_obj.cell(row=iLoop + 1, column=26).value = akshayboxnoText.get()
                sheet_obj.cell(row=iLoop + 1, column=27).value = isPatrikaSubsc_varaible.get()

                wb_obj.save(PATH_MEMBER)

                # check if login already exists for the changed designation if any

                bAccountExists = self.validateStaffAccountExists(member_Id.get())

                if not bAccountExists:
                    # if the designation has been modified to one of below , then check for staff assignment
                    # create the staff login as soon as staff is registered
                    if designation_varaible.get() == "Staff-Sevak" or \
                            designation_varaible.get() == "Manager" or \
                            designation_varaible.get() == "Accountant" or \
                            designation_varaible.get() == "President" or \
                            designation_varaible.get() == "Vice-President":
                        staff_login_path = PATH_STAFF_CREDENTIALS
                        wb = openpyxl.load_workbook(staff_login_path)
                        sheet = wb.active
                        totalRecords = self.totalrecords_excelDataBase(staff_login_path)
                        login_data = {}
                        for iLoop in range(1, 4):
                            if iLoop == 1:
                                if totalRecords == 1:
                                    serial_no = 1
                                else:
                                    serial_no = totalRecords + 1
                                # record serial number is total_rows - 1 , since excluding top header row
                                login_data[1] = serial_no
                            if iLoop == 2:
                                login_data[2] = str(member_Id.get())
                            if iLoop == 3:
                                login_data[3] = "Password@123"
                            sheet.cell(row=totalRecords + 2, column=iLoop).font = Font(size=12, name='Times New Roman',
                                                                                       bold=False)
                            sheet.cell(row=totalRecords + 2, column=iLoop).alignment = Alignment(horizontal='center',
                                                                                                 vertical='center')
                            sheet.cell(row=totalRecords + 2, column=iLoop).value = login_data[iLoop]
                        print("Login has been created for the staff ", str(member_Id.get()))
                        # save the sheet after data is written
                        wb.save(staff_login_path)
                    break

    def saveModifiedCommercialItemRecords(self, display_dataWindow, item_Id,
                                          itemname_Text,
                                          itemauthor_Text,
                                          unitpriceText,
                                          borrowFee_Text,
                                          itemQuantity_Text,
                                          rackno_Text,
                                          infoLabel, local_centerText):
        print("saveModifiedCommercialItemRecords - start")
        subdir_commercialstock = "..\\Library_Stock\\" + local_centerText.get() + "\\Commercial_Stock"
        filename = subdir_commercialstock + "\\Commercial_Stock.xlsx"
        wb_obj = openpyxl.load_workbook(filename)
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(filename)
        itemId = "CI-" + item_Id.get()
        for iLoop in range(1, total_records + 1):
            # if member id matches
            # over-write the respective records
            # save the file
            print("database id ->", sheet_obj.cell(row=iLoop + 1, column=2).value, " item id :", itemId)
            if sheet_obj.cell(row=iLoop + 1, column=2).value == itemId:
                print("condition is true")
                sheet_obj.cell(row=iLoop + 1, column=3).value = itemname_Text.get()
                sheet_obj.cell(row=iLoop + 1, column=4).value = itemauthor_Text.get()
                sheet_obj.cell(row=iLoop + 1, column=5).value = unitpriceText.get()
                sheet_obj.cell(row=iLoop + 1, column=6).value = borrowFee_Text.get()
                sheet_obj.cell(row=iLoop + 1, column=7).value = itemQuantity_Text.get()
                sheet_obj.cell(row=iLoop + 1, column=8).value = rackno_Text.get()
                sheet_obj.cell(row=iLoop + 1, column=10).value = local_centerText.get()

                wb_obj.save(filename)
                infoLabel.configure(fg='green')
                infoLabel['text'] = "Record has been successfully modified for Item Id :" + itemId
                self.print_button.configure(state=DISABLED, bg='light grey')
                break
        print("saveModifiedCommercialItemRecords - end")

    def modifyMemberBorrowRecord_Excel(self, member_Id, bookNo):
        filename = "..\\Member_Data\\" + member_Id + ".xlsx"
        """ Delete a line from a file at the given line number """
        wb_obj = openpyxl.load_workbook(filename)
        sheet_obj = wb_obj.active
        sheet_obj.cell(row=bookNo + 1, column=7).value = Date.today()

        wb_obj.save(filename)

    def return_book_Excel(self, bookReturn_window, member_Id, bReturn, dateOfReturn, middleFrame,
                          label_detail, event):
        gi = label_detail.grid_info()
        x = gi['row']
        y = gi['column']
        print("Dctioanry - ", self.bookId_Dict)
        book_id = self.bookId_Dict[x]
        print("return_book_Excel  bookNo:", book_id, " dateOfReturn :", dateOfReturn)
        if bReturn == 'y':
            wb_obj = openpyxl.load_workbook(PATH_STOCK)
            sheet_obj = wb_obj.active
            total_records = self.totalrecords_excelDataBase(PATH_STOCK)
            line_no = 0
            bookToReturn = ""
            for iLoop in range(1, total_records + 1):
                print("return_book  iteration:", book_id)
                if sheet_obj.cell(row=iLoop + 1, column=2).value == book_id:
                    print("Book found")
                    new_quantity = int(sheet_obj.cell(row=iLoop + 1, column=7).value) + 1
                    sheet_obj.cell(row=iLoop + 1, column=7).value = str(new_quantity)
                    wb_obj.save(PATH_STOCK)
                    filename = "..\\Member_Data\\" + member_Id.get() + ".xlsx"
                    wb_member = openpyxl.load_workbook(filename)
                    member_sheet = wb_member.active
                    total_record = self.totalrecords_excelDataBase(filename)
                    for iLoop in range(1, total_record + 1):
                        if member_sheet.cell(row=iLoop + 1, column=7).value is None and \
                                member_sheet.cell(row=iLoop + 1, column=4).value == book_id:
                            member_sheet.cell(row=iLoop + 1, column=7).font = Font(size=12, name='Times New Roman',
                                                                                   bold=False)
                            member_sheet.cell(row=iLoop + 1, column=7).alignment = Alignment(horizontal='center',
                                                                                             vertical='center')
                            member_sheet.cell(row=iLoop + 1, column=7).value = dateOfReturn
                            wb_member.save(filename)
                            break

                    messagebox.showinfo("Return Success", "Book is returned successfully")

                    # member borrow record shall be modified , if book is returned by user.
                    # self.modifyMemberBorrowRecord(member_Id.get(), bookNo)
                    # removing the file , since no borrow data exists now
                    # refresh the search display
                    print("Refreshing the borrow record for user ")
                    self.search_borrowRecords_Excel(bookReturn_window, member_Id, dateOfReturn)

    def generate_itemId(self, local_centerText):
        # To open the workbook
        # workbook object is created
        subdir_commercialstock = "..\\Library_Stock\\" + local_centerText + "\\Commercial_Stock"
        filename = subdir_commercialstock + "\\Commercial_Stock.xlsx"
        wb_obj = openpyxl.load_workbook(filename)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(filename)
        stock_id = total_records + 100
        return ("CI-" + str(stock_id))  # CI- Commercial Inventory

    def generate_itemId_nonCommenrcial(self, local_centerText):
        # To open the workbook
        # workbook object is created
        subdir_noncommercialstock = "..\\Library_Stock\\" + local_centerText + "\\NonCommercial_Stock"
        filename = subdir_noncommercialstock + "\\noncommercial_stock.xlsx"
        wb_obj = openpyxl.load_workbook(filename)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(filename)
        stock_id = total_records + 100
        return ("NCI-" + str(stock_id))  # CI- Non - Commercial Inventory

    def generate_advanceId(self):
        bid = randint(101, 10000)
        bid_exists = self.validate_advanceId_Excel(bid)
        if bid_exists is True:
            self.generate_advanceId()
        else:
            return bid

    # Generate a unique random receipt number for heavy donations .
    # This will never be shown in actual transactions as it is , hence
    # having a unique random number is OK
    def generate_SplitCandidate_ReceiptNo(self):
        bid = randint(1000, 100000)
        bid_exists = self.validate_splitcandidate_id(bid)
        if bid_exists is True:
            self.generate_SplitCandidate_ReceiptNo()
        else:
            return bid

    def generate_invoiceID(self):
        if os.stat("..\\Library_Stock\\Invoice_list.txt").st_size == 0:
            bid = 999
        else:
            bid = randint(999, 111111)

        return "INV" + str(bid) + "VYB"

    def generate_StockPurchase_invoiceID(self):
        if os.stat("..\\Library_Stock\\Invoice_list.txt").st_size == 0:
            bid = 999
        else:
            bid = randint(999, 111111)

        return "INV" + str(bid) + "STKVYB"

    def generate_invoiceID_sevaDeposit(self):
        if os.stat("..\\Library_Stock\\Invoice_list.txt").st_size == 0:
            bid = 999
        else:
            bid = randint(999, 111111)

        return "INV" + str(bid) + "CRVYB"

    def generate_invoiceID_splitcandidate(self):
        if os.stat("..\\Library_Stock\\Invoice_list.txt").st_size == 0:
            bid = 999
        else:
            bid = randint(999, 111111)

        return "INV" + str(bid) + "CRVYB"

    def generate_invoiceID_magazineSubscription(self):
        if os.stat("..\\Library_Stock\\Invoice_list.txt").st_size == 0:
            bid = 999
        else:
            bid = randint(999, 111111)

        return "INV" + str(bid) + "MAGVYB"

    def generate_invoiceID_nonMonetaryDeposit(self):
        if os.stat("..\\Library_Stock\\Invoice_list.txt").st_size == 0:
            bid = 999
        else:
            bid = randint(999, 111111)

        return "INV" + str(bid) + "NMDVYB"

    def generate_invoiceID_sevaExpanse(self):
        if os.stat("..\\Library_Stock\\Invoice_list.txt").st_size == 0:
            bid = 999
        else:
            bid = randint(999, 111111)

        return "INV" + str(bid) + "DBVYB"

    def validate_bookName_Excel(self, bookName, localCenterName):
        bLibExist = False
        bookId = ""
        print("validate_bookName_Excel--> Start for Book name: ", bookName)
        subdir_commercialstock = "..\\Library_Stock\\" + localCenterName + "\\Commercial_Stock"
        path = subdir_commercialstock + "\\Commercial_Stock.xlsx"

        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(path)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(path)
        print(" Total records for books: ", total_records)
        for iLoop in range(2, total_records + 1):
            cell_obj = sheet_obj.cell(row=iLoop, column=3)
            if cell_obj.value == bookName:
                bLibExist = True
                bookId = sheet_obj.cell(row=iLoop, column=2).value
                break

        return bLibExist

    def validate_bookbyId_Excel(self, bookId):
        bBookExist = False
        bookname = ""
        print("validate_bookName_Excel--> Start for bookId: ", bookId)
        path = PATH_STOCK

        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(path)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(path)
        print(" Total records for books: ", total_records)
        for iLoop in range(2, total_records + 1):
            cell_obj = sheet_obj.cell(row=iLoop, column=2)
            if cell_obj.value == bookId:
                bBookExist = True
                bookname = sheet_obj.cell(row=iLoop, column=3).value
                break

        return bBookExist, bookname

    def validate_nonCommerialItemName_Excel(self, itemName, localCenterName):
        bItemExists = False
        itemId = ""
        print("validate_nonCommerial ItemName_Excel--> Start for Book name: ", itemName)
        subdir_noncommercialstock = "..\\Library_Stock\\" + localCenterName + "\\NonCommercial_Stock"
        path = subdir_noncommercialstock + "\\noncommercial_stock.xlsx"

        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(path)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(path)
        print(" Total records for non commercial items: ", total_records)
        for iLoop in range(2, total_records + 1):
            cell_obj = sheet_obj.cell(row=iLoop, column=3)
            if cell_obj.value == itemName:
                bItemExists = True
                itemId = sheet_obj.cell(row=iLoop, column=2).value
                break

        return bItemExists

    def validate_nonCommerialItemId_Excel(self, itemid, localCenterName):
        bItemExists = False
        print("validate_nonCommerial ItemName_Excel--> Start for Item id: ", itemid)
        subdir_noncommercialstock = "..\\Library_Stock\\" + localCenterName + "\\NonCommercial_Stock"
        path = subdir_noncommercialstock + "\\noncommercial_stock.xlsx"

        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(path)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(path)
        print(" Total records for non commercial items: ", total_records)
        for iLoop in range(2, total_records + 1):
            cell_obj = sheet_obj.cell(row=iLoop, column=2)
            if cell_obj.value == itemid:
                bItemExists = True
                itemId = sheet_obj.cell(row=iLoop, column=2).value
                break

        return bItemExists

    def findMemberName_Excel(self, memberId, memberType):
        member_name = ""
        bLibExist = False
        print("Member id: ", memberId)
        if memberType == 1:
            path = PATH_MEMBER
        elif memberType == 2:
            path = PATH_STAFF
        else:
            pass

        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(path)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(path)
        print(" Total records: ", total_records)
        for iLoop in range(1, total_records + 1):
            cell_obj = sheet_obj.cell(row=iLoop + 1, column=2)
            if cell_obj.value == memberId:
                member_name = sheet_obj.cell(row=iLoop + 1, column=4).value
                break

        return member_name

    def findMemberContactNo_Excel(self, memberId, memberType):
        contact_no = ""
        bLibExist = False
        print("Member id: ", memberId)
        if memberType == 1:
            path = PATH_MEMBER
        elif memberType == 2:
            path = PATH_STAFF
        else:
            pass

        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(path)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.totalrecords_excelDataBase(path)
        print(" Total records: ", total_records)
        for iLoop in range(1, total_records + 1):
            cell_obj = sheet_obj.cell(row=iLoop + 1, column=2)
            if cell_obj.value == memberId:
                contact_no = sheet_obj.cell(row=iLoop + 1, column=13).value
                break

        return contact_no

    def validate_memberlibraryID_Excel(self, memberId, memberType):
        bLibExist = False
        print("Member id: ", memberId)
        if memberType == 1:
            path = PATH_MEMBER
        elif memberType == 2:
            path = PATH_STAFF
        else:
            pass

        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(path)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        totalrecords = self.totalrecords_excelDataBase(path)
        print("total records :", totalrecords)

        for iLoop in range(1, totalrecords + 1):
            print("Entering loop")
            cell_obj = sheet_obj.cell(row=iLoop + 1, column=2)
            print("cell_obj.value : ", cell_obj.value)
            if str(cell_obj.value) == memberId:
                bLibExist = True
                break
        return bLibExist

    def validate_advanceIdExcel(self, advanceId):
        bLibExist = False
        path = PATH_ADVANCE_SHEET
        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(path)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        totalrecords = self.totalrecords_excelDataBase(path)
        print("total records :", totalrecords)

        for iLoop in range(1, totalrecords + 1):
            print("Entering loop")
            cell_obj = sheet_obj.cell(row=iLoop + 1, column=11)
            print("cell_obj.value : ", cell_obj.value)
            if cell_obj.value == advanceId:
                bLibExist = True
                break
        return bLibExist

    def validate_memberSubscriptionCurrentYear(self, memberId, magazineCategoryText, requested_subscription_date):
        subscription_exists = False
        path = self.obj_initDatabase.get_magazine_subscription_database_name()

        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(path)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        totalrecords = self.totalrecords_excelDataBase(path)

        for iLoop in range(1, totalrecords + 1):
            if memberId == str(sheet_obj.cell(row=iLoop + 1, column=2).value) and \
                    magazineCategoryText.get() == str(sheet_obj.cell(row=iLoop + 1, column=5).value):
                cell_obj_date = sheet_obj.cell(row=iLoop + 1, column=4)
                print("cell_obj Date : ", cell_obj_date.value)
                subscription_date_from_database = self.prepare_dateFromString(cell_obj_date.value)
                subscription_year_from_database = subscription_date_from_database.year
                print("subscription_year_from_database :", subscription_year_from_database, "requested year:",
                      requested_subscription_date.year)
                if subscription_year_from_database == requested_subscription_date.year:
                    subscription_exists = True
                    print("Subscription Already exists !!!")
                    break
        return subscription_exists

    def validate_contactNo_Excel(self, memberId, memberType):
        bLibExist = False
        print("Member id: ", memberId)
        if memberType == 1:
            path = PATH_MEMBER
        elif memberType == 2:
            path = PATH_STAFF
        else:
            pass

        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(path)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        totalrecords = self.totalrecords_excelDataBase(path)
        print("total records :", totalrecords)

        for iLoop in range(1, totalrecords + 1):
            print("Entering loop")
            cell_obj = sheet_obj.cell(row=iLoop + 1, column=13)
            print("cell_obj.value : ", cell_obj.value)
            if cell_obj.value == memberId:
                bLibExist = True
                break
        return bLibExist

    def validate_memberGovtID_Excel(self, memberId):
        path = "..\\Member_Data\\Member.xlsx"
        bMemberExist = False
        if os.stat(path).st_size == 0:
            return False
        else:
            # To open the workbook
            # workbook object is created
            wb_obj = openpyxl.load_workbook(path)

            # Get workbook active sheet object
            # from the active attribute
            sheet_obj = wb_obj.active
            totalrecords = self.totalrecords_excelDataBase(path)
            for iLoop in range(2, totalrecords + 1):
                cell_obj = sheet_obj.cell(row=iLoop, column=3)
                if cell_obj.value == memberId:
                    bMemberExist = True
                    break;
        return bMemberExist

    # Function for clearing the
    # contents of text entry boxes
    def clear_form(self, name, author, price, quantity, borrowFee):
        # clear the content of text entry box
        name.delete(0, END)
        name.configure(fg='black')
        # author.delete(0, END)
        # author.configure(fg='black')
        price.delete(0, END)
        price.configure(fg='black')
        quantity.delete(0, END)
        quantity.configure(fg='black')
        borrowFee.delete(0, END)
        borrowFee.configure(fg='black')

    def clear_Memberform(self, member_govtId,
                         member_name,
                         member_fatherName,
                         member_mother,
                         member_gender,
                         member_address,
                         member_city,
                         member_state,
                         member_pincode,
                         member_contactNo,
                         member_country,
                         member_nationality,
                         member_emailId):
        # clear the content of text entry box
        member_govtId.delete(0, END)
        member_govtId.configure(fg='black')
        member_name.delete(0, END)
        member_name.configure(fg='black')
        member_fatherName.delete(0, END)
        member_fatherName.configure(fg='black')
        member_mother.delete(0, END)
        member_mother.configure(fg='black')
        member_dob = ""
        # member_gender
        member_address.delete('0', END)
        member_address.configure(fg='black')
        member_city.delete(0, END)
        member_city.configure(fg='black')
        member_state.configure(fg='black')
        member_state.delete(0, END)
        member_pincode.delete(0, END)
        member_pincode.configure(fg='black')
        member_country.delete(0, END)
        member_country.configure(fg='black')
        member_nationality.delete(0, END)
        member_nationality.configure(fg='black')
        member_contactNo.delete(0, END)
        member_contactNo.configure(fg='black')
        member_emailId.delete(0, END)
        member_emailId.configure(fg='black')

    def insert_commercial_data_Excel(self, newItem_window, item_name, author_name, item_price, item_borrowfee,
                                     item_quantity,
                                     rack_location, cal, local_centerText):

        dateTimeObj = cal.get_date()
        receival_date = dateTimeObj.strftime("%Y-%m-%d")
        item_id = self.generate_itemId(local_centerText.get())  # generates a unique item id
        if item_name.get() == "" or author_name.get() == "" or item_price.get() == "" or item_quantity.get() == "" or item_borrowfee.get() == "":
            messagebox.showinfo("Data Entry Error", "All fields are mandatory !!!")

        else:
            bBookExists = self.validate_bookName_Excel(item_name.get(), local_centerText.get())
            print("bBookExists :", bBookExists)
            if bBookExists:
                messagebox.showwarning("Duplicate Entry Error !", "Book already exists !!")
                item_name.configure(bd=2, fg='red')
                return
            else:
                subdir_commercialstock = "..\\Library_Stock\\" + local_centerText.get() + "\\Commercial_Stock"
                path = subdir_commercialstock + "\\Commercial_Stock.xlsx"
                wb_obj = openpyxl.load_workbook(path)
                sheet_obj = wb_obj.active
                total_records = self.totalrecords_excelDataBase(path)
                serial_no = 0
                row_no = 0
                if total_records is 0:
                    serial_no = 1
                    row_no = 2
                else:
                    serial_no = total_records + 1
                    row_no = total_records + 2

                # sets the general formatting for the new entry in new row
                for index in range(1, 9):
                    sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                         bold=False)
                    sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                   vertical='center')

                # new book data is assigned to respective cells in row
                sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                sheet_obj.cell(row=row_no, column=2).value = str(item_id)
                sheet_obj.cell(row=row_no, column=3).value = str(item_name.get())
                sheet_obj.cell(row=row_no, column=4).value = str(author_name.get())
                sheet_obj.cell(row=row_no, column=5).value = str(item_price.get())
                sheet_obj.cell(row=row_no, column=6).value = str(item_borrowfee.get())
                sheet_obj.cell(row=row_no, column=7).value = str(item_quantity.get())
                sheet_obj.cell(row=row_no, column=8).value = str(rack_location.get())
                sheet_obj.cell(row=row_no, column=9).value = str(receival_date)
                sheet_obj.cell(row=row_no, column=10).value = str(local_centerText.get())
                wb_obj.save(path)
                self.submit.configure(state=DISABLED, bg='light grey')
                self.clear_form(item_name, author_name, item_price, item_borrowfee, item_quantity)
                user_choice = messagebox.askquestion("Item insertion success", "Do you want to add another item ? ")
                # destroy the data entry form , if user do not want to add more records
                if user_choice == 'no':
                    newItem_window.destroy()

    def deposit_seva_rashi_Excel(self, new_noncommercial_Item_window,
                                 donator_idText,
                                 seva_amountText,
                                 categoryText,
                                 collector_nameText,
                                 cal,
                                 paymentMode_menu,
                                 paymentMode_text,
                                 authorizedby_Text,
                                 akshayPatra_Text,
                                 invoice_idText,
                                 infolabel, print_invoice, transId_Text):
        dateTimeObj = cal.get_date()
        dateOfCollection_calc = dateTimeObj.strftime("%Y-%m-%d ")
        if donator_idText.get() == "" or \
                seva_amountText.get() == "" or \
                collector_nameText.get() == "":
            infolabel.configure(text="All fields are mandatory", fg='red')

        else:
            today = date.today()
            if dateTimeObj <= today:
                bDonatorIdValid = self.validate_memberlibraryID_Excel(donator_idText.get(), 1)
                bReceiverIdValid = self.validate_memberlibraryID_Excel(collector_nameText.get(), 1)
                bAuthorizorIdValid = self.validate_memberlibraryID_Excel(authorizedby_Text.get(), 1)
                if bDonatorIdValid and bReceiverIdValid and bAuthorizorIdValid and (seva_amountText.get()).isnumeric():
                    if categoryText.get() == "Akshay-Patra Seva" and akshayPatra_Text.get() == "Not Available":
                        infolabel.configure(text="No Akshay patra Assigned to this member !!!", fg='red')
                    else:
                        member_data = self.retrieve_MemberRecords_Excel(donator_idText.get(), 1, SEARCH_BY_MEMBERID)
                        revceiver_data = self.retrieve_MemberRecords_Excel(collector_nameText.get(), 1,
                                                                           SEARCH_BY_MEMBERID)
                        authorizor_data = self.retrieve_MemberRecords_Excel(authorizedby_Text.get(), 1,
                                                                            SEARCH_BY_MEMBERID)
                        print("For debugging Seva Amt :", seva_amountText.get(), "Max donation allowed :",
                              MAX_ALLOWED_DONATION)
                        bGnericSevaCase = True

                        # If the amount > 10000 and by Cash medium,same is considered to be a split candidate
                        if (paymentMode_text.get() == "Cash") and (int(seva_amountText.get()) > MAX_ALLOWED_DONATION):
                            bGnericSevaCase = False

                        # if not a split candidate, proceed with normal deposit to accounts
                        if bGnericSevaCase:

                            filename_MonetarySheet = self.obj_initDatabase.get_seva_deposit_database_name()  # PATH_SEVA_SHEET

                            # writting the credit in Master Seva Sheet - starts
                            if categoryText.get() == "Gaushala Seva":
                                invoice_id = self.obj_commonUtil.generateMonetaryDonationReceiptId(
                                    SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST)
                            else:
                                invoice_id = self.obj_commonUtil.generateMonetaryDonationReceiptId(
                                    VIHANGAM_YOGA_KARNATAKA_TRUST)
                            # open seva rashi sheet and enter the data --start
                            wb_obj = openpyxl.load_workbook(filename_MonetarySheet)
                            sheet_obj = wb_obj.active
                            total_records = self.totalrecords_excelDataBase(filename_MonetarySheet)

                            if total_records is 0:
                                serial_no = 1
                                row_no = 2
                                balance_amount = 0
                            else:
                                serial_no = total_records + 1
                                row_no = total_records + 2
                                balance_amount = int(str(sheet_obj.cell(row=row_no - 1, column=3).value))

                            # sets the general formatting for the new entry in new row
                            for index in range(1, 14):
                                sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                                     bold=False)
                                sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                               vertical='center')

                            # new book data is assigned to respective cells in row
                            sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                            sheet_obj.cell(row=row_no, column=2).value = str(seva_amountText.get())
                            sheet_obj.cell(row=row_no, column=3).value = str(
                                balance_amount + int(seva_amountText.get()))
                            sheet_obj.cell(row=row_no, column=4).value = str(donator_idText.get())
                            sheet_obj.cell(row=row_no, column=5).value = str(member_data[2])
                            sheet_obj.cell(row=row_no, column=6).value = str(dateOfCollection_calc)
                            sheet_obj.cell(row=row_no, column=7).value = str(categoryText.get())
                            sheet_obj.cell(row=row_no, column=8).value = str(collector_nameText.get())
                            sheet_obj.cell(row=row_no, column=9).value = str(revceiver_data[2])
                            sheet_obj.cell(row=row_no, column=10).value = str(member_data[7])  # address
                            if paymentMode_text.get() == "Bank Transfer":
                                paymenttext = str(paymentMode_text.get()) + "Ref. " + str(transId_Text.get())
                            else:
                                paymenttext = str(paymentMode_text.get())
                            sheet_obj.cell(row=row_no, column=11).value = paymenttext
                            sheet_obj.cell(row=row_no, column=12).value = str(authorizedby_Text.get())
                            sheet_obj.cell(row=row_no, column=13).value = str(authorizor_data[2])
                            sheet_obj.cell(row=row_no, column=14).value = str(invoice_id)

                            wb_obj.save(filename_MonetarySheet)

                            # writting the credit in Master Seva Sheet - starts
                            bAlike_Seva = True
                            trust_type = VIHANGAM_YOGA_KARNATAKA_TRUST
                            if categoryText.get() == "Monthly Seva":
                                path_seva_sheet = self.obj_initDatabase.get_monthly_seva_database_name()  # PATH_MONTHLY_SEVA_SHEET
                            elif categoryText.get() == "Gaushala Seva":
                                trust_type = SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST
                                path_seva_sheet = self.obj_initDatabase.get_gaushala_seva_database_name()  # PATH_GAUSHALA_SEVA_SHEET
                            elif categoryText.get() == "Hawan Seva":
                                path_seva_sheet = self.obj_initDatabase.get_hawan_seva_database_name()  # PATH_HAWAN_SEVA_SHEET
                            elif categoryText.get() == "Event/Prachar Seva":
                                path_seva_sheet = self.obj_initDatabase.get_prachar_event_seva_database_name()  # PATH_EVENT_SEVA_SHEET
                            elif categoryText.get() == "Aarti Seva":
                                path_seva_sheet = self.obj_initDatabase.get_aarti_seva_database_name()  # PATH_AARTI_SEVA_SHEET
                            elif categoryText.get() == "Ashram Seva(Generic)":
                                path_seva_sheet = self.obj_initDatabase.get_ashram_seva_database_name()  # PATH_ASHRAM_GENERIC_SEVA_SHEET
                            elif categoryText.get() == "Ashram Nirmaan Seva":
                                path_seva_sheet = self.obj_initDatabase.get_ashram_nirmaan_seva_database_name()  # PATH_ASHRAM_NIRMAAN_SHEET
                            elif categoryText.get() == "Yoga Fees":
                                path_seva_sheet = self.obj_initDatabase.get_yoga_seva_database_name()  # PATH_YOGA_FEES_SHEET
                            elif categoryText.get() == "Akshay-Patra Seva":
                                path_seva_sheet = self.obj_initDatabase.get_akshay_patra_database_name()  # PATH_AKSHAY_PATRA_DATABASE
                            else:
                                bAlike_Seva = False
                                pass

                            # writting into respective seva sheet - start
                            if bAlike_Seva:
                                wb_sevasheetobj = openpyxl.load_workbook(path_seva_sheet)
                                sevasheet_obj = wb_sevasheetobj.active
                                total_records_seva_sheet = self.totalrecords_excelDataBase(path_seva_sheet)

                                if total_records_seva_sheet is 0:
                                    serial_no = 1
                                    row_no = 2
                                    balance_amount = 0
                                else:
                                    serial_no = total_records_seva_sheet + 1
                                    row_no = total_records_seva_sheet + 2
                                    balance_amount = int(str(sevasheet_obj.cell(row=row_no - 1, column=3).value))

                                # sets the general formatting for the new entry in new row
                                for index in range(1, 14):
                                    sevasheet_obj.cell(row=row_no, column=index).font = Font(size=12,
                                                                                             name='Times New Roman',
                                                                                             bold=False)
                                    sevasheet_obj.cell(row=row_no, column=index).alignment = Alignment(
                                        horizontal='center',
                                        vertical='center')

                                # new book data is assigned to respective cells in row
                                sevasheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                                sevasheet_obj.cell(row=row_no, column=2).value = str(seva_amountText.get())
                                sevasheet_obj.cell(row=row_no, column=3).value = str(
                                    balance_amount + int(seva_amountText.get()))
                                sevasheet_obj.cell(row=row_no, column=4).value = str(donator_idText.get())
                                sevasheet_obj.cell(row=row_no, column=5).value = str(member_data[2])
                                sevasheet_obj.cell(row=row_no, column=6).value = str(dateOfCollection_calc)
                                sevasheet_obj.cell(row=row_no, column=7).value = str(categoryText.get())
                                sevasheet_obj.cell(row=row_no, column=8).value = str(collector_nameText.get())
                                sevasheet_obj.cell(row=row_no, column=9).value = str(revceiver_data[2])
                                if categoryText.get() == "Akshay-Patra Seva":
                                    sevasheet_obj.cell(row=row_no, column=10).value = str(
                                        akshayPatra_Text.get())  # address
                                else:
                                    sevasheet_obj.cell(row=row_no, column=10).value = str(member_data[7])  # address
                                if paymentMode_text.get() == "Bank Transfer":
                                    paymenttext = str(paymentMode_text.get()) + "Ref. " + str(transId_Text.get())
                                else:
                                    paymenttext = str(paymentMode_text.get())
                                sevasheet_obj.cell(row=row_no, column=11).value = paymenttext
                                sevasheet_obj.cell(row=row_no, column=12).value = str(authorizedby_Text.get())
                                sevasheet_obj.cell(row=row_no, column=13).value = str(authorizor_data[2])
                                sevasheet_obj.cell(row=row_no, column=14).value = str(invoice_id)

                                wb_sevasheetobj.save(path_seva_sheet)

                                # writting into respective seva sheet - end

                            # open transaction sheet and enter the data
                            # receiving donation is a credit transaction for the organization
                            if categoryText.get() == "Gaushala Seva":
                                file_name_transaction = self.obj_initDatabase.get_gaushala_transaction_database_name()  # PATH_TRANSACTION_SHEET
                            else:
                                file_name_transaction = self.obj_initDatabase.get_transaction_database_name()  # PATH_TRANSACTION_SHEET

                            transaction_wb_obj = openpyxl.load_workbook(file_name_transaction)
                            transaction_sheet_obj = transaction_wb_obj.active
                            total_records_transaction = self.totalrecords_excelDataBase(file_name_transaction)

                            if total_records_transaction is 0:
                                serial_no = 1
                                row_no = 2
                                balance_amount = 0
                            else:
                                serial_no = total_records_transaction + 1
                                row_no = total_records_transaction + 2
                                balance_amount = int(str(transaction_sheet_obj.cell(row=row_no - 1, column=9).value))

                            # sets the general formatting for the new entry in new row
                            for index in range(1, 10):
                                transaction_sheet_obj.cell(row=row_no, column=index).font = Font(size=12,
                                                                                                 name='Times New Roman',
                                                                                                 bold=False)
                                transaction_sheet_obj.cell(row=row_no, column=index).alignment = Alignment(
                                    horizontal='center',
                                    vertical='center')

                            # new book data is assigned to respective cells in row
                            transaction_sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                            transaction_sheet_obj.cell(row=row_no, column=2).value = str(dateOfCollection_calc)
                            transaction_sheet_obj.cell(row=row_no, column=3).value = str(seva_amountText.get())
                            transaction_sheet_obj.cell(row=row_no, column=4).value = "---"
                            transaction_sheet_obj.cell(row=row_no, column=5).value = str(categoryText.get())
                            if paymentMode_text.get() == "Bank Transfer":
                                paymenttext = str(paymentMode_text.get()) + "Ref. " + str(transId_Text.get())
                            else:
                                paymenttext = str(paymentMode_text.get())
                            transaction_sheet_obj.cell(row=row_no, column=6).value = paymenttext
                            transaction_sheet_obj.cell(row=row_no, column=7).value = str(
                                authorizedby_Text.get())  # authorizor id
                            transaction_sheet_obj.cell(row=row_no, column=8).value = str(authorizor_data[2])
                            transaction_sheet_obj.cell(row=row_no, column=9).value = str(
                                balance_amount + int(seva_amountText.get()))
                            transaction_sheet_obj.cell(row=row_no, column=10).value = str(invoice_id)
                            transaction_wb_obj.save(file_name_transaction)

                            invoice_idText['text'] = invoice_id

                            # open transaction sheet and enter the data --end

                            self.generateDonationReceipt(donator_idText,
                                                         seva_amountText,
                                                         categoryText,
                                                         revceiver_data[2],
                                                         dateOfCollection_calc,
                                                         paymentMode_text,
                                                         invoice_id, member_data, print_invoice, self.submit_deposit)

                            text_withID = "Seva deposited successfully. Invoice  id :" + invoice_id
                            infolabel.configure(text=text_withID, fg='green')
                            # update the total balance
                            self.obj_commonUtil.calculateTotalAvailableBalance(trust_type)
                        else:
                            print("Donation amount is by CASH and greater than ", MAX_ALLOWED_DONATION)
                            filename_splitCandidate = self.obj_initDatabase.get_splittransaction_database_name()

                            # writting the credit in Master Seva Sheet - starts
                            invoice_id = self.generate_SplitCandidate_ReceiptNo()
                            # open seva rashi sheet and enter the data --start
                            wb_obj = openpyxl.load_workbook(filename_splitCandidate)
                            sheet_obj = wb_obj.active
                            total_records = self.totalrecords_excelDataBase(filename_splitCandidate)

                            if total_records is 0:
                                serial_no = 1
                                row_no = 2
                                balance_amount = 0
                            else:
                                serial_no = total_records + 1
                                row_no = total_records + 2
                                balance_amount = int(str(sheet_obj.cell(row=row_no - 1, column=3).value))

                            # sets the general formatting for the new entry in new row
                            for index in range(1, 11):
                                sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                                     bold=False)
                                sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                               vertical='center')

                            # new book data is assigned to respective cells in row
                            sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                            sheet_obj.cell(row=row_no, column=2).value = str(dateOfCollection_calc)
                            sheet_obj.cell(row=row_no, column=3).value = str(seva_amountText.get())
                            sheet_obj.cell(row=row_no, column=4).value = "Open"
                            sheet_obj.cell(row=row_no, column=5).value = str(categoryText.get())
                            sheet_obj.cell(row=row_no, column=6).value = str(paymentMode_text.get())
                            sheet_obj.cell(row=row_no, column=7).value = str(donator_idText.get())
                            sheet_obj.cell(row=row_no, column=8).value = str(member_data[2])
                            sheet_obj.cell(row=row_no, column=9).value = str(seva_amountText.get())
                            sheet_obj.cell(row=row_no, column=10).value = str(invoice_id)
                            wb_obj.save(filename_splitCandidate)

                            self.generateDonationReceipt(donator_idText,
                                                         seva_amountText,
                                                         categoryText,
                                                         revceiver_data[2],
                                                         dateOfCollection_calc,
                                                         paymentMode_text,
                                                         str(invoice_id), member_data, print_invoice,
                                                         self.submit_deposit)

                            text_withID = "Special Donation saved successfully. Receipt  id :" + str(invoice_id)
                            infolabel.configure(text=text_withID, fg='green')
                else:
                    infolabel.configure(text="Invalid Data for ID/IDs/Amount, please correct ...", fg='red')
            else:
                infolabel.configure(text="Transaction Date cannot be future!! please correct ...", fg='red')

    def update_pledge_detail(self, new_sankalp_window,
                             donator_idText,
                             pleadge_amountText,
                             trust_nametext,
                             coordinator_text,
                             cal,
                             duration_text,
                             paymentduration_text,
                             pledge_itemtext, remBalance_text,
                             infolabel):
        dateTimeObj = cal.get_date()
        dateOfCollection_calc = dateTimeObj.strftime("%Y-%m-%d ")
        if donator_idText.get() == "" or \
                pleadge_amountText.get() == "" or \
                coordinator_text.get() == "" or \
                duration_text.get() == "":
            infolabel.configure(text="All fields are mandatory", fg='red')

        else:
            today = date.today()
            if dateTimeObj <= today:
                bDonatorIdValid = self.validate_memberlibraryID_Excel(donator_idText.get(), 1)
                bReceiverIdValid = self.validate_memberlibraryID_Excel(coordinator_text.get(), 1)

                if bDonatorIdValid and bReceiverIdValid and (pleadge_amountText.get()).isnumeric():
                    member_data = self.retrieve_MemberRecords_Excel(donator_idText.get(), 1, SEARCH_BY_MEMBERID)
                    revceiver_data = self.retrieve_MemberRecords_Excel(coordinator_text.get(), 1,
                                                                       SEARCH_BY_MEMBERID)
                    pledge_file_path = self.obj_initDatabase.initilize_pledgeitem_database(trust_nametext,
                                                                                           pledge_itemtext.get())
                    pledge_master_file_path = self.obj_initDatabase.initilize_pledgeitem_database(trust_nametext,
                                                                                                  "Pledge_Master")

                    # open seva rashi sheet and enter the data --start
                    wb_obj = openpyxl.load_workbook(pledge_file_path)
                    sheet_obj = wb_obj.active
                    total_records = self.totalrecords_excelDataBase(pledge_file_path)

                    if total_records is 0:
                        serial_no = 1
                        row_no = 2

                    else:
                        serial_no = total_records + 1
                        row_no = total_records + 2

                    # sets the general formatting for the new entry in new row
                    for index in range(1, 14):
                        sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                             bold=False)
                        sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                       vertical='center')

                    # new book data is assigned to respective cells in row
                    sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                    sheet_obj.cell(row=row_no, column=2).value = str(donator_idText.get())
                    sheet_obj.cell(row=row_no, column=3).value = str(member_data[2])
                    sheet_obj.cell(row=row_no, column=4).value = str(member_data[11])
                    sheet_obj.cell(row=row_no, column=5).value = str(pleadge_amountText.get())
                    sheet_obj.cell(row=row_no, column=6).value = str(pledge_itemtext.get())
                    sheet_obj.cell(row=row_no, column=7).value = str(dateOfCollection_calc)
                    sheet_obj.cell(row=row_no, column=8).value = str(coordinator_text.get())
                    sheet_obj.cell(row=row_no, column=9).value = str(revceiver_data[2])
                    sheet_obj.cell(row=row_no, column=10).value = str(duration_text.get())  # address
                    sheet_obj.cell(row=row_no, column=11).value = str(paymentduration_text.get())
                    sheet_obj.cell(row=row_no, column=12).value = str(pleadge_amountText.get())
                    sheet_obj.cell(row=row_no, column=13).value = "Open"

                    remBalance_text['text'] = "Rs." + pleadge_amountText.get()

                    wb_obj.save(pledge_file_path)

                    # open pledge master and write the record there too
                    wb_pledgeMaster = openpyxl.load_workbook(pledge_master_file_path)
                    sheet_pledgeMaster = wb_pledgeMaster.active
                    total_records_pledge_master = self.totalrecords_excelDataBase(pledge_master_file_path)

                    if total_records_pledge_master is 0:
                        serial_no_pledge_master = 1
                        row_no_pledge_master = 2

                    else:
                        serial_no_pledge_master = total_records_pledge_master + 1
                        row_no_pledge_master = total_records_pledge_master + 2

                    # sets the general formatting for the new entry in new row
                    for index in range(1, 14):
                        sheet_pledgeMaster.cell(row=row_no_pledge_master, column=index).font = Font(size=12,
                                                                                                    name='Times New Roman',
                                                                                                    bold=False)
                        sheet_pledgeMaster.cell(row=row_no_pledge_master, column=index).alignment = Alignment(
                            horizontal='center',
                            vertical='center')

                    # new book data is assigned to respective cells in row
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=1).value = str(serial_no_pledge_master)
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=2).value = str(donator_idText.get())
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=3).value = str(member_data[2])
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=4).value = str(member_data[11])
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=5).value = str(pleadge_amountText.get())
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=6).value = str(pledge_itemtext.get())
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=7).value = str(dateOfCollection_calc)
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=8).value = str(coordinator_text.get())
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=9).value = str(revceiver_data[2])
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=10).value = str(
                        duration_text.get())  # address
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=11).value = str(paymentduration_text.get())
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=12).value = str(pleadge_amountText.get())
                    sheet_pledgeMaster.cell(row=row_no_pledge_master, column=13).value = "Open"
                    wb_pledgeMaster.save(pledge_master_file_path)
                    print("Pledge master writing complete for : ", trust_nametext.get())

                    text_withID = "Congratulation!!! Pledge Registered  for :", str(donator_idText.get())
                    infolabel.configure(text=text_withID, fg='green')
                else:
                    infolabel.configure(text="Invalid Data for ID/IDs/Amount, please correct ...", fg='red')
            else:
                infolabel.configure(text="Pledge payment Date cannot be future!! please correct ...", fg='red')

    def update_pledge_payment(self, new_noncommercial_Item_window,
                              donator_idText,
                              seva_amountText,
                              categoryText,
                              collector_nameText,
                              cal,
                              trust_nametext,
                              paymentMode_menu,
                              paymentMode_text,
                              authorizedby_Text,
                              invoice_idText,
                              infolabel, print_invoice, transId_Text):
        dateTimeObj = cal.get_date()
        dateOfCollection_calc = dateTimeObj.strftime("%Y-%m-%d ")
        if donator_idText.get() == "" or \
                seva_amountText.get() == "" or \
                collector_nameText.get() == "":
            infolabel.configure(text="All fields are mandatory", fg='red')

        else:
            today = date.today()
            if dateTimeObj <= today:
                bDonatorIdValid = self.validate_memberlibraryID_Excel(donator_idText.get(), 1)
                bReceiverIdValid = self.validate_memberlibraryID_Excel(collector_nameText.get(), 1)
                bAuthorizorIdValid = self.validate_memberlibraryID_Excel(authorizedby_Text.get(), 1)
                if bDonatorIdValid and bReceiverIdValid and bAuthorizorIdValid and (seva_amountText.get()).isnumeric():
                    member_data = self.retrieve_MemberRecords_Excel(donator_idText.get(), 1, SEARCH_BY_MEMBERID)
                    revceiver_data = self.retrieve_MemberRecords_Excel(collector_nameText.get(), 1,
                                                                       SEARCH_BY_MEMBERID)
                    authorizor_data = self.retrieve_MemberRecords_Excel(authorizedby_Text.get(), 1,
                                                                        SEARCH_BY_MEMBERID)
                    # open the pledge item master sheet and change the remaining balance item after this payment
                    pledge_file_path = self.obj_initDatabase.initilize_pledgeitem_database(trust_nametext,
                                                                                           categoryText.get())
                    rem_balance = 0
                    print("pledge_file_path :", pledge_file_path)
                    pledge_wb_obj = openpyxl.load_workbook(pledge_file_path)
                    pledge_sheet_obj = pledge_wb_obj.active
                    total_records = self.totalrecords_excelDataBase(pledge_file_path)
                    for iLoop in range(1, total_records + 1):
                        cell_obj = pledge_sheet_obj.cell(row=iLoop + 1, column=2)
                        if cell_obj.value == str(donator_idText.get()):
                            if pledge_sheet_obj.cell(row=iLoop + 1, column=13).value == "Open":
                                print("Pledge exists and currently open")
                                rem_balance = int(pledge_sheet_obj.cell(row=iLoop + 1, column=12).value) - int(
                                    str(seva_amountText.get()))
                                pledge_sheet_obj.cell(row=iLoop + 1, column=12).value = str(rem_balance)
                                if rem_balance == 0:
                                    pledge_sheet_obj.cell(row=iLoop + 1, column=13).value = "Closed"
                                pledge_wb_obj.save(pledge_file_path)
                                # open the pledge master sheet and change the remaining balance item after this payment
                                pledge_master_file_path = self.obj_initDatabase.initilize_pledgeitem_database(
                                    trust_nametext,
                                    "Pledge_Master")
                                print("pledge_master_file_path :", pledge_master_file_path)
                                rem_balance = 0
                                pledge_master_wb_obj = openpyxl.load_workbook(pledge_master_file_path)
                                pledge_master_sheet_obj = pledge_master_wb_obj.active
                                total_records = self.totalrecords_excelDataBase(pledge_master_file_path)
                                for iLoop in range(1, total_records + 1):
                                    cell_obj = pledge_master_sheet_obj.cell(row=iLoop + 1, column=2)
                                    # master sheet has multiple records for the same member , hence update comparing the pledge item
                                    if cell_obj.value == str(donator_idText.get()) and pledge_master_sheet_obj.cell(
                                            row=iLoop + 1, column=6).value == categoryText.get():
                                        rem_balance = int(
                                            pledge_master_sheet_obj.cell(row=iLoop + 1, column=12).value) - int(
                                            str(seva_amountText.get()))
                                        pledge_master_sheet_obj.cell(row=iLoop + 1, column=12).value = str(rem_balance)
                                        if rem_balance == 0:
                                            pledge_master_sheet_obj.cell(row=iLoop + 1, column=13).value = "Closed"
                                        break
                                pledge_master_wb_obj.save(pledge_master_file_path)

                                # update the global pledge item sheet
                                filename_MonetarySheet = self.obj_initDatabase.initilize_current_year_pledge_payment_database(
                                    trust_nametext.get(), categoryText.get())
                                if trust_nametext.get() == "Aadarsh gaushala Trust":
                                    invoice_id = self.obj_commonUtil.generateMonetaryDonationReceiptId(
                                        SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST)
                                else:
                                    invoice_id = self.obj_commonUtil.generateMonetaryDonationReceiptId(
                                        VIHANGAM_YOGA_KARNATAKA_TRUST)

                                # open current year pledge item sheet and write the data --start
                                wb_obj = openpyxl.load_workbook(filename_MonetarySheet)
                                sheet_obj = wb_obj.active
                                total_records = self.totalrecords_excelDataBase(filename_MonetarySheet)

                                if total_records is 0:
                                    serial_no = 1
                                    row_no = 2
                                else:
                                    serial_no = total_records + 1
                                    row_no = total_records + 2

                                # sets the general formatting for the new entry in new row
                                for index in range(1, 14):
                                    sheet_obj.cell(row=row_no, column=index).font = Font(size=12,
                                                                                         name='Times New Roman',
                                                                                         bold=False)
                                    sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                                   vertical='center')

                                # new book data is assigned to respective cells in row
                                sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                                sheet_obj.cell(row=row_no, column=2).value = str(seva_amountText.get())
                                sheet_obj.cell(row=row_no, column=3).value = str(donator_idText.get())
                                sheet_obj.cell(row=row_no, column=4).value = str(member_data[2])
                                sheet_obj.cell(row=row_no, column=5).value = str(dateOfCollection_calc)
                                sheet_obj.cell(row=row_no, column=6).value = str(categoryText.get())
                                sheet_obj.cell(row=row_no, column=7).value = str(collector_nameText.get())
                                sheet_obj.cell(row=row_no, column=8).value = str(revceiver_data[2])
                                sheet_obj.cell(row=row_no, column=9).value = str(member_data[7])
                                if paymentMode_text.get() == "Bank Transfer":
                                    paymenttext = str(paymentMode_text.get()) + "Ref. " + str(transId_Text.get())
                                else:
                                    paymenttext = str(paymentMode_text.get())
                                sheet_obj.cell(row=row_no, column=10).value = paymenttext
                                sheet_obj.cell(row=row_no, column=11).value = str(authorizedby_Text.get())
                                sheet_obj.cell(row=row_no, column=12).value = str(authorizor_data[2])
                                sheet_obj.cell(row=row_no, column=13).value = str(invoice_id)

                                wb_obj.save(filename_MonetarySheet)

                                filename_all_curYear_pledge_payment_sheet = self.obj_initDatabase.initilize_current_year_pledge_payment_database(
                                    trust_nametext.get(), "All")

                                # open current year pledge item sheet and write the data --start
                                wb_all_payment_curyear_obj = openpyxl.load_workbook(
                                    filename_all_curYear_pledge_payment_sheet)
                                sheet_obj_all_curPay_pledge = wb_all_payment_curyear_obj.active
                                total_records = self.totalrecords_excelDataBase(
                                    filename_all_curYear_pledge_payment_sheet)

                                if total_records is 0:
                                    serial_no = 1
                                    row_no = 2
                                else:
                                    serial_no = total_records + 1
                                    row_no = total_records + 2

                                # sets the general formatting for the new entry in new row
                                for index in range(1, 14):
                                    sheet_obj_all_curPay_pledge.cell(row=row_no, column=index).font = Font(size=12,
                                                                                                           name='Times New Roman',
                                                                                                           bold=False)
                                    sheet_obj_all_curPay_pledge.cell(row=row_no, column=index).alignment = Alignment(
                                        horizontal='center',
                                        vertical='center')

                                # new book data is assigned to respective cells in row
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=1).value = str(serial_no)
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=2).value = str(
                                    seva_amountText.get())
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=3).value = str(donator_idText.get())
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=4).value = str(member_data[2])
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=5).value = str(
                                    dateOfCollection_calc)
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=6).value = str(categoryText.get())
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=7).value = str(
                                    collector_nameText.get())
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=8).value = str(revceiver_data[2])
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=9).value = str(member_data[7])
                                if paymentMode_text.get() == "Bank Transfer":
                                    paymenttext = str(paymentMode_text.get()) + "Ref. " + str(transId_Text.get())
                                else:
                                    paymenttext = str(paymentMode_text.get())
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=10).value = paymenttext
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=11).value = str(
                                    authorizedby_Text.get())
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=12).value = str(authorizor_data[2])
                                sheet_obj_all_curPay_pledge.cell(row=row_no, column=13).value = str(invoice_id)

                                wb_all_payment_curyear_obj.save(filename_all_curYear_pledge_payment_sheet)

                                # This is a Monetary donation , hence updating the current year monetary donation sheet
                                # writting the credit in Master Seva Sheet - starts
                                filename_MonetarySevaSheet = self.obj_initDatabase.get_seva_deposit_database_name()  # PATH_SEVA_SHEET
                                # open seva rashi sheet and enter the data --start
                                wb_obj = openpyxl.load_workbook(filename_MonetarySevaSheet)
                                sheet_obj = wb_obj.active
                                total_records = self.totalrecords_excelDataBase(filename_MonetarySevaSheet)

                                if total_records is 0:
                                    serial_no = 1
                                    row_no = 2
                                    balance_amount = 0
                                else:
                                    serial_no = total_records + 1
                                    row_no = total_records + 2
                                    balance_amount = int(str(sheet_obj.cell(row=row_no - 1, column=3).value))

                                # sets the general formatting for the new entry in new row
                                for index in range(1, 14):
                                    sheet_obj.cell(row=row_no, column=index).font = Font(size=12,
                                                                                         name='Times New Roman',
                                                                                         bold=False)
                                    sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                                   vertical='center')

                                # new book data is assigned to respective cells in row
                                sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                                sheet_obj.cell(row=row_no, column=2).value = str(seva_amountText.get())
                                sheet_obj.cell(row=row_no, column=3).value = str(
                                    balance_amount + int(seva_amountText.get()))
                                sheet_obj.cell(row=row_no, column=4).value = str(donator_idText.get())
                                sheet_obj.cell(row=row_no, column=5).value = str(member_data[2])
                                sheet_obj.cell(row=row_no, column=6).value = str(dateOfCollection_calc)
                                sheet_obj.cell(row=row_no, column=7).value = str(categoryText.get())
                                sheet_obj.cell(row=row_no, column=8).value = str(collector_nameText.get())
                                sheet_obj.cell(row=row_no, column=9).value = str(revceiver_data[2])
                                sheet_obj.cell(row=row_no, column=10).value = str(member_data[7])  # address
                                if paymentMode_text.get() == "Bank Transfer":
                                    paymenttext = str(paymentMode_text.get()) + "Ref. " + str(transId_Text.get())
                                else:
                                    paymenttext = str(paymentMode_text.get())
                                sheet_obj.cell(row=row_no, column=11).value = paymenttext
                                sheet_obj.cell(row=row_no, column=12).value = str(authorizedby_Text.get())
                                sheet_obj.cell(row=row_no, column=13).value = str(authorizor_data[2])
                                sheet_obj.cell(row=row_no, column=14).value = str(invoice_id)
                                wb_obj.save(filename_MonetarySevaSheet)
                                # open seva rashi sheet and enter the data --end

                                # Pledge deposit shal also be updated in the respective trust transaction table
                                # receiving donation is a credit transaction for the organization
                                if trust_nametext.get() == "Aadarsh gaushala Trust":
                                    file_name_transaction = self.obj_initDatabase.get_gaushala_transaction_database_name()  # Gaushala main transaction sheet
                                else:
                                    file_name_transaction = self.obj_initDatabase.get_transaction_database_name()  # Trust transaction sheet

                                transaction_wb_obj = openpyxl.load_workbook(file_name_transaction)
                                transaction_sheet_obj = transaction_wb_obj.active
                                total_records_transaction = self.totalrecords_excelDataBase(file_name_transaction)

                                if total_records_transaction is 0:
                                    serial_no_trans_sheet = 1
                                    row_no_trans_sheet = 2
                                    balance_amount_trans_sheet = 0
                                else:
                                    serial_no_trans_sheet = total_records_transaction + 1
                                    row_no_trans_sheet = total_records_transaction + 2
                                    balance_amount_trans_sheet = int(
                                        str(transaction_sheet_obj.cell(row=row_no_trans_sheet - 1, column=9).value))

                                # sets the general formatting for the new entry in new row
                                for index in range(1, 10):
                                    transaction_sheet_obj.cell(row=row_no_trans_sheet, column=index).font = Font(
                                        size=12,
                                        name='Times New Roman',
                                        bold=False)
                                    transaction_sheet_obj.cell(row=row_no_trans_sheet,
                                                               column=index).alignment = Alignment(
                                        horizontal='center',
                                        vertical='center')

                                # new book data is assigned to respective cells in row
                                transaction_sheet_obj.cell(row=row_no_trans_sheet, column=1).value = str(
                                    serial_no_trans_sheet)
                                transaction_sheet_obj.cell(row=row_no_trans_sheet, column=2).value = str(
                                    dateOfCollection_calc)
                                transaction_sheet_obj.cell(row=row_no_trans_sheet, column=3).value = str(
                                    seva_amountText.get())
                                transaction_sheet_obj.cell(row=row_no_trans_sheet, column=4).value = "---"
                                transaction_sheet_obj.cell(row=row_no_trans_sheet, column=5).value = str(
                                    categoryText.get())
                                if paymentMode_text.get() == "Bank Transfer":
                                    paymenttext = str(paymentMode_text.get()) + "Ref. " + str(transId_Text.get())
                                else:
                                    paymenttext = str(paymentMode_text.get())
                                transaction_sheet_obj.cell(row=row_no_trans_sheet, column=6).value = paymenttext
                                transaction_sheet_obj.cell(row=row_no_trans_sheet, column=7).value = str(
                                    authorizedby_Text.get())  # authorizor id
                                transaction_sheet_obj.cell(row=row_no_trans_sheet, column=8).value = str(
                                    authorizor_data[2])
                                transaction_sheet_obj.cell(row=row_no_trans_sheet, column=9).value = str(
                                    balance_amount_trans_sheet + int(seva_amountText.get()))
                                transaction_sheet_obj.cell(row=row_no_trans_sheet, column=10).value = str(invoice_id)
                                transaction_wb_obj.save(file_name_transaction)
                                self.generateDonationReceipt(donator_idText,
                                                             seva_amountText,
                                                             categoryText,
                                                             revceiver_data[2],
                                                             dateOfCollection_calc,
                                                             paymentMode_text,
                                                             invoice_id, member_data, print_invoice,
                                                             self.submit_deposit)
                                invoice_idText['text'] = str(invoice_id)
                                text_withID = "Pledge amount successfully received. Receipt  id :" + str(invoice_id)
                                infolabel.configure(text=text_withID, fg='green')
                                break
                            else:
                                infolabel.configure(text="This pledge item is closed for member ...", fg='red')
                        else:
                            infolabel.configure(text="No pledge data for this member ...", fg='red')
                else:
                    infolabel.configure(text="Invalid Data for ID/IDs/Amount, please correct ...", fg='red')
            else:
                infolabel.configure(text="Transaction Date cannot be future!! please correct ...", fg='red')

    def subscribe_magazine_Excel(self, new_magazine_subscription_window,
                                 member_idText,
                                 seva_amountText,
                                 magazineCategoryText,
                                 quantityText,
                                 cal,
                                 paymentMode_text,
                                 authorizedby_idText,
                                 payableAmtText,
                                 infolabel, print_invoice):
        dateTimeObj = cal.get_date()
        dateOfCollection_calc = dateTimeObj.strftime("%d-%m-%Y ")
        if member_idText.get() == "" or \
                seva_amountText.get() == "" or \
                authorizedby_idText.get() == "":
            infolabel.configure(text="All fields are mandatory", fg='red')

        else:
            bValidateSubscription = self.validate_memberSubscriptionCurrentYear(member_idText.get(),
                                                                                magazineCategoryText, dateTimeObj)
            if not bValidateSubscription:
                today = date.today()
                if dateTimeObj <= today:
                    print("Member Id text :", member_idText.get())
                    bMemberIdValid = self.validate_memberlibraryID_Excel(member_idText.get(), 1)
                    bAuthorizorIdValid = self.validate_memberlibraryID_Excel(authorizedby_idText.get(), 1)
                    if bMemberIdValid and bAuthorizorIdValid and (seva_amountText.get()).isnumeric():

                        member_data = self.retrieve_MemberRecords_Excel(member_idText.get(), 1, SEARCH_BY_MEMBERID)

                        authorizor_data = self.retrieve_MemberRecords_Excel(authorizedby_idText.get(), 1,
                                                                            SEARCH_BY_MEMBERID)
                        filename_subscriptionSheet = self.obj_initDatabase.get_magazine_subscription_database_name()

                        # Subscription Amount is forwarded to transaction sheet

                        invoice_id = self.generate_invoiceID_magazineSubscription()
                        # open filename_subscriptionSheet  sheet and enter the data --start
                        wb_obj = openpyxl.load_workbook(filename_subscriptionSheet)
                        sheet_obj = wb_obj.active
                        total_records = self.totalrecords_excelDataBase(filename_subscriptionSheet)

                        if total_records is 0:
                            serial_no = 1
                            row_no = 2
                        else:
                            serial_no = total_records + 1
                            row_no = total_records + 2

                        # sets the general formatting for the new entry in new row
                        for index in range(1, 13):
                            sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                                 bold=False)
                            sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                           vertical='center',
                                                                                           wrapText=True)

                        # new book data is assigned to respective cells in row
                        sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                        sheet_obj.cell(row=row_no, column=2).value = str(member_idText.get())
                        sheet_obj.cell(row=row_no, column=3).value = str(member_data[2])
                        sheet_obj.cell(row=row_no, column=4).value = str(dateOfCollection_calc)
                        sheet_obj.cell(row=row_no, column=5).value = str(magazineCategoryText.get())
                        sheet_obj.cell(row=row_no, column=6).value = str(quantityText.get())
                        sheet_obj.cell(row=row_no, column=7).value = str(seva_amountText.get())
                        sheet_obj.cell(row=row_no, column=8).value = str(member_data[7])
                        sheet_obj.cell(row=row_no, column=9).value = str(authorizedby_idText.get())  # address
                        sheet_obj.cell(row=row_no, column=10).value = str(authorizor_data[2])
                        '''
                        if paymentMode_text.get() == "Bank Transfer":
                            paymenttext = str(paymentMode_text.get()) + "Ref. " + str(transId_Text.get())
                        else:
                            paymenttext = str(paymentMode_text.get())
                            '''
                        sheet_obj.cell(row=row_no, column=11).value = str(paymentMode_text.get())
                        sheet_obj.cell(row=row_no, column=12).value = str(invoice_id)
                        payableAmtText.configure(text=str(int(quantityText.get()) * int(seva_amountText.get())))

                        wb_obj.save(filename_subscriptionSheet)

                        # magazine distribution database is updated with the member ID
                        filename_distributionSheet = self.obj_initDatabase.get_magazine_distribution_database_name()

                        wb_distribution_database_obj = openpyxl.load_workbook(filename_distributionSheet)
                        distribution_sheet_obj = wb_distribution_database_obj.active
                        total_records = self.totalrecords_excelDataBase(filename_distributionSheet)

                        if total_records is 0:
                            serial_no = 1
                            row_no = 2
                        else:
                            serial_no = total_records + 1
                            row_no = total_records + 2

                        # sets the general formatting for the new entry in new row
                        for index in range(1, 15):
                            distribution_sheet_obj.cell(row=row_no, column=index).font = Font(size=12,
                                                                                              name='Times New Roman',
                                                                                              bold=False)
                            distribution_sheet_obj.cell(row=row_no, column=index).alignment = Alignment(
                                horizontal='center',
                                vertical='center')

                        # new book data is assigned to respective cells in row
                        distribution_sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                        distribution_sheet_obj.cell(row=row_no, column=2).value = str(member_idText.get())

                        # Fill the distribution record as NA for all the months
                        for index in range(3, 15):
                            distribution_sheet_obj.cell(row=row_no, column=index).value = "NA"

                        wb_distribution_database_obj.save(filename_distributionSheet)

                        '''' Future code
                        file_name_transaction = PATH_TRANSACTION_SHEET
                        transaction_wb_obj = openpyxl.load_workbook(file_name_transaction)
                        transaction_sheet_obj = transaction_wb_obj.active
                        total_records_transaction = self.totalrecords_excelDataBase(file_name_transaction)

                        if total_records_transaction is 0:
                            serial_no = 1
                            row_no = 2
                            balance_amount = 0
                        else:
                            serial_no = total_records_transaction + 1
                            row_no = total_records_transaction + 2
                            balance_amount = int(str(transaction_sheet_obj.cell(row=row_no - 1, column=9).value))

                        # sets the general formatting for the new entry in new row
                        for index in range(1, 13):
                            transaction_sheet_obj.cell(row=row_no, column=index).font = Font(size=12,
                                                                                             name='Times New Roman',
                                                                                             bold=False)
                            transaction_sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                                       vertical='center')

                        # new book data is assigned to respective cells in row
                        transaction_sheet_obj.cell(row=row_no, column=1).value = serial_no
                        transaction_sheet_obj.cell(row=row_no, column=2).value = dateOfCollection_calc
                        transaction_sheet_obj.cell(row=row_no, column=3).value = seva_amountText.get()
                        transaction_sheet_obj.cell(row=row_no, column=4).value = "---"
                        transaction_sheet_obj.cell(row=row_no, column=5).value = categoryText.get()
                        if paymentMode_text.get() == "Bank Transfer":
                            paymenttext = str(paymentMode_text.get()) + "Ref. " + str(transId_Text.get())
                        else:
                            paymenttext = str(paymentMode_text.get())
                        transaction_sheet_obj.cell(row=row_no, column=6).value = paymentMode_text.get()
                        transaction_sheet_obj.cell(row=row_no, column=7).value = authorizedby_Text.get()  # authorizor id
                        transaction_sheet_obj.cell(row=row_no, column=8).value = authorizor_data[2]
                        transaction_sheet_obj.cell(row=row_no, column=9).value = str(
                            balance_amount + int(seva_amountText.get()))
                        transaction_wb_obj.save(file_name_transaction)
                        '''
                        # open transaction sheet and enter the data --end

                        print_result = partial(self.generatePatrikaSubscription_Receipt,
                                               member_data,
                                               seva_amountText,
                                               quantityText,
                                               authorizor_data,
                                               dateOfCollection_calc,
                                               magazineCategoryText,
                                               paymentMode_text,
                                               invoice_id)
                        print_invoice.configure(state=NORMAL, bg='light cyan', command=print_result)
                        self.submit_deposit.configure(state=DISABLED, bg='light grey')
                        text_withID = "Seva deposited successfully. Invoice  id :" + invoice_id
                        infolabel.configure(text=text_withID, fg='green')
                    else:
                        infolabel.configure(text="Invalid Data for ID/IDs/Amount, please correct ...", fg='red')
                else:
                    infolabel.configure(text="Transaction Date cannot be future!! please correct ...", fg='red')
            else:
                infolabel.configure(text="Subscription Already Exists for requested year!! ", fg='red')

    def deposit_non_monetary_Excel(self, new_nonmonetary_donation_window,
                                   donator_idText,
                                   item_name,
                                   quantityText,
                                   collector_nameText,
                                   authorizorId_Text,
                                   cal,
                                   estValueText,
                                   invoice_idText,
                                   infolabel,
                                   print_invoice,
                                   submit_Btn,
                                   owner_type,
                                   location_text, local_centerText):
        dateTimeObj = cal.get_date()
        dateOfCollection_calc = dateTimeObj.strftime("%d-%m-%Y ")

        if donator_idText.get() == "" or \
                item_name.get() == "" or \
                collector_nameText.get() == "" or \
                quantityText.get() == "" or \
                estValueText.get() == "":

            infolabel.configure(text="All fields are mandatory", fg='red')

        else:
            today = date.today()
            if dateTimeObj <= today:
                bDonatorIdValid = self.validate_memberlibraryID_Excel(donator_idText.get(), 1)
                bReceiverIdValid = self.validate_memberlibraryID_Excel(collector_nameText.get(), 1)
                bAuthorizorIdValid = self.validate_memberlibraryID_Excel(authorizorId_Text.get(), 1)

                if bDonatorIdValid and bReceiverIdValid and bAuthorizorIdValid and \
                        (estValueText.get()).isnumeric() and (quantityText.get()).isnumeric():

                    member_data = self.retrieve_MemberRecords_Excel(donator_idText.get(), 1, SEARCH_BY_MEMBERID)
                    revceiver_data = self.retrieve_MemberRecords_Excel(collector_nameText.get(), 1, SEARCH_BY_MEMBERID)
                    authorizor_data = self.retrieve_MemberRecords_Excel(authorizorId_Text.get(), 1, SEARCH_BY_MEMBERID)

                    subdir_noncommercialstock = "..\\Library_Stock\\" + local_centerText.get() + "\\NonCommercial_Stock"
                    filename_nonMonetarySheet = subdir_noncommercialstock + "\\noncommercial_stock.xlsx"  # PATH_NON_MONETARY_SHEET

                    # Currently non monetary donations are not forwarded to transaction sheet
                    # if in case need arises , code respective to below transaction sheet can be enabled
                    # to process the same .
                    # file_name_transaction = PATH_TRANSACTION_SHEET
                    if owner_type != STOCK_OWNER_TYPE_ASHRAM:
                        invoice_id = self.generate_invoiceID_nonMonetaryDeposit()
                    else:
                        invoice_id = "N/A"
                    # open sheet and enter the data --start
                    wb_obj = openpyxl.load_workbook(filename_nonMonetarySheet)
                    sheet_obj = wb_obj.active
                    total_records = self.totalrecords_excelDataBase(filename_nonMonetarySheet)

                    if total_records is 0:
                        serial_no = 1
                        row_no = 2
                    else:
                        serial_no = total_records + 1
                        row_no = total_records + 2

                    # sets the general formatting for the new entry in new row
                    for index in range(1, 18):
                        sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                             bold=False)
                        sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                       vertical='center')

                    inventory_id = self.generate_itemId_nonCommenrcial(local_centerText.get())
                    # new book data is assigned to respective cells in row
                    sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                    sheet_obj.cell(row=row_no, column=2).value = str(inventory_id)
                    sheet_obj.cell(row=row_no, column=3).value = str(item_name.get())
                    sheet_obj.cell(row=row_no, column=4).value = str(quantityText.get())
                    sheet_obj.cell(row=row_no, column=5).value = str(estValueText.get())
                    sheet_obj.cell(row=row_no, column=6).value = str(donator_idText.get())
                    sheet_obj.cell(row=row_no, column=7).value = str(member_data[2])
                    sheet_obj.cell(row=row_no, column=8).value = str(member_data[7])
                    sheet_obj.cell(row=row_no, column=9).value = str(dateOfCollection_calc)
                    sheet_obj.cell(row=row_no, column=10).value = str(collector_nameText.get())
                    sheet_obj.cell(row=row_no, column=11).value = str(revceiver_data[2])
                    sheet_obj.cell(row=row_no, column=12).value = str(authorizorId_Text.get())
                    sheet_obj.cell(row=row_no, column=13).value = str(authorizor_data[2])
                    sheet_obj.cell(row=row_no, column=14).value = str(invoice_id)
                    sheet_obj.cell(row=row_no, column=15).value = str(owner_type)
                    sheet_obj.cell(row=row_no, column=16).value = str(location_text.get())
                    sheet_obj.cell(row=row_no, column=17).value = str(local_centerText.get())

                    wb_obj.save(filename_nonMonetarySheet)
                    # open  sheet and enter the data --End

                    # open transaction sheet and enter the data --start -future code , if required
                    # receiving donation is a credit transaction for the organization
                    '''
                    transaction_wb_obj = openpyxl.load_workbook(file_name_transaction)
                    transaction_sheet_obj = transaction_wb_obj.active
                    total_records_transaction = self.totalrecords_excelDataBase(file_name_transaction)

                    if total_records_transaction is 0:
                        serial_no = 1
                        row_no = 2
                    else:
                        serial_no = total_records_transaction + 1
                        row_no = total_records_transaction + 2

                    # sets the general formatting for the new entry in new row
                    for index in range(1, 8):
                        transaction_sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                                         bold=False)
                        transaction_sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                                   vertical='center')

                    # new book data is assigned to respective cells in row
                    transaction_sheet_obj.cell(row=row_no, column=1).value = serial_no
                    transaction_sheet_obj.cell(row=row_no, column=2).value = dateOfCollection_calc
                    transaction_sheet_obj.cell(row=row_no, column=3).value = seva_amountText.get()
                    transaction_sheet_obj.cell(row=row_no, column=4).value = "---"
                    transaction_sheet_obj.cell(row=row_no, column=5).value = categoryText.get()
                    transaction_sheet_obj.cell(row=row_no, column=6).value = paymentMode_text.get()
                    transaction_sheet_obj.cell(row=row_no, column=7).value = collector_nameText.get()
                    transaction_wb_obj.save(file_name_transaction)
                    '''
                    invoice_idText['text'] = invoice_id
                    invoice_idText.configure(font=TIMES_NEW_ROMAN_BIG)
                    # open transaction sheet and enter the data --end
                    submit_Btn.configure(state=DISABLED, bg='light grey')
                    if owner_type != STOCK_OWNER_TYPE_ASHRAM:
                        print_result = partial(self.generateDonation_NonMonetaryReceipt, donator_idText,
                                               item_name,
                                               quantityText,
                                               collector_nameText,
                                               str(revceiver_data[2]),
                                               dateOfCollection_calc,
                                               estValueText,
                                               invoice_id,
                                               member_data)
                        print_invoice.configure(state=NORMAL, bg='light cyan', command=print_result)
                    text_withID = "Item inserted with Invoice  id :" + invoice_id + " Stock Id :" + str(inventory_id)
                    infolabel.configure(text=text_withID, fg='green')
                else:
                    infolabel.configure(text="Invalid Data for ID/IDs/Amount, please correct ...", fg='red')
            else:
                infolabel.configure(text="Transaction Date cannot be future date, please correct ...", fg='red')

    def generateDonation_NonMonetaryReceipt(self, donator_idText,
                                            seva_amountText,
                                            categoryText,
                                            collector_id,
                                            collector_name,
                                            dateOfCollection_calc,
                                            paymentMode_text,
                                            invoice_id, member_data):
        currentYearDirName = self.obj_commonUtil.getCurrentYearFolderName()
        file_name = "..\\Expanse_Data\\" + currentYearDirName + "\\Seva_Rashi\\Receipts\\Template\\Donation_NonMonetary_Receipt_Template.xlsx"
        wb_obj = openpyxl.load_workbook(file_name)
        sheet_obj = wb_obj.active

        sheet_obj.cell(row=6, column=4).value = str(invoice_id)
        sheet_obj.cell(row=7, column=4).value = str(dateOfCollection_calc)
        sheet_obj.cell(row=8, column=4).value = str(member_data[2])
        sheet_obj.cell(row=9, column=4).value = str(member_data[7])  # Address
        sheet_obj.cell(row=10, column=4).value = str(member_data[8])  # city
        sheet_obj.cell(row=11, column=4).value = str(member_data[9])  # state
        sheet_obj.cell(row=12, column=4).value = str(member_data[10])  # postal code
        sheet_obj.cell(row=13, column=4).value = str(member_data[11])  # contact number
        sheet_obj.cell(row=15, column=4).value = str(seva_amountText.get())

        sheet_obj.cell(row=16, column=4).value = str(paymentMode_text.get())
        sheet_obj.cell(row=17, column=4).value = str(categoryText.get())
        sheet_obj.cell(row=18, column=4).value = str(collector_name) + "(" + str(collector_id.get()) + ")"

        wb_obj.save(file_name)
        pdf_file = self.obj_initDatabase.get_invoice_directory_name() + "\\" + invoice_id + ".pdf"
        desktop_backup_dir = self.obj_initDatabase.get_desktop_invoices_directory_path()
        self.obj_commonUtil.convertExcelToPdf(file_name, pdf_file)
        shutil.copy(pdf_file, desktop_backup_dir)
        os.startfile(pdf_file, 'print')

    def create_Advance_Excel(self, createExpanse_window,
                             receiver_nameText,
                             seva_amountText,
                             descriptionText,
                             authorizerText,
                             cal,
                             paymentMode_text,
                             infolabel):

        current_balance = self.obj_commonUtil.readcurrent_balance()
        if int(seva_amountText.get()) < int(current_balance):
            dateTimeObj = cal.get_date()
            dateOfExpanse_calc = dateTimeObj.strftime("%Y-%m-%d")
            if receiver_nameText.get() == "" or \
                    seva_amountText.get() == "" or \
                    authorizerText.get() == "" or \
                    paymentMode_text.get() == "":

                infolabel.configure(text="All fields are mandatory", fg='red')

            else:
                today = date.today()
                if dateTimeObj <= today:
                    bReceiverIdValid = self.validate_memberlibraryID_Excel(receiver_nameText.get(), 1)
                    bAuthorizorIdValid = self.validate_memberlibraryID_Excel(authorizerText.get(), 1)

                    if bReceiverIdValid and bAuthorizorIdValid and \
                            (seva_amountText.get()).isnumeric():

                        receiver_data = self.retrieve_MemberRecords_Excel(receiver_nameText.get(), 1,
                                                                          SEARCH_BY_MEMBERID)
                        authorizor_data = self.retrieve_MemberRecords_Excel(authorizerText.get(), 1, SEARCH_BY_MEMBERID)

                        file_name_sevaRashi = self.obj_initDatabase.get_advance_database_name()  # PATH_ADVANCE_SHEET

                        invoice_id = self.generate_advanceId()
                        # open expanse sheet and enter the data --start
                        wb_obj = openpyxl.load_workbook(file_name_sevaRashi)
                        sheet_obj = wb_obj.active
                        total_records = self.totalrecords_excelDataBase(file_name_sevaRashi)

                        if total_records is 0:
                            serial_no = 1
                            row_no = 2

                        else:
                            serial_no = total_records + 1
                            row_no = total_records + 2

                        # sets the general formatting for the new entry in new row
                        for index in range(1, 10):
                            sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                                 bold=False)
                            sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                           vertical='center')

                        sheet_obj.cell(row=row_no, column=1).value = serial_no
                        sheet_obj.cell(row=row_no, column=2).value = dateOfExpanse_calc
                        sheet_obj.cell(row=row_no, column=3).value = seva_amountText.get()
                        sheet_obj.cell(row=row_no, column=4).value = seva_amountText.get()
                        sheet_obj.cell(row=row_no, column=5).value = descriptionText.get()
                        sheet_obj.cell(row=row_no, column=6).value = receiver_nameText.get()
                        sheet_obj.cell(row=row_no, column=7).value = receiver_data[2]  # receiver name
                        sheet_obj.cell(row=row_no, column=8).value = authorizerText.get()  # authorizor id
                        sheet_obj.cell(row=row_no, column=9).value = authorizor_data[0]
                        sheet_obj.cell(row=row_no, column=10).value = paymentMode_text.get()
                        sheet_obj.cell(row=row_no, column=11).value = str(invoice_id)

                        wb_obj.save(file_name_sevaRashi)
                        text_withID = "Advance  has been issue to :" + str(invoice_id)
                        infolabel.configure(text=text_withID, fg='green')
                    else:
                        infolabel.configure(text="ID/IDs or Amount entered !! please correct", fg='red')
                else:
                    infolabel.configure(text="Future Advance cannot be created !! please correct", fg='red')
        else:
            infolabel.configure(text="In-sufficient balance to issue Advance !!!", fg='red')

    def update_Advance_Excel(self, advanceId,
                             retAmount,
                             infolabel):
        if advanceId.get() == "" or \
                retAmount.get() == "":
            infolabel.configure(text="All fields are mandatory", fg='red')
        else:
            bAdvanceIdValid = self.validate_advanceIdExcel(advanceId)
            if bAdvanceIdValid or (retAmount.get()).isnumeric():
                file_name_sevaRashi = self.obj_initDatabase.get_advance_database_name()  # PATH_ADVANCE_SHEET
                # open expanse sheet and enter the data --start
                wb_obj = openpyxl.load_workbook(file_name_sevaRashi)
                sheet_obj = wb_obj.active
                total_records = self.totalrecords_excelDataBase(file_name_sevaRashi)

                for row_idx in range(1, total_records + 1):
                    advance_sheetid = sheet_obj.cell(row=row_idx + 1, column=11).value
                    if advance_sheetid == advanceId.get():
                        advance_amount = sheet_obj.cell(row=row_idx + 1, column=4).value
                        if int(retAmount.get()) <= int(advance_amount):
                            sheet_obj.cell(row=row_idx + 1, column=4).value = str(
                                int(advance_amount) - int(retAmount.get()))
                            wb_obj.save(file_name_sevaRashi)
                            text_withID = "Advance  has been adjusted for :" + advanceId.get()
                            infolabel.configure(text=text_withID, fg='green')
                        else:
                            text_withID = "Returned Amount cannot be greater than Advance amount" + advanceId
                            infolabel.configure(text=text_withID, fg='red')
                    print("Balance is adjusted")
                    break
            else:
                infolabel.configure(text="Incorrect Advance ID/Amount entered !! please correct", fg='red')

    def data_validation(self, member_govtId, member_name,
                        member_fatherName,
                        member_mother,
                        member_address,
                        member_city,
                        member_state,
                        member_contactNo,
                        member_country):
        bValidData = False
        if str(member_address) == "" or member_city.get() == "" or \
                member_fatherName.get() == "" or \
                member_mother.get() == "":
            bValidData = False
            invalidReason = FIELD_BLANK
        elif len(member_name.get()) > 30 or len(member_fatherName.get()) > 30 or \
                len(member_mother.get()) > 30:
            bValidData = False
            invalidReason = MAX_LEN_EXCEED
        else:
            bValidData = True
            invalidReason = 0
        return bValidData, invalidReason

    def register_member_excel(self, addMember_window, member_id, member_govtId,
                              member_name,
                              member_fatherName,
                              member_mother,
                              cal,
                              member_gender,
                              member_address,
                              member_city,
                              member_state,
                              member_pincode,
                              member_contactNo,
                              member_country,
                              member_nationality,
                              member_emailId,
                              member_idType,
                              cal_associatedSince,
                              profession,
                              updestha,
                              updesh_stage,
                              designation_varaible,
                              akshayPatra_varaible,
                              akshayPatra_No,
                              magazine_subsVariable,
                              memberType, infoLabel):
        print("register_member_excel -->> Start")

        bValidData, invalid_reason = self.data_validation(member_govtId, member_name,
                                                          member_fatherName,
                                                          member_mother,
                                                          member_address,
                                                          member_city,
                                                          member_state,
                                                          member_contactNo,
                                                          member_country)
        dateTimeObj = cal.get_date()
        member_dob = dateTimeObj.strftime("%d-%b-%Y ")
        dateTimeObj_associatedSince = cal_associatedSince.get_date()
        member_associatedSince = dateTimeObj_associatedSince.strftime("%d-%b-%Y ")
        if not bValidData:
            if invalid_reason == -1:
                infoLabel.configure(fg='red')
                infoLabel['text'] = "Data Entry Error! Red fields are mandatory !!!"
            else:
                infoLabel[
                    'text'] = "Max length exceeded for Member Name,\n Parents name shall not exceed 30 characters !!!"
        else:
            bMemberExists = self.validate_memberGovtID_Excel(member_govtId.get())
            if bMemberExists:
                infoLabel.configure(fg='red')
                infoLabel['text'] = "Duplicate Entry Error ! Member already registered !!"
                member_name.configure(bd=2, fg='red')
                return
            else:
                # Call a Workbook() function of openpyxl
                # to create a new blank Workbook object

                path = PATH_MEMBER

                wb_obj = openpyxl.load_workbook(path)
                sheet_obj = wb_obj.active

                totalRecords = self.totalrecords_excelDataBase(path)
                # generate the new member id here itself
                self.newmember_id = self.obj_commonUtil.generate_new_memberId()
                # writing the values of the member data in respective column, row number remains the same .
                data_cell = {}
                # --------------------------Member Record entry - start----------------------------------
                if totalRecords is 0:
                    serial_no = 1
                    row_no = 2
                else:
                    serial_no = totalRecords + 1
                    row_no = totalRecords + 2

                # sets the general formatting for the new entry in new row
                for index in range(1, MAX_RECORD_ENTRY + 1):
                    sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                         bold=False)
                    sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                   vertical='center', wrapText=True)

                # new book data is assigned to respective cells in row
                sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                sheet_obj.cell(row=row_no, column=2).value = str(self.newmember_id)
                sheet_obj.cell(row=row_no, column=3).value = str(member_govtId.get())
                sheet_obj.cell(row=row_no, column=4).value = str(member_name.get())
                sheet_obj.cell(row=row_no, column=5).value = str(member_fatherName.get())
                sheet_obj.cell(row=row_no, column=6).value = str(member_mother.get())
                sheet_obj.cell(row=row_no, column=7).value = str(member_dob)
                sheet_obj.cell(row=row_no, column=8).value = str(member_gender.get())
                sheet_obj.cell(row=row_no, column=9).value = str(member_address.get())
                sheet_obj.cell(row=row_no, column=10).value = str(member_city.get())
                sheet_obj.cell(row=row_no, column=11).value = str(member_state.get())
                sheet_obj.cell(row=row_no, column=12).value = str(member_pincode.get())
                sheet_obj.cell(row=row_no, column=13).value = str(member_contactNo.get())
                sheet_obj.cell(row=row_no, column=14).value = str(member_country.get())
                sheet_obj.cell(row=row_no, column=15).value = str(member_nationality.get())
                sheet_obj.cell(row=row_no, column=16).value = str(member_emailId.get())
                sheet_obj.cell(row=row_no, column=17).value = str(self.member_photoFilePath)
                sheet_obj.cell(row=row_no, column=18).value = str(self.member_IdPhotoFilePath)
                sheet_obj.cell(row=row_no, column=19).value = str(member_idType.get())
                sheet_obj.cell(row=row_no, column=20).value = str(member_associatedSince)

                sheet_obj.cell(row=row_no, column=21).value = str(profession.get())
                sheet_obj.cell(row=row_no, column=22).value = str(updestha.get())
                sheet_obj.cell(row=row_no, column=23).value = str(updesh_stage.get())
                sheet_obj.cell(row=row_no, column=24).value = str(designation_varaible.get())
                sheet_obj.cell(row=row_no, column=25).value = str(akshayPatra_varaible.get())
                sheet_obj.cell(row=row_no, column=26).value = str(akshayPatra_No.get())
                sheet_obj.cell(row=row_no, column=27).value = str(magazine_subsVariable.get())

                wb_obj.save(PATH_MEMBER)
                # --------------------------Member Record entry - start----------------------------------
                info_text = "Registration Success !!! Member Id : " + str(self.newmember_id)

                # update the total member count
                self.obj_commonUtil.update_totalMemberRecords()

                # create the staff login as soon as staff is registered
                if designation_varaible.get() == "Staff-Sevak" or \
                        designation_varaible.get() == "Manager" or \
                        designation_varaible.get() == "Accountant" or \
                        designation_varaible.get() == "President" or \
                        designation_varaible.get() == "Vice-President":
                    staff_login_path = PATH_STAFF_CREDENTIALS
                    wb = openpyxl.load_workbook(staff_login_path)
                    sheet = wb.active
                    totalRecords = self.totalrecords_excelDataBase(staff_login_path)
                    login_data = {}
                    for iLoop in range(1, 4):
                        if iLoop == 1:
                            if totalRecords == 1:
                                serial_no = 1
                            else:
                                serial_no = totalRecords + 1
                            # record serial number is total_rows - 1 , since excluding top header row
                            login_data[1] = serial_no
                        if iLoop == 2:
                            login_data[2] = str(self.newmember_id)
                        if iLoop == 3:
                            login_data[3] = "Password@123"
                        sheet.cell(row=totalRecords + 2, column=iLoop).font = Font(size=12, name='Times New Roman',
                                                                                   bold=False)
                        sheet.cell(row=totalRecords + 2, column=iLoop).alignment = Alignment(horizontal='center',
                                                                                             vertical='center')
                        sheet.cell(row=totalRecords + 2, column=iLoop).value = login_data[iLoop]
                    print("Login has been created for the staff ", str(self.newmember_id))
                    # save the sheet after data is written
                    wb.save(staff_login_path)

                self.clear_Memberform(member_govtId,
                                      member_name,
                                      member_fatherName,
                                      member_mother,
                                      member_gender,
                                      member_address,
                                      member_city,
                                      member_state,
                                      member_pincode,
                                      member_contactNo,
                                      member_country,
                                      member_nationality,
                                      member_emailId)

                infoLabel.configure(fg='green')
                infoLabel['text'] = info_text
                self.submit.configure(state=DISABLED, bg='light grey')

            print("register_member_excel -->> End")

    def totalrecords_excelDataBase(self, path):
        # to open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj.active

        # print the total number of rows
        return sheet_obj.max_row - 1

    def validateStaffAccountExists(self, memberId):
        wb_obj = openpyxl.load_workbook(PATH_STAFF_CREDENTIALS)
        sheet_obj = wb_obj.active
        buserOk = False
        total_records = self.totalrecords_excelDataBase(PATH_STAFF_CREDENTIALS)
        # print("User entered  - username : ", username, "Password :", password, "Total records:", total_records)
        for iLoop in range(0, total_records):
            if str(sheet_obj.cell(row=iLoop + 2, column=2).value) == memberId:
                print("User Exists !!!!")
                buserOk = True
                break
        print("validateStaffCredentialsExistence->>End", buserOk)
        return buserOk

    def pic_upload(self, addMember_window, uploadType, myFrameId, canvasMem):
        filename = filedialog.askopenfilename(initialdir="/", title="Select File to Upload (Only Jpeg)",
                                              filetypes=(("jpeg files", "*.jpg"), ("all files", "*.*")))
        image1 = PIL.Image.open(filename)
        myimage = ImageTk.PhotoImage(image1.resize((150, 150)))
        canvasMem.create_image(0, 0, anchor=NW, image=myimage)
        if uploadType == 1:
            self.member_photoFilePath = filename
            print("Member photo path :", self.member_photoFilePath)
        elif uploadType == 2:
            self.member_IdPhotoFilePath = filename
            print("Id photo path :", self.member_IdPhotoFilePath)
        canvasMem.pack()
        mainloop()

    def read_webcam(self, addMember_window, uploadType, myFrameId, canvasMem):
        print("Read webcam entry ---")
        key = cv2.waitKey(1)
        webcam = cv2.VideoCapture(0)
        while True:
            try:
                check, frame = webcam.read()
                cv2.imshow("Capturing ...press 's' to click or 'q' to exit", frame)
                key = cv2.waitKey(1)
                if key == ord('s'):
                    cv2.imwrite(filename='..\\Images\\saved_img.jpg', img=frame)
                    webcam.release()
                    img_new = cv2.imread('..\\Images\\saved_img.jpg', cv2.IMREAD_ANYCOLOR)
                    img_new = cv2.imshow("Captured Image", img_new)
                    cv2.waitKey(1650)
                    cv2.destroyAllWindows()
                    img_ = cv2.imread('..\\Images\\saved_img.jpg', cv2.IMREAD_ANYCOLOR)
                    # print("Converting RGB image to grayscale...")
                    gray = cv2.cvtColor(img_, cv2.COLOR_BGR2GRAY)
                    # print("Converted RGB image to grayscale...")
                    # print("Resizing image to 28x28 scale...")
                    img_ = cv2.resize(gray, (150, 150))
                    # print("Resized...")
                    member_id = self.obj_commonUtil.get_current_new_memberID()
                    file_name = "..\\Images" + "\\" + str(member_id) + ".jpg"

                    if uploadType == 1:
                        self.member_photoFilePath = file_name
                    elif uploadType == 2:
                        self.member_IdPhotoFilePath = file_name
                    else:
                        pass

                    img_resized = cv2.imwrite(filename=file_name, img=img_)

                    image1 = Image.open(file_name)
                    myimage = ImageTk.PhotoImage(image1.resize((150, 150)))
                    canvasMem.create_image(0, 0, anchor=NW, image=myimage)
                    canvasMem.pack()
                    mainloop()
                    break
                elif key == ord('q'):
                    # print("Turning off camera.")
                    webcam.release()
                    # print("Camera off.")
                    # print("Program ended.")
                    cv2.destroyAllWindows()
                    break
            except(KeyboardInterrupt):
                # print("Turning off camera.")
                webcam.release()
                # print("Camera off.")
                # print("Program ended.")
                cv2.destroyAllWindows()
                break

    def generate_id_card(self, memberType):
        generateId_card_Window = Toplevel(self.master)
        if memberType == 1:
            headingForm = "Member Identity Card"
            generateId_card_Window.title("Member Identity Card")

        generateId_card_Window.geometry('790x480+250+50')
        generateId_card_Window.configure(background='powder blue')
        generateId_card_Window.resizable(width=True, height=True)

        upperFrame = Frame(generateId_card_Window, width=205, height=100, bd=4, relief='ridge', bg='light yellow')
        upperFrame.grid(row=1, column=0, padx=40, pady=10, sticky=W, columnspan=5)
        heading = Label(upperFrame, text=headingForm, font=('times new roman', 25, 'normal'), bg='light yellow')
        heading.grid(row=0, column=0, columnspan=3)
        middleFrame = Frame(generateId_card_Window, width=200, height=300, bd=8, relief='ridge', bg='light yellow')
        middleFrame.grid(row=2, column=2, padx=50, pady=10, sticky=W)

        memberIdLabel = Label(upperFrame, text="Member Id", width=9, anchor=W, justify=LEFT,
                              font=('arial narrow', 15, 'normal'), bg='light yellow')
        memberIdLabel.grid(row=1, column=0, padx=30, pady=10)
        member_Id = Entry(upperFrame, width=7, font=('arial narrow', 15, 'normal'), justify='center')
        member_Id.grid(row=1, column=1, ipadx=60, pady=10)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(upperFrame, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=1, column=2, padx=15, pady=10, sticky=W)

        # ---------------------------------Button Frame End----------------------------------------

        # ---------------------------------Preparing display Area - start ---------------------------------

        # Display Member Name - Row 1
        memberNameLabel = Label(middleFrame, text="Name", width=10, anchor=W, justify=LEFT,
                                font=('arial narrow', 12, 'normal'),
                                bg='light yellow')
        memberNameLabel.grid(row=1, column=2, padx=0, pady=5)
        memberNameText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                               font=('arial narrow', 13, 'normal'),
                               bg='light yellow')
        memberNameText.grid(row=1, column=3, padx=5, pady=5)

        # Display Member Id - Row 1
        memberIdLabel = Label(middleFrame, text="Member Id", width=10, anchor=W, justify=LEFT,
                              font=('arial narrow', 12, 'normal'),
                              bg='light yellow')
        memberIdLabel.grid(row=1, column=4, padx=10, pady=5)
        memberIdText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                             font=('arial narrow', 13, 'normal'),
                             bg='light yellow')
        memberIdText.grid(row=1, column=5, padx=5, pady=5)

        # Display Date Of Birth - Row 3
        memberDOBLabel = Label(middleFrame, text="Date Of Birth", width=10, anchor=W, justify=LEFT,
                               font=('arial narrow', 12, 'normal'),
                               bg='light yellow')
        memberDOBLabel.grid(row=3, column=2, padx=10, pady=5)
        memberDOBText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                              font=('arial narrow', 13, 'normal'),
                              bg='light yellow')
        memberDOBText.grid(row=3, column=3, padx=5, pady=5)

        # Display Mother Name - Row 3
        memberGenderLabel = Label(middleFrame, text="Gender", width=10, anchor=W, justify=LEFT,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='light yellow')
        memberGenderLabel.grid(row=3, column=4, padx=10, pady=5)
        memberGenderText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                 font=('arial narrow', 13, 'normal'),
                                 bg='light yellow')
        memberGenderText.grid(row=3, column=5, padx=5, pady=5)

        # Display Member Name - Row 4
        memberContactLabel = Label(middleFrame, text="Contact No", width=10, anchor=W, justify=LEFT,
                                   font=('arial narrow', 12, 'normal'),
                                   bg='light yellow')
        memberContactLabel.grid(row=4, column=2, padx=0, pady=5)
        memberContactNoText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                    font=('arial narrow', 13, 'normal'),
                                    bg='light yellow')
        memberContactNoText.grid(row=4, column=3, padx=5, pady=5)

        # Display Country Name - Row 8
        memberIdTypeLabel = Label(middleFrame, text="Designation", width=10, anchor=W, justify=LEFT,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='light yellow')
        memberIdTypeLabel.grid(row=4, column=4, padx=10, pady=5)
        memberDesignationText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                      font=('arial narrow', 13, 'normal'),
                                      bg='light yellow')
        memberDesignationText.grid(row=4, column=5, padx=5, pady=5)

        # Display Photo - Row 9
        memberPhotoLabel = Label(middleFrame, text="Member Photo", width=11, anchor=W, justify=LEFT,
                                 font=('arial narrow', 12, 'normal'),
                                 bg='light yellow')
        memberPhotoLabel.grid(row=9, column=2, padx=10, pady=5)

        # Display Photo - Row 9
        officeSealLabel = Label(middleFrame, text="Office Seal/Signature", width=16, anchor=W, justify=CENTER,
                                font=('arial narrow', 12, 'normal'), bg='light yellow')
        officeSealLabel.grid(row=9, column=4, padx=10, pady=5, columnspan=2)

        canvas_width, canvas_height = 150, 150
        canvasMem = Canvas(middleFrame, width=canvas_width, height=canvas_height, bg='light yellow')
        myimage = ImageTk.PhotoImage(Image.open("..\\Images\\default_member.jpg").resize((150, 150)))
        canvasMem.create_image(0, 0, anchor=NW, image=myimage)
        canvasMem.grid(row=9, column=3, padx=10, pady=5)

        print_btn = Button(buttonFrame, text="Print", fg="Black",
                           command=None,
                           font=NORM_FONT, width=11, bg='light grey', state=DISABLED, underline=0)

        generate_IDCard_result = partial(self.generate_IDCard, generateId_card_Window, member_Id, memberType,
                                         memberIdText,
                                         memberNameText,
                                         memberDOBText,
                                         memberGenderText,
                                         memberContactNoText,
                                         memberDesignationText,
                                         canvasMem, SEARCH_BY_MEMBERID, print_btn)

        id_button = Button(buttonFrame, text="View ID", fg="Black",
                           command=generate_IDCard_result,
                           font=NORM_FONT, width=11, bg='light cyan', state=NORMAL, underline=0)
        id_button.grid(row=0, column=0)
        print_btn.grid(row=0, column=1)

        # create a Close Button and place into the bookReturn_window window
        cancel = Button(buttonFrame, text="Close", fg="Black", command=generateId_card_Window.destroy,
                        font=NORM_FONT, width=10, bg='light cyan', underline=0)
        cancel.grid(row=0, column=2)

        generateId_card_Window.bind('<Alt-v>', lambda event=None: id_button.invoke())
        generateId_card_Window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        generateId_card_Window.bind('<Alt-p>', lambda event=None: print_btn.invoke())

        generateId_card_Window.focus()
        generateId_card_Window.grab_set()
        mainloop()

    def clearNonMonetaryDonationForm(self, donator_idText,
                                     seva_amountText,
                                     quantityText,
                                     collector_nameText,
                                     cal,
                                     estValueText,
                                     invoice_idText,
                                     infolabel, print_invoice):

        donator_idText.delete(0, END)
        donator_idText.configure(fg='black')
        donator_idText.focus_set()
        seva_amountText.delete(0, END)
        seva_amountText.configure(fg='black')
        collector_nameText.delete(0, END)
        collector_nameText.configure(fg='black')
        estValueText.delete(0, END)
        estValueText.configure(fg='black')
        invoice_idText['text'] = "----------"
        infolabel.configure(text="All fields are mandatory!!", fg='green')
        print_invoice.configure(state=DISABLED, bg='light grey')

    def create_Expanse_Excel(self, createExpanse_window,
                             receiver_IdText,
                             seva_amountText,
                             descriptionText,
                             authorizerText,
                             cal,
                             paymentMode_menu,
                             paymentMode_text,
                             invoice_idText,
                             infolabel, print_invoice, trust_nametext, var, receiver_nameText, receiver_phonenoText):

        current_balance = 0
        print("Radio Selection :", var.get())
        if trust_nametext.get() == "Vihangam Yoga (Karnataka) Trust":
            current_balance = self.obj_commonUtil.readcurrent_balance(VIHANGAM_YOGA_KARNATAKA_TRUST)
        if trust_nametext.get() == "Aadarsh gaushala Trust":
            current_balance = self.obj_commonUtil.readcurrent_balance(SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST)
        print("current_balance : ", current_balance, "in ", trust_nametext.get())
        if int(seva_amountText.get()) <= int(current_balance):
            dateTimeObj = cal.get_date()
            dateOfExpanse_calc = dateTimeObj.strftime("%Y-%m-%d")
            if receiver_IdText.get() == "" or \
                    seva_amountText.get() == "" or \
                    authorizerText.get() == "" or \
                    paymentMode_text.get() == "":

                infolabel.configure(text="All fields are mandatory", fg='red')

            else:
                today = date.today()
                if dateTimeObj <= today:
                    bReceiverIdValid = True  # intilaization required in case of Expanse by Name
                    if var.get() == 1:
                        bReceiverIdValid = self.validate_memberlibraryID_Excel(receiver_IdText.get(), 1)
                    bAuthorizorIdValid = self.validate_memberlibraryID_Excel(authorizerText.get(), 1)

                    if bReceiverIdValid and bAuthorizorIdValid and \
                            (seva_amountText.get()).isnumeric():
                        receiver_data = self.retrieve_MemberRecords_Excel(receiver_IdText.get(), 1,
                                                                          SEARCH_BY_MEMBERID)
                        authorizor_data = self.retrieve_MemberRecords_Excel(authorizerText.get(), 1, SEARCH_BY_MEMBERID)

                        if trust_nametext.get() == "Vihangam Yoga (Karnataka) Trust":
                            file_name_sevaRashi = InitDatabase.getInstance().get_expanse_database_name()  # PATH_EXPANSE_SHEET
                            file_name_transaction = InitDatabase.getInstance().get_transaction_database_name()  # PATH_TRANSACTION_SHEET
                            trust_type = VIHANGAM_YOGA_KARNATAKA_TRUST

                        elif trust_nametext.get() == "Aadarsh gaushala Trust":
                            print("Entered gaushala")
                            file_name_sevaRashi = InitDatabase.getInstance().get_gaushala_expanse_database_name()
                            file_name_transaction = InitDatabase.getInstance().get_gaushala_transaction_database_name()  # PATH_GAUSHALA_TRANSACTION_SHEET
                            trust_type = SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST

                        else:
                            # Not reachable code
                            pass

                        invoice_id = self.obj_commonUtil.generateExpanseVoucherReceiptId(trust_type)
                        # open expanse sheet and enter the data --start
                        print("File name :", file_name_sevaRashi)
                        wb_obj = openpyxl.load_workbook(file_name_sevaRashi)
                        sheet_obj = wb_obj.active
                        total_records = self.totalrecords_excelDataBase(file_name_sevaRashi)

                        if total_records is 0:
                            serial_no = 1
                            row_no = 2

                        else:
                            serial_no = total_records + 1
                            row_no = total_records + 2

                        # sets the general formatting for the new entry in new row
                        for index in range(1, 10):
                            sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                                 bold=False)
                            sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                           vertical='center')

                        sheet_obj.cell(row=row_no, column=1).value = serial_no
                        sheet_obj.cell(row=row_no, column=2).value = seva_amountText.get()
                        sheet_obj.cell(row=row_no, column=3).value = descriptionText.get()
                        if var.get() == 1:
                            sheet_obj.cell(row=row_no, column=4).value = receiver_IdText.get()  # receiver id
                            sheet_obj.cell(row=row_no, column=5).value = receiver_data[2]  # receiver name
                        else:
                            sheet_obj.cell(row=row_no, column=4).value = "NA"
                            sheet_obj.cell(row=row_no, column=5).value = receiver_nameText.get()  # receiver name
                        sheet_obj.cell(row=row_no, column=6).value = dateOfExpanse_calc
                        sheet_obj.cell(row=row_no, column=7).value = authorizerText.get()  # authorizor id
                        sheet_obj.cell(row=row_no, column=8).value = authorizor_data[0]
                        sheet_obj.cell(row=row_no, column=9).value = paymentMode_text.get()
                        sheet_obj.cell(row=row_no, column=10).value = invoice_id

                        wb_obj.save(file_name_sevaRashi)
                        # open expanse sheet and enter the data --End

                        # open transaction sheet and enter the data --start
                        # creating expanse is a debit transaction for the organization
                        transaction_wb_obj = openpyxl.load_workbook(file_name_transaction)
                        transaction_sheet_obj = transaction_wb_obj.active
                        total_records_transaction = self.totalrecords_excelDataBase(file_name_transaction)

                        if total_records_transaction is 0:
                            serial_no = 1
                            row_no = 2
                            balance_amount = 0
                        else:
                            serial_no = total_records_transaction + 1
                            row_no = total_records_transaction + 2
                            balance_amount = int(transaction_sheet_obj.cell(row=row_no - 1, column=9).value)

                        # sets the general formatting for the new entry in new row
                        for index in range(1, 10):
                            transaction_sheet_obj.cell(row=row_no, column=index).font = Font(size=12,
                                                                                             name='Times New Roman',
                                                                                             bold=False)
                            transaction_sheet_obj.cell(row=row_no, column=index).alignment = Alignment(
                                horizontal='center',
                                vertical='center')

                        # new book data is assigned to respective cells in row
                        transaction_sheet_obj.cell(row=row_no, column=1).value = serial_no
                        transaction_sheet_obj.cell(row=row_no, column=2).value = dateOfExpanse_calc
                        transaction_sheet_obj.cell(row=row_no, column=3).value = "---"
                        transaction_sheet_obj.cell(row=row_no, column=4).value = seva_amountText.get()
                        transaction_sheet_obj.cell(row=row_no, column=5).value = descriptionText.get()
                        transaction_sheet_obj.cell(row=row_no, column=6).value = paymentMode_text.get()
                        transaction_sheet_obj.cell(row=row_no, column=7).value = authorizerText.get()  # authorizor id
                        transaction_sheet_obj.cell(row=row_no, column=8).value = authorizor_data[2]
                        transaction_sheet_obj.cell(row=row_no, column=9).value = str(
                            balance_amount - int(seva_amountText.get()))
                        transaction_sheet_obj.cell(row=row_no, column=10).value = invoice_id

                        invoice_idText['text'] = invoice_id
                        transaction_wb_obj.save(file_name_transaction)
                        # open transaction sheet and enter the data --end
                        self.obj_commonUtil.updateExpanseVoucherReceiptBooklet(invoice_id, trust_type)
                        print_result = partial(self.generateExpanseVoucher, receiver_IdText,
                                               seva_amountText,
                                               descriptionText,
                                               authorizerText,
                                               dateOfExpanse_calc,
                                               paymentMode_text,
                                               invoice_id,
                                               receiver_data, trust_nametext,
                                               var,
                                               receiver_nameText,
                                               receiver_phonenoText)
                        print_invoice.configure(state=NORMAL, bg='light cyan', command=print_result)
                        self.submit_deposit.configure(state=DISABLED, bg='light grey')
                        text_withID = "Expanse registered with Invoice  id :" + invoice_id
                        infolabel.configure(text=text_withID, fg='green')

                        # updates the total balance
                        self.obj_commonUtil.calculateTotalAvailableBalance(trust_type)
                    else:
                        infolabel.configure(text="ID/IDs or Amount entered !! please correct", fg='red')
                else:
                    infolabel.configure(text="Future Expanse cannot created !! please correct", fg='red')
        else:
            infolabel.configure(text="Insufficient balance for this expanse !!", fg='red')

    def generate_IDCard(self, generateId_card_Window, member_Id, memberType,
                        memberIdText,
                        memberNameText,
                        memberDOBText,
                        memberGenderText,
                        memberContactNoText,
                        memberDesignationText,
                        canvasMem, search_criteria, print_btn):

        print("printIDCard--> Entry")
        member_data = self.retrieve_MemberRecords_Excel(member_Id.get(), memberType, search_criteria)
        memberIdText['text'] = member_data[0]
        memberNameText['text'] = member_data[2]
        memberDOBText['text'] = member_data[5]
        memberGenderText['text'] = member_data[6]
        memberContactNoText['text'] = member_data[11]
        memberDesignationText['text'] = member_data[22]
        myPhotoimage = ImageTk.PhotoImage(Image.open(member_data[15]).resize((150, 150)))
        canvasMem.create_image(0, 0, anchor=NW, image=myPhotoimage)
        print("Member photo path] :", member_data[15])

        if memberType == 1:
            designation_text = "Member"
            template_path = PATH_TEMPLATE_MEMBERID_CARD
            dest_file = "..\\Member_Data\\ID_Card\\IDs\\" + member_data[0] + "_" + member_data[2] + ".xlsx"
        elif memberType == 2:
            designation_text = "Staff"
            template_path = PATH_TEMPLATE_STAFFID_CARD
            dest_file = "..\\Staff_Data\\ID_Card\\IDs\\" + member_data[0] + "_" + member_data[2] + ".xlsx"
        else:
            pass

        wb_obj = openpyxl.load_workbook(template_path)
        sheet_obj = wb_obj.active
        if memberType == 1:
            designation_text = "Member"
        elif memberType == 2:
            designation_text = "Staff"
        else:
            pass
        sheet_obj.cell(row=2, column=3).value = member_data[2]  # name
        sheet_obj.cell(row=2, column=5).value = member_data[0]  # member id
        sheet_obj.cell(row=3, column=3).value = member_data[5]  # DOB
        sheet_obj.cell(row=3, column=5).value = member_data[6]  # Gender
        sheet_obj.cell(row=4, column=3).value = member_data[11]  # contact number
        sheet_obj.cell(row=4, column=5).value = member_data[22]  # designation

        self.insertPicInIDCard(template_path, member_data[15], wb_obj, sheet_obj, dest_file)

        print_result = partial(self.printIDCard, dest_file)
        print_btn.configure(bg='light cyan', state=NORMAL, command=print_result)

        mainloop()
        print("printIDCard--> Entry")

    def generate_MemberDetails_Form(self, member_Id):

        print("generate_MemberDetails_Form--> Entry")
        member_data = self.retrieve_MemberRecords_Excel(member_Id.get(), 1, 1)

        template_path = PATH_TEMPLATE_MEMBERID_DETAILS
        dest_file = "..\\Member_Data\\Member_Details\\" + member_data[0] + "_" + member_data[2] + ".xlsx"
        copyfile(template_path, dest_file)
        wb_obj = openpyxl.load_workbook(dest_file)
        sheet_obj = wb_obj.active

        sheet_obj.cell(row=3, column=3).value = member_data[2]  # name
        sheet_obj.cell(row=3, column=5).value = member_data[0]  # member id
        sheet_obj.cell(row=4, column=3).value = member_data[5]  # DOB
        sheet_obj.cell(row=4, column=5).value = member_data[1]  # National Identifier
        sheet_obj.cell(row=5, column=3).value = member_data[11]  # contact number
        sheet_obj.cell(row=5, column=5).value = member_data[8]  # Gender

        sheet_obj.cell(row=6, column=3).value = member_data[3]  # Father Name
        sheet_obj.cell(row=6, column=5).value = member_data[4]  # Mother Name
        sheet_obj.cell(row=7, column=3).value = member_data[8]  # City
        sheet_obj.cell(row=7, column=5).value = member_data[9]  # State
        sheet_obj.cell(row=8, column=3).value = member_data[12]  # country
        sheet_obj.cell(row=8, column=5).value = member_data[10]  # pin code

        sheet_obj.cell(row=9, column=3).value = member_data[7]  # Address
        sheet_obj.cell(row=10, column=3).value = member_data[13]  # Nationality
        sheet_obj.cell(row=10, column=5).value = member_data[17]  # id type
        sheet_obj.cell(row=11, column=3).value = member_data[14]  # Email
        sheet_obj.cell(row=12, column=3).value = member_data[19]  # profession
        sheet_obj.cell(row=12, column=5).value = member_data[18]  # member since
        sheet_obj.cell(row=13, column=3).value = member_data[22]  # designation
        sheet_obj.cell(row=13, column=5).value = member_data[21]  # initiated stage

        self.insertPicInMemberDeatails(dest_file, member_data[15], wb_obj, sheet_obj, member_data[0], member_data[2])
        print("generate_MemberDetails_Form--> End")
        mainloop()

    def insertPicInIDCard(self, template_path, imageToInsert, wb_obj, sheet_obj, dest_file):
        img = Image.open(imageToInsert)
        img = img.resize((150, 150), Image.NEAREST)
        img.save('photo.jpg')

        my_png = openpyxl.drawing.image.Image('photo.jpg')
        sheet_obj.add_image(my_png, 'C5')
        wb_obj.save(template_path)
        copyfile(template_path, dest_file)

    def insertPicInMemberDeatails(self, template_path, imageToInsert, wb_obj, sheet_obj, memberID, memberName):
        img = Image.open(imageToInsert)
        img = img.resize((150, 150), Image.NEAREST)
        img.save('photo.jpg')

        my_png = openpyxl.drawing.image.Image('photo.jpg')
        sheet_obj.add_image(my_png, 'C14')
        wb_obj.save(template_path)
        pdf_file = "..\\Member_Data\\Member_Details\\" + memberID + "_" + memberName + ".pdf"
        self.obj_commonUtil.convertExcelToPdf(template_path, pdf_file)

        # remove the temporary excel file
        os.remove(template_path)
        os.startfile(pdf_file, 'print')

    def printIDCard(self, dest_file):
        os.startfile(dest_file)

    def display_data(self, memberType, search_criteria):
        display_dataWindow = Toplevel(self.master)
        if memberType == 1:
            headingForm = "Displaying Member Data"
            display_dataWindow.title("Member Information Details ")
        elif memberType == 2:
            headingForm = "Staff Data Display"
            display_dataWindow.title("Staff Information Details")
        else:
            pass
        display_dataWindow.geometry('730x620+250+50')
        display_dataWindow.configure(background='wheat')
        display_dataWindow.resizable(width=True, height=True)

        if search_criteria == SEARCH_BY_MEMBERID:
            text_search = "Member Id"
        elif search_criteria == SEARCH_BY_CONTACTNO:
            text_search = "Contact No."
        else:
            pass
        upperFrame = Frame(display_dataWindow, width=205, height=100, bd=4, relief='ridge', bg='light yellow')
        upperFrame.grid(row=1, column=2, padx=20, pady=10, sticky=W)
        heading = Label(display_dataWindow, text=headingForm, font=('times new roman', 25, 'normal'), bg='wheat')
        heading.grid(row=0, column=1, columnspan=3)
        middleFrame = Frame(display_dataWindow, width=200, height=300, bd=8, relief='ridge', bg='light yellow')
        middleFrame.grid(row=2, column=2, padx=20, pady=10, sticky=W)

        memberIdLabel = Label(upperFrame, text=text_search, width=9, anchor=W, justify=LEFT,
                              font=('arial narrow', 15, 'normal'), bg='light yellow')
        memberIdLabel.grid(row=1, column=0, padx=30, pady=10)
        member_Id = Entry(upperFrame, width=7, font=('arial narrow', 15, 'normal'), justify='center')
        member_Id.grid(row=1, column=1, ipadx=60, pady=10)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(upperFrame, width=200, height=100, bd=4, relief='ridge', bg='light yellow')
        buttonFrame.grid(row=1, column=3, padx=15, pady=10, sticky=W)

        # ---------------------------------Button Frame End----------------------------------------

        # ---------------------------------Preparing display Area - start ---------------------------------

        # Display Member Name - Row 1
        memberNameLabel = Label(middleFrame, text="Name", width=10, anchor=W, justify=LEFT,
                                font=('arial narrow', 12, 'normal'),
                                bg='light yellow')
        memberNameLabel.grid(row=1, column=2, padx=0, pady=5)
        memberNameText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                               font=('arial narrow', 12, 'normal'),
                               bg='light yellow')
        memberNameText.grid(row=1, column=3, padx=5, pady=5)

        # Display Member Id - Row 1
        memberIdLabel = Label(middleFrame, text="Member Id", width=10, anchor=W, justify=LEFT,
                              font=('arial narrow', 12, 'normal'),
                              bg='light yellow')
        memberIdLabel.grid(row=1, column=4, padx=10, pady=5)
        memberIdText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                             font=('arial narrow', 12, 'normal'),
                             bg='light yellow')
        memberIdText.grid(row=1, column=5, padx=5, pady=5)

        # Display Father name - Row 2
        memberFatherLabel = Label(middleFrame, text="Father Name", width=10, anchor=W, justify=LEFT,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='light yellow')
        memberFatherLabel.grid(row=2, column=2, padx=10, pady=5)
        memberFatherText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                 font=('arial narrow', 12, 'normal'),
                                 bg='light yellow')
        memberFatherText.grid(row=2, column=3, padx=5, pady=5)

        # Display Mother Name - Row 2
        memberMotherLabel = Label(middleFrame, text="Mother Name", width=10, anchor=W, justify=LEFT,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='light yellow')
        memberMotherLabel.grid(row=2, column=4, padx=10, pady=5)
        memberMotherText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                 font=('arial narrow', 12, 'normal'),
                                 bg='light yellow')
        memberMotherText.grid(row=2, column=5, padx=5, pady=5)

        # Display Father name - Row 3
        memberDOBLabel = Label(middleFrame, text="Date Of Birth", width=10, anchor=W, justify=LEFT,
                               font=('arial narrow', 12, 'normal'),
                               bg='light yellow')
        memberDOBLabel.grid(row=3, column=2, padx=10, pady=5)
        memberDOBText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                              font=('arial narrow', 12, 'normal'),
                              bg='light yellow')
        memberDOBText.grid(row=3, column=3, padx=5, pady=5)

        # Display Mother Name - Row 3
        memberGenderLabel = Label(middleFrame, text="Gender", width=10, anchor=W, justify=LEFT,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='light yellow')
        memberGenderLabel.grid(row=3, column=4, padx=10, pady=5)
        memberGenderText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                 font=('arial narrow', 12, 'normal'),
                                 bg='light yellow')
        memberGenderText.grid(row=3, column=5, padx=5, pady=5)

        # Display Member Name - Row 4
        memberContactLabel = Label(middleFrame, text="Contact No", width=10, anchor=W, justify=LEFT,
                                   font=('arial narrow', 12, 'normal'),
                                   bg='light yellow')
        memberContactLabel.grid(row=4, column=2, padx=0, pady=5)
        memberContactNoText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                    font=('arial narrow', 12, 'normal'),
                                    bg='light yellow')
        memberContactNoText.grid(row=4, column=3, padx=5, pady=5)

        # Display Member Id - Row 4
        memberCityLabel = Label(middleFrame, text="City", width=10, anchor=W, justify=LEFT,
                                font=('arial narrow', 12, 'normal'),
                                bg='light yellow')
        memberCityLabel.grid(row=4, column=4, padx=10, pady=5)
        memberCityText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                               font=('arial narrow', 12, 'normal'),
                               bg='light yellow')
        memberCityText.grid(row=4, column=5, padx=5, pady=5)

        # Display Father name - Row 5
        memberStateLabel = Label(middleFrame, text="State", width=10, anchor=W, justify=LEFT,
                                 font=('arial narrow', 12, 'normal'),
                                 bg='light yellow')
        memberStateLabel.grid(row=5, column=2, padx=10, pady=5)
        memberStateText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                font=('arial narrow', 12, 'normal'),
                                bg='light yellow')
        memberStateText.grid(row=5, column=3, padx=5, pady=5)

        # Display Country Name - Row 5
        memberCountryLabel = Label(middleFrame, text="Country", width=10, anchor=W, justify=LEFT,
                                   font=('arial narrow', 12, 'normal'),
                                   bg='light yellow')
        memberCountryLabel.grid(row=5, column=4, padx=10, pady=5)
        memberCountryText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='light yellow')
        memberCountryText.grid(row=5, column=5, padx=5, pady=5)

        # Display Address - Row 6
        memberAddressLabel = Label(middleFrame, text="Address", width=10, anchor=W, justify=LEFT,
                                   font=('arial narrow', 12, 'normal'),
                                   bg='light yellow')
        memberAddressLabel.grid(row=6, column=2, padx=10, pady=5)
        memberAddressText = Label(middleFrame, text="", width=65, anchor=W, justify=LEFT,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='light yellow')
        memberAddressText.grid(row=6, column=3, columnspan=3, ipadx=10, padx=5, pady=5)

        # Display Father name - Row 7
        memberPinCodeLabel = Label(middleFrame, text="Pin-Code", width=10, anchor=W, justify=LEFT,
                                   font=('arial narrow', 12, 'normal'),
                                   bg='light yellow')
        memberPinCodeLabel.grid(row=7, column=2, padx=10, pady=5)
        memberPinCodeText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='light yellow')
        memberPinCodeText.grid(row=7, column=3, padx=5, pady=5)

        # Display Country Name - Row 7
        memberNationalityLabel = Label(middleFrame, text="Nationality", width=10, anchor=W, justify=LEFT,
                                       font=('arial narrow', 12, 'normal'),
                                       bg='light yellow')
        memberNationalityLabel.grid(row=7, column=4, padx=10, pady=5)
        memberNationalityText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                      font=('arial narrow', 12, 'normal'),
                                      bg='light yellow')
        memberNationalityText.grid(row=7, column=5, padx=5, pady=5)

        # Display Father name - Row 7
        memberEmailLabel = Label(middleFrame, text="Email-Id", width=10, anchor=W, justify=LEFT,
                                 font=('arial narrow', 12, 'normal'),
                                 bg='light yellow')
        memberEmailLabel.grid(row=8, column=2, padx=10, pady=5)
        memberEmailText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                font=('arial narrow', 12, 'normal'),
                                bg='light yellow')
        memberEmailText.grid(row=8, column=3, padx=5, pady=5)

        # Display Country Name - Row 8
        memberIdTypeLabel = Label(middleFrame, text="ID Type", width=10, anchor=W, justify=LEFT,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='light yellow')
        memberIdTypeLabel.grid(row=8, column=4, padx=10, pady=5)
        memberIdTypeText = Label(middleFrame, text="", width=25, anchor=W, justify=LEFT,
                                 font=('arial narrow', 12, 'normal'),
                                 bg='light yellow')
        memberIdTypeText.grid(row=8, column=5, padx=5, pady=5)

        # Display Photo - Row 9
        memberPhotoLabel = Label(middleFrame, text="Member Photo", width=11, anchor=W, justify=LEFT,
                                 font=('arial narrow', 12, 'normal'),
                                 bg='light yellow')
        memberPhotoLabel.grid(row=9, column=2, padx=10, pady=5)

        canvas_width, canvas_height = 150, 150
        canvasMem = Canvas(middleFrame, width=canvas_width, height=canvas_height, bg='light yellow')
        myimage = ImageTk.PhotoImage(Image.open("..\\Images\\default_member.jpg").resize((150, 150)))
        canvasMem.create_image(0, 0, anchor=NW, image=myimage)
        canvasMem.grid(row=9, column=3, padx=10, pady=5)

        memberPhotoIDLabel = Label(middleFrame, text="ID Photo", width=11, anchor=W, justify=LEFT,
                                   font=('arial narrow', 12, 'normal'),
                                   bg='light yellow')
        memberPhotoIDLabel.grid(row=9, column=4, padx=10, pady=5)

        canvas_width, canvas_height = 150, 150
        canvasMemID = Canvas(middleFrame, width=canvas_width, height=canvas_height, bg='light yellow')
        Idimage = ImageTk.PhotoImage(Image.open("..\\Images\\default_idcard.jpg").resize((150, 150)))
        canvasMemID.create_image(0, 0, anchor=NW, image=Idimage)
        canvasMemID.grid(row=9, column=5, padx=10, pady=5)

        print_button = Button(buttonFrame, text="Print", fg="Black",
                              command=None,
                              font=NORM_FONT, width=10, bg='light grey', state=DISABLED, underline=0)

        search_member = partial(self.assignDataForDisplay, display_dataWindow, memberType, member_Id,
                                memberIdText,
                                memberNameText,
                                memberFatherText,
                                memberMotherText,
                                memberDOBText,
                                memberGenderText,
                                memberContactNoText,
                                memberCityText,
                                memberStateText,
                                memberAddressText,
                                memberNationalityText,
                                memberCountryText,
                                memberPinCodeText,
                                memberIdTypeText,
                                memberEmailText,
                                canvasMem,
                                canvasMemID,
                                search_criteria, print_button)

        # create a Search Button and place into the bookReturn_window window
        submit = Button(buttonFrame, text="Search", fg="Black", command=search_member,
                        font=NORM_FONT, width=10, bg='light cyan', underline=0)
        submit.grid(row=0, column=0)

        # create a Close Button and place into the bookReturn_window window

        print_button.grid(row=0, column=1)

        # create a Close Button and place into the bookReturn_window window
        cancel = Button(buttonFrame, text="Close", fg="Black", command=display_dataWindow.destroy,
                        font=NORM_FONT, width=10, bg='light cyan', underline=0)
        cancel.grid(row=0, column=2)

        display_dataWindow.bind('<Return>', lambda event=None: submit.invoke())
        display_dataWindow.bind('<Alt-c>', lambda event=None: cancel.invoke())
        display_dataWindow.bind('<Alt-p>', lambda event=None: print_button.invoke())

        display_dataWindow.focus()
        display_dataWindow.grab_set()
        mainloop()

    def edit_member_data(self, memberType, search_criteria):
        display_dataWindow = Toplevel(self.master)
        if memberType == 1:
            headingForm = "Edit Member Details"
            display_dataWindow.title("Edit Information Details ")
        elif memberType == 2:
            headingForm = "Staff Data Display"
            display_dataWindow.title("Staff Information Details")
        else:
            pass
        display_dataWindow.geometry('700x500+250+150')
        display_dataWindow.configure(background='wheat')
        display_dataWindow.resizable(width=True, height=True)

        if search_criteria == SEARCH_BY_MEMBERID:
            text_search = "Member Id"
        elif search_criteria == SEARCH_BY_CONTACTNO:
            text_search = "Contact No."
        else:
            pass

        heading = Label(display_dataWindow, text=headingForm, font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        heading.grid(row=0, column=0, columnspan=4)
        upperFrame = Frame(display_dataWindow, width=205, height=100, bd=8, relief='ridge', bg='light yellow')
        upperFrame.grid(row=1, column=2, padx=20, pady=10, sticky=W)

        middleFrame = Frame(display_dataWindow, width=200, height=300, bd=8, relief='ridge', bg='light yellow')
        middleFrame.grid(row=2, column=2, padx=20, pady=10, sticky=W)

        infoFrame = Frame(display_dataWindow, width=200, height=100, bd=8, relief='ridge', bg='light yellow')
        infoFrame.grid(row=16, column=2, padx=80, pady=10, columnspan=5, sticky=W)

        memberIdLabel = Label(upperFrame, text=text_search, width=9, anchor=W, justify=LEFT,
                              font=('arial narrow', 13, 'normal'), bg='light yellow')
        memberIdLabel.grid(row=1, column=0, padx=30, pady=10)
        member_Id = Entry(upperFrame, width=25, font=('arial narrow', 13, 'normal'), justify='center')
        member_Id.grid(row=1, column=1, pady=10)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(upperFrame, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=1, column=3, padx=15, pady=10, sticky=W)

        # ---------------------------------Button Frame End----------------------------------------

        # ---------------------------------Preparing display Area - start ---------------------------------

        # Display Member Name - Row 4
        contactNo_text = StringVar(middleFrame)
        memberContactLabel = Label(middleFrame, text="Contact No", width=12, anchor=W, justify=LEFT,
                                   font=('arial narrow', 12, 'normal'),
                                   bg='light yellow')
        memberContactLabel.grid(row=4, column=2, padx=0, pady=5)
        memberContactNoText = Entry(middleFrame, textvariable=contactNo_text, width=25, justify=LEFT,
                                    font=('arial narrow', 12, 'normal'),
                                    bg='snow')
        memberContactNoText.grid(row=4, column=3, padx=5, pady=5)

        # Display Member Id - Row 4
        city_text = StringVar(middleFrame)
        memberCityLabel = Label(middleFrame, text="City", width=12, anchor=W, justify=LEFT,
                                font=('arial narrow', 12, 'normal'),
                                bg='light yellow')
        memberCityLabel.grid(row=4, column=4, padx=10, pady=5)
        memberCityText = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=city_text,
                               font=('arial narrow', 12, 'normal'),
                               bg='snow')
        memberCityText.grid(row=4, column=5, padx=5, pady=5)

        # Display Father name - Row 5
        state_text = StringVar(middleFrame)
        memberStateLabel = Label(middleFrame, text="State", width=12, anchor=W, justify=LEFT,
                                 font=('arial narrow', 12, 'normal'),
                                 bg='light yellow')
        memberStateLabel.grid(row=5, column=2, padx=10, pady=5)
        memberStateText = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=state_text,
                                font=('arial narrow', 12, 'normal'),
                                bg='snow')
        memberStateText.grid(row=5, column=3, padx=5, pady=5)

        # Display Country Name - Row 5
        country_text = StringVar(middleFrame)
        memberCountryLabel = Label(middleFrame, text="Country", width=12, anchor=W, justify=LEFT,
                                   font=('arial narrow', 12, 'normal'),
                                   bg='light yellow')
        memberCountryLabel.grid(row=5, column=4, padx=10, pady=5)
        memberCountryText = Entry(middleFrame, text="", textvariable=country_text, width=25, justify=LEFT,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='snow')
        memberCountryText.grid(row=5, column=5, padx=5, pady=5)

        # Display Address - Row 6
        address_text = StringVar(middleFrame)
        memberAddressLabel = Label(middleFrame, text="Address", width=12, anchor=W, justify=LEFT,
                                   font=('arial narrow', 12, 'normal'),
                                   bg='light yellow')
        memberAddressLabel.grid(row=6, column=2, padx=10, pady=5)
        memberAddressText = Entry(middleFrame, text="", width=71, justify=LEFT, textvariable=address_text,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='snow')
        memberAddressText.grid(row=6, column=3, columnspan=3, padx=5, pady=5)

        # Display Father name - Row 7
        pincode_text = StringVar(middleFrame)
        memberPinCodeLabel = Label(middleFrame, text="Pin-Code", width=12, anchor=W, justify=LEFT,
                                   font=('arial narrow', 12, 'normal'),
                                   bg='light yellow')
        memberPinCodeLabel.grid(row=7, column=2, padx=10, pady=5)
        memberPinCodeText = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=pincode_text,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='snow')
        memberPinCodeText.grid(row=7, column=3, padx=5, pady=5)

        # Display Father name - Row 7
        email_text = StringVar(middleFrame)
        memberEmailLabel = Label(middleFrame, text="Email-Id", width=12, anchor=W, justify=LEFT,
                                 font=('arial narrow', 12, 'normal'),
                                 bg='light yellow')
        memberEmailLabel.grid(row=8, column=2, padx=10, pady=5)
        memberEmailText = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=email_text,
                                font=('arial narrow', 12, 'normal'),
                                bg='snow')
        memberEmailText.grid(row=8, column=3, padx=5, pady=5)

        profession_text = StringVar(middleFrame)
        professionLabel = Label(middleFrame, text="Profession", width=12, anchor=W, justify=LEFT,
                                font=('arial narrow', 12, 'normal'),
                                bg='light yellow')
        professionLabel.grid(row=7, column=4, padx=10, pady=5)
        professionText = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=profession_text,
                               font=('arial narrow', 12, 'normal'),
                               bg='snow')
        professionText.grid(row=7, column=5, padx=5, pady=5)

        designation_varaible = StringVar(middleFrame)
        designationLabel = Label(middleFrame, text="Designation", width=12, anchor=W, justify=LEFT,
                                 font=('arial narrow', 12, 'normal'),
                                 bg='light yellow')
        designationLabel.grid(row=8, column=4, padx=10, pady=5)
        designation_varaible.set("Other")
        designation_Type = OptionMenu(middleFrame, designation_varaible, "Updestha", "President", "Vice-President",
                                      "Treasurer", "Trustee", "Manager",
                                      "Member", "Staff-Sevak", "Area Co-ordinator",
                                      "Other")
        designation_Type.configure(bg='white', width=17, fg='black', font=('verdana', 9, 'normal'), anchor=W,
                                   justify=LEFT)
        designation_Type.grid(row=8, column=5, padx=10, pady=5)

        akshyaAvailable_varaible = StringVar(middleFrame)
        akshayAvailableLabel = Label(middleFrame, text="Akshay Patra?", width=12, anchor=W, justify=LEFT,
                                     font=('arial narrow', 12, 'normal'),
                                     bg='light yellow')
        akshayAvailableLabel.grid(row=9, column=2, padx=10, pady=5)
        akshyaAvailable_varaible.set("Other")
        akshayAvailable_Type = OptionMenu(middleFrame, akshyaAvailable_varaible, "Yes", "No")
        akshayAvailable_Type.configure(bg='white', width=17, fg='black', font=('verdana', 9, 'normal'), anchor=W,
                                       justify=LEFT)
        akshayAvailable_Type.grid(row=9, column=3, padx=10, pady=5)

        akshayNo_text = StringVar(middleFrame)
        akshayboxLabel = Label(middleFrame, text="Akshay Box No.", width=12, anchor=W, justify=LEFT,
                               font=('arial narrow', 12, 'normal'),
                               bg='light yellow')
        akshayboxLabel.grid(row=9, column=4, padx=10, pady=5)
        akshayboxnoText = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=akshayNo_text,
                                font=('arial narrow', 12, 'normal'),
                                bg='snow')
        akshayboxnoText.grid(row=9, column=5, padx=5, pady=5)

        isPatrikaSubsc_varaible = StringVar(middleFrame)
        patrika_subscLabel = Label(middleFrame, text="Magazine Subs.?", width=12, anchor=W, justify=LEFT,
                                   font=('arial narrow', 12, 'normal'),
                                   bg='light yellow')
        patrika_subscLabel.grid(row=10, column=2, padx=10, pady=5)
        isPatrikaSubsc_varaible.set("Yes")
        patrilaSubs_Type = OptionMenu(middleFrame, isPatrikaSubsc_varaible, "Yes", "No")
        patrilaSubs_Type.configure(bg='white', width=17, fg='black', font=('verdana', 9, 'normal'), anchor=W,
                                   justify=LEFT)
        patrilaSubs_Type.grid(row=10, column=3, padx=10, pady=5)

        infoLabel = Label(infoFrame, text="Press Save button to save the modified records", width=60, anchor='center',
                          justify=CENTER,
                          font=('arial narrow', 13, 'normal'),
                          bg='light yellow')
        save_button = Button(buttonFrame, text="Save", fg="Black",
                             command=None,
                             font=NORM_FONT, width=9, bg='grey', state=DISABLED, underline=0)
        search_member = partial(self.assignDataForDisplay_editMemberInfo, display_dataWindow, memberType, member_Id,
                                memberContactNoText,
                                memberCityText,
                                memberStateText,
                                memberAddressText,
                                memberCountryText,
                                memberPinCodeText,
                                memberEmailText,
                                professionText,
                                designation_varaible,
                                akshyaAvailable_varaible,
                                akshayboxnoText,
                                isPatrikaSubsc_varaible,
                                search_criteria, infoLabel, save_button)

        # create a Search Button and place into the bookReturn_window window
        submit = Button(buttonFrame, text="Search", fg="Black", command=search_member,
                        font=NORM_FONT, width=9, bg='light cyan', underline=0)
        submit.grid(row=0, column=0)
        save_result = partial(self.saveModifiedMemberRecords, member_Id,
                              memberContactNoText,
                              memberCityText,
                              memberStateText,
                              memberAddressText,
                              memberCountryText,
                              memberPinCodeText,
                              memberEmailText,
                              professionText,
                              designation_varaible,
                              akshyaAvailable_varaible,
                              akshayboxnoText,
                              isPatrikaSubsc_varaible,
                              search_criteria, infoLabel)
        save_button.configure(command=save_result)
        save_button.grid(row=0, column=1)

        # create a Close Button and place into the bookReturn_window window
        cancel = Button(buttonFrame, text="Close", fg="Black", command=display_dataWindow.destroy,
                        font=NORM_FONT, width=9, bg='light cyan', underline=0)
        cancel.grid(row=0, column=2)

        # ---------------------------Button frame ends

        infoLabel.grid(row=16, column=1, padx=10, pady=5)

        display_dataWindow.bind('<Return>', lambda event=None: submit.invoke())
        display_dataWindow.bind('<Alt-c>', lambda event=None: cancel.invoke())
        display_dataWindow.bind('<Alt-r>', lambda event=None: self.print_button.invoke())

        display_dataWindow.focus()
        display_dataWindow.grab_set()
        mainloop()

    def edit_commercialItem_data(self):
        display_dataWindow = Toplevel(self.master)

        headingForm = "Edit Book\Sukrit Product Details"
        display_dataWindow.title("Edit Information Details ")

        display_dataWindow.geometry('700x425+250+150')
        display_dataWindow.configure(background='wheat')
        display_dataWindow.resizable(width=True, height=True)

        heading = Label(display_dataWindow, text=headingForm, font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        heading.grid(row=0, column=0, columnspan=4)
        upperFrame = Frame(display_dataWindow, width=205, height=100, bd=8, relief='ridge', bg='light yellow')
        upperFrame.grid(row=1, column=2, padx=30, pady=10, sticky=W)

        middleFrame = Frame(display_dataWindow, width=200, height=300, bd=8, relief='ridge', bg='light yellow')
        middleFrame.grid(row=2, column=2, padx=20, pady=10, sticky=W)

        infoFrame = Frame(display_dataWindow, width=200, height=100, bd=8, relief='ridge', bg='light yellow')
        infoFrame.grid(row=16, column=2, padx=80, pady=10, columnspan=5, sticky=W)

        itemIdLabel = Label(upperFrame, text="Item Id CI- ", width=9, anchor=W, justify=LEFT,
                            font=('arial narrow', 13, 'normal'), bg='light yellow')
        itemIdLabel.grid(row=1, column=0, padx=10, pady=10)
        item_Id = Entry(upperFrame, width=25, font=('arial narrow', 13, 'normal'), justify='center')
        item_Id.grid(row=1, column=1, pady=10)

        centerNameLabel = Label(upperFrame, text="Center Name", width=10, anchor=W, justify=LEFT,
                                font=('arial narrow', 12, 'normal'), bg='light yellow')
        centerNameLabel.grid(row=2, column=0, padx=10)
        local_centerText = StringVar(upperFrame)
        localCenterList = self.obj_commonUtil.getLocalCenterNames()
        print("Center list  - ", localCenterList)
        local_centerText.set(localCenterList[0])
        localcenter_menu = OptionMenu(upperFrame, local_centerText, *localCenterList)
        localcenter_menu.configure(width=50, font=('arial narrow', 12, 'normal'), bg='snow', anchor=W, justify=LEFT)
        localcenter_menu.grid(row=2, column=1, padx=5,columnspan = 3)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(upperFrame, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=1, column=3, padx=22, pady=15, sticky=W)

        # ---------------------------------Button Frame End----------------------------------------

        # ---------------------------------Preparing display Area - start ---------------------------------

        # Display item Id - Row 4
        itemnametext = StringVar(middleFrame)
        itemnamelabel = Label(middleFrame, text="Item Name", width=12, anchor=W, justify=LEFT,
                              font=('arial narrow', 12, 'normal'),
                              bg='light yellow')
        itemnamelabel.grid(row=4, column=2, padx=10, pady=5)
        itemname_Text = Entry(middleFrame, text="", width=73, justify=LEFT, textvariable=itemnametext,
                              font=('arial narrow', 12, 'normal'),
                              bg='snow')
        itemname_Text.grid(row=4, column=3, padx=5, pady=5, columnspan=4)

        borrowfeetext = StringVar(middleFrame)
        borrowfeeLabel = Label(middleFrame, text="Borrow Fee(Rs.)", width=12, anchor=W, justify=LEFT,
                               font=('arial narrow', 12, 'normal'),
                               bg='light yellow')
        borrowfeeLabel.grid(row=5, column=2, padx=10, pady=5)
        borrowFee_Text = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=borrowfeetext,
                               font=('arial narrow', 12, 'normal'),
                               bg='snow')
        borrowFee_Text.grid(row=5, column=3, padx=5, pady=5)

        unitpricetext = StringVar(middleFrame)
        unitpriceLabel = Label(middleFrame, text="Unit Price(Rs.)", width=12, anchor=W, justify=LEFT,
                               font=('arial narrow', 12, 'normal'),
                               bg='light yellow')
        unitpriceLabel.grid(row=5, column=4, padx=10, pady=5)
        unitpriceText = Entry(middleFrame, text="", textvariable=unitpricetext, width=25, justify=LEFT,
                              font=('arial narrow', 12, 'normal'),
                              bg='snow')
        unitpriceText.grid(row=5, column=5, padx=5, pady=5)

        racktext = StringVar(middleFrame)
        rackno_label = Label(middleFrame, text="Rack No.", width=12, anchor=W, justify=LEFT,
                             font=('arial narrow', 12, 'normal'),
                             bg='light yellow')
        rackno_label.grid(row=6, column=2, padx=10, pady=5)
        rackno_Text = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=racktext,
                            font=('arial narrow', 12, 'normal'),
                            bg='snow')
        rackno_Text.grid(row=6, column=3, padx=5, pady=5)

        quantitytext = StringVar(middleFrame)
        itemquantityLabel = Label(middleFrame, text="Quantity", width=12, anchor=W, justify=LEFT,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='light yellow')
        itemquantityLabel.grid(row=6, column=4, padx=10, pady=5)
        itemQuantity_Text = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=quantitytext,
                                  font=('arial narrow', 12, 'normal'),
                                  bg='snow')
        itemQuantity_Text.grid(row=6, column=5, padx=5, pady=5)
        itemauthorLabel = Label(middleFrame, text="Author Name", width=12, anchor=W, justify=LEFT,
                                font=('arial narrow', 12, 'normal'),
                                bg='light yellow')
        itemauthorLabel.grid(row=7, column=2, padx=10, pady=5)

        authorText = StringVar(middleFrame)
        authorList = self.obj_commonUtil.getAuthorNames()
        print("Author list  - ", authorList)
        authorText.set(authorList[0])
        author_menu = OptionMenu(middleFrame, authorText, *authorList)
        author_menu.configure(width=67, font=('arial narrow', 12, 'normal'), bg='snow', anchor=W, justify=LEFT)
        author_menu.grid(row=7, column=3, padx=10, pady=5, columnspan=3)
        infoLabel = Label(infoFrame, text="Press Save button to save the modified records", width=60, anchor='center',
                          justify=CENTER,
                          font=('arial narrow', 13, 'normal'),
                          bg='light yellow')

        search_item = partial(self.assignDataForDisplay_editCommercialItemInfo, display_dataWindow, item_Id,
                              itemname_Text,
                              authorText,
                              unitpriceText,
                              itemQuantity_Text,
                              borrowFee_Text,
                              rackno_Text,
                              infoLabel, local_centerText)

        # create a Search Button and place into the bookReturn_window window
        submit = Button(buttonFrame, text="Search", fg="Black", command=search_item,
                        font=NORM_FONT, width=9, bg='light cyan', underline=0)
        submit.grid(row=0, column=0)

        # create a Close Button and place into the bookReturn_window window

        save_result = partial(self.saveModifiedCommercialItemRecords, display_dataWindow, item_Id,
                              itemname_Text,
                              authorText,
                              unitpriceText,
                              borrowFee_Text,
                              itemQuantity_Text,
                              rackno_Text,
                              infoLabel, local_centerText)

        self.print_button = Button(buttonFrame, text="Save", fg="Black",
                                   command=save_result,
                                   font=NORM_FONT, width=9, bg='grey', state=DISABLED, underline=0)
        self.print_button.grid(row=0, column=1)

        # create a Close Button and place into the bookReturn_window window
        cancel = Button(buttonFrame, text="Close", fg="Black", command=display_dataWindow.destroy,
                        font=NORM_FONT, width=9, bg='light cyan', underline=0)
        cancel.grid(row=0, column=2)

        # ---------------------------Button frame ends

        infoLabel.grid(row=16, column=1, padx=10, pady=3)

        display_dataWindow.bind('<Return>', lambda event=None: submit.invoke())
        display_dataWindow.bind('<Alt-c>', lambda event=None: cancel.invoke())
        display_dataWindow.bind('<Alt-r>', lambda event=None: self.print_button.invoke())

        display_dataWindow.focus()
        display_dataWindow.grab_set()
        mainloop()

    def edit_noncommercialItem_data(self):
        obj_noncommerEdt = NoncommercialEdit(root)

    def view_monetarydonation_statement(self):
        obj_mondonactstatement = MonetarydonationStatement(root)

    def resetDatabase(self):
        self.obj_initDatabase.resetallExpanseData()

    def view_memberContribution(self):
        obj_memberContribution = MemberContribution(root)

    def view_stocksales_statement(self):
        obj_stocksalesstatement = StocksalesStatement(root)

    def view_main_account_statement(self):
        obj_accountstatement = AccountStatement(root)

    def view_gaushala_account_statement(self):
        obj_gaushalaaccountstatement = GaushalaAccountStatement(root)

    def import_database(self):
        obj_importDatabase = ImportDatabase(root)

    def view_split_donation_window(self):
        self.obj_splitdonation_window.split_donation_view(root)

    def view_split_donation_list(self):
        self.obj_splitdonation_window.split_donation_list_view(root)

    def view_pledge_statement_by_duration(self):
        obj_pledgeStatement = PledgeStatement(root)
        obj_pledgeStatement.pledge_byduration_statement_window()

    def view_pledge_statement_by_item(self):
        obj_pledgeStatement = PledgeStatement(root)
        obj_pledgeStatement.pledge_byItem_statement_window()

    def view_pledge_statement_by_member(self):
        obj_pledgeStatement = PledgeStatement(root)
        obj_pledgeStatement.pledge_byMember_statement_window()

    def view_stock_info(self):
        self.objStock_info.view_stock_info(root)

    # --------------------------------Preparing display area - end -----------------------------------

    def add_member(self, memberType):
        addMember_window = Toplevel(self.master)
        if memberType == 1:
            headingForm = "New Member Registration"
            addMember_window.title("Member Registration ")
        elif memberType == 2:
            headingForm = "Department Staff Registration"
            addMember_window.title("Department Staff Registration")

        addMember_window.geometry('900x700+220+90')
        addMember_window.configure(background='wheat')
        addMember_window.resizable(width=True, height=True)

        heading = Label(addMember_window, text=headingForm, font=('ariel narrow', 15, 'bold'),
                        bg='wheat')

        dataEntryFrame = Frame(addMember_window, width=300, height=50, bd=4, relief='ridge', bg='light yellow')
        id_number = Label(dataEntryFrame, text="Govt. Identifier", width=13, anchor=W, justify=LEFT,
                          font=NORM_VERDANA_FONT,
                          bg='light yellow')
        name = Label(dataEntryFrame, text="Member Name", width=13, anchor=W, justify=LEFT, font=NORM_VERDANA_FONT,
                     bg='light yellow', fg="red")
        dateOfBirth = Label(dataEntryFrame, text="Date of Birth", width=13, anchor=W, justify=LEFT,
                            font=NORM_VERDANA_FONT, bg='light yellow', fg="red")
        gender = Label(dataEntryFrame, text="Gender", width=13, anchor=W, justify=LEFT,
                       font=NORM_VERDANA_FONT, bg='light yellow', fg="red")
        father_name = Label(dataEntryFrame, text="Father Name", width=13, anchor=W, justify=LEFT,
                            font=NORM_VERDANA_FONT,
                            bg='light yellow', fg="red")
        mother_name = Label(dataEntryFrame, text="Mother Name", width=13, anchor=W, justify=LEFT,
                            font=NORM_VERDANA_FONT,
                            bg='light yellow', fg="red")
        address = Label(dataEntryFrame, text="Address", width=13, anchor=W, justify=LEFT,
                        font=NORM_VERDANA_FONT, bg='light yellow', fg="red")
        residentCity = Label(dataEntryFrame, text="City", width=13, anchor=W, justify=LEFT,
                             font=NORM_VERDANA_FONT, bg='light yellow', fg="red")
        residentState = Label(dataEntryFrame, text="State", width=13, anchor=W, justify=LEFT,
                              font=NORM_VERDANA_FONT, bg='light yellow')
        residentCountry = Label(dataEntryFrame, text="Country", width=13, anchor=W, justify=LEFT,
                                font=NORM_VERDANA_FONT, bg='light yellow')
        residentPincode = Label(dataEntryFrame, text="Pin Code", width=13, anchor=W, justify=LEFT,
                                font=NORM_VERDANA_FONT, bg='light yellow')
        contactNo = Label(dataEntryFrame, text="Contact No.", width=13, anchor=W, justify=LEFT,
                          font=NORM_VERDANA_FONT, bg='light yellow')
        nationality = Label(dataEntryFrame, text="Nationality", width=13, anchor=W, justify=LEFT,
                            font=NORM_VERDANA_FONT, bg='light yellow')
        emailId = Label(dataEntryFrame, text="Email Id", width=13, anchor=W, justify=LEFT,
                        font=NORM_VERDANA_FONT, bg='light yellow')
        residentIdType = Label(dataEntryFrame, text="Id Type", width=13, anchor=W, justify=LEFT,
                               font=NORM_VERDANA_FONT, bg='light yellow')
        associatedSince = Label(dataEntryFrame, text="Member Since", width=13, anchor=W, justify=LEFT,
                                font=NORM_VERDANA_FONT, bg='light yellow')

        profession = Label(dataEntryFrame, text="Profession", width=13, anchor=W, justify=LEFT,
                           font=NORM_VERDANA_FONT, bg='light yellow')
        updeshBy = Label(dataEntryFrame, text="Updesh by", width=13, anchor=W, justify=LEFT,
                         font=NORM_VERDANA_FONT, bg='light yellow')

        initiated_stage = Label(dataEntryFrame, text="Initiated Stage", width=13, anchor=W, justify=LEFT,
                                font=NORM_VERDANA_FONT, bg='light yellow')
        designation = Label(dataEntryFrame, text="Designation", width=13, anchor=W, justify=LEFT,
                            font=NORM_VERDANA_FONT, bg='light yellow')
        akshayPatra = Label(dataEntryFrame, text="Akshay Member", width=13, anchor=W, justify=LEFT,
                            font=NORM_VERDANA_FONT, bg='light yellow')

        akshay_no = Label(dataEntryFrame, text="Akshay Box No", width=13, anchor=W, justify=LEFT,
                          font=NORM_VERDANA_FONT, bg='light yellow')
        magazine = Label(dataEntryFrame, text="Patrika Subs.", width=13, anchor=W, justify=LEFT,
                         font=NORM_VERDANA_FONT, bg='light yellow')

        mem_pic = Label(dataEntryFrame, text="Member Photo", width=13, anchor=W, justify=LEFT,
                        font=NORM_VERDANA_FONT, bg='light yellow')
        residentIdPic = Label(dataEntryFrame, text="Id Picture", width=15, anchor=W, justify=LEFT,
                              font=NORM_VERDANA_FONT, bg='light yellow')

        heading.grid(row=0, column=0)
        dataEntryFrame.grid(row=1, column=0, padx=10, pady=5)
        id_number.grid(row=1, column=0)
        name.grid(row=1, column=2)
        father_name.grid(row=1, column=4, pady=5)
        mother_name.grid(row=2, column=0, pady=5)
        dateOfBirth.grid(row=2, column=2, pady=5)
        gender.grid(row=2, column=4, pady=5)
        address.grid(row=3, column=0, pady=2)
        residentCity.grid(row=3, column=4, pady=5)
        residentState.grid(row=4, column=0, pady=5)
        residentCountry.grid(row=4, column=2, pady=5)
        residentPincode.grid(row=4, column=4, pady=5)
        contactNo.grid(row=5, column=0, pady=5)
        nationality.grid(row=5, column=2, pady=5)
        emailId.grid(row=5, column=4, pady=5)
        residentIdType.grid(row=6, column=0, pady=5)
        associatedSince.grid(row=6, column=2, pady=5)
        profession.grid(row=6, column=4, pady=5)
        updeshBy.grid(row=7, column=0, pady=5)
        initiated_stage.grid(row=7, column=2, pady=5)
        designation.grid(row=7, column=4, pady=5)
        akshayPatra.grid(row=8, column=0, pady=5)
        akshay_no.grid(row=8, column=2, pady=5)
        magazine.grid(row=8, column=4, pady=5)
        mem_pic.grid(row=9, column=0, pady=5)
        residentIdPic.grid(row=9, column=3, pady=5)

        self.default_text1 = StringVar(dataEntryFrame, value='Not Available')
        self.default_text2 = StringVar(dataEntryFrame, value='Not Available')
        self.default_text3 = StringVar(dataEntryFrame, value='Not Available')

        self.default_text4 = StringVar(dataEntryFrame, value='Not Available')
        self.default_text5 = StringVar(dataEntryFrame, value='')
        self.default_text6 = StringVar(dataEntryFrame, value='')
        self.default_text7 = StringVar(dataEntryFrame, value='')

        member_govtId = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'),
                              textvariable=self.default_text4)
        member_name = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'),
                            textvariable=self.default_text5, justify=LEFT)
        member_fatherName = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'),
                                  textvariable=self.default_text1)
        member_mother = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'),
                              textvariable=self.default_text2)
        cal = DateEntry(dataEntryFrame, width=22, date_pattern='dd/MM/yyyy', font=('arial narrow', 11, 'normal'))

        gender_variable = StringVar(dataEntryFrame)
        gender_variable.set("Male")
        member_gender = OptionMenu(dataEntryFrame, gender_variable, "Male", "Female")
        member_gender.configure(bg='white', fg='black', font=('verdana', 9, 'normal'))
        member_address = Entry(dataEntryFrame, width=66, font=('arial narrow', 11, 'normal'),
                               textvariable=self.default_text6)
        member_city = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'),
                            textvariable=self.default_text7)
        member_state = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'))
        member_country = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'))
        member_pincode = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'),
                               textvariable=self.default_text3)
        member_contactNo = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'))
        member_nationality = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'))
        member_emailId = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'))

        variable_idType = StringVar(dataEntryFrame)
        variable_idType.set("Other")
        member_idType = OptionMenu(dataEntryFrame, variable_idType, "Other", "Aadhaar Card", "Passport", "Voter Id",
                                   "PAN Card", "Driving License",
                                   "Ration Card", "Caste Certificate",
                                   "Student Photo Id ")
        member_idType.configure(bg='white', width=17, fg='black', font=('verdana', 9, 'normal'))
        cal_associatedSince = DateEntry(dataEntryFrame, width=22, date_pattern='dd/MM/yyyy',
                                        font=('arial narrow', 11, 'normal'))
        profession = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'))
        updestha = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'))
        updesh_stage = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'))
        designation_varaible = StringVar(dataEntryFrame)
        designation_varaible.set("Other")
        designation_Type = OptionMenu(dataEntryFrame, designation_varaible, "Updestha", "President", "Vice-President",
                                      "Treasurer", "Trustee", "Manager",
                                      "Member", "Staff-Sevak", "Area Co-ordinator",
                                      "Other")
        designation_Type.configure(bg='white', width=17, fg='black', font=('verdana', 9, 'normal'))
        akshayPatra_varaible = StringVar(dataEntryFrame)
        akshayPatra_varaible.set("No")
        akshayPatra_Type = OptionMenu(dataEntryFrame, akshayPatra_varaible, "Yes", "No")
        akshayPatra_Type.configure(bg='white', width=17, fg='black', font=('verdana', 9, 'normal'))
        akshayPatra_No = Entry(dataEntryFrame, width=25, font=('arial narrow', 11, 'normal'))
        magazine_subsVariable = StringVar(dataEntryFrame)
        magazine_subsVariable.set("No")
        magazine_subsType = OptionMenu(dataEntryFrame, magazine_subsVariable, "Yes", "No")
        magazine_subsType.configure(bg='white', width=17, fg='black', font=('verdana', 9, 'normal'))

        member_govtId.grid(row=1, column=1, sticky=W)
        member_name.grid(row=1, column=3, sticky=W)
        member_fatherName.grid(row=1, column=5, sticky=W, pady=7)
        member_mother.grid(row=2, column=1, sticky=W)
        cal.grid(row=2, column=3, sticky=W, pady=5)
        member_gender.configure(width=17, height=1, pady=5)
        member_gender.grid(row=2, column=5, sticky=W, pady=5)
        member_address.grid(row=3, column=1, columnspan=3, sticky=W, pady=2)
        member_city.grid(row=3, column=5, sticky=W, pady=5)
        member_state.grid(row=4, column=1, sticky=W, pady=5)
        member_country.grid(row=4, column=3, sticky=W, pady=5)
        member_pincode.grid(row=4, column=5, sticky=W, pady=5)
        member_contactNo.grid(row=5, column=1, sticky=W, pady=5)
        member_nationality.grid(row=5, column=3, sticky=W, pady=5)
        member_emailId.grid(row=5, column=5, sticky=W, pady=5)
        member_idType.grid(row=6, column=1, sticky=W, pady=5)
        cal_associatedSince.grid(row=6, column=3, sticky=W, pady=5)
        profession.grid(row=6, column=5, sticky=W, pady=5)
        updestha.grid(row=7, column=1, sticky=W, pady=5)
        updesh_stage.grid(row=7, column=3, sticky=W, pady=5)
        designation_Type.grid(row=7, column=5, sticky=W, pady=5)
        akshayPatra_Type.grid(row=8, column=1, sticky=W, pady=5)
        akshayPatra_No.grid(row=8, column=3, sticky=W, pady=5)
        magazine_subsType.grid(row=8, column=5, sticky=W, pady=5)

        self.newmember_Type = memberType
        # --------------------------------------Member picture start ------------------------------------------
        myFrame = Frame(dataEntryFrame)
        myFrame.grid(row=9, column=1, columnspan=2)
        canvas_width, canvas_height = 150, 150
        canvasMem = Canvas(myFrame, width=canvas_width, height=canvas_height)
        myimage = ImageTk.PhotoImage(PIL.Image.open("..\\Images\\default_member.jpg").resize((150, 150)))
        canvasMem.create_image(0, 0, anchor=NW, image=myimage)
        canvasMem.pack()

        pic_result = partial(self.read_webcam, dataEntryFrame, 1, myFrame, canvasMem)
        member_Picupload = partial(self.pic_upload, dataEntryFrame, 1, myFrame, canvasMem)

        memberPicBtnFrame = Frame(dataEntryFrame, width=200, height=50, bd=4, relief='ridge', bg='light yellow')
        memberPicBtnFrame.grid(row=10, column=1, columnspan=2, padx=5, pady=5)
        member_pic = Button(memberPicBtnFrame, text="Camera", fg="Black", command=pic_result,
                            font=NORM_FONT, width=9, bg='light green')
        member_pic.grid(row=0, column=1)
        member_uploadPic = Button(memberPicBtnFrame, text="Upload", fg="Black", command=member_Picupload,
                                  font=NORM_FONT, width=9, bg='light cyan')
        member_uploadPic.grid(row=0, column=2)
        # --------------------------------------Member picture End ------------------------------------------

        # --------------------------------------Id picture start -------------------------------------------------

        myFrameId_idpic = Frame(dataEntryFrame)
        myFrameId_idpic.grid(row=9, column=4, columnspan=2)
        canvasId_idpic = Canvas(myFrameId_idpic, width=canvas_width, height=canvas_height)
        myIdimage = ImageTk.PhotoImage(PIL.Image.open("..\\Images\\default_idcard.jpg").resize((150, 150)))
        canvasId_idpic.create_image(0, 0, anchor=NW, image=myIdimage)
        canvasId_idpic.pack()

        idpic_result = partial(self.read_webcam, dataEntryFrame, 2, myFrameId_idpic, canvasId_idpic)
        idpic_upload = partial(self.pic_upload, dataEntryFrame, 2, myFrameId_idpic, canvasId_idpic)

        memberIdPicBtnFrame = Frame(dataEntryFrame, width=200, height=50, bd=4, relief='ridge', bg='light yellow')
        memberIdPicBtnFrame.grid(row=10, column=4, columnspan=2, padx=5, pady=5)
        member_pic = Button(memberIdPicBtnFrame, text="Camera", fg="Black", command=idpic_result,
                            font=NORM_FONT, width=9, bg='light green')
        member_pic.grid(row=0, column=4)
        member_uploadPic = Button(memberIdPicBtnFrame, text="Upload", fg="Black", command=idpic_upload,
                                  font=NORM_FONT, width=9, bg='light cyan')
        member_uploadPic.grid(row=0, column=5)

        # --------------------------------------Id picture end-------------------------------------------

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(addMember_window, width=300, height=100, bd=4, relief='ridge', bg='light yellow')
        buttonFrame.grid(column=0, columnspan=5, pady=15)
        infoFrame = Frame(addMember_window, width=300, height=100, bd=4, relief='ridge', bg='light yellow')
        infoLabel = Label(infoFrame, text="Red marked fields are mandatory", width=60, anchor='center', justify=CENTER,
                          font=BOLD_VERDANA_FONT, bg='light yellow', fg='red')
        insert_result = partial(self.register_member_excel, addMember_window, self.newmember_id, member_govtId,
                                member_name,
                                member_fatherName,
                                member_mother,
                                cal,
                                gender_variable,
                                member_address,
                                member_city,
                                member_state,
                                member_pincode,
                                member_contactNo,
                                member_country,
                                member_nationality,
                                member_emailId,
                                variable_idType,
                                cal_associatedSince,
                                profession,
                                updestha,
                                updesh_stage,
                                designation_varaible,
                                akshayPatra_varaible,
                                akshayPatra_No,
                                magazine_subsVariable,
                                memberType, infoLabel)

        # create a Save Button and place into the addMember_window window
        self.submit = Button(buttonFrame, text="Save", fg="Black", command=insert_result,
                             font=NORM_FONT, width=10, bg='light grey', state=DISABLED)
        self.submit.grid(row=0, column=0, padx=2)
        # check_result = partial(self.check_SaveBtn_state, submit)
        self.default_text1.trace("w", self.check_SaveBtn_state)
        self.default_text2.trace("w", self.check_SaveBtn_state)
        self.default_text3.trace("w", self.check_SaveBtn_state)
        self.default_text4.trace("w", self.check_SaveBtn_state)
        self.default_text5.trace("w", self.check_SaveBtn_state)
        self.default_text6.trace("w", self.check_SaveBtn_state)
        self.default_text7.trace("w", self.check_SaveBtn_state)

        clear_result = partial(self.clear_Memberform, member_govtId,
                               member_name,
                               member_fatherName,
                               member_mother,
                               member_gender,
                               member_address,
                               member_city,
                               member_state,
                               member_pincode,
                               member_contactNo,
                               member_country,
                               member_nationality,
                               member_emailId)

        # create a Clear Button and place into the addMember_window window
        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=10, bg='light cyan', underline=0)
        clear.grid(row=0, column=1, padx=1)

        # create a Close Button and place into the addMember_window window
        cancel = Button(buttonFrame, text="Close", fg="Black", command=addMember_window.destroy,
                        font=NORM_FONT, width=10, bg='light cyan', underline=0)
        cancel.grid(row=0, column=3, padx=1)
        # ---------------------------------Button Frame End----------------------------------------

        infoFrame.grid(row=12, column=0, columnspan=5, pady=15)

        infoLabel.grid(row=0, column=0, columnspan=5, pady=1)

        addMember_window.bind('<Return>', lambda event=None: self.submit.invoke())
        addMember_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        addMember_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        addMember_window.focus()
        addMember_window.grab_set()
        mainloop()

    def check_itemSearchBtnState(self, n, m, x, itemId_text, search_btn, submit):
        print("check_for_search button  enabling")

        if n.get() != "" and len(n.get()) > 2:
            m.configure(state=NORMAL, bg='light cyan')
            x.configure(state=NORMAL, bg='light cyan')
        else:
            m.configure(state=DISABLED, bg='light grey')
            x.configure(state=DISABLED, bg='light grey')

    def check_DepositBtn_state(self, *args):
        print("Tracing  entry input")

        if self.default_text1.get() != "" and \
                self.default_text2.get() != "" and \
                self.default_text3.get() != "" and \
                self.default_text4.get() != "" and \
                self.default_text5.get() != "":
            self.submit_deposit.configure(state=NORMAL, bg='light cyan')
        else:
            self.submit_deposit.configure(state=DISABLED, bg='light grey')

    def check_SaveItemBtn_state(self, *args):
        print("Tracing  entry input")

        if self.default_text1.get() != "" and \
                self.default_text3.get() != "" and \
                self.default_text4.get() != "" and \
                self.default_text5.get() != "":
            self.submit.configure(state=NORMAL, bg='light cyan')
        else:
            self.submit.configure(state=DISABLED, bg='light grey')

    def check_SaveBtn_state(self, *args):
        print("Tracing  entry input")

        if self.default_text1.get() != "" and self.default_text2.get() != "" and \
                self.default_text3.get() != "" and \
                self.default_text5.get() != "" and self.default_text4.get() != "" and \
                self.default_text6.get() != "" and self.default_text7.get() != "":
            self.submit.configure(state=NORMAL, bg='light cyan')

        else:
            self.submit.configure(state=DISABLED, bg='light grey')

    def destroyWindow(self, windowToClose):
        if windowToClose is self.newItem_window:
            self.itemEntryInstance = False
        if windowToClose is self.bookBorrow_window:
            self.bookBorrowInstance = False
        if windowToClose is self.bookReturn_window:
            self.bookReturnInstance = False
        windowToClose.destroy()

    def item_entry_commercial(self):
        self.newItem_window = Toplevel(self.master)
        self.newItem_window.title("New Data Entry ")
        self.newItem_window.geometry('970x300+240+200')
        self.newItem_window.configure(background='wheat')
        self.newItem_window.resizable(width=False, height=False)
        self.newItem_window.protocol('WM_DELETE_WINDOW', self.donothing)
        heading = Label(self.newItem_window, text="New Book/Sukrit Product Entry", font=('ariel narrow', 15, 'bold'),
                        bg='wheat')

        dataEntryFrame = Frame(self.newItem_window, width=200, height=100, bd=4, relief='ridge',
                               bg='snow')
        self.default_text1 = StringVar(dataEntryFrame, value='')
        self.default_text2 = StringVar(dataEntryFrame, value='')
        self.default_text3 = StringVar(dataEntryFrame, value='')
        self.default_text4 = StringVar(dataEntryFrame, value='')
        self.default_text5 = StringVar(dataEntryFrame, value='0')
        # create a Book Name label
        name = Label(dataEntryFrame, text="Book Name", width=13, anchor=W, justify=LEFT, font=TIMES_NEW_ROMAN_BIG,
                     bg='snow')

        # create a Author label
        author = Label(dataEntryFrame, text="Author", width=13, anchor=W, justify=LEFT, font=TIMES_NEW_ROMAN_BIG,
                       bg='snow')

        # create a Price label
        price = Label(dataEntryFrame, text="Price(Rs.)", width=13, anchor=W, justify=LEFT,
                      font=TIMES_NEW_ROMAN_BIG, bg='snow')

        # create a Quantity label
        quantity = Label(dataEntryFrame, text="Quantity", width=13, anchor=W, justify=LEFT,
                         font=TIMES_NEW_ROMAN_BIG, bg='snow')

        # create a borrow fee label
        borrowFee = Label(dataEntryFrame, text="Borrow Fee(Rs.)", width=13, anchor=W, justify=LEFT,
                          font=TIMES_NEW_ROMAN_BIG, bg='snow')

        # create a borrow fee label
        rackNumber = Label(dataEntryFrame, text="Rack Number", width=13, anchor=W, justify=LEFT,
                           font=TIMES_NEW_ROMAN_BIG, bg='snow')

        # create a borrow fee label
        date_label = Label(dataEntryFrame, text="Received On", width=13, anchor=W, justify=LEFT,
                           font=TIMES_NEW_ROMAN_BIG, bg='snow')

        center_location = Label(dataEntryFrame, text="Center Name", width=13, anchor=W, justify=LEFT,
                                font=TIMES_NEW_ROMAN_BIG, bg='snow')

        heading.grid(row=0, column=0, columnspan=2)
        dataEntryFrame.grid(row=1, column=1, padx=10, pady=5)
        name.grid(row=0, column=0, pady=7)
        author.grid(row=0, column=2, pady=7)
        price.grid(row=1, column=0, pady=7)
        quantity.grid(row=1, column=2, pady=7)
        borrowFee.grid(row=2, column=0, pady=7)
        rackNumber.grid(row=2, column=2, pady=7)
        date_label.grid(row=3, column=0, pady=7)
        center_location.grid(row=3, column=2, pady=7)

        # create a text entry box
        # for typing the information
        item_name = Entry(dataEntryFrame, width=27, font=TIMES_NEW_ROMAN_BIG, bg='light yellow',
                          textvariable=self.default_text1)

        authorText = StringVar(dataEntryFrame)
        authorList = self.obj_commonUtil.getAuthorNames()
        print("Author list  - ", authorList)
        authorText.set(authorList[0])
        author_menu = OptionMenu(dataEntryFrame, authorText, *authorList)
        author_menu.configure(width=24, font=TIMES_NEW_ROMAN_BIG, bg='light yellow', anchor=W, justify=LEFT)

        item_price = Entry(dataEntryFrame, width=27, font=TIMES_NEW_ROMAN_BIG, bg='light yellow',
                           textvariable=self.default_text3)
        item_quantity = Entry(dataEntryFrame, width=27, font=TIMES_NEW_ROMAN_BIG, bg='light yellow',
                              textvariable=self.default_text4)
        item_borrowfee = Entry(dataEntryFrame, width=27, text='0', state=DISABLED, font=TIMES_NEW_ROMAN_BIG,
                               bg='light grey', textvariable=self.default_text5)
        rack_location = Entry(dataEntryFrame, width=27, font=TIMES_NEW_ROMAN_BIG, bg='light yellow')

        cal = DateEntry(dataEntryFrame, width=25, date_pattern='dd/MM/yyyy', font=TIMES_NEW_ROMAN_BIG,
                        bg='light yellow',
                        justify='left')
        local_centerText = StringVar(dataEntryFrame)

        localCenterList = self.obj_commonUtil.getLocalCenterNames()
        print("Center list  - ", localCenterList)
        local_centerText.set(localCenterList[0])
        localcenter_menu = OptionMenu(dataEntryFrame, local_centerText, *localCenterList)
        localcenter_menu.configure(width=24, font=TIMES_NEW_ROMAN_BIG, bg='light yellow', anchor=W, justify=LEFT)

        item_name.grid(row=0, column=1, pady=7)
        author_menu.grid(row=0, column=3, pady=7)
        item_price.grid(row=1, column=1, pady=7)
        item_quantity.grid(row=1, column=3, pady=7)
        item_borrowfee.grid(row=2, column=1, pady=7)
        rack_location.grid(row=2, column=3, pady=7)
        cal.grid(row=3, column=1, pady=7)
        localcenter_menu.grid(row=3, column=3, pady=7)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(self.newItem_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=1, pady=5)
        self.submit = Button(buttonFrame)
        insert_result = partial(self.insert_commercial_data_Excel, self.newItem_window, item_name, authorText,
                                item_price,
                                item_borrowfee,
                                item_quantity, rack_location, cal, local_centerText)

        # create a Save Button and place into the self.newItem_window window
        self.submit.configure(text="Save", fg="Black", command=insert_result,
                              font=TIMES_NEW_ROMAN_BIG, width=8, bg='light grey', state=DISABLED)
        self.submit.grid(row=0, column=0)
        self.default_text1.trace("w", self.check_SaveItemBtn_state)
        # self.default_text2.trace("w", self.check_SaveItemBtn_state)
        self.default_text3.trace("w", self.check_SaveItemBtn_state)
        self.default_text4.trace("w", self.check_SaveItemBtn_state)
        self.default_text5.trace("w", self.check_SaveItemBtn_state)
        clear_result = partial(self.clear_form, item_name, authorText, item_price, item_borrowfee, item_quantity)

        # create a Clear Button and place into the self.newItem_window window
        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=TIMES_NEW_ROMAN_BIG, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=1)

        # create a Cancel Button and place into the self.newItem_window window
        cancel_Result = partial(self.destroyWindow, self.newItem_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_Result,
                        font=TIMES_NEW_ROMAN_BIG, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=2)
        # ---------------------------------Button Frame End----------------------------------------

        self.newItem_window.bind('<Return>', lambda event=None: self.submit.invoke())
        self.newItem_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.newItem_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.newItem_window.focus()
        self.newItem_window.grab_set()
        mainloop()

    def deposit_seva_rashi(self):
        self.new_sevaRashi_deposit_window = Toplevel(self.master)
        self.new_sevaRashi_deposit_window.title("Deposit Seva Rashi ")
        self.new_sevaRashi_deposit_window.geometry('700x335+240+200')
        self.new_sevaRashi_deposit_window.configure(background='wheat')
        self.new_sevaRashi_deposit_window.resizable(width=False, height=False)
        self.new_sevaRashi_deposit_window.protocol('WM_DELETE_WINDOW', self.donothing)
        heading = Label(self.new_sevaRashi_deposit_window, text="Deposit Seva Rashi Form",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        dataEntryFrame = Frame(self.new_sevaRashi_deposit_window, width=200, height=100, bd=4, relief='ridge',
                               bg='snow')
        self.default_text1 = StringVar(dataEntryFrame, value='')
        self.default_text2 = StringVar(dataEntryFrame, value='')
        self.default_text3 = StringVar(dataEntryFrame, value='')
        self.default_text4 = StringVar(dataEntryFrame, value='')
        self.default_text5 = StringVar(dataEntryFrame, value='None')
        # lower frame added to show the result of transactions
        infoFrame = Frame(self.new_sevaRashi_deposit_window, width=100, height=20, bd=4, relief='ridge')
        # create a Book Name label
        donator_name = Label(dataEntryFrame, text="Received From(Id)", width=14, anchor=W, justify=LEFT,
                             font=NORM_FONT,
                             bg='snow')

        # create a Author label
        seva_amount = Label(dataEntryFrame, text="Amount(Rs.)", width=13, anchor=W, justify=LEFT,
                            font=NORM_FONT,
                            bg='snow')

        # create a Price label
        category = Label(dataEntryFrame, text="Seva Type", width=13, anchor=W, justify=LEFT,
                         font=NORM_FONT, bg='snow')

        dateOfCollection = Label(dataEntryFrame, text="Date", width=13, anchor=W,
                                 justify=LEFT,
                                 font=NORM_FONT, bg='snow')

        # create a Quantity label
        collector_name = Label(dataEntryFrame, text="Receiver (Id)", width=13, anchor=W, justify=LEFT,
                               font=NORM_FONT, bg='snow')

        # create a borrow fee label
        modeOfpayment = Label(dataEntryFrame, text="Payment Mode", width=13, anchor=W, justify=LEFT,
                              font=NORM_FONT, bg='snow')

        authorizedBy_name = Label(dataEntryFrame, text="Authorizor (Id)", width=13, anchor=W, justify=LEFT,
                                  font=NORM_FONT, bg='snow')

        # create a borrow fee label
        akshay_boxNo = Label(dataEntryFrame, text="Akshay Patra Id", width=13, anchor=W, justify=LEFT,
                             font=NORM_FONT, bg='snow')

        infolabel = Label(infoFrame, text="All fields are mandatory !!", width=50, anchor='center',
                          justify=LEFT,
                          font=NORM_VERDANA_FONT, bg='snow')

        # create a borrow fee label
        invoice_id = Label(dataEntryFrame, text="Invoice Id", width=13, anchor=W, justify=LEFT,
                           font=NORM_FONT, bg='snow')

        heading.grid(row=0, column=0, columnspan=2)
        dataEntryFrame.grid(row=1, column=1, padx=10, pady=8)
        donator_name.grid(row=0, column=0, pady=3)
        seva_amount.grid(row=0, column=2, padx=10, pady=3)
        category.grid(row=1, column=0, pady=3)
        dateOfCollection.grid(row=1, column=2, pady=3)
        collector_name.grid(row=2, column=0, pady=3)
        modeOfpayment.grid(row=2, column=2, pady=3)
        authorizedBy_name.grid(row=3, column=0, pady=3)
        akshay_boxNo.grid(row=3, column=2, pady=3)
        invoice_id.grid(row=4, column=0, pady=3)

        # create a text entry box
        # for typing the information
        donator_idText = Entry(dataEntryFrame, width=20, font=NORM_FONT, bg='light yellow',
                               textvariable=self.default_text1)
        seva_amountText = Entry(dataEntryFrame, width=20, font=NORM_FONT, bg='light yellow',
                                textvariable=self.default_text2)
        categoryText = StringVar(dataEntryFrame)
        category_list = ['Monthly Seva', 'Gaushala Seva', 'Hawan Seva', 'Event/Prachar Seva',
                         'Aarti Seva', 'Akshay-Patra Seva', 'Ashram Seva(Generic)', 'Ashram Nirmaan Seva', 'Yoga Fees']
        categoryText.set("Other")
        category_menu = OptionMenu(dataEntryFrame, categoryText, *category_list)
        cal = DateEntry(dataEntryFrame, width=18, font=NORM_FONT, date_pattern='dd/MM/yyyy', bg='light yellow',
                        anchor=W, justify=LEFT)

        collector_nameText = Entry(dataEntryFrame, width=20, font=NORM_FONT, bg='light yellow',
                                   textvariable=self.default_text3)

        paymentMode_text = StringVar(dataEntryFrame)
        paymentMode_list = ['Cash', 'Bank Transfer', 'Cheque', 'Demand Draft', 'Other']
        paymentMode_text.set("Other")
        paymentMode_menu = OptionMenu(dataEntryFrame, paymentMode_text, *paymentMode_list)
        paymentMode_menu.configure(width=16, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)
        authorizedby_Text = Entry(dataEntryFrame, width=20, font=NORM_FONT, bg='light yellow',
                                  textvariable=self.default_text4)
        akshayPatra_Text = Entry(dataEntryFrame, width=20, font=NORM_FONT, bg='light yellow', state=DISABLED)

        invoice_idText = Label(dataEntryFrame, text="------------", width=16, anchor=W, justify=LEFT,
                               font=NORM_FONT, bg='snow')

        trans_idLabel = Label(dataEntryFrame, text="Transaction Id", width=13, anchor=W, justify=LEFT,
                              font=NORM_FONT, bg='snow')
        transId_Text = Entry(dataEntryFrame, width=20, font=NORM_FONT, bg='light yellow', state=DISABLED)

        donator_idText.grid(row=0, column=1, pady=3)
        seva_amountText.grid(row=0, column=3, pady=3)
        category_menu.configure(width=16, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)
        category_menu.grid(row=1, column=1, pady=3)
        cal.grid(row=1, column=3, pady=3)
        collector_nameText.grid(row=2, column=1, pady=3)
        check_result_transactionId = partial(self.transid_enable, paymentMode_text, transId_Text)
        paymentMode_menu.grid(row=2, column=3, pady=3)
        authorizedby_Text.grid(row=3, column=1, pady=3)
        akshayPatra_Text.grid(row=3, column=3, pady=3)
        check_result = partial(self.check_for_selection, categoryText, akshayPatra_Text, donator_idText)
        invoice_idText.grid(row=4, column=1, pady=3)
        trans_idLabel.grid(row=4, column=2)
        transId_Text.grid(row=4, column=3, pady=3)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(self.new_sevaRashi_deposit_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=1, pady=8)
        print_invoice = Button(buttonFrame, text="Print Receipt", fg="Black", command=None,
                               font=NORM_FONT, width=10, bg='light grey', state=DISABLED)
        self.submit_deposit = Button(buttonFrame)
        insert_result = partial(self.deposit_seva_rashi_Excel,
                                self.new_sevaRashi_deposit_window,
                                donator_idText,
                                seva_amountText,
                                categoryText,
                                collector_nameText,
                                cal,
                                paymentMode_menu,
                                paymentMode_text,
                                authorizedby_Text,
                                akshayPatra_Text,
                                invoice_idText,
                                infolabel, print_invoice, transId_Text)

        # create a Save Button and place into the self.new_sevaRashi_deposit_window window
        self.submit_deposit.configure(text="Deposit", fg="Black", command=insert_result,
                                      font=NORM_FONT, width=8, bg='light grey', underline=0, state=DISABLED)
        self.submit_deposit.grid(row=0, column=0)
        print_invoice.grid(row=0, column=1)
        self.default_text1.trace("w", self.check_DepositBtn_state)
        self.default_text2.trace("w", self.check_DepositBtn_state)
        self.default_text3.trace("w", self.check_DepositBtn_state)
        self.default_text4.trace("w", self.check_DepositBtn_state)

        categoryText.trace("w", check_result)
        paymentMode_text.trace("w", check_result_transactionId)
        # create a Clear Button and place into the self.new_sevaRashi_deposit_window window
        # clear the commercial form after data has been saved
        clear_result = partial(self.clearMonetarySevaDepositForm,
                               donator_idText,
                               seva_amountText,
                               categoryText,
                               collector_nameText,
                               cal,
                               paymentMode_menu,
                               paymentMode_text,
                               authorizedby_Text,
                               akshayPatra_Text,
                               invoice_idText, print_invoice)
        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=2)

        # create a Cancel Button and place into the self.new_sevaRashi_deposit_window window
        cancel_Result = partial(self.destroyWindow, self.new_sevaRashi_deposit_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_Result,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=3)
        # ---------------------------------Button Frame End----------------------------------------

        infoFrame.grid(row=9, column=1, pady=5)
        infolabel.grid(row=0, column=0, padx=2, pady=3)

        self.new_sevaRashi_deposit_window.bind('<Return>', lambda event=None: submit.invoke())
        self.new_sevaRashi_deposit_window.bind('<Alt-d>', lambda event=None: submit.invoke())
        self.new_sevaRashi_deposit_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.new_sevaRashi_deposit_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.new_sevaRashi_deposit_window.focus()
        self.new_sevaRashi_deposit_window.grab_set()
        mainloop()

    def fulfill_pledge_form(self):
        self.pledge_payment_window = Toplevel(self.master)
        self.pledge_payment_window.title("Pledge Amount Fulfillment ")
        self.pledge_payment_window.geometry('770x360+240+200')
        self.pledge_payment_window.configure(background='wheat')
        self.pledge_payment_window.resizable(width=False, height=False)
        self.pledge_payment_window.protocol('WM_DELETE_WINDOW', self.donothing)
        heading = Label(self.pledge_payment_window, text="Pledge Deposit",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        dataEntryFrame = Frame(self.pledge_payment_window, width=200, height=100, bd=4, relief='ridge',
                               bg='snow')
        self.default_text1 = StringVar(dataEntryFrame, value='')
        self.default_text2 = StringVar(dataEntryFrame, value='')
        self.default_text3 = StringVar(dataEntryFrame, value='')
        self.default_text4 = StringVar(dataEntryFrame, value='')
        self.default_text5 = StringVar(dataEntryFrame, value='None')
        # lower frame added to show the result of transactions
        infoFrame = Frame(self.pledge_payment_window, width=100, height=20, bd=4, relief='ridge')
        # create a Book Name label
        donator_name = Label(dataEntryFrame, text="Received From(Id)", width=13, anchor=W, justify=LEFT,
                             font=NORM_FONT,
                             bg='snow')

        seva_amount = Label(dataEntryFrame, text="Amount(Rs.)", width=13, anchor=W, justify=LEFT,
                            font=NORM_FONT,
                            bg='snow')

        category = Label(dataEntryFrame, text="Pledge Item", width=13, anchor=W, justify=LEFT,
                         font=NORM_FONT, bg='snow')

        trust_label = Label(dataEntryFrame, text="Trust Name", width=13, anchor=W, justify=LEFT,
                            font=NORM_FONT, bg='snow')

        dateOfCollection = Label(dataEntryFrame, text="Date", width=13, anchor=W,
                                 justify=LEFT,
                                 font=NORM_FONT, bg='snow')

        # create a Quantity label
        collector_name = Label(dataEntryFrame, text="Receiver (Id)", width=13, anchor=W, justify=LEFT,
                               font=NORM_FONT, bg='snow')

        # create a borrow fee label
        modeOfpayment = Label(dataEntryFrame, text="Payment Mode", width=13, anchor=W, justify=LEFT,
                              font=NORM_FONT, bg='snow')

        authorizedBy_name = Label(dataEntryFrame, text="Authorizor (Id)", width=13, anchor=W, justify=LEFT,
                                  font=NORM_FONT, bg='snow')

        infolabel = Label(infoFrame, text="All fields are mandatory !!", width=50, anchor='center',
                          justify=LEFT,
                          font=NORM_VERDANA_FONT, bg='snow')

        # create a borrow fee label
        invoice_id = Label(dataEntryFrame, text="Invoice Id", width=13, anchor=W, justify=LEFT,
                           font=NORM_FONT, bg='snow')

        heading.grid(row=0, column=0, columnspan=2)
        dataEntryFrame.grid(row=1, column=1, padx=10, pady=8)
        donator_name.grid(row=0, column=0, pady=5)
        seva_amount.grid(row=0, column=2, padx=10, pady=5)
        category.grid(row=1, column=0, pady=5)
        trust_label.grid(row=1, column=2, pady=5)
        collector_name.grid(row=2, column=0, pady=5)
        dateOfCollection.grid(row=2, column=2, pady=5)
        authorizedBy_name.grid(row=3, column=0, pady=5)
        modeOfpayment.grid(row=3, column=2, pady=5)

        invoice_id.grid(row=4, column=0, pady=5)

        # create a text entry box
        # for typing the information
        donator_idText = Entry(dataEntryFrame, width=25, font=NORM_FONT, bg='light yellow',
                               textvariable=self.default_text1)
        seva_amountText = Entry(dataEntryFrame, width=25, font=NORM_FONT, bg='light yellow',
                                textvariable=self.default_text2)
        pledge_itemList = self.obj_commonUtil.getPledgeItemNames()
        categoryText = StringVar(dataEntryFrame)
        categoryText.set(pledge_itemList[0])
        category_menu = OptionMenu(dataEntryFrame, categoryText, *pledge_itemList)
        category_menu.configure(width=21, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)
        trust_nametext = StringVar(dataEntryFrame)
        trustName_list = ['Vihangam Yoga (Karnataka) Trust', 'Aadarsh gaushala Trust']
        trust_nametext.set("Vihangam Yoga (Karnataka) Trust")
        trust_menu = OptionMenu(dataEntryFrame, trust_nametext, *trustName_list)
        trust_menu.configure(width=21, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)

        cal = DateEntry(dataEntryFrame, width=23, font=NORM_FONT, date_pattern='dd/MM/yyyy', bg='light yellow',
                        anchor=W, justify=LEFT)

        collector_nameText = Entry(dataEntryFrame, width=25, font=NORM_FONT, bg='light yellow',
                                   textvariable=self.default_text3)

        paymentMode_text = StringVar(dataEntryFrame)
        paymentMode_list = ['Cash', 'Bank Transfer', 'Cheque', 'Demand Draft', 'Other']
        paymentMode_text.set("Other")
        paymentMode_menu = OptionMenu(dataEntryFrame, paymentMode_text, *paymentMode_list)
        paymentMode_menu.configure(width=21, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)
        authorizedby_Text = Entry(dataEntryFrame, width=25, font=NORM_FONT, bg='light yellow',
                                  textvariable=self.default_text4)
        invoice_idText = Label(dataEntryFrame, text="------------", width=16, anchor=W, justify=LEFT,
                               font=NORM_FONT, bg='snow')

        trans_idLabel = Label(dataEntryFrame, text="Transaction Id", width=13, anchor=W, justify=LEFT,
                              font=NORM_FONT, bg='snow')
        transId_Text = Entry(dataEntryFrame, width=25, font=NORM_FONT, bg='light yellow', state=DISABLED)

        donator_idText.grid(row=0, column=1, pady=5)
        seva_amountText.grid(row=0, column=3, pady=5)
        category_menu.grid(row=1, column=1, pady=5)
        trust_menu.grid(row=1, column=3, pady=5)
        collector_nameText.grid(row=2, column=1, pady=5)
        check_result_transactionId = partial(self.transid_enable, paymentMode_text, transId_Text)
        cal.grid(row=2, column=3, pady=5)
        authorizedby_Text.grid(row=3, column=1, pady=5)
        invoice_idText.grid(row=4, column=1, pady=5)
        paymentMode_menu.grid(row=3, column=3)
        trans_idLabel.grid(row=4, column=2)
        transId_Text.grid(row=4, column=3, pady=5)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(self.pledge_payment_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=1, pady=8)
        print_invoice = Button(buttonFrame, text="Print Receipt", fg="Black", command=None,
                               font=NORM_FONT, width=10, bg='light grey', state=DISABLED)
        self.submit_deposit = Button(buttonFrame)
        insert_result = partial(self.update_pledge_payment,
                                self.pledge_payment_window,
                                donator_idText,
                                seva_amountText,
                                categoryText,
                                collector_nameText,
                                cal,
                                trust_nametext,
                                paymentMode_menu,
                                paymentMode_text,
                                authorizedby_Text,
                                invoice_idText,
                                infolabel, print_invoice, transId_Text)

        # create a Save Button and place into the self.pledge_payment_window window
        self.submit_deposit.configure(text="Deposit", fg="Black", command=insert_result,
                                      font=NORM_FONT, width=8, bg='light grey', underline=0, state=DISABLED)
        self.submit_deposit.grid(row=0, column=0)
        print_invoice.grid(row=0, column=1)
        self.default_text1.trace("w", self.check_DepositBtn_state)
        self.default_text2.trace("w", self.check_DepositBtn_state)
        self.default_text3.trace("w", self.check_DepositBtn_state)
        self.default_text4.trace("w", self.check_DepositBtn_state)

        paymentMode_text.trace("w", check_result_transactionId)
        # create a Clear Button and place into the self.pledge_payment_window window
        # clear the commercial form after data has been saved
        '''
        clear_result = partial(self.clearMonetarySevaDepositForm,
                               donator_idText,
                               seva_amountText,
                               categoryText,
                               collector_nameText,
                               cal,
                               paymentMode_menu,
                               paymentMode_text,
                               authorizedby_Text,
                               akshayPatra_Text,
                               invoice_idText, print_invoice)
                               '''
        clear = Button(buttonFrame, text="Reset", fg="Black", command=None,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=2)

        # create a Cancel Button and place into the self.pledge_payment_window window
        cancel_Result = partial(self.destroyWindow, self.pledge_payment_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_Result,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=3)
        # ---------------------------------Button Frame End----------------------------------------

        infoFrame.grid(row=9, column=1, pady=5)
        infolabel.grid(row=0, column=0, padx=2, pady=5)

        self.pledge_payment_window.bind('<Return>', lambda event=None: submit.invoke())
        self.pledge_payment_window.bind('<Alt-d>', lambda event=None: submit.invoke())
        self.pledge_payment_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.pledge_payment_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.pledge_payment_window.focus()
        self.pledge_payment_window.grab_set()
        mainloop()

    def register_new_sankalp_item(self):
        self.new_sankalp_item = Toplevel(self.master)
        self.new_sankalp_item.title("New Pledge Registration ")
        self.new_sankalp_item.geometry('420x220+240+200')
        self.new_sankalp_item.configure(background='wheat')
        self.new_sankalp_item.resizable(width=False, height=False)
        self.new_sankalp_item.protocol('WM_DELETE_WINDOW', self.donothing)
        heading = Label(self.new_sankalp_item, text="Create New Pledge Item",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        dataEntryFrame = Frame(self.new_sankalp_item, width=200, height=100, bd=4, relief='ridge',
                               bg='snow')
        self.default_text1 = StringVar(dataEntryFrame, value='')
        self.default_text2 = StringVar(dataEntryFrame, value='')

        # lower frame added to show the result of transactions
        infoFrame = Frame(self.new_sankalp_item, width=100, height=20, bd=4, relief='ridge')
        # create a Book Name label
        pledge_item = Label(dataEntryFrame, text="Pledge Item", width=15, anchor=W, justify=LEFT,
                            font=NORM_FONT,
                            bg='snow')

        dateofpledge = Label(dataEntryFrame, text="Pledge Date", width=15, anchor=W,
                             justify=LEFT,
                             font=NORM_FONT, bg='snow')

        trust_name = Label(dataEntryFrame, text="Trust Name", width=15, anchor=W,
                           justify=LEFT,
                           font=NORM_FONT, bg='snow')

        infolabel = Label(infoFrame, text="All fields are mandatory !!", width=50, anchor='center',
                          justify=LEFT,
                          font=NORM_VERDANA_FONT, bg='snow')

        heading.grid(row=0, column=0, columnspan=2)
        dataEntryFrame.grid(row=1, column=1, padx=10, pady=8)
        pledge_item.grid(row=0, column=0, pady=5)
        dateofpledge.grid(row=1, column=0, pady=5)
        trust_name.grid(row=2, column=0, pady=5)

        # create a text entry box
        # for typing the information
        pledge_text = Entry(dataEntryFrame, width=25, font=NORM_FONT, bg='light yellow',
                            textvariable=self.default_text1)

        trust_nametext = StringVar(dataEntryFrame)
        trustName_list = ['Vihangam Yoga (Karnataka) Trust', 'Aadarsh gaushala Trust']
        trust_nametext.set("Vihangam Yoga (Karnataka) Trust")
        trust_menu = OptionMenu(dataEntryFrame, trust_nametext, *trustName_list)
        trust_menu.configure(width=21, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)

        cal = DateEntry(dataEntryFrame, width=23, font=NORM_FONT, date_pattern='dd/MM/yyyy', bg='light yellow',
                        anchor=W, justify=LEFT)

        pledge_text.grid(row=0, column=1, pady=5)
        cal.grid(row=1, column=1, pady=5)
        trust_menu.grid(row=2, column=1, pady=5)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(self.new_sankalp_item, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=1, pady=8)
        submit_deposit = Button(buttonFrame)

        insert_result = partial(self.registerPledgeItem, trust_nametext, pledge_text)

        # create a Save Button and place into the self.new_sankalp_item window
        submit_deposit.configure(text="Register", fg="Black", command=insert_result,
                                 font=NORM_FONT, width=8, bg='light cyan', underline=0, state=NORMAL)
        submit_deposit.grid(row=0, column=0)

        clear_result = partial(self.clearRegisterPledgeForm,
                               pledge_text,
                               trust_nametext)

        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=1)

        # create a Cancel Button and place into the self.new_sankalp_item window
        cancel_Result = partial(self.destroyWindow, self.new_sankalp_item)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_Result,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=2)
        # ---------------------------------Button Frame End----------------------------------------

        infoFrame.grid(row=9, column=1, pady=5)
        infolabel.grid(row=0, column=0, padx=2, pady=3)

        self.new_sankalp_item.bind('<Return>', lambda event=None: submit_deposit.invoke())
        self.new_sankalp_item.bind('<Alt-d>', lambda event=None: submit_deposit.invoke())
        self.new_sankalp_item.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.new_sankalp_item.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.new_sankalp_item.focus()
        self.new_sankalp_item.grab_set()
        mainloop()

    def register_center(self):
        self.new_center_window = Toplevel(self.master)
        self.new_center_window.title("New Center Registration ")
        self.new_center_window.geometry('455x350+240+200')
        self.new_center_window.configure(background='wheat')
        self.new_center_window.resizable(width=False, height=False)
        self.new_center_window.protocol('WM_DELETE_WINDOW', self.donothing)
        heading = Label(self.new_center_window, text="New Center Registration",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        dataEntryFrame = Frame(self.new_center_window, width=200, height=100, bd=4, relief='ridge',
                               bg='snow')
        self.default_text1 = StringVar(dataEntryFrame, value='')
        self.default_text2 = StringVar(dataEntryFrame, value='')
        self.default_text3 = StringVar(dataEntryFrame, value='')
        self.default_text4 = StringVar(dataEntryFrame, value='')

        # lower frame added to show the result of transactions
        infoFrame = Frame(self.new_center_window, width=70, height=20, bd=4, relief='ridge')
        # create a Book Name label
        pledge_item = Label(dataEntryFrame, text="Center Name", width=15, anchor=W, justify=LEFT,
                            font=NORM_FONT,
                            bg='snow')

        dateofpledge = Label(dataEntryFrame, text="Registration Date", width=15, anchor=W,
                             justify=LEFT,
                             font=NORM_FONT, bg='snow')

        trust_name = Label(dataEntryFrame, text="Trust Name", width=15, anchor=W,
                           justify=LEFT,
                           font=NORM_FONT, bg='snow')

        managerlabel = Label(dataEntryFrame, text="Manager Name", width=15, anchor=W, justify=LEFT,
                             font=NORM_FONT,
                             bg='snow')

        addresslabel = Label(dataEntryFrame, text="Address", width=15, anchor=W,
                             justify=LEFT,
                             font=NORM_FONT, bg='snow')

        infolabel = Label(infoFrame, text="All fields are mandatory !!", width=40, anchor='center',
                          justify=LEFT,
                          font=NORM_VERDANA_FONT, bg='snow', fg="black")

        heading.grid(row=0, column=0, columnspan=2)
        dataEntryFrame.grid(row=1, column=1, padx=10, pady=8)
        pledge_item.grid(row=0, column=0, pady=5)
        dateofpledge.grid(row=1, column=0, pady=5)
        trust_name.grid(row=2, column=0, pady=5)
        managerlabel.grid(row=3, column=0, pady=5)
        addresslabel.grid(row=4, column=0, pady=5)

        # create a text entry box
        # for typing the information
        pledge_text = Entry(dataEntryFrame, width=30, font=NORM_FONT, bg='light yellow',
                            textvariable=self.default_text1)

        trust_nametext = StringVar(dataEntryFrame)
        trustName_list = ['Vihangam Yoga (Karnataka) Trust', 'Aadarsh gaushala Trust']
        trust_nametext.set("Vihangam Yoga (Karnataka) Trust")
        trust_menu = OptionMenu(dataEntryFrame, trust_nametext, *trustName_list)
        trust_menu.configure(width=26, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)

        cal = DateEntry(dataEntryFrame, width=28, font=NORM_FONT, date_pattern='dd/MM/yyyy', bg='light yellow',
                        anchor=W, justify=LEFT)
        manager_text = Entry(dataEntryFrame, width=30, font=NORM_FONT, bg='light yellow',
                             textvariable=self.default_text2)
        address_text = Entry(dataEntryFrame, width=30, font=NORM_FONT, bg='light yellow',
                             textvariable=self.default_text3)

        pledge_text.grid(row=0, column=1, pady=5)
        cal.grid(row=1, column=1, pady=5)
        trust_menu.grid(row=2, column=1, pady=5)
        manager_text.grid(row=3, column=1, pady=5)
        address_text.grid(row=4, column=1, pady=5)
        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(self.new_center_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=1, pady=8)
        submit_deposit = Button(buttonFrame)

        insert_result = partial(self.registerlocalCenter, trust_nametext, pledge_text, infolabel)

        # create a Save Button and place into the self.new_center_window window
        submit_deposit.configure(text="Register", fg="Black", command=insert_result,
                                 font=NORM_FONT, width=8, bg='light cyan', underline=0, state=NORMAL)
        submit_deposit.grid(row=0, column=0)

        clear_result = partial(self.clearRegisterPledgeForm,
                               pledge_text,
                               trust_nametext, infolabel)

        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=1)

        # create a Cancel Button and place into the self.new_center_window window
        cancel_Result = partial(self.destroyWindow, self.new_center_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_Result,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=2)
        # ---------------------------------Button Frame End----------------------------------------

        infoFrame.grid(row=9, column=1, pady=5)
        infolabel.grid(row=0, column=0, padx=2, pady=3)

        self.new_center_window.bind('<Return>', lambda event=None: submit_deposit.invoke())
        self.new_center_window.bind('<Alt-d>', lambda event=None: submit_deposit.invoke())
        self.new_center_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.new_center_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.new_center_window.focus()
        self.new_center_window.grab_set()
        mainloop()

    def clearRegisterPledgeForm(self, pledge_text,
                                trust_nametext, infolabel):
        pledge_text.delete(0, END)
        pledge_text.configure(fg='black')
        pledge_text.focus_set()
        trust_nametext.set("Vihangam Yoga (Karnataka) Trust")
        infolabel['fg'] = "black"
        infolabel['text'] = "All fields are mandatory !!!"

    def clearauthorregisterform(self, author_name):
        author_name.delete(0, END)
        author_name.configure(fg='black')
        author_name.focus_set()

    def registerPledgeItem(self, trust_nametext, pledge_item):
        self.obj_commonUtil.registerPledgeItem(pledge_item)
        self.obj_initDatabase.initilize_pledgeitem_database(trust_nametext, pledge_item)

    def registerlocalCenter(self, trust_nametext, pledge_item, infolabel):
        self.obj_commonUtil.registerlocalCenter(pledge_item, infolabel)

    def register_new_author(self):
        self.new_author_window = Toplevel(self.master)
        self.new_author_window.title("New Author Registration ")
        self.new_author_window.geometry('420x230+240+200')
        self.new_author_window.configure(background='wheat')
        self.new_author_window.resizable(width=False, height=False)
        self.new_author_window.protocol('WM_DELETE_WINDOW', self.donothing)
        heading = Label(self.new_author_window, text="Register New Author",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        dataEntryFrame = Frame(self.new_author_window, width=200, height=100, bd=4, relief='ridge',
                               bg='snow')
        self.default_text1 = StringVar(dataEntryFrame, value='')
        self.default_text2 = StringVar(dataEntryFrame, value='')

        # lower frame added to show the result of transactions
        infoFrame = Frame(self.new_author_window, width=100, height=20, bd=4, relief='ridge')
        # create a Book Name label
        pledge_item = Label(dataEntryFrame, text="Author Name", width=15, anchor=W, justify=LEFT,
                            font=NORM_FONT,
                            bg='snow')

        dateofpledge = Label(dataEntryFrame, text="Registration Date", width=15, anchor=W,
                             justify=LEFT,
                             font=NORM_FONT, bg='snow')

        infolabel = Label(infoFrame, text="All fields are mandatory !!", width=40, anchor='center',
                          justify=LEFT,
                          font=NORM_VERDANA_FONT, bg='snow')

        heading.grid(row=0, column=0, columnspan=2)
        dataEntryFrame.grid(row=1, column=1, padx=10, pady=8)
        pledge_item.grid(row=0, column=0, pady=5)
        dateofpledge.grid(row=1, column=0, pady=5)

        # create a text entry box
        # for typing the information
        author_name = Entry(dataEntryFrame, width=25, font=NORM_FONT, bg='light yellow',
                            textvariable=self.default_text1)

        cal = DateEntry(dataEntryFrame, width=23, font=NORM_FONT, date_pattern='dd/MM/yyyy', bg='light yellow',
                        anchor=W, justify=LEFT)

        author_name.grid(row=0, column=1, pady=5)
        cal.grid(row=1, column=1, pady=5)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(self.new_author_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=1, pady=8)
        submit_deposit = Button(buttonFrame)

        insert_result = partial(self.registernewauthor, author_name, infolabel)

        # create a Save Button and place into the self.new_author_window window
        submit_deposit.configure(text="Register", fg="Black", command=insert_result,
                                 font=NORM_FONT, width=8, bg='light cyan', underline=0, state=NORMAL)
        submit_deposit.grid(row=0, column=0)

        clear_result = partial(self.clearauthorregisterform, author_name)

        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=1)

        # create a Cancel Button and place into the self.new_author_window window
        cancel_Result = partial(self.destroyWindow, self.new_author_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_Result,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=2)
        # ---------------------------------Button Frame End----------------------------------------

        infoFrame.grid(row=9, column=1, pady=5)
        infolabel.grid(row=0, column=0, padx=2, pady=3)

        self.new_author_window.bind('<Return>', lambda event=None: submit_deposit.invoke())
        self.new_author_window.bind('<Alt-d>', lambda event=None: submit_deposit.invoke())
        self.new_author_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.new_author_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.new_author_window.focus()
        self.new_author_window.grab_set()
        mainloop()

    def registernewauthor(self, author_name, infolabel):
        self.obj_commonUtil.registerNewAuthor(author_name)
        infolabel.configure(text=" Author registration successful !!!", fg='green')

    def sankalp_form(self):
        self.new_sankalp_window = Toplevel(self.master)
        self.new_sankalp_window.title("Deposit Seva Rashi ")
        self.new_sankalp_window.geometry('800x355+240+200')
        self.new_sankalp_window.configure(background='wheat')
        self.new_sankalp_window.resizable(width=False, height=False)
        self.new_sankalp_window.protocol('WM_DELETE_WINDOW', self.donothing)
        heading = Label(self.new_sankalp_window, text="Create New Seva-Sankalp",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        dataEntryFrame = Frame(self.new_sankalp_window, width=200, height=100, bd=4, relief='ridge',
                               bg='snow')
        self.default_text1 = StringVar(dataEntryFrame, value='')
        self.default_text2 = StringVar(dataEntryFrame, value='')
        self.default_text3 = StringVar(dataEntryFrame, value='')
        self.default_text4 = StringVar(dataEntryFrame, value='')
        self.default_text5 = StringVar(dataEntryFrame, value='None')
        # lower frame added to show the result of transactions
        infoFrame = Frame(self.new_sankalp_window, width=100, height=20, bd=4, relief='ridge')
        # create a Book Name label
        donator_name = Label(dataEntryFrame, text="Member Id", width=15, anchor=W, justify=LEFT,
                             font=NORM_FONT,
                             bg='snow')

        # create a Author label
        seva_amount = Label(dataEntryFrame, text="Pledge Amt.(Rs.)", width=15, anchor=W, justify=LEFT,
                            font=NORM_FONT,
                            bg='snow')

        # create a Price label
        trust_name = Label(dataEntryFrame, text="Trust Name", width=15, anchor=W, justify=LEFT,
                           font=NORM_FONT, bg='snow')

        dateofpledge = Label(dataEntryFrame, text="Pledge Date", width=15, anchor=W,
                             justify=LEFT,
                             font=NORM_FONT, bg='snow')

        # create a Quantity label
        coordinator_id = Label(dataEntryFrame, text="Coordinator (Id)", width=15, anchor=W, justify=LEFT,
                               font=NORM_FONT, bg='snow')

        # create a borrow fee label
        modeOfpayment = Label(dataEntryFrame, text="Duration(Year)", width=15, anchor=W, justify=LEFT,
                              font=NORM_FONT, bg='snow')

        sankalp_period = Label(dataEntryFrame, text="Payment Period", width=15, anchor=W, justify=LEFT,
                               font=NORM_FONT, bg='snow')

        sankalp_purpose = Label(dataEntryFrame, text="Pledge For", width=15, anchor=W, justify=LEFT,
                                font=NORM_FONT, bg='snow')

        # create a borrow fee label
        balance_amt = Label(dataEntryFrame, text="Balance(Rs.)", width=15, anchor=W, justify=LEFT,
                            font=NORM_FONT, bg='snow')

        infolabel = Label(infoFrame, text="All fields are mandatory !!", width=50, anchor='center',
                          justify=LEFT,
                          font=NORM_VERDANA_FONT, bg='snow')

        heading.grid(row=0, column=0, columnspan=2)
        dataEntryFrame.grid(row=1, column=1, padx=10, pady=8)
        donator_name.grid(row=0, column=0, pady=5)
        seva_amount.grid(row=0, column=2, padx=10, pady=5)
        trust_name.grid(row=1, column=0, pady=5)
        dateofpledge.grid(row=1, column=2, pady=5)
        coordinator_id.grid(row=2, column=0, pady=5)
        modeOfpayment.grid(row=2, column=2, pady=5)
        sankalp_period.grid(row=3, column=0, pady=5)
        sankalp_purpose.grid(row=3, column=2, pady=5)
        balance_amt.grid(row=4, column=0, pady=5)

        # create a text entry box
        # for typing the information
        donator_idText = Entry(dataEntryFrame, width=25, font=NORM_FONT, bg='light yellow',
                               textvariable=self.default_text1)
        pleadge_amountText = Entry(dataEntryFrame, width=25, font=NORM_FONT, bg='light yellow',
                                   textvariable=self.default_text2)
        trust_nametext = StringVar(dataEntryFrame)
        trustName_list = ['Vihangam Yoga (Karnataka) Trust', 'Aadarsh gaushala Trust']
        trust_nametext.set("Vihangam Yoga (Karnataka) Trust")
        trust_menu = OptionMenu(dataEntryFrame, trust_nametext, *trustName_list)
        trust_menu.configure(width=21, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)
        cal = DateEntry(dataEntryFrame, width=23, font=NORM_FONT, date_pattern='dd/MM/yyyy', bg='light yellow',
                        anchor=W, justify=LEFT)

        coordinator_text = Entry(dataEntryFrame, width=25, font=NORM_FONT, bg='light yellow',
                                 textvariable=self.default_text3)

        duration_text = Entry(dataEntryFrame, width=25, font=NORM_FONT, bg='light yellow',
                              textvariable=self.default_text4)

        paymentduration_text = StringVar(dataEntryFrame)
        paymentMode_list = ['Monthly', 'Quarterly', 'Yearly', 'Random']
        paymentduration_text.set("Monthly")
        paymentMode_menu = OptionMenu(dataEntryFrame, paymentduration_text, *paymentMode_list)
        paymentMode_menu.configure(width=21, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)

        pledge_itemtext = StringVar(dataEntryFrame)

        pledgeItemName_list = self.obj_commonUtil.getPledgeItemNames()
        print("Pledge item list  - ", pledgeItemName_list)
        pledge_itemtext.set(pledgeItemName_list[0])
        pledgeItem_menu = OptionMenu(dataEntryFrame, pledge_itemtext, *pledgeItemName_list)
        pledgeItem_menu.configure(width=21, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)
        remBalance_text = Label(dataEntryFrame, text="------------", width=16, anchor=W, justify=LEFT,
                                font=NORM_FONT, bg='snow')

        donator_idText.grid(row=0, column=1, pady=5)
        pleadge_amountText.grid(row=0, column=3, pady=5)
        trust_menu.grid(row=1, column=1, pady=5)
        cal.grid(row=1, column=3, pady=5)
        coordinator_text.grid(row=2, column=1, pady=5)
        duration_text.grid(row=2, column=3, pady=5)
        paymentMode_menu.grid(row=3, column=1, pady=5)
        pledgeItem_menu.grid(row=3, column=3, pady=5)
        remBalance_text.grid(row=4, column=1, pady=5)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(self.new_sankalp_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=1, pady=8)
        print_invoice = Button(buttonFrame, text="Print Receipt", fg="Black", command=None,
                               font=NORM_FONT, width=10, bg='light grey', state=DISABLED)
        self.submit_deposit = Button(buttonFrame)

        insert_result = partial(self.update_pledge_detail,
                                self.new_sankalp_window,
                                donator_idText,
                                pleadge_amountText,
                                trust_nametext,
                                coordinator_text,
                                cal,
                                duration_text,
                                paymentduration_text,
                                pledge_itemtext, remBalance_text,
                                infolabel)

        # create a Save Button and place into the self.new_sankalp_window window
        self.submit_deposit.configure(text="Pledge", fg="Black", command=insert_result,
                                      font=NORM_FONT, width=8, bg='light grey', underline=0, state=DISABLED)
        self.submit_deposit.grid(row=0, column=0)
        print_invoice.grid(row=0, column=1)
        self.default_text1.trace("w", self.check_DepositBtn_state)
        self.default_text2.trace("w", self.check_DepositBtn_state)
        self.default_text3.trace("w", self.check_DepositBtn_state)
        self.default_text4.trace("w", self.check_DepositBtn_state)

        # create a Clear Button and place into the self.new_sankalp_window window
        # clear the commercial form after data has been saved

        clear_result = partial(self.clearSankalpForm,
                               donator_idText,
                               pleadge_amountText,
                               trust_nametext,
                               coordinator_text,
                               cal,
                               duration_text,
                               paymentduration_text,
                               pledge_itemtext, remBalance_text,
                               infolabel)

        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=2)

        # create a Cancel Button and place into the self.new_sankalp_window window
        cancel_Result = partial(self.destroyWindow, self.new_sankalp_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_Result,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=3)
        # ---------------------------------Button Frame End----------------------------------------

        infoFrame.grid(row=9, column=1, pady=5)
        infolabel.grid(row=0, column=0, padx=2, pady=3)

        self.new_sankalp_window.bind('<Return>', lambda event=None: self.submit_deposit.invoke())
        self.new_sankalp_window.bind('<Alt-d>', lambda event=None: self.submit_deposit.invoke())
        self.new_sankalp_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.new_sankalp_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.new_sankalp_window.focus()
        self.new_sankalp_window.grab_set()
        mainloop()

    def clearSankalpForm(self, donator_idText,
                         pleadge_amountText,
                         trust_nametext,
                         coordinator_text,
                         cal,
                         duration_text,
                         paymentduration_text,
                         pledge_itemtext, remBalance_text,
                         infolabel):

        donator_idText.delete(0, END)
        donator_idText.configure(fg='black')
        donator_idText.focus_set()
        pleadge_amountText.delete(0, END)
        pleadge_amountText.configure(fg='black')
        coordinator_text.delete(0, END)
        coordinator_text.configure(fg='black')
        duration_text.delete(0, END)
        duration_text.configure(fg='black')
        remBalance_text['text'] = "Rs. 0"

    def perform_patrika_subscription(self):
        self.new_akshay_allocation_window = Toplevel(self.master)
        self.new_akshay_allocation_window.title("Magazine Subscription ")
        self.new_akshay_allocation_window.geometry('730x350+240+200')
        self.new_akshay_allocation_window.configure(background='wheat')
        self.new_akshay_allocation_window.resizable(width=False, height=False)
        self.new_akshay_allocation_window.protocol('WM_DELETE_WINDOW', self.donothing)
        heading = Label(self.new_akshay_allocation_window, text="Magazine Subscription",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        dataEntryFrame = Frame(self.new_akshay_allocation_window, width=200, height=100, bd=4, relief='ridge',
                               bg='snow')
        self.default_text1 = StringVar(dataEntryFrame, value='')
        self.default_text2 = StringVar(dataEntryFrame, value='')
        self.default_text3 = StringVar(dataEntryFrame, value='')
        self.default_text4 = StringVar(dataEntryFrame, value='')
        self.default_text5 = StringVar(dataEntryFrame, value='None')
        # lower frame added to show the result of transactions
        infoFrame = Frame(self.new_akshay_allocation_window, width=100, height=20, bd=4, relief='ridge')
        # create a Book Name label
        member_idlabel = Label(dataEntryFrame, text="Member Id", width=13, anchor=W, justify=LEFT,
                               font=NORM_FONT,
                               bg='snow')

        dateOfSubscription = Label(dataEntryFrame, text="Subscription Date", width=14, anchor=W,
                                   justify=LEFT,
                                   font=NORM_FONT, bg='snow')

        # create a Price label
        magaiznename_label = Label(dataEntryFrame, text="Magazine", width=13, anchor=W, justify=LEFT,
                                   font=NORM_FONT, bg='snow')

        # create a Author label
        quantity_label = Label(dataEntryFrame, text="Quantity", width=14, anchor=W, justify=LEFT,
                               font=NORM_FONT,
                               bg='snow')

        # create a Quantity label
        amount_label = Label(dataEntryFrame, text="Unit Price(Rs.)", width=13, anchor=W, justify=LEFT,
                             font=NORM_FONT, bg='snow')

        # create a borrow fee label
        modeOfpayment = Label(dataEntryFrame, text="Payment Mode", width=14, anchor=W, justify=LEFT,
                              font=NORM_FONT, bg='snow')

        authorizedBy_name = Label(dataEntryFrame, text="Authorizor (Id)", width=13, anchor=W, justify=LEFT,
                                  font=NORM_FONT, bg='snow')

        infolabel = Label(infoFrame, text="All fields are mandatory !!", width=50, anchor='center',
                          justify=LEFT,
                          font=NORM_VERDANA_FONT, bg='snow')

        # create a borrow fee label
        transrefid_id = Label(dataEntryFrame, text="Transaction ID", width=14, anchor=W, justify=LEFT,
                              font=NORM_FONT, bg='snow')

        payableAmt_id = Label(dataEntryFrame, text="Payable Amt(Rs.)", width=13, anchor=W, justify=LEFT,
                              font=NORM_FONT, bg='snow')

        heading.grid(row=0, column=0, columnspan=2)
        dataEntryFrame.grid(row=1, column=1, padx=10, pady=8)
        member_idlabel.grid(row=0, column=0, pady=3)
        dateOfSubscription.grid(row=0, column=2, padx=5, pady=3)
        magaiznename_label.grid(row=1, column=0, padx=10, pady=3)
        quantity_label.grid(row=1, column=2, padx=5, pady=3)
        amount_label.grid(row=2, column=0, pady=3)
        modeOfpayment.grid(row=2, column=2, padx=5, pady=3)
        authorizedBy_name.grid(row=3, column=0, pady=3)
        transrefid_id.grid(row=3, column=2, pady=3)
        payableAmt_id.grid(row=4, column=0, padx=5, pady=3)

        # create a text entry box
        # for typing the information
        member_idText = Entry(dataEntryFrame, width=22, font=NORM_FONT, bg='light yellow',
                              textvariable=self.default_text1)
        cal = DateEntry(dataEntryFrame, width=20, font=NORM_FONT, date_pattern='dd/MM/yyyy', bg='light yellow',
                        anchor=W, justify=LEFT)

        magazineCategoryText = StringVar(dataEntryFrame)
        category_list = ['Vihangam Yog Sandesh', 'Arogya Vigyan', 'Vihangam Yoga Times', 'Other']
        magazineCategoryText.set("Vihangam Yog Sandesh")
        category_menu = OptionMenu(dataEntryFrame, magazineCategoryText, *category_list)

        quantityText = Entry(dataEntryFrame, width=22, font=NORM_FONT, bg='light yellow',
                             textvariable=self.default_text2)
        seva_amountText = Entry(dataEntryFrame, width=22, font=NORM_FONT, bg='light yellow',
                                textvariable=self.default_text3)
        paymentMode_text = StringVar(dataEntryFrame)
        paymentMode_list = ['Cash', 'Bank Transfer', 'Cheque', 'Demand Draft', 'Other']
        paymentMode_text.set("Other")
        paymentMode_menu = OptionMenu(dataEntryFrame, paymentMode_text, *paymentMode_list)
        paymentMode_menu.configure(width=18, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)

        authorizedby_idText = Entry(dataEntryFrame, width=22, font=NORM_FONT, bg='light yellow',
                                    textvariable=self.default_text4)
        transref_IdText = Entry(dataEntryFrame, width=22, font=NORM_FONT, bg='light yellow')
        payableAmtText = Label(dataEntryFrame, text="------------", width=18, anchor=W, justify=LEFT,
                               font=NORM_FONT, bg='snow')

        member_idText.grid(row=0, column=1, pady=3)
        cal.grid(row=0, column=3, pady=3)

        category_menu.configure(width=18, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)
        category_menu.grid(row=1, column=1, pady=3)

        quantityText.grid(row=1, column=3, pady=3)
        seva_amountText.grid(row=2, column=1, pady=3)
        paymentMode_menu.grid(row=2, column=3, pady=3)
        authorizedby_idText.grid(row=3, column=1, pady=3)

        transref_IdText.grid(row=3, column=3, pady=3)
        payableAmtText.grid(row=4, column=1, pady=3)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(self.new_akshay_allocation_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=1, pady=8)
        print_invoice = Button(buttonFrame, text="Print Receipt", fg="Black", command=None,
                               font=NORM_FONT, width=10, bg='light grey', state=DISABLED)
        self.submit_deposit = Button(buttonFrame)
        insert_result = partial(self.subscribe_magazine_Excel,
                                self.new_akshay_allocation_window,
                                member_idText,
                                seva_amountText,
                                magazineCategoryText,
                                quantityText,
                                cal,
                                paymentMode_text,
                                authorizedby_idText,
                                payableAmtText,
                                infolabel, print_invoice)

        # create a Save Button and place into the self.new_akshay_allocation_window window
        self.submit_deposit.configure(text="Subscribe", fg="Black", command=insert_result,
                                      font=NORM_FONT, width=8, bg='light grey', underline=0, state=DISABLED)
        self.submit_deposit.grid(row=0, column=0)
        print_invoice.grid(row=0, column=1)
        self.default_text1.trace("w", self.check_DepositBtn_state)
        self.default_text2.trace("w", self.check_DepositBtn_state)
        self.default_text3.trace("w", self.check_DepositBtn_state)
        self.default_text4.trace("w", self.check_DepositBtn_state)

        # create a Clear Button and place into the self.new_akshay_allocation_window window
        # clear the commercial form after data has been saved

        clear_result = partial(self.clearMagazineSubscriptionForm,
                               member_idText,
                               seva_amountText,
                               magazineCategoryText,
                               quantityText,
                               cal,
                               paymentMode_text,
                               authorizedby_idText,
                               payableAmtText,
                               infolabel, print_invoice)

        clear = Button(buttonFrame, text="Reset", fg="Black", command=None,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=2)

        # create a Cancel Button and place into the self.new_akshay_allocation_window window
        cancel_Result = partial(self.destroyWindow, self.new_akshay_allocation_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_Result,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=3)
        # ---------------------------------Button Frame End----------------------------------------

        infoFrame.grid(row=9, column=1, pady=5)
        infolabel.grid(row=0, column=0, padx=2, pady=3)

        self.new_akshay_allocation_window.bind('<Return>', lambda event=None: submit.invoke())
        self.new_akshay_allocation_window.bind('<Alt-d>', lambda event=None: submit.invoke())
        self.new_akshay_allocation_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.new_akshay_allocation_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.new_akshay_allocation_window.focus()
        self.new_akshay_allocation_window.grab_set()
        mainloop()

    def check_for_selection(self, n, m, x, categoryText, akshayPatra_Text, donator_idText):
        print("check_for_selection for Akshay Patra :", n.get())

        if n.get() == "Akshay-Patra Seva":
            donator_records = self.retrieve_MemberRecords_Excel(x.get(), 1, SEARCH_BY_MEMBERID)
            print(" Akshay Patra no is : ", donator_records[24])
            m.configure(state=NORMAL)
            m.delete(0, END)
            if donator_records[24] is not None:
                m.insert(0, donator_records[24])
            else:
                m.insert(0, "Not Available")
            m.configure(state=DISABLED)
        else:
            m.configure(state=NORMAL)
            m.delete(0, END)
            m.configure(state=DISABLED)

    def transid_enable(self, n, m, x, paymentMode, transactionId_text):
        print("check_for_transid_enable :", n.get())
        if n.get() == "Bank Transfer":
            m.configure(state=NORMAL)
            m.delete(0, END)
        else:
            m.delete(0, END)
            m.configure(state=DISABLED)

    def deposit_seva_nonmonetary_rashi(self, ownerType):
        self.new_sevaRashi_deposit_window = Toplevel(self.master)
        self.new_sevaRashi_deposit_window.title("Non-Monetary Inventory")
        self.new_sevaRashi_deposit_window.geometry('990x405+200+200')
        self.new_sevaRashi_deposit_window.configure(background='wheat')
        self.new_sevaRashi_deposit_window.resizable(width=False, height=False)
        self.new_sevaRashi_deposit_window.protocol('WM_DELETE_WINDOW', self.donothing)

        heading = Label(self.new_sevaRashi_deposit_window, text="Non-Commercial Stock/Donation Entry",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        dataEntryFrame = Frame(self.new_sevaRashi_deposit_window, width=200, height=100, bd=4, relief='ridge',
                               bg='snow')
        self.default_text1 = StringVar(dataEntryFrame, value='')
        self.default_text2 = StringVar(dataEntryFrame, value='')
        self.default_text3 = StringVar(dataEntryFrame, value='')
        self.default_text4 = StringVar(dataEntryFrame, value='')
        self.default_text5 = StringVar(dataEntryFrame, value='')

        # lower frame added to show the result of transactions
        infoFrame = Frame(self.new_sevaRashi_deposit_window, width=100, height=20, bd=4, relief='ridge')
        # create a Book Name label
        donator_name = Label(dataEntryFrame, text="Received From(Id)", width=14, anchor=W, justify=LEFT,
                             font=TIMES_NEW_ROMAN_BIG,
                             bg='snow')

        # create a Author label
        donated_item = Label(dataEntryFrame, text="Item Name", width=13, anchor=W, justify=LEFT,
                             font=TIMES_NEW_ROMAN_BIG,
                             bg='snow')

        # create a Price label
        quantity = Label(dataEntryFrame, text="Quantity", width=13, anchor=W, justify=LEFT,
                         font=TIMES_NEW_ROMAN_BIG, bg='snow')

        dateOfCollection = Label(dataEntryFrame, text="Date", width=13, anchor=W,
                                 justify=LEFT,
                                 font=TIMES_NEW_ROMAN_BIG, bg='snow')

        # create a Quantity label
        collector_name = Label(dataEntryFrame, text="Receiver (Id)", width=13, anchor=W, justify=LEFT,
                               font=TIMES_NEW_ROMAN_BIG, bg='snow')

        # create a borrow fee label
        estimatedValue = Label(dataEntryFrame, text="Est. Value(Rs.)", width=13, anchor=W, justify=LEFT,
                               font=TIMES_NEW_ROMAN_BIG, bg='snow')

        authorizor_id = Label(dataEntryFrame, text="Authorized By(Id)", width=13, anchor=W, justify=LEFT,
                              font=TIMES_NEW_ROMAN_BIG, bg='snow')

        location_label = Label(dataEntryFrame, text="Rack No.", width=13, anchor=W, justify=LEFT,
                               font=TIMES_NEW_ROMAN_BIG, bg='snow')

        invoice_idLabel = Label(dataEntryFrame, text="Invoice-Id", width=13, anchor=W, justify=LEFT,
                                font=TIMES_NEW_ROMAN_BIG, bg='snow')

        center_name = Label(dataEntryFrame, text="Center Name", width=13, anchor=W, justify=LEFT,
                            font=TIMES_NEW_ROMAN_BIG, bg='snow')

        infolabel = Label(infoFrame, text="All fields are mandatory !!", width=60, anchor='center',
                          justify=LEFT,
                          font=TIMES_NEW_ROMAN_BIG, bg='snow')

        heading.grid(row=0, column=0, columnspan=2)
        dataEntryFrame.grid(row=1, column=1, padx=10, pady=8)
        donator_name.grid(row=0, column=0, pady=7)
        donated_item.grid(row=0, column=2, padx=10, pady=7)
        quantity.grid(row=1, column=0, pady=7)
        dateOfCollection.grid(row=1, column=2, pady=7)
        collector_name.grid(row=2, column=0, pady=7)
        estimatedValue.grid(row=2, column=2, pady=7)
        authorizor_id.grid(row=3, column=0, pady=7)
        location_label.grid(row=3, column=2, pady=7)
        invoice_idLabel.grid(row=4, column=0, pady=7)
        center_name.grid(row=4, column=2, pady=7)

        # create a text entry box
        # for typing the information
        donator_idText = Entry(dataEntryFrame, width=27, font=TIMES_NEW_ROMAN_BIG, bg='light yellow',
                               textvariable=self.default_text1)

        seva_amountText = Entry(dataEntryFrame, width=27, font=TIMES_NEW_ROMAN_BIG, bg='light yellow',
                                textvariable=self.default_text2)
        quantityText = Entry(dataEntryFrame, width=27, font=TIMES_NEW_ROMAN_BIG, bg='light yellow',
                             textvariable=self.default_text3)
        cal = DateEntry(dataEntryFrame, width=25, font=TIMES_NEW_ROMAN_BIG, date_pattern='dd/MM/yyyy',
                        bg='light yellow',
                        anchor=W, justify=LEFT)

        collector_nameText = Entry(dataEntryFrame, width=27, font=TIMES_NEW_ROMAN_BIG, bg='light yellow',
                                   textvariable=self.default_text4)

        estValueText = Entry(dataEntryFrame, width=27, font=TIMES_NEW_ROMAN_BIG, bg='light yellow',
                             textvariable=self.default_text5)
        authorizorId_Text = Entry(dataEntryFrame, width=27, font=TIMES_NEW_ROMAN_BIG, bg='light yellow')
        location_text = Entry(dataEntryFrame, width=27, font=TIMES_NEW_ROMAN_BIG, bg='light yellow')
        invoice_idText = Label(dataEntryFrame, text="------------", width=27, anchor=W, justify=LEFT,
                               font=NORM_FONT, bg='snow')

        local_centerText = StringVar(dataEntryFrame)
        localCenterList = self.obj_commonUtil.getLocalCenterNames()
        print("Center list  - ", localCenterList)
        local_centerText.set(localCenterList[0])
        localcenter_menu = OptionMenu(dataEntryFrame, local_centerText, *localCenterList)
        localcenter_menu.configure(width=24, font=TIMES_NEW_ROMAN_BIG, bg='light yellow', anchor=W, justify=LEFT)

        donator_idText.grid(row=0, column=1, pady=7)
        seva_amountText.grid(row=0, column=3, pady=7)

        quantityText.grid(row=1, column=1, pady=7)
        cal.grid(row=1, column=3, pady=7)
        collector_nameText.grid(row=2, column=1, pady=7)
        estValueText.grid(row=2, column=3, pady=7)
        authorizorId_Text.grid(row=3, column=1, pady=7)
        location_text.grid(row=3, column=3, pady=7)
        invoice_idText.grid(row=4, column=1, pady=7)
        localcenter_menu.grid(row=4, column=3, pady=7)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(self.new_sevaRashi_deposit_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=1, pady=7)
        print_invoice = Button(buttonFrame, text="Print Receipt", fg="Black", command=None,
                               font=TIMES_NEW_ROMAN_BIG, width=10, bg='light grey', state=DISABLED)
        self.submit_deposit = Button(buttonFrame)
        insert_result = partial(self.deposit_non_monetary_Excel,
                                self.new_sevaRashi_deposit_window,
                                donator_idText,
                                seva_amountText,
                                quantityText,
                                collector_nameText,
                                authorizorId_Text,
                                cal,
                                estValueText,
                                invoice_idText,
                                infolabel,
                                print_invoice,
                                self.submit_deposit,
                                ownerType,
                                location_text, local_centerText)
        self.submit_deposit.configure(text="Deposit", fg="Black", command=insert_result,
                                      font=TIMES_NEW_ROMAN_BIG, width=8, bg='light grey', underline=0, state=DISABLED)

        self.default_text1.trace("w", self.check_DepositBtn_state)
        self.default_text2.trace("w", self.check_DepositBtn_state)
        self.default_text3.trace("w", self.check_DepositBtn_state)
        self.default_text4.trace("w", self.check_DepositBtn_state)
        self.default_text5.trace("w", self.check_DepositBtn_state)
        # create a Save Button and place into the self.new_sevaRashi_deposit_window window

        self.submit_deposit.grid(row=0, column=0)
        print_invoice.grid(row=0, column=1)
        # create a Clear Button and place into the self.new_sevaRashi_deposit_window window
        # clear the commercial form after data has been saved
        clear_result = partial(self.clearNonMonetaryDonationForm, donator_idText,
                               seva_amountText,
                               quantityText,
                               collector_nameText,
                               cal,
                               estValueText,
                               invoice_idText,
                               infolabel, print_invoice)
        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=TIMES_NEW_ROMAN_BIG, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=2)

        # create a Cancel Button and place into the self.new_sevaRashi_deposit_window window
        cancel_Result = partial(self.destroyWindow, self.new_sevaRashi_deposit_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_Result,
                        font=TIMES_NEW_ROMAN_BIG, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=3)
        # ---------------------------------Button Frame End----------------------------------------

        infoFrame.grid(row=9, column=1, pady=5)
        infolabel.grid(row=0, column=0, padx=2, pady=7)

        self.new_sevaRashi_deposit_window.bind('<Return>', lambda event=None: self.submit_deposit.invoke())
        self.new_sevaRashi_deposit_window.bind('<Alt-d>', lambda event=None: self.submit_deposit.invoke())
        self.new_sevaRashi_deposit_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.new_sevaRashi_deposit_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.new_sevaRashi_deposit_window.focus()
        self.new_sevaRashi_deposit_window.grab_set()
        mainloop()

    def create_expanse_entry(self):
        self.new_expanse_window = Toplevel(self.master)
        self.new_expanse_window.title("Create Ashram Expanse ")
        self.new_expanse_window.geometry('775x380+240+200')
        self.new_expanse_window.configure(background='wheat')
        self.new_expanse_window.resizable(width=False, height=False)
        self.new_expanse_window.protocol('WM_DELETE_WINDOW', self.donothing)
        heading = Label(self.new_expanse_window, text="Create Ashram Expanse Form",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        dataEntryFrame = Frame(self.new_expanse_window, width=220, height=100, bd=4, relief='ridge',
                               bg='snow')
        self.default_text1 = StringVar(dataEntryFrame, value='')
        self.default_text2 = StringVar(dataEntryFrame, value='')
        self.default_text3 = StringVar(dataEntryFrame, value='')
        self.default_text4 = StringVar(dataEntryFrame, value='')
        self.default_text5 = StringVar(dataEntryFrame, value='None')
        self.default_text6 = StringVar(dataEntryFrame, value='')
        self.default_text7 = StringVar(dataEntryFrame, value='')

        # lower frame added to show the result of transactions
        infoFrame = Frame(self.new_expanse_window, width=100, height=20, bd=4, relief='ridge')
        # create a Book Name label

        receiver_Id = Label(dataEntryFrame, text="Paid To(Id)", width=13, anchor=W, justify=LEFT,
                            font=NORM_FONT,
                            bg='snow')
        receiver_name = Label(dataEntryFrame, text="Paid To(Name)", width=13, anchor=W, justify=LEFT,
                              font=NORM_FONT,
                              bg='snow')
        # create a Price label
        description = Label(dataEntryFrame, text="Description", width=13, anchor=W, justify=LEFT,
                            font=NORM_FONT, bg='snow')
        receiver_phoneno = Label(dataEntryFrame, text="Phone No.", width=13, anchor=W, justify=LEFT,
                                 font=NORM_FONT,
                                 bg='snow')
        authorized_name = Label(dataEntryFrame, text="Authorized By", width=13, anchor=W, justify=LEFT,
                                font=NORM_FONT, bg='snow')

        # create a Author label
        seva_amount = Label(dataEntryFrame, text="Amount(Rs.)", width=13, anchor=W, justify=LEFT,
                            font=NORM_FONT,
                            bg='snow')

        dateOfexpanse = Label(dataEntryFrame, text="Date", width=13, anchor=W,
                              justify=LEFT,
                              font=NORM_FONT, bg='snow')

        # create a borrow fee label
        modeOfpayment = Label(dataEntryFrame, text="Payment Mode", width=13, anchor=W, justify=LEFT,
                              font=NORM_FONT, bg='snow')

        infolabel = Label(infoFrame, text="Please enter Expanse details", width=50, anchor='center',
                          justify=LEFT,
                          font=NORM_VERDANA_FONT, bg='snow')

        trust_label = Label(dataEntryFrame, text="Trust Account", width=13, anchor=W, justify=LEFT,
                            font=NORM_FONT, bg='snow')

        # create a borrow fee label
        invoice_id = Label(dataEntryFrame, text="Invoice Id", width=13, anchor=W, justify=LEFT,
                           font=NORM_FONT, bg='snow')

        heading.grid(row=0, column=0, columnspan=2)
        dataEntryFrame.grid(row=1, column=1, padx=10, pady=8)

        receiver_Id.grid(row=1, column=0, pady=3)
        receiver_name.grid(row=1, column=2, pady=3)
        description.grid(row=2, column=0, pady=3)
        receiver_phoneno.grid(row=2, column=2, pady=3)
        authorized_name.grid(row=3, column=0, pady=3)
        seva_amount.grid(row=3, column=2, padx=10, pady=3)
        dateOfexpanse.grid(row=4, column=2, pady=3)
        modeOfpayment.grid(row=5, column=2, pady=3)
        trust_label.grid(row=4, column=0, pady=3)
        invoice_id.grid(row=5, column=0, pady=3)

        # create a text entry box
        # for typing the information
        receiver_IDText = Entry(dataEntryFrame, width=23, font=NORM_FONT, bg='light yellow',
                                textvariable=self.default_text1)
        receiver_nameText = Entry(dataEntryFrame, width=23, font=NORM_FONT, bg='light yellow',
                                  textvariable=self.default_text6, state=DISABLED)
        descriptionText = Entry(dataEntryFrame, width=23, font=NORM_FONT, bg='light yellow',
                                textvariable=self.default_text3)
        receiver_phonenoText = Entry(dataEntryFrame, width=23, font=NORM_FONT, bg='light yellow',
                                     textvariable=self.default_text7, state=DISABLED)
        authorizerText = Entry(dataEntryFrame, width=23, font=NORM_FONT, bg='light yellow',
                               textvariable=self.default_text4)
        seva_amountText = Entry(dataEntryFrame, width=23, font=NORM_FONT, bg='light yellow',
                                textvariable=self.default_text2)

        cal = DateEntry(dataEntryFrame, width=21, font=NORM_FONT, date_pattern='dd/MM/yyyy', bg='light yellow',
                        anchor=W, justify=LEFT)

        paymentMode_text = StringVar(dataEntryFrame)
        paymentMode_list = ['Cash', 'Bank Transfer', 'Cheque', 'Demand Draft', 'Other']
        paymentMode_text.set("Other")
        paymentMode_menu = OptionMenu(dataEntryFrame, paymentMode_text, *paymentMode_list)
        paymentMode_menu.configure(width=19, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)

        trust_nametext = StringVar(dataEntryFrame)
        trustName_list = ['Vihangam Yoga (Karnataka) Trust', 'Aadarsh gaushala Trust']
        trust_nametext.set("Vihangam Yoga (Karnataka) Trust")
        trust_menu = OptionMenu(dataEntryFrame, trust_nametext, *trustName_list)
        trust_menu.configure(width=19, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)

        invoice_idText = Label(dataEntryFrame, text="------------", width=19, anchor=W, justify=LEFT,
                               font=NORM_FONT, bg='snow')

        receiver_IDText.grid(row=1, column=1, pady=3)
        receiver_nameText.grid(row=1, column=3, pady=3)
        descriptionText.grid(row=2, column=1, pady=3)
        receiver_phonenoText.grid(row=2, column=3, pady=3)
        seva_amountText.grid(row=3, column=3, pady=3)

        cal.grid(row=4, column=3, pady=3)
        authorizerText.grid(row=3, column=1, pady=3)
        paymentMode_menu.grid(row=5, column=3, pady=3)
        trust_menu.grid(row=4, column=1, pady=3)
        invoice_idText.grid(row=5, column=1, pady=3)

        var = IntVar()
        viewExpanse_Result = partial(self.enableIntenalID_RadioSelection, var, receiver_IDText,
                                     receiver_nameText, receiver_phonenoText)

        internalExpanse_radioBtn = Radiobutton(dataEntryFrame, text="Expanse By Id", variable=var, value=1,
                                               command=viewExpanse_Result, width=13, bg='snow',
                                               font=NORM_FONT, anchor=W, justify=LEFT)
        internalExpanse_radioBtn.select()
        internalExpanse_radioBtn.grid(row=0, column=0, pady=3)

        externalExpanse_radioBtn = Radiobutton(dataEntryFrame, text="Expanse By Name", variable=var, value=2,
                                               command=viewExpanse_Result, width=13, bg='snow',
                                               font=NORM_FONT, anchor=W, justify=LEFT)
        externalExpanse_radioBtn.grid(row=0, column=2, padx=10, pady=3)
        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(self.new_expanse_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=1, pady=8)
        print_invoice = Button(buttonFrame, text="Print Voucher", fg="Black", command=None,
                               font=NORM_FONT, width=10, bg='light grey', state=DISABLED)
        self.submit_deposit = Button(buttonFrame)
        insert_result = partial(self.create_Expanse_Excel,
                                self.new_expanse_window,
                                receiver_IDText,
                                seva_amountText,
                                descriptionText,
                                authorizerText,
                                cal,
                                paymentMode_menu,
                                paymentMode_text,
                                invoice_idText,
                                infolabel, print_invoice, trust_nametext, var, receiver_nameText, receiver_phonenoText)

        # create a Save Button and place into the self.new_expanse_window window
        self.submit_deposit.configure(text="Save", fg="Black", command=insert_result,
                                      font=NORM_FONT, width=8, bg='light grey', underline=0, state=DISABLED)
        self.submit_deposit.grid(row=0, column=0)
        print_invoice.grid(row=0, column=1)
        self.default_text1.trace("w", self.check_DepositBtn_state)
        self.default_text2.trace("w", self.check_DepositBtn_state)
        self.default_text3.trace("w", self.check_DepositBtn_state)
        self.default_text4.trace("w", self.check_DepositBtn_state)
        self.default_text5.trace("w", self.check_DepositBtn_state)

        # create a Clear Button and place into the self.new_expanse_window window
        # clear the commercial form after data has been saved
        clear_result = partial(self.clearNormalExpanseForm, receiver_nameText,
                               seva_amountText,
                               descriptionText,
                               authorizerText,
                               cal,
                               paymentMode_menu,
                               paymentMode_text,
                               invoice_idText, )
        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=2)

        # create a Cancel Button and place into the self.new_expanse_window window
        cancel_Result = partial(self.destroyWindow, self.new_expanse_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_Result,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=3)

        # ---------------------------------Button Frame End----------------------------------------

        infoFrame.grid(row=9, column=1, pady=5)
        infolabel.grid(row=0, column=0, padx=2, pady=3)

        self.new_expanse_window.bind('<Return>', lambda event=None: submit.invoke())
        self.new_expanse_window.bind('<Alt-d>', lambda event=None: submit.invoke())
        self.new_expanse_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.new_expanse_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.new_expanse_window.focus()
        self.new_expanse_window.grab_set()

        # Expanses cannot be created if total balance available  is 0
        '''
        balance = self.obj_commonUtil.readcurrent_balance()
        if int(balance) > 0:
            initial_text = "All fields are mandatory !!"
            fg_warning = "green"
        else:
            initial_text = "Insufficient Balance , please deposit  first !!!"
            fg_warning = "red"
            self.obj_commonUtil.disableChildren(dataEntryFrame)
        '''''

        # infolabel.configure(text="Please Enter Expanse Details", fg=fg_warning)

    def create_advance_expanse_entry(self):
        self.new_advance_window = Toplevel(self.master)
        self.new_advance_window.title("Issue Advance Amount ")
        self.new_advance_window.geometry('700x290+240+200')
        self.new_advance_window.configure(background='wheat')
        self.new_advance_window.resizable(width=False, height=False)
        self.new_advance_window.protocol('WM_DELETE_WINDOW', self.donothing)
        heading = Label(self.new_advance_window, text="Issue New Advance",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        dataEntryFrame = Frame(self.new_advance_window, width=200, height=100, bd=4, relief='ridge',
                               bg='snow')
        self.default_text1 = StringVar(dataEntryFrame, value='')
        self.default_text2 = StringVar(dataEntryFrame, value='')
        self.default_text3 = StringVar(dataEntryFrame, value='')
        self.default_text4 = StringVar(dataEntryFrame, value='')
        self.default_text5 = StringVar(dataEntryFrame, value='None')

        # lower frame added to show the result of transactions
        infoFrame = Frame(self.new_advance_window, width=100, height=20, bd=4, relief='ridge')
        # create a Book Name label
        receiver_name = Label(dataEntryFrame, text="Issued To(Id)", width=13, anchor=W, justify=LEFT,
                              font=NORM_FONT,
                              bg='snow')

        # create a Author label
        seva_amount = Label(dataEntryFrame, text="Amount(Rs.)", width=13, anchor=W, justify=LEFT,
                            font=NORM_FONT,
                            bg='snow')

        # create a Price label
        description = Label(dataEntryFrame, text="Description", width=13, anchor=W, justify=LEFT,
                            font=NORM_FONT, bg='snow')

        dateOfexpanse = Label(dataEntryFrame, text="Date", width=13, anchor=W,
                              justify=LEFT,
                              font=NORM_FONT, bg='snow')

        # create a Quantity label
        authorized_name = Label(dataEntryFrame, text="Authorized By", width=13, anchor=W, justify=LEFT,
                                font=NORM_FONT, bg='snow')

        # create a borrow fee label
        modeOfpayment = Label(dataEntryFrame, text="Payment Mode", width=13, anchor=W, justify=LEFT,
                              font=NORM_FONT, bg='snow')

        infolabel = Label(infoFrame, text="All fields are mandatory !!", width=50, anchor='center',
                          justify=LEFT,
                          font=NORM_VERDANA_FONT, bg='snow')

        heading.grid(row=0, column=0, columnspan=2)
        dataEntryFrame.grid(row=1, column=1, padx=10, pady=8)
        receiver_name.grid(row=0, column=0, pady=3)
        seva_amount.grid(row=0, column=2, padx=10, pady=3)
        description.grid(row=1, column=0, pady=3)
        dateOfexpanse.grid(row=1, column=2, pady=3)
        authorized_name.grid(row=2, column=0, pady=3)
        modeOfpayment.grid(row=2, column=2, pady=3)

        # create a text entry box
        # for typing the information
        receiver_nameText = Entry(dataEntryFrame, width=20, font=NORM_FONT, bg='light yellow',
                                  textvariable=self.default_text1)
        seva_amountText = Entry(dataEntryFrame, width=20, font=NORM_FONT, bg='light yellow',
                                textvariable=self.default_text2)
        descriptionText = Entry(dataEntryFrame, width=20, font=NORM_FONT, bg='light yellow',
                                textvariable=self.default_text3)

        cal = DateEntry(dataEntryFrame, width=18, font=NORM_FONT, date_pattern='dd/MM/yyyy', bg='light yellow',
                        anchor=W, justify=LEFT)

        authorizerText = Entry(dataEntryFrame, width=20, font=NORM_FONT, bg='light yellow',
                               textvariable=self.default_text4)

        paymentMode_text = StringVar(dataEntryFrame)
        paymentMode_list = ['Cash', 'Bank Transfer', 'Cheque', 'Demand Draft', 'Other']
        paymentMode_text.set("Other")
        paymentMode_menu = OptionMenu(dataEntryFrame, paymentMode_text, *paymentMode_list)
        paymentMode_menu.configure(width=16, font=NORM_FONT, bg='light yellow', anchor=W, justify=LEFT)

        receiver_nameText.grid(row=0, column=1, pady=3)
        seva_amountText.grid(row=0, column=3, pady=3)
        descriptionText.grid(row=1, column=1, pady=3)
        cal.grid(row=1, column=3, pady=3)
        authorizerText.grid(row=2, column=1, pady=3)
        paymentMode_menu.grid(row=2, column=3, pady=3)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(self.new_advance_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=1, pady=8)

        self.submit_deposit = Button(buttonFrame)
        insert_result = partial(self.create_Advance_Excel,
                                self.new_advance_window,
                                receiver_nameText,
                                seva_amountText,
                                descriptionText,
                                authorizerText,
                                cal,
                                paymentMode_text,
                                infolabel)

        # create a Save Button and place into the self.new_advance_window window
        self.submit_deposit.configure(text="Issue", fg="Black", command=insert_result,
                                      font=NORM_FONT, width=8, bg='light grey', underline=0, state=DISABLED)
        self.submit_deposit.grid(row=0, column=0)
        self.default_text1.trace("w", self.check_DepositBtn_state)
        self.default_text2.trace("w", self.check_DepositBtn_state)
        self.default_text3.trace("w", self.check_DepositBtn_state)
        self.default_text4.trace("w", self.check_DepositBtn_state)
        self.default_text5.trace("w", self.check_DepositBtn_state)

        # create a Clear Button and place into the self.new_advance_window window
        # clear the commercial form after data has been saved
        clear_result = partial(self.clearAdvanceIssueForm, receiver_nameText,
                               seva_amountText,
                               descriptionText,
                               authorizerText,
                               paymentMode_text,
                               infolabel)
        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=2)

        # create a Cancel Button and place into the self.new_advance_window window
        cancel_Result = partial(self.destroyWindow, self.new_advance_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_Result,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=3)
        # ---------------------------------Button Frame End----------------------------------------

        infoFrame.grid(row=9, column=1, pady=5)
        infolabel.grid(row=0, column=0, padx=2, pady=3)

        self.new_advance_window.bind('<Return>', lambda event=None: self.submit_deposit.invoke())
        self.new_advance_window.bind('<Alt-d>', lambda event=None: self.submit_deposit.invoke())
        self.new_advance_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.new_advance_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.new_advance_window.focus()
        self.new_advance_window.grab_set()

        # Advance cannot be issued if total balance available  is 0
        balance = self.obj_commonUtil.readcurrent_balance(VIHANGAM_YOGA_KARNATAKA_TRUST)
        if int(balance) > 0:
            initial_text = "All fields are mandatory !!"
            fg_warning = "green"
        else:
            initial_text = "Insufficient Balance, please deposit  first !!!"
            fg_warning = "red"
            self.obj_commonUtil.disableChildren(dataEntryFrame)

        infolabel.configure(text=initial_text, fg=fg_warning)

    def clearBookBorrowForm(self, bookBorrow_window, book_IdText, member_Id):
        book_IdText.delete(0, END)
        book_IdText.configure(fg='black')
        member_Id.delete(0, END)
        member_Id.configure(fg='black')
        member_Id.focus_set()

    def clearBookPurchaseForm(self, item_idText,
                              itemname_labelText,
                              authordetails_labelText,
                              unitprice_labelText,
                              quantitydetails_labelText,
                              cart_item_count,
                              amount_payableTextLabel,
                              member_IdText,
                              customer_nameText,
                              quantityText,
                              paymentMode_text,
                              itemid_labelText,
                              searchinfo_label):
        member_IdText.delete(0, END)
        member_IdText.configure(fg='black')
        customer_nameText.delete(0, END)
        customer_nameText.configure(fg='black')
        item_idText.delete(0, END)
        item_idText.configure(fg='black')
        quantityText.delete(0, END)
        quantityText.configure(fg='black')
        paymentMode_text.set("Other")
        itemname_labelText['text'] = ""
        authordetails_labelText['text'] = ""
        unitprice_labelText['text'] = ""
        quantitydetails_labelText['text'] = ""
        cart_item_count['text'] = "0"
        itemid_labelText['text'] = ""
        amount_payableTextLabel['text'] = "0"
        searchinfo_label.configure(text="Enter Item id for purchase", fg='green')
        self.list_InvoicePrint = []
        print("Total cart item: ", len(self.list_InvoicePrint))

    def clearNonCommercialEntryForm(self, item_name,
                                    descriptionText,
                                    categoryText,
                                    priceText,
                                    cal,
                                    quantityText,
                                    usage_locationTextt):
        categoryText.set("Other")
        item_name.delete(0, END)
        item_name.configure(fg='black')
        item_name.focus_set()
        quantityText.delete(0, END)
        quantityText.configure(fg='black')
        usage_locationTextt.delete(0, END)
        usage_locationTextt.configure(fg='black')
        priceText.delete(0, END)
        priceText.configure(fg='black')
        descriptionText.delete(0, END)
        descriptionText.configure(fg='black')

    def clearNormalExpanseForm(self, receiver_nameText,
                               seva_amountText,
                               descriptionText,
                               authorizerText,
                               cal,
                               paymentMode_menu,
                               paymentMode_text,
                               invoice_idText, ):
        paymentMode_text.set("Other")
        receiver_nameText.delete(0, END)
        receiver_nameText.configure(fg='black')
        receiver_nameText.focus_set()
        seva_amountText.delete(0, END)
        seva_amountText.configure(fg='black')
        authorizerText.delete(0, END)
        authorizerText.configure(fg='black')
        invoice_idText[text] = "------------"
        descriptionText.delete(0, END)
        descriptionText.configure(fg='black')

    def clearNonCommercialEntryForm(self, item_name,
                                    descriptionText,
                                    categoryText,
                                    priceText,
                                    cal,
                                    quantityText,
                                    usage_locationTextt):
        categoryText.set("Other")
        item_name.delete(0, END)
        item_name.configure(fg='black')
        item_name.focus_set()
        quantityText.delete(0, END)
        quantityText.configure(fg='black')
        usage_locationTextt.delete(0, END)
        usage_locationTextt.configure(fg='black')
        priceText.delete(0, END)
        priceText.configure(fg='black')
        descriptionText.delete(0, END)
        descriptionText.configure(fg='black')

    def clearAdvanceIssueForm(self, receiver_nameText,
                              seva_amountText,
                              descriptionText,
                              authorizerText,
                              paymentMode_text,
                              infolabel):

        receiver_nameText.delete(0, END)
        receiver_nameText.configure(fg='black')
        receiver_nameText.focus_set()
        authorizerText.delete(0, END)
        authorizerText.configure(fg='black')
        seva_amountText.delete(0, END)
        seva_amountText.configure(fg='black')
        descriptionText.delete(0, END)
        descriptionText.configure(fg='black')
        paymentMode_text.set("Cash")
        infolabel.configure(text="All fields are mandatory!!", fg='green')

    def clearMagazineSubscriptionForm(self, member_idText,
                                      seva_amountText,
                                      magazineCategoryText,
                                      quantityText,
                                      cal,
                                      paymentMode_text,
                                      authorizedby_idText,
                                      payableAmtText,
                                      infolabel, print_invoice):

        member_idText.delete(0, END)
        member_idText.configure(fg='black')
        member_idText.focus_set()
        authorizedby_idText.delete(0, END)
        authorizedby_idText.configure(fg='black')
        quantityText.delete(0, END)
        quantityText.configure(fg='black')
        seva_amountText.delete(0, END)
        seva_amountText.configure(fg='black')
        authorizedby_idText.delete(0, END)
        authorizedby_idText.configure(fg='black')
        magazineCategoryText.set("Other")
        payableAmtText['text'] = "----------"
        infolabel.configure(text="All fields are mandatory!!", fg='green')
        print_invoice.configure(state=DISABLED, bg='light grey')

    def clearMonetarySevaDepositForm(self, donator_idText,
                                     seva_amountText,
                                     categoryText,
                                     collector_nameText,
                                     cal,
                                     paymentMode_menu,
                                     paymentMode_text,
                                     authorizedby_Text,
                                     akshayPatra_Text,
                                     invoice_idText, print_invoice):

        donator_idText.delete(0, END)
        donator_idText.configure(fg='black')
        donator_idText.focus_set()
        seva_amountText.delete(0, END)
        seva_amountText.configure(fg='black')
        collector_nameText.delete(0, END)
        collector_nameText.configure(fg='black')
        authorizedby_Text.delete(0, END)
        authorizedby_Text.configure(fg='black')
        akshayPatra_Text.delete(0, END)
        akshayPatra_Text.configure(fg='black')
        invoice_idText['text'] = "----------"
        paymentMode_text.set("Other")
        categoryText.set("Monthly Seva")
        print_invoice.configure(state=DISABLED, bg='light grey')

    def clearSevaRashiDepositForm(self, donator_nameText,
                                  seva_amountText,
                                  categoryText,
                                  collector_nameText,
                                  dateOfCollection_calc,
                                  paymentMode_menu,
                                  paymentMode_text):
        categoryText.set("Other")
        donator_nameText.delete(0, END)
        donator_nameText.configure(fg='black')
        donator_nameText.focus_set()
        seva_amountText.delete(0, END)
        seva_amountText.configure(fg='black')
        collector_nameText.delete(0, END)
        collector_nameText.configure(fg='black')
        categoryText.set("Other")

    def borrow_book(self, newItem_window, bookName, member_Id):
        libMemberId = member_Id.get()
        totalBorrowedQuantity = self.calculate_TotalBorrowRecord(libMemberId)
        if totalBorrowedQuantity == 2:
            messagebox.showwarning("Limit Reached", "Only 2 items allowed per member")
            return
        # validate the member , if  already registered
        print("borrow_book->libMemberId :", libMemberId)
        bValidMember = self.validate_memberlibraryID(libMemberId, 1)
        print("Member bValidMember", bValidMember)
        # issue books only to the registered member
        if bValidMember is True:
            book_toBorrow = bookName.get()
            today = date.today()
            dateOfBorrow = today.strftime("%Y-%m-%d")

            # validate the book name , and issue only if book exists
            bBookExists = self.validate_bookName(book_toBorrow)
            if bBookExists:
                file_borrow = open("Stock.txt", 'r')
                line_no = 0
                for line in file_borrow:
                    record_field = line.split(',')
                    if record_field[0] == book_toBorrow:
                        new_quantity = int(record_field[4]) - 1

                        # check if  requested book is available in stock, issue only if stock is available
                        if new_quantity > 0:
                            infile = open("Stock.txt", 'r')
                            content = infile.readlines()
                            string = ""
                            content[line_no] = record_field[0] + "," + record_field[1] + "," + record_field[2] + "," + \
                                               record_field[3] + "," + str(new_quantity) + "," + record_field[5]
                            outfile = open("Stock.txt", 'w')
                            outfile.write(string.join(content))
                            infile.close()
                            outfile.close()
                            # writing record in member database
                            file_name = libMemberId + ".txt"
                            infile_member = open(file_name, 'a')
                            member_name = self.findMemberName(libMemberId)
                            borrow_entry = libMemberId + "," + member_name + "," + dateOfBorrow + "," + book_toBorrow + "\n"
                            infile_member.write(borrow_entry)
                            infile_member.close()
                            self.clearBookBorrowForm(newItem_window, bookName, member_Id)
                            user_choice = messagebox.askquestion("Successfully Borrowed",
                                                                 "Do you want to borrow another item ? ")
                            # destroy the data entry form , if user do not want to add more records
                            if user_choice == 'no':
                                newItem_window.destroy()
                        else:
                            messagebox.showwarning("Warning ", "No stock available")
                        break
                    line_no = line_no + 1

            else:
                messagebox.showwarning("Warning ", "No such book exists !!")
        else:
            messagebox.showwarning("Warning ", "Invalid Member id !!")

    def borrow_book_Excel(self, newItem_window, book_IdText, member_Id, cal, infoLabel):
        libMemberId = member_Id.get()
        totalBorrowedQuantity = self.calculate_TotalBorrowRecord_Excel(libMemberId)
        if totalBorrowedQuantity == 10:
            messagebox.showwarning("Limit Reached", "Only 10 items allowed per member")
            return
        # validate the member , if  already registered
        print("borrow_book->libMemberId :", libMemberId)
        bValidMember = self.validate_memberlibraryID_Excel(libMemberId, 1)
        print("Member bValidMember", bValidMember)
        # issue books only to the registered member
        if bValidMember is True:
            bookid_toBorrow = book_IdText.get()
            dateTimeObj = cal.get_date()
            dateOfBorrow = dateTimeObj.strftime("%Y-%m-%d")

            # validate the book name , and issue only if book exists
            bBookExists, book_toBorrow = self.validate_bookbyId_Excel(bookid_toBorrow)
            if bBookExists:
                # To open the workbook
                # workbook object is created
                wb_obj = openpyxl.load_workbook(PATH_STOCK)

                # Get workbook active sheet object
                # from the active attribute
                sheet_obj = wb_obj.active
                totalrecords = self.totalrecords_excelDataBase(PATH_STOCK)
                for iLoop in range(2, totalrecords + 1):
                    cell_obj = sheet_obj.cell(row=iLoop, column=2)
                    if cell_obj.value == bookid_toBorrow:
                        bookId = sheet_obj.cell(row=iLoop, column=2).value
                        # check if  requested book is available in stock, issue only if stock is available
                        if int(sheet_obj.cell(row=iLoop, column=7).value) > 0:  # if quantity > 0
                            new_quantity = int(sheet_obj.cell(row=iLoop, column=7).value) - 1
                            sheet_obj.cell(row=iLoop, column=7).value = str(new_quantity)
                            wb_obj.save(PATH_STOCK)
                            # writing record in member database
                            file_name = "..\\Member_Data\\" + libMemberId + ".xlsx"
                            self.obj_initDatabase.initilize_member_borrow_database(file_name)

                            member_file = openpyxl.load_workbook(file_name)
                            total_borrow_records = self.totalrecords_excelDataBase(file_name)
                            # Get workbook active sheet object
                            # from the active attribute
                            sheet_obj = member_file.active
                            member_name = self.findMemberName_Excel(libMemberId, 1)
                            serial_no = 0
                            row_no = 0
                            total_records = self.totalrecords_excelDataBase(file_name)
                            if total_records is 0:
                                serial_no = 1
                                row_no = 2
                            else:
                                serial_no = total_records + 1
                                row_no = total_records + 2

                            # sets the general formatting for the new entry in new row
                            for index in range(1, 8):
                                sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                                     bold=False)
                                sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                               vertical='center')
                            # write the newly borrow record into member specific file
                            sheet_obj.cell(row=row_no, column=1).value = serial_no
                            sheet_obj.cell(row=row_no, column=2).value = libMemberId
                            sheet_obj.cell(row=row_no, column=3).value = member_name
                            sheet_obj.cell(row=row_no, column=4).value = bookId
                            sheet_obj.cell(row=row_no, column=5).value = book_toBorrow
                            sheet_obj.cell(row=row_no, column=6).value = dateOfBorrow
                            sheet_obj.cell(row=row_no, column=7).value = ""  # date of return is kept as null
                            sheet_obj.cell(row=total_borrow_records + 1,
                                           column=8).value = ""  # paid fee is NULL in case of borrow
                            # write finished for member new borrow record

                            member_file.save(file_name)
                            self.clearBookBorrowForm(newItem_window, book_IdText, member_Id)
                            infoLabel.configure(text="Successfully Borrowed!!", fg='green')
                        else:
                            infoLabel.configure(text="No stock available", fg='red')
                        break
            else:
                infoLabel.configure(text="Invalid Item Id !!!", fg='red')
        else:
            infoLabel.configure(text="Invalid Member Id !!", fg='red')

    def purchase_stock_item(self, member_IdText, customer_nameText,
                            purchasebtn, print_btn, addtocart_btn, quantityText,
                            amount_payableTextLabel,
                            paymentMode_text, searchinfo_label, var, cal, local_centerText):
        if var.get() == 1:
            libMemberId = member_IdText.get()
        else:
            libMemberId = "Not Available"

        print("purchase_stock_item->MemberId :", libMemberId)
        dateTimeObj = cal.get_date()
        dateOfPurchase = dateTimeObj.strftime("%Y-%m-%d ")
        # validate the member , if  already registered
        print(" Cart items :", self.list_InvoicePrint)
        invoice_id = self.generate_StockPurchase_invoiceID()
        # reading the cart and filling the data to database/s
        for cartLoop in range(0, len(self.list_InvoicePrint)):

            print("Date is OK")
            book_toPurchase = self.list_InvoicePrint[cartLoop][3]
            # To open the workbook
            # workbook object is created
            subdir_commercialstock = "..\\Library_Stock\\" + local_centerText.get() + "\\Commercial_Stock"
            filename = subdir_commercialstock + "\\Commercial_Stock.xlsx"
            wb_obj = openpyxl.load_workbook(filename)

            # Get workbook active sheet object
            # from the active attribute
            sheet_obj = wb_obj.active
            totalrecords = self.totalrecords_excelDataBase(filename)
            print("Total records :", totalrecords)
            dict_index = 1
            for iLoop in range(2, totalrecords + 2):
                book_name = sheet_obj.cell(row=iLoop, column=3).value
                bookId = sheet_obj.cell(row=iLoop, column=2).value  # item id in database
                if bookId == book_toPurchase:
                    # check if  requested book is available in stock, issue only if stock is available
                    new_quantity = int(sheet_obj.cell(row=iLoop, column=7).value) - int(
                        self.list_InvoicePrint[cartLoop][1])
                    sheet_obj.cell(row=iLoop, column=7).value = str(new_quantity)
                    unit_mrp = sheet_obj.cell(row=iLoop, column=5).value
                    if var.get() == 1:
                        contact_no = self.findMemberContactNo_Excel(libMemberId, 1)
                    else:
                        contact_no = "Contact # Not Available"
                    print("STOCK is modified")
                    wb_obj.save(filename)
                    # write the purchase record in stock purchase file
                    # writing record in member database
                    file_name = InitDatabase.getInstance().get_purchase_record_database_name()

                    member_file = openpyxl.load_workbook(file_name)
                    total_borrow_records = self.totalrecords_excelDataBase(file_name)
                    # Get workbook active sheet object
                    # from the active attribute
                    sheet_obj = member_file.active
                    if var.get() == 1:
                        member_name = self.findMemberName_Excel(libMemberId, 1)
                    else:
                        member_name = customer_nameText.get()

                    print("contact_no :", contact_no, "member_name :", member_name)
                    total_records = self.totalrecords_excelDataBase(file_name)
                    if total_records is 0:
                        serial_no = 1
                        row_no = 2
                    else:
                        serial_no = total_records + 1
                        row_no = total_records + 2

                    # sets the general formatting for the new entry in new row

                    total_mrp = str(int(unit_mrp) * int(quantityText.get()))
                    if dict_index == 1:
                        balance = total_mrp
                    else:
                        balance = balance + int(sheet_obj.cell(row=row_no, column=7).value)
                    for index in range(1, 11):
                        sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                             bold=False)
                        sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                       vertical='center')

                    # write the newly borrow record into member specific file
                    total_mrp = str(int(unit_mrp) * int(quantityText.get()))
                    sheet_obj.cell(row=row_no, column=1).value = serial_no
                    sheet_obj.cell(row=row_no, column=2).value = libMemberId
                    sheet_obj.cell(row=row_no, column=3).value = member_name
                    sheet_obj.cell(row=row_no, column=4).value = bookId
                    sheet_obj.cell(row=row_no, column=5).value = book_toPurchase
                    sheet_obj.cell(row=row_no, column=6).value = dateOfPurchase
                    sheet_obj.cell(row=row_no, column=7).value = total_mrp
                    sheet_obj.cell(row=row_no, column=8).value = str(balance)
                    sheet_obj.cell(row=row_no, column=9).value = invoice_id
                    sheet_obj.cell(row=row_no, column=10).value = self.list_InvoicePrint[cartLoop][1]
                    print("Total MRP :", total_mrp)
                    print("PURCHASE STOCK is modified")
                    amount_payableTextLabel['text'] = str(total_mrp)
                    member_file.save(file_name)

                    '''
                    # ------------------------------------Transaction sheet update- start----------------------------------
                    # open transaction sheet and enter the data
                    # receiving donation is a credit transaction for the organization
                    # ENABLE IF REQUIRED IN FUTURE 
                    file_name_transaction = PATH_TRANSACTION_SHEET
                    transaction_wb_obj = openpyxl.load_workbook(file_name_transaction)
                    transaction_sheet_obj = transaction_wb_obj.active
                    total_records_transaction = self.totalrecords_excelDataBase(file_name_transaction)

                    if total_records_transaction is 0:
                        serial_no = 1
                        row_no = 2
                        balance_amount = 0
                    else:
                        serial_no = total_records_transaction + 1
                        row_no = total_records_transaction + 2
                        balance_amount = int(str(transaction_sheet_obj.cell(row=row_no - 1, column=9).value))

                    # sets the general formatting for the new entry in new row
                    for index in range(1, 10):
                        transaction_sheet_obj.cell(row=row_no, column=index).font = Font(size=12,
                                                                                         name='Times New Roman',
                                                                                         bold=False)
                        transaction_sheet_obj.cell(row=row_no, column=index).alignment = Alignment(
                            horizontal='center',
                            vertical='center', wrapText=True)

                    # new book data is assigned to respective cells in row
                    transaction_sheet_obj.cell(row=row_no, column=1).value = serial_no
                    transaction_sheet_obj.cell(row=row_no, column=2).value = dateOfPurchase
                    transaction_sheet_obj.cell(row=row_no, column=3).value = total_mrp
                    transaction_sheet_obj.cell(row=row_no, column=4).value = "---"
                    transaction_sheet_obj.cell(row=row_no, column=5).value = "Stock Purchase_" + book_name
                    transaction_sheet_obj.cell(row=row_no, column=6).value = paymentMode_text.get()
                    transaction_sheet_obj.cell(row=row_no,
                                               column=7).value = self.logged_staff_id  # authorizor id
                    transaction_sheet_obj.cell(row=row_no, column=8).value = self.logged_staff_name
                    transaction_sheet_obj.cell(row=row_no, column=9).value = str(
                        balance_amount + int(total_mrp))
                    transaction_wb_obj.save(PATH_TRANSACTION_SHEET)
                    print("MASTER TRANSACTION is modified")
                    '''
                    addtocart_btn.configure(state=DISABLED, bg='light grey')
                    purchasebtn.configure(state=DISABLED, bg='light grey')
                    break
            dict_index = dict_index + 1
            # ------------------------------------Transaction sheet update- End----------------------------------
        self.generateInvoicePage(member_name,
                                 libMemberId,
                                 dateOfPurchase,
                                 contact_no,
                                 book_toPurchase,
                                 unit_mrp, quantityText,
                                 print_btn, purchasebtn,
                                 addtocart_btn, var, searchinfo_label, invoice_id)

    def addto_cart_item(self, newItem_window, item_idText,
                        member_IdText,
                        purchase_btn,
                        print_btn,
                        addtocart_btn,
                        quantityText,
                        cart_item_count,
                        amount_payableTextLabel,
                        paymentMode_text,
                        searchinfo_label,
                        var,
                        cal, local_centerText):
        print("var = ", var.get())
        if var.get() == 1:
            libMemberId = member_IdText.get()
        else:
            libMemberId = "Not Available"

        dateTimeObj = cal.get_date()
        dateOfPurchase = dateTimeObj.strftime("%d-%m-%Y ")
        # validate the member , if  already registered

        print("purchase_stock_item->MemberId :", libMemberId)
        itemId = "CI-" + item_idText.get()
        if var.get() == 1:
            bValidMember = self.validate_memberlibraryID_Excel(libMemberId, 1)
        else:
            bValidMember = True
        print("var.get() : ", var.get(), bValidMember, "Member bValidMember :", bValidMember)
        # issue books only to the registered member
        if bValidMember is True:
            today = date.today()
            if dateTimeObj <= today and (quantityText.get()).isnumeric():
                print("Date is OK")
                book_toPurchase = itemId
                # To open the workbook
                # workbook object is created
                subdir_commercialstock = "..\\Library_Stock\\" + local_centerText.get() + "\\Commercial_Stock"
                filename = subdir_commercialstock + "\\Commercial_Stock.xlsx"
                wb_obj = openpyxl.load_workbook(filename)

                # Get workbook active sheet object
                # from the active attribute
                sheet_obj = wb_obj.active
                totalrecords = self.totalrecords_excelDataBase(filename)
                for iLoop in range(2, totalrecords + 2):
                    book_name = sheet_obj.cell(row=iLoop, column=3).value
                    bookId = sheet_obj.cell(row=iLoop, column=2).value  # item id in database
                    unit_mrp = sheet_obj.cell(row=iLoop, column=5).value
                    if bookId == book_toPurchase:
                        # check if  requested book is available in stock, issue only if stock is available
                        if int(sheet_obj.cell(row=iLoop, column=7).value) >= int(
                                quantityText.get()):  # if quantity >= requested quantity

                            # prepare the cart locally and retain it as long as purchase button is pressed
                            arr_InvoiceRecords = [book_name, quantityText.get(), unit_mrp, bookId]
                            self.list_InvoicePrint.append(arr_InvoiceRecords)

                            cart_item_count['text'] = str(len(self.list_InvoicePrint))

                            # calculate the total mrp
                            total_cart_mrp = 0
                            for iLoop in range(0, len(self.list_InvoicePrint)):
                                total_cart_mrp = total_cart_mrp + (int(self.list_InvoicePrint[iLoop][1]) * int(
                                    self.list_InvoicePrint[iLoop][2]))

                            amount_payableTextLabel['text'] = str(total_cart_mrp)
                            purchase_btn.configure(state=NORMAL, bg='light cyan')
                        else:
                            searchinfo_label.configure(text="No stock available!!!", fg='red')
                        break
                else:
                    searchinfo_label.configure(text="Item does not exists!!!", fg='red')
            else:
                searchinfo_label.configure(text="Invalid Qunatity/Furture Date Chosen", fg='red')
        else:
            searchinfo_label.configure(text="Invalid Member Id !!!", fg='red')

    def generateInvoicePage(self, member_name,
                            libMemberId,
                            dateOfPurchase,
                            contact_no,
                            book_toPurchase,
                            unit_mrp, quantityText,
                            print_btn, purchasebtn,
                            addtocart_btn, var, searchinfo_label, invoice_id):

        file_name = "..\\Library_Stock\\Invoices\\Template\\sales-invoice.xlsx"
        searchinfo_label.configure(text="Invoice is being generated. Please wait ...", fg="purple")
        wb_obj = openpyxl.load_workbook(file_name)
        sheet_obj = wb_obj.active
        print(" generateInvoicePage =>Var :", var.get())
        if var.get() == 1:
            member_data = self.retrieve_MemberRecords_Excel(libMemberId, 1, SEARCH_BY_MEMBERID)
            sheet_obj.cell(row=10, column=1).value = member_data[7]
            sheet_obj.cell(row=11, column=1).value = member_data[8] + "," + member_data[9]
            sheet_obj.cell(row=12, column=1).value = member_data[12] + "," + member_data[10]
        else:
            sheet_obj.cell(row=10, column=1).value = "Address : NA"
            sheet_obj.cell(row=11, column=1).value = "City : NA,Country :NA"
            sheet_obj.cell(row=12, column=1).value = "Pin-code : NA"

        sheet_obj.cell(row=2, column=6).value = dateOfPurchase
        sheet_obj.cell(row=3, column=6).value = invoice_id
        sheet_obj.cell(row=9, column=1).value = member_name

        sheet_obj.cell(row=13, column=1).value = contact_no

        sheet_obj.cell(row=16, column=1).value = self.logged_staff_name
        sheet_obj.cell(row=16, column=2).value = member_name
        sheet_obj.cell(row=16, column=3).value = libMemberId
        sheet_obj.cell(row=16, column=4).value = contact_no

        final_paymentValue = 0
        # clear the existing sales template
        for iLoop_row in range(0, 10):
            for iLoop_column in range(1, 7):
                sheet_obj.cell(row=19 + iLoop_row, column=iLoop_column).value = ""

        # filling the purchase details in invoice
        for iLoop in range(0, len(self.list_InvoicePrint)):
            tax = int(self.list_InvoicePrint[iLoop][2]) * (TAX_ON_MRP / 100)
            sheet_obj.cell(row=19 + iLoop, column=1).value = str(iLoop + 1)
            sheet_obj.cell(row=19 + iLoop, column=2).value = str(self.list_InvoicePrint[iLoop][0])  # Name
            sheet_obj.cell(row=19 + iLoop, column=3).value = str(self.list_InvoicePrint[iLoop][1])  # quantity
            sheet_obj.cell(row=19 + iLoop, column=4).value = str(self.list_InvoicePrint[iLoop][2])  # price of each item
            sheet_obj.cell(row=19 + iLoop, column=5).value = str(tax)
            sheet_obj.cell(row=19 + iLoop, column=6).value = str(
                ((int(self.list_InvoicePrint[iLoop][1])) * int(self.list_InvoicePrint[iLoop][2])) + int(tax))
            final_paymentValue = final_paymentValue + int(sheet_obj.cell(row=19 + iLoop, column=6).value)

        sheet_obj.cell(row=29, column=6).value = str(final_paymentValue)

        # print("Invoice records  :")
        # for iLoop in range(0, len(self.list_InvoicePrint)):
        # print(" Record :", iLoop + 1, " :", self.list_InvoicePrint[iLoop][1])

        wb_obj.save(file_name)
        dest_file = self.obj_initDatabase.get_invoice_directory_name() + "\\" + invoice_id + ".pdf "
        dest_desktop_file = self.obj_initDatabase.get_desktop_invoices_directory_path() + "\\" + invoice_id + ".pdf "
        self.obj_commonUtil.convertExcelToPdf(file_name, dest_file)

        invoice_info = "Invoice is ready ! Invoice Id : " + invoice_id
        searchinfo_label.configure(text=invoice_info, fg='purple')
        copyfile(dest_file, dest_desktop_file)

        print_btn.configure(state=NORMAL, bg='light cyan')
        print_result = partial(self.printInvoice, dest_file)
        print_btn.configure(command=print_result)

        # disable to add to cart and purchase button, so that same invoice is not generated twice
        purchasebtn.configure(state=DISABLED, bg='light grey')
        addtocart_btn.configure(state=DISABLED, bg='light grey')

        # update the invoice table
        self.obj_commonUtil.updateInvoiceTable(invoice_id, dest_file)

        # self.obj_commonUtil.clearSales_InvoiceData(file_name,len(self.list_InvoicePrint))

    def printInvoice(self, invoice_file):
        print("Invoice file name :", invoice_file)
        os.startfile(invoice_file, 'print')

    def generateSevaReceipt(self, donator_idText,
                            seva_amountText,
                            categoryText,
                            collector_nameText,
                            dateOfCollection_calc,
                            paymentMode_text,
                            invoice_id, member_data):

        currentYearDirectory = self.obj_commonUtil.getCurrentYearFolderName()
        file_name = "..\\Expanse_Data\\" + currentYearDirectory + "\\Seva_Rashi\\Receipts\\Template\\Donation_Receipt_Template.xlsx"
        wb_obj = openpyxl.load_workbook(file_name)
        sheet_obj = wb_obj.active

        sheet_obj.cell(row=6, column=4).value = str(invoice_id)
        sheet_obj.cell(row=7, column=4).value = str(dateOfCollection_calc)
        sheet_obj.cell(row=8, column=4).value = str(member_data[2])
        sheet_obj.cell(row=9, column=4).value = str(member_data[7])  # Address
        sheet_obj.cell(row=10, column=4).value = str(member_data[8])  # city
        sheet_obj.cell(row=11, column=4).value = str(member_data[9])  # state
        sheet_obj.cell(row=12, column=4).value = str(member_data[10])
        sheet_obj.cell(row=13, column=4).value = str(member_data[11])
        sheet_obj.cell(row=15, column=4).value = str(categoryText.get())

        sheet_obj.cell(row=16, column=4).value = str(seva_amountText.get())
        sheet_obj.cell(row=17, column=4).value = str(num2words.num2words(int(seva_amountText.get()))) + " Rs. only"
        sheet_obj.cell(row=18, column=4).value = str(paymentMode_text.get())
        sheet_obj.cell(row=19, column=4).value = str(collector_nameText)

        wb_obj.save(file_name)
        pdf_file = "..\\Expanse_Data\\" + currentYearDirectory + "\\Seva_Rashi\\Receipts\\Receipts\\" + invoice_id + ".pdf"

        self.obj_commonUtil.convertExcelToPdf(file_name, pdf_file)

        destdir_repo = self.obj_initDatabase.get_invoice_directory_name() + "\\" + invoice_id + ".pdf"
        self.obj_commonUtil.updateInvoiceTable(invoice_id, destdir_repo)
        desktop_repo = self.obj_initDatabase.get_desktop_invoices_directory_path() + "\\" + invoice_id + ".pdf"
        self.obj_commonUtil.updateMonetaryDonationReceiptBooklet(invoice_id)
        copyfile(pdf_file, destdir_repo)
        copyfile(pdf_file, desktop_repo)
        os.startfile(pdf_file, 'print')

    def generateDonationReceipt(self, donator_idText,
                                seva_amountText,
                                categoryText,
                                collector_nameText,
                                dateOfCollection_calc,
                                paymentMode_text,
                                invoice_id, member_data, print_invoice, submit_deposit):

        currentYearDirectory = self.obj_commonUtil.getCurrentYearFolderName()
        file_name = "..\\Expanse_Data\\" + currentYearDirectory + "\\Seva_Rashi\\Receipts\\Template\\Donation_Receipt_Template.xlsx"
        wb_obj = openpyxl.load_workbook(file_name)
        sheet_obj = wb_obj.active

        sheet_obj.cell(row=6, column=4).value = str(invoice_id)
        sheet_obj.cell(row=7, column=4).value = str(dateOfCollection_calc)
        sheet_obj.cell(row=8, column=4).value = str(member_data[2])
        sheet_obj.cell(row=9, column=4).value = str(member_data[7])  # Address
        sheet_obj.cell(row=10, column=4).value = str(member_data[8])  # city
        sheet_obj.cell(row=11, column=4).value = str(member_data[9])  # state
        sheet_obj.cell(row=12, column=4).value = str(member_data[10])
        sheet_obj.cell(row=13, column=4).value = str(member_data[11])
        sheet_obj.cell(row=15, column=4).value = str(categoryText.get())

        sheet_obj.cell(row=16, column=4).value = str(seva_amountText.get())
        sheet_obj.cell(row=17, column=4).value = str(num2words.num2words(int(seva_amountText.get()))) + " Rs. only"
        sheet_obj.cell(row=18, column=4).value = str(paymentMode_text.get())
        sheet_obj.cell(row=19, column=4).value = str(collector_nameText)

        wb_obj.save(file_name)
        pdf_file = "..\\Expanse_Data\\" + currentYearDirectory + "\\Seva_Rashi\\Receipts\\Receipts\\" + invoice_id + ".pdf"

        self.obj_commonUtil.convertExcelToPdf(file_name, pdf_file)
        print_result = partial(self.printInvoice, pdf_file)
        print_invoice.configure(state=NORMAL, bg='light cyan', command=print_result)
        submit_deposit.configure(state=DISABLED, bg='light grey')

        destdir_repo = self.obj_initDatabase.get_invoice_directory_name() + "\\" + invoice_id + ".pdf"
        self.obj_commonUtil.updateInvoiceTable(invoice_id, destdir_repo)
        desktop_repo = self.obj_initDatabase.get_desktop_invoices_directory_path() + "\\" + invoice_id + ".pdf"
        if categoryText.get() == "Gaushala Seva":
            trustType = SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST
        else:
            trustType = VIHANGAM_YOGA_KARNATAKA_TRUST

        self.obj_commonUtil.updateMonetaryDonationReceiptBooklet(invoice_id, trustType)
        copyfile(pdf_file, destdir_repo)
        copyfile(pdf_file, desktop_repo)

        # os.startfile(pdf_file, 'print')

    def generatePatrikaSubscription_Receipt(self, member_data,
                                            seva_amountText,
                                            quantityText,
                                            authorizor_data,
                                            dateOfCollection_calc,
                                            magazineCategoryText,
                                            paymentMode_text,
                                            invoice_id):

        currentYearDirectory = self.obj_commonUtil.getCurrentYearFolderName()
        file_name = "..\\Expanse_Data\\" + currentYearDirectory + "\\Magazine_Subscription\\Receipt\\Template\\Subscription_Receipt_Template.xlsx"
        wb_obj = openpyxl.load_workbook(file_name)
        sheet_obj = wb_obj.active
        sheet_obj.cell(row=4, column=4).font = Font(size=10, name='Arial',
                                                    bold=False)
        sheet_obj.cell(row=4, column=4).value = invoice_id
        sheet_obj.cell(row=4, column=6).font = Font(size=10, name='Arial',
                                                    bold=False)
        sheet_obj.cell(row=4, column=6).value = dateOfCollection_calc
        sheet_obj.cell(row=6, column=4).font = Font(size=10, name='Arial',
                                                    bold=False)
        sheet_obj.cell(row=6, column=4).value = member_data[0]
        sheet_obj.cell(row=6, column=6).font = Font(size=10, name='Arial',
                                                    bold=False)
        sheet_obj.cell(row=6, column=6).value = authorizor_data[0]  # Address
        sheet_obj.cell(row=7, column=4).font = Font(size=10, name='Arial',
                                                    bold=False)
        sheet_obj.cell(row=7, column=4).value = member_data[2]
        sheet_obj.cell(row=7, column=6).font = Font(size=10, name='Arial',
                                                    bold=False)
        sheet_obj.cell(row=7, column=6).value = authorizor_data[2]  # Address

        # address filling -start
        sheet_obj.cell(row=10, column=2).font = Font(size=10, name='Arial',
                                                     bold=False)
        sheet_obj.cell(row=10, column=2).value = member_data[2]  # ship to name
        sheet_obj.cell(row=11, column=2).font = Font(size=10, name='Arial',
                                                     bold=False)
        sheet_obj.cell(row=11, column=2).value = member_data[7]  # address
        sheet_obj.cell(row=12, column=2).font = Font(size=10, name='Arial',
                                                     bold=False)
        sheet_obj.cell(row=12, column=2).value = member_data[8] + "," + member_data[9]  # city ,state
        sheet_obj.cell(row=13, column=2).font = Font(size=10, name='Arial',
                                                     bold=False)
        sheet_obj.cell(row=13, column=2).value = member_data[12] + "," + member_data[10]  # country ,pincode
        sheet_obj.cell(row=14, column=2).font = Font(size=10, name='Arial',
                                                     bold=False)
        sheet_obj.cell(row=14, column=2).value = member_data[11]  # contact
        sheet_obj.cell(row=15, column=2).font = Font(size=10, name='Arial',
                                                     bold=False)
        sheet_obj.cell(row=15, column=2).value = member_data[14]  # email
        # address filling end

        sheet_obj.cell(row=18, column=2).font = Font(size=10, name='Arial',
                                                     bold=False)
        sheet_obj.cell(row=18, column=2).value = magazineCategoryText.get()
        sheet_obj.cell(row=18, column=4).font = Font(size=10, name='Arial',
                                                     bold=False)
        sheet_obj.cell(row=18, column=4).value = quantityText.get()
        sheet_obj.cell(row=18, column=5).font = Font(size=10, name='Arial',
                                                     bold=False)
        sheet_obj.cell(row=18, column=5).value = seva_amountText.get()
        sheet_obj.cell(row=18, column=6).font = Font(size=10, name='Arial',
                                                     bold=False)
        sheet_obj.cell(row=18, column=6).value = str(int(quantityText.get()) * int(seva_amountText.get()))
        sheet_obj.cell(row=21, column=6).font = Font(size=10, name='Arial',
                                                     bold=False)
        sheet_obj.cell(row=21, column=6).value = str(int(quantityText.get()) * int(seva_amountText.get()))
        sheet_obj.cell(row=22, column=6).font = Font(size=10, name='Arial',
                                                     bold=False)
        sheet_obj.cell(row=22, column=6).value = paymentMode_text.get()

        wb_obj.save(file_name)
        dest_file = "..\\Expanse_Data\\" + currentYearDirectory + "\\Magazine_Subscription\\Receipt\\Receipts\\" + invoice_id + ".xlsx"
        pdf_file = "..\\Expanse_Data\\" + currentYearDirectory + "\\Magazine_Subscription\\Receipt\\Receipts\\" + invoice_id + ".pdf"
        copyfile(file_name, dest_file)
        self.obj_commonUtil.convertExcelToPdf(dest_file, pdf_file)
        os.remove(dest_file)
        destdir_repo = self.obj_initDatabase.get_invoice_directory_name()
        desktop_repo = self.obj_initDatabase.get_desktop_invoices_directory_path()
        shutil.copy(pdf_file, desktop_repo)
        shutil.copy(pdf_file, destdir_repo)

        os.startfile(pdf_file, 'print')

    def generateExpanseVoucher(self, donator_idText,
                               seva_amountText,
                               categoryText,
                               collector_nameText,
                               dateOfCollection_calc,
                               paymentMode_text,
                               invoice_id, member_data,
                               trust_nametext,
                               var,
                               receiver_nameText,
                               receiver_phonenoText):

        currentYearDirectory = self.obj_commonUtil.getCurrentYearFolderName()

        if trust_nametext.get() == "Gaushala Seva":
            trustType = SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST
            file_name = "..\\Expanse_Data\\" + currentYearDirectory + "\\Expanse\\Receipts\\Template\\Gaushala_Voucher_Receipt_Template.xlsx"
        else:
            trustType = VIHANGAM_YOGA_KARNATAKA_TRUST
            file_name = "..\\Expanse_Data\\" + currentYearDirectory + "\\Expanse\\Receipts\\Template\\Ashram_Voucher_Receipt_Template.xlsx"

        wb_obj = openpyxl.load_workbook(file_name)
        sheet_obj = wb_obj.active

        sheet_obj.cell(row=6, column=4).value = invoice_id
        sheet_obj.cell(row=7, column=4).value = str(dateOfCollection_calc)
        print("Var is :", var)
        if var.get() == 1:
            sheet_obj.cell(row=8, column=4).value = str(member_data[2])
            sheet_obj.cell(row=9, column=4).value = str(member_data[7])  # Address
            sheet_obj.cell(row=10, column=4).value = str(member_data[8])  # city
            sheet_obj.cell(row=11, column=4).value = str(member_data[9])  # state
            sheet_obj.cell(row=12, column=4).value = str(member_data[10])
            sheet_obj.cell(row=13, column=4).value = str(member_data[11])
        else:
            sheet_obj.cell(row=8, column=4).value = receiver_nameText.get()
            sheet_obj.cell(row=9, column=4).value = "Not Available"  # Address
            sheet_obj.cell(row=10, column=4).value = "Not Available"  # city
            sheet_obj.cell(row=11, column=4).value = "Not Available"  # state
            sheet_obj.cell(row=12, column=4).value = "Not Available"
            sheet_obj.cell(row=13, column=4).value = receiver_phonenoText.get()

        sheet_obj.cell(row=15, column=4).value = str(categoryText.get())
        sheet_obj.cell(row=16, column=4).value = str(seva_amountText.get())
        sheet_obj.cell(row=17, column=4).value = str(num2words.num2words(int(seva_amountText.get()))) + " Rs. only"
        sheet_obj.cell(row=18, column=4).value = str(paymentMode_text.get())
        sheet_obj.cell(row=19, column=4).value = str(collector_nameText.get())

        wb_obj.save(file_name)
        dest_file = "..\\Expanse_Data\\" + currentYearDirectory + "\\Expanse\\Receipts\\Receipts\\" + invoice_id + ".xlsx "
        pdf_file = "..\\Expanse_Data\\" + currentYearDirectory + "\\Expanse\\Receipts\\Receipts\\" + invoice_id + ".pdf"
        copyfile(file_name, dest_file)
        self.obj_commonUtil.convertExcelToPdf(dest_file, pdf_file)
        os.remove(dest_file)
        os.startfile(pdf_file)

    def borrow_bookForm(self):
        self.bookBorrow_window = Toplevel(self.master)
        self.bookBorrow_window.title("Borrow Book ")
        self.bookBorrow_window.geometry('355x255+450+280')
        self.bookBorrow_window.configure(background='wheat')
        self.bookBorrow_window.resizable(width=False, height=False)

        heading = Label(self.bookBorrow_window, text="Book Borrow Requisition", font=('arial', 17, 'bold'),
                        bg='wheat')
        upperFrame = Frame(self.bookBorrow_window, width=210, height=100, bd=4, relief='ridge', bg='snow')
        # create a Book Name label
        nameLabel = Label(upperFrame, text="Member Id", width=10, anchor=W, justify=LEFT,
                          font=NORM_VERDANA_FONT,
                          bg='snow')
        booknameLabel = Label(upperFrame, text="Book Id", width=10, anchor=W, justify=LEFT,
                              font=NORM_VERDANA_FONT,
                              bg='snow')
        date_Label = Label(upperFrame, text="Date", width=10, anchor=W, justify=LEFT,
                           font=('arial narrow', 13, 'normal'), bg='snow')

        heading.grid(row=0, column=0, columnspan=2, padx=20)
        upperFrame.grid(row=1, column=0, columnspan=2, padx=20)
        nameLabel.grid(row=1, column=0, padx=10, pady=5)
        booknameLabel.grid(row=2, column=0, padx=10, pady=5)
        date_Label.grid(row=3, column=0, padx=10, pady=5)

        member_Id = Entry(upperFrame, width=25, font=('arial narrow', 11, 'normal'), bg='light yellow')
        book_IdText = Entry(upperFrame, width=25, font=('arial narrow', 11, 'normal'), bg='light yellow')
        cal = DateEntry(upperFrame, width=22, date_pattern='dd/MM/yyyy', font=('arial narrow', 12, 'normal'),
                        justify=LEFT)
        member_Id.grid(row=1, column=1, padx=10, pady=5)
        book_IdText.grid(row=2, column=1, padx=10, pady=5)
        cal.grid(row=3, column=1, padx=10, pady=5)

        infoFrame = Frame(self.bookBorrow_window, width=200, height=100, bd=4, relief='ridge', bg='snow')
        infoLabel = Label(infoFrame, text="All details are mandatory", width=35, anchor='center', justify=CENTER,
                          font=NORM_VERDANA_FONT,
                          bg='snow', fg='green')
        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(self.bookBorrow_window, width=200, height=100, bd=4, relief='ridge')

        buttonFrame.grid(row=8, column=0, pady=6, padx=10, columnspan=2)
        infoFrame.grid(row=9, column=0, pady=6, padx=10, columnspan=2)
        infoLabel.grid(row=1, column=0, pady=6, padx=2, columnspan=2)
        insert_result = partial(self.borrow_book_Excel, self.bookBorrow_window, book_IdText, member_Id, cal, infoLabel)

        # create a Borrow Button and place into the self.bookBorrow_window window
        submit = Button(buttonFrame, text="Borrow", fg="Black", command=insert_result,
                        font=NORM_FONT, width=8, bg='light cyan')
        submit.grid(row=0, column=0)

        clear_result = partial(self.clearBookBorrowForm, self.bookBorrow_window, book_IdText, member_Id)
        # create a Reser Button and place into the self.bookBorrow_window window
        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=1)

        # create a Close Button and place into the self.bookBorrow_window window
        cancel_result = partial(self.destroyWindow, self.bookBorrow_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_result,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=2)
        # ---------------------------------Button Frame End----------------------------------------

        # shortcut keys for keyboard actions
        self.bookBorrow_window.bind('<Return>', lambda event=None: submit.invoke())
        self.bookBorrow_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.bookBorrow_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.bookBorrow_window.focus()
        self.bookBorrow_window.grab_set()
        mainloop()

    def purchase_bookForm(self):
        self.bookPurchase_window = Toplevel(self.master)
        self.bookPurchase_window.title("Purchase Book/Sukrit Product ")
        self.bookPurchase_window.geometry('690x540+280+200')
        self.bookPurchase_window.configure(background='wheat')
        self.bookPurchase_window.resizable(width=False, height=False)

        heading = Label(self.bookPurchase_window, text="Purchase Book/Sukrit Product Requisition",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        heading.grid(row=0, column=0)
        # upper frame start
        itemFrame = Frame(self.bookPurchase_window, width=210, height=100, bd=4, relief='ridge', bg='snow')
        item_detailsFrame = Frame(self.bookPurchase_window, width=210, height=100, bd=4, relief='ridge', bg='snow')
        upperFrame = Frame(self.bookPurchase_window, width=210, height=100, bd=4, relief='ridge', bg='snow')

        itemFrame.grid(row=1, column=0, padx=20, pady=5)
        item_detailsFrame.grid(row=2, column=0, padx=20, pady=5)
        upperFrame.grid(row=3, column=0, padx=20, pady=5)

        default_text1 = StringVar(itemFrame, value='')

        # design item frame

        itembtnFrame = Frame(itemFrame, width=210, height=100, bd=4, relief='ridge', bg='snow')

        search_btn = Button(itembtnFrame, text="Search", fg="Black",
                            font=NORM_FONT, width=12, bg='light grey', state=DISABLED)
        close_btn = Button(itembtnFrame, text="Close", fg="Black",
                           font=NORM_FONT, width=12, bg='light cyan', state=NORMAL)
        item_id = Label(itemFrame, text="Item Id", width=12, anchor=W, justify=LEFT,
                        font=('arial', 13, 'normal'),
                        bg='snow')
        item_id.grid(row=0, column=0, pady=5)
        itemId_Text = Entry(itemFrame, text="", width=29, justify=CENTER,
                            font=('arial narrow', 15, 'normal'),
                            bg='light yellow', textvariable=default_text1)
        itemId_Text.grid(row=0, column=1, pady=5)

        center_namelabel = Label(itemFrame, text="Center Name", width=12, anchor=W, justify=LEFT,
                                 font=('arial', 13, 'normal'),
                                 bg='snow')
        center_namelabel.grid(row=1, column=0, pady=5)
        local_centerText = StringVar(itemFrame)
        localCenterList = self.obj_commonUtil.getLocalCenterNames()
        print("Center list  - ", localCenterList)
        local_centerText.set(localCenterList[0])
        localcenter_menu = OptionMenu(itemFrame, local_centerText, *localCenterList)
        localcenter_menu.configure(width=32, font=('arial narrow', 12, 'normal'), bg='snow', anchor=W, justify=LEFT)
        localcenter_menu.grid(row=1, column=1, pady=5)
        itembtnFrame.grid(row=0, column=2, padx=5)
        search_btn.grid(row=0, column=0, padx=1)
        close_btn.grid(row=0, column=1, padx=1)
        searchinfo_label = Label(itemFrame, text="Enter Item Id for purchase", width=40, anchor='center',
                                 justify=CENTER,
                                 font=('arial narrow', 13, 'normal'),
                                 bg='snow', fg='green')
        searchinfo_label.grid(row=2, column=0, padx=1, columnspan=3)

        # design item details frame - starts
        itemid_label = Label(item_detailsFrame, text="Item Id", width=12, anchor=W, justify=LEFT,
                             font=('arial narrow', 13, 'normal'),
                             bg='snow')
        itemid_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                 font=('arial narrow', 13, 'normal'),
                                 bg='light yellow')

        itemname_label = Label(item_detailsFrame, text="Item Name", width=12, anchor=W, justify=LEFT,
                               font=('arial narrow', 13, 'normal'),
                               bg='snow')
        itemname_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                   font=('arial narrow', 13, 'normal'),
                                   bg='light yellow')

        authordetails_label = Label(item_detailsFrame, text="Author", width=12, anchor=W, justify=LEFT,
                                    font=('arial narrow', 13, 'normal'),
                                    bg='snow')
        authordetails_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                        font=('arial narrow', 13, 'normal'),
                                        bg='light yellow')
        unitprice_label = Label(item_detailsFrame, text="Unit Price(Rs.)", width=12, anchor=W, justify=LEFT,
                                font=('arial narrow', 13, 'normal'),
                                bg='snow')
        unitprice_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                    font=('arial narrow', 13, 'normal'),
                                    bg='light yellow')

        quantitydetails_label = Label(item_detailsFrame, text="Stock Quantity", width=12, anchor=W, justify=LEFT,
                                      font=('arial narrow', 13, 'normal'),
                                      bg='snow')
        quantitydetails_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                          font=('arial narrow', 13, 'normal'),
                                          bg='light yellow')

        itemid_label.grid(row=0, column=0, pady=5)
        itemid_labelText.grid(row=0, column=1, padx=5, pady=5)
        itemname_label.grid(row=0, column=2, pady=5)
        itemname_labelText.grid(row=0, column=3, padx=5, pady=5)
        authordetails_label.grid(row=1, column=0, pady=5)
        authordetails_labelText.grid(row=1, column=1, padx=5, pady=5)
        unitprice_label.grid(row=1, column=2, pady=5)
        unitprice_labelText.grid(row=1, column=3, padx=5, pady=5)
        quantitydetails_label.grid(row=2, column=0, pady=5)
        quantitydetails_labelText.grid(row=2, column=1, padx=5, pady=5)
        searchitemid_result = partial(self.searchStockItemRecords, itemId_Text, itemid_labelText,
                                      itemname_labelText, authordetails_labelText,
                                      unitprice_labelText, quantitydetails_labelText, searchinfo_label,
                                      local_centerText)
        search_btn.configure(command=searchitemid_result)
        cancel_result = partial(self.destroyWindow, self.bookPurchase_window)
        close_btn.configure(command=cancel_result)
        # design item details frame - end

        member_idLabel = Label(upperFrame, text="Member Id", width=10, anchor=W, justify=LEFT,
                               font=('arial narrow', 13, 'normal'), bg='snow')

        member_IdText = Entry(upperFrame, width=22, font=('arial narrow', 13, 'normal'), justify=LEFT,
                              bg='light yellow', state=DISABLED)
        customer_namelabel = Label(upperFrame, text="Customer Name", width=13, anchor=W, justify=LEFT,
                                   font=('arial narrow', 13, 'normal'), bg='snow')
        customer_nameText = Entry(upperFrame, width=22, font=('arial narrow', 13, 'normal'), justify=LEFT,
                                  state=DISABLED)

        item_idLabel = Label(upperFrame, text="Item Id (CI-)", width=10, anchor=W, justify=LEFT,
                             font=('arial narrow', 13, 'normal'), bg='snow')
        item_idText = Entry(upperFrame, width=22, font=('arial narrow', 13, 'normal'), justify=LEFT, bg='light yellow',
                            textvariable=default_text1)

        date_Label = Label(upperFrame, text="Date", width=13, anchor=W, justify=LEFT,
                           font=('arial narrow', 13, 'normal'), bg='snow')

        cal = DateEntry(upperFrame, width=22, date_pattern='dd/MM/yyyy', font=('arial narrow', 12, 'normal'),
                        justify=LEFT)

        quantity_label = Label(upperFrame, text="Quantity", width=10, anchor=W, justify=LEFT,
                               font=('arial narrow', 13, 'normal'), bg='snow')

        quantityText = Entry(upperFrame, width=22, font=('arial narrow', 13, 'normal'), justify=LEFT, bg='light yellow')

        paymentmode_label = Label(upperFrame, text="Payment Mode", width=13, anchor=W, justify=LEFT,
                                  font=('arial narrow', 13, 'normal'), bg='snow')
        paymentMode_text = StringVar(upperFrame)
        paymentMode_list = ['Cash', 'Bank Transfer', 'Cheque', 'Demand Draft', 'Other']
        paymentMode_text.set("Other")
        paymentMode_menu = OptionMenu(upperFrame, paymentMode_text, *paymentMode_list)
        paymentMode_menu.configure(width=20, font=('arial narrow', 12, 'normal'), bg='light yellow', anchor=W,
                                   justify=LEFT)

        cart_item_countLabel = Label(upperFrame, text="Cart Count", width=10, anchor=W, justify=LEFT,
                                     font=('arial narrow', 13, 'normal'), bg='snow')

        cart_item_count = Label(upperFrame, text="0", width=22, anchor='center', justify=LEFT,
                                font=('arial narrow', 13, 'normal'), bg='snow')

        amount_payableLabel = Label(upperFrame, text="Bill Amount(Rs.)", width=13, anchor=W, justify=LEFT,
                                    font=('arial narrow', 13, 'normal'), bg='snow')

        amount_payableTextLabel = Label(upperFrame, text="0", width=10, justify=CENTER,
                                        font=('arial narrow', 14, 'bold'), bg='navy', fg='white')

        var = IntVar()
        viewPurchaseBy_Result = partial(self.enablePurchaseViewBy_RadioSelection, var, member_IdText,
                                        customer_nameText, itemId_Text, item_idText, cart_item_count,
                                        amount_payableTextLabel)
        purchasebyMemeberId_radioBtn = Radiobutton(upperFrame, text="Purchase by ID", variable=var, value=1,
                                                   command=viewPurchaseBy_Result, width=12, bg='snow',
                                                   font=('arial narrow', 12, 'normal'), anchor=W, justify=LEFT)
        purchasebyName_radioBtn = Radiobutton(upperFrame, text="Purchase by Name", variable=var, value=2,
                                              command=viewPurchaseBy_Result, width=14, bg='snow',
                                              font=('arial narrow', 12, 'normal'), anchor=W, justify=LEFT)

        heading.grid(row=0, column=0, columnspan=3, padx=10)
        purchasebyMemeberId_radioBtn.grid(row=1, column=0, columnspan=1)
        purchasebyName_radioBtn.grid(row=1, column=2)
        member_idLabel.grid(row=2, column=0)
        member_IdText.grid(row=2, column=1)
        customer_namelabel.grid(row=2, column=2, padx=5, pady=5)
        customer_nameText.grid(row=2, column=3, pady=5)
        item_idLabel.grid(row=3, column=0)
        item_idText.grid(row=3, column=1, pady=5)
        date_Label.grid(row=3, column=2)
        cal.grid(row=3, column=3, padx=5, pady=5)
        quantity_label.grid(row=4, column=0)
        quantityText.grid(row=4, column=1)
        paymentmode_label.grid(row=4, column=2)
        paymentMode_menu.grid(row=4, column=3)
        cart_item_countLabel.grid(row=5, column=0)
        cart_item_count.grid(row=5, column=1, padx=5, pady=5)
        amount_payableLabel.grid(row=5, column=2, padx=5)
        amount_payableTextLabel.grid(row=5, column=3)
        # upper frame end

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(upperFrame, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=0, pady=6, columnspan=4)

        purchase_btn = Button(buttonFrame, text="Purchase", fg="Black",
                              font=NORM_FONT, width=12, bg='light grey', state=DISABLED)
        print_btn = Button(buttonFrame, text="Print Invoice", fg="Black",
                           font=NORM_FONT, width=12, bg='light grey', state=DISABLED)
        addtocart_btn = Button(buttonFrame, text="Add to Cart", fg="Black", font=NORM_FONT, width=12, bg='light cyan')
        insert_result = partial(self.addto_cart_item, self.bookPurchase_window,
                                item_idText,
                                member_IdText,
                                purchase_btn,
                                print_btn,
                                addtocart_btn,
                                quantityText,
                                cart_item_count,
                                amount_payableTextLabel,
                                paymentMode_text,
                                searchinfo_label,
                                var,
                                cal, local_centerText)
        addtocart_btn.configure(command=insert_result)
        result_searcbtnState = partial(self.check_itemSearchBtnState, default_text1, search_btn, addtocart_btn)
        default_text1.trace("w", result_searcbtnState)

        addtocart_btn.grid(row=0, column=0)
        purchase_result = partial(self.purchase_stock_item,
                                  member_IdText, customer_nameText,
                                  purchase_btn, print_btn, addtocart_btn, quantityText,
                                  amount_payableTextLabel,
                                  paymentMode_text, searchinfo_label, var, cal, local_centerText)
        purchase_btn.configure(command=purchase_result)
        purchase_btn.grid(row=0, column=1)
        print_btn.grid(row=0, column=2)

        clear_result = partial(self.clearBookPurchaseForm, itemId_Text,
                               itemname_labelText,
                               authordetails_labelText,
                               unitprice_labelText,
                               quantitydetails_labelText,
                               cart_item_count,
                               amount_payableTextLabel,
                               member_IdText,
                               customer_nameText,
                               quantityText,
                               paymentMode_text,
                               itemid_labelText,
                               searchinfo_label)
        # create a Reset Button and place into the self.bookPurchase_window window
        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=12, bg='light cyan', underline=0)
        clear.grid(row=0, column=3)

        # ---------------------------------Button Frame End----------------------------------------
        # shortcut keys for keyboard actions
        self.bookPurchase_window.bind('<Return>', lambda event=None: search_btn.invoke())
        self.bookPurchase_window.bind('<Alt-c>', lambda event=None: close_btn.invoke())
        self.bookPurchase_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.bookPurchase_window.focus()
        self.bookPurchase_window.grab_set()
        mainloop()

    def advance_returnForm(self):
        self.advance_return_window = Toplevel(self.master)
        self.advance_return_window.title("Return Advance ")
        self.advance_return_window.geometry('690x340+280+200')
        self.advance_return_window.configure(background='wheat')
        self.advance_return_window.resizable(width=False, height=False)

        heading = Label(self.advance_return_window, text="Advance Return Form",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        heading.grid(row=0, column=0)
        # upper frame start
        itemFrame = Frame(self.advance_return_window, width=210, height=100, bd=4, relief='ridge', bg='snow')
        item_detailsFrame = Frame(self.advance_return_window, width=210, height=100, bd=4, relief='ridge', bg='snow')
        upperFrame = Frame(self.advance_return_window, width=210, height=100, bd=4, relief='ridge', bg='snow')

        itemFrame.grid(row=1, column=0, padx=20, pady=5)
        item_detailsFrame.grid(row=2, column=0, padx=10, pady=5)
        upperFrame.grid(row=3, column=0, padx=10, pady=5)

        default_text1 = StringVar(itemFrame, value='')

        # design item frame

        itembtnFrame = Frame(itemFrame, width=210, height=100, bd=4, relief='ridge', bg='snow')

        search_btn = Button(itembtnFrame, text="Search", fg="Black",
                            font=NORM_FONT, width=13, bg='light cyan', state=NORMAL)
        close_btn = Button(itembtnFrame, text="Close", fg="Black",
                           font=NORM_FONT, width=13, bg='light cyan', state=NORMAL)
        item_id = Label(itemFrame, text="Advance Id", width=12, anchor=W, justify=LEFT,
                        font=('arial narrow', 13, 'normal'),
                        bg='snow')
        item_id.grid(row=0, column=0, pady=5)
        itemId_Text = Entry(itemFrame, text="", width=29, justify=CENTER,
                            font=('arial narrow', 15, 'normal'),
                            bg='light yellow', textvariable=default_text1)
        itemId_Text.grid(row=0, column=1, pady=5)
        itembtnFrame.grid(row=0, column=2, padx=5)
        search_btn.grid(row=0, column=0, padx=1)
        close_btn.grid(row=0, column=1, padx=1)
        searchinfo_label = Label(itemFrame, text="Enter Advance Id to view details", width=40, anchor='center',
                                 justify=CENTER,
                                 font=('arial narrow', 13, 'normal'),
                                 bg='snow', fg='green')
        searchinfo_label.grid(row=1, column=0, padx=1, columnspan=3)

        # design item details frame - starts
        advanceId_label = Label(item_detailsFrame, text="Advance Id", width=12, anchor=W, justify=LEFT,
                                font=('arial narrow', 13, 'normal'),
                                bg='snow')
        advanceId_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                    font=('arial narrow', 13, 'normal'),
                                    bg='light yellow')
        issuedTo_label = Label(item_detailsFrame, text="Receiver(Id)", width=12, anchor=W, justify=LEFT,
                               font=('arial narrow', 13, 'normal'),
                               bg='snow')
        issuedTo_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                   font=('arial narrow', 13, 'normal'),
                                   bg='light yellow')

        issuedName_label = Label(item_detailsFrame, text="Receiver(Name)", width=12, anchor=W, justify=LEFT,
                                 font=('arial narrow', 13, 'normal'),
                                 bg='snow')
        issuedName_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                     font=('arial narrow', 13, 'normal'),
                                     bg='light yellow')
        dateofissue_label = Label(item_detailsFrame, text="Date of Issue", width=12, anchor=W, justify=LEFT,
                                  font=('arial narrow', 13, 'normal'),
                                  bg='snow')
        dateofissue_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                      font=('arial narrow', 13, 'normal'),
                                      bg='light yellow')

        issuedAmt_label = Label(item_detailsFrame, text="Issued Amt(Rs.)", width=12, anchor=W, justify=LEFT,
                                font=('arial narrow', 13, 'normal'),
                                bg='snow')
        issuedAmt_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                    font=('arial narrow', 13, 'normal'),
                                    bg='light yellow')

        description_label = Label(item_detailsFrame, text="Description", width=12, anchor=W, justify=LEFT,
                                  font=('arial narrow', 13, 'normal'),
                                  bg='snow')
        description_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                      font=('arial narrow', 13, 'normal'),
                                      bg='light yellow')

        amtReturned_labelText = Label(item_detailsFrame, text="Spent Amt.(Rs.)", width=12, anchor=W, justify=LEFT,
                                      font=('arial narrow', 13, 'normal'),
                                      bg='snow')
        amtReturnText = Entry(item_detailsFrame, width=25, font=('arial narrow', 13, 'normal'), justify=LEFT,
                              bg='light yellow', state=NORMAL)

        advanceId_label.grid(row=0, column=0, pady=5)
        advanceId_labelText.grid(row=0, column=1, padx=5, pady=5)
        issuedTo_label.grid(row=0, column=2, pady=5)
        issuedTo_labelText.grid(row=0, column=3, padx=5, pady=5)
        issuedName_label.grid(row=1, column=0, pady=5)
        issuedName_labelText.grid(row=1, column=1, padx=5, pady=5)
        dateofissue_label.grid(row=1, column=2, pady=5)
        dateofissue_labelText.grid(row=1, column=3, padx=5, pady=5)
        issuedAmt_label.grid(row=2, column=0, padx=5, pady=5)
        issuedAmt_labelText.grid(row=2, column=1, padx=5, pady=5)
        description_label.grid(row=2, column=2, padx=5, pady=5)
        description_labelText.grid(row=2, column=3, padx=5, pady=5)
        amtReturned_labelText.grid(row=3, column=0, pady=5)
        amtReturnText.grid(row=3, column=1, pady=5)
        savebtn = Button(upperFrame, text="Save", fg="Black",
                         font=NORM_FONT, width=12, bg='light cyan', underline=0)
        searchitemid_result = partial(self.searchAdvanceIdExcel,
                                      itemId_Text,
                                      advanceId_labelText,
                                      issuedTo_labelText,
                                      dateofissue_labelText,
                                      issuedName_labelText,
                                      issuedAmt_labelText,
                                      amtReturnText,
                                      description_labelText,
                                      searchinfo_label, savebtn)
        search_btn.configure(command=searchitemid_result)
        cancel_result = partial(self.destroyWindow, self.advance_return_window)
        close_btn.configure(command=cancel_result)
        # design item details frame - end

        # ---------------------------------Button Frame Start----------------------------------------
        save_result = partial(self.update_Advance_Excel, itemId_Text, amtReturnText, searchinfo_label)
        savebtn.configure(command=save_result)
        savebtn.grid(row=0, column=0)

        # create a Reset Button and place into the self.advance_return_window window
        reset_result = partial(self.clearAdvanceReturnForm, advanceId_labelText,
                               issuedTo_labelText,
                               dateofissue_labelText,
                               issuedName_labelText,
                               issuedAmt_labelText,
                               amtReturnText,
                               description_labelText,
                               searchinfo_label)
        clear = Button(upperFrame, text="Reset", fg="Black", command=reset_result,
                       font=NORM_FONT, width=12, bg='light cyan', underline=0)
        clear.grid(row=0, column=1)

        # ---------------------------------Button Frame End----------------------------------------
        # shortcut keys for keyboard actions
        self.advance_return_window.bind('<Return>', lambda event=None: search_btn.invoke())
        self.advance_return_window.bind('<Alt-c>', lambda event=None: close_btn.invoke())
        self.advance_return_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        self.advance_return_window.focus()
        self.advance_return_window.grab_set()
        mainloop()

    def enablePurchaseViewBy_RadioSelection(self, var, member_IdText,
                                            customer_nameText, itemId_Text, item_idText, cart_item_count,
                                            amount_payableTextLabel):
        print("Enabling the view by date section Radiobutton :", var.get())
        self.list_InvoicePrint = []
        cart_item_count.configure(text="0")
        amount_payableTextLabel.configure(text="0")
        # all elements are disabled in beginning
        # based on the selection of the radio button, respective ones are enabled
        member_IdText.configure(state=DISABLED)
        customer_nameText.configure(state=DISABLED)
        item_idText.configure(state=DISABLED)

        if var.get() == 1:
            print("Enabling view by date")
            member_IdText.configure(state=NORMAL, bg='light yellow')
            customer_nameText.insert(0, "")
        elif var.get() == 2:
            customer_nameText.configure(state=NORMAL, bg='light yellow')
            member_IdText.insert(0, "")

        else:
            pass

    def enableIntenalID_RadioSelection(self, var, receiver_IDText,
                                       receiver_nameText, receiver_phonenoText):
        print("Enabling the view by date section Radiobutton :", var.get())

        # all elements are disabled in beginning
        # based on the selection of the radio button, respective ones are enabled
        receiver_IDText.configure(state=DISABLED)
        receiver_nameText.configure(state=DISABLED)
        receiver_phonenoText.configure(state=DISABLED)

        if var.get() == 1:
            print("Enabling view by date")
            receiver_IDText.configure(state=NORMAL, bg='light yellow')
            receiver_IDText.insert(0, "")
        elif var.get() == 2:
            receiver_nameText.configure(state=NORMAL, bg='light yellow')
            receiver_nameText.insert(0, "")
            receiver_phonenoText.configure(state=NORMAL, bg='light yellow')
            receiver_phonenoText.insert(0, "")

        else:
            pass

    def book_ReturnForm(self):
        self.bookReturn_window = Toplevel(self.master)
        self.bookReturn_window.title("Return Book ")
        self.bookReturn_window.geometry('715x450+150+80')
        self.bookReturn_window.configure(background='wheat')
        self.bookReturn_window.resizable(width=False, height=False)

        # delete "X" button in window will be not-operational
        self.bookReturn_window.protocol('WM_DELETE_WINDOW', self.donothing)
        heading = Label(self.bookReturn_window, text="Book Return Requisition", font=('arial', 20, 'normal'),
                        bg='wheat')
        heading.grid(row=0, column=0)
        upperFrame = Frame(self.bookReturn_window, width=280, height=100, bd=4, relief='ridge', bg='snow')
        upperFrame.grid(row=1, column=0, padx=10, pady=10, sticky=W)

        self.middleFrame_bookdisplay = Frame(self.bookReturn_window, width=0, height=0, bd=8, relief='ridge')
        self.middleFrame_bookdisplay.grid(row=2, column=0, padx=20, pady=10, sticky=W)

        memberIdLabel = Label(upperFrame, text="Member Id", width=10, anchor=W, justify=LEFT,
                              font=('arial narrow', 11, 'normal'), bg='snow')
        dateOfReturn = Label(upperFrame, text="Date of Return", width=12, anchor=W, justify=LEFT,
                             font=('arial narrow', 11, 'normal'), bg='snow')

        memberIdLabel.grid(row=1, column=0, padx=10, pady=10)
        dateOfReturn.grid(row=1, column=2, padx=30, pady=10)

        member_Id = Entry(upperFrame, width=10, font=('arial narrow', 13, 'normal'), justify='left', bg='light yellow')
        member_Id.grid(row=1, column=1, padx=5, pady=10)

        cal = DateEntry(upperFrame, width=10, date_pattern='dd/MM/yyyy', font=('arial narrow', 13, 'normal'),
                        justify='center', bg='light yellow')

        cal.grid(row=1, column=3)
        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(upperFrame, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=1, column=4, padx=8, pady=10, sticky=W)

        search_result = partial(self.search_borrowRecords_Excel, self.bookReturn_window, member_Id, cal)

        # create a Search Button and place into the self.bookReturn_window window
        submit = Button(buttonFrame, text="Search", fg="Black", command=search_result,
                        font=NORM_FONT, width=11, bg='light cyan')
        submit.grid(row=0, column=0)

        # create a Close Button and place into the self.bookReturn_window window
        cancel_result = partial(self.destroyWindow, self.bookReturn_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_result,
                        font=NORM_FONT, width=11, bg='light cyan')
        cancel.grid(row=0, column=1)
        # ---------------------------------Button Frame End----------------------------------------

        self.bookReturn_window.bind('<Return>', lambda event=None: submit.invoke())
        self.bookReturn_window.bind('<Alt-c>', lambda event=None: cancel.invoke())

        self.bookReturn_window.focus()
        self.bookReturn_window.grab_set()
        mainloop()

    def enableViewBy_RadioSelection(self, var, fromDate, cal_dateFrom, toDate, cal_toDate,
                                    viewByMonth_Month, viewbyMonth_monthTxt, viewByMonth_Year, viewbymonth_yearTxt,
                                    viewByYear_Year, viewbyYear_yearTxt):
        print("Enabling the view by date section Radiobutton :", var.get())

        # all elements are disabled in begining
        # based on the selection of the radio button, respective ones are enabled

        fromDate.configure(state=DISABLED)
        cal_dateFrom.configure(state=DISABLED)
        toDate.configure(state=DISABLED)
        cal_toDate.configure(state=DISABLED)
        viewByMonth_Month.configure(state=DISABLED)
        viewbyMonth_monthTxt.configure(state=DISABLED)
        viewByMonth_Year.configure(state=DISABLED)
        viewbymonth_yearTxt.configure(state=DISABLED)
        viewByYear_Year.configure(state=DISABLED)
        viewbyYear_yearTxt.configure(state=DISABLED)

        if var.get() == 1:
            print("Enabling view by date")
            fromDate.configure(state=NORMAL, bg='light yellow')
            cal_dateFrom.configure(state=NORMAL)
            toDate.configure(state=NORMAL, bg='light yellow')
            cal_toDate.configure(state=NORMAL)
        elif var.get() == 2:
            viewByMonth_Month.configure(state=NORMAL, bg='light yellow')
            viewbyMonth_monthTxt.configure(state=NORMAL)
            viewByMonth_Year.configure(state=NORMAL, bg='light yellow')
            viewbymonth_yearTxt.configure(state=NORMAL)
        elif var.get() == 3:
            viewByYear_Year.configure(state=NORMAL, bg='light yellow')
            viewbyYear_yearTxt.configure(state=NORMAL)
        else:
            pass

    def display_book_info(self):
        self.display_bookInfo_window = Toplevel(self.master)
        bgcolor = 'antiquewhite'
        self.display_bookInfo_window.title("Display Book Info ")
        self.display_bookInfo_window.geometry('550x650+150+80')
        self.display_bookInfo_window.configure(background='powder blue')
        # self.display_bookInfo_window.resizable(width=False, height=False)

        # delete "X" button in window will be not-operational
        self.display_bookInfo_window.protocol('WM_DELETE_WINDOW', self.donothing)
        upperFrame = Frame(self.display_bookInfo_window, width=205, height=100, bd=4, relief='ridge', bg=bgcolor)
        upperFrame.grid(row=1, column=2, padx=30, pady=10, sticky=W)
        heading = Label(upperFrame, text="\tDisplay Book Info", font=('times new roman', 25, 'normal'), bg=bgcolor)

        middleFrame = Frame(self.display_bookInfo_window, width=200, height=150, bd=8, relief='ridge')
        middleFrame.grid(row=2, column=2, padx=30, pady=10, sticky=W)

        bookId_label = Label(upperFrame, text="Book Id", width=9, anchor=W, justify=LEFT,
                             font=('arial narrow', 14, 'normal'), bg=bgcolor)

        bookName_label = Label(upperFrame, text="Book Name", width=9, anchor=W, justify=LEFT,
                               font=('arial narrow', 14, 'normal'), bg=bgcolor)

        condition_label = Label(upperFrame, text="OR", width=3, anchor=W, justify=CENTER,
                                font=('arial narrow', 16, 'bold'), fg='blue', bg=bgcolor)
        bookDetailsLabel = Label(middleFrame, text="No Data", width=45, height=14, anchor=W, justify=LEFT,
                                 font=('arial narrow', 15, 'normal'),
                                 bg='light yellow')

        # heading.grid(row=0, column=0)
        bookId_label.grid(row=1, column=0, padx=20, pady=10)

        book_id = Entry(upperFrame, width=7, font=('arial narrow', 15, 'normal'), justify='center')
        book_id.grid(row=1, column=1, padx=10, ipadx=60, pady=10)
        condition_label.grid(row=2, column=1, padx=5, pady=5)
        bookName_label.grid(row=3, column=0, padx=5, pady=5)

        book_name = Entry(upperFrame, width=7, font=('arial narrow', 15, 'normal'), justify='center')
        book_name.grid(row=3, column=1, padx=10, ipadx=60, pady=10)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(upperFrame, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=4, column=0, padx=65, columnspan=2, pady=10, sticky=W)

        printbtn = Button(buttonFrame, text="Print", fg="Black", command=None,
                          font=NORM_FONT, width=11, bg='light grey', state=DISABLED)

        search_result = partial(self.search_bookInfo_Excel, book_id, book_name, bookDetailsLabel, printbtn)
        # create a Search Button and place into the self.display_bookInfo_window window
        submit = Button(buttonFrame, text="Search", fg="Black", command=search_result,
                        font=NORM_FONT, width=11, bg='light cyan')
        submit.grid(row=0, column=0)

        # create a Search Button and place into the self.display_bookInfo_window window

        printbtn.grid(row=0, column=1)

        # create a Close Button and place into the self.display_bookInfo_window window
        cancel_result = partial(self.destroyWindow, self.display_bookInfo_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=cancel_result,
                        font=NORM_FONT, width=11, bg='light cyan')
        cancel.grid(row=0, column=2)
        # ---------------------------------Button Frame End----------------------------------------
        self.display_bookInfo_window.bind('<Return>', lambda event=None: submit.invoke())
        self.display_bookInfo_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        self.display_bookInfo_window.bind('<Alt-c>', lambda event=None: printbtn.invoke())

        bookDetailsLabel.grid(row=2, column=0, padx=5, pady=15, sticky=W)

        self.display_bookInfo_window.focus()
        self.display_bookInfo_window.grab_set()
        mainloop()

    def getStaffUserName(self, memberId, memtype):
        print("getStaffUserName-> Start")
        member_name = ""
        if memtype == 1:
            file_name = "Member.txt"
        if memtype == 2:
            file_name = "Staff.txt"
        # Fail-safe protection  - if database is deleted anonmously at back end while reaching here
        if not os.path.isfile(file_name):
            messagebox.showerror("Database error", "No Members available ....")
            return
        member_file = open(file_name, 'r')
        for line in member_file:
            record = line.split(',')
            if record[0] == memberId:
                member_name = record[2]
                break
        print("getStaffUserName-> End")
        return member_name

    def getStaffUserName_Excel(self, memberId, memtype):
        # print("getStaffUserName-> Start memberId :", memberId)
        member_name = ""
        member_designation = ""
        file_name = PATH_MEMBER

        # Fail-safe protection  - if database is deleted anonymously at back end while reaching here
        if not os.path.isfile(file_name):
            messagebox.showerror("Database error", "No Members available ....")
            return
        wb_obj = openpyxl.load_workbook(file_name)
        sheet_obj = wb_obj.active
        total_record = self.totalrecords_excelDataBase(file_name)
        for iLoop in range(0, total_record):
            if str(sheet_obj.cell(row=iLoop + 2, column=2).value) == memberId:
                member_name = str(sheet_obj.cell(row=iLoop + 2, column=4).value)
                member_designation = str(sheet_obj.cell(row=iLoop + 2, column=24).value)
                break
        # print("getStaffUserName Excel-> End member name :", member_name)
        return member_name, member_designation

    def validateStaffPassword(self, username, password):
        login_file = open("StaffLogin.txt", 'r')
        buserOk = False
        for line in login_file:
            record = line.split(',')
            pwd_entry = record[1].splitlines()
            if username == record[0] and password == pwd_entry[0]:
                buserOk = True
                break
        print("validateStaffPassword->>", buserOk)
        return buserOk

    def validateStaffPassword_Excel(self, username, password):
        wb_obj = openpyxl.load_workbook(PATH_STAFF_CREDENTIALS)
        sheet_obj = wb_obj.active
        buserOk = False
        total_records = self.totalrecords_excelDataBase(PATH_STAFF_CREDENTIALS)
        # print("User entered  - username : ", username, "Password :", password, "Total records:", total_records)
        for iLoop in range(0, total_records):
            if str(sheet_obj.cell(row=iLoop + 2, column=2).value) == username:
                # print("Match found for username")
                if str(sheet_obj.cell(row=iLoop + 2, column=3).value) == password:
                    buserOk = True
                    # print("Match found --- aborting loop now")
                    break
        print("validateStaffPassword_Excel->>End", buserOk)
        return buserOk

    def validateStaff_Login_Excel(self, login_window, lowerFrame, buttonFrame, username, password, labelLogin):
        # user is permitted 3 trials of login, system exit beyond that
        print("validateStaff_Login_Excel - >Entry")

        wb_obj = openpyxl.load_workbook(PATH_STAFF_CREDENTIALS)
        sheet_obj = wb_obj.active

        buserOk = self.validateStaffPassword_Excel(username.get(), password.get())
        print("buserOk", buserOk)
        strUserName, strUserDesignation = self.getStaffUserName_Excel(username.get(), 1)
        print("Staff Name :", strUserName, "Designation :", strUserDesignation)
        if buserOk is True:
            self.logged_staff_id = username.get()
            self.logged_staff_name = strUserName
            labelLogin.configure(height=4, font=('arial narrow', 18, 'normal'), bg='antiquewhite', fg='midnightblue')
            labelLogin['text'] = "Success !!\nWelcome " + strUserName

            lowerFrame.destroy()
            buttonFrame.destroy()
            closeFrame = Frame(login_window, width=200, height=100, bd=4, relief='ridge')
            closeFrame.grid(row=3, column=0, columnspan=2)
            # close_result = partial(self.closeFromLogin, login_window)
            cancel = Button(closeFrame, text="Close", fg="Black", command=login_window.destroy,
                            font=NORM_FONT, width=9, bg='khaki')
            cancel.grid(row=0, column=0)
            self.main_screen_design(strUserDesignation)
            # ensures that Close button reacts by pressing keyboard "Enter" key
            login_window.bind('<Return>', lambda event=None: cancel.invoke())
        else:
            labelLogin.configure(fg='red')
            labelLogin['text'] = "Login Failed !! Try Again"
            self.clear_loginForm(username, password)
            username.focus()
        mainloop()
        print("validateStaff_Login - >End")

    def closeFromLogin(self, window_name):
        self.obj_commonUtil.encryptDatabase()
        window_name.destroy()

    def main_screen_design(self, userControlPrivilage):

        # prepares the menu bar
        self.main_bar = Menu(self.master)
        self.main_bar.config()
        self.main_bar.entryconfig(1, state=DISABLED)

        '''
        ############################################################################
        # main menu bar contains 4  - menus
        # 1. New
        #       - Stock Entry
        #            - Commercial(Books & Sukrit) Entry
        #            - Non-Commercial(Others) Entry
        #       - Member Registration
        #       - Expanse
        #       - Donation
        #           - Monetary
        #           - Non-Monetary
        #       - Magazine Subscription
        # 2. Stock Transaction
        #       - Sell Stock
        #       - Borrow Stock
        #       - Return Stock
        # 3. Edit
        #       - Member Info
        #       - Stock Details Info
        #           - Commercial(Books & Sukrit)
        #           - Non-Commercial(Others)
        #    
        # 4. View
        #       - Member Info
        #           - By Details
        #       - Inventory\Stock 
        #           - Critical Stock Info
        #           - Stock Info
        #               - Commercial(Books & Sukrit)
        #               - Non-Commercial(Others)
        #               - Stock Sales Statement
        #                   - Search by ID
        #                   - Generic Sales Statement
        #               - Complete Stock Info
        #                   - Commercial
        #                   - Non-Commercial
        #       - Donation Statement
        #       - Ashram Account Statement
        #       - Gaushala Account Statement
        #  5. User
        #       - Current User
        #       - Change Password
        #       - Logout
        #       - Exit
        # #########################################################################
        '''

        self.file_menu = Menu(self.main_bar, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                              activeforeground='light yellow')

        self.stock_transaction_menu = Menu(self.main_bar, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                           activeforeground='light yellow')

        self.edit_menu = Menu(self.main_bar, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                              activeforeground='light yellow')
        self.option_menu = Menu(self.main_bar, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                activeforeground='light yellow')
        self.view_menu = Menu(self.main_bar, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                              activeforeground='light yellow')
        self.account_menu = Menu(self.main_bar, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                 activeforeground='light yellow')

        self.main_bar.add_cascade(label="New", menu=self.file_menu, underline=0)
        self.main_bar.add_cascade(label="Stock Transaction", menu=self.stock_transaction_menu, underline=0)
        if userControlPrivilage == "Staff-Sevak" or userControlPrivilage == "Manager" or userControlPrivilage == "Admin":
            self.main_bar.add_cascade(label='Edit', menu=self.edit_menu, underline=0)
        self.main_bar.add_cascade(label='View', menu=self.view_menu, underline=0)
        self.main_bar.add_cascade(label='User', menu=self.account_menu, underline=0)

        # "File" menu contents
        # ------
        item_entry_submenu = Menu(self.file_menu, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                  activeforeground='light yellow')
        self.file_menu.add_cascade(label='Stock Entry', menu=item_entry_submenu, state=NORMAL, underline=0)
        item_entry_submenu.add_command(label="Register New Author", command=self.register_new_author,
                                       state=NORMAL,
                                       underline=0)
        item_entry_submenu.add_command(label="Commercial(Books & Sukrit) Entry", command=self.item_entry_commercial,
                                       state=NORMAL,
                                       underline=0)

        result_entry_noncommercial_stock = partial(self.deposit_seva_nonmonetary_rashi, STOCK_OWNER_TYPE_ASHRAM)
        item_entry_submenu.add_command(label="Non-Commercial(Others) Entry", command=result_entry_noncommercial_stock,
                                       state=NORMAL,
                                       underline=0)
        memberType = 1
        addMember = partial(self.add_member, memberType)
        self.file_menu.add_cascade(label="Member Registration", command=addMember,
                                   state=NORMAL,
                                   underline=0)
        expanse_submenu = Menu(self.file_menu, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                               activeforeground='light yellow')
        expanse_submenu.add_cascade(label="Normal", command=self.create_expanse_entry,
                                    state=NORMAL,
                                    underline=0)
        expanse_submenu.add_cascade(label="Advance", command=self.create_advance_expanse_entry,
                                    state=NORMAL,
                                    underline=0)
        self.file_menu.add_cascade(label="Expanse", menu=expanse_submenu,
                                   state=NORMAL,
                                   underline=0)

        item_donation_submenu = Menu(self.file_menu, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                     activeforeground='light yellow')
        self.file_menu.add_cascade(label='Donation', menu=item_donation_submenu, state=NORMAL, underline=0)

        item_donation_submenu.add_cascade(label="Monetary", command=self.deposit_seva_rashi,
                                          state=NORMAL,
                                          underline=0)
        result_entry_noncommercial_donation_stock = partial(self.deposit_seva_nonmonetary_rashi,
                                                            STOCK_OWNER_TYPE_DONATED)
        item_donation_submenu.add_command(label="Non-Monetary", command=result_entry_noncommercial_donation_stock,
                                          state=NORMAL,
                                          underline=0)

        sankalp_submenu = Menu(self.file_menu, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                               activeforeground='light yellow')
        self.file_menu.add_cascade(label='Sankalp(Pledge)', menu=sankalp_submenu, state=NORMAL, underline=0)

        sankalp_submenu.add_cascade(label="Register New Pledge", command=self.register_new_sankalp_item,
                                    state=NORMAL,
                                    underline=0)
        sankalp_submenu.add_command(label="Take Pledge(Sankalp)", command=self.sankalp_form,
                                    state=NORMAL,
                                    underline=0)
        sankalp_submenu.add_command(label="Fulfill Pledge(Sankalp)", command=self.fulfill_pledge_form,
                                    state=NORMAL,
                                    underline=0)
        '''
        self.file_menu.add_cascade(label="Magazine Subscription", command=self.perform_patrika_subscription,
                                   state=NORMAL,
                                   underline=0)
        '''
        self.file_menu.add_cascade(label="New Center Registration", command=self.register_center,
                                   state=NORMAL,
                                   underline=0)
        # -------
        self.stock_transaction_menu.add_command(label="Purchase Book", command=self.purchase_bookForm, state=NORMAL,
                                                underline=0)
        self.stock_transaction_menu.add_command(label="Borrow Book", command=self.borrow_bookForm, state=NORMAL,
                                                underline=0)
        self.stock_transaction_menu.add_command(label="Return Book", command=self.book_ReturnForm, state=NORMAL,
                                                underline=0)
        self.stock_transaction_menu.add_command(label="Return Advance", command=self.advance_returnForm, state=NORMAL,
                                                underline=0)

        editMember_memberId = partial(self.edit_member_data, 1, SEARCH_BY_MEMBERID)
        self.edit_menu.add_command(label="Member Info", command=editMember_memberId, state=NORMAL, underline=0)

        item_detailsmodify_submenu = Menu(self.file_menu, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                          activeforeground='light yellow')
        self.edit_menu.add_cascade(label='Stock Details Info', menu=item_detailsmodify_submenu, state=NORMAL,
                                   underline=0)

        item_detailsmodify_submenu.add_command(label="Commercial", command=self.edit_commercialItem_data,
                                               state=NORMAL,
                                               underline=0)

        item_detailsmodify_submenu.add_command(label="Non-Commercial", command=self.edit_noncommercialItem_data,
                                               state=NORMAL,
                                               underline=0)
        self.edit_menu.add_command(label="Split Donation", command=self.view_split_donation_window, state=NORMAL,
                                   underline=0)

        # Adding Search Menu

        member_search_submenu = Menu(self.view_menu, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                     activeforeground='light yellow')
        self.view_menu.add_cascade(label='Member Info', menu=member_search_submenu, state=NORMAL, underline=0)

        displayMember_memberId = partial(self.display_data, 1, SEARCH_BY_MEMBERID)
        member_search_submenu.add_command(label="By Member Id", command=displayMember_memberId, state=NORMAL,
                                          underline=0)
        displayMember_contactno = partial(self.display_data, 1, SEARCH_BY_CONTACTNO)
        member_search_submenu.add_command(label="By Contact No", command=displayMember_contactno, state=NORMAL,
                                          underline=0)
        member_idcard_result = partial(self.generate_id_card, 1)
        member_search_submenu.add_command(label="ID card", command=member_idcard_result, state=NORMAL,
                                          underline=0)

        info_stock_menu = Menu(self.view_menu, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                               activeforeground='light yellow')
        self.view_menu.add_cascade(label='Inventory/Stock', menu=info_stock_menu, state=NORMAL, underline=0)

        info_stock_menu.add_command(label='Generic Stock Info', command=self.view_stock_info, state=NORMAL,
                                    underline=0)
        info_stock_menu.add_command(label='Stock Sales Statement', command=self.view_stocksales_statement,
                                    state=NORMAL, underline=0)

        info_donation_menu = Menu(self.view_menu, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                  activeforeground='light yellow')
        self.view_menu.add_cascade(label='Donation', menu=info_donation_menu, state=NORMAL, underline=0)

        info_donation_menu.add_command(label='Split Donation List', command=self.view_split_donation_list,
                                       state=NORMAL, underline=0)
        info_donation_menu.add_command(label='Donation Statement', command=self.view_monetarydonation_statement,
                                       state=NORMAL, underline=0)
        info_donation_menu.add_command(label='Member Contribution', command=self.view_memberContribution,
                                       state=NORMAL, underline=0)

        pledge_donation_menu = Menu(self.view_menu, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                    activeforeground='light yellow')
        self.view_menu.add_cascade(label='Pledge Statement', menu=pledge_donation_menu, state=NORMAL, underline=0)

        pledge_donation_menu.add_command(label='Statement By Duration', command=self.view_pledge_statement_by_duration,
                                         state=NORMAL, underline=0)
        pledge_donation_menu.add_command(label='Statement By Pledge Item', command=self.view_pledge_statement_by_item,
                                         state=NORMAL, underline=0)
        pledge_donation_menu.add_command(label='Statement By Member', command=self.view_pledge_statement_by_member,
                                         state=NORMAL, underline=0)

        if userControlPrivilage == "Accountant" or userControlPrivilage == "Manager" or \
                userControlPrivilage == "President" or userControlPrivilage == "Vice-President" \
                or userControlPrivilage == "Admin":
            self.view_menu.add_command(label='Ashram Account Statement', command=self.view_main_account_statement,
                                       state=NORMAL, underline=0)
            self.view_menu.add_command(label='Gaushala Account Statement', command=self.view_gaushala_account_statement,
                                       state=NORMAL, underline=0)

        self.account_menu.add_command(label='Current User', command=self.loggedStaffDetails, state=NORMAL, underline=0)
        self.account_menu.add_command(label='Change Password', command=self.resetLoginPassword_window, state=NORMAL,
                                      underline=0)
        self.account_menu.add_command(label='Logout', command=self.login_window, state=NORMAL, underline=0)
        if userControlPrivilage == "Admin":
            self.account_menu.add_command(label='Import Database', command=self.import_database, state=NORMAL,
                                          underline=0)
            self.account_menu.add_command(label='Reset Database', command=self.resetDatabase, state=NORMAL, underline=0)
            self.account_menu.add_command(label='Re-initialize Database',
                                          command=self.obj_initDatabase.initilizealldatabase, state=NORMAL, underline=0)
        self.account_menu.add_command(label='Create Backup & Exit', command=self.exit_application, state=NORMAL,
                                      underline=0)
        self.account_menu.add_command(label='Exit', command=self.simpleExit, state=NORMAL, underline=0)
        self.master.config(menu=self.main_bar)

        self.master.bind('<Alt-l>', lambda event=None: self.file_menu.invoke())
        self.master.bind('<Alt-r>', lambda event=None: self.edit_menu.invoke())
        self.master.bind('<Alt-v>', lambda event=None: self.option_menu.invoke())
        self.master.bind('<Alt-a>', lambda event=None: self.account_menu.invoke())

    def exit_application(self):
        print("Creating Backup before exit ")
        today = datetime.now()
        backup_folder = today.strftime("%d_%b_%Y_%H%M%S")

        src_folder = "..\\Expanse_Data\\"
        dest_folder = "C:\\VYOAM\\VYOAM_Backup\\" + backup_folder
        self.obj_commonUtil.create_backup(src_folder, dest_folder)
        print("Creating Backup Completed !!! ")
        self.obj_commonUtil.encryptDatabase()
        self.master.destroy()

    def simpleExit(self):
        self.obj_commonUtil.encryptDatabase()
        self.master.destroy()

    def resume_permissions(self):
        print("Opening database transactions")
        today = datetime.now()
        backup_folder = today.strftime("%d_%b_%Y_%H%M%S")

        src_folder = "..\\Expanse_Data\\"
        dest_folder = "C:\\VYOAM\\VYOAM_Backup\\" + backup_folder
        # self.obj_commonUtil.change_permissions_recursive(src_folder,0o777)
        # self.obj_commonUtil.change_permissions_recursive(dest_folder,0o777)
        print("Database is open for transactions ")

    def donothing(self, event=None):
        print("Button is disabled")
        pass

    def clear_loginForm(self, userNameText, passwordText):
        userNameText.delete(0, END)
        userNameText.configure(fg='black')
        userNameText.focus_set()
        passwordText.delete(0, END)
        passwordText.configure(fg='black')

    def clear_ChangePasswordForm(self, userNameText,
                                 old_passwordText, new_passwordText, confirm_passwordText):
        userNameText.delete(0, END)
        userNameText.configure(fg='black')
        userNameText.focus_set()
        old_passwordText.delete(0, END)
        old_passwordText.configure(fg='black')
        new_passwordText.delete(0, END)
        new_passwordText.configure(fg='black')
        confirm_passwordText.delete(0, END)
        confirm_passwordText.configure(fg='black')
        userNameText.focus()

    def loggedStaffDetails(self):
        logged_info_screen = Toplevel(self.master)  # create a GUI window

        # Get the master screen width and height , and place the child screen accordingly
        xSize = self.master.winfo_screenwidth()
        ySize = self.master.winfo_screenheight()

        # set the configuration of GUI window
        logged_info_screen.geometry(
            '{}x{}+{}+{}'.format(400, 150, (int(xSize / 5) + 150), (int(ySize / 5) + 50)))
        logged_info_screen.title("Account Login Info")  # set the title of GUI window
        logged_info_screen.configure(bg="lemonchiffon")
        logged_info_screen.protocol('WM_DELETE_WINDOW', self.donothing)

        upperFrame = Frame(logged_info_screen, width=300, height=200, bd=8, relief='ridge', bg="white")
        upperFrame.grid(row=1, column=0, padx=20, pady=10, columnspan=2)

        label_logged_info = Label(upperFrame, text="", width=30, anchor=CENTER, justify=CENTER,
                                  font=('arial narrow', 18, 'normal'), bg='antiquewhite', fg='midnightblue')
        label_logged_info.grid(row=0, column=0, padx=1, pady=1)
        label_logged_info[
            'text'] = "User Name : " + self.logged_staff_name + "\n" + "Staff Id : " + self.logged_staff_id

        closeFrame = Frame(logged_info_screen, width=200, height=100, bd=4, relief='ridge')
        closeFrame.grid(row=3, column=0, columnspan=2)
        cancel = Button(closeFrame, text="Close", fg="Black", command=logged_info_screen.destroy,
                        font=NORM_FONT, width=10, bg='light cyan')
        cancel.grid(row=0, column=0)
        logged_info_screen.focus()
        logged_info_screen.grab_set()
        mainloop()

    '''
    Primary conditions for password validation :
    ---------------------------------------------------------
    Minimum 8 characters.
    The alphabets must be between [a-z]
    At least one alphabet should be of Upper Case [A-Z]
    At least 1 number or digit between [0-9].
    At least 1 character from [ _ or @ or $ ].
    ---------------------------------------------------------
    '''

    def password_validity(self, password):
        print("password_validity -> start :", password)
        flag = 0
        while True:
            if len(password) < 8:
                flag = -1
                break
            elif not re.search("[a-z]", password):
                flag = -1
                break
            elif not re.search("[A-Z]", password):
                flag = -1
                break
            elif not re.search("[0-9]", password):
                flag = -1
                break
            elif not re.search("[_@$]", password):
                flag = -1
                break
            elif re.search("\s", password):
                flag = -1
                break
            else:
                flag = 0
                break
        return flag

    def change_login_password_Excel(self, login_window, lowerFrame, buttonFrame, userNameText,
                                    old_passwordText, new_passwordText, confirm_passwordText, labelLogin):
        bOldPasswordValid = self.validateStaffPassword_Excel(userNameText.get(), old_passwordText.get())
        if bOldPasswordValid is True:
            wb_obj = openpyxl.load_workbook(PATH_STAFF_CREDENTIALS)
            sheet_obj = wb_obj.active
            total_record = self.totalrecords_excelDataBase(PATH_STAFF_CREDENTIALS)
            new_pwd = new_passwordText.get()
            reconfirm_newPwd = confirm_passwordText.get()
            # new password and confirmation must match to proceed with new password setting
            print("new pwd ", new_pwd, "   reconfirm_pwd:", reconfirm_newPwd)
            if new_pwd == reconfirm_newPwd:
                bValidPassword = self.password_validity(new_pwd)
                print("bValidPassword :", bValidPassword)
                if bValidPassword == 0:
                    line_no = 0
                    for iLoop in range(0, total_record):
                        if str(sheet_obj.cell(row=iLoop + 2, column=2).value) == userNameText.get():
                            sheet_obj.cell(row=iLoop + 2, column=3).value = new_pwd
                            wb_obj.save(PATH_STAFF_CREDENTIALS)

                            # Get the master screen width and height , and place the child screen accordingly
                            xSize = self.master.winfo_screenwidth()
                            ySize = self.master.winfo_screenheight()

                            # set the configuration of GUI window
                            login_window.geometry(
                                '{}x{}+{}+{}'.format(400, 200, (int(xSize / 5) + 150),
                                                     (int(ySize / 5) + 50)))
                            labelLogin.configure(height=4, font=('arial narrow', 18, 'normal'), bg='antiquewhite',
                                                 fg='midnightblue')
                            name = self.getStaffUserName_Excel(userNameText.get(), 2)
                            strtemp = "Password change Success !!" + "\nStaff Id :" + userNameText.get() + "\nUser Name :" + str(
                                name)
                            labelLogin['text'] = strtemp
                            lowerFrame.destroy()
                            buttonFrame.destroy()
                            closeFrame = Frame(login_window, width=200, height=100, bd=4, relief='ridge')
                            closeFrame.grid(row=3, column=0, columnspan=2)
                            cancel = Button(closeFrame, text="Close", fg="Black", command=login_window.destroy,
                                            font=NORM_FONT, width=9, bg='khaki')
                            cancel.grid(row=0, column=0)
                            break
                else:
                    labelLogin.configure(font=('arial narrow', 13, 'normal'), fg='red', bg='light cyan')
                    labelLogin['text'] = "Password set criteria failed!! Try Again"
            else:
                labelLogin.configure(font=('arial narrow', 13, 'normal'), fg='red', bg='light cyan', width=45)
                labelLogin['text'] = "New and Confirm password do not match!! Try Again"
        else:
            labelLogin.configure(font=('arial narrow', 13, 'normal'), fg='red', bg='light cyan', width=45)
            labelLogin['text'] = "Incorrect Old password !! Try Again"

    def resetLoginPassword_window(self):
        login_window = Toplevel(self.master)  # create a GUI window

        # Get the master screen width and height , and place the child screen accordingly
        xSize = self.master.winfo_screenwidth()
        ySize = self.master.winfo_screenheight()

        # set the configuration of GUI window
        login_window.geometry(
            '{}x{}+{}+{}'.format(430, 420, (int(xSize / 4.7)), (int(ySize / 4.8) + 25)))
        login_window.title("Account Login")  # set the title of GUI window
        login_window.configure(bg="white")
        login_window.protocol('WM_DELETE_WINDOW', self.donothing)
        upperFrame = Frame(login_window, width=300, height=200, bd=8, relief='ridge', bg="white")
        upperFrame.grid(row=1, column=0, padx=10, pady=5, columnspan=2)

        labelLogin = Label(upperFrame, text="Change Password", width=32, anchor=CENTER, justify=CENTER,
                           font=('arial narrow', 18, 'normal'), fg='blue', bg='light cyan')
        labelLogin.grid(row=0, column=0, padx=1, pady=1)

        lowerFrame = Frame(login_window, width=300, height=110, bd=8, relief='ridge', bg="white")
        lowerFrame.grid(row=2, column=0, padx=20, pady=5)
        userNameLabel = Label(lowerFrame, text="User Name", width=18, anchor=W, justify=LEFT,
                              font=('arial narrow', 13, 'normal'), bg="white", bd=2, relief='ridge')
        userNameLabel.grid(row=2, column=0)
        userNameText = Entry(lowerFrame, width=22, font=('Yu Gothic', 12, 'normal'), bd=2, relief='ridge',
                             bg='light yellow')
        userNameText.grid(row=2, column=1, padx=5)

        old_passwordLabel = Label(lowerFrame, text="Old Password", width=18, anchor=W, justify=LEFT,
                                  font=('arial narrow', 13, 'normal'), bg="white", bd=2, relief='ridge')
        old_passwordLabel.grid(row=3, column=0, pady=2)
        old_passwordText = Entry(lowerFrame, width=22, show='*', font=('Yu Gothic', 12, 'normal'), bd=2, relief='ridge',
                                 bg='light yellow')
        old_passwordText.grid(row=3, column=1, padx=5, pady=2)

        new_passwordLabel = Label(lowerFrame, text="New Password", width=18, anchor=W, justify=LEFT,
                                  font=('arial narrow', 13, 'normal'), bg="white", bd=2, relief='ridge')
        new_passwordLabel.grid(row=4, column=0, pady=2)
        new_passwordText = Entry(lowerFrame, width=22, show='*', font=('Yu Gothic', 12, 'normal'), bd=2, relief='ridge',
                                 bg='light yellow')
        new_passwordText.grid(row=4, column=1, padx=5, pady=2)

        confirm_passwordLabel = Label(lowerFrame, text="Confirm New Password", width=18, anchor=W, justify=LEFT,
                                      font=('arial narrow', 13, 'normal'), bg="white", bd=2, relief='ridge')
        confirm_passwordLabel.grid(row=5, column=0, pady=2)
        confirm_passwordText = Entry(lowerFrame, width=22, show='*', font=('Yu Gothic', 12, 'normal'), bd=2,
                                     relief='ridge',
                                     bg='light yellow')
        confirm_passwordText.grid(row=5, column=1, padx=5, pady=2)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(login_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=4, column=0)

        reset_pwd_result = partial(self.change_login_password_Excel, login_window, lowerFrame, buttonFrame,
                                   userNameText,
                                   old_passwordText, new_passwordText, confirm_passwordText, labelLogin)
        # create a Login Button and place into the button frame window
        submit = Button(buttonFrame, text="Change", fg="Black", command=reset_pwd_result,
                        font=NORM_FONT, width=8, bg='light cyan')
        submit.grid(row=0, column=0)

        # create a Clear Button and place into the self.newItem_window window
        clear_result = partial(self.clear_ChangePasswordForm, userNameText,
                               old_passwordText, new_passwordText, confirm_passwordText)
        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0)
        clear.grid(row=0, column=1)

        # create a Cancel Button and place into the self.newItem_window window
        # cancel_Result = partial(self.destroyWindow, self.newItem_window)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=login_window.destroy,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0)
        cancel.grid(row=0, column=2)

        # ---------------------------------Button Frame End----------------------------------------
        # ---------------------------------Password Tip Start----------------------------------------
        password_tip_text = "1. Minimum 8 characters.\n" + \
                            "2. The alphabets must be between [a-z].\n" + \
                            "3. At least one alphabet should be of Upper Case [A-Z].\n" + \
                            "4. At least 1 number or digit between [0-9].\n" + \
                            "5. At least 1 character from [ _ or @ or $ ]."
        pwd_tipframe = Frame(login_window, width=300, height=100, bd=4, relief='ridge')
        pwd_tipframe.grid(row=5, column=0, pady=15)
        password_tip_label = Label(pwd_tipframe, text=password_tip_text, width=44, anchor=W, justify=LEFT,
                                   font=('arial narrow', 13, 'normal'), bg="wheat", bd=2, relief='ridge')
        password_tip_label.grid(row=5, column=0, pady=2)

        # shortcut keys for window operations
        login_window.bind('<Return>', lambda event=None: submit.invoke())
        login_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        login_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        login_window.focus()
        login_window.grab_set()
        login_window.mainloop()  # start the GUI

    def login_window(self):
        login_window = Toplevel(self.master, takefocus=True)  # create a GUI window
        # login_window.tk.call('tk', 'scaling', 2.0)
        # Get the master screen width and height , and place the child screen accordingly
        xSize = self.master.winfo_screenwidth()
        ySize = self.master.winfo_screenheight()

        # set the configuration of GUI window
        login_window.geometry(
            '{}x{}+{}+{}'.format(410, 200, (int(xSize / 4.7)), (int(ySize / 4.8) + 50)))
        login_window.title("Account Login")  # set the title of GUI window
        login_window.configure(bg="white")
        login_window.protocol('WM_DELETE_WINDOW', self.donothing)
        upperFrame = Frame(login_window, width=300, height=200, bd=8, relief='ridge', bg="white")
        upperFrame.grid(row=1, column=0, padx=20, pady=5, columnspan=2)

        labelLogin = Label(upperFrame, text="System Authentication", width=30, anchor=CENTER, justify=CENTER,
                           font=('arial narrow', 18, 'normal'), fg='blue', bg='light cyan')
        labelLogin.grid(row=0, column=0, padx=1, pady=1)

        lowerFrame = Frame(login_window, width=300, height=110, bd=8, relief='ridge', bg="white")
        lowerFrame.grid(row=2, column=0, padx=20, pady=5)
        userNameLabel = Label(lowerFrame, text="User Name", width=12, anchor=W, justify=LEFT,
                              font=('arial narrow', 15, 'normal'), bg="white", bd=2, relief='ridge')
        userNameLabel.grid(row=2, column=0)
        userNameText = Entry(lowerFrame, width=22, font=('Yu Gothic', 12, 'normal'), bd=2, relief='ridge',
                             bg='light yellow')
        userNameText.grid(row=2, column=1, padx=5)
        userNameText.focus_set()

        passwordLabel = Label(lowerFrame, text="Password", width=12, anchor=W, justify=LEFT,
                              font=('arial narrow', 15, 'normal'), bg="white", bd=2, relief='ridge')
        passwordLabel.grid(row=3, column=0, pady=2)
        passwordText = Entry(lowerFrame, width=22, show='*', font=('Yu Gothic', 12, 'normal'), bd=2, relief='ridge',
                             bg='light yellow')
        passwordText.grid(row=3, column=1, padx=5, pady=2)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(login_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=4, column=0)

        reset_pwd_result = partial(self.validateStaff_Login_Excel, login_window, lowerFrame, buttonFrame, userNameText,
                                   passwordText, labelLogin)
        # create a Login Button and place into the button frame window
        submit = Button(buttonFrame, text="Login", fg="Black", command=reset_pwd_result,
                        font=NORM_FONT, width=8, bg='light cyan', highlightcolor="snow")
        submit.grid(row=0, column=0)

        # create a Clear Button and place into the self.newItem_window window
        clear_result = partial(self.clear_loginForm, userNameText, passwordText)
        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0, highlightcolor="black")
        clear.grid(row=0, column=1)

        # create a Cancel Button and place into the self.newItem_window window
        # cancel_Result = partial(self.destroyWindow, self.newItem_window)
        # close_result = partial(self.closeFromLogin, self.master)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=self.master.destroy,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0, highlightcolor="black")
        cancel.grid(row=0, column=2)
        # ---------------------------------Button Frame End----------------------------------------

        login_window.bind('<Return>', lambda event=None: submit.invoke())
        login_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        login_window.bind('<Alt-r>', lambda event=None: clear.invoke())
        login_window.focus_set()
        login_window.grab_set()

        login_window.mainloop()  # start the GUI

    # constructor for Library class
    def __init__(self, master):
        self.list_InvoicePrint = []
        self.itemEntryInstance = False
        self.bookReturnInstance = False
        self.bookBorrowInstance = False
        self.newmember_id = 0
        self.logged_staff_name = ""
        self.logged_staff_id = ""
        self.bookBorrow_window = ""
        self.bookReturn_window = ""
        self.newItem_window = ""
        self.master = master
        self.print_button = ""
        self.member_IdPhotoFilePath = ""
        self.member_photoFilePath = ""
        self.returnBook1Name = ""
        self.returnBook2Name = ""
        self.memberId_forIDCard = ""
        self.middleFrame_bookdisplay = ""
        self.bookId_Dict = {}
        self.label_identities = []
        # sets the configuration of main screen
        self.master.title("Vihangam Yoga Regional Center Management ")
        # ========================================================================

        # ========================================================================
        # self.master.attributes('-disabled', True)
        self.obj_commonUtil = CommonUtil()

        # "X" button window becomes not - operational
        self.master.protocol('WM_DELETE_WINDOW', self.donothing)

        self.obj_initDatabase = InitDatabase()
        # self.startloading(self.master)
        self.obj_commonUtil.disableAllLogingPrints()
        self.obj_commonUtil.decryptDatabase()
        self.obj_splitdonation_window = SplitDonation(root)
        # Data base initialization for all .
        # It is ensured that this is executed only once during installation
        self.objStock_info = StockInfo()
        self.obj_initDatabase.initilizealldatabase()

        self.obj_commonUtil.calculateTotalAvailableBalance(VIHANGAM_YOGA_KARNATAKA_TRUST)
        self.obj_commonUtil.calculateTotalAvailableBalance(SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST)
        self.main_menu()

    def loading_animation_end(self):
        pass

    def main_menu(self):
        width, height = pyautogui.size()
        self.master.geometry('{}x{}+{}+{}'.format(width, height, 0, 0))
        # canvas designed to display the library image on main screen
        canvas_width, canvas_height = width, height
        canvas = Canvas(self.master, width=canvas_width, height=canvas_height)
        myimage = ImageTk.PhotoImage(PIL.Image.open("..\\Images\\Logos\\loading2.JPG").resize((width, height)))
        canvas.create_image(0, 0, anchor=NW, image=myimage)
        canvas.pack()

        self.master.lift()
        # prevents the application been closed by alt + F4 etc.
        # self.master.overrideredirect(True)
        self.login_window()
        self.master.mainloop()

    def startloading(self, _master):
        self.loading_window = Toplevel(_master, takefocus=True)
        # canvas designed to display the library image on main screen

        self.loading_window.title("                      VYOAM Loading ...")  # set the title of GUI self.loading_window
        self.loading_window.configure(bg="white")
        self.loading_window.protocol('WM_DELETE_self.loading_window', self.donothing)

        canvas_width, canvas_height = 410, 200
        canvas = Canvas(self.loading_window, width=canvas_width, height=canvas_height)
        myimage = ImageTk.PhotoImage(PIL.Image.open("..\\Images\\Logos\\loading2.JPG").resize((410, 200)))
        canvas.create_image(0, 0, anchor=NW, image=myimage)
        canvas.grid(row=0, column=0)
        obj_threadClass = myThread(1, "loadingvyoam", 3, self.loading_window,
                                   "Dummy", "Dummy", "Dummy", "Dummy", "Dummy",
                                   "Dummy")
        obj_threadClass.start()
        self.loading_window.mainloop()


# obj_animation = LoadingAnimation()
root = Tk()

# Query DPI Awareness (Windows 10 and 8)
awareness = ctypes.c_int()
errorCode = ctypes.windll.shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
print(awareness.value)

# Set DPI Awareness  (Windows 10 and 8)
# errorCode = ctypes.windll.shcore.SetProcessDpiAwareness(2)
# the argument is the awareness level, which can be 0, 1 or 2:
# for 1-to-1 pixel control I seem to need it to be non-zero (I'm using level 2)
dpi = root.winfo_fpixels('1i')
factor = dpi / 72
root.call('tk', 'scaling', factor)

libraryObj = Library(root)
