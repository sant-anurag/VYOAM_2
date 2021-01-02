"""
# Copyright 2020 by Vihangam Yoga Karnataka.
# All rights reserved.
# This file is part of the Vihangan Yoga Operations of Ashram Management Software Package(VYOAM),
# and is released under the "VY License Agreement". Please see the LICENSE
# file that should have been included as part of this package.
# Vihangan Yoga Operations  of Ashram Management Software
# File Name : member_contribution.py
# Developer : Sant Anurag Deo
# Version : 2.0
"""

from app_defines import *
from app_common import *
from init_database import *
from app_thread import *


# Class definition for member donation statement
class MemberContribution:
    # constructor for Member Donation class
    def __init__(self, master):
        print("constructor called for noncommercial edit ")
        # creating handlers of common util classes
        self.obj_commonUtil = CommonUtil()
        self.dateTimeOp = DatetimeOperation()
        self.memberDonation_statement(master)


    def memberDonation_statement(self, master):
        """
        Method to display the Member donation statment window
        :param master:  root object
        :return: None
        """
        # creating the toplevel window as child of master
        transaction_summary_window = Toplevel(master)

        #specifying window attributes
        transaction_summary_window.title("Individual Contribution Statement ")
        transaction_summary_window.geometry('780x560+120+40')
        transaction_summary_window.configure(background='wheat')

        # Window cannot be resized
        transaction_summary_window.resizable(width=False, height=False)

        # delete "X" button in window will be not-operational
        transaction_summary_window.protocol('WM_DELETE_WINDOW', self.obj_commonUtil.donothing)

        # declaring the image frame to display the Organization logo
        imageFrame = Frame(transaction_summary_window, width=65, height=60,
                           bg="wheat")
        canvas_width, canvas_height = 60, 60
        canvas = Canvas(imageFrame, width=canvas_width, height=canvas_height, highlightthickness=0)
        myimage = ImageTk.PhotoImage(PIL.Image.open("..\\Images\\a_wheat.png").resize((60, 60)))
        canvas.create_image(0, 0, anchor=NW, image=myimage)
        imageFrame.grid(row=0, column=0, pady=2)
        canvas.grid(row=0, column=0)

        infoFrame = Frame(transaction_summary_window, width=300, height=100, bd=6, relief='ridge',
                          bg="wheat")
        topFrame = Frame(transaction_summary_window, width=300, height=100, bd=6, relief='ridge',
                         bg="light yellow")
        topFrame.grid(row=2, column=0, padx=20, pady=10, sticky=W)
        upperFrame = Frame(topFrame, width=300, height=100, bg="light yellow")
        upperFrame.grid(row=1, column=0, padx=5, pady=10, sticky=W)
        infoFrame.grid(row=3, column=0, padx=80, pady=10, sticky=W)
        infoLabel = Label(infoFrame, text="Select appropriate parameters and press Search", width=65,
                          justify='center', font=('arial narrow', 13, 'bold'),
                          bg='snow', fg='green', state=NORMAL)
        infoLabel.grid(row=0, column=0)
        middleFrame_bookdisplay = Frame(transaction_summary_window, width=0, height=0, bd=8, relief='ridge')

        heading = Label(transaction_summary_window, text="Member Donation Statement",
                        font=('times new roman', 20, 'bold'),
                        bg="wheat")

        heading.grid(row=1, column=0)
        category_frame = Frame(upperFrame, width=100, height=100, bd=0, relief='ridge', bg='light yellow')
        category_frame.grid(row=1, column=0, padx=10, pady=10)
        memberId_label = Label(category_frame, text="Member ID", width=14, justify='center', font=NORM_FONT,
                               bg='light yellow', state=NORMAL)
        member_IdText = Entry(category_frame, width=20, font=('arial narrow', 15, 'normal'), justify='center',
                              bg='snow', state=NORMAL)
        category_label = Label(category_frame, text="Donation Type", width=12, justify='center', font=NORM_FONT,
                               bg='light yellow', state=NORMAL)
        memberId_label.grid(row=1, column=0, padx=5, pady=10)
        member_IdText.grid(row=1, column=1, padx=5, pady=10)
        category_label.grid(row=1, column=2, padx=5, pady=10)

        categoryText = StringVar(category_frame)
        category_list = ['Monthly Seva', 'Gaushala Seva', 'Hawan Seva', 'Event/Prachar Seva',
                         'Aarti Seva', 'Akshay-Patra Seva', 'Ashram Seva(Generic)', 'Ashram Nirmaan Seva', 'Yoga Fees',
                         'All']
        categoryText.set("All")
        category_menu = OptionMenu(category_frame, categoryText, *category_list)
        category_menu.configure(font=NORM_FONT, width=18, anchor=W, justify='left', bg='snow')
        category_menu.grid(row=1, column=3, padx=10, pady=10)

        dateFrame = Frame(upperFrame, width=100, height=100, bd=2, relief='ridge', bg='light yellow')
        fromDate = Label(dateFrame, text="From Date", width=10, anchor=W, justify='center',
                         font=NORM_FONT,
                         bg='light yellow', state=DISABLED)
        cal_dateFrom = DateEntry(dateFrame, width=15, date_pattern='dd/MM/yyyy', font=NORM_FONT,
                                 state=DISABLED, justify=LEFT, anchor=W)
        toDate = Label(dateFrame, text="To Date", width=10, justify='center', font=NORM_FONT,
                       bg='light yellow', state=DISABLED)
        cal_toDate = DateEntry(dateFrame, width=15, date_pattern='dd/MM/yyyy', font=NORM_FONT,
                               state=DISABLED, justify='center')
        viewByMonth_Month = Label(dateFrame, text="Month", width=10, justify=LEFT, anchor=W,
                                  font=NORM_FONT,
                                  bg='light yellow', state=DISABLED)
        viewByMonth_Year = Label(dateFrame, text="Year", width=10, justify=LEFT, anchor=W,
                                 font=NORM_FONT,
                                 bg='light yellow', state=DISABLED)
        viewByYear_Year = Label(dateFrame, text="Year", width=10, justify=LEFT, anchor=W,
                                font=NORM_FONT,
                                bg='light yellow', state=NORMAL)

        month_variable = StringVar(dateFrame)
        now = datetime.now()
        month_variable.set(self.dateTimeOp.fetchMonthName(now.month))

        viewbyMonth_monthTxt = OptionMenu(dateFrame, month_variable, 'January', 'February', 'March', 'April', 'May',
                                          'June', 'July', 'August', 'September', 'October',
                                          'November', 'December')
        viewbyMonth_monthTxt.configure(bg='snow', width=13, fg='black', font=NORM_FONT,
                                       state=DISABLED)

        year_variable = StringVar(dateFrame)
        year_variable.set("2020")

        viewbymonth_yearTxt = OptionMenu(dateFrame, year_variable, '2019', '2020', '2021', '2022', '2023', '2024',
                                         '2025', '2026', '2027', '2028', '2029', '2030')
        viewbymonth_yearTxt.configure(bg='snow', width=13, fg='black', font=NORM_FONT,
                                      state=DISABLED)

        year_yearvariable = StringVar(dateFrame)
        year_yearvariable.set("2020")
        viewbyYear_yearTxt = OptionMenu(dateFrame, year_yearvariable, '2019', '2020', '2021', '2022', '2023', '2024',
                                        '2025', '2026', '2027', '2028', '2029', '2030')
        viewbyYear_yearTxt.configure(bg='snow', width=13, fg='black', font=NORM_FONT,
                                     state=NORMAL)

        var = IntVar()
        var.set(3)
        viewSelFrame_Result = partial(self.enableViewBy_RadioSelection, var, fromDate, cal_dateFrom, toDate, cal_toDate,
                                      viewByMonth_Month, viewbyMonth_monthTxt, viewByMonth_Year, viewbymonth_yearTxt,
                                      viewByYear_Year, viewbyYear_yearTxt)
        viewbydate_radioBtn = Radiobutton(dateFrame, text="View By Date", variable=var, value=1,
                                          command=viewSelFrame_Result, width=12, bg='light yellow',
                                          font=NORM_FONT, anchor=W, justify=LEFT)
        viewbymonth_radioBtn = Radiobutton(dateFrame, text="View By Month", variable=var, value=2,
                                           command=viewSelFrame_Result, width=12, bg='light yellow',
                                           font=NORM_FONT, anchor=W, justify=LEFT)
        viewbyyear_radioBtn = Radiobutton(dateFrame, text="View By Year", variable=var, value=3,
                                          command=viewSelFrame_Result, width=12, bg='light yellow',
                                          font=NORM_FONT, anchor=W, justify=LEFT)
        viewbydate_radioBtn.grid(row=0, column=0, padx=20)
        fromDate.grid(row=1, column=0, padx=20)
        cal_dateFrom.grid(row=1, column=1)
        toDate.grid(row=1, column=2, padx=10)
        cal_toDate.grid(row=1, column=3)

        viewbymonth_radioBtn.grid(row=2, column=0, padx=30)
        viewByMonth_Month.grid(row=3, column=0, padx=20)
        viewbyMonth_monthTxt.grid(row=3, column=1)
        viewByMonth_Year.grid(row=3, column=2, padx=10)
        viewbymonth_yearTxt.grid(row=3, column=3)

        viewbyyear_radioBtn.grid(row=4, column=0, padx=30)
        viewByYear_Year.grid(row=5, column=1, padx=20)
        viewbyYear_yearTxt.grid(row=5, column=2)
        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(category_frame, width=100, height=100, bd=4, relief='ridge', bg="light yellow")
        buttonFrame.grid(row=2, column=0, padx=10, pady=5, columnspan=4)

        viewPDF = Button(buttonFrame, text="View Statement", fg="Black",
                         font=NORM_FONT, width=12, bg='light grey', state=DISABLED)

        printBtn = Button(buttonFrame, text="Print Statement", fg="Black",
                          font=NORM_FONT, width=12, bg='light grey', state=DISABLED)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=transaction_summary_window.destroy,
                        font=NORM_FONT, width=12, bg='light cyan')
        search_result = partial(self.prepare_account_statement_Excel,
                                transaction_summary_window,
                                member_IdText,
                                printBtn,
                                cal_dateFrom,
                                cal_toDate,
                                month_variable,
                                year_variable,
                                year_yearvariable,
                                var, categoryText,
                                viewPDF,
                                infoLabel,
                                cancel)
        submit = Button(buttonFrame, text="Search", fg="Black", command=search_result,
                        font=NORM_FONT, width=12, bg='light cyan')
        submit.grid(row=0, column=0)
        viewPDF.grid(row=0, column=1)
        printBtn.grid(row=0, column=2)

        dateFrame.grid(row=2, column=0, padx=10, pady=10)
        fromDate.grid(row=1, column=0, padx=10)
        cal_dateFrom.grid(row=1, column=1)
        toDate.grid(row=1, column=2, padx=10)
        cal_toDate.grid(row=1, column=3)

        # create a Close Button and place into the transaction_summary_window window

        cancel.grid(row=0, column=3)
        # ---------------------------------Button Frame End----------------------------------------
        middleFrame_bookdisplay.grid(row=2, column=2, padx=65, pady=10, sticky=W)

        transaction_summary_window.bind('<Return>', lambda event=None: submit.invoke())
        transaction_summary_window.bind('<Alt-c>', lambda event=None: cancel.invoke())

        transaction_summary_window.focus()
        transaction_summary_window.grab_set()
        mainloop()

    def checkCategoryChange(self, n, m, x, src_file, starting_index, dummy):
        """
        Method to clear the template data when category changes
        :param master:  root object
        :return: None
        """
        print("Category has been changed !!!")
        src_filename = n
        wb_template = openpyxl.load_workbook(src_filename)
        template_sheet = wb_template.active
        print("Source file name: ", src_filename)

        for rows in range(15, m + 1):
            for columns in range(1, 6):
                template_sheet.cell(row=rows, column=columns).value = ""

        wb_template.save(src_filename)

    def enableViewBy_RadioSelection(self, var, fromDate, cal_dateFrom, toDate, cal_toDate,
                                    viewByMonth_Month, viewbyMonth_monthTxt, viewByMonth_Year, viewbymonth_yearTxt,
                                    viewByYear_Year, viewbyYear_yearTxt):
        """
        This method controls the enablin/disabling of frame elements in filter duration
        :param var:  Radio button for choosing the filter duration
        :param fromDate:  From date for statement generation
        :param cal_dateFrom:  Calender From Date
        :param toDate: "To" date for statement generation
        :param cal_toDate:   Calender To Date
        :param viewByMonth_Month:  "Month" label in View by Month
        :param viewbyMonth_monthTxt: Month" text field in View by Month
        :param viewByMonth_Year: "Month" label in View by Year
        :param viewbymonth_yearTxt:  "Year" text field in View by Month
        :param viewByYear_Year: "Year" label field in View by Month
        :param viewbyYear_yearTxt "Year" text field in View by Month
        :return:
        """
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

    def myfunction(self, mycanvas, event):
        """
        This method specifies the canvas parameters for generated statement display
        :param mycanvas: canvas object from caller
        :param event: scrolling event
        :return: None
        """
        mycanvas.configure(scrollregion=mycanvas.bbox("all"), width=725, height=407)

    def closepage(self, src_file, starting_index, account_statement_window):
        """
        This method closes the donation statment window, and clears the template records for next usage
        :param src_file: template file path
        :param starting_index: staring row number to start the erasing of records
        :param account_statement_window: donation widow object refernece
        :return: None
        """
        print("closepage :", src_file)
        account_statement_window.destroy()
        # erase the written records in template sheet
        # executed only if
        wb_template = openpyxl.load_workbook(src_file)
        template_sheet = wb_template.active

        for rows in range(15, starting_index + 1):
            for columns in range(1, 6):
                template_sheet.cell(row=rows, column=columns).value = ""

        wb_template.save(src_file)

    def prepare_account_statement_Excel(self, account_statement_window,
                                        member_IdText,
                                        printBtn,
                                        cal_dateFrom,
                                        cal_toDate,
                                        viewbyMonth_monthTxt,
                                        viewbymonth_yearTxt,
                                        viewbyYear_yearTxt, var,
                                        categoryText,
                                        viewPDF,
                                        infoLabel,
                                        cancelbtn):
        """
        This method prepares the donation statement for a specific member id
        :param account_statement_window:  donation window object refernece
        :param member_IdText:  Memebr id , to search the recctive records
        :param printBtn: print button reference of the main statement window
        :param cal_dateFrom: From date from calender
        :param cal_toDate: To date from calender
        :param viewbyMonth_monthTxt:
        :param viewbymonth_yearTxt:
        :param viewbyYear_yearTxt:
        :param var: radiobutton choice
        :param categoryText: Donation Category Text
        :param viewPDF: View PDF button reference of the main statement window
        :param infoLabel: Information label object refernce from statement window
        :param cancelbtn: Close button reference of the main statement window
        :return: None
        """
        print("prepare_account_statement_Excel --> var:", var.get())
        bMemberIdValid = self.obj_commonUtil.validate_memberId_Excel(member_IdText.get(), 1)
        if bMemberIdValid:
            viewPDF.configure(state=DISABLED, bg="light grey")
            printBtn.configure(state=DISABLED, bg="light grey")
            if var.get() == 1:
                dateTimeObj_From = cal_dateFrom.get_date()
                from_Date = dateTimeObj_From.strftime("%Y-%m-%d")
                dateTimeObj_To = cal_toDate.get_date()
                to_Date = dateTimeObj_To.strftime("%Y-%m-%d")
                fromDate = self.dateTimeOp.prepare_dateFromString(from_Date)
                toDate = self.dateTimeOp.prepare_dateFromString(to_Date)
            elif var.get() == 2:
                noOfDays, month_number = self.dateTimeOp.calculateNoOfDaysInMonth(viewbyMonth_monthTxt.get(),
                                                                                  viewbymonth_yearTxt.get())
                fromDate, toDate = self.dateTimeOp.getFromAndToDates_Account_Statement(month_number,
                                                                                       viewbymonth_yearTxt.get(),
                                                                                       noOfDays)
                frdateforstatement = fromDate.strftime("%b-%d-%Y")
                todateforstatement = toDate.strftime("%b-%d-%Y")
            else:
                print("Requested year is :", viewbyYear_yearTxt.get())
                noOfDays = self.dateTimeOp.calculateNoOfDaysInYear(viewbyYear_yearTxt.get())
                fromDate, toDate = self.dateTimeOp.getFromAndToDates_Account_Statement(1, viewbyYear_yearTxt.get(),
                                                                                       noOfDays)

            from_year = fromDate.strftime("%Y")
            to_year = toDate.strftime("%Y")
            to_month = toDate.strftime('%m')
            today_date = datetime.now()
            formatted_date = today_date.strftime("%Y-%m-%d")
            currentDate = self.dateTimeOp.prepare_dateFromString(formatted_date)
            print("From Year :", from_year, " To Year :", to_year)
            current_year = currentDate.strftime("%Y")
            print("From Year :", from_year, " To Year :", to_year)
            bDateConditionsValid = True
            if fromDate > toDate:
                error_info = "From date cannot be grater than to date !!!"
                bDateConditionsValid = False
            elif ((toDate > currentDate) or (fromDate > currentDate)) and \
                    ((var.get() == 1) or (var.get() == 2)):

                if var.get() == 2:
                    current_month = datetime.today().month
                    current_year = datetime.today().year
                    print("to_year : ", to_year, "current_year :", current_year, "to_month :", to_month,
                          "current_month :",
                          current_month)
                    if int(to_year) > int(current_year) or int(to_month) > int(current_month):
                        error_info = "Year/Month can not be future !!!"
                        bDateConditionsValid = False
                    if int(to_year) == int(current_year) and int(to_month) == int(current_month):
                        bDateConditionsValid = True
                else:
                    error_info = "From/To Date cannot be greater than today!!!"
                    bDateConditionsValid = False
            elif ((toDate - fromDate).days > 180) and (var.get() == 1):
                error_info = "Statements can be generated for maximum of 180 days !!! "
                bDateConditionsValid = False
            else:
                # check if the selected years are less than or equal to current date ,
                # but no database exists for them
                dir_name = "..\\Expanse_Data\\" + str(from_year)
                if not os.path.exists(dir_name):
                    error_info = "Database does not exists for " + str(from_year) + " Please correct !!!"
                    bDateConditionsValid = False
                pass

            # algorithm generates the statement when from and start date has the same year
            # same current year directory needs to be referred for these statement generations
            print("bDateConditionsValid :", bDateConditionsValid)
            if bDateConditionsValid:
                if from_year == self.obj_commonUtil.getCurrentYearFolderName() and \
                        to_year == self.obj_commonUtil.getCurrentYearFolderName():
                    bAlike_seva = True
                    print("This is current year transaction")
                    if categoryText.get() == "Monthly Seva":
                        path_seva_sheet = InitDatabase.getInstance().get_monthly_seva_database_name()
                        InitDatabase.getInstance().initilize_sortedmonthly_seva_database()
                        sorted_seva_sheet = InitDatabase.getInstance().get_sortedmonthly_database_name()
                    elif categoryText.get() == "Gaushala Seva":
                        path_seva_sheet = InitDatabase.getInstance().get_gaushala_seva_database_name()
                        InitDatabase.getInstance().initilize_sortedgaushala_seva_database()
                        sorted_seva_sheet = InitDatabase.getInstance().get_sortedgaushala_database_name()
                    elif categoryText.get() == "Hawan Seva":
                        path_seva_sheet = InitDatabase.getInstance().get_hawan_seva_database_name()
                        InitDatabase.getInstance().initilize_sortedhawan_seva_database()
                        sorted_seva_sheet = InitDatabase.getInstance().get_sortedhawan_database_name()
                    elif categoryText.get() == "Event/Prachar Seva":
                        path_seva_sheet = InitDatabase.getInstance().get_prachar_event_seva_database_name()
                        InitDatabase.getInstance().initilize_sortedprachar_event_seva_database()
                        sorted_seva_sheet = InitDatabase.getInstance().get_sortedprachar_event_database_name()
                    elif categoryText.get() == "Aarti Seva":
                        path_seva_sheet = InitDatabase.getInstance().get_aarti_seva_database_name()
                        InitDatabase.getInstance().initilize_sortedaarti_seva_database()
                        sorted_seva_sheet = InitDatabase.getInstance().get_sortedaarti_seva_database_name()
                    elif categoryText.get() == "Ashram Seva(Generic)":
                        path_seva_sheet = InitDatabase.getInstance().get_ashram_seva_database_name()
                        InitDatabase.getInstance().initilize_sortedashram_seva_database()
                        sorted_seva_sheet = InitDatabase.getInstance().get_sortedashram_seva_database_name()
                    elif categoryText.get() == "Ashram Nirmaan Seva":
                        path_seva_sheet = InitDatabase.getInstance().get_ashram_nirmaan_seva_database_name()
                        InitDatabase.getInstance().initilize_sortedashramnirmaan_seva_database()
                        sorted_seva_sheet = InitDatabase.getInstance().get_sortedashramnirmaan_seva_database_name()
                    elif categoryText.get() == "Yoga Fees":
                        path_seva_sheet = InitDatabase.getInstance().get_yoga_seva_database_name()
                        InitDatabase.getInstance().initilize_sortedyoga_seva_database()
                        sorted_seva_sheet = InitDatabase.getInstance().get_sortedyoga_seva_database_name()
                    elif categoryText.get() == "All":
                        path_seva_sheet = InitDatabase.getInstance().get_seva_deposit_database_name()
                        InitDatabase.getInstance().initilize_sortedseva_deposit_database()
                        sorted_seva_sheet = InitDatabase.getInstance().get_sortedseva_deposit_database_name()
                    elif categoryText.get() == "Akshay-Patra Seva":
                        path_seva_sheet = InitDatabase.getInstance().get_akshay_patra_database_name()
                        InitDatabase.getInstance().initilize_sortedakshay_patra_database()
                        sorted_seva_sheet = InitDatabase.getInstance().get_sortedakshay_patra_database_name()
                    else:
                        pass

                    if bAlike_seva:
                        print("path_seva_sheet :", path_seva_sheet)
                        wb_obj = openpyxl.load_workbook(path_seva_sheet)
                        wb_sorted_obj = openpyxl.load_workbook(sorted_seva_sheet)
                        sheet_obj = wb_obj.active
                        sheet_sorted_obj = wb_sorted_obj.active
                        total_records = self.obj_commonUtil.totalrecords_excelDataBase(path_seva_sheet)
                        if total_records > 0:
                            sort_sheet_index = 2
                            print("Total records  in transaction sheet:", total_records)
                            for row_index in range(0, total_records):
                                # critical stock ->stock with quantity is 0 or 1
                                # print("Date from sheet is :", sheet_obj.cell(row=row_index + 2, column=6).value)
                                dateFromTransactionSheet = self.dateTimeOp.prepare_dateFromString(
                                    sheet_obj.cell(row=row_index + 2, column=6).value)
                                memberIdFromSheet = str(sheet_obj.cell(row=row_index + 2, column=4).value)
                                # print("dateFromMon_DepositSheet :", dateFromTransactionSheet, "fromDate :", fromDate, " toDate:", toDate)

                                # if date conditions are satisfied , and member id matches with the one uin the monhly sheet
                                if ((dateFromTransactionSheet > fromDate or dateFromTransactionSheet == fromDate)
                                        and (dateFromTransactionSheet < toDate or dateFromTransactionSheet == toDate)
                                        and (member_IdText.get() == memberIdFromSheet)):

                                    for column_index in range(1, 15):
                                        text_value = str(sheet_obj.cell(row=row_index + 2, column=column_index).value)

                                        sheet_sorted_obj.cell(row=sort_sheet_index, column=column_index).font = Font(
                                            size=8,
                                            name='Arial',
                                            bold=False)
                                        sheet_sorted_obj.cell(row=sort_sheet_index,
                                                              column=column_index).alignment = Alignment(
                                            horizontal='left', vertical='center', wrapText=True)
                                        sheet_sorted_obj.cell(row=sort_sheet_index,
                                                              column=column_index).value = text_value

                                    sort_sheet_index = sort_sheet_index + 1

                            today = date.today()
                            dt_today = today.strftime("%d-%b-%Y")
                            wb_sorted_obj.save(sorted_seva_sheet)
                            print("Sorted sheet created for sorting")
                            self.obj_commonUtil.sortExcelSheetByDate(sorted_seva_sheet, sorted_seva_sheet)

                            now = datetime.now()
                            dt_string = now.strftime("%d_%b_%Y_%H%M%S")
                            currentyear = now.strftime("%Y")
                            destination_file = "..\\Expanse_Data\\" + currentyear + "\\Seva_Rashi\\Statements\\" + member_IdText.get() + "_Donation" + dt_string + ".pdf"
                            # write the  sorted record in statement template
                            template_statement = "..\\Expanse_Data\\" + currentyear + "\\Seva_Rashi\\Template\\member_statement_template.xlsx"
                            wb_critical_stock = openpyxl.load_workbook(template_statement)
                            critical_stock_sheet = wb_critical_stock.active
                            total_sorted_records = self.obj_commonUtil.totalrecords_excelDataBase(sorted_seva_sheet)
                            wb_sort = openpyxl.load_workbook(sorted_seva_sheet)
                            sort_sheet = wb_sort.active

                            dict_index = 1
                            starting_index = 15
                            print("Total sorted records :", total_sorted_records)
                            if total_sorted_records > 0:
                                text_info = "Statement is being generated for " + categoryText.get() + ".Please wait ...."
                                infoLabel.configure(text=text_info, fg='purple')
                                for row_index in range(1, total_sorted_records + 1):
                                    print("credit amount is :", sort_sheet.cell(row=row_index + 1, column=2).value)
                                    if dict_index == 1:
                                        balance = int(sort_sheet.cell(row=row_index + 1, column=2).value)
                                    else:
                                        balance = balance + int(sort_sheet.cell(row=row_index + 1, column=2).value)

                                    for column_index in range(1, 6):
                                        if column_index == 1:  # Date
                                            text_value = sort_sheet.cell(row=row_index + 1, column=6).value
                                            text_value = text_value.strftime("%d-%b-%Y")
                                        elif column_index == 2:  # Invoice
                                            text_value = sort_sheet.cell(row=row_index + 1, column=14).value
                                        elif column_index == 3:  # Description
                                            text_value = sort_sheet.cell(row=row_index + 1, column=7).value + "-By " + \
                                                         sort_sheet.cell(row=row_index + 1, column=11).value
                                        elif column_index == 4:  # credit
                                            text_value = sort_sheet.cell(row=row_index + 1, column=2).value
                                        elif column_index == 5:  # balance
                                            text_value = str(balance)
                                        else:
                                            pass
                                        critical_stock_sheet.cell(row=starting_index, column=column_index).font = Font(
                                            size=8,
                                            name='Arial',
                                            bold=False)
                                        if column_index == 4 or column_index == 5:
                                            critical_stock_sheet.cell(row=starting_index,
                                                                      column=column_index).alignment = Alignment(
                                                horizontal='center', vertical='center', wrapText=True)
                                        else:
                                            critical_stock_sheet.cell(row=starting_index,
                                                                      column=column_index).alignment = Alignment(
                                                horizontal='left', vertical='center', wrapText=True)
                                        critical_stock_sheet.cell(row=starting_index,
                                                                  column=column_index).value = text_value
                                    dict_index = dict_index + 1
                                    starting_index = starting_index + 1

                                frdateforstatement = fromDate.strftime("%d-%b-%Y")
                                todateforstatement = toDate.strftime("%d-%b-%Y")
                                critical_stock_sheet.cell(row=3, column=5).value = dt_today
                                critical_stock_sheet.cell(row=5, column=5).value = categoryText.get()
                                critical_stock_sheet.cell(row=8, column=5).value = str(balance)
                                critical_stock_sheet.cell(row=9, column=5).value = str(starting_index - 15)
                                critical_stock_sheet.cell(row=10, column=5).value = frdateforstatement
                                critical_stock_sheet.cell(row=11, column=5).value = todateforstatement
                                wb_critical_stock.save(template_statement)
                                print("File has been saved for template")
                                destination_copy_folder = InitDatabase.getInstance().get_desktop_statement_directory_path()
                                obj_threadClass = myThread(14, "memberDonation", 1, template_statement,
                                                           destination_file, starting_index, viewPDF, printBtn,
                                                           infoLabel,
                                                           destination_copy_folder)
                                obj_threadClass.start()

                                print_result = partial(self.obj_commonUtil.print_statement_file,
                                                       template_statement,
                                                       destination_file, starting_index)
                                printBtn.configure(command=print_result)
                                view_result = partial(self.obj_commonUtil.open_statement_file, template_statement,
                                                      destination_file, starting_index)
                                viewPDF.configure(command=view_result)

                                cancel_result = partial(self.closepage, template_statement, starting_index,
                                                        account_statement_window)
                                cancelbtn.configure(command=cancel_result)
                                result_categoryChangeState = partial(self.checkCategoryChange, template_statement,
                                                                     starting_index,
                                                                     "None")
                                categoryText.trace("w", result_categoryChangeState)
                            else:
                                text_error = "No records present for " + categoryText.get() + " in specified period!!!"
                                infoLabel.configure(text=text_error, fg='red')
                        else:
                            text_error = "No records present for " + categoryText.get()
                            infoLabel.configure(text=text_error, fg='red')
                else:
                    # algorithm generates the statement when from and start date has different year
                    # same current year directory needs to be referred for these statement generations
                    print(" From year and to year are different")

                    # Since the maximum viewed transaction are only 6 months
                    # only 2 year numbers can be considered at max
                    # hence same loop with different from and to dates in executed twice
                    # this is possible only in case of view by date

                    yearDiff = int(to_year) - int(from_year)
                    loop_range = yearDiff + 2

                    print("Year diff :", yearDiff, " loop_range :", loop_range)
                    for year_loop in range(1, loop_range):
                        if var.get() == 1:
                            if year_loop == 1:
                                fDate = fromDate
                                yearOfFromdate = fDate.strftime("%Y")
                                yearFolderToSearch = yearOfFromdate
                                tDate = self.obj_commonUtil.prepare_dateFromString(
                                    "31" + "-" + "12" + "-" + yearOfFromdate)
                            elif year_loop == 2:
                                yearOfTodate = toDate.strftime("%Y")
                                fDate = self.obj_commonUtil.prepare_dateFromString("1" + "-" + "1" + "-" + yearOfTodate)
                                yearOfTodate = toDate.strftime("%Y")
                                yearFolderToSearch = yearOfTodate
                                tDate = toDate
                            else:
                                pass
                        elif var.get() == 2 or var.get() == 3:
                            fDate = fromDate
                            tDate = toDate
                            yearOfFromdate = fDate.strftime("%Y")
                            yearFolderToSearch = yearOfFromdate
                        else:
                            pass

                        print("fDate :", fDate, "tDate:", tDate, "yearFolderToSearch :", yearFolderToSearch)

                        if categoryText.get() == "Monthly Seva":
                            path_seva_sheet = "..\\Expanse_Data\\" + yearFolderToSearch + "\\Seva_Rashi\\Donation\\Monthly_Donation.xlsx"
                            if year_loop == 1:
                                InitDatabase.getInstance().initilize_sortedmonthly_seva_database()
                                sorted_seva_sheet = InitDatabase.getInstance().get_sortedmonthly_database_name()
                        elif categoryText.get() == "Gaushala Seva":
                            path_seva_sheet = "..\\Expanse_Data\\" + yearFolderToSearch + "\\Seva_Rashi\\Donation\\Gaushala_Donation.xlsx"
                            if year_loop == 1:
                                InitDatabase.getInstance().initilize_sortedgaushala_seva_database()
                                sorted_seva_sheet = InitDatabase.getInstance().get_sortedgaushala_database_name()
                        elif categoryText.get() == "Hawan Seva":
                            path_seva_sheet = "..\\Expanse_Data\\" + yearFolderToSearch + "\\Seva_Rashi\\Donation\\Hawan_Donation.xlsx"
                            if year_loop == 1:
                                InitDatabase.getInstance().initilize_sortedhawan_seva_database()
                                sorted_seva_sheet = InitDatabase.getInstance().get_sortedhawan_database_name()
                        elif categoryText.get() == "Event/Prachar Seva":
                            path_seva_sheet = "..\\Expanse_Data\\" + yearFolderToSearch + "\\Seva_Rashi\\Donation\\Event_prachar_Donation.xlsx"
                            if year_loop == 1:
                                InitDatabase.getInstance().initilize_sortedprachar_event_seva_database()
                                sorted_seva_sheet = InitDatabase.getInstance().get_sortedprachar_event_database_name()
                        elif categoryText.get() == "Aarti Seva":
                            path_seva_sheet = "..\\Expanse_Data\\" + yearFolderToSearch + "\\Seva_Rashi\\Donation\\Aarti_Donation.xlsx"
                            if year_loop == 1:
                                InitDatabase.getInstance().initilize_sortedaarti_seva_database()
                                sorted_seva_sheet = InitDatabase.getInstance().get_sortedaarti_seva_database_name()
                        elif categoryText.get() == "Ashram Seva(Generic)":
                            path_seva_sheet = "..\\Expanse_Data\\" + yearFolderToSearch + "\\Seva_Rashi\\Donation\\Ashram_Generic_Donation.xlsx"
                            if year_loop == 1:
                                InitDatabase.getInstance().initilize_sortedashram_seva_database()
                                sorted_seva_sheet = InitDatabase.getInstance().get_sortedashram_seva_database_name()
                        elif categoryText.get() == "Ashram Nirmaan Seva":
                            path_seva_sheet = "..\\Expanse_Data\\" + yearFolderToSearch + "\\Seva_Rashi\\Donation\\Ashram_Nirmaan_Donation.xlsx"
                            if year_loop == 1:
                                InitDatabase.getInstance().initilize_sortedashramnirmaan_seva_database()
                                sorted_seva_sheet = InitDatabase.getInstance().get_sortedashramnirmaan_seva_database_name()
                        elif categoryText.get() == "Yoga Fees":
                            path_seva_sheet = "..\\Expanse_Data\\" + yearFolderToSearch + "\\Seva_Rashi\\Donation\\Yoga_Fees_Donation.xlsx"
                            if year_loop == 1:
                                InitDatabase.getInstance().initilize_sortedyoga_seva_database()
                                sorted_seva_sheet = InitDatabase.getInstance().get_sortedyoga_seva_database_name()
                        elif categoryText.get() == "All":
                            path_seva_sheet = "..\\Expanse_Data\\" + yearFolderToSearch + "\\Seva_Rashi\\Donation\\Monetary_Donation.xlsx"
                            if year_loop == 1:
                                InitDatabase.getInstance().initilize_sortedseva_deposit_database()
                                sorted_seva_sheet = InitDatabase.getInstance().get_sortedseva_deposit_database_name()
                        elif categoryText.get() == "Akshay-Patra Seva":
                            path_seva_sheet = "..\\Expanse_Data\\" + yearFolderToSearch + "\\Seva_Rashi\\Donation\\Akshay_patra.xlsx"
                            if year_loop == 1:
                                InitDatabase.getInstance().initilize_sortedakshay_patra_database()
                                sorted_seva_sheet = InitDatabase.getInstance().get_sortedakshay_patra_database_name()
                        else:
                            pass

                        print("path_seva_sheet :", path_seva_sheet)
                        wb_obj = openpyxl.load_workbook(path_seva_sheet)
                        wb_sorted_obj = openpyxl.load_workbook(sorted_seva_sheet)
                        sheet_obj = wb_obj.active
                        sheet_sorted_obj = wb_sorted_obj.active
                        total_records = self.obj_commonUtil.totalrecords_excelDataBase(path_seva_sheet)
                        if total_records > 0:
                            sort_sheet_index = self.obj_commonUtil.totalrecords_excelDataBase(sorted_seva_sheet) + 2
                            print("Sorted sheet will now start from row:", sort_sheet_index)
                            print("Total records  in transaction sheet:", total_records)
                            for row_index in range(0, total_records):
                                # critical stock ->stock with quantity is 0 or 1
                                # print("Date from sheet is :", sheet_obj.cell(row=row_index + 2, column=6).value)
                                dateFromTransactionSheet = self.dateTimeOp.prepare_dateFromString(
                                    sheet_obj.cell(row=row_index + 2, column=6).value)
                                memberIdFromSheet = str(sheet_obj.cell(row=row_index + 2, column=4).value)
                                # print("dateFromMon_DepositSheet :", dateFromTransactionSheet, "fromDate :", fromDate, " toDate:", toDate)

                                if ((dateFromTransactionSheet > fDate or dateFromTransactionSheet == fDate)
                                        and (dateFromTransactionSheet < tDate or dateFromTransactionSheet == tDate)
                                        and (member_IdText.get() == memberIdFromSheet)):
                                    for column_index in range(1, 15):
                                        text_value = str(sheet_obj.cell(row=row_index + 2, column=column_index).value)

                                        sheet_sorted_obj.cell(row=sort_sheet_index, column=column_index).font = Font(
                                            size=8,
                                            name='Arial',
                                            bold=False)
                                        sheet_sorted_obj.cell(row=sort_sheet_index,
                                                              column=column_index).alignment = Alignment(
                                            horizontal='left', vertical='center', wrapText=True)
                                        sheet_sorted_obj.cell(row=sort_sheet_index,
                                                              column=column_index).value = text_value

                                    sort_sheet_index = sort_sheet_index + 1

                            today = date.today()
                            dt_today = today.strftime("%d-%b-%Y")
                            wb_sorted_obj.save(sorted_seva_sheet)
                            print("Sorted sheet created for sorting")
                            self.obj_commonUtil.sortExcelSheetByDate(sorted_seva_sheet, sorted_seva_sheet)
                            now = datetime.now()
                            dt_string = now.strftime("%d_%b_%Y_%H%M%S")
                            currentyear = now.strftime("%Y")
                            destination_file = "..\\Expanse_Data\\" + currentyear + "\\Seva_Rashi\\Statements\\" + member_IdText.get() + "_Donation" + dt_string + ".pdf"
                            # write the  sorted record in statement template
                            template_statement = "..\\Expanse_Data\\" + currentyear + "\\Seva_Rashi\\Template\\member_statement_template.xlsx"
                            wb_critical_stock = openpyxl.load_workbook(template_sheet)
                            critical_stock_sheet = wb_critical_stock.active
                            total_sorted_records = self.obj_commonUtil.totalrecords_excelDataBase(sorted_seva_sheet)
                            wb_sort = openpyxl.load_workbook(sorted_seva_sheet)
                            sort_sheet = wb_sort.active
                            dict_index = 1
                            starting_index = 15
                            print("Total sorted records :", total_sorted_records)
                            if total_sorted_records > 0:
                                text_info = "Statement is being generated for " + categoryText.get() + ".Please wait ...."
                                infoLabel.configure(text=text_info, fg='purple')
                                for row_index in range(1, total_sorted_records + 1):
                                    if dict_index == 1:
                                        balance = int(sort_sheet.cell(row=row_index + 1, column=2).value)
                                    else:
                                        balance = balance + int(sort_sheet.cell(row=row_index + 1, column=2).value)

                                    for column_index in range(1, 6):
                                        if column_index == 1:  # Date
                                            text_value = sort_sheet.cell(row=row_index + 1, column=6).value
                                            text_value = text_value.strftime("%d-%b-%Y")
                                        elif column_index == 2:  # Invoice
                                            text_value = sort_sheet.cell(row=row_index + 1, column=14).value
                                        elif column_index == 3:  # Description
                                            text_value = sort_sheet.cell(row=row_index + 1, column=7).value + "-By " + \
                                                         sort_sheet.cell(row=row_index + 1, column=11).value

                                        elif column_index == 4:  # credit
                                            text_value = sort_sheet.cell(row=row_index + 1, column=2).value
                                        elif column_index == 5:  # balance
                                            text_value = str(balance)
                                        else:
                                            pass
                                        # print("Text value :", text_value)
                                        critical_stock_sheet.cell(row=starting_index, column=column_index).font = Font(
                                            size=8,
                                            name='Arial',
                                            bold=False)
                                        if column_index == 4 or column_index == 5:
                                            critical_stock_sheet.cell(row=starting_index,
                                                                      column=column_index).alignment = Alignment(
                                                horizontal='center', vertical='center', wrapText=True)
                                        else:
                                            critical_stock_sheet.cell(row=starting_index,
                                                                      column=column_index).alignment = Alignment(
                                                horizontal='left', vertical='center', wrapText=True)
                                        critical_stock_sheet.cell(row=starting_index,
                                                                  column=column_index).value = text_value
                                    dict_index = dict_index + 1
                                    starting_index = starting_index + 1

                                frdateforstatement = fromDate.strftime("%d-%b-%Y")
                                todateforstatement = toDate.strftime("%d-%b-%Y")
                                critical_stock_sheet.cell(row=3, column=5).value = dt_today
                                critical_stock_sheet.cell(row=5, column=5).value = categoryText.get()
                                critical_stock_sheet.cell(row=8, column=5).value = str(balance)
                                critical_stock_sheet.cell(row=9, column=5).value = str(starting_index - 15)
                                critical_stock_sheet.cell(row=10, column=5).value = frdateforstatement
                                critical_stock_sheet.cell(row=11, column=5).value = todateforstatement
                                wb_critical_stock.save(template_sheet)
                                print("File has been saved for template")
                                destination_copy_folder = InitDatabase.getInstance().get_desktop_statement_directory_path()
                                obj_threadClass = myThread(14, "memberDonation", 1, template_sheet,
                                                           destination_file, starting_index, viewPDF, printBtn,
                                                           infoLabel,
                                                           destination_copy_folder)
                                obj_threadClass.start()

                                print_result = partial(self.obj_commonUtil.open_statement_file,
                                                       template_sheet,
                                                       destination_file, starting_index)
                                printBtn.configure(command=print_result)
                                view_result = partial(self.obj_commonUtil.open_statement_file, template_sheet,
                                                      destination_file, starting_index)
                                viewPDF.configure(command=view_result, )

                                cancel_result = partial(self.closepage, template_sheet, starting_index,
                                                        account_statement_window)
                                cancelbtn.configure(command=cancel_result)
                                result_categoryChangeState = partial(self.checkCategoryChange, template_sheet,
                                                                     starting_index,
                                                                     "None")
                                categoryText.trace("w", result_categoryChangeState)
                            else:
                                text_error = "No records present for " + categoryText.get() + " in specified period!!!"
                                infoLabel.configure(text=text_error, fg='red')
                        else:
                            text_error = "No records present for " + categoryText.get() + " in specified period!!!"
                            infoLabel.configure(text=text_error, fg='red')
            else:
                infoLabel.configure(text=error_info, fg='red')
        else:
            error_info = "Member id is invalid  !!!"
            infoLabel.configure(text=error_info, fg='red')
