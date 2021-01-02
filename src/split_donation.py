from app_defines import *
from app_common import *
from init_database import *
from app_thread import *


class SplitDonation:
    # constructor for Library class
    def __init__(self, master):
        print("constructor called for noncommercial edit ")
        self.obj_commonUtil = CommonUtil()
        self.dateTimeOp = DatetimeOperation()
        self.splitTransactionList = []
        self.list_split_open_record = []

    def split_donation_view(self, master):
        split_donation_window = Toplevel(master)
        split_donation_window.title("Split Donation Window ")
        split_donation_window.geometry('1280x500+0+200')
        split_donation_window.configure(background='wheat')
        split_donation_window.resizable(width=True, height=False)
        # delete "X" button in window will be not-operational
        split_donation_window.protocol('WM_DELETE_WINDOW', self.obj_commonUtil.donothing)
        heading = Label(split_donation_window, text="Split Donation Form",
                        font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        heading.grid(row=0, column=0)
        heading_list = Label(split_donation_window, text="Open Split Candidates",
                             font=('ariel narrow', 15, 'bold'),
                             bg='wheat')
        heading_list.grid(row=0, column=5, padx=170)
        # upper frame start
        itemFrame = Frame(split_donation_window, width=210, height=100, bd=4, relief='ridge', bg='snow')
        # splitCandidateListFrame = Frame(split_donation_window, width=210, height=100, bd=4, relief='ridge', bg='snow')
        item_detailsFrame = Frame(split_donation_window, width=210, height=100, bd=4, relief='ridge', bg='snow')
        upperFrame = Frame(split_donation_window, width=210, height=100, bd=4, relief='ridge', bg='snow')

        itemFrame.grid(row=1, column=0, padx=20, pady=5)
        # splitCandidateListFrame.grid(row=1, column=1, padx=20, pady=5)
        item_detailsFrame.grid(row=2, column=0, padx=20, pady=5)
        upperFrame.grid(row=3, column=0, padx=20, pady=5)

        default_text1 = StringVar(itemFrame, value='')

        # design item frame

        itembtnFrame = Frame(itemFrame, width=210, height=100, bd=4, relief='ridge', bg='snow')

        search_btn = Button(itembtnFrame, text="Search", fg="Black",
                            font=NORM_FONT, width=12, bg='light grey', state=DISABLED)
        close_btn = Button(itembtnFrame, text="Close", fg="Black",
                           font=NORM_FONT, width=12, bg='light cyan', state=NORMAL)
        receipt_id = Label(itemFrame, text="Receipt Id", width=12, anchor=W, justify=LEFT,
                           font=('arial', 13, 'normal'),
                           bg='snow')
        receipt_id.grid(row=0, column=0, pady=5)
        receiptid_Text = Entry(itemFrame, text="", width=29, justify=CENTER,
                               font=('arial narrow', 15, 'normal'),
                               bg='light yellow', textvariable=default_text1)
        receiptid_Text.grid(row=0, column=1, pady=5)
        itembtnFrame.grid(row=0, column=2, padx=5)
        search_btn.grid(row=0, column=0, padx=1)
        close_btn.grid(row=0, column=1, padx=1)
        searchinfo_label = Label(itemFrame, text="Enter Split Receipt Id", width=60, anchor='center',
                                 justify=CENTER,
                                 font=('arial narrow', 13, 'normal'),
                                 bg='snow', fg='green')
        searchinfo_label.grid(row=1, column=0, padx=1, columnspan=3)

        # design item details frame - starts
        itemid_label = Label(item_detailsFrame, text="Receipt Id", width=12, anchor=W, justify=LEFT,
                             font=('arial narrow', 13, 'normal'),
                             bg='snow')
        receiptId_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                    font=('arial narrow', 13, 'normal'),
                                    bg='light yellow')
        donatorname_label = Label(item_detailsFrame, text="Donated By", width=12, anchor=W, justify=LEFT,
                                  font=('arial narrow', 13, 'normal'),
                                  bg='snow')
        donatorname_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                      font=('arial narrow', 13, 'normal'),
                                      bg='light yellow')

        date_label = Label(item_detailsFrame, text="Date", width=12, anchor=W, justify=LEFT,
                           font=('arial narrow', 13, 'normal'),
                           bg='snow')
        date_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                               font=('arial narrow', 13, 'normal'),
                               bg='light yellow')
        balance_label = Label(item_detailsFrame, text="Balance(Rs.)", width=12, anchor=W, justify=LEFT,
                              font=('arial narrow', 13, 'normal'),
                              bg='snow')
        balance_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                  font=('arial narrow', 13, 'normal'),
                                  bg='light yellow')

        splitstatus_label = Label(item_detailsFrame, text="Split Status", width=12, anchor=W, justify=LEFT,
                                  font=('arial narrow', 13, 'normal'),
                                  bg='snow')
        splitstatus_labelText = Label(item_detailsFrame, text="", width=25, anchor=W, justify=LEFT,
                                      font=('arial narrow', 13, 'normal'),
                                      bg='light yellow')

        itemid_label.grid(row=0, column=0, pady=5)
        receiptId_labelText.grid(row=0, column=1, padx=5, pady=5)
        donatorname_label.grid(row=0, column=2, pady=5)
        donatorname_labelText.grid(row=0, column=3, padx=5, pady=5)
        date_label.grid(row=1, column=0, pady=5)
        date_labelText.grid(row=1, column=1, padx=5, pady=5)
        balance_label.grid(row=1, column=2, pady=5)
        balance_labelText.grid(row=1, column=3, padx=5, pady=5)
        splitstatus_label.grid(row=2, column=0, pady=5)
        splitstatus_labelText.grid(row=2, column=1, padx=5, pady=5)
        rembalance_TextLabel = Label(upperFrame, text="0", width=10, justify=CENTER,
                                     font=('arial narrow', 14, 'bold'), bg='navy', fg='white')
        searchitemid_result = partial(self.searchSplitOpenCandidates, receiptid_Text, receiptId_labelText,
                                      donatorname_labelText, date_labelText,
                                      balance_labelText, splitstatus_labelText, searchinfo_label, rembalance_TextLabel)
        search_btn.configure(command=searchitemid_result)
        close_btn.configure(command=split_donation_window.destroy)
        # design item details frame - end

        member_idLabel = Label(upperFrame, text="Member Id", width=10, anchor=W, justify=LEFT,
                               font=('arial narrow', 13, 'normal'), bg='snow')

        member_IdText = Entry(upperFrame, width=22, font=('arial narrow', 13, 'normal'), justify=LEFT,
                              bg='light yellow', state=DISABLED)
        customer_namelabel = Label(upperFrame, text="Customer Name", width=13, anchor=W, justify=LEFT,
                                   font=('arial narrow', 13, 'normal'), bg='snow')
        customer_nameText = Entry(upperFrame, width=22, font=('arial narrow', 13, 'normal'), justify=LEFT,
                                  state=DISABLED)

        item_idLabel = Label(upperFrame, text="Item Id", width=10, anchor=W, justify=LEFT,
                             font=('arial narrow', 13, 'normal'), bg='snow')
        item_idText = Entry(upperFrame, width=22, font=('arial narrow', 13, 'normal'), justify=LEFT, bg='light yellow',
                            textvariable=default_text1)

        date_Label = Label(upperFrame, text="Date", width=13, anchor=W, justify=LEFT,
                           font=('arial narrow', 13, 'normal'), bg='snow')

        cal = DateEntry(upperFrame, width=22, date_pattern='dd/MM/yyyy', font=('arial narrow', 12, 'normal'),
                        justify=LEFT)

        splitamt_label = Label(upperFrame, text="Split Amt.(Rs.)", width=11, anchor=W, justify=LEFT,
                               font=('arial narrow', 13, 'normal'), bg='snow')

        splitamtText = Entry(upperFrame, width=22, font=('arial narrow', 13, 'normal'), justify=LEFT, bg='light yellow')

        paymentmode_label = Label(upperFrame, text="Payment Mode", width=13, anchor=W, justify=LEFT,
                                  font=('arial narrow', 13, 'normal'), bg='snow')
        paymentMode_text = StringVar(upperFrame)
        paymentMode_list = ['Cash', 'Bank Transfer', 'Cheque', 'Demand Draft', 'Other']
        paymentMode_text.set("Other")
        paymentMode_menu = OptionMenu(upperFrame, paymentMode_text, *paymentMode_list)
        paymentMode_menu.configure(width=20, font=('arial narrow', 12, 'normal'), bg='light yellow', anchor=W,
                                   justify=LEFT)

        cart_item_countLabel = Label(upperFrame, text="Split Count", width=10, anchor=W, justify=LEFT,
                                     font=('arial narrow', 13, 'normal'), bg='snow')

        cart_item_count = Label(upperFrame, text="0", width=22, anchor='center', justify=LEFT,
                                font=('arial narrow', 13, 'normal'), bg='snow')

        rem_balanceLabel = Label(upperFrame, text="Rem. Amt.(Rs.)", width=13, anchor=W, justify=LEFT,
                                 font=('arial narrow', 13, 'normal'), bg='snow')

        var = IntVar()
        viewPurchaseBy_Result = partial(self.enablePurchaseViewBy_RadioSelection, var, member_IdText,
                                        customer_nameText, receiptid_Text, item_idText, cart_item_count,
                                        rembalance_TextLabel, balance_labelText)
        purchasebyMemeberId_radioBtn = Radiobutton(upperFrame, text="Split by ID", variable=var, value=1,
                                                   command=viewPurchaseBy_Result, width=12, bg='snow',
                                                   font=('arial narrow', 12, 'normal'), anchor=W, justify=LEFT)
        purchasebyName_radioBtn = Radiobutton(upperFrame, text="Split by Name", variable=var, value=2,
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
        splitamt_label.grid(row=4, column=0)
        splitamtText.grid(row=4, column=1)
        paymentmode_label.grid(row=4, column=2)
        paymentMode_menu.grid(row=4, column=3)
        cart_item_countLabel.grid(row=5, column=0)
        cart_item_count.grid(row=5, column=1, padx=5, pady=5)
        rem_balanceLabel.grid(row=5, column=2, padx=5)
        rembalance_TextLabel.grid(row=5, column=3)
        # upper frame end

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(upperFrame, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=8, column=0, pady=6, columnspan=4)

        purchase_btn = Button(buttonFrame, text="Commit", fg="Black",
                              font=NORM_FONT, width=12, bg='light grey', state=DISABLED)
        addtocart_btn = Button(buttonFrame, text="Add Split", fg="Black", font=NORM_FONT, width=12, bg='light cyan')

        insert_result = partial(self.addSplit, split_donation_window,
                                item_idText,
                                member_IdText,
                                customer_nameText,
                                purchase_btn,
                                addtocart_btn,
                                splitamtText,
                                cart_item_count,
                                rembalance_TextLabel,
                                paymentMode_text,
                                searchinfo_label,
                                var,
                                cal, balance_labelText)

        addtocart_btn.configure(command=insert_result)
        result_searcbtnState = partial(self.check_itemSearchBtnState, default_text1, search_btn, addtocart_btn)
        default_text1.trace("w", result_searcbtnState)

        addtocart_btn.grid(row=0, column=0)

        purchase_result = partial(self.deposit_splt_donation_to_transaction,
                                  split_donation_window, purchase_btn,
                                  addtocart_btn,
                                  searchinfo_label,
                                  var, cal,
                                  item_idText)

        purchase_btn.configure(command=purchase_result)
        purchase_btn.grid(row=0, column=1)

        clear_result = partial(self.clear_split_donation_form, split_donation_window,
                               item_idText,
                               member_IdText,
                               customer_nameText,
                               purchase_btn,
                               addtocart_btn,
                               splitamtText,
                               cart_item_count,
                               rembalance_TextLabel,
                               paymentMode_text,
                               searchinfo_label,
                               var,
                               cal, balance_labelText,
                               receiptId_labelText,
                               donatorname_labelText,
                               date_labelText,
                               splitstatus_labelText,
                               receiptid_Text)

        # create a Reset Button and place into the split_donation_window window
        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=12, bg='light cyan', underline=0)
        clear.grid(row=0, column=3)

        # ---------------------------------Button Frame End----------------------------------------
        # shortcut keys for keyboard actions
        split_donation_window.bind('<Return>', lambda event=None: search_btn.invoke())
        split_donation_window.bind('<Alt-c>', lambda event=None: close_btn.invoke())
        split_donation_window.bind('<Alt-r>', lambda event=None: clear.invoke())

        split_donation_window.focus()
        split_donation_window.grab_set()
        self.display_open_splitDonation_Details(split_donation_window, 680, 35, 560, 542)
        mainloop()

    def clear_split_donation_form(self, split_donation_window,
                                  item_idText,
                                  member_IdText,
                                  customer_nameText,
                                  purchase_btn,
                                  addtocart_btn,
                                  splitamtText,
                                  cart_item_count,
                                  rembalance_TextLabel,
                                  paymentMode_text,
                                  searchinfo_label,
                                  var,
                                  cal,
                                  receiptId_labelText,
                                  donatorname_labelText,
                                  date_labelText,
                                  balance_labelText,
                                  splitstatus_labelText, receiptid_Text):
        member_IdText.delete(0, END)
        member_IdText.configure(fg='black')
        customer_nameText.delete(0, END)
        customer_nameText.configure(fg='black')
        item_idText.delete(0, END)
        item_idText.configure(fg='black')
        splitamtText.delete(0, END)
        splitamtText.configure(fg='black')
        paymentMode_text.set("Other")
        cart_item_count['text'] = "0"
        balance_labelText['text'] = ""
        rembalance_TextLabel['text'] = "0"
        receiptId_labelText['text'] = ""
        donatorname_labelText['text'] = ""
        date_labelText['text'] = ""
        splitstatus_labelText['text'] = ""
        receiptid_Text.delete(0, END)
        searchinfo_label.configure(text="Enter Donation Id", fg='green')
        self.list_split_open_record = []
        purchase_btn.configure(state=DISABLED, bg="light grey")
        addtocart_btn.configure(state=DISABLED, bg="light grey")

    def enablePurchaseViewBy_RadioSelection(self, var, member_IdText,
                                            customer_nameText, itemId_Text, item_idText, cart_item_count,
                                            amount_payableTextLabel, balance_labelText):
        print("Enabling the view by date section Radiobutton :", var.get())
        self.splitTransactionList = []
        cart_item_count.configure(text="0")
        amount_payableTextLabel.configure(text=balance_labelText.cget("text"))
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

    def check_itemSearchBtnState(self, n, m, x, itemId_text, search_btn, submit):
        print("check_for_search button  enabling")

        if n.get() != "" and len(n.get()) > 2:
            m.configure(state=NORMAL, bg='light cyan')
            x.configure(state=NORMAL, bg='light cyan')
        else:
            m.configure(state=DISABLED, bg='light grey')
            x.configure(state=DISABLED, bg='light grey')

    def addSplit(self, split_donation_window,
                 item_idText,
                 member_IdText,
                 customer_nameText,
                 purchase_btn,
                 addtocart_btn,
                 splitamtText,
                 cart_item_count,
                 rembalance_TextLabel,
                 paymentMode_text,
                 searchinfo_label,
                 var,
                 cal, balance_labelText):
        print("var = ", var.get())
        if var.get() == 1:
            libMemberId = member_IdText.get()
        else:
            libMemberId = "Not Available"

        dateTimeObj = cal.get_date()
        dateOfPurchase = dateTimeObj.strftime("%Y-%m-%d")
        # validate the member , if  already registered

        print("addSplit->MemberId :", libMemberId)
        bMemberValidForSplit = True
        if var.get() == 1:
            bValidMember = self.obj_commonUtil.validate_memberId_Excel(libMemberId, 1)
            for iLoop in range(0, len(self.splitTransactionList)):
                if (self.splitTransactionList[iLoop][4] == libMemberId):
                    bMemberValidForSplit = False
                    break;
        else:
            bValidMember = True
        print("var.get() : ", var.get(), bValidMember, "Member bValidMember :", bValidMember)
        # issue books only to the registered member
        if bValidMember == True:
            if bMemberValidForSplit == True:
                today = date.today()
                if (dateTimeObj <= today) and (splitamtText.get()).isnumeric():
                    print("Date is OK")
                    splitCandidate_receiptId = item_idText.get()
                    # To open the workbook
                    # workbook object is created
                    filename = InitDatabase.getInstance().get_splittransaction_database_name()
                    wb_obj = openpyxl.load_workbook(filename)

                    # Get workbook active sheet object
                    # from the active attribute
                    sheet_obj = wb_obj.active
                    totalrecords = self.obj_commonUtil.totalrecords_excelDataBase(filename)
                    for iLoop in range(1, totalrecords + 1):
                        donationReceiptId = str(sheet_obj.cell(row=iLoop + 1, column=10).value)  # item id in database
                        if donationReceiptId == splitCandidate_receiptId:
                            # check if  requested donation receipt is available in database
                            if (int(sheet_obj.cell(row=iLoop + 1, column=3).value) >= int(
                                    splitamtText.get())) and (int(rembalance_TextLabel.cget("text")) >= int(
                                splitamtText.get())):  # if amount specified is less than equal to balance amount
                                description = sheet_obj.cell(row=iLoop + 1, column=5).value
                                modeOftransaction = paymentMode_text.get()
                                # prepare the cart locally and retain it as long as purchase button is pressed
                                arr_SplitList = [dateTimeObj, splitamtText.get(), description, modeOftransaction,
                                                 libMemberId]
                                self.splitTransactionList.append(arr_SplitList)

                                cart_item_count['text'] = str(len(self.splitTransactionList))

                                # calculate the total mrp
                                rem_balance = int(balance_labelText.cget("text"))
                                for iLoop in range(0, len(self.splitTransactionList)):
                                    rem_balance = rem_balance - (int(self.splitTransactionList[iLoop][1]))

                                rembalance_TextLabel['text'] = str(rem_balance)
                                purchase_btn.configure(state=NORMAL, bg='light cyan')
                            else:
                                searchinfo_label.configure(
                                    text="Split amount cannot be greater than original/remaining balance ", fg='red')
                            break
                else:
                    searchinfo_label.configure(text="Invalid Split Amt/Future Date Chosen !!!", fg='red')
            else:
                searchinfo_label.configure(text="Member already exists for this split !!!", fg='red')
        else:
            searchinfo_label.configure(text="Invalid Member Id !!!", fg='red')

    def deposit_splt_donation_to_transaction(self, split_donation_window, purchase_btn,
                                             addtocart_btn,
                                             searchinfo_label,
                                             var, cal,
                                             item_idText):

        purchase_btn.configure(state=DISABLED, bg='light grey')
        addtocart_btn.configure(state=DISABLED, bg='light grey')
        print("Donation Split process starts for :")
        print("--------------------------------------------------------------")
        print(self.splitTransactionList)
        print("--------------------------------------------------------------")

        '''
            [dateTimeObj, splitamtText.get(), description,modeOftransaction,libMemberId]
            '''

        for iLoop in range(0, len(self.splitTransactionList)):
            print("For debugging Seva Amt :", self.splitTransactionList[iLoop][1], "Max donation allowed :",
                  MAX_ALLOWED_DONATION)
            # central moentary sheet
            filename_MonetarySheet = InitDatabase.getInstance().get_seva_deposit_database_name()  # PATH_SEVA_SHEET

            # writting the credit in Master Seva Sheet - starts

            # open seva rashi sheet and enter the data --start
            wb_obj = openpyxl.load_workbook(filename_MonetarySheet)
            sheet_obj = wb_obj.active
            total_records = self.obj_commonUtil.totalrecords_excelDataBase(filename_MonetarySheet)
            # Receipt voucher is generated serially and is hence always unique
            invoice_id = self.obj_commonUtil.generateMonetaryDonationReceiptId(VIHANGAM_YOGA_KARNATAKA_TRUST)
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

            # new transaction data is assigned in the moentary sheet
            sheet_obj.cell(row=row_no, column=1).value = serial_no
            sheet_obj.cell(row=row_no, column=2).value = self.splitTransactionList[iLoop][1]
            sheet_obj.cell(row=row_no, column=3).value = str(balance_amount + int(self.splitTransactionList[iLoop][1]))
            sheet_obj.cell(row=row_no, column=4).value = self.splitTransactionList[iLoop][4]

            member_data = self.obj_commonUtil.retrieve_MemberRecords(self.splitTransactionList[iLoop][4],
                                                                     1, SEARCH_BY_MEMBERID)
            sheet_obj.cell(row=row_no, column=5).value = member_data[2]
            sheet_obj.cell(row=row_no, column=6).value = str(self.splitTransactionList[iLoop][0])  # Date
            sheet_obj.cell(row=row_no, column=7).value = str(self.splitTransactionList[iLoop][2])
            sheet_obj.cell(row=row_no, column=8).value = "Not Applicable"
            sheet_obj.cell(row=row_no, column=9).value = "Center Manager"
            sheet_obj.cell(row=row_no, column=10).value = member_data[7]  # address
            sheet_obj.cell(row=row_no, column=11).value = str(self.splitTransactionList[iLoop][3])
            sheet_obj.cell(row=row_no, column=12).value = "Not Applicable"
            sheet_obj.cell(row=row_no, column=13).value = "Center Manager"
            sheet_obj.cell(row=row_no, column=14).value = invoice_id

            wb_obj.save(filename_MonetarySheet)

            # writting the credit in Master Seva Sheet - starts
            bAlike_Seva = True
            if self.splitTransactionList[iLoop][2] == "Monthly Seva":
                path_seva_sheet = InitDatabase.getInstance().get_monthly_seva_database_name()  # PATH_MONTHLY_SEVA_SHEET
            elif self.splitTransactionList[iLoop][2] == "Gaushala Seva":
                path_seva_sheet = InitDatabase.getInstance().get_gaushala_seva_database_name()  # PATH_GAUSHALA_SEVA_SHEET
            elif self.splitTransactionList[iLoop][2] == "Hawan Seva":
                path_seva_sheet = InitDatabase.getInstance().get_hawan_seva_database_name()  # PATH_HAWAN_SEVA_SHEET
            elif self.splitTransactionList[iLoop][2] == "Event/Prachar Seva":
                path_seva_sheet = InitDatabase.getInstance().get_prachar_event_seva_database_name()  # PATH_EVENT_SEVA_SHEET
            elif self.splitTransactionList[iLoop][2] == "Aarti Seva":
                path_seva_sheet = InitDatabase.getInstance().get_aarti_seva_database_name()  # PATH_AARTI_SEVA_SHEET
            elif self.splitTransactionList[iLoop][2] == "Ashram Seva(Generic)":
                path_seva_sheet = InitDatabase.getInstance().get_ashram_seva_database_name()  # PATH_ASHRAM_GENERIC_SEVA_SHEET
            elif self.splitTransactionList[iLoop][2] == "Ashram Nirmaan Seva":
                path_seva_sheet = InitDatabase.getInstance().get_ashram_nirmaan_seva_database_name()  # PATH_ASHRAM_NIRMAAN_SHEET
            elif self.splitTransactionList[iLoop][2] == "Yoga Fees":
                path_seva_sheet = InitDatabase.getInstance().get_yoga_seva_database_name()  # PATH_YOGA_FEES_SHEET
            elif self.splitTransactionList[iLoop][2] == "Akshay-Patra Seva":
                path_seva_sheet = InitDatabase.getInstance().get_akshay_patra_database_name()  # PATH_AKSHAY_PATRA_DATABASE
            else:
                bAlike_Seva = False
                pass

            # writting into respective seva sheet - start
            if bAlike_Seva:
                wb_sevasheetobj = openpyxl.load_workbook(path_seva_sheet)
                sevasheet_obj = wb_sevasheetobj.active
                total_records_seva_sheet = self.obj_commonUtil.totalrecords_excelDataBase(path_seva_sheet)

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
                    sevasheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                             bold=False)
                    sevasheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                       vertical='center')

                # new book data is assigned to respective cells in row
                sevasheet_obj.cell(row=row_no, column=1).value = serial_no
                sevasheet_obj.cell(row=row_no, column=2).value = str(self.splitTransactionList[iLoop][1])
                sevasheet_obj.cell(row=row_no, column=3).value = str(
                    balance_amount + int(self.splitTransactionList[iLoop][1]))
                sevasheet_obj.cell(row=row_no, column=4).value = str(self.splitTransactionList[iLoop][4])
                sevasheet_obj.cell(row=row_no, column=5).value = member_data[2]
                sevasheet_obj.cell(row=row_no, column=6).value = str(self.splitTransactionList[iLoop][0])
                sevasheet_obj.cell(row=row_no, column=7).value = str(self.splitTransactionList[iLoop][2])
                sevasheet_obj.cell(row=row_no, column=8).value = "Not Applicable"
                sevasheet_obj.cell(row=row_no, column=9).value = "Center Manager"
                sevasheet_obj.cell(row=row_no, column=10).value = member_data[7]  # address
                sevasheet_obj.cell(row=row_no, column=11).value = str(self.splitTransactionList[iLoop][3])
                sevasheet_obj.cell(row=row_no, column=12).value = "Not Applicable"
                sevasheet_obj.cell(row=row_no, column=13).value = "Center Manager"
                sevasheet_obj.cell(row=row_no, column=14).value = invoice_id

                wb_sevasheetobj.save(path_seva_sheet)
                # writting into respective seva sheet - end

                # open transaction sheet and enter the data
                # receiving donation is a credit transaction for the organization

                file_name_transaction = InitDatabase.getInstance().get_transaction_database_name()  # PATH_TRANSACTION_SHEET

                transaction_wb_obj = openpyxl.load_workbook(file_name_transaction)
                transaction_sheet_obj = transaction_wb_obj.active
                total_records_transaction = self.obj_commonUtil.totalrecords_excelDataBase(file_name_transaction)

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
                    transaction_sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                               vertical='center')

                # new book data is assigned to respective cells in row
                transaction_sheet_obj.cell(row=row_no, column=1).value = serial_no
                transaction_sheet_obj.cell(row=row_no, column=2).value = str(self.splitTransactionList[iLoop][0])
                transaction_sheet_obj.cell(row=row_no, column=3).value = str(self.splitTransactionList[iLoop][1])
                transaction_sheet_obj.cell(row=row_no, column=4).value = "---"
                transaction_sheet_obj.cell(row=row_no, column=5).value = str(self.splitTransactionList[iLoop][2])
                transaction_sheet_obj.cell(row=row_no, column=6).value = str(self.splitTransactionList[iLoop][3])
                transaction_sheet_obj.cell(row=row_no, column=7).value = str(
                    self.splitTransactionList[iLoop][4])  # authorizor id
                transaction_sheet_obj.cell(row=row_no, column=8).value = member_data[2]
                transaction_sheet_obj.cell(row=row_no, column=9).value = str(
                    balance_amount + int(self.splitTransactionList[iLoop][1]))
                transaction_sheet_obj.cell(row=row_no, column=10).value = invoice_id
                transaction_wb_obj.save(file_name_transaction)
                # open transaction sheet and enter the data --end
                self.generateSplitDonationReceipt(self.splitTransactionList[iLoop][4],
                                                  self.splitTransactionList[iLoop][1],
                                                  self.splitTransactionList[iLoop][2],
                                                  "Center Manager",
                                                  self.splitTransactionList[iLoop][0],
                                                  self.splitTransactionList[iLoop][3],
                                                  invoice_id, member_data)


        # workbook object is created for Split transaction sheet
        filename_splitSheet = InitDatabase.getInstance().get_splittransaction_database_name()

        wb_obj = openpyxl.load_workbook(filename_splitSheet)

        # Get workbook active sheet object
        # from the active attribute
        splitsheet_obj = wb_obj.active
        totalrecords = self.obj_commonUtil.totalrecords_excelDataBase(filename_splitSheet)
        print("totalrecords :", totalrecords)
        for iLoop in range(1, totalrecords + 1):
            donationReceiptId = str(splitsheet_obj.cell(row=iLoop + 1, column=10).value)  # item id in database
            print("donationReceiptId :", donationReceiptId, "item_idText :", item_idText.get())
            if donationReceiptId == item_idText.get():
                print("Match found 1")
                initial_balance = str(splitsheet_obj.cell(row=iLoop + 1, column=3).value)  # current balance
                print("initial_balance :", initial_balance)
                # check if  requested donation receipt is available in database
                adjusted_balance = 0
                for iLoop_adjBalance in range(0, len(self.splitTransactionList)):
                    adjusted_balance = adjusted_balance + (int(self.splitTransactionList[iLoop_adjBalance][1]))

                new_balance = int(initial_balance) - adjusted_balance
                print("adjusted_balance :", adjusted_balance, "new_balance :", new_balance)
                splitsheet_obj.cell(row=iLoop + 1, column=3).value = str(new_balance)
                if new_balance == 0:
                    # if new balance is 0, split status of the donation must be changed to "Closed"
                    splitsheet_obj.cell(row=iLoop + 1, column=4).value = "Closed"
                break
        wb_obj.save(filename_splitSheet)

        # update the total balance
        self.obj_commonUtil.calculateTotalAvailableBalance()

        print("Split complete")

        text_withID = "Split Success"
        searchinfo_label.configure(text=text_withID, fg='green')

        # Refresh the open table
        self.display_open_splitDonation_Details(split_donation_window, 680, 35, 560, 542)

    def generateSplitDonationReceipt(self, donator_idText,
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
        sheet_obj.cell(row=15, column=4).value = str(categoryText)

        sheet_obj.cell(row=16, column=4).value = str(seva_amountText)
        sheet_obj.cell(row=17, column=4).value = str(num2words.num2words(int(seva_amountText))) + " Rs. only"
        sheet_obj.cell(row=18, column=4).value = str(paymentMode_text)
        sheet_obj.cell(row=19, column=4).value = str(collector_nameText)

        wb_obj.save(file_name)
        pdf_file = "..\\Expanse_Data\\" + currentYearDirectory + "\\Seva_Rashi\\Receipts\\Receipts\\" + invoice_id + ".pdf"

        self.obj_commonUtil.convertExcelToPdf(file_name, pdf_file)

        destdir_repo = InitDatabase.getInstance().get_invoice_directory_name() + "\\" + invoice_id + ".pdf"
        self.obj_commonUtil.updateInvoiceTable(invoice_id, destdir_repo)
        desktop_repo = InitDatabase.getInstance().get_desktop_invoices_directory_path() + "\\" + invoice_id + ".pdf"
        self.obj_commonUtil.updateMonetaryDonationReceiptBooklet(invoice_id)
        copyfile(pdf_file, destdir_repo)
        copyfile(pdf_file, desktop_repo)
        # os.startfile(pdf_file, 'print')

    def display_open_splitDonation_Details(self, split_open_window, startx, starty, xwidth, xheight):
        myframe = Frame(split_open_window, relief=GROOVE, width=520, height=407, bd=4)
        myframe.place(x=startx, y=starty)

        mycanvas = Canvas(myframe)
        frame = Frame(mycanvas, width=200, height=100)
        myscrollbar = Scrollbar(myframe, orient="vertical", command=mycanvas.yview)
        mycanvas.configure(yscrollcommand=myscrollbar.set)

        myscrollbar.pack(side="right", fill="y")
        mycanvas.pack(side="left")
        mycanvas.create_window((0, 0), window=frame, anchor='nw')

        result = partial(self.myfunction, xwidth, xheight, mycanvas)

        frame.bind("<Configure>", result)

        label_Sno = Label(frame, text="S.No", width=5, height=1, anchor='center',
                          justify=CENTER,
                          font=('arial narrow', 11, 'bold'),
                          bg='wheat')

        label_detail1 = Label(frame, text="RV No", width=10, height=1, anchor='center',
                              justify=CENTER,
                              font=('arial narrow', 11, 'bold'),
                              bg='wheat')

        label_detail2 = Label(frame, text="Donar Name", width=25, height=1, anchor='center',
                              justify=CENTER,
                              font=('arial narrow', 11, 'bold'),
                              bg='wheat')

        label_detail3 = Label(frame, text="Date", width=15, height=1,
                              anchor='center',
                              justify=CENTER,
                              font=('arial narrow', 11, 'bold'),
                              bg='wheat')

        label_detail4 = Label(frame, text="Amount(Rs.)", width=10, height=1,
                              anchor='center',
                              justify=CENTER,
                              font=('arial narrow', 11, 'bold'),
                              bg='wheat')
        label_Sno.grid(row=0, column=1, padx=2, pady=5)
        label_detail1.grid(row=0, column=2, padx=2, pady=5)
        label_detail2.grid(row=0, column=3, padx=2, pady=5)
        label_detail3.grid(row=0, column=4, padx=2, pady=5)
        label_detail4.grid(row=0, column=5, padx=2, pady=5)
        open_split_records = self.retrieveOpenSplitRecords()

        for row_index in range(0, len(open_split_records)):
            # critical stock ->stock with quantity is 0 or 1
            for column_index in range(1, 6):
                if column_index == 5:
                    width_column = 10
                elif column_index == 1:
                    width_column = 5
                elif column_index == 2:
                    width_column = 10
                elif column_index == 3:
                    width_column = 25
                else:
                    width_column = 15

                label_detail = Label(frame, text="No Data", width=width_column, height=1,
                                     anchor='center', justify=LEFT,
                                     font=('arial narrow', 13, 'normal'),
                                     bg='light yellow')
                label_detail.grid(row=row_index + 1, column=column_index, padx=2, pady=5, sticky=W)

                if column_index == 1:
                    label_detail['text'] = str(row_index + 1)
                else:
                    label_detail['text'] = open_split_records[row_index][column_index - 2]

    def myfunction(self, xwidth, yheight, mycanvas, event):
        mycanvas.configure(scrollregion=mycanvas.bbox("all"), width=xwidth, height=yheight)

    def searchSplitOpenCandidates(self, receiptid_Text,
                                  receiptId_labelText,
                                  donatorname_labelText,
                                  date_labelText,
                                  balance_labelText,
                                  splitstatus_labelText,
                                  searchinfo_label,
                                  rembalance_TextLabel):
        # To open the workbook
        # workbook object is created
        filename = InitDatabase.getInstance().get_splittransaction_database_name()
        wb_obj = openpyxl.load_workbook(filename)
        bValidItem = False
        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.obj_commonUtil.totalrecords_excelDataBase(filename)
        print(" Total records in split transaction database: ", total_records)
        for iLoop in range(1, total_records + 1):
            cell_book_id = sheet_obj.cell(row=iLoop + 1, column=10)
            split_status = sheet_obj.cell(row=iLoop + 1, column=4).value
            print("Entered Receipt id :", receiptid_Text.get(), "Receipt id from sheet :", cell_book_id.value)
            if (str(cell_book_id.value) == receiptid_Text.get()):  # if receipt id is found
                bValidItem = True
                # print("Donation Id is found")
                receiptId_labelText['text'] = sheet_obj.cell(row=iLoop + 1, column=10).value
                donatorname_labelText['text'] = sheet_obj.cell(row=iLoop + 1, column=8).value
                date_labelText['text'] = sheet_obj.cell(row=iLoop + 1, column=2).value
                balance_labelText['text'] = str(sheet_obj.cell(row=iLoop + 1, column=3).value)
                splitstatus_labelText['text'] = sheet_obj.cell(row=iLoop + 1, column=4).value

                if (split_status == "Closed"):
                    searchinfo_label.configure(text="Split for this Donation is closed !!", fg='green')
                elif (split_status == "Open"):
                    searchinfo_label.configure(text="Donation found !!", fg='green')
                else:
                    pass
                rembalance_TextLabel['text'] = str(sheet_obj.cell(row=iLoop + 1, column=3).value)
                break
            else:
                pass
        if bValidItem == False:
            searchinfo_label.configure(text="Invalid ID for split", fg='red')

    def retrieveOpenSplitRecords(self):
        print("searchOpenSplitRecords->>Entry")
        self.splitTransactionList = []

        # To open the workbook
        # workbook object is created
        filename = InitDatabase.getInstance().get_splittransaction_database_name()
        wb_obj = openpyxl.load_workbook(filename)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        list_split_open_record = []
        total_records = self.obj_commonUtil.totalrecords_excelDataBase(filename)
        print(" Total records in split transaction database: ", total_records)
        for iLoop in range(1, total_records + 1):
            split_status = sheet_obj.cell(row=iLoop + 1, column=4).value

            if split_status == "Open":  # if split status is Open

                print("Donation Id is found")
                rv_no = str(sheet_obj.cell(row=iLoop + 1, column=10).value)
                donar_name = str(sheet_obj.cell(row=iLoop + 1, column=8).value)
                dateOfDonation = str(sheet_obj.cell(row=iLoop + 1, column=2).value)
                donated_Amt = str(sheet_obj.cell(row=iLoop + 1, column=3).value)
                arr_split_record = []
                arr_split_record = [rv_no, donar_name, dateOfDonation, donated_Amt]
                list_split_open_record.append(arr_split_record)

        print("searchOpenSplitRecords->>End")
        return list_split_open_record

    def split_donation_list_view(self, master):
        split_donation_list_window = Toplevel(master)
        split_donation_list_window.title("Split Donation Window ")
        split_donation_list_window.geometry('620x400+150+100')
        split_donation_list_window.configure(background='wheat')
        split_donation_list_window.resizable(width=False, height=False)
        # delete "X" button in window will be not-operational
        split_donation_list_window.protocol('WM_DELETE_WINDOW', self.obj_commonUtil.donothing)

        heading_list = Label(split_donation_list_window, text="Open Split Candidates",
                             font=('ariel narrow', 15, 'bold'),
                             bg='wheat')
        heading_list.grid(row=0, column=2, padx=50)

        self.display_open_splitDonation_Details(split_donation_list_window, 10, 35, 570, 300)

        buttonFrame = Frame(split_donation_list_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=1, column=2, pady=320, padx=120)
        close_btn = Button(buttonFrame, text="Close", fg="Black", font=NORM_FONT, width=12, bg='light cyan',
                           command=split_donation_list_window.destroy)

        print_btn = Button(buttonFrame, text="Print", fg="Black",
                           font=NORM_FONT, width=12, bg='light grey', state=DISABLED)

        viewPDF_btn = Button(buttonFrame, text="View PDF", fg="Black", font=NORM_FONT, width=12, bg='light cyan',
                             state=NORMAL)
        viewpdf_result = partial(self.viewPDF_splitdonation_list, split_donation_list_window, close_btn, print_btn,
                                 viewPDF_btn)
        viewPDF_btn.configure(command=viewpdf_result)

        viewPDF_btn.grid(row=0, column=0)
        print_btn.grid(row=0, column=1)
        close_btn.grid(row=0, column=2)

    def close_split_donation_list_view(self, split_donation_list_window):
        # Existing data in the template sheet is erased for reuse of the template 
        # for the next operation

        now = datetime.now()
        dt_string = now.strftime("%d_%b_%Y_%H%M%S")
        currentyear = now.strftime("%Y")
        src_file = "..\\Expanse_Data\\" + currentyear + "\\Transaction\\Account_Statement\\Template\\Split_Donation_Open.xlsx"
        # change file permissions to write the data into file

        wb_template = openpyxl.load_workbook(src_file)
        template_sheet = wb_template.active

        # calculate the total number of records in the file
        total_records = self.obj_commonUtil.totalrecords_excelDataBase(src_file)
        for rows in range(4, total_records + 1):
            for columns in range(2, 7):
                # write blank in all respective columns 
                template_sheet.cell(row=rows, column=columns).value = ""

        wb_template.save(src_file)

        # make the file readonly

        # Close the window after template sheet data has been cleared
        split_donation_list_window.destroy()

    def viewPDF_splitdonation_list(self, split_donation_list_window, close_btn, print_btn, viewPDF_btn):
        now = datetime.now()
        dt_string = now.strftime("%d_%b_%Y_%H%M%S")
        datforPrint = now.strftime("%d_%b_%Y")
        currentyear = now.strftime("%Y")
        destination_file = "..\\Expanse_Data\\" + currentyear + "\\Transaction\\Account_Statement\\Statements\\Split_Donation_Open_List_" + dt_string + ".pdf"
        # write the  sorted record in list template
        template_sheet = "..\\Expanse_Data\\" + currentyear + "\\Transaction\\Account_Statement\\Template\\Split_Donation_Open.xlsx"

        wb_template = openpyxl.load_workbook(template_sheet)
        template_sheet_obj = wb_template.active
        split_donation_list = self.retrieveOpenSplitRecords()
        dict_index = 1
        starting_index = 4
        for row_index in range(0, len(split_donation_list)):
            for column_index in range(1, 6):
                # print("Inserting elemnt :",row_index +1)
                if column_index == 1:
                    template_sheet_obj.cell(row=starting_index, column=column_index + 1).value = str(dict_index)
                else:
                    template_sheet_obj.cell(row=starting_index, column=column_index + 1).value = str(
                        split_donation_list[row_index - 1][column_index - 2])
            starting_index = starting_index + 1
            dict_index = dict_index + 1

        # fill the date column in prepraed excel template
        template_sheet_obj.cell(row=1, column=6).value = str(datforPrint)
        wb_template.save(template_sheet)

        # revoke write permissions after the data has been changed

        destination_copy_folder = InitDatabase.getInstance().get_statement_directory_path()
        '''
        obj_threadClass = myThread(11, "splitdonation_thread", 1, template_sheet,
                            destination_file, starting_index, "Dummy", print_btn, "Dummy",
                            destination_copy_folder)
        obj_threadClass.start()
        '''

        # Convert the excel to pdf
        # and copy the prepared statement to Desktop prepared statement folder
        self.obj_commonUtil.convertExcelToPdf(template_sheet, destination_file)
        self.obj_commonUtil.copyTheStatementFileToDesktop_file(destination_file, destination_copy_folder)

        # just opens the prepared PDF file on Desktop/STDOUT
        os.startfile(destination_file)

        # redefine command for print action
        print_result = partial(self.printPDFfile, destination_file)
        print_btn.configure(state=NORMAL, bg='light cyan', command=print_result)

        # redefine command for close button action
        # After the PDF preparation ,
        # close button must ensure the data wipe in template sheet upon exit
        close_result = partial(self.close_split_donation_list_view, split_donation_list_window)
        close_btn.configure(command=close_result)

        # redefine command for print action
        # If again pressed , doesn't need to prepare the pdf again , instead simply open the same 
        viewPdf_result = partial(self.openPDFfile, destination_file)
        viewPDF_btn.configure(command=viewPdf_result)

    def printPDFfile(self, dest_file):
        os.startfile(dest_file, 'print')

    def openPDFfile(self, dest_file):
        os.startfile(dest_file, 'print')
