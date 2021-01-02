from tkinter import *
from app_defines import *
from app_common import *

class NoncommercialEdit:
    # constructor for Library class
    def __init__(self, master):
        print("constructor called for noncommercial edit ")
        self.obj_commonUtil = CommonUtil()
        self.edit_noncommercialItem_data(master)


    def validate_noncommercialitemId_Excel(self, itemId, local_centerText):
        bIdExist = False
        print("validate_itemId_Excel--> Start for Item: ", itemId)

        # To open the workbook
        # workbook object is created
        subdir_commercialstock = "..\\Library_Stock\\" + local_centerText.get() + "\\NonCommercial_Stock"
        filename = subdir_commercialstock + "\\noncommercial_stock.xlsx"
        wb_obj = openpyxl.load_workbook(filename)

        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active

        total_records = self.obj_commonUtil.totalrecords_excelDataBase(filename)
        print(" Total records for books: ", total_records)
        for iLoop in range(2, total_records + 2):
            cell_obj = sheet_obj.cell(row=iLoop, column=2)
            if cell_obj.value == itemId:
                bIdExist = True

        return bIdExist

    def edit_noncommercialItem_data(self, master):
        display_dataWindow = Toplevel(master)

        headingForm = "Edit Non-Commercial Item Details"
        display_dataWindow.title("Edit Information Details ")

        display_dataWindow.geometry('800x375+250+150')
        display_dataWindow.configure(background='wheat')
        display_dataWindow.resizable(width=True, height=True)

        heading = Label(display_dataWindow, text=headingForm, font=('ariel narrow', 15, 'bold'),
                        bg='wheat')
        heading.grid(row=0, column=0, columnspan=4)
        upperFrame = Frame(display_dataWindow, width=205, height=100, bd=8, relief='ridge', bg='light yellow')
        upperFrame.grid(row=1, column=2, padx=20, pady=10, sticky=W)

        middleFrame = Frame(display_dataWindow, width=200, height=300, bd=8, relief='ridge', bg='light yellow')
        middleFrame.grid(row=2, column=2, padx=20, pady=10, sticky=W)

        infoFrame = Frame(display_dataWindow, width=200, height=100, bd=8, relief='ridge', bg='light yellow')
        infoFrame.grid(row=16, column=2, padx=90, pady=10, columnspan=5, sticky=W)

        itemIdLabel = Label(upperFrame, text="Item Id (NCI-)", width=9, anchor=W, justify=LEFT,
                            font=NORM_FONT, bg='light yellow')
        itemIdLabel.grid(row=1, column=0, padx=10, pady=10)
        item_Id = Entry(upperFrame, width=25, font=NORM_FONT, justify='center')
        item_Id.grid(row=1, column=1, pady=10)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(upperFrame, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=1, column=3, padx=22, pady=15, sticky=W)

        # ---------------------------------Button Frame End----------------------------------------

        # ---------------------------------Preparing display Area - start ---------------------------------

        itemnametext = StringVar(middleFrame)
        itemnamelabel = Label(middleFrame, text="Item Name", width=12, anchor=W, justify=LEFT,
                              font=NORM_FONT,
                              bg='light yellow')
        itemnamelabel.grid(row=4, column=2, padx=10, pady=5)
        itemname_Text = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=itemnametext,
                              font=NORM_FONT,
                              bg='snow')
        itemname_Text.grid(row=4, column=3, pady=5)

        # Display item Id - Row 4
        descriptiontext = StringVar(middleFrame)
        descriptionlabel = Label(middleFrame, text="Donar Name", width=12, anchor=W, justify=LEFT,
                                 font=NORM_FONT,
                                 bg='light yellow')
        descriptionlabel.grid(row=4, column=4, padx=10, pady=5)
        description_Text = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=descriptiontext,
                                 font=NORM_FONT,
                                 bg='snow')
        description_Text.grid(row=4, column=5, padx=5, pady=5)

        # Display Father name - Row 5

        # Display Country Name - Row 5
        quantitytext = StringVar(middleFrame)
        quantityLabel = Label(middleFrame, text="Quantity", width=12, anchor=W, justify=LEFT,
                              font=NORM_FONT,
                              bg='light yellow')
        quantityLabel.grid(row=5, column=2, padx=10, pady=5)
        quantity_Text = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=quantitytext,
                              font=NORM_FONT,
                              bg='snow')
        quantity_Text.grid(row=5, column=3, pady=5)

        unitpricetext = StringVar(middleFrame)
        unitpriceLabel = Label(middleFrame, text="Est. Value(Rs.)", width=12, anchor=W, justify=LEFT,
                               font=NORM_FONT,
                               bg='light yellow')
        unitpriceLabel.grid(row=5, column=4, padx=10, pady=5)
        unitprice_Text = Entry(middleFrame, text="", textvariable=unitpricetext, width=25, justify=LEFT,
                               font=NORM_FONT,
                               bg='snow')
        unitprice_Text.grid(row=5, column=5, padx=5, pady=5)

        racktext = StringVar(middleFrame)
        rackno_label = Label(middleFrame, text="Rack No.", width=12, anchor=W, justify=LEFT,
                             font=NORM_FONT,
                             bg='light yellow')
        rackno_label.grid(row=6, column=2, padx=10, pady=5)
        rackno_Text = Entry(middleFrame, text="", width=25, justify=LEFT, textvariable=racktext,
                            font=NORM_FONT,
                            bg='snow')
        rackno_Text.grid(row=6, column=3, pady=5)

        centerNameLabel = Label(middleFrame, text="Center Name", width=12, anchor=W, justify=LEFT,
                                font=NORM_FONT,
                                bg='light yellow')
        centerNameLabel.grid(row=6, column=4, padx=10, pady=5)

        local_centerText = StringVar(middleFrame)
        localCenterList = self.obj_commonUtil.getLocalCenterNames()
        print("Center list  - ", localCenterList)
        local_centerText.set(localCenterList[0])
        localcenter_menu = OptionMenu(middleFrame, local_centerText, *localCenterList)
        localcenter_menu.configure(width=27, font=('arial narrow', 12, 'normal'), bg='snow', anchor=W, justify=LEFT)
        localcenter_menu.grid(row=6, column=5, padx=10, pady=5)
        infoLabel = Label(infoFrame, text="Press Save button to save the modified records", width=60,
                          anchor='center',
                          justify=CENTER,
                          font=NORM_FONT,
                          bg='light yellow')

        print_button = Button(buttonFrame, text="Save", fg="Black",
                              font=NORM_FONT, width=12, bg='grey', state=DISABLED, underline=0)

        search_item = partial(self.assignDataForDisplay_editnonCommercialItemInfo, display_dataWindow, item_Id,
                              itemname_Text,
                              description_Text,
                              quantity_Text,
                              unitprice_Text,
                              rackno_Text,
                              infoLabel, print_button, local_centerText)

        # create a Search Button and place into the bookReturn_window window
        submit = Button(buttonFrame, text="Search", fg="Black", command=search_item,
                        font=NORM_FONT, width=12, bg='light cyan', underline=0)
        submit.grid(row=0, column=0)

        # create a Close Button and place into the bookReturn_window window

        save_result = partial(self.saveModifiedNonCommercialItemRecords, display_dataWindow, item_Id,
                              itemname_Text,
                              description_Text,
                              quantity_Text,
                              unitprice_Text,
                              rackno_Text,
                              infoLabel, print_button, local_centerText)

        print_button.configure(command=save_result)
        print_button.grid(row=0, column=1)

        # create a Close Button and place into the bookReturn_window window
        cancel = Button(buttonFrame, text="Close", fg="Black", command=display_dataWindow.destroy,
                        font=NORM_FONT, width=12, bg='light cyan', underline=0)
        cancel.grid(row=0, column=2)

        # ---------------------------Button frame ends

        infoLabel.grid(row=16, column=1, padx=10, pady=5)

        display_dataWindow.bind('<Return>', lambda event=None: submit.invoke())
        display_dataWindow.bind('<Alt-c>', lambda event=None: cancel.invoke())
        display_dataWindow.bind('<Alt-r>', lambda event=None: self.print_button.invoke())

        display_dataWindow.focus()
        display_dataWindow.grab_set()
        mainloop()

    def retrieve_nonCommercialItemRecords_Excel(self, itemid, local_centerText):
        print("retrieve_CommercialItemRecords_Excel->Start")
        recordList = []
        subdir_commercialstock = "..\\Library_Stock\\" + local_centerText.get() + "\\NonCommercial_Stock"
        file_name = subdir_commercialstock + "\\noncommercial_stock.xlsx"
        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(file_name)
        print(" Data extraction logic will be executed for : ", itemid)
        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active
        total_records = self.obj_commonUtil.totalrecords_excelDataBase(file_name)

        for iLoop in range(1, total_records + 2):
            cell_obj = sheet_obj.cell(row=iLoop + 1, column=2)
            if cell_obj.value == itemid:
                for iColumn in range(2, 18):
                    cell_value = sheet_obj.cell(row=iLoop + 1, column=iColumn).value
                    print("record[", iColumn, "] :", cell_value)
                    recordList.append(cell_value)
                break
        print("retrieve_CommercialItemRecords_Excel->End")
        return recordList

    def assignDataForDisplay_editnonCommercialItemInfo(self, display_dataWindow, item_Id,
                                                       itemname_Text,
                                                       description_Text,
                                                       quantity_Text,
                                                       unitprice_Text,
                                                       rackno_Text,
                                                       infoLabel, print_button, local_centerText):
        print("assignDataForDisplay_editCommercialItemInfo for :", item_Id.get())
        itemId_str = "NCI-" + item_Id.get()
        bItemValid = self.validate_noncommercialitemId_Excel(itemId_str, local_centerText)
        print_button.configure(bg='light grey', state=DISABLED)
        if bItemValid:
            infoLabel.configure(fg='green', text="Please press Save button to modify item records")
            item_data = self.retrieve_nonCommercialItemRecords_Excel(itemId_str, local_centerText)
            if len(item_data) > 0:
                itemname_Text.delete(0, END)
                itemname_Text.insert(0, item_data[1])
                description_Text.delete(0, END)
                description_Text.insert(0, item_data[5])
                quantity_Text.delete(0, END)
                quantity_Text.insert(0, item_data[2])
                unitprice_Text.delete(0, END)
                unitprice_Text.insert(0, item_data[3])
                rackno_Text.delete(0, END)
                rackno_Text.insert(0, item_data[14])
                local_centerText.set(item_data[15])

                print_button.configure(bg='light cyan', state=NORMAL)
        else:
            infoText = "Invalid Item Id :" + itemId_str
            infoLabel.configure(fg='red', text=infoText)
            itemname_Text.delete(0, END)
            itemname_Text.configure(fg='black')
            description_Text.delete(0, END)
            description_Text.configure(fg='black')
            quantity_Text.delete(0, END)
            quantity_Text.configure(fg='black')
            unitprice_Text.delete(0, END)
            unitprice_Text.configure(fg='black')
            rackno_Text.delete(0, END)
            rackno_Text.configure(fg='black')

    def saveModifiedNonCommercialItemRecords(self, display_dataWindow, item_Id,
                                             itemname_Text,
                                             description_Text,
                                             quantity_Text,
                                             unitprice_Text,
                                             rackno_Text,
                                             infoLabel, print_button, local_centerText):
        print("saveModifiedNonCommercialItemRecords - start")
        subdir_commercialstock = "..\\Library_Stock\\" + local_centerText.get() + "\\NonCommercial_Stock"
        filename = subdir_commercialstock + "\\noncommercial_stock.xlsx"
        wb_obj = openpyxl.load_workbook(filename)
        sheet_obj = wb_obj.active
        total_records = self.obj_commonUtil.totalrecords_excelDataBase(filename)
        for iLoop in range(1, total_records + 1):
            # if member id matches
            # over-write the respective records
            # save the file
            print("database id ->", sheet_obj.cell(row=iLoop + 1, column=2).value, " item id :", item_Id.get())
            if sheet_obj.cell(row=iLoop + 1, column=2).value == item_Id.get():
                print("condition is true")
                sheet_obj.cell(row=iLoop + 1, column=3).value = itemname_Text.get()
                sheet_obj.cell(row=iLoop + 1, column=6).value = description_Text.get()
                sheet_obj.cell(row=iLoop + 1, column=4).value = quantity_Text.get()
                sheet_obj.cell(row=iLoop + 1, column=5).value = unitprice_Text.get()
                sheet_obj.cell(row=iLoop + 1, column=16).value = rackno_Text.get()
                sheet_obj.cell(row=iLoop + 1, column=17).value = local_centerText.get()

                wb_obj.save(filename)
                infoLabel.configure(fg='green')
                infoLabel['text'] = "Record has been successfully modified for Item Id :" + item_Id.get()
                print_button.configure(state=DISABLED, bg='light grey')
                break
        print("saveModifiedCommercialItemRecords - end")