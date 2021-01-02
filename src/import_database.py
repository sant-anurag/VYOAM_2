"""
# Copyright 2020 by Vihangam Yoga Karnataka.
# All rights reserved.
# This file is part of the Vihangan Yoga Operations of Ashram Management Software Package(VYOAM),
# and is released under the "VY License Agreement". Please see the LICENSE
# file that should have been included as part of this package.
# Vihangan Yoga Operations  of Ashram Management Software
# File Name : import_database.py
# Developer : Sant Anurag Deo
# Version : 2.0
"""

from app_defines import *
from app_common import *
from init_database import *
from app_thread import *

from tkinter import filedialog


class ImportDatabase:
    # constructor for Library class
    def __init__(self, master):
        print("constructor called for noncommercial edit ")
        self.obj_commonUtil = CommonUtil()
        self.dateTimeOp = DatetimeOperation()
        self.file_for_import = ''
        self.import_database_dialog(master)

    def import_database_dialog(self, master):
        import_window = Toplevel(master)
        import_window.title("Import Member Database ")
        import_window.geometry('770x250+120+40')
        import_window.configure(background='wheat')
        import_window.resizable(width=False, height=False)

        # delete "X" button in window will be not-operational
        import_window.protocol('WM_DELETE_WINDOW', self.obj_commonUtil.donothing)

        imageFrame = Frame(import_window, width=65, height=60,
                           bg="wheat")
        canvas_width, canvas_height = 60, 60
        canvas = Canvas(imageFrame, width=canvas_width, height=canvas_height, highlightthickness=0)
        myimage = ImageTk.PhotoImage(PIL.Image.open("..\\Images\\a_wheat.png").resize((60, 60)))
        canvas.create_image(0, 0, anchor=NW, image=myimage)
        imageFrame.grid(row=0, column=0, pady=2)
        canvas.grid(row=0, column=0)

        infoFrame = Frame(import_window, width=300, height=100, bd=6, relief='ridge',
                          bg="wheat")
        infoFrame.grid(row=2, column=0, padx=80, pady=10, sticky=W)
        infoLabel = Label(infoFrame, text="Press <Select> to choose file", width=65,
                          justify='center', font=('arial narrow', 13, 'bold'),
                          bg='snow', fg='green', state=NORMAL)
        infoLabel.grid(row=0, column=0)

        heading = Label(import_window, text="Import Member Database",
                        font=('times new roman', 20, 'bold'),
                        bg="wheat")

        heading.grid(row=1, column=0)

        buttonFrame = Frame(import_window, width=100, height=100, bd=4, relief='ridge', bg="light yellow")
        buttonFrame.grid(row=3, column=0, padx=10, pady=5, columnspan=3)

        cancel = Button(buttonFrame, text="Close", fg="Black", command=import_window.destroy,
                        font=NORM_FONT, width=12, bg='light cyan')
        search_result = partial(self.browseFiles, infoLabel)
        submit = Button(buttonFrame, text="Select", fg="Black", command=search_result,
                        font=NORM_FONT, width=12, bg='light cyan')
        submit.grid(row=0, column=0)
        import_btn = Button(buttonFrame, text="Import", fg="Black", command=self.import_database,
                            font=NORM_FONT, width=12, bg='light cyan')
        import_btn.grid(row=0, column=1)

        # create a Close Button and place into the import_window window

        cancel.grid(row=0, column=3)
        # ---------------------------------Button Frame End----------------------------------------

        import_window.bind('<Return>', lambda event=None: submit.invoke())
        import_window.bind('<Alt-c>', lambda event=None: cancel.invoke())

        import_window.focus()
        import_window.grab_set()
        mainloop()

    # Function for opening the
    # file explorer window
    def browseFiles(self, label_file_explorer):
        filename = filedialog.askopenfilename(initialdir="/",
                                              title="Select a File",
                                              filetypes=(("Excel files",
                                                          "*.xlsx*"),
                                                         ("all files",
                                                          "*.*")))

        # Change label contents
        if filename == "":
            text_label = "No File Selected !!!"
        else:
            text_label = filename
        label_file_explorer.configure(text=text_label)
        self.file_for_import = filename

    def import_database(self):
        print("Starting Import of Web Database !!!")
        # opening the source excel file

        wb1 = openpyxl.load_workbook(self.file_for_import)
        ws1 = wb1.active

        # opening the destination excel file
        wb2 = openpyxl.load_workbook(PATH_MEMBER)
        ws2 = wb2.active
        total_destination_records = self.obj_commonUtil.totalrecords_excelDataBase(PATH_MEMBER)
        total_source_records = self.obj_commonUtil.totalrecords_excelDataBase(self.file_for_import)
        print("Destination Records :", total_destination_records, "Importing >>", total_source_records, " Records")
        # new imported records will be updated after the existing member records in database
        # copying the cell values from source
        # excel file to destination excel file
        new_serial_number = total_destination_records + 1
        for row_source in range(1, total_source_records + 1):
            print("Record < ", row_source, " > Imported")
            for column_source in range(1, 29):
                # reading cell value from source excel file
                destination_cell = ws2.cell(row=new_serial_number + 1, column=column_source)
                source_cell = ws1.cell(row=row_source + 1, column=column_source)
                # writing the read value to destination excel file
                if column_source == 1:
                    cell_value = str(new_serial_number)
                elif column_source == 3 or column_source == 16 or \
                        column_source == 17 or column_source == 18 or column_source == 19:
                    cell_value = "Not Available"
                else:
                    cell_value = str(source_cell.value)

                ws2.cell(row=new_serial_number + 1, column=column_source).font = Font(size=12, name='Times New Roman',
                                                                                      bold=False)
                ws2.cell(row=new_serial_number + 1, column=column_source).alignment = Alignment(horizontal='center',
                                                                                                vertical='center',
                                                                                                wrapText=True)
                destination_cell.value = cell_value
            new_serial_number = new_serial_number + 1

            # saving the destination excel file
        wb2.save(PATH_MEMBER)
        self.obj_commonUtil.update_totalMemberRecords()
        print("Import of Web Database  Completed Successfully!!!")
