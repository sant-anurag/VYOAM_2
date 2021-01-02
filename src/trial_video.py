from tkinter import *
root = Tk()
root.geometry('{}x{}+{}+{}'.format(200, 100, 100,100))
root.tk.call('tk', 'scaling', 2.0)
labelToDisplay = Label(root,text = "Hello India is country of all", fg = "RED")
labelToDisplay.grid(row =1, column = 1)
mainloop()