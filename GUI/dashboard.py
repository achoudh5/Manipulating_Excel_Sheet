from tkinter import *
#import <api>
from PIL import ImageTk, Image
root = Tk()

root.title("Check Validity")

canvas1 = Canvas(root, width = 400, height = 300)
canvas1.pack()
# The Canvas is your display where you can place items, such as entry boxes, buttons, charts and more.
# You can control the dimensions of your Canvas by changing the width and height values:

my_img = ImageTk.PhotoImage(Image.open("../images/Screen Shot 2020-10-03 at 1.19.35 AM.png"))
my_label = Label(image=my_img)
my_label.pack()

def myClick():
    global counter
    counter+=1
    #s= <api call happening, this is already in place, check import>
    mylabel = Label(root, text=str(counter)+' '+'You Entered:'\
                               +input_box.get()+'\nRequestor'+' '+str(s)+'\n')
    mylabel.pack()
    input_box.delete(0, END)

#counts how many time you used this tool to check validity
counter=0

#An input_box can be used to get the user’s input
input_box = Entry(root, width=40, bg='#30CFBB', fg='black', selectforeground='red',selectbackground='white',borderwidth=5)
canvas1.create_window(250,140,window=input_box)


label2 = Label(root, text='*The database is case sensitive!*')
label2.config(font=('helvetica', 16))
canvas1.create_window(250, 180, window=label2)
input_box.pack()
#input_box.insert() #placeholder text in the box

#validity_Button = Button(root, text="Checking validity of the Requestor!", padx=70, pady=70, command=lambda: myClick(bluegroup_api.Bluegroup(input_box.get())))
validity_Button = Button(text="Checking validity!",command= myClick, bg='brown', fg='white', font=('helvetica', 9, 'bold'))
canvas1.create_window(280, 180, window=validity_Button)
validity_Button.pack()

#Exit Button
#button_quit = Button(root, text="Exit Program", command=root.quit)
#button_quit.pack()

root.mainloop()