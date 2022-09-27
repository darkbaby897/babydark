from tkinter import *
from tkinter import Label, Button, Toplevel, Radiobutton

def booknow():
    root.withdraw()
    global booking_window

    booking_window = Toplevel(root)
    booking_window.title("Booking Now")
    booking_window.geometry("580x360")
    booking_window.iconbitmap('images/logo.ico')
    booking_window.configure(bg='#454545')

    start_line = Label(booking_window, text='',width=100, bg='black', height=2)
    start_line.pack()

    go_back = Button(booking_window, text="<<< Go Back", command=goback0, font='calibri, 14', cursor='hand2', bg='black', fg='white', activebackground='black', borderwidth=0)
    go_back.place(x=0, y=0)

    end_line = Label(booking_window, text='',width=100, bg='black', height=2)
    end_line.pack(side='bottom')

    v = IntVar()

    bus_select0 = Radiobutton(booking_window, text='Solan-to-Nalagarh', variable=v, value=1)
    bus_select0.pack(anchor=W)

    bus_select1 = Radiobutton(booking_window, text='nalagarh-to-solan', variable=v, value=2)
    bus_select1.pack(anchor=W)


    bus_select1 = Radiobutton(booking_window, text='Solan-to-Shimla', variable=v, value=3)
    bus_select1.pack(anchor=W)



def yourbooking():
    root.withdraw()
    global your_booking_window

    your_booking_window = Toplevel(root)
    your_booking_window.title("Your Bookings")
    your_booking_window.geometry("580x360")
    your_booking_window.iconbitmap('images/logo.ico')
    your_booking_window.configure(bg='#454545')

    start_line = Label(your_booking_window, text='',width=100, bg='black', height=2)
    start_line.pack()

    go_back = Button(your_booking_window, text="<<< Go Back", command=goback1, font='calibri, 14', cursor='hand2', bg='black', fg='white', activebackground='black', borderwidth=0)
    go_back.place(x=0, y=0)

    end_line = Label(your_booking_window, text='',width=100, bg='black', height=2)
    end_line.pack(side='bottom')

def goback0():
    booking_window.withdraw()
    root.deiconify()

def goback1():
    your_booking_window.withdraw()
    root.deiconify()


root = Tk()
root.title("Bus Booking System")
root.geometry('580x360') # width = 580, height = 360
root.iconbitmap('images/logo.ico')
root.configure(bg='#454545')

label1 = Label(root, text="This Bus Online Booking System", font='calibri, 17', width=100, bg='black', fg='white')
label1.pack()

photo1 = PhotoImage()

book_now = Button(root, text="Book Now", font='calibri, 17', command=booknow, borderwidth=0, bg='#454545', activebackground='#454545', cursor='hand2', fg='white')
book_now.pack(pady=12)

your_booking = Button(root, text="Your Bookings", font='calibri, 17', command=yourbooking, borderwidth=0, bg='#454545', activebackground='#454545', cursor='hand2', fg='white')
your_booking.pack(pady=10)

label2 =Label(root, text='', width=100, bg='black', height=2)
label2.pack(side='bottom')

root.mainloop()
