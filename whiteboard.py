import tkinter as tk
from tkinter.colorchooser import askcolor
from tkinter import ttk
from tkinter import *
from tkinter import messagebox
from main import show_secondary_page
import re

#functions 
common_passwords = [
    "password", "123456", "12345678", "qwerty", "abc123", "monkey", "1234567",
    "letmein", "trustno1", "dragon", "baseball", "111111", "iloveyou",
    "master", "sunshine", "ashley", "bailey", "shadow", "superman", "qazwsx"
]

def is_password_strong(password):
    
    # Check the password against the criteria
    if len(password) < 8:
        return False, "Password must be at least 8 characters long."
    if not re.search("[a-z]", password):
        return False, "Password must include lowercase letters."
    if not re.search("[A-Z]", password):
        return False, "Password must include uppercase letters."
    if not re.search("[0-9]", password):
        return False, "Password must include numbers."
    if not re.search("[!@#$%^&*(),.?\":{}|<>]", password):
        return False, "Password must include special characters."
    if password in common_passwords:
        return False, "Please avoid common passwords."
    return True, "Password is strong."
# Placeholder clear function
def clear_placeholder(event):
    if event.widget == user and user.get() == 'username':
        user.delete(0, tk.END)
    elif event.widget == password and password.get() == 'Password':
        password.delete(0, tk.END)
        password.config(show='*')
    elif event.widget == authentication_key and authentication_key.get() == 'Authentication Key':
        authentication_key.delete(0, tk.END)

# Restore placeholder if empty
def add_placeholder(event):
    if event.widget == user and not user.get():
        user.insert(0, 'username')
    elif event.widget == password and not password.get():
        password.config(show='')
        password.insert(0, 'Password')
    elif event.widget == authentication_key and not authentication_key.get():
        authentication_key.insert(0, 'Authentication Key')
def signup():
    username_input = user.get()
    password_input = password.get()
    auth_key_input = authentication_key.get()
    
    if user.get() == 'username' or password.get() == 'Password' or authentication_key.get() == 'Authentication Key':
        messagebox.showwarning("Warning", "Please enter your details!")
    
    # Check if the password is strong
    password_strength, message = is_password_strong(password_input)
    if not password_strength:
        messagebox.showwarning("Weak Password", message)
        return

    # Check if the password and authentication key match
    if password_input != auth_key_input:
        messagebox.showwarning("Mismatch", "The password and authentication key do not match.")
        return
    else:
        messagebox.showinfo("Success", "Login Successful!") 
        root.destroy() # Close the login window
        show_secondary_page()

root = Tk()
root.title('EMS - A Business Intelligence tool')
root.geometry('925x550+300+200')
root.configure(bg="#fff")
root.resizable(False, False)

image_icon = PhotoImage(file="logo.png")
root.iconphoto(False,image_icon)

img = PhotoImage(file='login.png')
Label(root,image=img,bg='white').place(x=45,y=105)

frame = Frame(root,width=350,height=450,bg="white")
frame.place(x=480,y=70)

heading = Label(frame, text= "User Login", fg="#57a1f8", bg="white", font=('Microsoft YaHei UI Light', 23, 'bold'))
heading.place(x=130,y=5)

user = Entry(frame, text='login', width=90, fg='black', border=0, bg="white", font=('Microsoft YaHei UI Light', 11))
user.place(x=10,y=80)
user.insert(0, 'username')
Frame(frame, width=405, height=2, bg='black').place(x=10, y=107)

password = Entry(frame, text='Password', width=90, fg='black', border=0, bg="white", font=('Microsoft YaHei UI Light', 11))
password.place(x=10,y=155)
password.insert(0, 'Password')
Frame(frame, width=405, height=2, bg='black').place(x=10, y=187)

authentication_key = Entry(frame, text='Authentication Key', width=90, fg='black', border=0, bg="white", font=('Microsoft YaHei UI Light', 11))
authentication_key.place(x=10,y=235)
authentication_key.insert(0, 'Authentication Key')
Frame(frame, width=405, height=2, bg='black').place(x=10, y=267)

# Binding entry widgets to clear and restore placeholder text
user.bind("<FocusIn>", clear_placeholder)
user.bind("<FocusOut>", add_placeholder)
password.bind("<FocusIn>", clear_placeholder)
password.bind("<FocusOut>", add_placeholder)
authentication_key.bind("<FocusIn>", clear_placeholder)
authentication_key.bind("<FocusOut>", add_placeholder)

Button(frame, width=39, padx=5, pady=5, text='Log In', bg='#57a1f8', fg='white', border=0, command=signup, font=('Microsoft YaHei UI Light', 11)).place(x=10, y=295)

label = Label(frame, text='Don\'t have a secruity key?',fg='black',bg='white', font=('Microsoft YaHei UI Light', 11))
label.place(x=55, y=345)

click_here = Button(frame, width=8, text='click here', border=0, bg='white', cursor='hand2', fg='#57a1f8', font=('Microsoft YaHei UI Light', 11))
click_here.place(x=230, y=343)
root.mainloop()
