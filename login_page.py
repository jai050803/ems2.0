import customtkinter as ctk
import tkinter as tk
from tkinter import *
from tkinter import messagebox
import mysql.connector
import bcrypt

app = ctk.CTk()
app.title("Welcome to EMS App")
app.geometry("700x600")
app.config(bg="#001220")

font1 = ("helvetica", 40, "bold")
font2 = ("Arial", 20)
font3 = ("Arial", 12)

# Connect to MySQL database
try:
    conn = mysql.connector.connect(
        host="localhost",
        user="jai",
        password="jai@2301420045",
        database="user_credentials"
    )
    cursor = conn.cursor()

        
except mysql.connector.Error as err:
    messagebox.showerror("MySQL Error", f"Error: {err}")
    
def signup():
    username = username_entry.get()
    password = password_entry.get()
    confirm_password = confirm_password_entry.get()

    if password == confirm_password:
        hashed_password = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())
        
        try:
            # Insert user data into the database
            cursor.execute("INSERT INTO users (username, password) VALUES (%s, %s)", (username, hashed_password))
            conn.commit()
            messagebox.showinfo("Success", "Account created successfully!")

        except mysql.connector.Error as err:
            messagebox.showerror("MySQL Error", f"Error: {err}")
        
    else:
        messagebox.showerror("Error", "Password and Confirm Password do not match")

def go_to_login():
    app.destroy()  # Close the current window
    # Add code to open the login window or navigate to the login page

frame1 = ctk.CTkFrame(app, bg_color="#001220", width=350, height= 600)
frame1.place(x=0,y=0)

frame2 = ctk.CTkFrame(app, bg_color="#001220", fg_color="#000000", width=350, height= 600)
frame2.place(x=350,y=0)

img1 = PhotoImage(file = "blue.png")
img1_label = Label(frame1, image = img1, bg = "#001220", width=350, height= 600)
img1_label.place(x=0,y=0)


signup_label = ctk.CTkLabel(frame2, text = "Sign Up", text_color = "#fff", font=font1)
signup_label.place(x=110,y=20)

username_entry = ctk.CTkEntry(frame2, text_color = "#fff",fg_color="#001a2e",bg_color="#121111",
                            border_color="#004780", border_width=3, placeholder_text="username",
                            placeholder_text_color="#a3a3a3", width=200, height=50, font=font2)
username_entry.place(x=80,y=150)

password_entry = ctk.CTkEntry(frame2, text_color = "#fff",fg_color="#001a2e",bg_color="#121111",
                            border_color="#004780", border_width=3, placeholder_text="password",
                            placeholder_text_color="#a3a3a3", width=200, height=50, font=font2, show= "*")
password_entry.place(x=80,y=240)

confirm_password_entry = ctk.CTkEntry(frame2, text_color = "#fff",fg_color="#001a2e",bg_color="#121111",
                            border_color="#004780", border_width=3, placeholder_text="Confirm password",
                            placeholder_text_color="#a3a3a3", width=200, height=50, font=font2, show= "*")
confirm_password_entry.place(x=80,y=330)

signup_button = ctk.CTkButton(frame2, font=font2, text_color="#fff", text='sign up', fg_color="#00965d", 
                            hover_color='#006e44',bg_color="#121111", cursor="hand2",corner_radius=5, width=120)
signup_button.place(x=115, y= 430)

login_label = ctk.CTkLabel(frame2, font=font3, text='Already have an account?', text_color="#fff")
login_label.place(x=100,y=470)

login_button = ctk.CTkButton(frame2, font=font2, text_color="#fff", text='Login Here', fg_color="#00965d", 
                            hover_color='#006e44',bg_color="#121111", cursor="hand2",corner_radius=5, width=120)
login_button.place(x=115, y= 510)

app.mainloop()