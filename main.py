import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl


def load_data(treeview):
    # Load data from Excel into the Treeview
    path = "C:/Users/PCA/Mozar GUI Project/People.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active


    list_values = list(sheet.values)
    treeview.delete(*treeview.get_children())
    for info_name in list_values[0]:
        treeview.heading(info_name, text=info_name)


    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)


def show_main_window():
    # Setup and display the main application window with treeview and input forms.
    root.deiconify()
    root.title("Bilang-bilang Service Employee's Record")
    root.geometry('650x500')
   
    for widget in root.winfo_children():
        widget.destroy()

    frame = ttk.Frame(root)
    frame.pack(fill="both", expand=True)
   
    treeframe = ttk.Frame(frame)
    treeframe.pack(fill="both", expand=True, side="right", padx=20, pady=20)
   
    treescroll = ttk.Scrollbar(treeframe)
    treescroll.pack(side="right", fill="y")

    treeview = ttk.Treeview(treeframe, columns=("Name", "Age", "Work Time", "Employment Status"), show="headings", yscrollcommand=treescroll.set)
    treeview.column("Name", width=100)
    treeview.column("Age", width=50)
    treeview.column("Work Time", width=100)
    treeview.column("Employment Status", width=120)
    treeview.pack(fill="both", expand=True)
    treescroll.config(command=treeview.yview)


    load_data(treeview)


    input_frame = ttk.Frame(frame)
    input_frame.pack(fill="both", expand=True, side="right", padx=20, pady=20)


    ttk.Label(input_frame, text="Name:").pack()
    name_entry = ttk.Entry(input_frame)
    name_entry.pack()


    ttk.Label(input_frame, text="Age:").pack()
    age_spinbox = ttk.Spinbox(input_frame, from_=18, to=100)
    age_spinbox.pack()


    ttk.Label(input_frame, text="Work Time:").pack()
    work_combobox = ttk.Combobox(input_frame, values=["8 hours", "5 hours", "On leave"])
    work_combobox.current(0)
    work_combobox.pack()


    ttk.Label(input_frame, text="Employment Status:").pack()
    employed_var = tk.BooleanVar()
    check_button = ttk.Checkbutton(input_frame, text="Employed", variable=employed_var)
    check_button.pack()

    insert_button = ttk.Button(input_frame, text="Insert",
                               command=lambda: insert_row(name_entry, age_spinbox, work_combobox, employed_var, check_button, treeview))
    insert_button.pack()


    logout_button = ttk.Button(root, text="Logout", command=logout)
    logout_button.pack(pady=10)


def insert_row(name_entry, age_spinbox, work_combobox, employed_var, check_button, treeview):
    # Function to insert a new row into the Excel and update the treeview
    name = name_entry.get()
    age = age_spinbox.get()
    work_status = work_combobox.get()
    employment_status = "Employed" if employed_var.get() else "Unemployed"

    path = "C:/Users/PCA/Mozar GUI Project/People.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    row_values = [name, age, work_status, employment_status]
    sheet.append(row_values)
    workbook.save(path)

    treeview.insert('', tk.END, values=row_values)
    name_entry.delete(0, "end")
    age_spinbox.delete(0, "end")
    work_combobox.set("8 hours")
    check_button.state(["!selected"])


def logout():
    # Handle logout and show login window again.
    root.withdraw()
    show_login_window()


def show_login_window():
    # Create and display the login window.
    global login_window, username_entry, password_entry


    login_window = tk.Toplevel()
    login_window.title("Mozar Project: Bilang-bilang Service")
    login_window.geometry('500x400')


    username_label = ttk.Label(login_window, text="Username:")
    username_label.pack(pady=(20, 5))


    username_entry = ttk.Entry(login_window)
    username_entry.pack(pady=5)
    username_entry.bind("<Return>", lambda event: verify_credentials())


    password_label = ttk.Label(login_window, text="Password:")
    password_label.pack(pady=(5, 5))


    password_entry = ttk.Entry(login_window, show="*")
    password_entry.pack(pady=5)
    password_entry.bind("<Return>", lambda event: verify_credentials())


    login_button = ttk.Button(login_window, text="Login", command=verify_credentials)
    login_button.pack(pady=20)


def verify_credentials():
    # Verify username and password against some credentials.
    username = username_entry.get()
    password = password_entry.get()
    if username == "owner" and password == "manpower":
        messagebox.showinfo("Login Successful", "You have successfully logged in.")
        login_window.destroy()
        show_main_window()
    else:
        messagebox.showerror("Login Failed", "Invalid username or password")


root = tk.Tk()
root.withdraw()
style = ttk.Style()
style.theme_use('clam')
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")


show_login_window()
root.mainloop()