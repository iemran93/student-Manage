import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, Menu
from PIL import ImageTk
from awesometkinter.bidirender import add_bidi_support, render_text
import pendulum
from random import choice
import glob
import pandas as pd


# connection
class ConnectSQL():
    def __init__(self, database_file="students.db"):
        self.databas_file = database_file

    def connect(self):
        connection = sqlite3.connect(self.databas_file)
        return connection


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Student Management")
        width = int((self.winfo_screenwidth() / 2) - 350)
        height = int((self.winfo_screenheight() / 2) - 300)
        self.geometry(f"700x600+{width}+{height}")
        self.resizable(False, False)
        date = pendulum.now().to_day_datetime_string()[:-8]

        self.name_var = tk.StringVar()
        self.ac_var = tk.IntVar(value=2021)
        self.class_var = tk.StringVar()
        self.date_var = tk.StringVar(value=date)
        # combobox
        self.items = ["مكتب خدمة المجتمع", "الاشراف", "الادارة"]
        self.place_var = tk.StringVar(value=self.items[0])
        # menubar
        menubar = Menu(self)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="Insert", command=self.insert_data)
        menubar.add_cascade(label="File", menu=filemenu)
        self.configure(menu=menubar)

        self.widgets()
        self.layout()

    def widgets(self):
        # date label
        self.label_date = tk.Label(
            self, textvariable=self.date_var, relief="ridge")
        # header
        self.label_header = tk.Label(self, text="ادخال طالب جديد",
                                     bg="orange", font="Tahoma 12", width=30)
        # name
        self.entry_name = tk.Entry(self, justify="right",
                                   textvariable=self.name_var)
        add_bidi_support(self.entry_name)
        self.name_label = tk.Label(self, text="الاسم")
        # ac number
        self.entry_ac = tk.Entry(
            self, justify="right", textvariable=self.ac_var)
        self.ac_label = tk.Label(self, text="الرقم الأكاديمي")
        # class
        self.entry_class = tk.Entry(
            self, justify="right", textvariable=self.class_var)
        self.class_label = tk.Label(self, text="الشعبة")
        # buttons
        self.button = tk.Button(self, text="ادخال",
                                command=self.add_student)
        self.button_show = tk.Button(
            self, text="عرض", command=self.show_window)
        self.button_search = tk.Button(
            self, text="بحث", command=self.search_window)
        self.button_all = tk.Button(
            self, text="على الكل", command=self.apply_all, background="lightblue")
        # image
        self.logo_img = ImageTk.PhotoImage(file="assets\database_logo.png")
        self.img_label = tk.Label(self, image=self.logo_img)
        self.img_label.image = self.logo_img
        # combobox
        self.combo = ttk.Combobox(
            self, values=self.items, textvariable=self.place_var)
        self.combo_label = tk.Label(
            self, text="مكان الخدمة")

    def layout(self):
        # date
        self.label_date.place(relx=0.95, rely=0.05, anchor="ne")
        # header
        self.label_header.place(relx=0.5, rely=0.33, anchor="center")
        # name
        self.entry_name.place(relx=0.5, rely=0.4, anchor="center")
        self.name_label.place(relx=0.62, rely=0.4, anchor="center")
        # ac number
        self.entry_ac.place(relx=0.5, rely=0.5, anchor="center")
        self.ac_label.place(relx=0.65, rely=0.5, anchor="center")
        # class
        self.entry_class.place(relx=0.2, rely=0.5, anchor="center")
        self.class_label.place(relx=0.32, rely=0.5, anchor="center")
        # buttons
        self.button.place(relx=0.5, rely=0.7, anchor="center")
        self.button_show.place(relx=0.4, rely=0.7, anchor="center")
        self.button_search.place(relx=0.3, rely=0.7, anchor="center")
        self.button_all.place(relx=0.65, rely=0.7, anchor="center")
        # image
        self.img_label.place(relx=0.5, rely=0, anchor="n")
        # combo
        self.combo.place(relx=0.5, rely=0.6, anchor="center")
        self.combo_label.place(relx=0.65, rely=0.6, anchor="center")

    def add_student(self):
        try:
            if self.class_var.get() and self.name_var.get():
                name = self.name_var.get()
                ac_id = self.ac_var.get()
                place = self.place_var.get()
                class_ = self.class_var.get()
                connection = ConnectSQL().connect()
                cursor = connection.cursor()
                cursor.execute(
                    "INSERT INTO students (AcNum,Name,Class,Place) VALUES(?,?,?,?)", (ac_id, name, class_, place))
                connection.commit()
                cursor.close()
                connection.close()
            else:
                messagebox.showinfo("Showinfo", "ادخل اسم الطالب والشعبة")
        except sqlite3.IntegrityError:
            messagebox.showinfo("Showinfo", "الرقم الأكاديمي موجود")

    def show_window(self):
        show_window = ShowStudents()

    def search_window(self):
        search_window = SearchStudents(self.ac_var, self.items)

    def apply_all(self):
        answer = messagebox.askquestion(
            "Are you sure?", "تطبيق المكان العشوائي على كل الطلبة؟")
        if answer == "yes":
            # Apply randon place choice for all students
            connection = ConnectSQL().connect()
            cursor = connection.cursor()
            cursor.execute("SELECT * FROM students")
            result = cursor.fetchall()
            for row in result:
                place_choice = choice(self.items)
                ac_num = row[0]
                cursor.execute("UPDATE students SET place=? WHERE AcNum=?",
                               (place_choice, ac_num))
            connection.commit()
            cursor.close()
            connection.close()

    def insert_data(self):
        if "data.xlsx" in glob.glob("*.xlsx"):
            df = pd.read_excel("data.xlsx")
            conn = ConnectSQL().connect()
            df.to_sql("students", conn, if_exists="replace", index=False)
            conn.close()
            messagebox.showinfo("Done", "تم ادخال البيانات")


class Treeview(ttk.Treeview):
    def __init__(self, parent):
        super().__init__(parent, columns=(
            "first", "second", "third", "fourth"), show="headings")
        s = ttk.Style()
        s.theme_use('clam')
        s.configure('Treeview.Heading', background="green")
        self.heading("first", text="Ac Number")
        self.heading("second", text="Name")
        self.heading("third", text="Class")
        self.heading("fourth", text="Plcae")

        self.pack(fill="both", expand=True)


class ShowStudents(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Students Table")
        self.geometry("800x600")
        # self.resizable(False, False)

        self.table = Treeview(self)
        # widgets
        self.button_class = tk.Button(
            self, text="حسب الصف", background="lightblue", command=self.choose_class)
        self.button_place = tk.Button(
            self, text="حسب المكان", background="lightblue", command=self.choose_place)
        # layout
        self.button_class.pack(side="right")
        self.button_place.pack(side="right")
        # class var

        self.load_data()

    def load_data(self, class_="all", place=""):
        connection = ConnectSQL().connect()
        # load all data
        if class_ == "all":
            if not place:
                result = connection.execute("SELECT * FROM students")
                for row_data in result:
                    self.table.insert(parent="", index=tk.END, values=row_data)
                connection.close()
            else:
                # clear table
                for item in self.table.get_children():
                    self.table.delete(item)
                result = connection.execute("SELECT * FROM students Where place=?",
                                            (place,))
                for row_data in result:
                    self.table.insert(parent="", index=tk.END, values=row_data)
                connection.close()
        else:
            # clear table
            for item in self.table.get_children():
                self.table.delete(item)
            # load new data
            result = connection.execute("SELECT * FROM students WHERE class=?",
                                        (class_,))
            for row_data in result:
                self.table.insert(parent="", index=tk.END, values=row_data)
            connection.close()

    def choose_class(self):
        top = tk.Toplevel(self)
        top.geometry("300x75+100+100")
        top.title("Select Class")
        class_ent = tk.Entry(top, width=15, justify="right")
        btn = tk.Button(top, text="اختيار")
        class_ent.pack()
        btn.pack()
        btn.bind("<Button>", lambda event: (self.load_data(
            class_ent.get()), top.destroy()))

    def choose_place(self):
        top = tk.Toplevel(self)
        top.geometry("300x75+100+100")
        top.title("Select Place")
        place_ent = tk.Entry(top, width=15, justify="right")
        add_bidi_support(place_ent)
        btn = tk.Button(top, text="اختيار")
        place_ent.pack()
        btn.pack()
        btn.bind("<Button>", lambda event: (self.load_data("all",
                                                           place_ent.get()), top.destroy()))


class SearchStudents(tk.Tk):
    def __init__(self, ac_var, items):
        super().__init__()
        self.title("Students Search")
        self.geometry("800x600")
        self.resizable(False, False)
        self.places = items

        # ac number
        self.ac_id = ac_var.get()

        self.table = Treeview(self)
        self.table.bind("<<TreeviewSelect>>", self.select_item)
        self.table.bind("<F1>", self.edit_student)
        self.load_data()

    def load_data(self):
        connection = ConnectSQL().connect()
        cursor = connection.cursor()
        result = cursor.execute(
            "SELECT * FROM students WHERE AcNum = ?", (self.ac_id,))
        for row_data in result:
            self.table.insert(parent="", index=tk.END, values=row_data)
        cursor.close()
        connection.close()

    def select_item(self, event):
        for i in self.table.selection():
            self.item_selected = self.table.item(i)['values']

    def edit_student(self, event):
        edit_student = EditStudent(self.item_selected, self.places)
        self.destroy()


class EditStudent(tk.Tk):
    def __init__(self, item, places):
        super().__init__()
        self.title("Edit Student")
        self.geometry("400x200")
        self.resizable(False, False)
        # data
        self.ac_num = item[0]
        self.name = item[1]
        self.class_ = item[2]
        self.place = item[3]
        self.places = places
        self.combo_var = tk.StringVar(value=self.place)
        combo_index = self.places.index(self.place)
        # entry widgets
        self.ac_ent = tk.Entry(self, justify="right")
        self.ac_ent.insert(0, self.ac_num)
        self.ac_ent.configure(state="readonly")
        self.ac_ent.pack(expand=True)

        self.name_ent = tk.Entry(self, justify="right")
        self.name_ent.pack(expand=True)
        add_bidi_support(self.name_ent)
        self.name_ent.insert(0, render_text(self.name))

        self.class_ent = tk.Entry(self, justify="right")
        self.class_ent.pack(expand=True)
        self.class_ent.insert(0, self.class_)

        self.combo = ttk.Combobox(self)
        self.combo.pack(expand=True)
        self.combo.config(values=self.places)
        self.combo.current(combo_index)

        button_edit = tk.Button(
            self, text="تعديل", command=self.edit_student).pack(expand=True)
        button_delete = tk.Button(
            self, text="حذف", command=self.delete_student).pack(expand=True)

    def delete_student(self):
        connection = ConnectSQL().connect()
        cursor = connection.cursor()
        cursor.execute("DELETE from students WHERE AcNum = ?", (self.ac_num,))
        connection.commit()
        cursor.close()
        connection.close()
        messagebox.showinfo("ShowInfo", "تم الحذف")
        self.status_finish = True
        self.destroy()

    def edit_student(self):
        connection = ConnectSQL().connect()
        cursor = connection.cursor()
        cursor.execute("UPDATE students SET name=?, class=?, place=? WHERE AcNum = ?",
                       (self.name_ent.get(),
                        self.class_ent.get(),
                        self.combo.get(),
                        self.ac_num))
        connection.commit()
        cursor.close()
        connection.close()
        messagebox.showinfo("ShowInfo", "تم التعديل")
        self.status_finish = True
        self.destroy()


if __name__ == '__main__':
    app = App()
    app.mainloop()
