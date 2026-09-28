import openpyxl
import os
import tkinter as tk
from tkinter import messagebox, ttk


FILE_NAME = "electronic_shop.xlsx"


# ===========================
# CREATE EXCEL FILE
# ===========================

def create_excel_file():

    if not os.path.exists(FILE_NAME):

        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Products"

        headers = [
            "Product ID",
            "Product Name",
            "Category",
            "Brand",
            "Price",
            "Quantity",
            "Supplier",
            "Stock Status"
        ]

        sheet.append(headers)

        workbook.save(FILE_NAME)

        print("Excel file created successfully!")


create_excel_file()


# ========================
# LOGIN
# ========================

def login():

    username = username_entry.get()
    password = password_entry.get()

    if username == "admin" and password == "1234":

        messagebox.showinfo(
            "Login Successful",
            "Welcome to Electronic Shop Management System!"
        )

        login_window.withdraw()
        dashboard()

    else:

        messagebox.showerror(
            "Login Failed",
            "Invalid Username or Password"
        )


def clear_login():

    username_entry.delete(0, tk.END)
    password_entry.delete(0, tk.END)


# =======================
# ADD RECORD
# =======================

def add_record():

    add_window = tk.Toplevel()
    add_window.title("Add Product Record")
    add_window.geometry("400x600")


    # ---------------- CANVAS ----------------

    canvas = tk.Canvas(add_window)


    # ---------------- SCROLLBAR ----------------

    scrollbar = tk.Scrollbar(
        add_window,
        orient="vertical",
        command=canvas.yview
    )

    canvas.configure(
        yscrollcommand=scrollbar.set
    )

    scrollbar.pack(
        side="right",
        fill="y"
    )

    canvas.pack(
        side="left",
        fill="both",
        expand=True
    )


    # ---------------- FRAME INSIDE CANVAS ----------------

    form_frame = tk.Frame(canvas)

    canvas_window = canvas.create_window(
        (0, 0),
        window=form_frame,
        anchor="nw"
    )


    # ---------------- UPDATE SCROLL AREA ----------------

    def update_scroll_region(event=None):

        canvas.configure(
            scrollregion=canvas.bbox("all")
        )


    form_frame.bind(
        "<Configure>",
        update_scroll_region
    )


    # ---------------- MAKE FRAME WIDTH EQUAL TO WINDOW ----------------

    def resize_frame(event):

        canvas.itemconfig(
            canvas_window,
            width=event.width
        )


    canvas.bind(
        "<Configure>",
        resize_frame
    )


    # ========================
    # TITLE
    # ========================

    tk.Label(
        form_frame,
        text="ADD PRODUCT RECORD",
        font=("Arial", 14, "bold")
    ).pack(pady=25)


    # =====================
    # PRODUCT ID
    # =====================

    tk.Label(
        form_frame,
        text="Product ID",
        font=("Arial", 8)
    ).pack(pady=(5, 3))

    product_id_entry = tk.Entry(
        form_frame,
        width=35
    )

    product_id_entry.pack()


    # ===========================
    # PRODUCT NAME
    # ===========================

    tk.Label(
        form_frame,
        text="Product Name",
        font=("Arial", 8)
    ).pack(pady=(15, 3))

    product_name_entry = tk.Entry(
        form_frame,
        width=35
    )

    product_name_entry.pack()


    # ===================
    # CATEGORY
    # ===================

    tk.Label(
        form_frame,
        text="Category",
        font=("Arial", 8)
    ).pack(pady=(15, 3))

    category_entry = tk.Entry(
        form_frame,
        width=35
    )

    category_entry.pack()


    # =======================
    # BRAND
    # =======================

    tk.Label(
        form_frame,
        text="Brand",
        font=("Arial", 8)
    ).pack(pady=(15, 3))

    brand_entry = tk.Entry(
        form_frame,
        width=35
    )

    brand_entry.pack()


    # ===================
    # PRICE
    # ===================

    tk.Label(
        form_frame,
        text="Price",
        font=("Arial", 8)
    ).pack(pady=(15, 3))

    price_entry = tk.Entry(
        form_frame,
        width=35
    )

    price_entry.pack()


    # =======================
    # QUANTITY
    # =======================

    tk.Label(
        form_frame,
        text="Quantity",
        font=("Arial", 8)
    ).pack(pady=(15, 3))

    quantity_entry = tk.Entry(
        form_frame,
        width=35
    )

    quantity_entry.pack()


    # ========================
    # SUPPLIER
    # ========================

    tk.Label(
        form_frame,
        text="Supplier",
        font=("Arial", 8)
    ).pack(pady=(15, 3))

    supplier_entry = tk.Entry(
        form_frame,
        width=35
    )

    supplier_entry.pack()


    # ========================
    # STOCK STATUS
    # ========================

    tk.Label(
        form_frame,
        text="Stock Status",
        font=("Arial", 8)
    ).pack(pady=(15, 3))

    stock_status_entry = tk.Entry(
        form_frame,
        width=35
    )

    stock_status_entry.pack()


    # ========================
    # CLEAR FIELDS
    # ========================

    def clear_fields():

        product_id_entry.delete(0, tk.END)
        product_name_entry.delete(0, tk.END)
        category_entry.delete(0, tk.END)
        brand_entry.delete(0, tk.END)
        price_entry.delete(0, tk.END)
        quantity_entry.delete(0, tk.END)
        supplier_entry.delete(0, tk.END)
        stock_status_entry.delete(0, tk.END)


    # =========================
    # SAVE RECORD
    # =========================

    def save_record():

        product_id = product_id_entry.get()
        product_name = product_name_entry.get()
        category = category_entry.get()
        brand = brand_entry.get()
        price = price_entry.get()
        quantity = quantity_entry.get()
        supplier = supplier_entry.get()
        stock_status = stock_status_entry.get()


        if (
            product_id == "" or
            product_name == "" or
            category == "" or
            brand == "" or
            price == "" or
            quantity == "" or
            supplier == "" or
            stock_status == ""
        ):

            messagebox.showerror(
                "Error",
                "Please fill all fields."
            )

            return


        try:

            workbook = openpyxl.load_workbook(FILE_NAME)
            sheet = workbook["Products"]

            sheet.append([
                product_id,
                product_name,
                category,
                brand,
                price,
                quantity,
                supplier,
                stock_status
            ])

            workbook.save(FILE_NAME)

            messagebox.showinfo(
                "Success",
                "Product record added successfully!"
            )

            clear_fields()

        except Exception as e:

            messagebox.showerror(
                "Error",
                f"Something went wrong:\n{e}"
            )


    # =========================
    # BUTTONS
    # =========================

    tk.Button(
        form_frame,
        text="SAVE RECORD",
        width=20,
        height=2,
        command=save_record
    ).pack(pady=25)


    tk.Button(
        form_frame,
        text="CLEAR",
        width=20,
        height=2,
        command=clear_fields
    ).pack(pady=(0, 30))
    
    
# ========================
# VIEW RECORDS
# ========================

def view_records():

    view_window = tk.Toplevel()
    view_window.title("View Records")
    view_window.geometry("950x550")
    view_window.configure(bg="#F4F7FB")


    # ---------------- TITLE ----------------

    title_label = tk.Label(
        view_window,
        text="Electronic Shop - All Records",
        font=("Arial", 18, "bold"),
        bg="#F4F7FB",
        fg="#1F3C88"
    )

    title_label.pack(pady=15)


    # ---------------- TABLE FRAME ----------------

    table_frame = tk.Frame(
        view_window,
        bg="#F4F7FB"
    )

    table_frame.pack(
        fill="both",
        expand=True,
        padx=15,
        pady=10
    )


    # ---------------- COLUMNS ----------------

    columns = (
        "Product ID",
        "Product Name",
        "Category",
        "Brand",
        "Quantity",
        "Price",
        "Supplier",
        "Contact",
        "Date"
    )


    # ---------------- TREEVIEW ----------------

    tree = ttk.Treeview(
        table_frame,
        columns=columns,
        show="headings"
    )


    # ---------------- COLUMN HEADINGS ----------------

    for column in columns:

        tree.heading(
            column,
            text=column
        )


    # ---------------- COLUMN WIDTHS ----------------

    widths = {

        "Product ID": 90,
        "Product Name": 150,
        "Category": 120,
        "Brand": 100,
        "Quantity": 80,
        "Price": 90,
        "Supplier": 130,
        "Contact": 120,
        "Date": 100
    }


    for column in columns:

        tree.column(
            column,
            width=widths[column],
            anchor="center"
        )


    # ---------------- VERTICAL SCROLLBAR ----------------

    vertical_scrollbar = ttk.Scrollbar(
        table_frame,
        orient="vertical",
        command=tree.yview
    )


    # ---------------- HORIZONTAL SCROLLBAR ----------------

    horizontal_scrollbar = ttk.Scrollbar(
        table_frame,
        orient="horizontal",
        command=tree.xview
    )


    tree.configure(
        yscrollcommand=vertical_scrollbar.set,
        xscrollcommand=horizontal_scrollbar.set
    )


    tree.grid(
        row=0,
        column=0,
        sticky="nsew"
    )


    vertical_scrollbar.grid(
        row=0,
        column=1,
        sticky="ns"
    )


    horizontal_scrollbar.grid(
        row=1,
        column=0,
        sticky="ew"
    )


    table_frame.grid_rowconfigure(
        0,
        weight=1
    )

    table_frame.grid_columnconfigure(
        0,
        weight=1
    )


    # ---------------- READ RECORDS FROM EXCEL ----------------

    try:

        workbook = openpyxl.load_workbook(FILE_NAME)
        sheet = workbook["Products"]


        for row in sheet.iter_rows(
            min_row=2,
            values_only=True
        ):

            tree.insert(
                "",
                "end",
                values=row
            )


        workbook.close()


    except Exception as e:

        messagebox.showerror(
            "Error",
            "Unable to read records.\n\n" + str(e)
        )

# ---------------- UPDATE RECORD ----------------

def update_record():

    update_window = tk.Toplevel()
    update_window.title("Update Product Record")
    update_window.geometry("450x650")
    update_window.configure(bg="#F4F7FB")

    # ---------------- SCROLLABLE AREA ----------------

    canvas = tk.Canvas(
        update_window,
        bg="#F4F7FB",
        highlightthickness=0
    )

    scrollbar = tk.Scrollbar(
        update_window,
        orient="vertical",
        command=canvas.yview
    )

    scrollable_frame = tk.Frame(
        canvas,
        bg="#F4F7FB"
    )

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window(
        (0, 0),
        window=scrollable_frame,
        anchor="nw",
        width=430
    )

    canvas.configure(
        yscrollcommand=scrollbar.set
    )

    canvas.pack(
        side="left",
        fill="both",
        expand=True
    )

    scrollbar.pack(
        side="right",
        fill="y"
    )

    # ---------------- MOUSE / TOUCH SCROLL ----------------

    def mouse_scroll(event):
        canvas.yview_scroll(
            int(-1 * (event.delta / 120)),
            "units"
        )

    canvas.bind_all(
        "<MouseWheel>",
        mouse_scroll
    )

    # ---------------- TITLE ----------------

    tk.Label(
        scrollable_frame,
        text="UPDATE PRODUCT RECORD",
        font=("Arial", 18, "bold"),
        bg="#F4F7FB",
        fg="#1F3C88"
    ).pack(pady=20)

    # ---------------- PRODUCT ID ----------------

    tk.Label(
        scrollable_frame,
        text="Enter Product ID",
        font=("Arial", 11),
        bg="#F4F7FB"
    ).pack(pady=(10, 5))

    search_id_entry = tk.Entry(
        scrollable_frame,
        width=30
    )
    search_id_entry.pack()

    # ---------------- FORM FRAME ----------------

    form_frame = tk.Frame(
        scrollable_frame,
        bg="#F4F7FB"
    )

    form_frame.pack(pady=15)

    # ---------------- PRODUCT NAME ----------------

    tk.Label(
        form_frame,
        text="Product Name",
        bg="#F4F7FB"
    ).grid(
        row=0,
        column=0,
        padx=10,
        pady=8
    )

    product_name_entry = tk.Entry(
        form_frame,
        width=25
    )

    product_name_entry.grid(
        row=0,
        column=1
    )

    # ---------------- CATEGORY ----------------

    tk.Label(
        form_frame,
        text="Category",
        bg="#F4F7FB"
    ).grid(
        row=1,
        column=0,
        padx=10,
        pady=8
    )

    category_entry = tk.Entry(
        form_frame,
        width=25
    )

    category_entry.grid(
        row=1,
        column=1
    )

    # ---------------- BRAND ----------------

    tk.Label(
        form_frame,
        text="Brand",
        bg="#F4F7FB"
    ).grid(
        row=2,
        column=0,
        padx=10,
        pady=8
    )

    brand_entry = tk.Entry(
        form_frame,
        width=25
    )

    brand_entry.grid(
        row=2,
        column=1
    )

    # ---------------- PRICE ----------------

    tk.Label(
        form_frame,
        text="Price",
        bg="#F4F7FB"
    ).grid(
        row=3,
        column=0,
        padx=10,
        pady=8
    )

    price_entry = tk.Entry(
        form_frame,
        width=25
    )

    price_entry.grid(
        row=3,
        column=1
    )

    # ---------------- QUANTITY ----------------

    tk.Label(
        form_frame,
        text="Quantity",
        bg="#F4F7FB"
    ).grid(
        row=4,
        column=0,
        padx=10,
        pady=8
    )

    quantity_entry = tk.Entry(
        form_frame,
        width=25
    )

    quantity_entry.grid(
        row=4,
        column=1
    )

    # ---------------- SUPPLIER ----------------

    tk.Label(
        form_frame,
        text="Supplier",
        bg="#F4F7FB"
    ).grid(
        row=5,
        column=0,
        padx=10,
        pady=8
    )

    supplier_entry = tk.Entry(
        form_frame,
        width=25
    )

    supplier_entry.grid(
        row=5,
        column=1
    )

    # ---------------- STOCK STATUS ----------------

    tk.Label(
        form_frame,
        text="Stock Status",
        bg="#F4F7FB"
    ).grid(
        row=6,
        column=0,
        padx=10,
        pady=8
    )

    stock_status_entry = tk.Entry(
        form_frame,
        width=25
    )

    stock_status_entry.grid(
        row=6,
        column=1
    )

    # ---------------- CLEAR FORM ----------------

    def clear_update_fields():

        search_id_entry.delete(0, tk.END)

        product_name_entry.delete(0, tk.END)
        category_entry.delete(0, tk.END)
        brand_entry.delete(0, tk.END)
        price_entry.delete(0, tk.END)
        quantity_entry.delete(0, tk.END)
        supplier_entry.delete(0, tk.END)
        stock_status_entry.delete(0, tk.END)

    # ---------------- SEARCH RECORD ----------------

    def search_record():

        product_id = search_id_entry.get().strip()

        if product_id == "":
            messagebox.showerror(
                "Error",
                "Please enter Product ID."
            )
            return

        try:

            workbook = openpyxl.load_workbook(FILE_NAME)
            sheet = workbook["Products"]

            found = False

            for row in sheet.iter_rows(min_row=2):

                if str(row[0].value) == product_id:

                    product_name_entry.delete(0, tk.END)
                    product_name_entry.insert(
                        0,
                        row[1].value or ""
                    )

                    category_entry.delete(0, tk.END)
                    category_entry.insert(
                        0,
                        row[2].value or ""
                    )

                    brand_entry.delete(0, tk.END)
                    brand_entry.insert(
                        0,
                        row[3].value or ""
                    )

                    price_entry.delete(0, tk.END)
                    price_entry.insert(
                        0,
                        row[4].value or ""
                    )

                    quantity_entry.delete(0, tk.END)
                    quantity_entry.insert(
                        0,
                        row[5].value or ""
                    )

                    supplier_entry.delete(0, tk.END)
                    supplier_entry.insert(
                        0,
                        row[6].value or ""
                    )

                    stock_status_entry.delete(0, tk.END)
                    stock_status_entry.insert(
                        0,
                        row[7].value or ""
                    )

                    found = True
                    break

            workbook.close()

            if not found:

                messagebox.showerror(
                    "Not Found",
                    "No product found with this Product ID."
                )

        except Exception as e:

            messagebox.showerror(
                "Error",
                "Unable to search record.\n\n" + str(e)
            )

    # ---------------- SAVE UPDATE ----------------

    def save_update():

        product_id = search_id_entry.get().strip()

        if product_id == "":
            messagebox.showerror(
                "Error",
                "Please enter Product ID."
            )
            return

        values = [
            product_name_entry.get().strip(),
            category_entry.get().strip(),
            brand_entry.get().strip(),
            price_entry.get().strip(),
            quantity_entry.get().strip(),
            supplier_entry.get().strip(),
            stock_status_entry.get().strip()
        ]

        if any(value == "" for value in values):

            messagebox.showerror(
                "Error",
                "Please fill all fields."
            )
            return

        try:

            workbook = openpyxl.load_workbook(FILE_NAME)
            sheet = workbook["Products"]

            found = False

            for row in sheet.iter_rows(min_row=2):

                if str(row[0].value) == product_id:

                    row[1].value = values[0]
                    row[2].value = values[1]
                    row[3].value = values[2]
                    row[4].value = values[3]
                    row[5].value = values[4]
                    row[6].value = values[5]
                    row[7].value = values[6]

                    found = True
                    break

            if found:

                workbook.save(FILE_NAME)
                workbook.close()

                messagebox.showinfo(
                    "Success",
                    "Product record updated successfully!"
                )

                clear_update_fields()

            else:

                workbook.close()

                messagebox.showerror(
                    "Not Found",
                    "No product found with this Product ID."
                )

        except Exception as e:

            messagebox.showerror(
                "Error",
                "Unable to update record.\n\n" + str(e)
            )

    # ---------------- BUTTONS ----------------

    tk.Button(
        scrollable_frame,
        text="SEARCH RECORD",
        width=20,
        height=2,
        command=search_record
    ).pack(pady=10)

    tk.Button(
        scrollable_frame,
        text="UPDATE RECORD",
        width=20,
        height=2,
        command=save_update
    ).pack(pady=10)

    tk.Button(
        scrollable_frame,
        text="CLEAR",
        width=20,
        height=2,
        command=clear_update_fields
    ).pack(pady=10)


# =====================
# DELETE RECORD
# =====================

def delete_record():

    delete_window = tk.Toplevel()

    delete_window.title(
        "Delete Record"
    )

    delete_window.geometry(
        "400x300"
    )


    # ---------------- TITLE ----------------

    tk.Label(
        delete_window,
        text="DELETE RECORD",
        font=("Arial", 12, "bold")
    ).pack(pady=20)


    # ---------------- PRODUCT ID LABEL ----------------

    tk.Label(
        delete_window,
        text="Enter Product ID",
        font=("Arial", 10)
    ).pack(pady=5)


    # ---------------- PRODUCT ID ENTRY ----------------

    product_id_entry = tk.Entry(
        delete_window,
        width=25,
        font=("Arial", 11)
    )

    product_id_entry.pack(pady=5)


    # ==================
    # DELETE DATA
    # ==================

    def delete_data():

        product_id = product_id_entry.get().strip()


        if product_id == "":

            messagebox.showwarning(
                "Warning",
                "Please enter Product ID"
            )

            return


        try:

            workbook = openpyxl.load_workbook(
                FILE_NAME
            )

            sheet = workbook["Products"]

            found = False


            for row in range(
                2,
                sheet.max_row + 1
            ):

                if str(
                    sheet.cell(
                        row=row,
                        column=1
                    ).value
                ) == product_id:

                    product_name = sheet.cell(
                        row=row,
                        column=2
                    ).value


                    confirm = messagebox.askyesno(
                        "Confirm Delete",
                        f"Are you sure you want to delete\n"
                        f"{product_name}?"
                    )


                    if confirm:

                        sheet.delete_rows(
                            row,
                            1
                        )

                        workbook.save(
                            FILE_NAME
                        )

                        messagebox.showinfo(
                            "Success",
                            "Record deleted successfully!"
                        )

                        product_id_entry.delete(
                            0,
                            tk.END
                        )


                    found = True

                    break


            workbook.close()


            if not found:

                messagebox.showerror(
                    "Error",
                    "Product ID not found!"
                )


        except Exception as e:

            messagebox.showerror(
                "Error",
                "Unable to delete record.\n\n" + str(e)
            )


    # =======================
    # DELETE BUTTON
    # =======================

    tk.Button(
        delete_window,
        text="Delete Record",
        width=25,
        height=2,
        command=delete_data
    ).pack(pady=15)


    # =======================
    # CLEAR BUTTON
    # =======================

    tk.Button(
        delete_window,
        text="Clear",
        width=20,
        height=2,
        command=lambda: product_id_entry.delete(
            0,
            tk.END
        )
    ).pack(pady=5)


# =========================
# DASHBOARD
# =========================

def dashboard():

    global dashboard_window

    dashboard_window = tk.Tk()

    dashboard_window.title(
        "Dashboard - Electronic Shop Management System"
    )

    dashboard_window.attributes("-fullscreen", True)


    # ===================
    # TITLE
    # ===================

    title_label = tk.Label(
        dashboard_window,
        text="ELECTRONIC SHOP",
        font=("Arial", 22, "bold")
    )

    title_label.pack(
        pady=(30, 5)
    )


    # =========================
    # SUBTITLE
    # =========================

    subtitle_label = tk.Label(
        dashboard_window,
        text="Management System",
        font=("Arial", 14)
    )

    subtitle_label.pack(
        pady=5
    )


    # ========================
    # WELCOME
    # ========================

    welcome_label = tk.Label(
        dashboard_window,
        text="Welcome, Admin!",
        font=("Arial", 12)
    )

    welcome_label.pack(
        pady=20
    )


    # =========================
    # ADD RECORD
    # =========================

    tk.Button(
        dashboard_window,
        text="Add Record",
        width=25,
        height=2,
        command=add_record
    ).pack(
        pady=8
    )


    # ========================
    # VIEW RECORDS
    # ========================

    tk.Button(
        dashboard_window,
        text="View Records",
        width=25,
        height=2,
        command=view_records
    ).pack(
        pady=8
    )

    # =========================
    # UPDATE RECORD
    # =========================

    tk.Button(
        dashboard_window,
        text="Update Record",
        width=25,
        height=2,
        command=update_record
    ).pack(
        pady=8
    )


    # =========================
    # DELETE RECORD
    # =========================

    tk.Button(
        dashboard_window,
        text="Delete Record",
        width=25,
        height=2,
        command=delete_record
    ).pack(
        pady=8
    )


    # ========================
    # LOGOUT FUNCTION
    # ========================

    def logout():

        dashboard_window.destroy()
        login_window.deiconify()


    # ========================
    # LOGOUT
    # ========================

    tk.Button(
        dashboard_window,
        text="Logout",
        width=25,
        height=2,
        command=logout
    ).pack(
        pady=20
    )


    dashboard_window.mainloop()


# =========================
# LOGIN GUI
# =========================

login_window = tk.Tk()

login_window.title(
    "Login - Electronic Shop Management System"
)

login_window.geometry(
    "400x350"
)


# =========================
# TITLE
# =========================

title_label = tk.Label(
    login_window,
    text="ELECTRONIC SHOP",
    font=("Arial", 18, "bold")
)

title_label.pack(
    pady=25
)


# ==========================
# SUBTITLE
# ==========================

subtitle_label = tk.Label(
    login_window,
    text="MANAGEMENT SYSTEM LOGIN",
    font=("Arial", 11)
)

subtitle_label.pack(
    pady=5
)


# ==========================
# USERNAME
# ==========================

tk.Label(
    login_window,
    text="Username",
    font=("Arial", 11)
).pack(
    pady=(20, 5)
)


username_entry = tk.Entry(
    login_window,
    width=30
)

username_entry.pack()


# ==========================
# PASSWORD
# ==========================

tk.Label(
    login_window,
    text="Password",
    font=("Arial", 11)
).pack(
    pady=(15, 5)
)


password_entry = tk.Entry(
    login_window,
    width=30,
    show="*"
)

password_entry.pack()


# ==========================
# LOGIN BUTTON
# ==========================

tk.Button(
    login_window,
    text="Login",
    width=15,
    command=login
).pack(
    pady=20
)


# ==========================
# CLEAR BUTTON
# ==========================

tk.Button(
    login_window,
    text="Clear",
    width=15,
    command=clear_login
).pack()


# ==========================
# START APPLICATION
# ==========================

login_window.mainloop()