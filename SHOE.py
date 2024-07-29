import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import mysql.connector
from mysql.connector import Error
from tkinter import messagebox
from PIL import Image, ImageTk
import pandas as pd

def fetch_data():
    try:
        conn = mysql.connector.connect(
            host='localhost',
            user='root',
            password='123123',
            database='shoe_factory'
        )
        if conn.is_connected():
            print("连接成功")
        cursor = conn.cursor()
        cursor.execute('''
            SELECT po.id, po.order_date, po.product_name, po.quantity as order_quantity,
                   m.name as material_name, om.quantity as material_quantity, m.unit
            FROM production_orders po
            JOIN order_materials om ON po.id = om.order_id
            JOIN materials m ON om.material_id = m.id
        ''')
        rows = cursor.fetchall()
        if rows:
            print("查询成功，获取到数据：", rows)
        else:
            print("查询成功，但未获取到数据")
        update_treeview(rows)
        cursor.close()
        conn.close()
    except Error as err:
        messagebox.showerror("Database Error", str(err))

def update_treeview(rows):
    for row in tree.get_children():
        tree.delete(row)
    for row in rows:
        tree.insert("", "end", values=row)

def insert_data():
    global order_date_entry, product_name_entry, order_quantity_entry, material_name_entry, material_quantity_entry, unit_entry

    def submit_data():
        order_date = order_date_entry.get()
        product_name = product_name_entry.get()
        order_quantity = order_quantity_entry.get()
        material_name = material_name_entry.get()
        material_quantity = material_quantity_entry.get()
        unit = unit_entry.get()

        try:
            conn = mysql.connector.connect(
                host='localhost',
                user='root',
                password='123123',
                database='shoe_factory'
            )
            cursor = conn.cursor()
            
            # 插入材料数据
            cursor.execute("INSERT INTO materials (name, unit, stock_quantity) VALUES (%s, %s, %s)", 
                           (material_name, unit, material_quantity))
            material_id = cursor.lastrowid

            # 插入生产订单数据
            cursor.execute("INSERT INTO production_orders (order_date, product_name, quantity) VALUES (%s, %s, %s)",
                           (order_date, product_name, order_quantity))
            order_id = cursor.lastrowid

            # 插入订单材料数据
            cursor.execute("INSERT INTO order_materials (order_id, material_id, quantity) VALUES (%s, %s, %s)",
                           (order_id, material_id, material_quantity))

            conn.commit()
            cursor.close()
            conn.close()
            fetch_data()  # 更新数据视图
            messagebox.showinfo("Success", "数据插入成功")
            clear_entries()  # 插入数据后清空输入框
        except Error as err:
            messagebox.showerror("Database Error", str(err))
        except Exception as e:
            messagebox.showerror("Error", str(e))

    insert_window = tk.Toplevel(app)
    insert_window.title("插入数据")

    tk.Label(insert_window, text="订单日期 (YYYY-MM-DD):").pack()
    order_date_entry = tk.Entry(insert_window)
    order_date_entry.pack()

    tk.Label(insert_window, text="产品名称:").pack()
    product_name_entry = tk.Entry(insert_window)
    product_name_entry.pack()

    tk.Label(insert_window, text="订单数量:").pack()
    order_quantity_entry = tk.Entry(insert_window)
    order_quantity_entry.pack()

    tk.Label(insert_window, text="材料名称:").pack()
    material_name_entry = tk.Entry(insert_window)
    material_name_entry.pack()

    tk.Label(insert_window, text="材料数量:").pack()
    material_quantity_entry = tk.Entry(insert_window)
    material_quantity_entry.pack()

    tk.Label(insert_window, text="单位:").pack()
    unit_entry = tk.Entry(insert_window)
    unit_entry.pack()

    submit_button = tk.Button(insert_window, text="提交", command=submit_data)
    submit_button.pack()

def delete_data():
    def submit_deletion():
        selected_item = tree.selection()
        if selected_item:
            order_id = tree.item(selected_item)['values'][0]
            try:
                conn = mysql.connector.connect(
                    host='localhost',
                    user='root',
                    password='123123',
                    database='shoe_factory'
                )
                cursor = conn.cursor()
                
                # 删除订单材料数据
                cursor.execute("DELETE FROM order_materials WHERE order_id = %s", (order_id,))
                
                # 删除生产订单数据
                cursor.execute("DELETE FROM production_orders WHERE id = %s", (order_id,))
                
                conn.commit()
                cursor.close()
                conn.close()
                fetch_data()  # 更新数据视图
                messagebox.showinfo("Success", "数据删除成功")
            except Error as err:
                messagebox.showerror("Database Error", str(err))
        else:
            messagebox.showwarning("Selection Error", "请选择要删除的数据")

    delete_window = tk.Toplevel(app)
    delete_window.title("删除数据")

    tk.Label(delete_window, text="选择要删除的数据并点击删除按钮").pack()
    delete_button = tk.Button(delete_window, text="删除", command=submit_deletion)
    delete_button.pack()

def export_to_excel():
    try:
        conn = mysql.connector.connect(
            host='localhost',
            user='root',
            password='123123',
            database='shoe_factory'
        )
        cursor = conn.cursor()
        cursor.execute('''
            SELECT po.id, po.order_date, po.product_name, po.quantity as order_quantity,
                   m.name as material_name, om.quantity as material_quantity, m.unit
            FROM production_orders po
            JOIN order_materials om ON po.id = om.order_id
            JOIN materials m ON om.material_id = m.id
        ''')
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        
        # 使用 pandas 将数据保存到 Excel
        df = pd.DataFrame(rows, columns=["Order ID", "Order Date", "Product Name", "Order Quantity", "Material Name", "Material Quantity", "Unit"])
        df.to_excel("生产订单和材料信息.xlsx", index=False)
        messagebox.showinfo("Success", "数据已导出到生产订单和材料信息.xlsx")
    except Error as err:
        messagebox.showerror("Database Error", str(err))

def clear_entries():
    order_date_entry.delete(0, tk.END)
    product_name_entry.delete(0, tk.END)
    order_quantity_entry.delete(0, tk.END)
    material_name_entry.delete(0, tk.END)
    material_quantity_entry.delete(0, tk.END)
    unit_entry.delete(0, tk.END)

def import_image():
    file_path = filedialog.askopenfilename()
    if file_path:
        try:
            img = Image.open(file_path)
            img = img.resize((200, 200), Image.Resampling.LANCZOS)
            img_tk = ImageTk.PhotoImage(img)
            panel = tk.Label(app, image=img_tk)
            panel.image = img_tk
            panel.grid(row=0, column=1, rowspan=4, padx=10, pady=10)
        except Exception as e:
            messagebox.showerror("Error", str(e))

app = tk.Tk()
app.title("鞋厂生产所需材料报表系统")

# 创建数据展示表格
tree = ttk.Treeview(app, columns=("Order ID", "Order Date", "Product Name", "Order Quantity", "Material Name", "Material Quantity", "Unit"), show='headings')
tree.heading("Order ID", text="订单ID")
tree.heading("Order Date", text="订单日期")
tree.heading("Product Name", text="产品名称")
tree.heading("Order Quantity", text="订单数量")
tree.heading("Material Name", text="材料名称")
tree.heading("Material Quantity", text="材料数量")
tree.heading("Unit", text="单位")
tree.grid(row=0, column=0, columnspan=4, padx=10, pady=10, sticky='nsew')

# 创建按钮
fetch_button = tk.Button(app, text="获取数据", command=fetch_data)
fetch_button.grid(row=1, column=0, padx=5, pady=5, sticky='ew')

insert_button = tk.Button(app, text="插入数据", command=insert_data)
insert_button.grid(row=1, column=1, padx=5, pady=5, sticky='ew')

delete_button = tk.Button(app, text="删除选中数据", command=delete_data)
delete_button.grid(row=1, column=2, padx=5, pady=5, sticky='ew')

export_button = tk.Button(app, text="导出为Excel", command=export_to_excel)
export_button.grid(row=1, column=3, padx=5, pady=5, sticky='ew')

import_button = tk.Button(app, text="导入图片", command=import_image)
import_button.grid(row=2, column=0, columnspan=4, padx=5, pady=5, sticky='ew')

# 设置窗口布局
app.grid_rowconfigure(0, weight=1)
app.grid_columnconfigure(0, weight=1)

app.geometry("800x600")
app.mainloop()
