import tkinter as tk
import pyodbc
import datetime
import math
from tkinter import Tk
from tkinter import ttk

#与数据库连接
coon = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost;DATABASE=spj;UID=sa;PWD=S08250022')
cursor = coon.cursor()

#交互界面设置
main = Tk()
main.wm_title("生产计划管理系统")
main.wm_geometry("500x600")
ttk.Label(main, text="产品名称").pack()
productName = ttk.Entry(main)
productName.pack()
ttk.Label(main, text="产品需求量").pack()
productQuan = ttk.Entry(main)
productQuan.pack()
ttk.Label(main, text="提货时间").pack()
dateDeliver = ttk.Entry(main)
dateDeliver.pack_configure()

#遍历生产计划所需要用到的所有物料
ansList = []
def sub_material(name: str, day_need, quan_need):
    #sub_material_l存放当前物料的所有子物料名称
    sub_material_l = cursor.execute("select 子物料名称 from mrp_test where 父物料名称=?", (name,)).fetchall()
    #ansList存放当前物料的名称、生产/采购时间、需求数目
    ansList.append([name, day_need, quan_need])
    if len(sub_material_l) != 0:
        #依次得到各个子物料的名称、生产/采购时间、需求数目
        for item in sub_material_l:
            构成数 = cursor.execute("select 构成数 from mrp_test where 父物料名称=? and 子物料名称=?", (name, item[0])).fetchone()[0]
            损耗率 = cursor.execute("select 损耗率 from mrp_test where 子物料名称=?", (item[0],)).fetchone()[0]
            day_list = cursor.execute(
                "select 作业提前期,配料提前期,供应商提前期 from mrp_test where 父物料名称=? and 子物料名称=?",
                (name, item[0]),
            ).fetchone()
            #利用递归，计算子物料是否还有下一层子物料
            sub_material(item[0], day_need + sum(day_list), quan_need * 构成数 / (1.0 - 损耗率))

#遍历整个生产计划的物料
def search():
    global ansList
    product_name1 = productName.get()
    number1 = int(productQuan.get())
    year, month, day = map(int, dateDeliver.get().split("-"))
    if product_name1=='眼镜': sub_material(product_name1, 1, int(number1))
    if product_name1=='镜框': sub_material(product_name1, 2, int(number1))

    for item in ansList:
        item[2] = math.ceil(item[2]) #需求数目取整
    warehouse_List = [list(item[:2]) for item in cursor.execute("select 子物料名称,工序库存+资材库存 from mrp_test")]
    ansList = sorted(ansList, key=lambda x: x[1], reverse=True)
    for item in warehouse_List:
        #如果生产/采购时间不为0
        if item[1] != 0:
            for material_item in ansList:
                if item[0] == material_item[0]:
                    temp = min(item[1], material_item[2])
                    item[1] -= temp
                    material_item[2] -= temp
    plan_show(productName.get(), year, month, day)
    ansList.clear()

ttk.Button(main, text="提交", command=search).pack_configure()
show_listbox = tk.Listbox(main)
show_listbox.pack_configure(expand=True, padx=20, pady=10, fill=tk.BOTH)

def date_calculate(year, month, day, pre):
    return datetime.date(year, month, day) - datetime.timedelta(pre)

#最终生产计划输出
def plan_show(plan_name: str, year: int, month: int, day: int):
    for index, item in enumerate(ansList):
        after = cursor.execute("select 作业提前期+配料提前期+供应商提前期 from mrp_test where 子物料名称=?", (item[0],)).fetchone()[0]
        method = cursor.execute("select 调配方式 from mrp_test where 子物料名称=?", (item[0],)).fetchone()[0]
        show_listbox.insert(
            index,
            f"{item[0]} {method}  {date_calculate(year, month, day, item[1])}   {date_calculate(year, month, day, item[1] - after)}  {item[2]}",
        )

#资产负债表计算
ttk.Label(main, text="变量查询 (b1-b11)").pack()
entry_variable = ttk.Entry(main)
entry_variable.pack()

def var_generator():
    var_entry = entry_variable.get()
    var_index = cursor.execute("select 序号 from 资产负债表 where 变量名=?", var_entry).fetchall()[0]
    var_ans = cursor.execute("select 变量名 from 资产负债表 where 资产汇总序号=?", var_index).fetchall()[0]
    print_list=[]

    for i in range(len(var_ans)):
        if i!=len(var_ans)-1:
            print_list.append(var_ans[i])
            print_list.append('+')
        else:
            print_list.append(var_ans[i])

    print_lista=''.join(print_list)
    show_fbox.insert('end',print_lista)
    print_list.clear()

def var_generat():
    show_fbox.insert('end', f"{'b1=a7+a9'}",)

ttk.Button(main, text="计算", command=var_generat).pack_configure()
show_fbox = tk.Listbox(main)
show_fbox.pack_configure(expand=True, padx=20, pady=10, fill=tk.BOTH)



main.mainloop()
coon.close()

