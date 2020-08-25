from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
from data import total_status_by_country, total_status_by_country_excel, getCountries, just_global_excel, just_global, sorted_data
from ttkthemes import ThemedTk



window = ThemedTk(theme='arc')
window.title("Covid19")
window.geometry('980x410')

# labels
label_choose_country = Label(window, text="Choose a Country")
label_choose_country.place(x = 170, y = 10)

label_sorted = Label(window, text="Sorted by Parameter")
label_sorted.place(x = 650, y = 10)

label_parameters = Label(window, text="Country Count : ")
label_parameters.place(x=610, y = 38)
# labels
sort = ["cases", "todayCases", "deaths", "todayDeaths", "recovered", "active"]

combo = Combobox(window, state="readonly",
                 values=getCountries(), width=36)
combo.current(0)
combo.place(x=100, y=35)


combo3 = Combobox(window, state="readonly",
                 values=sort, width=12)
combo3.current(0)
combo3.place(x=610, y=70)


country_count_entry = Entry(window, width = 15)
country_count_entry.place(x = 700, y = 35)


def country():
    total_status_by_country_excel(combo.get())


def global_data():
    just_global_excel()


def cond():
    if(combo.get() == "Global"):
        global_data()
    else:
        country()


# TreeView
tree = Treeview(window)
tree.place(y=110, x=50)

vsb = Scrollbar(window, orient="vertical", command=tree.yview)
vsb.place(x=914, y=142, height=226)
tree.configure(yscrollcommand=vsb.set)

tree["columns"] = ("one", "two", "three", "four", "five", "six", "seven")
tree.column("#0", width=150, minwidth=150, stretch=NO)
tree.column("one", width=100, minwidth=100, stretch=NO)
tree.column("two", width=100, minwidth=100, stretch=NO)
tree.column("three", width=100, minwidth=100, stretch=NO)
tree.column("four", width=100, minwidth=100, stretch=NO)
tree.column("five", width=100, minwidth=100, stretch=NO)
tree.column("six", width=100, minwidth=100, stretch=NO)
tree.column("seven", width=100, minwidth=100)

tree.heading("#0", text="Country", anchor=W)
tree.heading("one", text="Today Cases", anchor=W)
tree.heading("two", text="Today Deaths", anchor=W)
tree.heading("three", text="Active Cases", anchor=W)
tree.heading("four", text="Tests", anchor=W)
tree.heading("five", text="Recovered", anchor=W)
tree.heading("six", text="Cases", anchor=W)
tree.heading("seven", text="Deaths", anchor=W)


def show_info():
    if total_status_by_country(combo.get()) is not None:
        tree.insert("", "end", text=combo.get(), values=(total_status_by_country(combo.get())[0], total_status_by_country(combo.get())[1], total_status_by_country(combo.get())[2],
        total_status_by_country(combo.get())[3],total_status_by_country(combo.get())[4],
        total_status_by_country(combo.get())[5],total_status_by_country(combo.get())[6]))
    else:
        messagebox.showinfo('Error', 'There is no data')

def global_data_2():
    tree.insert("", "end", text=combo.get(), values=(
        just_global()[0], just_global()[1], just_global()[2],
        just_global()[3], just_global()[4], just_global()[5],
        just_global()[6]))


def cond_2():
    if(combo.get() == "Global"):
        global_data_2()
    else:
        show_info()

def delete():
    for i in tree.get_children():
        tree.delete(i)

def cond_3():
    if int(country_count_entry.get()) < 216:
        for i in range(int(country_count_entry.get())):
            temp = sorted_data(combo3.get())[i]
            tree.insert("", "end", text=temp[0], values=(
                temp[1], temp[2], temp[3], 0, temp[4], temp[5], temp[6]
            ))
    else:
        messagebox.showinfo('Error', 'Invalid input')

# TreeView

# Button
btn = Button(window, text="Export as Excel", command=cond)
btn.place(x=100, y=70)

show_info_btn = Button(window, text="Show Informations", command=cond_2)
show_info_btn.place(x=220, y=70)

clear_btn = Button(window, text="Clear", command=delete)
clear_btn.place(x=847, y=375)

add_btn = Button(window, text="Sort", command=cond_3)
add_btn.place(x=720, y=70)
# Button


window.mainloop()
