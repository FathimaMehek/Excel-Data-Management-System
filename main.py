from tkinter import *
import tkinter as tk
import tkinter.messagebox as messagebox
from tkinter import ttk
import pandas as pd                         # pip install pandas
import matplotlib.pyplot as plt             # pip install matplotlib
import numpy as np

class RescueDB:

    def __init__(self, root):
        self.root = root
        self.root.title("Data Management Systems")
        self.root.geometry("1920x900+0+0")

        TitleFrame = Frame(self.root, bd=14, width=1920, height=150, padx=12, relief=RIDGE)
        TitleFrame.grid(row=0, column=0)
        MainFrame = Frame(self.root)
        MainFrame.grid(row=1, column=0)

        TopFrame = Frame(MainFrame, bd=14, width=1350, height=550, padx=4, relief=RIDGE)
        TopFrame.grid(row=0, column=0)

        LeftFrameMain = Frame(TopFrame, bd=10, width=400, height=750, relief=RIDGE)
        LeftFrameMain.grid(row=0, column=0)
        LeftFrame = Frame(LeftFrameMain, bd=10, width=450, height=500, relief=RIDGE)
        LeftFrame.grid(row=0, column=0)
        Leftbottom = Frame(LeftFrameMain, bd=10, width=650, height=190, relief=RIDGE)
        Leftbottom.grid(row=1, column=0, pady=1)

        RightFrame = Frame(TopFrame, bd=10, width=1200, height=550, pady=6, relief=RIDGE)
        RightFrame.grid(row=0, column=1)

        BottomFrame = Frame(MainFrame, bd=10, width=1350, height=150, padx=14, relief=RIDGE)
        BottomFrame.grid(row=1, column=0)

        # =====================================Function=================================================

        def calculate_total_and_percentage():
            try:
                mongodb_marks = float(mongodb_entry.get()) if mongodb_entry.get() else 0
                php_marks = float(php_entry.get()) if php_entry.get() else 0
                data_structure_marks = float(data_structure_entry.get()) if data_structure_entry.get() else 0

                total_marks = mongodb_marks + php_marks + data_structure_marks
                percentage = (total_marks / 300) * 100 if total_marks > 0 else 0  # Assuming each subject is out of 100
                total_marks_entry.delete(0, END)
                total_marks_entry.insert(0, str(total_marks))
                percentage_entry.delete(0, END)
                percentage_entry.insert(0, f"{percentage:.2f}")
            except ValueError:
                messagebox.showerror("Input Error", "Please enter valid marks.")

        def add_data():
            try:
                calculate_total_and_percentage()  # Calculate before adding
                df = pd.read_excel("Student_Data.xlsx")
                usn = usn_entry.get().strip()
                new_data = {
                    'USN': [usn],
                    'Student_Name': [student_name_entry.get()],
                    'MongoDB': [mongodb_entry.get()],
                    'PHP': [php_entry.get()],
                    'Data_Structure': [data_structure_entry.get()],
                    'Total_Marks': [total_marks_entry.get()],
                    'Percentage': [percentage_entry.get()]
                }
                new_df = pd.DataFrame(new_data)

                # Check if USN already exists
                if usn in df['USN'].astype(str).values:
                    messagebox.showerror("Error", "USN already exists. Please use a unique USN.")
                    return

                df = pd.concat([df, new_df], ignore_index=True)
                df = df.sort_values(by='USN')  # Sort by USN in ascending order
                df.to_excel("Student_Data.xlsx", index=False)
                messagebox.showinfo("Success", "Data added successfully.")
                reset_entries()
                refresh_treeview()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        def update_data():
            try:
                calculate_total_and_percentage()  # Calculate before updating
                df = pd.read_excel("Student_Data.xlsx")
                usn = usn_entry.get().strip()  # Strip any leading/trailing spaces

                # Check if the USN exists in the DataFrame
                if usn in df['USN'].astype(str).values:  # Ensure USN is treated as string
                    # Update the row with the new data
                    df.loc[df['USN'].astype(str) == usn, ['Student_Name', 'MongoDB', 'PHP', 'Data_Structure', 'Total_Marks', 'Percentage']] = [
                        student_name_entry.get(),
                        mongodb_entry.get(),
                        php_entry.get(),
                        data_structure_entry.get(),
                        total_marks_entry.get(),
                        percentage_entry.get()
                    ]
                    df.to_excel("Student_Data.xlsx", index=False)
                    messagebox.showinfo("Success", "Data updated successfully.")
                else:
                    messagebox.showerror("Error", "USN not found.")

                reset_entries()
                refresh_treeview()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        def delete_data():
            try:
                df = pd.read_excel("Student_Data.xlsx")
                usn = usn_entry.get().strip()  # Strip any leading/trailing spaces

                # Check if the USN exists in the DataFrame
                if usn in df['USN'].astype(str).values:  # Ensure USN is treated as string
                    df = df[df['USN'].astype(str) != usn]  # Remove the row
                    df.to_excel("Student_Data.xlsx", index=False)
                    messagebox.showinfo("Success", "Data deleted successfully.")
                    reset_entries()
                    refresh_treeview()
                else:
                    messagebox.showerror("Error", "USN not found.")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        def reset_entries():
            usn_entry.delete(0, END)
            student_name_entry.delete(0, END)
            mongodb_entry.delete(0, END)
            php_entry.delete(0, END)
            data_structure_entry.delete(0, END)
            total_marks_entry.delete(0, END)
            percentage_entry.delete(0, END)

        def refresh_treeview():
            try:
                df = pd.read_excel('Student_Data.xlsx')
                treeview.delete(*treeview.get_children())
                for index, row in df.iterrows():
                    treeview.insert('', 'end', values=(row['USN'], row['Student_Name'], row['MongoDB'], row['PHP'], row['Data_Structure'], row['Total_Marks'], row['Percentage']))
            except Exception as e:
                messagebox.showerror('Error', str(e))

        def plot_graph():
            try:
                df = pd.read_excel('Student_Data.xlsx')
                plt.figure(figsize=(10, 6))
                colors = np.random.rand(len(df))
                plt.bar(df['Student_Name'], df['Percentage'], color=colors)
                plt.xlabel('Student Name')
                plt.ylabel('Percentage')
                plt.title('Students Percentage by Name')
                plt.xticks(rotation=45)
                plt.tight_layout()
                plt.show()
            except Exception as e:
                messagebox.showerror('Error', str(e))

        # Create the Title widgets
        dataTitle = Label(TitleFrame, font=('arial', 90, 'bold'), padx=16, text='Excel Data Management System')
        dataTitle.grid(row=0, column=0)
        subTitle = Label(Leftbottom, font=('arial', 90, 'bold'), padx=16, text='Excel Data')
        subTitle.grid(row=0, column=0)

        # Create the Entry widgets
        usn_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='USN:')
        usn_label.grid(row=0, column=0)
        usn_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        usn_entry.grid(row=0, column=1)

        student_name_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='Student Name:')
        student_name_label.grid(row=1, column=0)
        student_name_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        student_name_entry.grid(row=1, column=1)

        mongodb_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='MongoDB:')
        mongodb_label.grid(row=2, column=0)
        mongodb_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        mongodb_entry.grid(row=2, column=1)

        php_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='PHP:')
        php_label.grid(row=3, column=0)
        php_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        php_entry.grid(row=3, column=1)

        data_structure_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='Data Structure:')
        data_structure_label.grid(row=4, column=0)
        data_structure_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        data_structure_entry.grid(row=4, column=1)

        total_marks_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='Total Marks:')
        total_marks_label.grid(row=5, column=0)
        total_marks_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        total_marks_entry.grid(row=5, column=1)

        percentage_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='Percentage:')
        percentage_label.grid(row=6, column=0)
        percentage_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        percentage_entry.grid(row=6, column=1)

        # Create the buttons
        add_button = Button(BottomFrame, pady=1, bd=4, font=('arial', 40, 'bold'),
                            width=11, height=1, text='Add Data', command=add_data)
        add_button.grid(row=0, column=0, padx=3)

        update_button = Button(BottomFrame, pady=1, bd=4, font=('arial', 40, 'bold'),
                               width=11, height=1, text='Update', command=update_data)
        update_button.grid(row=0, column=1, padx=3)

        delete_button = Button(BottomFrame, pady=1, bd=4, font=('arial', 40, 'bold'),
                               width=11, height=1, text='Delete', command=delete_data)
        delete_button.grid(row=0, column=2, padx=3)

        plot_button = Button(BottomFrame, pady=1, bd=4, font=('arial', 40, 'bold'),
                             width=11, height=1, text='Plot Graph', command=plot_graph)
        plot_button.grid(row=0, column=3, padx=3)

        reset_button = Button(BottomFrame, pady=1, bd=4, font=('arial', 40, 'bold'),
                              width=11, height=1, text='Reset', command=reset_entries)
        reset_button.grid(row=0, column=4, padx=3)

        # Create the Treeview widget to display the data        
        style = ttk.Style()
        style.configure('Treeview.Heading', font=('TkDefaultFont', 18))
        style.configure('Treeview', rowheight=40, font=('TkDefaultFont', 18))

        treeview_columns = ('USN', 'Student Name', 'MongoDB', 'PHP', 'Data Structure', 'Total Marks', 'Percentage')
        treeview = ttk.Treeview(RightFrame, columns=treeview_columns, show='headings', height=10)
        treeview.grid(row=0, columnspan=10, pady=34)

        for col in treeview_columns:
            treeview.heading(col, text=col)
            treeview.column(col, width=170)
            treeview.column(col, anchor='center')

        # Load data from the Excel spreadsheet and display it in the Treeview
        try:
            df = pd.read_excel('Student_Data.xlsx')
            df = df.sort_values(by='USN')  # Sort by USN in ascending order
            for index, row in df.iterrows():
                treeview.insert('', 'end', values=(row['USN'], row['Student_Name'], row['MongoDB'], row['PHP'],
                                                   row['Data_Structure'], row['Total_Marks'], row['Percentage']))
        except Exception as e:
            messagebox.showerror('Error', str(e))

        # Function to handle the Treeview row selection event
        def on_treeview_select(event):
            selected_item = treeview.focus()
            if selected_item:
                values = treeview.item(selected_item, 'values')
                usn_entry.delete(0, tk.END)
                student_name_entry.delete(0, tk.END)
                mongodb_entry.delete(0, tk.END)
                php_entry.delete(0, tk.END)
                data_structure_entry.delete(0, tk.END)
                total_marks_entry.delete(0, tk.END)
                percentage_entry.delete(0, tk.END)

                usn_entry.insert(0, values[0])
                student_name_entry.insert(0, values[1])
                mongodb_entry.insert(0, values[2])
                php_entry.insert(0, values[3])
                data_structure_entry.insert(0, values[4])
                total_marks_entry.insert(0, values[5])
                percentage_entry.insert(0, values[6])

        # Bind the function to the Treeview selection event
        treeview.bind('<<TreeviewSelect>>', on_treeview_select)

if __name__ == '__main__':
    root = Tk()
    application = RescueDB(root)
    root.mainloop()
