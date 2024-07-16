# NOTE: Not finish incorporting suggestions and making changes to code. 

from tkinter import *
from datetime import *
from tkinter.scrolledtext import ScrolledText
from tkcalendar import Calendar
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
import openpyxl as op

class App(Tk):
    # df = pd.read_excel("RunLog.xlsx")

    def __init__(self):
        Tk.__init__(self)
        self.title('Running Log App')
        self.df = pd.read_excel("RunLog.xlsx")

         # Index label
        self.index_label = Label(self, text="Index (ONLY to Edit or Delete data):")
        self.index_label.grid(row=0, column=0, padx=10, pady=10, sticky="e")

        # Index dropdown menu
        self.index_user_input = IntVar()
        self.index_list = list(self.df.index.values) # Gets index values and makes it a list
        self.index_select = OptionMenu(self, self.index_user_input, *self.index_list)
        self.index_select.grid(row=0, column=1, padx=10, pady=10, sticky="e")

        # Columns label
        self.col_label = Label(self, text="Select column (ONLY to Edit data):")
        self.col_label.grid(row=1, column=0, padx=10, pady=10, sticky="e")

        # Column names dropdown menu
        self.col_name = StringVar()
        self.col_list = ["Date", "Run Type", "Hours", "Minutes", "Seconds", "Miles"]
        self.col = OptionMenu(self, self.col_name, *self.col_list)
        self.col.grid(row=1, column=1, padx=10, pady=10, sticky="e")

        # Date label
        self.date_label = Label(self, text="Date (MM/DD/YYYY):")
        self.date_label.grid(row=1, column=3, padx=10, pady=10, sticky="e")

        # Creating a calendar for user to select date 
        current_year = datetime.now().year
        current_month = datetime.now().month
        current_day = datetime.now().day

        self.date_input = Calendar(self, selectmode="day", font="Arial 7", locale="en_US", year=current_year, month=current_month, day=current_day, date_pattern="MM/DD/YYYY")
        self.date_input.grid(row=2, rowspan=4, column=3, padx=10, pady=10, sticky="nsew")
 
        # Run type label
        self.run_type_label = Label(self, text="Run Type (select one):")
        self.run_type_label.grid(row=3, column=0, padx=10, pady=10, sticky="e")
        
        # Road and trail run radio button selections 
        self.run_select = StringVar()
        self.road_select = Radiobutton(self, width=10, text="Road", variable=self.run_select, value="road")
        self.road_select.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

        self.trail_select = Radiobutton(self, width=10, text="Trail", variable=self.run_select, value="trail")
        self.trail_select.grid(row=3, column=2, padx=10, pady=10, sticky="ew")

        # Hours label
        self.hours_label = Label(self, text="Hours")
        self.hours_label.grid(row=4, column=0, padx=10, pady=10, sticky="e")

        # Hours scale. User slides bar to select how many hours
        self.hours_input = IntVar()
        self.hours_entry = Scale(self, from_=0, to=23, orient=HORIZONTAL, length=200, variable=self.hours_input)
        self.hours_entry.grid(row=4, column=1, padx=10, pady=10, sticky="ew")

        # Minutes label
        self.minutes_label = Label(self, text="Minutes:")
        self.minutes_label.grid(row=5, column=0, padx=10, pady=10, sticky="e")

        # Minutes scale. User slides bar to select how many minutes
        self.minutes_input = IntVar()
        self.minutes_entry = Scale(self, from_=0, to=59, orient=HORIZONTAL, length=200, variable=self.minutes_input)
        self.minutes_entry.grid(row=5, column=1, padx=10, pady=10, sticky="ew")

        # Seconds label
        self.seconds_label = Label(self, text="Seconds:")
        self.seconds_label.grid(row=6, column=0, padx=10, pady=10, sticky="e")

        # Seconds scale. User slides bar to select how many seconds
        self.seconds_input = IntVar()
        self.seconds_entry = Scale(self, from_=0, to=59, orient=HORIZONTAL, length=200, variable=self.seconds_input)
        self.seconds_entry.grid(row=6, column=1, padx=10, pady=10, sticky="ew")
 
        # Miles label
        self.miles_label = Label(self, text="Miles:")
        self.miles_label.grid(row=7, column=0, padx=10, pady=10, sticky="e")

        # Miles scale. User slides bar to select how many miles
        self.miles_input = DoubleVar()
        self.miles_entry = Scale(self, from_=0.1, to=100, resolution=0.1, orient=HORIZONTAL, length=200, variable=self.miles_input)
        self.miles_entry.grid(row=7, column=1, padx=10, pady=10, sticky="ew")

        # Insert new data button
        self.insert_button = Button(self, text="Enter New Data", command=self.input_data)
        self.insert_button.grid(row=8, column=0, padx=10, pady=10, sticky="w")

        # Edit data button
        self.save_button = Button(self, text="Save Edit to Table", command=self.edit_data)
        self.save_button.grid(row=8, column=1, padx=10, pady=10, sticky="w")
        
        # Delete data button
        self.delete_button = Button(self, text="Delete Last Row", command=self.delete_row)
        self.delete_button.grid(row=8, column=2, padx=10, pady=10, sticky="w")

        # Calculate button
        self.calculate_button = Button(self, text="Calculate Daily Average", command=self.monthly_averages_results)
        self.calculate_button.grid(row=8, column=3, padx=10, pady=10, sticky="w")

        # Scrolled text widget for calculations 
        self.calculations_text = ScrolledText(self, width=100, height=10)
        self.calculations_text.grid(row=12, column=0, columnspan=10, padx=10, pady=10,)
        # self.calculations_text.insert(self.results)

        # Scrolled text widget for displaying excel spreadsheet data
        self.excel_text = ScrolledText(self, width=100, height=20)
        self.excel_text.grid(row=14, column=0, columnspan=10, padx=10, pady=10,)
        self.excel_text.insert(END, self.df)


    # Add new row of data
    def input_data(self):
        # User inputs
        date_str = self.date_input.get_date() # Convert date string into date object (has time) and extract just date part
        date = datetime.strptime(date_str, "%m/%d/%Y").date()

        run_type = self.run_select.get()
        hours = self.hours_input.get()
        minutes = self.minutes_input.get()
        seconds = self.seconds_input.get()
        miles = self.miles_input.get()

        # Calculating pace
        convert_seconds = seconds / 60
        convert_hours = hours * 60
        total_minutes = minutes + convert_seconds + convert_hours
        pace = round(total_minutes / miles)

        # Creating new row of data based on user inputs
        new_row = pd.DataFrame({'Date': [date], 'Run Type': [run_type], 'Hours': [hours], 'Minutes': [minutes], 'Seconds': [seconds], 'Miles': [miles], 'Pace':[pace]})

        # Adding new row 
        self.df = pd.concat([self.df, new_row], ignore_index=True)

        # Sends new row to excel spreadsheet
        self.df.to_excel("RunLog.xlsx", index=False)

        # FIXME: Figure out how to update index list. Code below not working
        # Updates and gets list of index values
        self.df.reset_index()
        self.index_list = list(self.df.index.values) 
        
        # # Clears and diplays changes in excel textbox
        self.excel_text.delete("1.0", END) 
        self.excel_text.insert(END, self.df)


    # Edit data
    def edit_data(self):
        index_input = self.index_user_input.get()
        select_column = self.col_name.get()
    
        # Depending on column, will convert user input to correct data type
        if select_column == "Date":
            user_edit = self.date_input.get_date()
            datetime_object = datetime.strptime(user_edit, "%m/%d/%Y")
            date = datetime_object.date()
            self.df.at[index_input, select_column] = date # Changes the value at the selected index and column
        elif select_column == "Run Type":
            user_edit = self.run_select.get()
            self.df.at[index_input, select_column] = user_edit
        elif select_column == "Hours":
            user_edit = self.hours_input.get()
            self.df.at[index_input, select_column] = user_edit
        elif select_column == "Minutes":
            user_edit = self.minutes_input.get()
            self.df.at[index_input, select_column] = user_edit
        elif select_column == "Seconds":
            user_edit = self.seconds_input.get()
            self.df.at[index_input, select_column] = user_edit
        elif select_column == "Miles":
            self.df.at[index_input, 'Miles'] = None
            user_edit = self.miles_input.get()
            self.df.at[index_input, select_column] = user_edit

        # Recalculating pace
        convert_seconds = self.df.loc[index_input].at["Seconds"] / 60
        convert_hours = self.df.loc[index_input].at["Hours"] * 60
        total_minutes = self.df.loc[index_input].at["Minutes"] + convert_seconds + convert_hours
        pace = round(total_minutes / self.df.loc[index_input].at["Miles"])
        self.df.at[index_input, "Pace"] = pace

        # Sends edit to excel spreadsheet
        self.df.to_excel("RunLog.xlsx", index=False)

        # Clears and diplays changes in excel textbox
        self.excel_text.delete("1.0", END)
        self.excel_text.insert(END, self.df)


    # Delete (last) row
    def delete_row(self):
        # Getting index of row
        #index_input = self.index_user_input.get()
        last_index = self.df.index[-1]

        # Delete row base on index
        self.df.drop([index_input], inplace=True)

        # Sends edit to excel spreadsheet
        self.df.to_excel("RunLog.xlsx", index=False)

        # FIXME: Figure out how to update index list. Code below not working
        # Getting index values again and make it a list
        #self.df.reset_index()
        #self.index_list = list(self.df.index.values) 

        # Clears and diplays changes in excel textbox
        #self.excel_text.delete("1.0", END)
        #self.excel_text.insert(END, self.df)

    
    def monthly_averages_results(self):

        # Function converts the time values in the DataFrame to total hours and calculates the pace
        def convert_to_TotalHours(df_copy):
            # Make a new column for total time in hours by adding hours, minutes, and seconds as timedeltas    
            df_copy['TotalHours'] = (pd.to_timedelta(df_copy['Hours'], unit='h') + 
                        pd.to_timedelta(df_copy['Minutes'], unit='m') + 
                        pd.to_timedelta(df_copy['Seconds'], unit='s')).dt.total_seconds() / 3600

            # Drop the Hours, Minutes, and Seconds columns as they are no longer needed            
            columns_to_drop = ['Hours', 'Minutes', 'Seconds']
            df_copy.drop(columns=columns_to_drop, inplace=True)

            # Add a Pace column
            df_copy['Pace'] = df_copy['TotalHours'] / df_copy['Miles']
            return df_copy

        # Function that takes a df, a month as an integer and a runtype and returns the daily average miles per month
        def get_monthly_averages(df_copy, month, runtype):
            df_month = df_copy[df_copy.index.month == month]
            if len(df_month) == 0:
                return None, None, None # need 3 None's to match the return type
            runtype_group = df_month.groupby('Run Type')   
            monthly = runtype_group.resample('ME')
            monthly_mean = monthly.mean() 
            row = monthly_mean.loc[runtype]
            return round(row['Miles'].iloc[0], 2), round(row['TotalHours'].iloc[0], 2), round(row['Pace'].iloc[0], 2)

        # make a copy of the DataFrame and set the index to the Date column
        df_copy = self.df.copy()
        df_copy.set_index('Date', inplace=True)

        df_copy = convert_to_TotalHours(df_copy) # convert the time values to total hours and calculate the pace

        for month in range(1, 13):
            for runtype in ['road', 'trail']:
                miles, total_hours, pace = get_monthly_averages(df_copy, month, runtype)
                if miles is not None: # all three are None if there is no data for that month
                    results = (f"Daily average {runtype} miles for Month {month}: {miles} miles\n")
                    # results_list = results
                    print(results) # NOTE: remove after done testing
                    self.calculations_text.insert(END, results)

        
if __name__ == "__main__":
    app = App()
    app.mainloop()

