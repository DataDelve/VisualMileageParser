#!/usr/bin/env python3
# coding: utf-8
# Written by Jon Hoover

"""
VisualMileageParser

Dependencies:
PyQt6
pandas
openpyxl
"""

import PyQt6.QtWidgets as qtw
import PyQt6.QtGui as qtg
import pandas as pd

class MainWindow(qtw.QWidget):
    def __init__(self):
        super().__init__()
        # Add a title
        self.setWindowTitle("Visual Mileage Parser")

        # Set Vertical layout
        self.setLayout(qtw.QVBoxLayout())

        # Create a label
        my_label = qtw.QLabel("Paste mileage from Timeero here:")

        # Change the fot size of label
        my_label.setFont(qtg.QFont("Verdana", 14))
        self.layout().addWidget(my_label)

        # Create an entry box
        self.my_entry = qtw.QTextEdit()
        self.my_entry.setObjectName("text_field")
        self.my_entry.setText("")
        self.layout().addWidget(self.my_entry)

        # Create another label
        label2 = qtw.QLabel("!! WARNING !! This will overwrite any existing Excel file!")
        label2.setFont(qtg.QFont("Verdana", 18))
        label2.setStyleSheet("QLabel {color: #FF0000;}")
        self.layout().addWidget(label2)

        # Create a button
        my_button = qtw.QPushButton("Export to MS Excel!")
        my_button.clicked.connect(self.press_it)
        self.layout().addWidget(my_button)

        self.show()

    def press_it(self):
        # Create a list[] and store input, by line
        input_text = str(self.my_entry.toPlainText()).split("\n")
        # Clear text imput box
        self.my_entry.setText("")

        # Begin MileageParser

        # Parse data and create initial dataframe
        timeEntry = []
        mileageDF = pd.DataFrame(columns=["branch", "time_in", "date_in", "time_out",
                                          "date_out", "duration", "length"])

        for line in input_text:
            # remove tab and newline from line
            line = line.replace("\t", "")
            line = line.replace("\n", "")

            # remove invalid data and blank lines
            if line in ["CST", "", "-"]:
                continue

            # The last line in each entry always contains the word "miles".
            # Therefore, we check if "miles" in line; then we check if the entire entry
            # is valid, where if it has 7 items in the list, then it is valid.
            # If not, we reject the entire entry and move on.
            if "miles" in line:
                timeEntry.append(line)

                if len(timeEntry) == 7:
                    mileageDF.loc[len(mileageDF)] = timeEntry
                    timeEntry.clear()
                else:
                    # print("Ignoring invalid entry: " + str(timeEntry))
                    timeEntry.clear()
                    continue
            else:
                timeEntry.append(line)

        #mileageDF.to_csv("rawdata.csv")

        # Clean data

        # combine date and time columns, then convert to datetime format
        mileageDF["datetime_in"] = pd.to_datetime((mileageDF["date_in"] + " " + mileageDF["time_in"]),
                                                  format="%b %d, %Y %I:%M %p")
        mileageDF["datetime_out"] = pd.to_datetime((mileageDF["date_out"] + " " + mileageDF["time_out"]),
                                                   format="%b %d, %Y %I:%M %p")

        # Change date format to remove the commas
        for col in ["date_in", "date_out"]:
            mileageDF[col] = pd.to_datetime(mileageDF[col], format="%b %d, %Y")
            mileageDF[col] = mileageDF[col].dt.strftime("%Y-%m-%d")

        for col in ["time_in", "time_out"]:
            mileageDF[col] = pd.to_datetime(mileageDF[col], format="%I:%M %p")
            mileageDF[col] = mileageDF[col].dt.strftime("%I:%M %p")

        # Remove rows where duration = "0:00"
        mileageDF = mileageDF.loc[mileageDF["duration"] != "0:00"]

        # sort rows by date and time
        mileageDF = mileageDF.sort_values(by="datetime_in")
        mileageDF.reset_index(drop=True, inplace=True)

        # Determine actual mileage

        # Get list of unique dates from dataframe
        uniqueDates = mileageDF["date_in"].unique()

        # Timeero does not accurately calculate the mileage between locations.
        # Therefore, Mileage Chart (contains updated mileage for all locations) will be
        # referenced and the correct mileage will be generated on the final report.

        # Enter correct mileage for each location
        finalDF = pd.DataFrame(columns=["Date", "From Branch", "To Branch", "Distance"])
        from_branch = ""
        to_branch = ""
        distance = 0.0

        # create dictionary of library locations for conversion
        conv_dict = {"Main Library": "MAIN",
                     "Birmingham Branch": "BIRM",
                     "Heatherdowns Branch": "HED",
                     "Holland Branch": "HOLL",
                     "Kent Branch": "KENT",
                     "Mobile Services": "KINGRD",
                     "King Road Branch": "KINGRD",
                     "Lagrange Branch": "LAG",
                     "Locke Branch": "LOCKE",
                     "Maumee Branch": "MAUM",
                     "Mott Branch": "MOTT",
                     "Oregon Branch": "OREG",
                     "Friends of the Library": "WAREHOUSE",
                     "Point Place Branch": "PTPL",
                     "Reynolds Corners Branch": "RC",
                     "Sanger Branch": "SANG",
                     "South Branch": "SOUTH",
                     "Sylvania Branch": "SYLV",
                     "Toledo Heights Branch": "TH",
                     "Washington Branch": "WASH",
                     "Waterville Branch": "WATV",
                     "West Toledo Branch": "WTOL",
                     "Cherry Street Mission": "MAIN"}

        # Create dataframe for mileage chart
        chartDF = pd.read_csv("mileage-chart.csv")
        chartDF = chartDF.set_index("LOCATION")


        # Function returns correct mileage from Mileage Chart.
        def get_mileage(from_loc, to_loc):
            from_conv = conv_dict[from_loc]
            to_conv = conv_dict[to_loc]
            return chartDF.loc[from_conv][to_conv]


        # Take entries for each date and calculate starting and ending mileage for each
        # location on that date. Generate data for final report.
        for date in uniqueDates:
            selected_rows = mileageDF.loc[mileageDF["date_in"] == date]

            sr_len = len(selected_rows)

            if sr_len != 1:
                for x in range(0, sr_len):
                    if x == 0:
                        continue
                    elif x == 1:
                        # print(selected_rows)
                        from_branch = selected_rows.iloc[0, 0]
                        to_branch = selected_rows.iloc[1, 0]

                        if from_branch != to_branch:
                            distance = get_mileage(from_branch, to_branch)

                            if distance == 0.00:
                                continue

                            finalDF.loc[len(finalDF.index)] = [date, from_branch,
                                                               to_branch, distance]
                    else:
                        from_branch = to_branch
                        to_branch = selected_rows.iloc[x, 0]

                        if from_branch != to_branch:
                            distance = get_mileage(from_branch, to_branch)

                            if distance == 0.00:
                                continue

                            finalDF.loc[len(finalDF.index)] = [date, from_branch,
                                                               to_branch, distance]

        # Create final report
        # Note: module "openpyxl" must be installed.
        #     pip install openpyxl
        finalDF.to_excel("final-mileage.xlsx", index=False, freeze_panes=(1, 0))
        self.my_entry.setFont(qtg.QFont("Courier", 10))
        self.my_entry.setText("File saved as: final-mileage.xlsx")


app = qtw.QApplication([])
mw = MainWindow()

# Run the app
app.exec()
