import os
import ctypes
from selenium import webdriver
import xlsxwriter
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
import matplotlib.pyplot as plt
import seaborn as sns

# Locating the webdriver
path = "C:\Program Files (x86)\chromedriver.exe"
web_driver = webdriver.Chrome(path)

# User input for year + race number
year = input("Enter the year: ")
web_driver.get(f"http://ergast.com/api/f1/{year}")

info_row = len(web_driver.find_elements_by_xpath('//*[@id="content"]/div[1]/div/div[2]/table/tbody/tr')) - 2

year_data = ""
for singular in range(3, info_row + 3):
    year_data += web_driver.find_element_by_xpath(f'//*[@id="content"]/div[1]/div/div[2]/table/tbody/tr[{singular}]/td[2]').text + "       " + web_driver.find_element_by_xpath(f'//*[@id="content"]/div[1]/div/div[2]/table/tbody/tr[{singular}]/td[3]').text + "\n"

ctypes.windll.user32.MessageBoxW(0, year_data, f"{year} Formula 1 Season", 64)
race_number = input("Enter the race number: ")

# Create file according to user input
os.mkdir(
    f"C:\\Users\\Aatmiya Silwal\\PycharmProjects\\WebScraping\\venv\\Formula One Project\\{year} Round{race_number}")
outWorkbook = xlsxwriter.Workbook(
    f"C:\\Users\\Aatmiya Silwal\\PycharmProjects\\WebScraping\\venv\\Formula One Project\\{year} Round{race_number}\\{year} Round{race_number}.xlsx")
outSheet = outWorkbook.add_worksheet("Race Data")

insert_row = 1
valid_bool = True
laps = 1

# Headers for the dataframe
outSheet.write(0, 0, "Drivers")
outSheet.write(0, 1, "Position")
outSheet.write(0, 2, "Lap Times")

# Open webpage for lap-wise data and writing it into the worksheet
while valid_bool:
    web_driver.get(f"http://ergast.com/api/f1/{year}/{race_number}/laps/{str(laps)}")
    rows = len(web_driver.find_elements_by_xpath('//*[@id="content"]/div[1]/div/div[2]/table[2]/tbody/tr'))
    columns = len(web_driver.find_elements_by_xpath('//*[@id="content"]/div[1]/div/div[2]/table[2]/tbody/tr[3]/th'))
    if columns == 0:
        valid_bool = False

    for r in range(4, rows + 1):
        insert_col = 0
        for c in range(1, columns + 1):
            table_data = web_driver.find_element_by_xpath(
                f'//*[@id="content"]/div[1]/div/div[2]/table[2]/tbody/tr[{str(r)}]/td[{str(c)}]').text
            if c == 3:
                lap_time = round(float(table_data[0]) * 60 + float(table_data[2:4]) + float(table_data[-3:]) / 1000, 3)
                outSheet.write(insert_row, insert_col, lap_time)
            elif c == 2:
                outSheet.write(insert_row, insert_col, int(table_data))
            else:
                outSheet.write(insert_row, insert_col, table_data)
            insert_col += 1
        insert_row += 1

    insert_row += 2
    laps += 1
web_driver.quit()
outWorkbook.close()

laps -= 2
# Filtering the data and copying it into columns with driver as the header

filepath = f"C:\\Users\\Aatmiya Silwal\\PycharmProjects\\WebScraping\\venv\\Formula One Project\\{year} Round{race_number}\\{year} Round{race_number}.xlsx"
writer = pd.ExcelWriter(
    f"C:\\Users\\Aatmiya Silwal\\PycharmProjects\\WebScraping\\venv\\Formula One Project\\{year} Round{race_number}\\Test.xlsx",
    engine='xlsxwriter')

sheets = pd.read_excel(filepath, header=0, sheet_name="Race Data", usecols=[0, 2], index_col=False)
race_data_df = pd.DataFrame(sheets)

driver_list = [x for x in sheets["Drivers"].unique() if str(x) != 'nan']
fil_xl = race_data_df[race_data_df.iloc[:, 0] == driver_list[0]]

lap_time_df = pd.DataFrame()
lap_time_df["Drivers"] = list(range(1, laps + 1))

for driver in driver_list:
    fil_xl = race_data_df[race_data_df.iloc[:, 0] == driver]["Lap Times"]
    temp_list = list(fil_xl)
    while len(temp_list) != laps:
        temp_list.append(0)
    lap_time_df[driver] = np.array(temp_list)

lap_time_df.to_excel(writer, sheet_name="Lap Data")
writer.save()

# Driver-wise KDE plot generation
os.mkdir(
    f"C:\\Users\\Aatmiya Silwal\\PycharmProjects\\WebScraping\\venv\\Formula One Project\\{year} Round{race_number}\\Driver Distributions")
raw_data = pd.read_excel(
    f"C:\\Users\\Aatmiya Silwal\\PycharmProjects\\WebScraping\\venv\\Formula One Project\\{year} Round{race_number}\\Test.xlsx")

for one_driver in driver_list:
    x_val = list(raw_data[one_driver].values)
    x_val.sort()
    temp = []

    for items in x_val:
        if np.percentile(x_val, 80) > items > 0:
            temp.append(items)

    if len(temp) > (laps / 3):
        distribution = sns.kdeplot(pd.DataFrame(temp, columns=["Lap Times"])["Lap Times"].values, x="Lap Times",
                                   fill=True, color=(tuple(np.random.random_sample(3, ))))
        distribution.set_xlabel("Time (seconds)")
        distribution.set(yticklabels=[])
        plt.title(f"{one_driver[0].upper()}{one_driver[1:]}")
        driver_figure = distribution.get_figure()
        driver_figure.savefig(
            f"C:\\Users\\Aatmiya Silwal\\PycharmProjects\\WebScraping\\venv\\Formula One Project\\{year} Round{race_number}\\Driver Distributions\\{one_driver[0].upper()}{one_driver[1:]}.png")
        plt.close()

os.remove(f"C:\\Users\\Aatmiya Silwal\\PycharmProjects\\WebScraping\\venv\\Formula One Project\\{year} Round{race_number}\\Test.xlsx")
