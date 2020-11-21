import os
import ctypes
from selenium import webdriver
import xlsxwriter
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
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
    year_data += web_driver.find_element_by_xpath(
        f'//*[@id="content"]/div[1]/div/div[2]/table/tbody/tr[{singular}]/td[2]').text + "       " + web_driver.find_element_by_xpath(
        f'//*[@id="content"]/div[1]/div/div[2]/table/tbody/tr[{singular}]/td[3]').text + "\n"

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
                lap_time = round(float(table_data[0:table_data.find(':')]) * 60 + float(
                    table_data[table_data.find(':') + 1: table_data.find('.')]) + float(table_data[-2:]) / 1000, 3)
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

# Filtering the data and copying it into columns with driver as the header
laps -= 2
os.mkdir(
    f"C:\\Users\\Aatmiya Silwal\\PycharmProjects\\WebScraping\\venv\\Formula One Project\\{year} Round{race_number}\\Driver Distributions")

filepath = f"C:\\Users\\Aatmiya Silwal\\PycharmProjects\\WebScraping\\venv\\Formula One Project\\{year} Round{race_number}\\{year} Round{race_number}.xlsx"

sheets = pd.read_excel(filepath, header=0, sheet_name="Race Data", usecols=[0, 1, 2], index_col=False)
race_data_df = pd.DataFrame(sheets)
minimum_limit = min(list(race_data_df['Lap Times'])) - 1

driver_list = [x for x in sheets["Drivers"].unique() if str(x) != 'nan']

lap_time_df = positions_df = pd.DataFrame()
lap_time_df["Drivers"] = positions_df["Drivers"] = list(range(1, laps + 1))

plt.figure(figsize=(19, 16))

for driver in driver_list:
    fil_xl = race_data_df[race_data_df.iloc[:, 0] == driver]["Lap Times"]
    temp_list = list(fil_xl)

    position_fil_xl = race_data_df[race_data_df.iloc[:, 0] == driver]["Position"]
    position_temp_list = list(position_fil_xl)

    while len(position_temp_list) != laps:
        position_temp_list.append(list(fil_xl)[-1])
        temp_list.append(0)

    positions_df[driver] = np.array(position_temp_list)
    positions_plot = sns.lineplot(data=positions_df, x="Drivers", y=driver)

    lap_time_df[driver] = np.array(temp_list)

# Settings for position plot optimization
positions_plot.yaxis.set_major_locator(ticker.MultipleLocator(1))
positions_plot.xaxis.set_major_locator(ticker.MultipleLocator(5))
positions_plot.set_ylim(0, len(driver_list) + 1)
positions_plot.set_xlim(left=1, right=laps)
positions_plot.invert_yaxis()
positions_plot.set_xlabel("Laps", size=10)
positions_plot.set_ylabel("Position", size=10)
plt.title("Position Change Chart", size=15)
plt.legend(labels=driver_list, loc="center", bbox_to_anchor=(1.024, -0.4, .08, 1.9), fontsize=9)
position_change = positions_plot.get_figure()
position_change.savefig(
    f"C:\\Users\\Aatmiya Silwal\\PycharmProjects\\WebScraping\\venv\\Formula One Project\\{year} Round{race_number}\\Driver Distributions\\Position Change Graph.png")
plt.close()

# Driver-wise KDE plot generation
for one_driver in driver_list:
    x_val = list(lap_time_df[one_driver].values)
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
        distribution.set_xlim(left=minimum_limit)
        plt.title(f"{one_driver[0].upper()}{one_driver[1:]}")
        driver_plot = distribution.get_figure()
        driver_plot.savefig(
            f"C:\\Users\\Aatmiya Silwal\\PycharmProjects\\WebScraping\\venv\\Formula One Project\\{year} Round{race_number}\\Driver Distributions\\{one_driver[0].upper()}{one_driver[1:]}.png")
        plt.close()
