import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

def makeHeatmap(filepath, sheet_name, title):
    sleep_data = pd.read_excel(filepath, sheet_name=sheet_name) # read the file and the sheet in 
    sleep_data.drop(sleep_data.columns[:2], axis=1, inplace=True) # drop the dates so heatmap can work
    ax = sns.heatmap(sleep_data) # create a heatmap of the data
    ax.set_title(title) # decide the title of the data
    plt.show() #show the heatmap

# example usage
makeHeatmap('newFall23rawSleepData.xlsx', 'TotalMinutesAsleep', 'Original Sleep Dataset Fall 2023')
