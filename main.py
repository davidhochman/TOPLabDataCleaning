from IPython.display import display
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt


def read_data(filename, sheetname, drop):
    # Prepare to parse Excel sheets into DataFrame
    xls = pd.ExcelFile('./data/' + filename)
    # Read data into DataFrame from appropriate sheet
    df = pd.read_excel(xls, sheetname)

    if drop is True:
        # Drop column 0 because it reads in the labels
        df.drop(df.columns[0], axis=1, inplace=True)

    return df


def create_histogram(df, filename):
    # Fill in NaN with zeros first then retrieve number of nonzero values in each column
    series = df.fillna(0).astype(bool).sum(axis=0)

    # Get data recorded percentage of each participant
    series = (series / len(df)) * 100

    plt.xlabel("Percentage of Participant Data")
    plt.ylabel("Frequency")

    sns.histplot(series)
    plt.savefig("./images/" + filename, dpi=600)
    plt.show()


# Calling the functions
if __name__ == '__main__':
    # Reading in files
    df1 = read_data('newFall23rawSleepData.xlsx', 'TotalTimeInBed', True)
    df2 = read_data('newSpring23rawSleepData.xlsx', 'TotalTimeInBed', True)
    df3 = read_data('SP24_rawSleepFormatted.xlsx', 'TotalTimeInBed', True)

    create_histogram(df1, 'F23_Histogram')
    create_histogram(df2, 'SP23_histogram')
    create_histogram(df3, 'SP24_histogram')

