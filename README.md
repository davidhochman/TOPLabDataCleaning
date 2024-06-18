Here’s how the EM method is applied in the code:

Initialization (Expectation Step):

The initial missing values are filled with estimates (e.g., mean or median of the feature).
Imputation (Maximization Step):

For each feature with missing values, the IterativeImputer fits a regression model using the other features.
Missing values are then predicted (imputed) using this model.
Iteration:

The process is repeated iteratively, using the newly imputed values to update the models until convergence or the maximum number of iterations is reached.

To reference the code:
By using IterativeImputer, the code leverages the EM method to handle missing data, ensuring a more robust and statistically sound imputation process compared to simpler methods like mean or median imputation.

  This Python script is designed to preprocess and clean a dataset of sleep activity from Fitbit users, stored in Excel files. It automates the process of handling multiple sheets within each Excel file using Iterative Imputer which imputes missing values by iteratively modeling each feature with missing entries as a function of other features, using a sequential, cyclic approach.
  
  The script begins by reading all sheets from an input Excel file into a dictionary of DataFrames. It replaces zero values with NaN and restores the 0th participant. It then removes columns that have more than 70% zeros, and applies the model “IterativeImputer” to fill missing values. It then caps values based on Z-scores to reduce outliers and applies winsorization (which limits extreme values), to normalize the data distribution.The script writes the cleaned data back to Excel, adds “CLEANED(EM)” to the title, and keeps everything else the same. It formats the content, and saves it in a specific folder.

Example Usage:
```
import pandas as pd
import numpy as np
data = {
    'A': [1, 2, 3, 4, 100],  # 100 is a potential outlier
    'B': [1.2, 2.5, 3.5, 2.1, 1.9]
}
df = pd.DataFrame(data)
# Cap outliers using the default threshold
df_no_outliers = cap_outliers_zscore(df)
# Cap outliers using a custom threshold
df_custom_threshold = cap_outliers_zscore(df, threshold=2.5)
```
2. winsorize_data(df, limits=[0.05, 0.05])
Description: Applies winsorization to the DataFrame to limit the influence of extreme values at both the lower and upper tails of the data distribution.
Example Usage:
```
df_winsorized = winsorize_data(df)
# Applying winsorization with different limits
df_custom_limits = winsorize_data(df, limits=[0.01, 0.01])  # More aggressive winsorization
```
3. process_excel_file(file_path, output_folder, stddev_folder, heatmap_folder)
Description: Processes a single Excel file by imputing missing values, handling outliers, and saving the cleaned data to a new Excel file.
Example Usage:
```
file_path = 'C:\\path_to_excel\\sample_data.xlsx'
output_folder = 'C:\\path_to_output\\cleaned_data'
stddev_folder = 'C:\\path_to_output\\stddev'
heatmap_folder = 'C:\\path_to_output\\heatmaps'
process_excel_file(file_path, output_folder, stddev_folder, heatmap_folder)
```
4. process_all_excel_files(input_folder, output_folder, stddev_folder, heatmap_folder)
Description: Processes all Excel files within a specified directory, applying data cleaning operations to each file and saving the results.


Example Usage:
```
input_folder = 'C:\\path_to_excel_files'
output_folder = 'C:\\path_to_output\\cleaned_data'
stddev_folder = 'C:\\path_to_output\\stddev'
heatmap_folder = 'C:\\path_to_output\\heatmaps'
process_all_excel_files(input_folder, output_folder, stddev_folder, heatmap_folder)
```



