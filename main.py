# Import necessary libraries
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# --- Step 1: Load the Dataset ---
# Load the supply chain data from the Excel file into a pandas DataFrame.
# Note: Reading Excel files requires the 'openpyxl' library.
# You can install it by running: pip install openpyxl
#
# Make sure the 'q-excel-correlation-heatmap.xlsx' file is in the same directory as this script.
try:
    df = pd.read_excel('q-excel-correlation-heatmap.xlsx')
    print("Successfully loaded the dataset.")
except FileNotFoundError:
    print("Error: 'q-excel-correlation-heatmap.xlsx' not found.")
    print("Please make sure the data file is in the same folder as the script.")
    exit()
except Exception as e:
    print(f"An error occurred: {e}")
    print("Please ensure you have the 'openpyxl' library installed (`pip install openpyxl`).")
    exit()

# --- Step 2: Calculate the Correlation Matrix ---
# Use the .corr() method on the DataFrame to compute the pairwise correlation of columns.
correlation_matrix = df.corr()
print("\nCorrelation Matrix:")
print(correlation_matrix)

# --- Step 3: Export the Correlation Matrix to CSV ---
# Save the calculated correlation matrix to a file named 'correlation.csv'.
# The 'index=True' argument ensures that the row labels (the variable names) are included.
correlation_matrix.to_csv('correlation.csv', index=True)
print("\nSuccessfully saved correlation matrix to 'correlation.csv'.")

# --- Step 4: Generate and Save the Heatmap ---
# Create a heatmap to visualize the correlation matrix.
plt.figure(figsize=(8, 6)) # Set the figure size for better readability

# Use seaborn's heatmap function.
# - 'correlation_matrix': The data to plot.
# - 'annot=True': Display the correlation values on the heatmap.
# - 'cmap='coolwarm'': Use a color map where low values are blue (cool) and high values are red (warm).
#   This is similar to the Red-White-Green scale, showing negative and positive correlations clearly.
# - 'fmt=".2f"': Format the annotations to two decimal places.
sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', fmt=".2f", linewidths=.5)

# Add a title to the plot.
plt.title('Supply Chain Metrics Correlation Matrix')

# Ensure the plot layout is tight and clean.
plt.tight_layout()

# Save the generated heatmap as a PNG image file.
plt.savefig('heatmap.png', dpi=300) # Use a high DPI for better quality
print("Successfully generated and saved heatmap to 'heatmap.png'.")

# --- (Optional) Step 5: Display the Plot ---
# If you are running this script in an environment that supports GUI,
# this will show the plot in a new window.
# plt.show()
