import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

print("Starting Excel data analysis script...")

try:
    # Excel file path
    file_path = r"C:\Users\100995420\Desktop\May16Data.xlsx"
    print(f"Attempting to read Excel file: {file_path}")
    
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name="Refined Gerdau Data")
    
    # Display basic info
    print("\nSuccessfully loaded the Excel file!")
    print(f"Dimensions: {df.shape[0]} rows, {df.shape[1]} columns")
    
    print("\nFirst 5 rows of the data:")
    print(df.head())
    
    print("\nColumn names in the dataset:")
    print(df.columns.tolist())
    
    # Save results to a text file for easier viewing
    with open("analysis_results.txt", "w") as f:
        f.write("DATA ANALYSIS RESULTS\n\n")
        f.write("First 5 rows:\n")
        f.write(df.head().to_string())
        f.write("\n\nColumn names:\n")
        f.write(str(df.columns.tolist()))
        f.write("\n\nBasic statistics:\n")
        f.write(df.describe().to_string())
    
    print("\nSaved detailed results to 'analysis_results.txt'")
    
    # Create and save a simple visualization
    print("\nCreating visualizations...")
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    if len(numeric_cols) > 0:
        plt.figure(figsize=(10, 6))
        for i, col in enumerate(numeric_cols[:3]):  # First 3 numeric columns
            plt.subplot(1, 3, i+1)
            plt.hist(df[col].dropna(), bins=20)
            plt.title(col)
        plt.tight_layout()
        plt.savefig('data_visualization.png')
        print("Saved visualization to 'data_visualization.png'")
    
    print("\nAnalysis completed successfully!")
    
except Exception as e:
    print(f"Error occurred: {e}")

# Keep the window open
input("\nPress Enter to exit...")

