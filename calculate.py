#!/usr/bin/env python3
"""
Excel Processor for Gas Flow Analysis

This script processes an Excel file containing gas flow data, performs calculations based on 
the Shomate equation and gas flow parameters, and outputs a new Excel file with the calculated values.

Usage:
    python excel_processor.py input_excel_file.xlsx [output_excel_file.xlsx]

If output_excel_file.xlsx is not provided, it will default to "processed_output.xlsx"
"""

import pandas as pd
import numpy as np
import sys
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment

def process_excel_file(diameter, radius, flow_rate, input_file, output_file="processed_output.xlsx"):
    """
    Process the Excel file with gas flow data and output a new file with calculated values.
    
    Args:
        input_file (str): Path to the input Excel file
        output_file (str): Path to the output Excel file
    """
    print(f"Processing {input_file}...")
    
    try:
        # 1. Open the excel file - try to read the "Refined Gerdau Data" sheet
        try:
            df = pd.read_excel(input_file, sheet_name="Refined Gerdau Data", header=15)
            print("Reading 'Refined Gerdau Data' sheet...")
        except:
            # If the sheet doesn't exist, read the first sheet
            print("'Refined Gerdau Data' sheet not found. Reading the first sheet...")
            df = pd.read_excel(input_file, header=15)
        
        # 2. Define constants from the PDF
        # Constants from the PDF that will be used in calculations
        constants = {
            # Pressure (atm)
            'pressure': 1,
            # R constant (atm*L/mol*K)
            'R_constant': 0.0821,
            # Reference temperature T1 (°C)
            'T1': 300,
            # Molecular weights
            'MW_CO2': 44.01,
            'MW_H2O': 18.01528,
            'MW_N2': 28.0134,
            'MW_O2': 31.999,
            'MW_CO': 28.01,
            'MW_NOx': 30.0061,
            # Constants for energy calculations
            'energy_factor_CO2': 1.2135,
            'energy_factor_H2O': 1.996,
            'energy_factor_N2': 1.158,
            'energy_factor_O2': 1.087552941,
            'energy_factor_CO': 1.23723875,
            'energy_factor_NOx': 1.183859,
            # Gas composition factors from the formulas
            'factor_CO2': 1.03552318948896,
            'factor_H2O': 3.76603433650984,
            'factor_N2': 17.440390559814,
            'factor_O2': 2.51572679168908,
            'factor_CO': 0.000976810511041857,
            'factor_NOx': 0.011187614171608,
        }
        
        # Check for expected columns and add them if they don't exist
        required_columns = ['Date', 'End Date', 'Time', 'NG High Flow (NCMH)', 
                           'NG Low Flow (NCMH)', 'Flue gas exhaust temperature (ºC)']
        
        # Rename columns if needed to match the expected format
        if 'Date' not in df.columns and 'date' in df.columns:
            df.rename(columns={'date': 'Date'}, inplace=True)
            
        for col in required_columns:
            if col not in df.columns:
                print(f"Warning: Column '{col}' not found in the input file. Adding empty column.")
                df[col] = np.nan
        
        # 3. Calculate the natural gas flow rate (SCM/second)
        # Formula: =(F17+E17)/3600
        print("Calculating natural gas flow rate...")
        df['Natural GasFlow rate (SCM/second)'] = (df['NG Low Flow (NCMH)'] + df['NG High Flow (NCMH)']) / 3600
        
        # 4. Calculate the molar flow rate (moles/second)
        # Formula: =((101.325)*(H17)*1000)/((8.314)*(273.15))
        print("Calculating molar flow rate...")
        df['n (Moles/second)'] = ((101.325) * df['Natural GasFlow rate (SCM/second)'] * 1000) / ((8.314) * (273.15))
        
        # 5. Calculate delta T in Kelvin
        # Formula: =(G17+273.15)-($J$7+273.15)
        print("Calculating temperature differences...")
        df['Delta T (K)'] = (df['Flue gas exhaust temperature (ºC)'] + 273.15) - (constants['T1'] + 273.15)
        
        # 6. Calculate delta T in degrees Celsius
        # Formula: =G17-$J$7
        df['Delta T (ºC)'] = df['Flue gas exhaust temperature (ºC)'] - constants['T1']
        
        # 7. Calculate the flow rate of the gases in kj/s
        # Formula: =(K17)*2.22*(K17)
        print("Calculating natural gas energy flow...")
        df['Q (NG) (kJ/s)'] = df['Delta T (K)'] * 2.22 * df['Delta T (K)']
        
        # 8-9. Calculate gas component molar flow rates, mass flow rates, and energy contribution
        print("Calculating gas component flows and energies...")
        
        # CO2 calculations
        # Formula: =(1.03552318948896)*(I17)
        df['n (CO2)(Moles/second)'] = constants['factor_CO2'] * df['n (Moles/second)']
        
        # Formula: =M17*(44.01)
        df['m (CO2) (g/s)'] = df['n (CO2)(Moles/second)'] * constants['MW_CO2']
        
        # Formula: =(K17)*(N17/1000)*(1.2135)
        df['Q (CO2) (kJ/s)'] = df['Delta T (K)'] * (df['m (CO2) (g/s)'] / 1000) * constants['energy_factor_CO2']
        
        # H2O calculations
        # Formula: =(3.76603433650984)*(I17)
        df['n (H2O)(Moles/second)'] = constants['factor_H2O'] * df['n (Moles/second)']
        
        # Formula: =(P17)*(18.01528)
        df['m (H2O) (g/s)'] = df['n (H2O)(Moles/second)'] * constants['MW_H2O']
        
        # Formula: =(K17)*(Q17/1000)*(1.996)
        df['Q (H2O) (kJ/s)'] = df['Delta T (K)'] * (df['m (H2O) (g/s)'] / 1000) * constants['energy_factor_H2O']
        
        # N2 calculations
        # Formula: =(17.440390559814)*(I17)
        df['n (N2)(mo/s)'] = constants['factor_N2'] * df['n (Moles/second)']
        
        # Formula: =(S17)*(28.0134)
        df['m (N2) (g/s)'] = df['n (N2)(mo/s)'] * constants['MW_N2']
        
        # Formula: =(K17)*(T17/1000)*(1.158)
        df['Q (N2) (kJ/s)'] = df['Delta T (K)'] * (df['m (N2) (g/s)'] / 1000) * constants['energy_factor_N2']
        
        # O2 calculations
        # Formula: =(2.51572679168908)*(I17)
        df['n (O2)(mol/s)'] = constants['factor_O2'] * df['n (Moles/second)']
        
        # Formula: =(V17)*(31.999)
        df['m (O2) (g/s)'] = df['n (O2)(mol/s)'] * constants['MW_O2']
        
        # Formula: =(K17)*(W17/1000)*(1.087552941)
        df['Q (O2) (kJ/s)'] = df['Delta T (K)'] * (df['m (O2) (g/s)'] / 1000) * constants['energy_factor_O2']
        
        # CO calculations
        # Formula: =(0.000976810511041857)*(I17)
        df['n (CO)(mol/s)'] = constants['factor_CO'] * df['n (Moles/second)']
        
        # Formula: =(Y17)*(28.01)
        df['m (CO) (g/s)'] = df['n (CO)(mol/s)'] * constants['MW_CO']
        
        # Formula: =(K17)*(Z17/1000)*(1.23723875)
        df['Q (CO) (kJ/s)'] = df['Delta T (K)'] * (df['m (CO) (g/s)'] / 1000) * constants['energy_factor_CO']
        
        # NOx calculations
        # Formula: =(0.011187614171608)*(I17)
        df['n (NOx)(mol/s)'] = constants['factor_NOx'] * df['n (Moles/second)']
        
        # Formula: =(AB17)*(30.0061)
        df['m (NOx) (g/s)'] = df['n (NOx)(mol/s)'] * constants['MW_NOx']
        
        # Formula: =(K17)*(AC17/1000)*(1.183859)
        df['Q (NOx) (kJ/s)'] = df['Delta T (K)'] * (df['m (NOx) (g/s)'] / 1000) * constants['energy_factor_NOx']
        
        # Calculate sum of exhaust gases energy
        # Formula: =SUM(O17,R17,U17,X17,AA17,AD17)
        print("Calculating total energy in exhaust gases...")
        df['Sum of the Energy in exhaust gases (kJ/s)'] = (
            df['Q (CO2) (kJ/s)'] + df['Q (H2O) (kJ/s)'] + df['Q (N2) (kJ/s)'] + 
            df['Q (O2) (kJ/s)'] + df['Q (CO) (kJ/s)'] + df['Q (NOx) (kJ/s)']
        )
        
        # Calculate volumetric flow rates for each gas component
        print("Calculating volumetric flow rates...")
        
        # Formula: =(((M17)*($N$5)*(G17+273.15))*0.001)/($N$4)
        df['Q (C02, m^3/s)'] = (
            ((df['n (CO2)(Moles/second)'] * constants['R_constant'] * (df['Flue gas exhaust temperature (ºC)'] + 273.15)) * 0.001)
            / constants['pressure']
        )
        
        # Formula: =((((P17)*($N$5)*(G17+273.15))*0.001))/($N$4)
        df['Q (H2O, m^3/s)'] = (
            ((df['n (H2O)(Moles/second)'] * constants['R_constant'] * (df['Flue gas exhaust temperature (ºC)'] + 273.15)) * 0.001)
            / constants['pressure']
        )
        
        # Formula: =(((S17)*(0.0821)*(G17+273.15))*0.001)
        df['Q (N2, m^3/s)'] = (
            ((df['n (N2)(mo/s)'] * constants['R_constant'] * (df['Flue gas exhaust temperature (ºC)'] + 273.15)) * 0.001)
        )
        
        # Formula: =((((V17)*($N$5)*(G17+273.15))*0.001))/($N$4)
        df['Q (O2, m^3/s)'] = (
            ((df['n (O2)(mol/s)'] * constants['R_constant'] * (df['Flue gas exhaust temperature (ºC)'] + 273.15)) * 0.001)
            / constants['pressure']
        )
        
        # Formula: =((((Y17)*($N$5)*(G17+273.15))*0.001))/($N$4)
        df['Q (CO, m^3/s)'] = (
            ((df['n (CO)(mol/s)'] * constants['R_constant'] * (df['Flue gas exhaust temperature (ºC)'] + 273.15)) * 0.001)
            / constants['pressure']
        )
        
        # Formula: =((((AB17)*($N$5)*(G17+273.15))*0.001))/($N$4)
        df['Q (NOx, m^3/s)'] = (
            ((df['n (NOx)(mol/s)'] * constants['R_constant'] * (df['Flue gas exhaust temperature (ºC)'] + 273.15)) * 0.001)
            / constants['pressure']
        )
        
        # Calculate sum of volumetric flow rates
        # Formula: =SUM(AF17,AG17,AH17,AI17,AJ17,AK17)
        df['Sum of the Volumetric Flow Rates (m^3/s)'] = (
            df['Q (C02, m^3/s)'] + df['Q (H2O, m^3/s)'] + df['Q (N2, m^3/s)'] + 
            df['Q (O2, m^3/s)'] + df['Q (CO, m^3/s)'] + df['Q (NOx, m^3/s)']
        )
        
        # Calculate the velocity of the exhaust gases based on area
        # We assume area is calculated as PI()*((Diameter/2)^2)
        # Where Diameter = (Diameter_ft/3.281)-(Refractory_inches/39.37)
        # From the PDF: Diameter (ft) = 6, Refractory (inches) = 6
        diameter_ft = 6
        refractory_inches = 6
        diameter_m = (diameter_ft / 3.281) - (refractory_inches / 39.37)
        area_m2 = np.pi * ((diameter_m / 2) ** 2)
        
        # Formula: =AL17/$P$3
        print("Calculating gas velocities...")
        df['Velocity of the exhaust gases (m/s)'] = df['Sum of the Volumetric Flow Rates (m^3/s)'] / area_m2
        
        # Calculate summary statistics
        print("Calculating summary statistics...")
        summary = {
            'Maximum Energy (kJ/s)': df['Sum of the Energy in exhaust gases (kJ/s)'].max(),
            'Minimum Energy (kJ/s)': df['Sum of the Energy in exhaust gases (kJ/s)'].min(),
            'Average Energy (kJ/s)': df['Sum of the Energy in exhaust gases (kJ/s)'].mean(),
            'Average Volumetric Flow Rate (m^3/s)': df['Sum of the Volumetric Flow Rates (m^3/s)'].mean(),
            'Maximum Velocity (m/s)': df['Velocity of the exhaust gases (m/s)'].max(),
            'Minimum Velocity (m/s)': df['Velocity of the exhaust gases (m/s)'].min(),
            'Average Velocity (m/s)': df['Velocity of the exhaust gases (m/s)'].mean(),
            'Upper bound temperature (°C)': df['Flue gas exhaust temperature (ºC)'].max(),
            'Lower bound temperature (°C)': df['Flue gas exhaust temperature (ºC)'].min(),
            'Average temperature (°C)': df['Flue gas exhaust temperature (ºC)'].mean(),
        }
        
        # Create a new workbook using openpyxl for more control over formatting
        wb = Workbook()
        ws = wb.active
        ws.title = "Processed Data"
        
        # Write the data
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        
        # Add summary worksheet
        ws_summary = wb.create_sheet("Summary")
        ws_summary.append(['Metric', 'Value'])
        for key, value in summary.items():
            ws_summary.append([key, value])
        
        # Format the summary sheet
        for cell in ws_summary["A"]:
            cell.font = Font(bold=True)
        
        # Set column widths
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = max_length + 2
            ws.column_dimensions[column_letter].width = adjusted_width
            
        for column in ws_summary.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = max_length + 2
            ws_summary.column_dimensions[column_letter].width = adjusted_width
            
        # Save the workbook
        print(f"Saving results to {output_file}...")
        wb.save(output_file)
        print(f"Processing complete. Results saved to {output_file}")
        
        return True
        
    except Exception as e:
        print(f"Error processing the file: {e}")
        return False

if __name__ == "__main__":
    # Check if input file is provided as command line argument
    if len(sys.argv) < 2:
        print("Usage: python excel_processor.py input_excel_file.xlsx [output_excel_file.xlsx]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    # Check if output file is provided as command line argument
    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
    else:
        # Create default output filename based on input filename
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        output_file = f"{base_name}_processed.xlsx"
    
    # Process the file
    result = process_excel_file(input_file, output_file)
    
    if result:
        print("\nSUCCESS: Excel file processed successfully.")
    else:
        print("\nERROR: Failed to process the Excel file.")
        sys.exit(1)