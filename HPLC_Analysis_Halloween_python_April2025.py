import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os


class ExcelFileOpener:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Opener")
        self.selected_dataframes = {}

        # Create a button to open the file dialog
        self.open_button = tk.Button(root, text="Open Excel File", command=self.open_file)
        self.open_button.pack(pady=20)

    def open_file(self):
        # Open a file dialog to select an Excel file
        file_path = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=(("Excel files", "*.xls;*.xlsx"), ("All files", "*.*"))
        )
        if file_path:
            try:
                # Read all sheets from the Excel file into a dictionary of DataFrames
                all_sheets = pd.read_excel(file_path, sheet_name=None)

                # Iterate through all sheets
                for sheet_name, data in all_sheets.items():
                    selected_rows = data.iloc[13:24]
                    cleaned_rows = selected_rows.dropna(how='all')

                    # Ensure cleaned_rows has at least 2 rows
                    if cleaned_rows.shape[0] > 1:
                        file_name = str(cleaned_rows.iloc[1, 0]).strip()  # Get name from second row
                    else:
                        print(f"Warning: Sheet '{sheet_name}' does not have enough rows. Saving as 'corrupted'.")
                        valid_file_name = "corrupted_" + sheet_name.replace(" ", "_")
                        self.selected_dataframes[valid_file_name] = pd.DataFrame()  # Empty DataFrame for corrupted sheets
                        continue

                    valid_file_name = file_name.replace(" ", "_")

                    if not valid_file_name:
                        valid_file_name = "unknown_title"

                    # Drop the first seven columns and clean the DataFrame
                    remaining_cols = cleaned_rows.drop(columns=cleaned_rows.columns[:7], errors='ignore')
                    if not remaining_cols.empty:
                        remaining_cols.iloc[:, 0] = remaining_cols.iloc[:, 0].shift(-2)
                        remaining_cols = remaining_cols.iloc[:, 1:]

                    # Replace the first row with headers
                    remaining_cleaned = remaining_cols.apply(lambda row: row.dropna().reset_index(drop=True), axis=1)
                    remaining_cleaned.columns = remaining_cleaned.iloc[0]  # Set the first row as header
                    remaining_cleaned = remaining_cleaned[1:]  # Remove the header row from the data

                    self.selected_dataframes[valid_file_name] = remaining_cleaned.reset_index(drop=True)

                # Create a summary DataFrame
                summary_df = self.create_summary()

                # Calculate initial glycerol
                initial_glycerol = self.calculate_initial_glycerol(summary_df)
                print(f"Initial Glycerol: {initial_glycerol}")

                # Create Conversion DataFrame
                conversion_df = self.create_conversion_df(summary_df, initial_glycerol)

                # Create Yield DataFrame
                yield_df = self.create_yield_df(conversion_df, summary_df)

                # Create Selectivity DataFrame
                selectivity_df = self.create_selectivity_df(summary_df)

                # Save all DataFrames and the summary to a new Excel file
                output_file_path = os.path.splitext(file_path)[0] + "_output.xlsx"
                with pd.ExcelWriter(output_file_path) as writer:
                    summary_df.to_excel(writer, sheet_name="Summary", index=False)
                    conversion_df.to_excel(writer, sheet_name="Conversion", index=False)
                    yield_df.to_excel(writer, sheet_name="Yield", index=False)
                    selectivity_df.to_excel(writer, sheet_name="Selectivity", index=False)

                    # Save each DataFrame as a separate sheet without index and with headers
                    for df_name, df in self.selected_dataframes.items():
                        df.reset_index(drop=True, inplace=True)  # Reset index to drop original index
                        df.to_excel(writer, sheet_name=df_name[:31], index=False)

                print(f"\nDataFrames, summary, conversion, yield, and selectivity saved to: {output_file_path}")

                # Close the GUI after opening the file
                self.root.destroy()

            except Exception as e:
                print(f"Error opening file: {e}")

    def create_summary(self):
        summary_data = []
        unique_samples = set()

        for df_name, df in self.selected_dataframes.items():
            if df.empty:
                summary_data.append((df_name, {'GLY': 0}))  # Placeholder for corrupted sheets
                continue

            samples = df.iloc[:, 0].dropna().unique()
            unique_samples.update(samples)

            sample_values = {sample: 0 for sample in samples}

            for sample in samples:
                value_row = df[df.iloc[:, 0] == sample]
                if not value_row.empty:
                    sample_values[sample] = float(value_row.iloc[0, 3])  # Convert to float

            summary_data.append((df_name, sample_values))

        unique_samples_list = sorted(unique_samples)
        summary_df = pd.DataFrame(index=unique_samples_list)

        for df_name, sample_values in summary_data:
            for sample in unique_samples_list:
                summary_df.loc[sample, df_name] = float(sample_values.get(sample, 0))

        summary_df.fillna(0, inplace=True)
        summary_df.reset_index(inplace=True)
        summary_df.rename(columns={'index': 'Sample Name'}, inplace=True)

        return summary_df

    def calculate_initial_glycerol(self, summary_df):
        gly_row = summary_df[summary_df['Sample Name'] == 'GLY']
        if not gly_row.empty:
            initial_values = gly_row.iloc[0, 1:4].astype(float)
            initial_glycerol = initial_values.max()
            return initial_glycerol
        return 0

    def create_conversion_df(self, summary_df, initial_glycerol):
        conversion_data = []
        measured_gly = summary_df[summary_df['Sample Name'] == 'GLY'].iloc[0, 1:].values

        for i, df_name in enumerate(summary_df.columns[1:]):
            measured_value = measured_gly[i]
            consumed_gly = initial_glycerol - measured_value
            consumed_gly = consumed_gly if consumed_gly <= 0.4 else 0

            conversion_data.append({
                'Sample Name': df_name,
                'Initial GLY': initial_glycerol,
                'Measured GLY': measured_value,
                'Consumed GLY': consumed_gly,
                'Conversion (%)': round((consumed_gly / initial_glycerol * 100), 3) if initial_glycerol > 0 else 0
            })

        conversion_df = pd.DataFrame(conversion_data)
        return conversion_df

    def create_yield_df(self, conversion_df, summary_df):
        yield_data = []
        
        for _, row in conversion_df.iterrows():
            sample_name = row['Sample Name']
            initial_gly = row['Initial GLY']
            measured_gly = row['Measured GLY']
            consumed_gly = row['Consumed GLY']

            yield_row = {
                'Sample Name': sample_name,
                'Initial GLY': initial_gly,
                'Measured GLY': measured_gly,
                'Consumed GLY': consumed_gly
            }

            for sample in summary_df['Sample Name'][:]:
                value = summary_df.loc[summary_df['Sample Name'] == sample, sample_name].values[0]
                yield_row[sample] = 0 if pd.isna(value) else value
                
                if consumed_gly > 0:
                    yield_row[f"{sample} Yield %"] = round((yield_row[sample] / consumed_gly * 100), 3) if sample != 'GLY' else None
                
                    if sample == 'Glycolic Acid':
                        yield_row[f"{sample} Yield %"] /= 1.5
                    elif sample == 'FA':
                        yield_row[f"{sample} Yield %"] /= 3
                    elif sample == 'OA':
                        yield_row[f"{sample} Yield %"] /= 1.5
                else:
                    yield_row[f"{sample} Yield %"] = None
            
            yield_data.append(yield_row)

        yield_df = pd.DataFrame(yield_data)

        # Calculate Faradaic Efficiency after Yield is computed
        yield_df = self.calculate_faradaic_efficiency(yield_df)  # Add Faradaic Efficiency

        column_order = ['Sample Name', 'Initial GLY', 'Measured GLY', 'Consumed GLY'] + \
                       [col for col in summary_df['Sample Name'][1:]] + \
                       [f"{col} Yield %" for col in summary_df['Sample Name'][1:]] + \
                       ['Faradaic Efficiency (%)']  # Add Faradaic Efficiency to column order
        yield_df = yield_df[column_order]

        yield_df.fillna(0, inplace=True)

        return yield_df

    def calculate_faradaic_efficiency(self, yield_df, theoretical_product_col='Initial GLY'):
        """Calculate Faradaic Efficiency for each product."""
        faradaic_efficiency_data = []

        for _, row in yield_df.iterrows():
            sample_name = row['Sample Name']
            actual_product = row[sample_name]  # Actual product amount from the yield calculations
            theoretical_product = row[theoretical_product_col]  # Theoretical product (e.g., Initial GLY)

            # Calculate Faradaic Efficiency: (Actual Product / Theoretical Product) * 100
            if theoretical_product > 0:
                faradaic_efficiency = (actual_product / theoretical_product) * 100
            else:
                faradaic_efficiency = 0  # To avoid division by zero, set FE to 0 if theoretical product is 0.

            # Add Faradaic Efficiency to the row
            faradaic_efficiency_data.append({
                'Sample Name': sample_name,
                'Faradaic Efficiency (%)': round(faradaic_efficiency, 3)
            })

        # Convert the Faradaic Efficiency data into a DataFrame
        faradaic_efficiency_df = pd.DataFrame(faradaic_efficiency_data)
        
        # Merge the Faradaic Efficiency data with the existing yield DataFrame
        yield_df = pd.merge(yield_df, faradaic_efficiency_df, on='Sample Name', how='left')

        return yield_df

    def create_selectivity_df(self, summary_df):
        selectivity_data = []
        for _, row in summary_df.iterrows():
            selectivity_data.append({
                'Sample Name': row['Sample Name'],
                'Selectivity': row['Selectivity'] if 'Selectivity' in row else 0  # Placeholder logic
            })

        selectivity_df = pd.DataFrame(selectivity_data)
        return selectivity_df

# Initialize and run the Tkinter window
root = tk.Tk()
app = ExcelFileOpener(root)
root.mainloop()
