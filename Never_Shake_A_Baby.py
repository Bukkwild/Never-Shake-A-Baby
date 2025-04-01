import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import numpy as np
import os
import glob
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import multiprocessing
import openpyxl
import ttkthemes
from tkinter import scrolledtext
import datetime
import re

def parse_model_string(value):
    if pd.isna(value) or value == '-':
        return ('-', '-', '-')
    
    value = str(value).strip()
    
    # Try to match pattern: YYYY MAKE MODEL
    match = re.match(r'^(\d{4})\s+([A-Za-z-]+)\s+(.+)$', value)
    if match:
        year, make, model = match.groups()
        return (year, make.upper(), model.strip())
    
    # Try to match pattern: MAKE MODEL
    match = re.match(r'^([A-Za-z-]+)\s+(.+)$', value)
    if match:
        make, model = match.groups()
        return ('-', make.upper(), model.strip())
    
    return ('-', '-', value)

class ExcelCombinerApp:
    def __init__(self):
        # Initialize processing settings
        cpu_count = multiprocessing.cpu_count()
        pd.options.mode.chained_assignment = None
        pd.set_option('compute.use_bottleneck', True)
        pd.set_option('compute.use_numexpr', True)
        self.chunk_size = 10000
        
        # Define expected columns in the correct order
        self.expected_columns = [
            'File Name', 'Assesment Year', 'Client Name', 'State', 'Asset Number',
            'VIN?', 'VIN', 'Year', 'Make', 'Model', 'Fuel Type', 'Miles/Hours',
            'Meter Reading', 'Class Code', 'Class Description', 'Department',
            'VEU?', 'VEUs', 'Original Purchase Cost', 'Current Age (months)',
            'Age at Plan Beginning (months)', 'Projected Age at Replacment (months)',
            'Years Past Due', 'Replacement Year',
            'Projected Replacement Asset\'s Life Cycle (months)',
            'Projected Replacement Cost Today', 'Projected Replacement Cost at Replacement',
            'Residual Value of Replacement Asset',
            'Total Depreciation-Based Charge-Back for Replacement',
            'Loan Term', 'Lease Term', '% Used',
            'Residual Value of Current Asset When Replaced', 'Net Capital Cost',
            'Net Straight Line Depreciation/Mo.', 'Imputed LTD Depreciation',
            'Imputed Current Book Value', 'Remaining Life Annual Charge-Back Rate',
            'Number of Replacements in 40 Years'
        ]
        
        # Initialize GUI variables BEFORE creating GUI
        self.root = ttkthemes.ThemedTk(theme="arc")
        self.progress_var = tk.DoubleVar(self.root)
        self.processing = False
        self.selected_directory = None
        
        self.log_enabled = tk.BooleanVar(value=False)  # Default to not creating logs
        self.log_file_path = None
        
        # Create the GUI
        self.create_gui()

    def create_gui(self):
        # Configure root window
        self.root.title("Excel Sheet Combiner")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')

        # Create main frame
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Create control frame
        control_frame = ttk.LabelFrame(self.main_frame, text="Controls", padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 20))

        # Directory selection
        self.dir_frame = ttk.Frame(control_frame)
        self.dir_frame.pack(fill=tk.X, pady=5)
        
        self.dir_label = ttk.Label(
            self.dir_frame,
            text="No directory selected",
            font=('Helvetica', 10)
        )
        self.dir_label.pack(side=tk.LEFT, padx=5)
        
        # Select Directory Button
        self.select_btn = ttk.Button(
            self.dir_frame,
            text="Select Directory",
            command=self.select_directory,  # Changed to select_directory
            style='Accent.TButton'
        )
        self.select_btn.pack(side=tk.RIGHT, padx=5)

        # Add Options Frame
        options_frame = ttk.LabelFrame(control_frame, text="Options", padding="5")
        options_frame.pack(fill=tk.X, pady=5)

        # Add Logging Checkbox
        self.log_checkbox = ttk.Checkbutton(
            options_frame,
            text="Create processing log file",
            variable=self.log_enabled,
            style='Switch.TCheckbutton'
        )
        self.log_checkbox.pack(side=tk.LEFT, padx=5)

        # Add a tooltip/help icon next to the checkbox
        help_label = ttk.Label(
            options_frame,
            text="ⓘ",
            cursor="question_arrow"
        )
        help_label.pack(side=tk.LEFT)
        
        # Bind tooltip
        self.create_tooltip(help_label, 
            "When enabled, creates a detailed log file in the 'logs' folder\n"
            "containing all processing steps and any errors encountered."
        )

        # Button frame for Start/Cancel buttons
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill=tk.X, pady=5)

        # Start button
        self.start_btn = ttk.Button(
            button_frame,
            text="Start Processing",
            command=self.process_directory,
            state='disabled',  # Initially disabled until directory is selected
            style='Accent.TButton'
        )
        self.start_btn.pack(side=tk.RIGHT, padx=5)

        # Cancel button (initially hidden)
        self.cancel_btn = ttk.Button(
            button_frame,
            text="Cancel",
            command=self.cancel_processing,
            style='Secondary.TButton'
        )
        # Don't pack the cancel button initially

        # Progress frame
        progress_frame = ttk.LabelFrame(self.main_frame, text="Progress", padding="10")
        progress_frame.pack(fill=tk.BOTH, expand=True)

        # Progress bar
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))

        # Status text
        self.status_text = tk.Text(
            progress_frame,
            height=12,
            width=70,
            font=('Consolas', 10),
            wrap=tk.WORD,
            bg='#ffffff',
            fg='#333333'
        )
        self.status_text.pack(fill=tk.BOTH, expand=True)

        # Scrollbar for status text
        scrollbar = ttk.Scrollbar(progress_frame, command=self.status_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.configure(yscrollcommand=scrollbar.set)

        # Bottom button frame
        bottom_frame = ttk.Frame(self.main_frame)
        bottom_frame.pack(fill=tk.X, pady=(20, 0))

        # Exit button
        self.exit_btn = ttk.Button(
            bottom_frame,
            text="Exit",
            command=self.root.quit,
            style='Secondary.TButton'
        )
        self.exit_btn.pack(side=tk.RIGHT, padx=5)

    def create_tooltip(self, widget, text):
        """Create a tooltip for a given widget"""
        def show_tooltip(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")

            label = ttk.Label(tooltip, text=text, justify=tk.LEFT,
                            background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                            padding=(5, 5))
            label.pack()

            def hide_tooltip():
                tooltip.destroy()

            tooltip.bind('<Leave>', lambda e: hide_tooltip())
            widget.bind('<Leave>', lambda e: hide_tooltip())

        widget.bind('<Enter>', show_tooltip)

    def write_to_log(self, message, level="INFO"):
        """Write a message to both log file and GUI"""
        if not self.log_enabled.get() or not self.log_file_path:
            # Only show in GUI if logging is disabled
            self.log_status(message)
            return

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {level}: {message}"
        
        try:
            with open(self.log_file_path, 'a', encoding='utf-8') as f:
                f.write(log_message + "\n")
        except Exception as e:
            self.log_status(f"Error writing to log file: {str(e)}")
        
        # Also show in GUI
        self.log_status(message)

    def find_header_row(self, df, sheet_name):
        """Find the likely header row based on common patterns"""
        try:
            # Convert all values to string and lowercase for pattern matching
            df_lower = df.astype(str).apply(lambda x: x.str.lower())
            
            # Key patterns that indicate a header row
            key_patterns = [
                ['asset', 'number'],
                ['vin'],
                ['year'],
                ['make'],
                ['model'],
                ['department'],
                ['cost']
            ]
            
            # Check first 15 rows for header patterns
            for idx in range(min(15, len(df))):
                row_values = df_lower.iloc[idx].astype(str)
                pattern_matches = 0
                
                for pattern in key_patterns:
                    # Check if any cell in this row contains all parts of the pattern
                    for cell in row_values:
                        if all(part in str(cell).lower() for part in pattern):
                            pattern_matches += 1
                            break
                
                # If we find at least 2 pattern matches, this is likely our header row
                if pattern_matches >= 2:
                    self.log_status(f"Found header row at index {idx} in {sheet_name} with {pattern_matches} matching patterns")
                    return idx
            
            self.log_status(f"No header row found in {sheet_name}, will try reading without header specification")
            return None
            
        except Exception as e:
            self.log_status(f"Error finding header row in {sheet_name}: {str(e)}")
            return None

    def find_matching_column(self, df, target_column):
        """Find matching column name accounting for variations"""
        column_variations = {
            'File Name': ['file name', 'filename', 'file'],
            'Assessment Year': ['assessment year', 'assesment year', 'assess year', 'year assessed'],
            'Client Name': ['client name', 'client', 'customer name', 'customer'],
            'State': ['state', 'location state', 'st'],
            'Asset Number': ['asset number', 'asset #', 'asset no', 'asset id', 'equipment number'],
            'VIN': ['vin', 'vin number', 'vehicle id number'],
            'Year': ['year', 'model year', 'vehicle year'],
            'Make': ['make', 'manufacturer', 'vehicle make'],
            'Model': ['model', 'vehicle model'],
            'Fuel Type': ['fuel type', 'fuel', 'fuel source'],
            'Miles/Hours': ['miles/hours', 'miles', 'hours', 'odometer', 'meter reading'],
            'Meter Reading': ['meter reading', 'current reading', 'odometer reading'],
            'Class Code': ['class code', 'asset class', 'class'],
            'Class Description': ['class description', 'class desc', 'asset class description'],
            'Department': ['department', 'dept', 'division'],
            'VEUs': ['veus', 'vehicle equivalent units', 'equivalent units'],
            'Original Purchase Cost': ['original purchase cost', 'purchase cost', 'original cost', 'acquisition cost'],
            'Current Age (months)': ['current age', 'age (months)', 'current age in months'],
            'Replacement Year': ['replacement year', 'replace year', 'year of replacement'],
            'Net Capital Cost': ['net capital cost', 'capital cost', 'net cost'],
            'Age at Plan Beginning (months)': [
                'age at plan beginning (months)',
                'age at beginning of plan',
                'plan beginning age',
                'age at plan start',
                'starting age',
                'initial age months',
                'age at beginning',
                'plan start age'
            ],
            'Projected Age at Replacment (months)': [
                'projected age at replacment (months)',
                'projected replacement age',
                "current asset's projected age at replacement (months)",
                'projected age at replacement',
                'replacement age projection',
                'expected replacement age',
                'age at replacement',
                'projected replacement month age'
            ],
        }

        # Convert target column to lowercase for matching
        target_lower = target_column.lower()
        
        # First try exact match
        if target_column in df.columns:
            return target_column
        
        # Then try lowercase match
        df_columns_lower = [col.lower() for col in df.columns]
        if target_lower in df_columns_lower:
            return df.columns[df_columns_lower.index(target_lower)]
        
        # Try variations with more flexible matching
        if target_column in column_variations:
            for variation in column_variations[target_column]:
                # Exact variation match
                for col in df.columns:
                    if variation == col.lower():
                        return col
                
                # Partial match (if no exact match found)
                for col in df.columns:
                    if variation in col.lower():
                        return col
                    
                # Try matching without parentheses and special characters
                cleaned_variation = ''.join(c.lower() for c in variation if c.isalnum() or c.isspace())
                for col in df.columns:
                    cleaned_col = ''.join(c.lower() for c in col if c.isalnum() or c.isspace())
                    if cleaned_variation in cleaned_col:
                        return col
        
        return None

    def handle_missing_columns(self, df, required_columns):
        """
        Handle missing columns in the DataFrame
        """
        self.log_status("Checking for missing columns...")
        
        # Create a list of missing columns
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            self.log_status(f"Missing columns found: {', '.join(missing_columns)}")
            # Add missing columns with NaN values
            for col in missing_columns:
                df[col] = np.nan
                self.log_status(f"Added missing column: {col}")
        else:
            self.log_status("No missing columns found")
        
        # Ensure all columns are in the correct order
        self.log_status("Reordering columns to match expected format...")
        df = df[required_columns]
        
        return df

    def process_single_file(self, file):
        try:
            file_name = os.path.basename(file)
            self.log_status(f"\n{'='*50}")
            self.log_status(f"Processing file: {file_name}")
            self.log_status(f"{'='*50}")
            
            xl = pd.ExcelFile(file)
            all_sheets = xl.sheet_names
            self.log_status(f"Found sheets in file: {all_sheets}")
            
            target_sheets = ["4.ASSET DATA", "3.FLEET INVENTORY"]
            self.log_status(f"Looking for target sheets: {target_sheets}")
            
            sheet_dfs = {}
            
            for sheet_name in target_sheets:
                self.log_status(f"\nProcessing target sheet: {sheet_name}")
                found_sheet = None
                
                if sheet_name in all_sheets:
                    found_sheet = sheet_name
                    self.log_status(f"Found exact match for sheet: {sheet_name}")
                    
                    try:
                        self.log_status(f"Attempting to read sheet: {found_sheet}")
                        df = pd.read_excel(file, sheet_name=found_sheet, skiprows=5, header=0)
                        self.log_status(f"Successfully read sheet. Initial shape: {df.shape}")
                        
                        # Basic cleanup
                        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
                        df = df.loc[:,~df.columns.duplicated()]
                        self.log_status(f"After cleanup. Final shape: {df.shape}")
                        
                        # Add source sheet info
                        df['Source_Sheet'] = sheet_name
                        sheet_dfs[found_sheet] = df
                        
                    except Exception as e:
                        self.log_status(f"Error reading sheet {found_sheet}: {str(e)}")
                        self.log_status(f"Full error details: {repr(e)}")
            
            self.log_status(f"\nSuccessfully processed sheets: {list(sheet_dfs.keys())}")
            
            if len(sheet_dfs) == 2:
                asset_data = sheet_dfs["4.ASSET DATA"]
                fleet_inventory = sheet_dfs["3.FLEET INVENTORY"]
                
                # Create column mappings
                asset_data_mapping = {
                    'Current Age in Months': 'Current Age (months)',
                    'Age at Beginning of Plan': 'Age at Plan Beginning (months)',
                    "Current Asset's Projected Age At Replacement (months)": 'Projected Age at Replacment (months)',
                    'Current Asset Replacement (Fiscal) Year': 'Replacement Year',
                    'Last Meter Reading (miles or hours)': 'Meter Reading',
                    'Current Asset Class Code': 'Class Code',
                    'Current Asset Class Description': 'Class Description'
                }
                
                fleet_inventory_mapping = {
                    'Asset Model Year': 'Year',
                    'Manufacturer': 'Make',
                    'Most Recent Meter Reading': 'Miles/Hours'
                }
                
                # Rename columns
                asset_data = asset_data.rename(columns=asset_data_mapping)
                fleet_inventory = fleet_inventory.rename(columns=fleet_inventory_mapping)
                
                # Merge the dataframes
                final_df = pd.merge(
                    asset_data,
                    fleet_inventory,
                    on='Asset Number',
                    how='outer',
                    suffixes=('_asset', '_fleet')
                )
                
                # Add missing required columns with default values
                final_df['File Name'] = file_name
                final_df['Assesment Year'] = datetime.datetime.now().year
                final_df['Client Name'] = 'Berkeley'  # Or extract from filename
                final_df['State'] = 'CA'  # Default for Berkeley
                final_df['VIN?'] = final_df['VIN'].notna()
                final_df['Fuel Type'] = ''  # Default empty if not available
                final_df['VEU?'] = final_df['VEUs'].notna()
                
                # Clean up duplicate columns
                for col in final_df.columns:
                    if col.endswith('_asset') and col[:-6] + '_fleet' in final_df.columns:
                        base_col = col[:-6]
                        final_df[base_col] = final_df[col].combine_first(final_df[base_col + '_fleet'])
                        final_df = final_df.drop([col, base_col + '_fleet'], axis=1)
                
                # Ensure all expected columns are present
                for col in self.expected_columns:
                    if col not in final_df.columns:
                        final_df[col] = np.nan
                
                # Reorder columns to match expected order
                final_df = final_df[self.expected_columns]
                
                return final_df
            
            elif len(sheet_dfs) == 1:
                return next(iter(sheet_dfs.values()))
            else:
                raise ValueError(f"No valid sheets found in {file_name}")
            
        except Exception as e:
            self.log_status(f"Error processing {file_name}: {str(e)}")
            return None

    def get_unique_filename(self, base_path):
        """Generate a unique filename by appending a number or timestamp if file exists"""
        directory = os.path.dirname(base_path)
        filename = os.path.basename(base_path)
        name, ext = os.path.splitext(filename)
        
        # First try with numbers
        counter = 1
        new_path = base_path
        while os.path.exists(new_path) and counter < 100:
            new_path = os.path.join(directory, f"{name}_{counter}{ext}")
            counter += 1
        
        # If still exists, use timestamp
        if os.path.exists(new_path):
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            new_path = os.path.join(directory, f"{name}_{timestamp}{ext}")
        
        return new_path

    def process_excel_files(self, directory_path):
        # Look for both .xlsm and .csv files, excluding temporary Excel files
        xlsm_files = [f for f in glob.glob(os.path.join(directory_path, "*.xlsm")) 
                     if not os.path.basename(f).startswith('~$')]
        csv_files = glob.glob(os.path.join(directory_path, "*.csv"))
        all_files = xlsm_files + csv_files
        
        if not all_files:
            raise Exception("No .xlsm or .csv files found in the selected directory")

        total_files = len(all_files)
        self.log_status(f"Found {total_files} files to process")
        self.log_status("Initializing parallel processing...")
        
        # Reset progress bar
        self.progress_var.set(0)
        processed_files = 0

        # Process files in parallel
        with ThreadPoolExecutor(max_workers=multiprocessing.cpu_count()) as executor:
            self.log_status(f"Created thread pool with {multiprocessing.cpu_count()} workers")
            future_to_file = {executor.submit(self.process_single_file, file): file 
                            for file in all_files}
            
            all_data = []
            for future in as_completed(future_to_file):
                if not self.processing:  # Check if processing was canceled
                    self.log_status("Processing canceled by user")
                    executor.shutdown(wait=False)
                    return None
                
                file = future_to_file[future]
                try:
                    self.log_status(f"Starting to process: {os.path.basename(file)}")
                    df = future.result()
                    if df is not None:
                        self.log_status(f"Successfully read file, shape: {df.shape}")
                        all_data.append(df)
                        self.log_status(f"✓ Processed: {os.path.basename(file)}")
                    else:
                        self.log_status(f"✗ Failed: {os.path.basename(file)} (returned None)")
                except Exception as e:
                    self.log_status(f"✗ Error processing {os.path.basename(file)}: {str(e)}")
                    self.log_status(f"Detailed error: {repr(e)}")
                
                processed_files += 1
                progress = (processed_files / total_files) * 100
                self.update_progress(progress)
                self.log_status(f"Progress: {processed_files}/{total_files} files ({progress:.1f}%)")

        if not self.processing:  # Check if processing was canceled
            return None

        if all_data:
            self.log_status(f"\nProcessing complete. Successfully processed {len(all_data)} files")
            self.log_status("Starting data combination process...")
            
            # Ensure all DataFrames have unique column names before concatenation
            cleaned_data = []
            for idx, df in enumerate(all_data):
                self.log_status(f"Cleaning DataFrame {idx + 1}/{len(all_data)}")
                self.log_status(f"Current columns: {', '.join(df.columns)}")
                # Remove any duplicate columns
                df = df.loc[:,~df.columns.duplicated()]
                cleaned_data.append(df)
            
            # Combine all data efficiently
            self.log_status("Concatenating all DataFrames...")
            combined_data = pd.concat(cleaned_data, ignore_index=True)
            self.log_status(f"Combined data shape: {combined_data.shape}")
            
            # Create final DataFrame efficiently
            self.log_status("Creating final DataFrame...")
            final_df = combined_data.copy()
            
            # Add missing columns efficiently
            missing_cols = set(self.expected_columns) - set(final_df.columns)
            if missing_cols:
                self.log_status(f"Adding missing columns: {', '.join(missing_cols)}")
            for col in missing_cols:
                final_df[col] = np.nan
            
            # Reorder columns
            self.log_status("Reordering columns...")
            final_df = final_df[self.expected_columns]
            
            # Get unique output filename
            base_output_file = os.path.join(directory_path, "Combined_Data.xlsx")
            output_file = self.get_unique_filename(base_output_file)
            
            self.log_status(f"Writing to Excel file: {os.path.basename(output_file)}...")
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
                final_df.to_excel(writer, sheet_name='Combined_Data', index=False)
                self.format_excel(writer)
            
            self.log_status(f"File saved successfully: {output_file}")
            return output_file
        else:
            raise Exception("No data was processed successfully")

    def format_excel(self, writer):
        worksheet = writer.sheets['Combined_Data']
        
        # Batch format operations
        for idx, col in enumerate(self.expected_columns, 1):
            col_letter = openpyxl.utils.get_column_letter(idx)
            worksheet.column_dimensions[col_letter].width = max(len(col) + 2, 15)
        
        # Format headers in one go
        for cell in worksheet[1]:
            cell.font = openpyxl.styles.Font(bold=True)
            cell.alignment = openpyxl.styles.Alignment(horizontal='center')
        
        # Format number columns efficiently
        number_format = '$#,##0.00'
        number_columns = ['Original Purchase Cost', 'Projected Replacement Cost Today',
                         'Projected Replacement Cost at Replacement', 'Net Capital Cost']
        
        for col in number_columns:
            if col in self.expected_columns:
                col_letter = openpyxl.utils.get_column_letter(
                    self.expected_columns.index(col) + 1)
                for cell in worksheet[col_letter][1:]:
                    cell.number_format = number_format

    def log_status(self, message):
        # Use after() to safely update GUI from a non-main thread
        self.root.after(0, self._log_status_safe, message)

    def _log_status_safe(self, message):
        if message.startswith("✓"):
            tag = "success"
            self.status_text.tag_configure("success", foreground="green")
        elif message.startswith("✗"):
            tag = "error"
            self.status_text.tag_configure("error", foreground="red")
        else:
            tag = "info"
            self.status_text.tag_configure("info", foreground="black")
            
        self.status_text.insert(tk.END, f"{message}\n", tag)
        self.status_text.see(tk.END)

    def update_progress(self, value):
        # Use after() to safely update GUI from a non-main thread
        self.root.after(0, self._update_progress_safe, value)

    def _update_progress_safe(self, value):
        self.progress_var.set(value)

    def select_directory(self):
        """Handle directory selection"""
        directory = filedialog.askdirectory(title="Select Directory containing Excel files")
        if directory:
            self.selected_directory = directory
            self.dir_label.config(text=f"Selected: {os.path.basename(directory)}")
            self.start_btn.config(state='normal')  # Enable start button
            self.log_status(f"Selected directory: {directory}")

    def process_directory(self):
        if not self.selected_directory:
            self.log_status("No directory selected. Please select a directory first.")
            return
        
        # Initialize log file if enabled
        if self.log_enabled.get():
            self.initialize_log_file()
            self.write_to_log(f"Starting processing for directory: {self.selected_directory}")
        
        # Set processing flag and update UI
        self.processing = True
        self.select_btn.config(state='disabled')
        self.start_btn.config(state='disabled')
        self.cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        # Start processing in a separate thread
        processing_thread = threading.Thread(
            target=self._process_directory_thread,
            args=(self.selected_directory,)
        )
        processing_thread.daemon = True
        processing_thread.start()

    def cancel_processing(self):
        self.processing = False
        self.log_status("Canceling processing...")

    def initialize_log_file(self):
        """Create a new log file with timestamp in name"""
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_dir = "logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        self.log_file_path = os.path.join(log_dir, f"processing_log_{timestamp}.txt")
        
        # Write initial log entry
        with open(self.log_file_path, 'w') as f:
            f.write(f"Processing Log - Started at {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("-" * 80 + "\n\n")

    def _process_directory_thread(self, directory):
        try:
            self.write_to_log("Processing files...")
            output_file = self.process_excel_files(directory)
            if output_file:
                self.write_to_log(f"Success! Output saved to: {output_file}", "SUCCESS")
            
            # Update title to show completion
            self.root.after(0, lambda: self.root.title(f"Never Shake A Baby - {os.path.basename(directory)} (Completed)"))
        except Exception as e:
            self.write_to_log(f"Error: {str(e)}", "ERROR")
            self.root.after(0, lambda: self.root.title(f"Never Shake A Baby - {os.path.basename(directory)} (Error)"))
        finally:
            if self.log_enabled.get():
                self.write_to_log("Processing completed")
            # Reset UI
            self.processing = False
            self.root.after(0, lambda: self.cancel_btn.pack_forget())
            self.root.after(0, lambda: self.select_btn.config(state='normal'))
            self.root.after(0, lambda: self.start_btn.config(state='normal'))
            self.selected_directory = None
            self.log_file_path = None  # Reset log file path

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelCombinerApp()
    app.run()








