import os
import datetime
import shutil
import re
import pandas as pd
import logging
from typing import Optional

class DataProcessor:
    """
    A class to process Excel data by reading from a source file, transforming the data,
    and updating or creating a target Excel file with the processed data.
    """

    def __init__(self):
        """
        Initializes the DataProcessor with header mappings, modification rules,
        and a dictionary of technical abbreviations for replacement.
        """
        # Mapping from target column names to source column names or None if special handling is required
        self.header_mapping = {
            "Id": "Id",
            "Last update": "Completion time",
            "Name": None,  # Not directly mapped
            "Contact email": "Contact email",
            "Type": "Type",
            "Stage": "Stage",
            "label": None,  # Special handling
            "Business / Organisation Description": None,  # Special handling
            "Website": None,  # Special handling
            "Tags": None,  # Special handling
            "Image": "Organisation / business logo url",
            "Are you a member of the Artificial Intelligence Researchers Association?": None,
            "Full Name": None,
            "Which of the following apply to you?": "Which of the following apply to you?",
            "Organisation": None,
            "Professional website/ Social URL": None,
            "AI focus Area": "AI focus Area ",
            "Region": "Region",
            "Full Address": "Full Address",
            "Market presence": "Market presence",
            "AI Technologies": "Which AI technologies is your organisation or you as an individual currently using or developing? (Select all that apply)",
            "AI Enablement Capabilities": "What AI enablement capabilities does your organisation or you as an individual have? (Select all that apply)",
            "Business Areas Or Research Fields": "Which business areas or research fields does your organisation or you as an individual focus on? (Select all that apply)",
            "Industry Sector(S)": "Which industry sector(s) does your organisation or you as an individual focus on? (Select all that apply)",
            "Upload a profile photo for yourself or your business": None,
            "Additional information": "Would you like to include additional information? "
        }

        # Rules to modify content of specific columns
        self.rules = {
            "Which of the following apply to you?": self.replace_semicolon_with_pipe,
            "AI Technologies": self.replace_semicolon_with_pipe_and_title,
            "AI Enablement Capabilities": self.replace_semicolon_with_pipe_and_title,
            "Business Areas Or Research Fields": self.replace_semicolon_with_pipe_and_title,
            "Industry Sector(S)": self.replace_semicolon_with_pipe_and_title
        }

        # Dictionary of technical abbreviations to be replaced with their capitalized forms
        self.tech_abbreviations = {
            'ai': 'AI',
            'nlp': 'NLP',
            'ml': 'ML',
            'ar': 'AR',
            'vr': 'VR',
            'iot': 'IoT',
            'api': 'API',
            'saas': 'SaaS',
            'paas': 'PaaS',
            'iaas': 'IaaS',
            'gpu': 'GPU',
            'cpu': 'CPU',
            'ui': 'UI',
            'ux': 'UX',
            'crm': 'CRM',
            'erp': 'ERP',
            'poc': 'PoC',
            'mvp': 'MVP',
            'cv': 'CV',
            'it': 'IT',
        }

        # Configure logging
        logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    def read_excel(self, file_path: str) -> Optional[pd.DataFrame]:
        """
        Reads an Excel file and returns a pandas DataFrame.

        Parameters:
            file_path (str): The path to the Excel file.

        Returns:
            pd.DataFrame: The DataFrame containing the Excel data, or None if an error occurs.
        """
        try:
            df = pd.read_excel(file_path)
            logging.info(f"Successfully read Excel file: {file_path}")
            return df
        except FileNotFoundError:
            logging.error(f"File not found: {file_path}")
            return None
        except Exception as e:
            logging.error(f"Error reading file: {file_path}, Error: {str(e)}")
            return None

    def process_data(self, src_df: pd.DataFrame) -> pd.DataFrame:
        """
        Processes the source DataFrame by mapping headers and modifying content.

        Parameters:
            src_df (pd.DataFrame): The source DataFrame.

        Returns:
            pd.DataFrame: The processed DataFrame ready for export.
        """
        # Initialize the target DataFrame with the target columns
        target_columns = list(self.header_mapping.keys())
        target_df = pd.DataFrame(columns=target_columns)

        # Map and fill headers from source to target DataFrame
        target_df = self.match_and_fill_headers(src_df, target_df)

        # Modify the content according to the rules
        target_df = self.modify_content(target_df)

        return target_df

    def match_and_fill_headers(self, src_df: pd.DataFrame, target_df: pd.DataFrame) -> pd.DataFrame:
        """
        Maps source DataFrame columns to target DataFrame columns and fills data.

        Parameters:
            src_df (pd.DataFrame): The source DataFrame.
            target_df (pd.DataFrame): The target DataFrame to be filled.

        Returns:
            pd.DataFrame: The target DataFrame with mapped and filled data.
        """
        # Filter out specific rows based on conditions
        filtered_src_df = src_df[
            ~(
                (src_df["Type"] == "Individual (e.g. researcher, academic or engineers)") &
                (src_df["Are you a member of the Artificial Intelligence Researchers Association?"] == "Yes")
            )
        ]

        for target_col, src_col in self.header_mapping.items():
            if target_col == "Tags":
                # Special handling for 'Tags' column
                target_df["Tags"] = filtered_src_df.apply(
                    lambda row: "AI Forum NZ" if (
                        row["Type"] == "Business / Organisation" and
                        row["Is your organisation a member of AI Forum NZ?"] == "Yes"
                    ) else None,
                    axis=1
                )
            elif target_col == "label":
                # Special handling for 'label' column
                target_df["label"] = filtered_src_df.apply(
                    lambda row: row["Full Name"] if row["Type"] == "Individual (e.g. researcher, academic or engineers)"
                    else row["Business / Organisation Name"],
                    axis=1
                )
            elif target_col == "Website":
                # Special handling for 'Website' column
                target_df["Website"] = filtered_src_df.apply(
                    lambda row: row["Professional website/ Social URL"] if row["Type"] == "Individual (e.g. researcher, academic or engineers)"
                    else row["Website"],
                    axis=1
                )
            elif target_col == "Business / Organisation Description":
                # Special handling for 'Business / Organisation Description' column
                target_df["Business / Organisation Description"] = filtered_src_df.apply(
                    lambda row: row["Organisation"] if row["Type"] == "Individual (e.g. researcher, academic or engineers)"
                    else row["Business / Organisation Description"],
                    axis=1
                )
            elif src_col in filtered_src_df.columns:
                # Direct mapping of columns
                target_df[target_col] = filtered_src_df[src_col]
            else:
                # Assign None if source column not found
                target_df[target_col] = None

        return target_df

    def modify_content(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Modifies the content of the DataFrame based on predefined rules and fixes abbreviations.

        Parameters:
            df (pd.DataFrame): The DataFrame to be modified.

        Returns:
            pd.DataFrame: The modified DataFrame.
        """
        # Apply modification rules to the specified columns
        for column, rule in self.rules.items():
            if column in df.columns:
                df[column] = df[column].apply(rule)
                # Fix abbreviations in the modified columns
                df[column] = df[column].apply(self.fix_abbreviations)
        return df

    def fix_abbreviations(self, text: str) -> str:
        """
        Replaces abbreviations in the text with their capitalized forms.

        Parameters:
            text (str): The text to be processed.

        Returns:
            str: The text with abbreviations fixed.
        """
        if not isinstance(text, str):
            return text

        # Iterate over each abbreviation and replace it in the text
        for abbr, replacement in self.tech_abbreviations.items():
            # Regular expression pattern to match the abbreviation as a whole word
            pattern = r'(?<![a-zA-Z])' + re.escape(abbr) + r'(?![a-zA-Z])'
            # Replace the abbreviation with the capitalized form
            text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
        return text

    @staticmethod
    def replace_semicolon_with_pipe(text: str) -> str:
        """
        Replaces semicolons with pipes in the text.

        Parameters:
            text (str): The text to be modified.

        Returns:
            str: The modified text.
        """
        if isinstance(text, str):
            return text.replace(";", "|")
        return text

    @staticmethod
    def replace_semicolon_with_pipe_and_title(text: str) -> str:
        """
        Replaces semicolons with pipes and converts text to title case.

        Parameters:
            text (str): The text to be modified.

        Returns:
            str: The modified text in title case.
        """
        if isinstance(text, str):
            return text.replace(";", "|").title()
        return text

    def update_excel(self, src_file_path: str, target_file_path: str) -> bool:
        """
        Updates the target Excel file with new data from the source file.

        Parameters:
            src_file_path (str): Path to the source Excel file.
            target_file_path (str): Path to the target Excel file.

        Returns:
            bool: True if the update is successful, False otherwise.
        """
        src_df = self.read_excel(src_file_path)
        if src_df is None:
            return False

        new_data = self.process_data(src_df)

        if os.path.exists(target_file_path):
            try:
                # Create a backup of the existing target file
                timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                backup_path = f"backup_{timestamp}_{os.path.basename(target_file_path)}"
                shutil.copy2(target_file_path, backup_path)
                logging.info(f"Backup created: {backup_path}")

                existing_data = self.read_excel(target_file_path)
                if existing_data is not None:
                    updated_count = 0
                    final_data = existing_data.copy()

                    # Update existing records
                    for index, new_row in new_data.iterrows():
                        mask = existing_data['Id'] == new_row['Id']
                        if mask.any():
                            old_row = existing_data.loc[mask].iloc[0]
                            columns_to_compare = [col for col in new_row.index if col != 'Id']
                            has_changes = False
                            changes_log = []

                            logging.info(f"Checking record with ID {new_row['Id']}, label: {new_row['label']}")

                            for col in columns_to_compare:
                                if pd.notna(new_row[col]) and new_row[col] != old_row[col]:
                                    has_changes = True
                                    final_data.loc[mask, col] = new_row[col]
                                    changes_log.append(
                                        f"Field: {col}\n"
                                        f"  - Old value: {old_row[col]}\n"
                                        f"  - New value: {new_row[col]}"
                                    )

                            if has_changes:
                                updated_count += 1
                                logging.info(f"Updates found for ID {new_row['Id']}:")
                                for change in changes_log:
                                    logging.info(change)
                            else:
                                logging.info(f"No updates needed for ID {new_row['Id']}")
                        else:
                            pass  # ID not found in existing data

                    # Add new records
                    existing_ids = set(existing_data['Id'].values)
                    new_records = new_data[~new_data['Id'].isin(existing_ids)]

                    if not new_records.empty:
                        logging.info(f"New records to add: {len(new_records)}")
                        for _, row in new_records.iterrows():
                            logging.info(f"Adding new record - ID: {row['Id']}, Label: {row['label']}")
                        final_data = pd.concat([final_data, new_records], ignore_index=True)

                    if updated_count > 0 or not new_records.empty:
                        # Save the updated data to the target file
                        final_data.to_excel(target_file_path, index=False)
                        logging.info(f"Update summary:\n - Updated {updated_count} records\n - Added {len(new_records)} new records")
                    else:
                        logging.info("Data is up to date. No changes made.")
                else:
                    # If existing data could not be read, write new data to the target file
                    new_data.to_excel(target_file_path, index=False)
                    logging.info(f"Created new file: {target_file_path}")
            except Exception as e:
                logging.error(f"Error updating file: {target_file_path}, Error: {str(e)}")
                return False
        else:
            # If target file does not exist, create it with the new data
            new_data.to_excel(target_file_path, index=False)
            logging.info(f"Created new file: {target_file_path}")

        return True

def main():
    """
    Main function to execute the data processing and updating of the Excel file.
    """
    processor = DataProcessor()
    src_file_path = "AI Ecosystem & Capability Map for Aotearoa.xlsx"
    target_file_path = "target2.xlsx"
    success = processor.update_excel(src_file_path, target_file_path)
    if success:
        logging.info("Data processing and update completed successfully.")
    else:
        logging.error("Data processing and update failed.")

if __name__ == "__main__":
    main()
