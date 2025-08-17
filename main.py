from ExtractTable import ExtractTable
import pandas as pd
from datetime import datetime

# Replace with your valid API key
# Replace YOUR_API_KEY with your actual API key
API_KEY = "OcMPHc25bMlYCNfMc4mWV7FyZS8vViPqXT7z32HS"
IMAGE_PATH = "./IMG_5786.JPG"  # Path to the input image
# Path to save the output Excel file
OUTPUT_EXCEL_PATH = "./driver_duty_roster.xlsx"


def convert_image_to_excel(api_key, image_path, output_excel_path):
    try:
        # Initialize ExtractTable session
        et_sess = ExtractTable(api_key=api_key)

        # Check API key validity and usage
        usage_info = et_sess.check_usage()
        print("API Key Usage Info:", usage_info)

        # Process the image file to extract table data
        table_data = et_sess.process_file(
            filepath=image_path, output_format="df")

        if table_data:
            # Get current date information
            current_date = datetime.now()
            current_month = current_date.month
            current_year = current_date.year

            # Process cell values in each DataFrame
            processed_tables = []
            for df in table_data:
                df = df.applymap(lambda x: x.upper()
                                 if isinstance(x, str) else x)
                df = df.replace({
                    "AL": "A1", "AI": "A1",
                    "BL": "B1", "BI": "B1",
                    "CL": "C1", "CI": "C1",
                    "OFFL": "OFFL1", "OFFI": "OFFL1"
                })

                # Handle date logic
                def handle_date(cell):
                    if isinstance(cell, int) and 1 <= cell <= 31:
                        month = current_month + 1 if cell < current_date.day else current_month
                        if month > 12:
                            month = 1
                            year = current_year + 1
                        else:
                            year = current_year
                        return f"{cell} tháng {month} năm {year}"
                    return cell

                df = df.applymap(handle_date)

                # Filter rows where column B contains "Bean"
                filtered_df = df[df.iloc[:, 1] == "Bean"]
                processed_tables.append(filtered_df)

            # Combine all processed tables into a single DataFrame
            combined_df = pd.concat(processed_tables, ignore_index=True)

            # Save the DataFrame to an Excel file
            combined_df.to_excel(output_excel_path, index=False)
            print(f"Excel file saved at: {output_excel_path}")
        else:
            print("No table data found in the image.")
    except Exception as e:
        print("Error processing image:", str(e))


# Example usage
convert_image_to_excel(API_KEY, IMAGE_PATH, OUTPUT_EXCEL_PATH)
