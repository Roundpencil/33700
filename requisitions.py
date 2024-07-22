import os
import pandas as pd


def clean_phone_number(phone_number):
    """
    Clean phone number according to the specified rules:
    - If the number starts with + and has 9 or 14 digits, replace + with 0.
    - If the number starts with + and has a different number of digits, replace + with 00.
    """
    phone_str = str(phone_number)
    if phone_str.startswith('+33'):
        digits_only = phone_str[3:]
        if len(digits_only) == 9 or len(digits_only) == 14:
            return '0' + digits_only
        else:
            return '00' + digits_only
    return phone_str


def filter_excel_files(input_folder, phone_numbers, urls, output_file):
    # Clean the phone numbers before processing
    cleaned_phone_numbers = [clean_phone_number(phone) for phone in phone_numbers]

    # Initialize an empty dataframe to hold the results
    result_df = pd.DataFrame()

    # Iterate through all files in the input folder
    for file_name in os.listdir(input_folder):
        try:
            if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
                file_path = os.path.join(input_folder, file_name)
                print(f"current file : {file_path}")

                # Read the Excel file
                df = pd.read_excel(file_path)

                # Filter rows where "expediteur_nettoye" is in the cleaned phone numbers
                phone_filter = df['expediteur_nettoye'].astype(str).isin(cleaned_phone_numbers)

                # Filter rows where "URL_REBOND_SIGNALE" contains any of the provided URLs
                url_filter = df['URL_REBOND_SIGNALE'].astype(str).apply(lambda x: any(url in x for url in urls))

                # Combine both filters
                combined_filter = phone_filter | url_filter

                # Append the filtered data to the result dataframe
                result_df = pd.concat([result_df, df[combined_filter]])
        except Exception as e:
            print(f"Exception : {e}")
            continue

    # Write the result dataframe to an output Excel file
    result_df.to_excel(output_file, index=False)
    print(f"Filtered data has been saved to {output_file}")


# Example usage
# input_folder = '/path/to/your/input/folder'
# phone_numbers = ['+754220004', '+another_phone_number']
# urls = ['part_of_url1', 'part_of_url2']
# output_file = '/path/to/your/output/file.xlsx'

input_folder = r'C:\Users\Pierre TROCME\OneDrive - AFMM\data 33700\data retraitées pour V2 rapports'

phone_numbers = [
    '+33744749865',
    '+33757812168',
    '+33628422338',
    '+33638038037',
    '+33627031849',
    '+33744749865',
    '+33744896447',
    '+33749255964'
]

### valeurs de test
# input_folder = r'C:\Users\Pierre TROCME\OneDrive - AFMM\data 33700\test requisition'
# phone_numbers = [
#     '+33637702211'
# ]
urls = [
    'iledefr.com',
    'navigo-agence.com',
    'ship-swiss.info',
    'fr-disneyplus.com',
    'amendes.gouv-paiement.info',
    'dhl.com-suivre.info',
    'connexion-navigo.com',
    'agences-navigo.com',
    'ligne-prixtel.com'
]
output_file = r'C:\Users\Pierre TROCME\OneDrive - AFMM\data 33700\00_output_requisition.xlsx'

filter_excel_files(input_folder, phone_numbers, urls, output_file)


# import os
# import pandas as pd
#
#
# def filter_excel_files(input_folder, phone_numbers, urls, output_file):
#     # Initialize an empty dataframe to hold the results
#     result_df = pd.DataFrame()
#
#     # Iterate through all files in the input folder
#     for file_name in os.listdir(input_folder):
#         if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
#             file_path = os.path.join(input_folder, file_name)
#
#             # Read the Excel file
#             df = pd.read_excel(file_path)
#
#             # Filter rows where "expediteur_nettoye" is in the provided phone numbers
#             phone_filter = df['expediteur_nettoye'].astype(str).isin(phone_numbers)
#
#             # Filter rows where "URL_REBOND_SIGNALE" contains any of the provided URLs
#             url_filter = df['URL_REBOND_SIGNALE'].astype(str).apply(lambda x: any(url in x for url in urls))
#
#             # Combine both filters
#             combined_filter = phone_filter | url_filter
#
#             # Append the filtered data to the result dataframe
#             result_df = pd.concat([result_df, df[combined_filter]])
#
#     # Write the result dataframe to an output Excel file
#     result_df.to_excel(output_file, index=False)
#     print(f"Filtered data has been saved to {output_file}")
#
#
# # Example usage
# input_folder = '/path/to/your/input/folder'
# phone_numbers = ['754220004', 'another_phone_number']
# urls = ['part_of_url1', 'part_of_url2']
# output_file = '/path/to/your/output/file.xlsx'
#
# filter_excel_files(input_folder, phone_numbers, urls, output_file)
