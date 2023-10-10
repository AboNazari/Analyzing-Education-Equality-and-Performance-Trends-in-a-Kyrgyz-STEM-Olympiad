import pandas as pd

def extract_distinct_schools(input_file, output_file):
    # Read the excel file
    df = pd.read_excel(input_file)

    # Select distinct rows by 'school' and 'region'
    distinct_df = df[['school', 'region']].drop_duplicates()

    # Write the new DataFrame to an Excel file
    distinct_df.to_excel(output_file, index=False)

if __name__ == "__main__":
    input_file = 'allClasses.xlsx'  # replace with your input file name
    output_file = 'extractedSchools.xlsx'  # replace with your output file name

    extract_distinct_schools(input_file, output_file)


