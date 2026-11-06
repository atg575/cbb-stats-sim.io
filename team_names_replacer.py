import pandas as pd

file_path = "/Users/alexvazquez/Coding/cbb_stats/team_names.csv"
from team_name_extractor import extract_teams

try:
    df = pd.read_csv(file_path)
    
 # --- FIX: Convert column data to lowercase immediately after reading ---
    # Use the .str.lower() method provided by pandas to lowercase all strings in the column
    cbbref_data_list = df["School"].str.lower().tolist()
    kenpom_data_list = df["Team"].str.lower().tolist()
    warren_teams = extract_teams(url = "https://www.warrennolan.com/basketball/2026/teams-az")
    # Convert lists to sets for efficient 'in' checking
    cbbref_schools = set(cbbref_data_list)
    kenpom_teams = set(kenpom_data_list)

    print(f"CBBRef Schools: {cbbref_schools}")
    print(f"KenPom Teams: {kenpom_teams}")
    print(f"Warren-Nolan Teams: {warren_teams}")
    
    print("\nSchools in CBBRef data but not in KenPom data:")
    # Iterate through the cbbref schools set
    for school_name in cbbref_schools:
        if school_name not in kenpom_teams:
            print(school_name)
        if school_name not in warren_teams:
            print(school_name)
    print("-"*20)
    for school_name in kenpom_teams:
        if school_name not in warren_teams:
            print(school_name)


except FileNotFoundError:
    print("Error: The specified CSV file was not found.")
except KeyError as e:
    # Print which specific column caused the error
    print(f"Error: The specified column name {e} was not found in the CSV file.")
except Exception as e:
    # Catch any other potential errors during file reading or processing
    print(f"An unexpected error occurred: {e}")