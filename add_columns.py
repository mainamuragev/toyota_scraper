import pandas as pd
import re

FILE_NAME = "Toyota_parts_WITH_MODELS.xlsx"

def parse_models(model_string):
    if pd.isna(model_string) or str(model_string).strip() in ["", "Not Found", "Error", "nan"]:
        return "", "", ""
        
    series_list = []
    type_list = []
    models_list = []
    
    # Split by the pipe separator used in our scraping
    entries = str(model_string).split(" | ")
    
    for entry in entries:
        # Use Regex to split the entry into: Series Code, Type, and Specific Models
        # Example: "G121 ForkliftTrucks(Engine) 7FG/7FD10,15,20,25,30,J35,K20-K30"
        match = re.match(r"^(\S+)\s+(.*?)\s+(.*)$", entry)
        
        if match:
            series_list.append(match.group(1))
            type_list.append(match.group(2))
            models_list.append(match.group(3))
        else:
            # Fallback if regex fails
            series_list.append(entry)
            
    return " | ".join(series_list), " | ".join(list(set(type_list))), " | ".join(models_list)

def main():
    print(f"Loading {FILE_NAME}...")
    # Load only the Toyota sheet
    df = pd.read_excel(FILE_NAME, sheet_name="Part Numbers")
    
    # Create the new columns
    print("Parsing Model data into separate columns...")
    parsed_data = df["Model"].apply(parse_models)
    
    # Assign the parsed data to new columns
    df["Series Code"] = [x[0] for x in parsed_data]
    df["Type"] = [x[1] for x in parsed_data]
    df["Specific Models"] = [x[2] for x in parsed_data]
    
    print(f"Saving updated data back to {FILE_NAME}...")
    with pd.ExcelWriter(FILE_NAME, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Part Numbers", index=False)
        
    print("Done! You can now open the Excel file.")

if __name__ == "__main__":
    main()
