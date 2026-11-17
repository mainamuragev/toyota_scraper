import pandas as pd
import re

FILE_NAME = "Toyota_parts_WITH_MODELS.xlsx"
OUTPUT_FILE = "Toyota_EXACT_FLEET.xlsx"

# --- YOUR EXACT FLEET LIST ---
my_exact_fleet = [
    "8FD80N", "8FD70N", "8FD50N", "8FD40N", "8FD30", "8FDZN30", 
    "8FDZN25", "8FDZN20", "8FG30", "8FG35N", "8FGZN30", 
    "8FBMT40", "8FB30", "8FBN30", "8FB25", "8FBN25", 
    "8FB20", "8FBN20", "8FB18", "8FBN18", "8FB15",
    "RRE140B", "RRE140H", "RRE250H", "LPE200", "SPE120", 
    "SPE140S", "LHM230", "8FGJ35"
]

def find_exact_matches(model_text):
    if pd.isna(model_text) or str(model_text).strip() in ["", "Not Found", "Error", "nan"]:
        return ""
    
    matched_machines = []
    model_text = str(model_text).upper()
    
    for machine in my_exact_fleet:
        machine = machine.upper()
        
        # 1. Direct match (e.g., the text literally says "8FD80N")
        if machine in model_text:
            matched_machines.append(machine)
            continue
            
        # 2. Smart Match: Look for the series (e.g., "8FD") and the capacity number (e.g., "80")
        # Example: Machine is 8FD80N -> Series: 8FD, Number: 80
        match = re.match(r"([A-Z0-9]+?)(\d+)[A-Z]?$", machine)
        if match:
            series = match.group(1) # e.g., 8FD
            capacity = match.group(2) # e.g., 80
            
            # If the catalog text has both the series and the capacity near each other
            if series in model_text:
                # Check if the capacity appears in the text
                if re.search(rf"\b{capacity}\b", model_text) or capacity in model_text:
                    matched_machines.append(machine)
                    
    # Remove duplicates and format
    unique_matches = list(set(matched_machines))
    if unique_matches:
        return ", ".join(unique_matches)
    return ""

def main():
    print(f"Loading {FILE_NAME}...")
    df = pd.read_excel(FILE_NAME, sheet_name="Part Numbers")
    
    print("Finding exactly which of your machines use each part...")
    
    target_col = "Specific Models" if "Specific Models" in df.columns else "Model"
    
    # Apply the matching function
    df["Your Fleet That Uses This Part"] = df[target_col].apply(find_exact_matches)
    
    print(f"Saving results to {OUTPUT_FILE}...")
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Part Numbers", index=False)
        
    # Count how many parts are used by at least one of your machines
    used_parts = df[df["Your Fleet That Uses This Part"] != ""]
    
    print("\n--- SUMMARY ---")
    print(f"Total parts in your file: {len(df)}")
    print(f"Parts used by your fleet: {len(used_parts)}")
    print("Check the new column 'Your Fleet That Uses This Part' in the Excel file.")
    print(f"Done! Check {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
