import pandas as pd
from faker import Faker
from datetime import datetime
from random import randint 
import argparse
from openpyxl import load_workbook


# --- CLI argument parsing ---
parser = argparse.ArgumentParser(description="Merge, clean, and sort prospect lists.")
parser.add_argument('--sort', type=str, default='Name', help="Column to sort by (e.g. Name, Email, Company, Date Added)")
args = parser.parse_args()

fake = Faker()

def generate_leads(n=5):
    leads = []
    for _ in range(n):
        leads.append({
            'Name': fake.name(),
            'Email': fake.unique.email(),
            'Company': fake.company(),
            'Date Added': datetime.now().strftime('%Y-%m-%d'),
            'Value': randint(1000,5000)
        })
    return pd.DataFrame(leads)

# Create original prospect list
original_df = generate_leads(10)
original_df.to_excel('prospect_list.xlsx', index=False)

# Generate new incoming leads (simulated)
new_leads_df = generate_leads(10)
new_leads_df.to_excel('new_leads.xlsx', index=False)

print("✔️ Files created: prospect_list.xlsx & new_leads.xlsx")

# Load only one sheet first to get columns
test_df = pd.read_excel('prospect_list.xlsx')

if args.sort not in test_df.columns:
    print(f"❌ Invalid sort key: '{args.sort}'")
    print("Available columns are:", list(test_df.columns))
    exit(1)

# Load original and new data
original_df = test_df  # reuse loaded one
new_leads_df = pd.read_excel('new_leads.xlsx')

# Combine the two DataFrames
combined_df = pd.concat([original_df, new_leads_df], ignore_index=True)

# Drop duplicates based on Email
clean_df = combined_df.drop_duplicates(subset='Email', keep='first')

# --- Sort ---
clean_df = clean_df.sort_values(by=args.sort)
print(f"✔️ Sorted by '{args.sort}'")

# Save final version
clean_df.to_excel('prospect_list_updated.xlsx', index=False)
clean_df.to_csv('prospect_list_updated.csv', index=False)

print(f"✔️ Updated list saved with {len(clean_df)} entries.")

# Load the Excel file just created
wb = load_workbook('prospect_list_updated.xlsx')
ws = wb.active

# Adjust column widths based on max length in each column
for column_cells in ws.columns:
    max_length = 0
    col_letter = column_cells[0].column_letter  # e.g., 'A'
    for cell in column_cells:
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    adjusted_width = max_length + 2
    ws.column_dimensions[col_letter].width = adjusted_width

ws.freeze_panes = 'A2'  # Freezes everything above row 2 (i.e. keeps row 1 visible)
ws.auto_filter.ref = ws.dimensions  # Apply filters to the entire data range

# Save the styled file
wb.save('prospect_list_updated.xlsx')
print("✨ Column widths adjusted using openpyxl.")