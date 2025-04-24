# Excel Prospect List Updater (Python)

This script automates the process of updating and maintaining a prospect list stored in Excel format. It merges existing leads with new entries, removes duplicates based on email addresses, allows optional sorting by column, and exports the result in both Excel and CSV formats.

## 🚀 Features

- Merge two Excel sheets: `prospect_list.xlsx` and `new_leads.xlsx`
- Remove duplicate leads (by email)
- Sort the final output by any valid column (e.g., Name, Email, Company)
- Automatically adjusts Excel column widths using `openpyxl`
- Adds freeze panes and filters to the first row in Excel output
- Outputs final data as:
  - `prospect_list_updated.xlsx`
  - `prospect_list_updated.csv`

## 🛠 Requirements

Install dependencies using pip:

```
pip install pandas openpyxl faker
```

## 📦 Usage

Place your Excel files in the working directory:
- `prospect_list.xlsx` – your original contact list
- `new_leads.xlsx` – new entries to be added

Run the script from the command line:

```
python prospect_updater.py --sort Name
```

### Optional Arguments:
- `--sort [COLUMN]` – Column to sort by (default: Name)

If an invalid column is specified, the script will list valid options and exit safely.

## 📂 Output

- `prospect_list_updated.xlsx` (with formatted columns and filters)
- `prospect_list_updated.csv`

## 🔧 Example Use Cases

- Keeping your CRM prospect list up to date
- Automating lead imports from different sources
- Pre-cleaning Excel data before analytics

## 👨‍💻 Built With

- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [faker](https://faker.readthedocs.io/)

---

Happy automating!
