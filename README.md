# complementary-items
# Complementary Items Finder

This Python script analyzes purchase data in Excel invoices to identify top co-purchased items. It helps uncover patterns in customer buying behavior, which can be useful for product recommendations, cross-selling, and marketing strategies.

How It Works

# 1. Load Data
- Reads an Excel file (e.g., invoices.xlsx) into a pandas DataFrame.
- Assumes the relevant data is in the first sheet of the workbook.

# 2. Group Items by Invoice
- Groups the purchased items by invoice number.
- Collects all items bought in the same transaction.

# 3. Count Co-occurrences
- For each invoice, generates all item pairs.
- Counts how many times each pair of items was purchased together across invoices.

# 4. Identify Top Co-Purchased Items
- For each item, identifies the top 3 items most frequently purchased alongside it.
- Stores the results in a convenient, structured format.

# 5. Output Results
Creates a pandas DataFrame with columns:

- Base Item – the original item
- Top Co-purchased Item – the most frequent associated item
- Invoice Count – how many invoices included both items

Saves the results to an Excel file and prints the top 10 results for quick inspection.
