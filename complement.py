import pandas as pd
from collections import Counter
from itertools import combinations

# Load the Excel file into a pandas DataFrame
# Assuming your Excel file is named 'invoices.xlsx' and the sheet is the first sheet.
df = pd.read_excel(r'invoices.xlsx',sheet_name=0)

# Group the data by invoice number and collect the items purchased in each invoice
grouped = df.groupby('invoice_number')['item_no'].apply(list).reset_index()

# Create a dictionary to count co-occurrences of items
co_occurrence_counter = Counter()

# For each invoice, find all item pairs and count their co-occurrences
for items in grouped['item_no']:
    for item1, item2 in combinations(items, 2):
        if item1 != item2:
            # Use frozenset to make sure the pairs are considered as undirected (item1, item2) == (item2, item1)
            co_occurrence_counter[frozenset([item1, item2])] += 1

# Create a dictionary to store the top 3 co-purchased items for each item
top_3_items = {}

# Collect co-occurrence data into a more convenient structure
co_occurrence_data = {}
for (item1, item2), count in co_occurrence_counter.items():
    if item1 not in co_occurrence_data:
        co_occurrence_data[item1] = Counter()
    if item2 not in co_occurrence_data:
        co_occurrence_data[item2] = Counter()
    co_occurrence_data[item1][item2] += count
    co_occurrence_data[item2][item1] += count

# Find the top 3 items for each item
for item, counter in co_occurrence_data.items():
    top_3 = counter.most_common(3)
    top_3_items[item] = top_3

# Create a DataFrame to display the results
results = []
for item, top_items in top_3_items.items():
    for top_item, count in top_items:
        results.append([item, top_item, count])

result_df = pd.DataFrame(results, columns=['Base Item', 'Top Co-purchased Item', 'Invoice Count'])

# Save the result to an Excel file
result_df.to_excel('top_3_co_purchased_items.xlsx', index=False)

# Print the result DataFrame
print(result_df.head(10))
