import pandas as pd
import openpyxl


# Load and group Epicor data
epicor = pd.read_excel('excel_sheets/Epicor.xlsx')
epicor_grouped = epicor.groupby('order_#')['MRR'].sum().reset_index()

# Load SFD data
sfd = pd.read_excel('excel_sheets/SFD.xlsx')

# Merge grouped Epicor data with SFD data
merged_df = pd.merge(epicor_grouped, sfd, on='order_#', how='inner', suffixes=('_epicor', '_sfd'))

# Compute absolute difference 
merged_df['MRR_diff'] = abs(merged_df['MRR_epicor'] - merged_df['MRR_sfd'])

# Filter only rows where the absolute difference is greater than 1 
variances = merged_df[merged_df['MRR_diff'] > 1]
# After merging and filtering your data

# Rename the columns
variances = variances.rename(columns={
    'order_#': 'Order #',
    'MRR_epicor': 'MRR Epicor',
    'customer_name': 'Customer',
    'MRR_sfd': 'MRR SF',
    'MRR_diff': 'Variance'
})

# Save the renamed DataFrame to a new Excel file
variances.to_excel('excel_sheets/Variances.xlsx', index=False)
