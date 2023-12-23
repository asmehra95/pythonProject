import pandas as pd
import numpy as np

def get_marketplace_code(marketplace, selected_countries):
    """
    Convert marketplace to its corresponding ISO country code.
    If not in selected countries, return 'Other'.
    """
    marketplace_codes = {
        'Amazon.com': 'US',
        'Amazon.ca': 'CA',
        'Amazon.co.uk': 'UK',
        'Amazon.com.au': 'AU'
    }
    code = marketplace_codes.get(marketplace, 'Other')
    return code if code in selected_countries else 'Other'

def generate_sheet_name(marketplace, title, idx, selected_countries):
    """
    Generate a valid Excel sheet name using marketplace ISO code,
    first two words of the title, and an index.
    """
    code = get_marketplace_code(marketplace, selected_countries)
    short_title = '_'.join(title.split()[:2])
    clean_title = ''.join(char for char in short_title if char not in '[]:*?/\\')[:10]
    name = f"{code}_{clean_title}_{idx}"
    return name[:31]

def aggregate_data(df, selected_books):
    """
    Aggregate the sales data based on Title, Marketplace, Royalty Date, and Currency.
    Only include selected books.
    """
    filtered_df = df[df['Title'].isin(selected_books)]
    return filtered_df.groupby(['Title', 'Marketplace', 'Royalty Date', 'Currency']).agg(
        {'Units Sold': 'sum', 'Net Units Sold': 'sum', 'Royalty': 'sum'}).reset_index()

def create_full_date_df(start_date, end_date):
    """
    Create a DataFrame with a full range of dates between start_date and end_date.
    """
    date_range = pd.date_range(start=start_date, end=end_date)
    return pd.DataFrame(date_range, columns=['Royalty Date'])

def export_to_excel(aggregated_data, output_file_path, selected_countries):
    """
    Export the aggregated data to an Excel file with separate sheets.
    """
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        for idx, ((marketplace, title, currency), group) in enumerate(aggregated_data.groupby(['Marketplace', 'Title', 'Currency'])):
            sheet_name = generate_sheet_name(marketplace, title, idx, selected_countries)
            start_date = aggregated_data['Royalty Date'].min()
            end_date = aggregated_data['Royalty Date'].max()
            full_dates_df = create_full_date_df(start_date, end_date)
            full_dates_df['Marketplace'] = marketplace
            full_dates_df['Title'] = title
            full_dates_df['Currency'] = currency
            merged_df = pd.merge(full_dates_df, group, on=['Marketplace', 'Title', 'Royalty Date', 'Currency'], how='left')
            merged_df[['Units Sold', 'Net Units Sold', 'Royalty']] = merged_df[['Units Sold', 'Net Units Sold', 'Royalty']].fillna(0)
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)

def process_sales_data(file_path, output_file_path, selected_books, selected_countries):
    """
    Main function to process and export sales data.
    """
    df = pd.read_excel(file_path, sheet_name="Combined Sales")
    df['Royalty Date'] = pd.to_datetime(df['Royalty Date'])
    aggregated_data = aggregate_data(df, selected_books)
    export_to_excel(aggregated_data, output_file_path, selected_countries)

# Configure your selection
selected_books = ['b2', 'b1']  # Replace with the titles you're interested in
selected_countries = ['US', 'CA', 'UK', 'AU']  # Add or remove countries as needed

# Paths to your Excel file and the output file
file_path = '<AddPath>'
output_file_path = '<AddPath>'

# Call the main function
process_sales_data(file_path, output_file_path, selected_books, selected_countries)

print(f"Data exported to sheets for selected books and countries in: {output_file_path}")
