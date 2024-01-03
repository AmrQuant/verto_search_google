import requests
from urllib.parse import urlparse
import pandas as pd
from datetime import datetime
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment


def query_google(query, api_key, cse_id, start_page=1, num_results=10, date_range=None):
    url = "https://www.googleapis.com/customsearch/v1"
    results = []
    while len(results) < num_results:
        params = {
            'q': query,
            'key': api_key,
            'cx': cse_id,
            'start': start_page
        }
        if date_range:
            params['dateRestrict'] = date_range

        response = requests.get(url, params=params)
        if response.status_code == 200:
            response_data = response.json()
            results.extend(response_data.get("items", []))
            start_page += 10
            total_results = int(response_data.get('searchInformation', {}).get('totalResults', "0"))
            if start_page > total_results:
                break
        else:
            print(f"Error occurred: {response.status_code}")
            break

    return results[:num_results]

def extract_domain(url):
    parsed_url = urlparse(url)
    domain_parts = parsed_url.netloc.split('.')
    if len(domain_parts) >= 2:
        return domain_parts[-2].capitalize()
    return "Unknown"

def extract_search_results(results):
    extracted_results = []
    social_media_domains = ['facebook', 'twitter', 'instagram', 'tiktok', 'snapchat', 'linkedin', 'youtube']
    
    for item in results:
        domain_name = extract_domain(item.get('link'))
        # Check if the domain is a known social media platform
        is_social_media = 'Yes' if domain_name.lower() in social_media_domains else 'No'

        # Attempt to extract the publication date
        date = item.get('pagemap', {}).get('metatags', [{}])[0].get('article:published_time', 'Unknown')

        extracted_data = {
            'date': date,
            'title': item.get('title'),
            'link': item.get('link'),
            'source_name': domain_name,
            'Social_Media': is_social_media,
            'search_engine_name': 'Google'
        }
        extracted_results.append(extracted_data)
    return extracted_results




def sanitize_filename(filename):
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '')
    return filename


import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment

# ... (Other parts of the script) ...

def save_to_excel(data, search_query, directory_path=''):
    today_date = datetime.now().strftime('%Y-%m-%d')
    sanitized_query = sanitize_filename(search_query.replace(' ', '_'))
    base_filename = f"{sanitized_query}_{today_date}"
    extension = ".xlsx"
    filename = base_filename + extension
    full_path = os.path.join(directory_path, filename)

    # Check if the file exists and change the filename by appending a number in brackets
    counter = 1
    while os.path.exists(full_path):
        filename = f"{base_filename}({counter}){extension}"
        full_path = os.path.join(directory_path, filename)
        counter += 1

    wb = Workbook()
    ws = wb.active

    # Convert DataFrame to rows
    df = pd.DataFrame(data)
    rows = dataframe_to_rows(df, index=False, header=True)

    # Write rows, set hyperlink, and format headers
    for r_idx, row in enumerate(rows, 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if r_idx == 1:
                # For headers, make the text bold and capitalize
                cell.font = Font(bold=True)
                cell.value = cell.value.title()
            elif r_idx > 1 and c_idx == df.columns.get_loc('link') + 1:
                cell.hyperlink = value
                cell.font = Font(color='0000FF', underline='single')

    # Freeze the top row and apply filter
    ws.freeze_panes = 'A2'  # Freeze the top row
    ws.auto_filter.ref = ws.dimensions  # Apply filter to all columns

    wb.save(full_path)
    print(f"Results saved to {full_path}")





def main():
    api_key = 'COPY_PAST_HERE'
    cse_id = 'COPY_PAST_HERE'
    user_query = input("Enter your search query: ")
    num_results = int(input("Enter the number of results you want: "))
    date_range_input = input("Enter the date range (e.g., 'd5' for last 5 days): ")
    directory_path = input("Enter the directory path to save the file (leave empty for current directory): ")

    results = query_google(user_query, api_key, cse_id, num_results=num_results, date_range=date_range_input)
    if results:
        search_results = extract_search_results(results)
        save_to_excel(search_results, user_query, directory_path)

if __name__ == "__main__":
    main()
