import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time
import os

# Function to clear the terminal
def clear_terminal():
    os.system('cls' if os.name == 'nt' else 'clear')

# Function to clean up the email addresses
def clean_email(email):
    # Regular expression to identify the valid end of an email address
    match = re.search(r'@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', email)
    if match:
        end_index = match.end()
        return email[:end_index]
    return email

# Function to get the email address using web scraping
def get_company_email(company_name, search_engines):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    
    max_attempts = len(search_engines)
    initial_delay = 5
    delay = initial_delay
    search_engine_index = 0

    print(f"Searching for email for {company_name}...")

    for attempt in range(max_attempts):
        search_url = search_engines[search_engine_index].format(company_name)
        search_engine_name = search_engines[search_engine_index].split('.')[1]
        print(f"Attempt {attempt + 1}: Using {search_engine_name.capitalize()} - {search_url}")

        try:
            response = requests.get(search_url, headers=headers)
            print(f"Response Status: {response.status_code}")
            
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, "html.parser")
                email = None
                
                # Regular expression to find email addresses
                email_regex = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
                emails = re.findall(email_regex, soup.text)
                
                if emails:
                    cleaned_email = clean_email(emails[0])
                    print(f"Email found: {cleaned_email}")
                    return cleaned_email
                print("No email found in search results.")
            elif response.status_code == 429:
                print("Rate limit exceeded. Switching search engine...")
                search_engine_index = (search_engine_index + 1) % len(search_engines)
                time.sleep(delay)
                delay = initial_delay  # Reset delay
            else:
                response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"Request failed: {e}")
            search_engine_index = (search_engine_index + 1) % len(search_engines)
            time.sleep(delay)
            delay = initial_delay  # Reset delay

    print("Max attempts reached. No email found.")
    return '404'

# Main function to read company names from Excel and write results back to Excel
def main():
    input_file = 'C:\\Users\\######\\Downloads\\Map6.xlsx'  # Your own path to the file
    output_file = 'C:\\Users\\######\\Downloads\\Hotel_with_emails.xlsx'
    
    search_engines = [
        "https://www.google.com/search?q=Hotel+Belgie+{}+email",
        "https://www.bing.com/search?q=Hotel+Belgie+{}+email",
        "https://duckduckgo.com/?q=Hotel+Belgie+{}+email"
    ]

    # Load the Excel file without headers
    df = pd.read_excel(input_file, header=None)

    # Add a new column for the emails if it doesn't exist
    if 'Email' not in df.columns:
        df['Email'] = None

    start_time = time.time()

    while df['Email'].str.contains('404').any():
        for idx, row in df.iterrows():
            if row['Email'] == '404' or pd.isna(row['Email']):
                clear_terminal()
                company_name = row[0]  # Accessing the first column
                print(f"\nProcessing company {idx + 1}/{len(df)}: {company_name}")
                print("-" * 60)
                email = get_company_email(company_name, search_engines)
                df.at[idx, 'Email'] = email
                elapsed_time = time.time() - start_time
                estimated_total_time = (elapsed_time / (idx + 1)) * len(df)
                estimated_remaining_time = estimated_total_time - elapsed_time
                print("-" * 60)
                print(f"Estimated time remaining: {time.strftime('%H:%M:%S', time.gmtime(estimated_remaining_time))}")
                if email != '404':
                    print(f"\nEmail for {company_name}: {email}")
                    time.sleep(1)

        # Save the intermediate results to a new Excel file
        df.to_excel(output_file, index=False, header=False)

    print(f"\nFinal results saved to {output_file}")

if __name__ == "__main__":
    main()
