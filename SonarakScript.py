import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
import re
import urllib.parse

# Define a configuration mapping for each website
CONFIG = {
    'Solarkal': {
        'description': {'section': 'home-hero', 'tag': 'p', 'class': 'subhead'},
        'hq_offices': {'tag': 'p', 'class': 'contact-p'},  # Updated for HQ and offices
        'clients': {'div': 'slider', 'tag': 'img', 'attribute': 'src'},
        'news': {'div': 'w-dyn-list', 'item_class': 'collection-item-3 w-dyn-item', 'title_tag': 'h3', 'date_index': 2}
    }
}

def extract_client_name_from_url(url):
    # Decode URL-encoded characters
    decoded_url = urllib.parse.unquote(url)

    # Extract the part of the URL before '.png' or '.jpg' etc.
    match = re.search(r'/([^/]+)\.(png|jpg|jpeg|gif)$', decoded_url)
    if match:
        name_part = match.group(1)

        # Split by underscores and remove unwanted parts
        name_parts = name_part.split('_')  # Split by underscores
        client_name = name_parts[-1]  # Get the last part

        # Clean the client name
        client_name = client_name.replace('-logo', '').strip()  # Remove '-logo'
        client_name = client_name.replace('%20', ' ').strip()  # Replace '%20' with space

        return client_name

    return ''

def scrape_company_info(base_url, blog_url, company_name):
    data = {
        'Description': '',
        'HQ and Offices': '',
        'Clients': [],
        'News': []
    }

    config = CONFIG.get(company_name, {})

    def extract_description(url):
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        desc_config = config.get('description', {})
        if desc_config:
            desc_tag = soup.find(desc_config.get('tag'), class_=desc_config.get('class'))
            if desc_tag:
                data['Description'] = desc_tag.get_text(strip=True)

    def extract_hq_offices(url):
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        hq_config = config.get('hq_offices', {})
        if hq_config:
            hq_section = soup.find(hq_config.get('tag'), class_=hq_config.get('class'))
            if hq_section:
                # Extract text and format address
                address = hq_section.get_text(separator=' ', strip=True)  # Join lines with space
                data['HQ and Offices'] = address

    def extract_clients(url):
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        clients_config = config.get('clients', {})
        if clients_config:
            clients_section = soup.find('div', class_=clients_config.get('div'))
            if clients_section:
                img_tags = clients_section.find_all(clients_config.get('tag'))
                for img in img_tags:
                    img_src = img.get(clients_config.get('attribute'), '').strip()
                    if img_src:
                        client_name = extract_client_name_from_url(img_src)
                        if client_name and client_name not in data['Clients']:
                            data['Clients'].append(client_name)

    def extract_news(url):
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        news_config = config.get('news', {})
        if news_config:
            news_section = soup.find('div', class_=news_config.get('div'))
            if news_section:
                news_items = news_section.find_all('div', class_=news_config.get('item_class'))
                for item in news_items:
                    title = item.find(news_config.get('title_tag')).get_text(strip=True) if item.find(news_config.get('title_tag')) else ''
                    date = item.find_all('p')[news_config.get('date_index')].get_text(strip=True) if len(item.find_all('p')) > news_config.get('date_index') else ''
                    url = item.find('a').get('href') if item.find('a') else ''
                    summary = item.find('div').find('p').get_text(strip=True) if item.find('div') and item.find('div').find('p') else ''
                    data['News'].append({
                        'Title': title,
                        'Date': date,
                        'URL': url,
                        'Summary': summary
                    })

    # Extract information from the base URL
    extract_description(base_url)
    # Update URL for HQ and Offices extraction
    extract_hq_offices('https://www.solarkal.com/contact-us')
    extract_clients(base_url)

    # Extract news from the blog URL
    extract_news(blog_url)

    return data

def save_to_excel(data, filename='company_data.xlsx'):
    try:
        df_main = pd.DataFrame({
            'Company': [name for name in data],
            'Description': [info['Description'] for info in data.values()],
            'HQ and Offices': [info['HQ and Offices'] for info in data.values()],
            'Clients': [', '.join(info['Clients']) for info in data.values()]
        })

        df_news = pd.DataFrame([news for info in data.values() for news in info['News']])

        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df_main.to_excel(writer, sheet_name='Company Info', index=False)
            df_news.to_excel(writer, sheet_name='News', index=False)

        print(f"Data has been successfully saved to '{filename}'")

    except ImportError:
        print("Error: The 'openpyxl' package is required for saving to Excel. Install it using 'pip install openpyxl'.")
    except Exception as e:
        print(f"An error occurred while saving to Excel: {e}")

if __name__ == '__main__':
    companies = [
        (5875, 'Solarkal', 'https://www.solarkal.com/'),
        # Add more companies as needed
    ]

    all_data = {}
    blog_suffix = '/blog'

    for company_id, company_name, website in companies:
        print(f"Scraping data for {company_name}...")
        company_data = scrape_company_info(website, website + blog_suffix, company_name)
        print(company_data)
        all_data[company_name] = company_data

    if all_data:
        save_to_excel(all_data)

    print("All data has been successfully scraped and saved.")
