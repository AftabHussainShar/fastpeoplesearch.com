import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager
from concurrent.futures import ThreadPoolExecutor, as_completed
import os

def extract_data(url):
    chrome_options = Options()
    service = Service(ChromeDriverManager().install())
    
    with webdriver.Chrome(service=service, options=chrome_options) as driver:
        driver.set_page_load_timeout(10)
        results = []
        try:
            driver.get(url)
        except TimeoutException:
            driver.execute_script("window.stop();")
            driver.execute_script("document.dispatchEvent(new KeyboardEvent('keydown', {'key': 'Escape'}));")
        
        try:
            # driver.get(url)
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, 'card'))
            )
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            cards = soup.find_all('div', class_='card')
            
            for card in cards:
                try:
                    if card.get('id') and card.get('data-link'):
                        span = card.find('span', class_='larger')
                        name = span.text.strip() if span else 'No Name'
                        phone_element = card.select_one('div.card-block > strong > a')
                        phone_number = phone_element.text.strip() if phone_element else 'No Number'
                        results.append({'Name': name, 'Number': phone_number})
                except AttributeError:
                    continue
        
        except TimeoutException:
            print(f"Timeout occurred while processing URL: {url}")
        except NoSuchElementException:
            print(f"Element not found for URL: {url}")
        except Exception as e:
            print(f"Unexpected error occurred while processing URL {url}: {e}")
    
    return results

def process_chunk(chunk, output_file):
    results = []
    for index, row in chunk.iterrows():
        first_name = row['First Name']
        last_name = row['Last Name']
        url = row['URL']
        
        print(f"Processing URL for {first_name} {last_name}: {url}")
        
        extracted_data = extract_data(url)
        
        person_results = {
            'First Name': first_name,
            'Last Name': last_name,
            'URL': url
        }
        for i in range(7):
            if i < len(extracted_data):
                person_results[f'Result {i+1} Name'] = extracted_data[i]['Name']
                person_results[f'Result {i+1} Number'] = extracted_data[i]['Number']
            else:
                person_results[f'Result {i+1} Name'] = None
                person_results[f'Result {i+1} Number'] = None
        
        results.append(person_results)
    
    output_df = pd.DataFrame(results)
    file_exists = os.path.isfile(output_file)
    mode = 'a' if file_exists else 'w'
    output_df.to_csv(output_file, index=False, mode=mode, header=not file_exists)
    print(f"Saved")
    
    return results

def main():
    csv_file = 'people.csv'
    output_file = 'output_results.csv'
    
    if not os.path.isfile(csv_file):
        print(f"CSV file {csv_file} not found.")
        return
    
    df = pd.read_csv(csv_file)
    df = df.drop_duplicates(subset='URL').reset_index(drop=True)
    
    if df.empty:
        print("No unique rows to process.")
        return
    
    num_workers = 5
    chunk_size = len(df) // num_workers + (len(df) % num_workers > 0)
    chunks = [df[i:i + chunk_size] for i in range(0, len(df), chunk_size)]
    
    with ThreadPoolExecutor(max_workers=num_workers) as executor:
        future_to_chunk = {executor.submit(process_chunk, chunk, output_file): chunk for chunk in chunks}
        
        for future in as_completed(future_to_chunk):
            chunk = future_to_chunk[future]
            try:
                result = future.result()
                print(f"Saved")
            except Exception as e:
                print(f"Error occurred while processing chunk {chunk}: {e}")

if __name__ == "__main__":
    main()
