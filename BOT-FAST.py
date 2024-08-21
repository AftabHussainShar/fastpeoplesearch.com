import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import os
from webdriver_manager.chrome import ChromeDriverManager
import os

def extract_data(url):
    chrome_options = Options()
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.set_page_load_timeout(10)
    
    try:
        driver.get(url)
    except TimeoutException:
        driver.execute_script("window.stop();")
        driver.execute_script("document.dispatchEvent(new KeyboardEvent('keydown', {'key': 'Escape'}));")

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, 'card'))
        )
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        cards = soup.find_all('div', class_='card')
        
        results = []
        for card in cards:
            try:
                if card.get('id') and card.get('data-link'):
                    span = card.find('span', class_='larger')
                    name = span.text.strip()
                    phone_number = None
                    phone_element = card.select_one('div[class="card-block"] > strong > a')
                    if phone_element:
                        phone_number = phone_element.text.strip()
                    results.append({'Name': name, 'Number': phone_number})
                    
            except AttributeError:
                continue
        
    except TimeoutException:
        print(f"Timeout occurred while processing URL: {url}")
        results = None
    
    finally:
        driver.quit()
    
    return results

csv_file = 'people.csv'
df = pd.read_csv(csv_file)


for index, row in df.iterrows():
    output_data = []
    first_name = row['First Name']
    last_name = row['Last Name']
    url = row['URL']
    
    print(f"Processing URL for {first_name} {last_name}: {url}")
    
    results = extract_data(url)
    
    if results:
        person_results = {
            'First Name': first_name,
            'Last Name': last_name,
            'URL': url
        }
        for i in range(min(len(results), 7)):
            person_results[f'Result {i+1} Name'] = results[i]['Name']
            person_results[f'Result {i+1} Number'] = results[i]['Number']
        output_data.append(person_results)
    
    else:
        output_data.append({
            'First Name': first_name,
            'Last Name': last_name,
            'URL': url
        })

    output_df = pd.DataFrame(output_data)
    file_exists = os.path.isfile('output_results.csv')
    mode = 'a' if file_exists else 'w'  
    output_df.to_csv('output_results.csv', index=False, mode=mode, header=not file_exists)
    print("Extraction and processing completed. Results saved to output_results.csv.")
