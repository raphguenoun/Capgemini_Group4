from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
import pandas as pd
import time
import random

class TrustpilotScraper:
    def __init__(self, url):
        self.url = url
        self.setup_driver()

    def setup_driver(self):
        chrome_options = Options()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--window-size=1920,1080')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.wait = WebDriverWait(self.driver, 15)

    def random_sleep(self, min_time=2, max_time=5):
        time.sleep(random.uniform(min_time, max_time))

    def get_review_title(self, review):
        try:
            return review.find_element(By.CSS_SELECTOR, '[data-service-review-title-typography="true"]').text
        except NoSuchElementException:
            return None

    def get_review_text(self, review):
        try:
            return review.find_element(By.CSS_SELECTOR, '[data-service-review-text-typography="true"]').text
        except NoSuchElementException:
            return None

    def get_rating(self, review):
        try:
            rating_div = review.find_element(By.CSS_SELECTOR, '[data-service-review-rating]')
            return rating_div.get_attribute('data-service-review-rating')
        except NoSuchElementException:
            return None

    def get_date(self, review):
        try:
            date_element = review.find_element(By.CSS_SELECTOR, '[data-service-review-date-of-experience-typography="true"]')
            return date_element.text.replace('Date de l\'expérience: ', '')
        except NoSuchElementException:
            return None

    def get_supplier_response(self, review):
        try:
            return review.find_element(By.CSS_SELECTOR, '[data-service-review-business-reply-text-typography="true"]').text
        except NoSuchElementException:
            return None

    def scrape_page(self):
        reviews = []
        max_retries = 3
        retries = 0

        while retries < max_retries:
            try:
                review_elements = self.wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, '[data-service-review-card-paper="true"]'))
                )

                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                self.random_sleep(1, 2)
                self.driver.execute_script("window.scrollTo(0, 0);")
                self.random_sleep(1, 2)

                for review in review_elements:
                    review_data = {
                        'title': self.get_review_title(review),
                        'text': self.get_review_text(review),
                        'rating': self.get_rating(review),
                        'date': self.get_date(review),
                        'supplier_response': self.get_supplier_response(review)
                    }
                    reviews.append(review_data)
                break

            except (TimeoutException, WebDriverException) as e:
                print(f"Tentative {retries + 1} échouée: {str(e)}")
                retries += 1
                if retries < max_retries:
                    print("Réessai après une courte pause...")
                    self.random_sleep(5, 10)
                    self.setup_driver()
                else:
                    print("Échec après plusieurs tentatives")
                    raise

        return reviews

    def scrape_pages(self, num_pages=10):
        all_reviews = []
        
        for page in range(1, num_pages + 1):
            current_url = f"{self.url}?page={page}"
            print(f"Scraping de la page {page}...")
            
            try:
                self.driver.get(current_url)
                self.random_sleep()
                
                page_reviews = self.scrape_page()
                all_reviews.extend(page_reviews)
                
                print(f"Page {page} : {len(page_reviews)} avis récupérés")
                
            except Exception as e:
                print(f"Erreur sur la page {page}: {str(e)}")
                break

        return all_reviews

    def save_to_excel(self, reviews, filename="reviews_total_energies.xlsx"):
        df = pd.DataFrame(reviews)
        df.to_excel(filename, index=False, engine='openpyxl')
        print(f"Données sauvegardées dans {filename}")

    def close(self):
        try:
            self.driver.quit()
        except:
            pass

def main():
    url = "https://fr.trustpilot.com/review/totalenergies.fr"
    scraper = TrustpilotScraper(url)
    
    try:
        reviews = scraper.scrape_pages(10)
        scraper.save_to_excel(reviews)
    finally:
        scraper.close()

if __name__ == "__main__":
    main()