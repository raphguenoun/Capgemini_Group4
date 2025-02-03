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
    """
    Web scraper for collecting customer reviews from Trustpilot.
    Uses Selenium for dynamic content handling and implements anti-detection measures.
    """
    
    def __init__(self, url):
        self.url = url
        self.setup_driver()

    def setup_driver(self):
        """Configure Chrome driver with anti-bot detection settings"""
        chrome_options = Options()
        # Set browser configurations to mimic human behavior
        chrome_options.add_argument('--no-sandbox') 
        chrome_options.add_argument('--disable-dev-shm-usage') 
        chrome_options.add_argument('--window-size=1920,1080')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        # Mask automation signatures
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.wait = WebDriverWait(self.driver, 15) # Timeout for waiting for elements

    def random_sleep(self, min_time=2, max_time=5):
        """Add random delays to mimic human browsing patterns"""
        time.sleep(random.uniform(min_time, max_time))

    def scrape_page(self):
        """
        Extract review data from current page with retry mechanism
        Returns: list of dictionaries containing review data
        """
        reviews = []
        max_retries = 3
        retries = 0

        while retries < max_retries:
            try:
                # Wait for reviews to load and simulate scrolling for dynamic content
                review_elements = self.wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, '[data-service-review-card-paper="true"]'))
                )
                
                # Extract structured data from each review
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
                retries += 1
                if retries < max_retries:
                    self.random_sleep(5, 10)
                    self.setup_driver()
                else:
                    raise

        return reviews

    def scrape_pages(self, num_pages=10):
        """
        Iterate through multiple pages and collect all reviews
        Args:
            num_pages: Number of pages to scrape
        Returns:
            List of all reviews
        """
        all_reviews = []
        
        for page in range(1, num_pages + 1):
            current_url = f"{self.url}?page={page}"
            try:
                self.driver.get(current_url)
                self.random_sleep()
                page_reviews = self.scrape_page()
                all_reviews.extend(page_reviews)
            except Exception as e:
                break

        return all_reviews

    def save_to_excel(self, reviews, filename="reviews_total_energies.xlsx"):
        """
        Export collected reviews to Excel for data analysis
        Args:
            reviews: List of review dictionaries
            filename: Output Excel file name
        """
        df = pd.DataFrame(reviews)
        df.to_excel(filename, index=False, engine='openpyxl')

def main():
    # Example usage targeting TotalEnergies reviews
    url = "https://fr.trustpilot.com/review/totalenergies.fr"
    scraper = TrustpilotScraper(url)
    
    try:
        reviews = scraper.scrape_pages(10)  # Collect 10 pages of reviews
        scraper.save_to_excel(reviews)
    finally:
        scraper.close()

if __name__ == "__main__":
    main()