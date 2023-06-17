import scrapy
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
from ..items import SeleniumscrapyItem




class SeleniumscraperSpider(scrapy.Spider):
    name = "seleniumscraper"


    def start_requests(self):

        #Part-I - Extracting links using selenium to join with url to access the desired webpage

        chrome_options = Options()

        chrome_options.headless = False
        # chrome_options.add_argument("--window-position=-2000,0")

        chromedriver_path = r"C:\\Users\\muhammad.naveed\\Downloads\\chromedriver_win32\\chromedriver.exe"

        website = 'https://www.zipmec.eu/en/cards.html'

        service = Service(executable_path=chromedriver_path)

        driver = webdriver.Chrome(service=service, options=chrome_options)

        driver.get(website)

        wait = WebDriverWait(driver, 5)

        cookie_reject = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="c-s-bn"]'))).click()

        choose_product_button = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div[2]/div/div/aside/div[2]/div[2]/div[1]/div[1]'))).click()

        # on selection of each product, webpage shows list of companies that deals with that product
        # so first step is to collect all products in a list

        product_list = []

        products = wait.until(
            EC.presence_of_all_elements_located((By.XPATH, '//*[@id="mod_rg_select_prodotti_chzn"]/div/ul/li')))


        # iterating over each product so that program itself select the product and then seacrh for companies

        for product in products[1:]:
            y = product.text
            product_list.append(y)

            total_href_list = []

        for product_name in product_list:

            product_field = driver.find_element(By.XPATH, '//*[@id="mod_rg_select_prodotti_chzn"]/div/div/input')
            product_field.send_keys(product_name)
            product_field.send_keys(Keys.ENTER)

            time.sleep(2)

            company_found = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="mod_rg_risultati"]'))).text

            if company_found == "0":
                new_search = wait.until(EC.visibility_of_element_located((By.ID, 'mod_rg_btn_reset'))).click()

                choose_product_button = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, '/html/body/div[3]/div[2]/div/div/aside/div[2]/div[2]/div[1]/div[1]'))).click()
                continue

            see_companies = wait.until(EC.element_to_be_clickable((By.ID, 'mod_rg_btn_vedi'))).click()

            time.sleep(3)

            # after selection of product and clicking on search,minimum 10 companies are shown on first page
            # that's why iteration in range of 10

            for x in range(10):
                x = x + 1
                try:
                    hrefs = wait.until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.riga-' + str(x) + ' div div h4 a')))

                    for href in hrefs:
                        value = href.get_attribute("data-campo")
                        if value not in total_href_list:
                            total_href_list.append(value)
                except:
                    # In case the first page contains less than 4 (or any number less than 10),
                    # in that case it will come to except block and breakout of loop
                    # This will save time over unrequired iteration
                    break

            print("Product Name:", product_name, "completed")

            driver.execute_script("window.scrollTo(0, 0);")

            time.sleep(2)

            new_search = wait.until(EC.visibility_of_element_located((By.ID, 'mod_rg_btn_reset'))).click()

            choose_product_button = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div[2]/div/div/aside/div[2]/div[2]/div[1]/div[1]'))).click()


        #Step - II: Ieterating over each link and joining with URL to accesss desired webpage

        url = 'https://www.zipmec.eu'

        for link in total_href_list:
            join_url = url + link

            yield scrapy.Request(url=join_url, callback=self.parse)



    def parse(self, response):

        #Step - III : parsing data from desired webpage

        items = SeleniumscrapyItem()

        Company_name = response.css(".edysmabox-body li:nth-child(1)::text").get()
        Company_address = response.css(".edysmabox-body li:nth-child(2)::text").get()
        town_city = response.css(".edysmabox-body li:nth-child(3)::text").get()
        province = response.css(".edysmabox-body li:nth-child(4)::text").get()
        state = response.css(".edysmabox-body li:nth-child(5)::text").get()
        country = response.css(".edysmabox-body li:nth-child(6)::text").get()
        latitude = response.css(".edysmabox-body li:nth-child(7)::text").get()
        longitude = response.css(".edysmabox-body li:nth-child(8)::text").get()

        try:
            fax = response.css(".uk-icon-fax~ span .contatti-click").css('::attr(data-campo)').extract()

        except:
            fax = None

        try:
            telephone = response.css(".uk-icon-phone~ span .contatti-click").css('::attr(data-campo)').extract()

        except:
            telephone = None

        try:
            telephone_fax = response.css(".uk-icon-volume-control-phone~ span .contatti-click").css(
                '::attr(data-campo)').extract()

        except:
            telephone_fax = None

        try:
            email = response.css(".uk-icon-envelope-o~ span .contatti-click").css('::attr(data-campo)').extract()
        except:
            email = None

        try:
            website = response.css(".uk-icon-share-square-o~ span .contatti-click").css('::attr(data-campo)').extract()
        except:
            website = None

        try:
            mobile = response.css(".uk-icon-mobile-phone~ span .contatti-click").css('::attr(data-campo)').extract()

        except:
            mobile = None


        items['Company_name'] = Company_name
        items['Company_address'] = Company_address
        items['Town_city'] = town_city
        items['Province'] = province
        items['Region_state'] = state
        items['Country'] = country
        items['Latitude'] = latitude
        items['Longitude'] = longitude
        items['fax'] = fax
        items['telephone'] = telephone
        items['telephone_fax'] = telephone_fax
        items['email'] = email
        items['website'] = website
        items['mobile'] = mobile


        yield items
