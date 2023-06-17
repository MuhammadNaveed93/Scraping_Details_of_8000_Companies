# Scraping_Details_of_8000_Companies
The project involved scraping details of 8000+ companies listed on a website. The goal was to retrieve essential data such as names, addresses, city, country, and contact details for each company

Challenge:
The challenge was to interact with the website by entering product names one by one in the search bar to retrieve the list of companies.Each company had to be clicked individually, opening a new window (with unique URL) where the desired company information was available

Project Solution:
To overcome the challenge, I utilized both Selenium and Scrapy for the project. With Selenium, first I extracted the product lists and performed actions such as entering each product name one by one and clicking the search button to display associated companies. Subsequently, scraped URLs for each company. 
Scrapy was then used to access each company (utilizing extracted web links) for scraping the desired information. This two-step approach enabled me to retrieve data from both dynamic and static webpages, ensuring a comprehensive data collection process.

Project Success: 
Using this approach, I successfully extracted detailed information from approximately 8,000 companies. The final Excel file presents accurate data, organized in an optimal format. 
