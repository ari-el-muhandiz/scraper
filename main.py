from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from requests import get
import time
from xlwt import Workbook 

domain = 'https://www.tripadvisor.co.id'
link_urls = [
    # {
    #     "url": domain + '/Hotels-g294230-Yogyakarta_Region_Java-Hotels.html',
    #     "key": 'hotel-jogja',
    #     "link_selector": "#taplc_hsx_hotel_list_lite_dusty_hotels_combined_sponsored_0 .photo-wrapper > a",
    #     "more_link_selector": ".location-review-review-list-parts-ExpandableReview__cta--2mR2g",
    #     "review_content_class": "hotels-community-tab-common-Card__card--ihfZB",
    #     "review": {
    #         "tag": "q", "class": "location-review-review-list-parts-ExpandableReview__reviewText--gOmRC"
    #     },
    #     "person": {
    #         "tag": "a", "class": "ui_header_link social-member-event-MemberEventOnObjectBlock__member--35-jC"
    #     },
    #     "name": {
    #         "tag": "h1", "class": "hotels-hotel-review-atf-info-parts-Heading__heading--2ZOcD"
    #     },
    #     "tanggal": {
    #         "tag": "div", "class": "social-member-event-MemberEventOnObjectBlock__event_type--3njyv"
    #     },
    #     "nav_class": "nav next ui_button primary"
    # },
    # {
    #     "url": domain + '/Hotels-g294226-Bali-Hotels.html',
    #     "key": 'hotel-bali',
    #     "link_selector": "#taplc_hsx_hotel_list_lite_dusty_hotels_combined_sponsored_0 .photo-wrapper > a",
    #     "more_link_selector": ".location-review-review-list-parts-ExpandableReview__cta--2mR2g",
    #     "review_content_class": "hotels-community-tab-common-Card__card--ihfZB",
    #     "review": {
    #         "tag": "q", "class": "location-review-review-list-parts-ExpandableReview__reviewText--gOmRC"
    #     },
    #     "person": {
    #         "tag": "a", "class": "ui_header_link social-member-event-MemberEventOnObjectBlock__member--35-jC"
    #     },
    #     "name": {
    #         "tag": "h1", "class": "hotels-hotel-review-atf-info-parts-Heading__heading--2ZOcD"
    #     },
    #     "tanggal": {
    #         "tag": "div", "class": "social-member-event-MemberEventOnObjectBlock__event_type--3njyv"
    #     },
    #     "nav_class": "nav next ui_button primary"
    # },
    # {
    #     "url": domain + '/Restaurants-g14782503-Yogyakarta_Yogyakarta_Region_Java.html',
    #     "key": 'restoran-jogja',
    #     "link_selector": "#component_2 div.restaurants-list-ListCell__photoWrapper--1umtU > span > a",
    #     "more_link_selector": ".taLnk.ulBlueLinks",
    #     "review_content_class": "ppr_rup ppr_priv_location_reviews_list_resp",
    #     "review": {
    #         "tag": "p", "class": "partial_entry"
    #     },
    #     "person": {
    #         "tag": "div", "class": "info_text pointer_cursor"
    #     },
    #     "name": {
    #         "tag": "h1", "class": "ui_header h1"
    #     },
    #     "tanggal": {
    #         "tag": "span", "class": "ratingDate"
    #     },
    #     "nav_class": "nav next rndBtn ui_button primary taLnk"
    # },
    {
        "url": domain + '/RestaurantSearch-g294226-oa360-Bali.html#EATERY_LIST_CONTENTS',
        "key": 'restoran-bali',
        "link_selector": "#component_2 > div > div > span > div._1kNOY9zw > div._2Q7zqOgW > div._2kbTRHSI > div > span > a",
        "more_link_selector": ".taLnk.ulBlueLinks",
        "review_content_class": "ppr_rup ppr_priv_location_reviews_list_resp",
        "review": {
            "tag": "p", "class": "partial_entry"
        },
        "person": {
            "tag": "div", "class": "info_text pointer_cursor"
        },
        "name": {
            "tag": "h1", "class": "ui_header h1"
        },
        "tanggal": {
            "tag": "span", "class": "ratingDate"
        },
        "nav_class": "nav next rndBtn ui_button primary taLnk",
        "address": {
            "tag": "span", "class": "detail"
        }
    }
]
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--disable-browser-side-navigation")

driver = webdriver.Chrome(executable_path="./chromedriver", options=chrome_options)
wb = Workbook() 
  
for link_url in link_urls:
    # link_count = 0
    url = link_url['url']
    print(url)
    # add_sheet is used to create sheet.
    sheet = wb.add_sheet('Sheet '+link_url['key']) 
    sheet_row = 0
    sheet.write(sheet_row, 0, 'URL') 
    sheet.write(sheet_row, 1, 'Hotel / Restoran Name') 
    sheet.write(sheet_row, 2, 'Address')    
    sheet.write(sheet_row, 3, 'Reviewer') 
    sheet.write(sheet_row, 4, 'Rating') 
    sheet.write(sheet_row, 5, 'Ulasan') 
    sheet.write(sheet_row, 6, 'Tanggal Ulasan') 
    
    while True:
        try:
            response = get(url)
            soup = BeautifulSoup(response.text, 'lxml')
            data_links = soup.select(link_url['link_selector'])
            if(len(data_links) == 0):
                break
            for link in data_links:
                # link_count += 1
                detail_url = domain + link['href']
                same_page_counter = 0
                driver.get(detail_url)
                wait(driver, 10).until(EC.url_to_be(detail_url))
                # link_detail_count = 0
                prev_url = detail_url
                while True:
                    try:
                        print(driver.current_url) 
                        if(driver.current_url == prev_url):
                            same_page_counter += 1
                        prev_url = driver.current_url
                        
                        if same_page_counter == 5:
                            break
                        # link_detail_count += 1
                        detail_page_source = driver.page_source

                        # click all more link
                        try:
                            more_el = driver.find_element_by_css_selector(link_url['more_link_selector'])
                            if (more_el.is_displayed() and more_el.is_enabled()):
                                driver.execute_script("document.querySelector('"+link_url['more_link_selector']+"').click();")
                                time.sleep(1)
                        except Exception as e:
                            print('Error when clicking more link')
                            print(e)
                       
                        soup_detail = BeautifulSoup(detail_page_source, 'lxml')
                        review_sections = soup_detail.find_all("div", class_=link_url['review_content_class'])
                        for rs in review_sections:
                            review = rs.find(link_url['review']['tag'], class_=link_url['review']['class'])
                            person = rs.find(link_url['person']['tag'], class_=link_url['person']['class'])
                            name = soup_detail.find(link_url['name']['tag'], class_=link_url['name']['class'])
                            rating = rs.find('span', class_="ui_bubble_rating")
                            tanggal = rs.find(link_url['tanggal']['tag'], class_=link_url['tanggal']['class'])
                            address = soup_detail.find(link_url['address']['tag'], class_=link_url['address']['class'])
                            sheet_row += 1
                            sheet.write(sheet_row, 0, driver.current_url) 
                            sheet.write(sheet_row, 1, name.text) 
                            sheet.write(sheet_row, 2, address.text)
                            sheet.write(sheet_row, 3, person.text) 
                            sheet.write(sheet_row, 4, rating['class'][1]) 
                            sheet.write(sheet_row, 5, review.text) 
                            sheet.write(sheet_row, 6, tanggal.text) 
                        if sheet_row >= 10000:
                            break
                        #driver.execute_script("document.querySelector('.ui_button.nav.next.primary').click();")
                        wait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.ui_button.nav.next.primary'))).click()
                        time.sleep(3)
                    except Exception as err:
                        print('Error in detail ' + detail_url)
                        print(err)     
                        break
                if sheet_row >= 10000:
                  break
            next_link = soup.find('a', class_=link_url['nav_class'])
            url = domain + next_link['href']
            if sheet_row >= 10000:
                break
        except (Exception, KeyboardInterrupt) as e:
            print("Last page reached")
            print(e)
            break  
wb.save('dataset-bu-yuli.xls') 
driver.close()