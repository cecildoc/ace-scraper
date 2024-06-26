from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import subprocess
import csv
import time

# Scraper requires ChromeDriver
subprocess.Popen(
   '"C:/PATH/TO/chrome-win64/chrome.exe" --remote-debugging-port=9222', shell=True) # Update this path as needed.
# Preferred using chrome test version instead of actual chrome installation
# YOU MUST CLOSE ANY EXISTING WINDOWS OF CHROME TEST BROWSER OR IT WILL NOT WORK



# Initialize Chrome options
options = webdriver.ChromeOptions()
options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/33.0.1750.517 Safari/537.36'
options.add_argument('user-agent={0}'.format(user_agent))
chrome_driver_path  = 'C:/PATH/TO/chromedriver-win64/chromedriver.exe'  # Update this path as needed

# Initialize the WebDriver with ChromeDriver path and options
driver = webdriver.Chrome(options=options)
driver.maximize_window() # Window mustn't be minimized or out of focus else it cannot search
# List of addresses to check
addresses = ["7041 Pecos StDenver, CO 80221-7206", "202 San Juan AveLa Junta, CO 81050-1521", "6900 Eudora DrCommerce City, CO 80022-1841", "423 Main AveFlagler, CO 80815-9236", "215 S College AveFort Collins, CO 80524-2810", "1727 Main StLongmont, CO 80501-2035", "341 - 1st StreetStratton, CO 80836", "5914 S Kipling PkwyLittleton, CO 80127-5572", "1030 E 9th AveDenver, CO 80218-3322", "425 ZerexFraser, CO 80442", "1650 Highway 6 And 50Fruita, CO 81521-2042", "21761 Highway 40/287Limon, CO 80828-9315", "151 S Oak AveEaton, CO 80615-8808", "525 3rd StBerthoud, CO 80513-2679", "12035 W Alameda PkwyLakewood, CO 80228-2701", "Co-op Country995 MainNucla, CO 81424", "55 W Bromley LnBrighton, CO 80601-3025", "2563 Kipling StLakewood, CO 80215-1526", "121 W Gunnison River DrDelta, CO 81416-1814", "225 NW Frontage RdFort Collins, CO 80524-9265", "74 Hwy 119 S #1450Nederland, CO 80466", "300 Puppy Smith StAspen, CO 81611-1455", "200 E Colorado AveTelluride, CO 81435-5049", "525 N BroadwayCortez, CO 81321-2001", "2640 E 12th AveDenver, CO 80206-3208", "1011 Highway 133Carbondale, CO 81623-1874", "203 MainCollbran, CO 81624", "1075 Teller County Road 1Cripple Creek, CO 80813-5219", "151 West Mineral Ste 100Littleton, CO 80120-4510", "3333 S Tamarac DrDenver, CO 80231-4362", "2315 W 1st StCraig, CO 81625-3612", "561 Lone Pine DrEstes Park, CO 80517-9407", "1 Enterprise DrWestcliffe, CO 81252-8557", "16555 State Highway 136La Jara, CO 81140-9473", "205 W 8th StPalisade, CO 81526-8662", "121 E. Bridge StreetHotchkiss, CO 81419", "5195 County Road 64Bailey, CO 80421", "1555 S 1st StreetBennett, CO 80102-8638", "1719 Sheridan BlvdEdgewater, CO 80214-1323", "401 E Market StMeeker, CO 81641-9621", "1225 N Circle DrColorado Springs, CO 80909-3136", "8258 Colorado BlvdFirestone, CO 80504-6800", "2626 A 11th AveGreeley, CO 80631-8441", "700 E 8th AveYuma, CO 80759-2136", "333 Dexter StWray, CO 80758-1625", "627 W Agate AveGranby, CO 80446", "8080 S Holly StCentennial, CO 80122-4001", "2155 Curve PlzSteamboat Springs, CO 80487-4967", "22 S Townsend AveMontrose, CO 81401-3954", "525 Navajo Trail DrPagosa Springs, CO 81147-8867", "4201 Centennial BlvdColorado Springs, CO 80907-3770", "1245 Main StWindsor, CO 80550-5918", "569 32 Road #4Grand Junction, CO 81504-7053", "2300 N Wahsatch AveColorado Springs, CO 80907-6941", "201 S Rollie AveFort Lupton, CO 80621-1508", "15530 W 64th Ave Ste GArvada, CO 80007-6855", "4509 S BroadwayEnglewood, CO 80113-5723", "5910 S UniversityGreenwood Village, CO 80121-2882", "1830 W Uintah StColorado Springs, CO 80904-5902", "2800 W 104th AveDenver, CO 80234-3539", "3540 W 10th StGreeley, CO 80634-1824", "9579 S University BlvdHighlands Ranch, CO 80126-8106", "2500 S Colorado BlvdDenver, CO 80222-5909", "7777 W Jewell Ave #1BLakewood, CO 80232-6843", "2730 S Academy BlvdColorado Springs, CO 80916-2806", "157 Craft DrAlamosa, CO 81101-2273", "7100 E Colfax AveDenver, CO 80220-1806", "3090 E Main StCanon City, CO 81212-2731", "186 Mount Evans BlvdPine, CO 80470-7899", "3758 Osage StDenver, CO 80211-2657", "9979 Wadsworth PkwyWestminster, CO 80021-4241", "5944 Stetson Hills Boulevard Ste 180Colorado Springs, CO 80923-3510", "269 E 29th StLoveland, CO 80538-2721", "2602 S Glen AveGlenwood Springs, CO 81601-4414", "24 Main Street, PO Box 428Red Feather Lakes, CO 80545-0428", "15181 East 104th AveCommerce City, CO 80022", "1920 Grand AveNorwood, CO 81423", "3640 Austin Bluffs PkwyColorado Springs, CO 80918-6680", "609 Highway 64 WRangely, CO 81648-9600", "3851 E 120th AveThornton, CO 80233-1660", "1920 County Road 31Florissant, CO 80816-7064", "1001 E Harmony Rd Unit BFort Collins, CO 80525-8890", "13355 Voyager PkwyColorado Springs, CO 80921-7656", "1298 Martin AveBurlington, CO 80807-1755", "111 N 6th StHayden, CO 81639", "1417 S Holly StDenver, CO 80222-3509", "1335 Park StCastle Rock, CO 80109-1596", "2525 Arapahoe AveBoulder, CO 80302-6720", "6850 Highway 165Colorado City, CO 81019", "10000 Ralston RdArvada, CO 80004-4959", "170 E 2nd SCheyenne Wells, CO 80810", "Salida Ace Hardware", "220 Cooley Mesa RdGypsum, CO 81637-9707", "299 Highway 285 PO Box 94Fairplay, CO 80440-0094", "607 6th StreetCrested Butte, CO 81224", "1040 15th StDenver, CO 80202-2345", "5005 W 72nd AveWestminster, CO 80030-5199", "820 W Tomichi AveGunnison, CO 81230-3438", "4830 Chambers RdDenver, CO 80239-5152", "1722 9th StGreeley, CO 80631-3134", "19850 Cockriel DriveParker, CO 80134", "8214 South KiplingLittleton, CO 80127", "17190 E Iliff AveAurora, CO 80013-1522", "200 SW 2nd StCedaredge, CO 81413-3837", "2401 North Ave, Unit 20Grand Junction, CO 81501-6408", "1343 E South Boulder RdLouisville, CO 80027-2301", "9 S Parish AveJohnstown, CO 80534-9099", "1153 Bergen Pkwy Space OEvergreen, CO 80439-9501", "155 E Main St Box 566Aguilar, CO 81020-5067", "7420 S Gartrell Rd Unit BAurora, CO 80016-4330", "8 Town PlzDurango, CO 81301-5104", "29785 Hwy 24 NBuena Vista, CO 81211-9601", "1010 N Market PlzPueblo West, CO 81007-1530", "1402 Mill StBrush, CO 80723-1809", "17720 S Golden RdGolden, CO 80401-6012", "CONNECTICUT", "220 Albany TpkeCanton, CT 06019-2520", "2687 Main StGlastonbury, CT 06033-2023", "385 Main StRidgefield, CT 06877-4601", "415 Park StHartford, CT 06106-1534", "1145 Main StBranford, CT 06405-3717", "25 E High St # 3-4-5East Hampton, CT 06424-1001", "81 Main StHebron, CT 06248-1541", "132 Main StOld Saybrook, CT 06475-2373", "690 Main St SWoodbury, CT 06798-3701", "146 W Town StNorwich, CT 06360-2132", "480 S Main StMiddletown, CT 06457-4215", "231 Maple AveCheshire, CT 06410-2506", "71-73 Windsor AveVernon, CT 06066-2450", "296 Broad StWindsor, CT 06095-2939", "114 E Main StClinton, CT 06413-2112", "945 Cromwell AveRocky Hill, CT 06067-3008", "982 Farmington AveWest Hartford, CT 06107-4100", "300 Oxford RdOxford, CT 06478-1656", "212 Old Hartford RdColchester, CT 06415-2743", "18 Kent Green BlvdKent, CT 06757-1519", "268 Sound Beach AveOld Greenwich, CT 06870-1626", "9 Boston StGuilford, CT 06437-2855", "595 Straits TpkeWatertown, CT 06795-3356", "40 Boston Post RdWaterford, CT 06385-2424", "95 Wolcott RdWolcott, CT 06716-2688", "59 Church StNorth Canaan, CT 06018-2482", "211 Greenwood AveBethel, CT 06801-2124", "348 Bantam RdLitchfield, CT 06759-3318", "224 Watertown RdThomaston, CT 06787-1920", "724 Portland Cobalt RdPortland, CT 06480-1760"]
# Open the store locator page
driver.get("https://www.acehardware.com/store-locator")

# Define a function to extract store information
def extract_store_info(driver):
    placeholder_email = "N/A"
    placeholder_number = "N/A"
    placeholder_address = "N/A"
    placeholder_manager = "N/A"
    placeholder_owner = "N/A"

    store_info = {
    "email": placeholder_email,
    "phone": placeholder_number,
    "address": placeholder_address,
    "manager": placeholder_manager,
    "owner": placeholder_owner
    }
      # Check for presence of elements using XPATH and strings  
      # Wait until the store info is loaded and extract data
    try:
        element = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="page-content"]/div[4]/div/div[1]/div[1]/div/div[2]/div[1]/div[3]/span'))  
        )
        if element.text == "Phone":
            store_phone = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="page-content"]/div[4]/div/div[1]/div[1]/div/div[2]/div[1]/div[3]/a' ))
            ).text
            store_info["phone"] = store_phone
    except TimeoutException:
        pass # Uses placeholder

    try:
        # Wait until the store info is loaded and extract data
        element = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="page-content"]/div[4]/div/div[1]/div[1]/div/div[2]/div[2]/div[1]/span'))  
        )
        if element.text == "Email":
            store_email = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="page-content"]/div[4]/div/div[1]/div[1]/div/div[2]/div[2]/div[1]/a' ))
            ).text
            store_info["email"] = store_email
        else:
            store_email = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="page-content"]/div[4]/div/div[1]/div[1]/div/div[2]/div[2]/div[2]/a' ))
            ).text
            store_info["email"] = store_email

    except TimeoutException:
        pass # Uses placeholder
   # Pull partial address for manual verification if needed. Not necessary. Can also add city, zip and state
   # by manually pulling appropriate XPATH from website
    try:    
        store_address = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="page-content"]/div[4]/div/div[1]/div[1]/div/div[2]/div[1]/div[2]/div[1]/span[1]'))
        ).text
        store_info["address"] = store_address
    except TimeoutException:
        pass # Uses placeholder 

    try:   
        element = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="page-content"]/div[4]/div/div[1]/div[1]/div/div[3]/div/span'))
        )
        if element.text == "Managers":
            store_manager = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="page-content"]/div[4]/div/div[1]/div[1]/div/div[3]/div/div'))
            ).text
            store_info["manager"] = store_manager
        elif element.text == "Owned by":
            store_owner = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="page-content"]/div[4]/div/div[1]/div[1]/div/div[3]/div[1]/div'))
            ).text                                         
            store_info["owner"] = store_owner 
    except TimeoutException:
        pass  # Use placeholder 
    
    try:
        store_manager = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="page-content"]/div[4]/div/div[1]/div[1]/div/div[3]/div[2]/div[2]'))
        ).text
        store_info["manager"] = store_manager
    except TimeoutException:
        pass  # Use placeholder  

    return store_info

# Creating CSV. Change arguements appropriately. Currently will always create a new blank document
with open('store_info.csv', mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["Email", "Phone Number", "Address", "Manager", "Owner"])

    for address in addresses:
        try:
            # Find the search input field and enter the address
            
            search_field = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "mapsearch"))
            )
            search_field.clear()
            search_field.send_keys(address)
            search_field.send_keys(Keys.RETURN)

            time.sleep(1) # Allows elements to load, wasn't exprerienced enough with selenium to use
            # the explicit function included with the package
            # Match result with address
            ''' element = WebDriverWait(driver,3).until(
                EC.element_to_be_clickable((By.XPATH,'//*[@id="location-list"]/div/div[1]/div[1]/div[2]'))
            ).text
            element2 = WebDriverWait(driver,3).until(
                EC.element_to_be_clickable((By.XPATH,'//*[@id="location-list"]/div/div[1]/div[1]/div[3]'))
            ).text
            elementsearch = element + element2
            
            if elementsearch == address:
                elementsearch = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="location-list"]/div/div[1]'))
            )
                elementsearch.click()
            '''    
            first_result = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="location-list"]/div/div[1]'))
            )
            first_result.click()
            
            # Click on "Store Details" button/link
            store_details_button = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="mz-map-canvas"]/div/div[3]/div[1]/div[2]/div/div[4]/div/div/div[2]/div[1]/a/button'))
            )
            store_details_button.click()
    
            # Extract store information from the redirected page
            store_info = extract_store_info(driver)
            if store_info:
                writer.writerow([store_info['email'], store_info['phone'], store_info['address'], store_info['manager'], store_info['owner']])

            # Go back to the store locator page
            driver.back()
            # Allows page to load, similarly to before
            time.sleep(1)
        except Exception as e:
            print(f"An error occurred with address {address}: {e}")

# Close the WebDriver although because it's a subprocess make sure to close out the window before starting again. 
# Also check for any remaining processes.
driver.quit()