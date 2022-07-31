from email.quoprimime import quote
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import pickle
import json
import time
from datetime import datetime
from datetime import timedelta
import pypyodbc as podbc

import undetected_chromedriver as uc

def checkDollar(str):
    str = str.replace("$", "")
    str = str.replace(",", "")
    return str
    
def convertFloat(str):
    if "N/A" in str:
        value = 0
    else:
        value = float(str)
    return value

if __name__ == "__main__":

    #connect to mssql server
    server = 'EC2AMAZ-CNTDMIS'
    database = 'DailyProcess'
    username = 'Administrator'
    password = 'Pazzword33$'

    try:
        connection = podbc.connect("Driver={ODBC Driver 17 for SQL Server};Server=OFFICEDESKTOP;Database=DailyProcess;Trusted_Connection=yes;") # Connection string
        cursor = connection.cursor()
        time.sleep(2)
    except:
        print (" Database connection Error")
        quit()

    selectSymbolQuery = "select distinct symbol from DailyProcess.dbo.stockcache where quotedate > DATEADD(DD, -7, getdate()) order by symbol"
    cursor.execute(selectSymbolQuery)
    symbols =cursor.fetchall()
    symbolAry = []

    for symbol in symbols:
        symbolAry.append(symbol[0])

    print(symbolAry)
    if len(symbolAry) == 0:
        print("Nothing symbols")
        connection.close()
        quit()

    #getting chrome driver
    uc_options = uc.ChromeOptions()
    uc_options.add_argument("--start-maximized")
    driver = uc.Chrome(options=uc_options)

    url = "https://www.google.com"
    driver.get(url)
    time.sleep(2)

    loginUrl = "https://client.schwab.com/Trade/OrderStatus/ViewOrderStatus.aspx"
    driver.get(loginUrl)
    time.sleep(10)

    iframe = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.TAG_NAME, "iframe#lmsSecondaryLogin"))
    )
    driver.switch_to.frame(iframe)

    inputs = WebDriverWait(driver, 15).until(
        EC.presence_of_all_elements_located((By.TAG_NAME, "input"))
    )
    inputs[0].send_keys("rickfortier")
    time.sleep(2)
    inputs[1].send_keys("VfvBjFdD2ZH2RRy")
    time.sleep(2)
    submitButton = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.TAG_NAME, "button"))
    )
    submitButton.click()
    time.sleep(2)

    # pickle.dump(driver.get_cookies(), open("cookies.pkl", "wb"))
    # cookies = pickle.load(open("cookies.pkl", "rb"))
    # for cookie in cookies:
    #     driver.add_cookie(cookie)
    
    # driver.get("https://client.schwab.com/secure/cc/research/stocks/stocks.html?path=/research/Client/Stocks/Summary&symbol=AAPL")
    # time.sleep(10)
    
    redirectedUrls = []
    totalData = []
    temp = 0
    for symbol in symbolAry:
        temp += 1
        print(symbol)
        
        stockUrl = "https://client.schwab.com/secure/cc/research/stocks/stocks.html?path=/research/Client/Stocks/Summary&symbol=" + symbol

        driver.get(stockUrl)
        time.sleep(8)

        redirectedUrl = driver.current_url
        redirectedUrls.append(redirectedUrl)

        todayOpen = ""
        highToday = ""
        lowToday = ""
        lastPrice = ""
        todayVolume = ""
        quoteTimeDate = ""

        try:
            topTabelIframe = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#wsodIFrame"))
            )
            driver.switch_to.frame(topTabelIframe)
        except:
            pass

        noInforFlag = False
        try:
            noInforMsg =  WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.message.important>p"))
            ).text
            if "No information" in noInforMsg:
                noInforFlag = True
        except:
            pass
        if noInforFlag == True:
            print(symbol, " : has no information")
            continue

        if "/stocks/stocks" in redirectedUrl:

            detailsTable = WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div#modQuoteDetails>div>div.colRight>div>table"))
            )[0]
            detailsTableTrs = detailsTable.find_elements(by=By.CSS_SELECTOR, value="tbody>tr")

            todayOpen = detailsTableTrs[0].find_element(by=By.CSS_SELECTOR, value="td").text
            
            dayRangeStr = detailsTableTrs[2].find_element(by=By.CSS_SELECTOR, value="td").text
            dayRangeStrAry = dayRangeStr.split(" - ")
            highToday = dayRangeStrAry[1]
            lowToday = dayRangeStrAry[0]

            topTable = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table.snapQuote"))
            )
            valueTr = topTable.find_elements(by=By.CSS_SELECTOR, value = "tbody>tr")[1]
            valueTds = valueTr.find_elements(by=By.CSS_SELECTOR, value = "td")

            lastPrice = valueTds[0].find_elements(by=By.CSS_SELECTOR, value = "span")[0].text

            todayVolume = valueTds[4].find_elements(by=By.CSS_SELECTOR, value = "span")[0].text

            quoteTimeDateStr = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#modFirstGlance>span"))
            ).text
            quoteTimeDateStrAry = quoteTimeDateStr.split(", ")
            if "As of close" in quoteTimeDateStrAry[0]:
                quoteTimeDate = "16:00:00" + " " + quoteTimeDateStrAry[1]
            else:
                quoteTimeDate = quoteTimeDateStrAry[0] + " " + quoteTimeDateStrAry[1]
            # print("LastPrice", lastPrice)  
            # print("TodayVolume", todayVolume)
            # print("TodayOpen", todayOpen)
            # print("HighToday", highToday)
            # print("LowToday", lowToday) 
            # print("QuoteTime", quoteTimeDate)
        
        elif "/etfs/etfs.html" in redirectedUrl:
            
            detailsTable = WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.quoteDetailsModule>div>div.colRight>div>table"))
            )[0]
            detailsTableTrs = detailsTable.find_elements(by=By.CSS_SELECTOR, value="tbody>tr")

            todayOpen = detailsTableTrs[0].find_element(by=By.CSS_SELECTOR, value="td").text
            
            dayRangeStr = detailsTableTrs[2].find_element(by=By.CSS_SELECTOR, value="td").text
            dayRangeStrAry = dayRangeStr.split(" - ")
            highToday = dayRangeStrAry[1]
            lowToday = dayRangeStrAry[0]
            

            topTable = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table.snapQuote"))
            )
            valueTr = topTable.find_elements(by=By.CSS_SELECTOR, value = "tbody>tr")[1]

            valueTds = valueTr.find_elements(by=By.CSS_SELECTOR, value = "td")
            lastPrice = valueTds[0].text

            todayVolume = valueTds[5].find_elements(by=By.CSS_SELECTOR, value = "span")[0].text

            quoteTimeDateStr = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div#modFirstGlance>div>div.contain>div.subLabel"))
            ).text
            quoteTimeDateStrAry = quoteTimeDateStr.split(", ")
            if "As of close" in quoteTimeDateStrAry[0]:
                quoteTimeDate = "16:00:00" + " " + quoteTimeDateStrAry[1]
            else:
                quoteTimeDate = quoteTimeDateStrAry[0] + " " + quoteTimeDateStrAry[1]
            
        elif "RequestType=Summary&SecurityType=INDEX" in redirectedUrl:
            
            tableRows = WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table.tblAltWhtGrayRow>tbody>tr"))
            )

            for tableRow in tableRows:
                tableTds = tableRow.find_elements(by=By.CSS_SELECTOR, value="td>span")
                tdIndex = 0
                for tableTd in tableTds:
                    tdIndex += 1
                    tdText = tableTd.text
                    if "Last Trade" == tdText:
                        lastPrice = tableTds[tdIndex].text
                        
                    elif "Open" == tdText:
                        todayOpen = tableTds[tdIndex].text

                    elif "Day High" == tdText:
                        highToday = tableTds[tdIndex].text

                    elif "Day Low" == tdText:
                        lowToday = tableTds[tdIndex].text

                    elif "Volume" == tdText:
                        todayVolume = tableTds[tdIndex].text

                    else:
                        continue
            
            quoteTimeStr = WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table.tblDrkBlueCrv>tbody>tr>td>table>tbody>tr>td>table>tbody>tr>td>div>span"))
            )[0].text.lstrip()
            quoteDateStr = WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table.tblDrkBlueCrv>tbody>tr>td>table>tbody>tr>td>table>tbody>tr>td>div>span"))
            )[1].text
            if "At Close" in quoteTimeStr:
                quoteTimeStr = "16:00:00"
            quoteTimeDate = quoteTimeStr + " " + quoteDateStr
        else:
            print(symbol, " : strange")  
            continue      
        

        # chaning values
        # if "$" in todayOpen:
        #     todayOpen = todayOpen.split("$")[1]
        # if "$" in highToday:
        #     highToday = highToday.split("$")[1]
        # if "$" in lowToday:
        #     lowToday = lowToday.split("$")[1]
        # if "$" in lastPrice:
        #     lastPrice = lastPrice.split("$")[1]
        
        todayOpen = checkDollar(todayOpen)
        highToday = checkDollar(highToday)
        lowToday = checkDollar(lowToday)
        lastPrice = checkDollar(lastPrice)
        todayVolume = checkDollar(todayVolume)
        
        todayOpen = convertFloat(todayOpen)
        highToday = convertFloat(highToday)
        lowToday = convertFloat(lowToday)
        lastPrice = convertFloat(lastPrice)
        todayVolume = convertFloat(todayVolume)

        sourceStr = "schwab"

        now = datetime.now()
        datetimeformat = "%Y-%m-%d %H:%M:%S"
        insertedtime = now.strftime(datetimeformat)

        print(quoteTimeDate)
        if quoteTimeDate[0] == " ":
            quoteTimeDate = quoteTimeDate.lstrip()
        if "As of" in quoteTimeDate:
            quoteTimeDate = quoteTimeDate.replace("As of ", "")
        pmFlag = False
        if "PM " in quoteTimeDate:
            quoteTimeDate = quoteTimeDate.replace("PM ", "")
            pmFlag = True
        if "ET " in quoteTimeDate:
            quoteTimeDate = quoteTimeDate.replace("ET ", "")
        if "ET, " in quoteTimeDate:
            quoteTimeDate = quoteTimeDate.replace("ET, ", "")
        
        
        timeStr = quoteTimeDate.split(" ")[0]
        timeLen = len(timeStr.split(":"))
        if timeLen == 2:
            timeStr += ":00"
        if pmFlag == True:
            timeAry = timeStr.split(":")
            hrTxt = int(timeAry[0]) + 12
            timeStr = str(hrTxt) + ":" + str(timeAry[1]) + ":" + str(timeAry[2])

        dateStr = quoteTimeDate.split(" ")[1]
        yStr = str(dateStr.split("/")[2])
        mStr = str(dateStr.split("/")[0])
        dStr = str(dateStr.split("/")[1])

        # date_time_str = "18/09/19 01:55:19"
        date_time_str = dStr + "/" + mStr + "/" + str(yStr[2:4]) + " " + timeStr
        date_time_obj = datetime.strptime(date_time_str, '%d/%m/%y %H:%M:%S')
        quoteTimeDate = date_time_obj.strftime(datetimeformat)
        # print("QuotedTime", quoteTimeDate)

        now = datetime.now()
        # print("now", now)
        datetimeformat = "%Y-%m-%d %H:%M:%S"
        insertedtime = now.strftime(datetimeformat)

        outDateFlag = True
        if now-timedelta(hours=24) <= date_time_obj:
            outDateFlag = False

        if outDateFlag == True:
            print(symbol, " : Out of Date")
        else:
            print(symbol, " : Insert to database")
            query = "INSERT INTO dbo.stockcache (symbol, quotedate,open_price, high_price, low_price, close_price, volume, source, insertedtime) VALUES (?,?,?,?,?,?,?,?,?)"
            parameters = symbol, quoteTimeDate,todayOpen, highToday, lowToday, lastPrice, todayVolume, sourceStr, insertedtime
                
            cursor.execute(query, parameters)
            connection.commit()
            time.sleep(2)

        matchData = {
            "symbol": symbol,
            "quotedate": quoteTimeDate,
            "open_price": todayOpen,
            "high_price": highToday,
            "low_price": lowToday,
            "close_price": lastPrice,
            "volume": todayVolume,
            "source": sourceStr,
            "insertedtime": insertedtime
        }

        totalData.append(matchData)

        # if temp == 10:
        #     break
    
    totalDataJson = json.dumps(totalData)
    with open("scrappedResult.json", "w") as outfile:
        outfile.write(totalDataJson)

    
    # json_object = json.dumps(redirectedUrls)
    # with open("RedirectedUrls.json", "w") as outfile:
    #     outfile.write(json_object)

    connection.close()
    driver.close()