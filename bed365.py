#! env python
# -*- coding: utf-8 -*

from argparse import ArgumentParser
import re
import os
import sys
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
import csv

def set_up(url):
    options = Options()
    options.binary_location = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
#    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    capabilities = webdriver.DesiredCapabilities.CHROME.copy()
    capabilities['platform'] = "WINDOWS"
    capabilities['version'] = "7"
    # install chromedriver if not found and start chrome
    driver = webdriver.Chrome(chrome_options=options, desired_capabilities=capabilities)
    driver.get(url)
    print('Searching in ' + url)
    time.sleep(waittime)
    driver.save_screenshot('home.png')
    driver.find_elements_by_tag_name("a")[18].click()
    time.sleep(waittime)
    return driver

def scraping(driver, match_num, roop_count):
    """
    SCRAPING CODE
    RETURN : DICT LIST
    """
    count = 0
    All_Res_Dict = {}
    while count < roop_count:
        if driver.find_element_by_css_selector(
                "div.li-InPlayClassificationButton_Header>\
            div.li-InPlayClassificationButton_HeaderLabel"
        ).text == "Soccer":
            print('Soccer term found!')
            print(driver.find_elements_by_css_selector(
                    "div.li-InPlayClassificationButton_Header+\
                div.li-InPlayClassification_League>\
                div.li-InPlayLeague"))
            matchnum = match_num if match_num else len(driver.find_elements_by_css_selector(
                    "div.li-InPlayClassificationButton_Header+\
                div.li-InPlayClassification_League>\
                div.li-InPlayLeague"))
            for n in range(matchnum):
            # Specify target matchs
                print(n)
                try:
                    i = driver.find_elements_by_css_selector(
                        "div.li-InPlayClassificationButton_Header+\
                    div.li-InPlayClassification_League>\
                    div.li-InPlayLeague")[n].find_element_by_css_selector("div.li-InPlayEventHeader")
                    match_name = str(i.text.split('\n')[0])
                except Exception:
           
                    print("Match Name can't get")
                    continue
                if match_name not in All_Res_Dict:
                    print('====================' + match_name + ' start scraping!====================')
                    All_Res_Dict[match_name] = []
                i.click()
                time.sleep(waittime)
                res_dict = {}
                driver.save_screenshot('game{}.png'.format(str(n)))
                try:
                    time_reg = driver.find_element_by_css_selector(
                        "div.ipe-SoccerHeaderLayout>div.ipe\
                        -SoccerHeaderLayout_ExtraData").text
                    print(time_reg)
                    res_dict['Time'] = time_reg
                except:
                    print("Time can't get")
                for h in driver.find_elements_by_css_selector(
                        "div.gl-MarketGroup"):
                    try:
                        s = h.find_element_by_css_selector(
                            "div.gl-MarketGroupButton")
                    except:
                        print("gl-MarketGroupButton not found")
                        continue
                    print('s : ' + s.text)
                    if s.text == "Fulltime Result":
                        res_dict['Fulltime Result'] = {}
                        for m in h.find_elements_by_css_selector(
                                "div.gl-Participant"):
                            print("Fulltime Result\n :".format(n) + m.text)
                            print(m.text.split('\n'))
                            try:
                                res_dict['Fulltime Result'][m.text.split('\n')[0]] = m.text.split('\n')[1]
                            except:
                                print('Fulltime Result not showing now.')
                    if s.text == "Half Time Result":
                        for m in h.find_elements_by_css_selector(
                                "div.gl-Participant"):
                            print("Halftime Result\n" + m.text)
                            try:
                                res_dict['Half Time Result'][m.text.split('\n')[0]] = m.text.split('\n')[1]
                            except:
                                print('Half Time Result not showing now.')
                    if s.text == "2nd Goal":
                        for m in h.find_elements_by_css_selector(
                                "div.gl-Participant"):
                            print("2nd Goal\n" + m.text)
                            try:
                                res_dict['2nd Goal'][m.text.split('\n')[0]] = m.text.split('\n')[1]
                            except:
                                print('Fulltime Result not showing now.')
                    if s.text == "Alternative Match Goals":
                        val = h.find_element_by_css_selector('div.gl-MarketLabel').text
                        over = h.find_elements_by_css_selector('div.gl-MarketValues')[0].text
                        under = h.find_elements_by_css_selector('div.gl-MarketValues')[1].text
                        altdict = {m : (n, l) for m, n, l in zip(val.split('\n'), over.split('\n'), under.split('\n'))}
                        print("Alternative Match Goals")
                        print(altdict)
                        res_dict['Alternative Match Goals'] = altdict
                try:
                    allstats = driver.find_element_by_css_selector("div.ml1-AllStats").text
                    print(allstats)
                    res_dict['allstats'] = allstats
                    Possesion = driver.find_elements_by_css_selector("div.ml1-StatWheel")[2].text
                    print("Possession " + Posession)
                    res_dict['possesion'] = Posession
                except:
                    pass
                All_Res_Dict[match_name].append(res_dict)
                print(All_Res_Dict)
                print('\nNew appending dict')
                print(res_dict)
                driver.back()
                print("back now")
                time.sleep(30)
                print("take picture{}".format(n))
                driver.save_screenshot('game{}.png'.format(str(n + 100)))
            roop_count += 1
        else:
            print('Waiting...')
            time.sleep(20)
    driver.quit()
    return All_Res_Dict

    """"
    while True:
        html = driver.page_source.encode('utf-8')
        try:
            soup = BeautifulSoup(html, "lxml")
        except:
            soup = BeautifulSoup(html, "html5lib")
    Tagdd = Soup.find_all('dd')
        TagRemovePatt = r'<.*?>'
        TagRemover = re.compile(TagRemovePatt)
        PaperID = [TagRemover.sub('', str(j)) for j in Tagdd]
        PaperURLs = [PubHome + '/' + i for i in PaperID]
        PaperURLList.extend(PaperURLs)
    return(PaperURLList)
    driver.quit()url = 'https://www.bet365.com/#/IP/'
    driver = webdriver.PhantomJS()
    driver.get(url)
    PaperURLList = [
    while True:
        html = driver.page_source.encode('utf-8')
        try:
            soup = BeautifulSoup(html, "lxml")
        except:
            soup = BeautifulSoup(html, "html5lib")
        Tagdd = Soup.find_all('dd')
        TagRemovePatt = r'<.*?>'
    """

def ExcelWriter(All_Res_Dict):
    def merge(ws):
        ws.merge_cells('A%s:A%s' % (1, 2), 'Merged Range')
        ws.merge_cells('B%s:D%s' % (1, 1), 'Merged Range')
        ws.merge_cells('E%s:G%s' % (1, 1), 'Merged Range')
        ws.merge_cells('H%s:J%s' % (1, 1), 'Merged Range')
        ws.merge_cells('K%s:M%s' % (2, 2), 'Merged Range')
        ws.merge_cells('N%s:P%s' % (1, 1), 'Merged Range')
        ws.merge_cells('Q%s:R%s' % (1, 1), 'Merged Range')
        ws.merge_cells('S%s:T%s' % (1, 1), 'Merged Range')
        ws.merge_cells('U%s:V%s' % (1, 1), 'Merged Range')
        ws.merge_cells('X%s:X%s' % (1, 2), 'Merged Range')
        ws.merge_cells('Y%s:AF%s' % (1, 2), 'Merged Range')

    def copy_range(fromsheet, tosheet, source_range, target_range):
        for i, j in zip(fromsheet[source_range],
                        tosheet[target_range]):
            for cell, new_cell in zip(i, j):
                new_cell._style = copy(cell._style)
                new_cell.value = cell.value

    try:
        os.mkdir('Report')
    except Exception:
        print("Result is included in Report directory.")
        pass
    Template = './Template/Template.xlsx'
    Target = px.load_workbook(Template)

    print("Excel Writer Start!")
    # Summary Sheet
    for match in All_Res_Dict.keys():
        Row = 3
        ActiveWS = Target.get_sheet_by_name('Match_name')
        SummarySheet = Target.copy_worksheet(ActiveWS)
        SummarySheet.title = match[:31]
        merge(SummarySheet)
        for info_dict in match:
            SummarySheet['A' + Row].value = str(info_dict.get('Time', 'NoData'))
            for col in zip(['B', 'C', 'D']):
                try:
                    for key in info_dict.get('Fulltime Result'):
                        SummarySheet[col + Row].value = str(info_dict['Fulltime Result'][key])
                except:
                    SummarySheet[col + Row].value = 'NoData'
                    break
            for col in zip(['E', 'F', 'G']):
                try:
                    for key in info_dict.get('Halftime Result'):
                        SummarySheet[col + Row].value = str(info_dict['Halftime Result'][key])
                except:
                    SummarySheet[col + Row].value = 'NoData'
                    break
            for col in zip(['H', 'I', 'J']):
                try:
                    for key in info_dict.get('2nd Goal'):
                        SummarySheet[col + Row].value = str(info_dict['2nd Goal'][key])
                except:
                    SummarySheet[col + Row].value = 'NoData'
                    break
            Row += 1

    # Delete tmp Sheet
    ActiveWS = Target.get_sheet_by_name('Match_name')
    Target.remove_sheet(ActiveWS)
    # Save
    Target.save('Report/Result.xlsx')


def write_list_to_csv(res_file, fieldnamelist, dictlist):
    with open(res_file, 'w') as wf:
        writter = csv.DictWritter(wf, fieldnames = fieldnamelist)
        writer.writeheader()
        for d in dictlist:
            writer.writerow(d)

def Argument_Parser():
    """
    Parse Arguments
    """
    parser = ArgumentParser(
            description='bed365_scraper.py')
    parser.add_argument('-t', '--waittime',
            action='store',
            dest='wait',
            default=15,
            type=int,
            help="specify waiting time with option '-t'",
            )
    parser.add_argument('-mn', '--matchnum',
            action='store',
            dest='mn',
            default=3,
            type=int,
            help="specify match number with option '-mn'")
    parser.add_argument('-rc', '--roop_count',
            action='store',
            dest='r_count',
            default=10,
            type=int,
            help="specify roop count with option '-rc'")

 
    args = parser.parse_args()
    return args

def main():
    print('bed365_scraper START!!')
    driver = set_up('https://www.bet365.com/home/')
    All_Res_Dict = scraping(driver, args.mn, args.r_count)
    ExcelWriter(All_Res_Dict)
    print('END!!')



if __name__ == '__main__':
    args = Argument_Parser()
    waittime = args.wait
    main()
