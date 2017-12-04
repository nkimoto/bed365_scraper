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
import openpyxl as px
import datetime

def set_up(url):
    options = Options()
#    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    prefs = {"profile.default_content_setting_values.notifications" : 2}
    options.add_experimental_option("prefs",prefs)
    # install chromedriver if not found and start chrome
# chrome_options=options
    driver = webdriver.Chrome(chrome_options=options)
    driver.get(url)
    print('Searching in ' + url)
    time.sleep(waittime)
    driver.find_elements_by_tag_name("a")[18].click()
    time.sleep(waittime)
    return driver

def scraping(driver, match_num, roop_count):
    """
    SCRAPING CODE
    RETURN : DICT LIST
    """
    count = 1
    All_Res_Dict = {}
    events_dict = {}
    while count <= roop_count:
        print('Now roop count is ' + str(count))
        if e_wait(driver, 'div.li-InPlayClassificationButton_Header>\
                div.li-InPlayClassificationButton_HeaderLabel').text == "Soccer":
            print('Soccer term found!')
            matchnum = match_num if match_num else len(driver.find_elements_by_css_selector(
            "div.li-InPlayClassificationButton_Header+\
                    div.li-InPlayClassification_League>\
                    div.li-InPlayLeague"))
            print(match_num, matchnum)
            for n in range(matchnum):
            # Specify target matchs
                e_wait(driver, "div.li-InPlayClassificationButton_Header+\
                    div.li-InPlayClassification_League>\
                    div.li-InPlayLeague")
                i = driver.find_elements_by_css_selector(
                    "div.li-InPlayClassificationButton_Header+\
                div.li-InPlayClassification_League>\
                div.li-InPlayLeague")[n].find_element_by_css_selector("div.li-InPlayEventHeader")
                j = "div.li-InPlayClassificationButton_Header+\
                div.li-InPlayClassification_League>\
                div.li-InPlayLeague"
                match_name = str(i.text.split('\n')[0])
                if match_name not in All_Res_Dict:
                    print('====================' + match_name + ' start scraping!====================')
                    All_Res_Dict[match_name] = []
                time.sleep(10)
                i.click()
                res_dict = {}
                try:
                    time_reg = e_wait(driver, "div.ipe-SoccerHeaderLayout_ExtraData")
                    print(time_reg.text)
                    res_dict['Time'] = time_reg.text
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
                    if s.text == "1st Goal":
                        for m in h.find_elements_by_css_selector(
                                "div.gl-Participant"):
                            print("1st Goal\n" + m.text)
                            try:
                                res_dict['1st Goal'][m.text.split('\n')[0]] = m.text.split('\n')[1]
                            except:
                                print('1st Goal not showing now.')
                    if s.text == "2nd Goal":
                        for m in h.find_elements_by_css_selector(
                                "div.gl-Participant"):
                            print("2nd Goal\n" + m.text)
                            try:
                                res_dict['2nd Goal'][m.text.split('\n')[0]] = m.text.split('\n')[1]
                            except:
                                print('2nd Goal not showing now.')
                    if s.text == "Alternative Match Goals":
                        val = h.find_element_by_css_selector('div.gl-MarketLabel').text
                        over = h.find_elements_by_css_selector('div.gl-MarketValues')[0].text
                        under = h.find_elements_by_css_selector('div.gl-MarketValues')[1].text
                        altlist = [[m , n, l] for m, n, l in zip(val.split('\n'), over.split('\n')[1:], under.split('\n')[1:])]
                        print("Alternative Match Goals")
                        print(altlist)
                        res_dict['Alternative Match Goals'] = [val, over.lstrip('Over\n'), under.lstrip('Under\n')]
                # Wheel nums
                e_wait(driver, "div.ml1-StatWheel_Team1Text")
                for i, num in zip(('Attacks', 'Dangerous Attacks', 'Possession'), range(3)):
                    try:
                        element1 = driver.find_elements_by_css_selector("div.ml1-StatWheel_Team1Text")[num].text
                        element2 = driver.find_elements_by_css_selector("div.ml1-StatWheel_Team2Text")[num].text
                        print("Possession " + Possession1 + ' : ' + Possession2)
                        res_dict['Possession'] = (Possession1, Possession2)
                    except:
                        print('There is no ' + str(i))
                # Bar nums
                for i, num in zip(('On Target', 'Off Target'), range(2)):
                    try:    
                        element1 = driver.find_elements_by_css_selector("div.ml1-SoccerStatsBar_MiniBarValue-1")[num].text
                        element2 = driver.find_elements_by_css_selector("div.ml1-SoccerStatsBar_MiniBarValue-2")[num].text
                        print(i + element1 + ':' + element2)
                        res_dict[i] = (element1, element2)
                    except:
                        print('There is no ' + str(i))
                All_Res_Dict[match_name].append(res_dict)
                print(All_Res_Dict)
                print('\nNew appending dict')
                print(res_dict)
                if count == roop_count:
                    try:
                        events = e_wait(driver, 'div.ipe-SummaryNativeScroller_Content')
                        events_list = events.split('\n')
                        events_dict['match_name'] = events_list
                    except:
                        print("events list can't get.")
                        pass
                driver.back()
                print("back now")
            count += 1
        else:
            print('Waiting...')
            time.sleep(60)
    driver.quit()
    return All_Res_Dict, events_dict

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

def ExcelWriter(All_Res_Dict, events_dict):
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
        info_Row = 3
        ActiveWS = Target.get_sheet_by_name('Match_name')
        SummarySheet = Target.copy_worksheet(ActiveWS)
        SummarySheet.title = match[:31]
        merge(SummarySheet)
        # complement header
        name1_list = ('B', 'E', 'H', 'K', 'T', 'V', 'X')
        name2_list = ('D', 'G', 'J', 'M', 'U', 'W', 'Y')
        for col1, col2 in zip(name1_list, name2_list):
            SummarySheet[col1 + '2'].value = match.split(' v ')[0]
            SummarySheet[col2 + '2'].value = match.split(' v ')[1]
        for info_dict in All_Res_Dict[match]:
            SummarySheet['A' + str(Row)].value = str(info_dict.get('Time', 'NoData'))
            for col in ['B', 'C', 'D']:
                try:
                    for key in info_dict.get('Fulltime Result'):
                        SummarySheet[col + str(Row)].value = str(info_dict['Fulltime Result'][key])
                except:
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col in ['E', 'F', 'G']:
                try:
                    for key in info_dict.get('Halftime Result'):
                        SummarySheet[col + str(Row)].value = str(info_dict['Halftime Result'][key])
                except:
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col in ['H', 'I', 'J']:
                try:
                    for key in info_dict.get('1st Goal'):
                        SummarySheet[col + str(Row)].value = str(info_dict['1st Goal'][key])
                except:
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col in ['K', 'L', 'M']:
                try:
                    for key in info_dict.get('2nd Goal'):
                        SummarySheet[col + str(Row)].value = str(info_dict['2nd Goal'][key])
                except:
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col, num in zip(['N', 'O', 'P'], range(3)):
                try:
                    SummarySheet[col + str(Row)].value = str(info_dict['Alternative Match Goals'][num])
                except:
                    SummarySheet[col + str(Row)].value = 'NoData'
            for i, j in zip([('T', 'U'), ('V', 'W'), ('X', 'Y'), ('Z', 'AA'), ('AB', 'AC')],(
                'Attacks', 'Dangerous Attacks', 'Possession', 'Off Target', 'On Target')):
                for col, num in zip(i, range(2)):
                    try:
                        SummarySheet[col + str(Row)].value = str(info_dict[j][num])
                    except:
                        SummarySheet[col + str(Row)].value = 'NoData'


            Row += 1
        try:
            for event in events_dict[match_name]:
                SummarySheet['AE' + str(info_Row)].value = event
                info_Row += 1
        except:
            SummarySheet['AE' + str(info_Row)].value = 'NoData'
    # Delete tmp Sheet
    ActiveWS = Target.get_sheet_by_name('Match_name')
    Target.remove_sheet(ActiveWS)
    # Save
    Target.save('Report/Res_' + NowTime + '.xlsx')


def e_wait(driver, element, max_time = 60):
    return_element = None
    try:
        return_element = WebDriverWait(driver, max_time).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, element))
                )
    except:
        print(element + " can't get!")
        pass
    return return_element


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
    All_Res_Dict, events_dict = scraping(driver, args.mn, args.r_count)
    ExcelWriter(All_Res_Dict, events_dict)
    print('END!!')


if __name__ == '__main__':
    NowTime = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    args = Argument_Parser()
    waittime = args.wait
    main()
