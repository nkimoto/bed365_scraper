#! env python
# coding=shift_jis

from argparse import ArgumentParser
import re
import os
import sys
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import csv
import openpyxl as px
import datetime

def set_up(url):
    options = Options()
#    options.add_argument('--headless')
    options.binary_location = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    options.add_argument('--disable-gpu')
    prefs = {"profile.default_content_setting_values.notifications" : 2}
    options.add_experimental_option("prefs",prefs)
#    capabilities['platform'] = "WINDOWS"
#    capabilities['version'] = "7"
    # install chromedriver if not found and start chrome
# chrome_options=options
    driver = webdriver.Chrome(chrome_options=options)
    driver.get(url)
    print('Searching in ' + url)
    time.sleep(waittime)
    driver.find_elements_by_tag_name("a")[18].click()
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
        # initiate
        print('Now roop count is ' + str(count))
        first_term = x_wait(driver, "//div[contains(@class, 'li-InPlayClassificationButton_HeaderLabel')]")
        while first_term is None:
            first_term = e_wait(driver, 'div.li-InPlayClassificationButton_Header>' + \
            'div.li-InPlayClassificationButton_HeaderLabel')
            print("First term can't get in the page. Waiting...")
            driver.rehash()
            time.sleep(30)
        else:
            pass
        print('Now first term : ' + first_term.text)
        if first_term.text == "Soccer":
            try:
                print('Soccer term found!')
                soccer = e_wait(driver, 'div.li-InPlayClassification_League')
                leagues = soccer.find_elements_by_css_selector('div.li-InPlayLeague')
                for i in range(len(leagues)):
                    league_num = i + 1
                    game_x_list = []
                    if len(leagues) == 1:
                        soccer = e_wait(driver, 'div.li-InPlayClassification_League')
                        league = x_wait(soccer, "//div[@class='li-InPlayLeague ']")
                    else:
                        soccer = e_wait(driver, 'div.li-InPlayClassification_League')
                        soccer.find_elements_by_xpath("//div[@class='li-InPlayLeague ' and position()=2]")
                        league = x_wait(soccer, "//div[@class='li-InPlayLeague ' and position()=%s]"%league_num)
                    try:
                        league_matchs = league.find_elements_by_css_selector('div.li-InPlayEvent')       
                    except (NoSuchElementException, TimeoutException):
                        continue
                    print(league_matchs)
                    for j in range(len(league_matchs)):
#                        game1 = "//div[contains(text(), 'Soccer') and @class='li-InPlayClassificationButton_Header ']/following-sibling::div//div[@class='li-InPlayLeague ' and position()=%s]/div[2]/div/div[@class='li-InPlayEventHeader ']/div[1]/div[1]"%league_num
                        if len(league_matchs) == 1:
                        # div[contains(@class, 'li-InPlayClassificationButton_HeaderLabel')]
                            game_x = "//div[contains(@class, 'li-InPlayClassification ') and contains(., 'Soccer')]//div[@class='li-InPlayLeague ' and position()=%s]/div[2]/div/div[@class='li-InPlayEventHeader ']/div[1]/div[1]"%league_num
#                            game_x_list.append(game_x)
                        elif len(league_matchs) > 1:
                            match_num = j + 1
                            game_x = "//div[contains(@class, 'li-InPlayClassification ') and contains(., 'Soccer')]//div[@class='li-InPlayLeague ' and position()=%s]/div[2]/div[@class='li-InPlayEvent ' and position()=%s]/div[@class='li-InPlayEventHeader ']/div[1]/div[1]" %(league_num, match_num)
#                                game_x_list.append(game_x)
#                        for game_x in game_x_list:
                        print(game_x)
                        c = 0
#                        game = WebDriverWait(league, 50).until(
#                                EC.element_to_be_clickable((By.XPATH, "//div[1]/div/div[2]/div[1]/div/div[2]/div[2]/div/div/div[5]/div/div[3]/div[2]/div[7]/div[2]/div[1]/div/div[1]/div[1]"
# ))))  )
                        while c < 5:
                            try:
                                print('roop count : ' + str(count))
                                print('league num : ' + str(league_num))
                                game = c_wait(driver, game_x)
                                match_name = str(game.text.split('\n')[0])
                                game.click()
                                break
                            except Exception: # match can't get : not exist or not visible
                                print(game_x)
                                print(' does not clickable...\nscroll page!')
                                driver.execute_script('window.scrollBy(0,300);')
                                c += 1
                        else: # not exist
                            driver.execute_script('window.scrollTo(0,0);')
                            print('!!!!!!!!!!')
                            break
                        if match_name not in All_Res_Dict:
                            try:
                                print('====================' + match_name + ' start scraping!====================')
                            except UnicodeError:
                                print('====================start scraping!===========================')
                            All_Res_Dict[match_name] = []
                        res_dict = {}
                        try:
                            time_reg  = e_wait(driver, "div.ipe-SoccerHeaderLayout_ExtraData")
                            print(time_reg.text)
                            res_dict['Time'] = time_reg.text
                        except AttributeError:
                            print("Time can't get")
                        for h in driver.find_elements_by_css_selector(
                                "div.gl-MarketGroup"):
                            try:
                                s = e_wait(h, "div.gl-MarketGroupButton")
                                print('s : ' + s.text)
                            except NoSuchElementException:
                                print("gl-MarketGroupButton not found")
                                continue
                            if s.text == "First Half Goals":
                                res_dict['First Half Goals'] = {}
                                val = e_wait(h, 'div.gl-MarketLabel').text
                                over = h.find_elements_by_css_selector('div.gl-MarketValues')[0].text
                                under = h.find_elements_by_css_selector('div.gl-MarketValues')[1].text
                                altlist = [[m , n, l] for m, n, l in zip(val.split('\n'), over.split('\n')[1:], under.split('\n')[1:])]
                                print("First Half Goals")
                                print(altlist)
                                res_dict['First Half Goals'] = [val, over.lstrip('Over\n'), under.lstrip('Under\n')]
                            if s.text == "Fulltime Result":
                                res_dict['Fulltime Result'] = {}
                                e_wait(h, "div.gl-Participant")
                                for m in h.find_elements_by_css_selector(
                                        "div.gl-Participant"):
                                    try:
                                        res_dict['Fulltime Result'][m.text.split('\n')[0]] = m.text.split('\n')[1]
                                    except AttributeError:
                                        print('Fulltime Result not showing now.')
                            if s.text == "Half Time Result":
                                res_dict['Half Time Result'] = {}
                                for m in h.find_elements_by_css_selector(
                                        "div.gl-Participant"):
                                    try:
                                        res_dict['Half Time Result'][m.text.split('\n')[0]] = m.text.split('\n')[1]
                                    except AttributeError:
                                        print('Half Time Result not showing now.')
                            if s.text == "1st Goal":
                                res_dict['1st Goal'] = {}
                                for m in h.find_elements_by_css_selector(
                                        "div.gl-Participant"):
                                    try:
                                        print("1st Goal\n" + m.text)
                                    except UnicodeEncodeError:
                                        pass
                                    try:
                                        res_dict['1st Goal'][m.text.split('\n')[0]] = m.text.split('\n')[1]
                                    except AttributeError:
                                        print('1st Goal not showing now.')
                            if s.text == "2nd Goal":
                                res_dict['2nd Goal'] = {}
                                for m in h.find_elements_by_css_selector(
                                        "div.gl-Participant"):
                                    try:
                                        print("2nd Goal\n" + m.text)
                                    except UnicodeEncodeError:
                                        pass
                                    try:
                                        res_dict['2nd Goal'][m.text.split('\n')[0]] = m.text.split('\n')[1]
                                    except AttributeError:
                                        print('2nd Goal not showing now.')
                            if s.text == "Alternative Match Goals":
                                val = e_wait(h, 'div.gl-MarketLabel').text
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
                                print(i + " " + element1 + ' : ' + element2)
                                res_dict[i] = (element1, element2)
                            except Exception:
                                print('There is no ' + str(i))
                        # Bar nums
                        for i, num in zip(('On Target', 'Off Target'), range(2)):
                            try:
                                print(driver.find_elements_by_css_selector("div.ml1-SoccerStatsBar"))
                                element1 = driver.find_elements_by_css_selector("div.ml1-SoccerStatsBar")[num].text.split('\n')[0]
                                element2 = driver.find_elements_by_css_selector("div.ml1-SoccerStatsBar")[num].text.split('\n')[1]
                                print(i + element1 + ':' + element2)
                                res_dict[i] = (element1, element2)
                            except Exception:
                                print('There is no ' + str(i))
                        if res_dict == {}:
                            continue
                        All_Res_Dict[match_name].append(res_dict)
#                        print(All_Res_Dict)
#                        print('\nNew appending dict')
#                        print(res_dict)
                        if count == roop_count:
                            try:
                                delta = x_wait(driver, "//div[@class='ipe-SummaryButton_Icon']")
                                delta.click()
                                events = x_wait(driver, "//div[contains(@class, 'ipe-SummaryNativeScroller_Content')]")
                                events_list = events.text.split('\n')
                                print(events_list)
                                events_dict[match_name] = events_list
                            except Exception:
                                print("events list can't get.")
                        driver.back()
                        print("back now")
                count += 1
            except Exception:
                pass
        else:
            print('Waiting...')
            time.sleep(60)

    driver.quit()
    del_key_list = []
    for key in All_Res_Dict:
        if not All_Res_Dict[key]:
            del_key_list.append(key)
    for k in del_key_list:
        del All_Res_Dict[k]
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
        # complement header
        name1_list = ('B', 'E', 'H', 'K', 'T', 'V', 'X', 'Z', 'AB')
        name2_list = ('D', 'G', 'J', 'M', 'U', 'W', 'Y', 'AA', 'AC')
        for col1, col2 in zip(name1_list, name2_list):
            try:
                SummarySheet[col1 + '2'].value = match.split(' v ')[0]
                SummarySheet[col2 + '2'].value = match.split(' v ')[1]
            except (KeyError, AttributeError):
                continue
        for info_dict in All_Res_Dict[match]:
            SummarySheet['A' + str(Row)].value = str(info_dict.get('Time', 'NoData'))
            for col in ['B', 'C', 'D']:
                try:
                    for key in info_dict.get('Fulltime Result'):
                        SummarySheet[col + str(Row)].value = str(info_dict['Fulltime Result'][key])
                except (KeyError, TypeError):
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col in ['E', 'F', 'G']:
                try:
                    for key in info_dict.get('Half Time Result'):
                        SummarySheet[col + str(Row)].value = str(info_dict['Half Time Result'][key])
                except (KeyError, TypeError):
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col in ['H', 'I', 'J']:
                try:
                    for key in info_dict.get('1st Goal'):
                        SummarySheet[col + str(Row)].value = str(info_dict['1st Goal'][key])
                except (KeyError, TypeError):
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col in ['K', 'L', 'M']:
                try:
                    for key in info_dict.get('2nd Goal'):
                        SummarySheet[col + str(Row)].value = str(info_dict['2nd Goal'][key])
                except (KeyError, TypeError):
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col, num in zip(['N', 'O', 'P'], range(3)):
                try:
                    SummarySheet[col + str(Row)].value = str(info_dict['Alternative Match Goals'][num])
                except (KeyError, TypeError):
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col, num in zip(['Q', 'R', 'S'], range(3)):
                try:
                    SummarySheet[col + str(Row)].value = str(info_dict['First Half Goals'][num])
                except (KeyError, TypeError):
                    SummarySheet[col + str(Row)].value = 'NoData'
            for i, j in zip([('T', 'U'), ('V', 'W'), ('X', 'Y'), ('Z', 'AA'), ('AB', 'AC')],(
                'Attacks', 'Dangerous Attacks', 'Possession', 'Off Target', 'On Target')):
                for col, num in zip(i, range(2)):
                    try:
                        SummarySheet[col + str(Row)].value = str(info_dict[j][num])
                    except (KeyError, TypeError):
                        SummarySheet[col + str(Row)].value = 'NoData'


            Row += 1
        try:
            for event in events_dict[match]:
                SummarySheet['AE' + str(info_Row)].value = event
                info_Row += 1
        except KeyError:
            SummarySheet['AE' + str(info_Row)].value = 'NoData'
    # Delete tmp Sheet
    ActiveWS = Target.get_sheet_by_name('Match_name')
    Target.remove_sheet(ActiveWS)
    # Save
    Target.save('Report/Res_' + NowTime + '.xlsx')


def e_wait(driver, element, max_time = 10):
    return_element = None
    try:
        return_element = WebDriverWait(driver, max_time).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, element))
                )
    except (NoSuchElementException, TimeoutException):
        print(element + " can't get!")
        pass
    return return_element

def c_wait(driver, element, max_time = 10):
    return_element = None
    try:
        return_element = WebDriverWait(driver, max_time).until(
                EC.element_to_be_clickable((By.XPATH, element))
                )
    except (NoSuchElementException, TimeoutException):
        print(element + " can't get!")
        pass
    return return_element


def x_wait(driver, element, max_time = 30):
    return_element = None
    try:
        return_element = WebDriverWait(driver, max_time).until(
                EC.presence_of_element_located((By.XPATH, element))
                )
    except (NoSuchElementException, TimeoutException):
        print(element + " can't get!")
        pass
    return return_element


def v_wait(driver, element, max_time = 10):
    return_element = None
    try:
        return_element = WebDriverWait(driver, max_time).until(
                EC.visibility_of_element_located((By.XPATH, element))
                )
    except NoSuchElementException:
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
