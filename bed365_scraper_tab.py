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
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import csv
import openpyxl as px
import datetime

def set_up(url):
    options = Options()
#    options.add_argument('--headless')
#    options.binary_location = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    options.add_argument('--disable-gpu')
    prefs = {"profile.default_content_setting_values.notifications" : 2}
    options.add_experimental_option("prefs",prefs)
#    capabilities['platform'] = "WINDOWS"
#    capabilities['version'] = "7"
    # install chromedriver if not found and start chrome
    driver = webdriver.Chrome(chrome_options=options)
    driver.get(url)
    print('Searching in ' + url)
    element_to_home = x_wait(driver, '//div[@id="dv1"]/a/div[1]/img')
    element_to_home.click()
    return driver

def get_url(browser, url):
    driver.get(url)
    time.sleep(10)

def get_with_tab(d, element):
    main_window = d.current_window_handle
    print(element)
    element.send_keys(Keys.COMMAND + Keys.RETURN)
    d.switch_to.window(d.window_handles[1])
    time.sleep(10)
    return d

def make_tab(driver):
    main_window = driver.current_window_handle
    driver.execute_script("window.open('%s');" %url)
    return driver


def delete_tab(driver):
    driver.close()

def change_tab(driver, n):
    driver.switch_to.window(driver.window_handles[n])
    return driver

def get_games(driver):
    """
    SCRAPING CODE
    RETURN : DICT LIST
    """

    first_term = x_wait(driver, "//div[contains(@class, 'li-InPlayClassificationButton_HeaderLabel')]")
    print('Now first term : ' + first_term.text)
    if first_term.text == "Soccer":
        print('Soccer term found!')
        soccer = e_wait(driver, 'div.li-InPlayClassification_League')
        leagues = soccer.find_elements_by_css_selector('div.li-InPlayLeague')
        game_list = []
        for i in range(len(leagues)):
            league_num = i + 1
            if len(leagues) == 1:
                soccer = e_wait(driver, 'div.li-InPlayClassification_League')
                league = x_wait(soccer, "//div[@class='li-InPlayLeague ']")
            else:
                soccer = e_wait(driver, 'div.li-InPlayClassification_League')
                league = x_wait(soccer, "//div[@class='li-InPlayLeague ' and position()=%s]"%league_num)
            try:
                league_matchs = league.find_elements_by_css_selector('div.li-InPlayEvent')       
            except Exception:
                continue
            for j in range(len(league_matchs)):
                if len(league_matchs) == 1:
                    game_x = "//div[contains(@class, 'li-InPlayClassification ') and contains(., 'Soccer')]//div[@class='li-InPlayLeague ' and position()=%s]/div[2]/div/div[contains(@class, 'li-InPlayEventHeader ')]/div[1]/div[1]"%league_num
                elif len(league_matchs) > 1:
                    match_num = j + 1
                    game_x = "//div[contains(@class, 'li-InPlayClassification ') and contains(., 'Soccer')]//div[@class='li-InPlayLeague ' and position()=%s]/div[2]/div[contains(@class, 'li-InPlayEvent') and position()=%s]/div[@class='li-InPlayEventHeader ']/div[1]/div[1]" %(league_num, match_num)
                c = 0
                while c < 5:
                    try:
                        game = x_wait(driver, game_x)
                        match_name = str(game.text)#.split('\n')[0])
                        game_list.append((match_name, game_x))
                        break
                    except Exception: # match can't get : not exist or not visible
                        print(' does not clickable...\nscroll page!')
                        driver.execute_script('window.scrollBy(0,300);')
                        c += 1
                else: # not exist
                    driver.execute_script('window.scrollTo(0,0);')
                    print('!!!!!!!!!!')
                    break
        print(game_list)
        return game_list
    else:
        print('Waiting...')
        time.sleep(60)
        get_data(driver)

def crawl_tabs_setup(driver, game_list, now_scraping  = []):
    del_list = []
    for game_x in game_list:
        make_tab(driver)
    for n, game_x in enumerate([i[1] for i in game_list]):
        change_tab(driver, n + len(now_scraping) + 1)
        try:
            game = x_wait(driver, game_x)
            game.click()
        except:
            del_list.append(n)
        time.sleep(1)
    else:
        change_tab(driver, 0)
    for i in del_list:
        del game_list[i]
    return game_list


def crawl_tabs_getdata(driver, game_list):
    for i, game_name in zip(range(len(driver.window_handles[1:])), [j[1] for j in game_list]):
        print(len(driver.window_handles[1:]), len(game_list))
        change_tab(driver, i + 1)
        time.sleep(3)
        end_flag = get_data(driver, game_name)
        if end_flag == 1:
            delete_tab(driver)
            change_tab(driver, 0)
            break
    else:
        change_tab(driver, 0)


def get_data(driver, match_name):
    end_flag = 0
    if match_name not in AllResDict:
        try:
            print('====================' + match_name + ' start scraping!====================')
        except UnicodeError:
            print('====================start scraping!===========================')
        AllResDict[match_name] = []
    res_dict = {}
    try:
        time_reg  = e_wait(driver, "div.ipe-SoccerHeaderLayout_ExtraData")
        if int(time_reg[0]) == 9:
            end_flag += 1
            print('break!!')
        res_dict['Time'] = time_reg.text
    except Exception:
        print("Time can't get")
    for h in driver.find_elements_by_css_selector(
            "div.gl-MarketGroup"):
        try:
            s = e_wait(h, "div.gl-MarketGroupButton")
            print('s : ' + s.text)
        except Exception:
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
                except Exception:
                    print('Fulltime Result not showing now.')
        if s.text == "Half Time Result":
            res_dict['Half Time Result'] = {}
            for m in h.find_elements_by_css_selector(
                    "div.gl-Participant"):
                try:
                    res_dict['Half Time Result'][m.text.split('\n')[0]] = m.text.split('\n')[1]
                except Exception:
                    print('Half Time Result not showing now.')
        if s.text == "1st Goal":
            res_dict['1st Goal'] = {}
            for m in h.find_elements_by_css_selector(
                    "div.gl-Participant"):
                try:
                    print("1st Goal\n" + m.text)
                except Exception:
                    pass
                try:
                    res_dict['1st Goal'][m.text.split('\n')[0]] = m.text.split('\n')[1]
                except Exception:
                    print('1st Goal not showing now.')
        if s.text == "2nd Goal":
            res_dict['2nd Goal'] = {}
            for m in h.find_elements_by_css_selector(
                    "div.gl-Participant"):
                try:
                    print("2nd Goal\n" + m.text)
                except Exception:
                    pass
                try:
                    res_dict['2nd Goal'][m.text.split('\n')[0]] = m.text.split('\n')[1]
                except Exception:
                    print('2nd Goal not showing now.')
        if s.text == "Alternative Match Goals":
            try:
                val = e_wait(h, 'div.gl-MarketLabel').text
                over = h.find_elements_by_css_selector('div.gl-MarketValues')[0].text
                under = h.find_elements_by_css_selector('div.gl-MarketValues')[1].text
                altlist = [[m , n, l] for m, n, l in zip(val.split('\n'), over.split('\n')[1:], under.split('\n')[1:])]
                print("Alternative Match Goals")
                print(altlist)
                res_dict['Alternative Match Goals'] = [val, over.lstrip('Over\n'), under.lstrip('Under\n')]
            except Exception:
                print('Alternative Match Goals not showing now.')

# Boad
    corner = x_wait(driver, "//div[contains(@class, 'ipe-SoccerGridColumn_ICorner')]")
    yellow = x_wait(driver, "//div[@class='ipe-SoccerGridContainer ']/div[contains(@class, 'ipe-SoccerGridColumn_IYellowCard')]")
    red = x_wait(driver, "//div[@class='ipe-SoccerGridContainer ']/div[contains(@class, 'ipe-SoccerGridColumn_IRedCard')]")
    penalty = x_wait(driver, "//div[@class='ipe-SoccerGridContainer ']/div[contains(@class, 'ipe-SoccerGridColumn_IPenalty')]")
    goal = driver.find_elements_by_xpath("//div[@class='ipe-SoccerGridContainer ']/div[contains(@class, 'ipe-SoccerGridColumn')]")[-1]
    for i, j in zip(("Corner", "Yellow", "Red", "Penalty", "Goal"), (corner, yellow, red, penalty, goal)):
        try:
            print(i + ' : ' + ','.join(j.text.split("\n")))
            res_dict[i] = j.text.split("\n")
        except Exception:
            print(i + " can't get")
            res_dict[i] = 'NoData'
    # Wheel nums
    right_button = c_wait(driver, "//div[contains(@class, 'lv-ButtonBar_MatchLive')]")
    right_button.click()
    e_wait(driver, "div.ml1-StatWheel_Team1Text")
    for i, num in zip(('Attacks', 'Dangerous Attacks', 'Possession'), range(3)):
        try:
            element1 = driver.find_elements_by_css_selector("div.ml1-StatWheel_Team1Text")[num].text
            element2 = driver.find_elements_by_css_selector("div.ml1-StatWheel_Team2Text")[num].text
            print(i + ' ' + element1 + ':' + element2)
            res_dict[i] = (element1, element2)
        except Exception:
            print('There is no ' + str(i))
    # Bar nums
    for i, num in zip(('On Target', 'Off Target'), range(2)):
        try:
            element1 = driver.find_elements_by_css_selector("div.ml1-SoccerStatsBar")[num].text.split('\n')[0]
            element2 = driver.find_elements_by_css_selector("div.ml1-SoccerStatsBar")[num].text.split('\n')[1]
            print(i + ' ' + element1 + ':' + element2)
            res_dict[i] = (element1, element2)
        except Exception:
            print('There is no ' + str(i))
    AllResDict[match_name].append(res_dict)
#    try:
#        print(AllResDict)
#    except UnicodeError:
#        print("All_Res_Dict can't show up")
    print('\nNew appending dict')
    try:
        print(res_dict)
    except UnicodeError:
        print("res_dict can't show up")
    try:
        delta = x_wait(driver, "//div[@class='ipe-SummaryButton_Icon']")
        delta.click()
        events = x_wait(driver, "//div[contains(@class, 'ipe-SummaryNativeScroller_Content')]")
        events_list = events.text.split('\n')
        print(events_list)
        EventsDict[match_name] = events_list
    except Exception:
        print("events list can't get.")
    return end_flag

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

def get_diff(now_scraping, now_opening):
    appending_list = []
    for i in now_opening:
        if i[0] not in [k[0] for k in now_scraping]:
            appending_list.append(i)
    return appending_list


def ExcelWriter(AllResDict, EventsDict):
    def copy_range(fromsheet, tosheet, source_range, target_range):
        for i, j in zip(fromsheet[source_range],
                        tosheet[target_range]):
            for cell, new_cell in zip(i, j):
                new_cell._style = copy(cell._style)
                new_cell.value = cell.value

    try:
        os.mkdir('Report')
    except FileExistsError:
        print("Result is included in Report directory.")
        pass
    Template = './Template/Template.xlsx'
    Target = px.load_workbook(Template)

    print("Excel Writer Start!")
    # Summary Sheet
    for match in AllResDict.keys():
        Row = 3
        info_Row = 3
        ActiveWS = Target.get_sheet_by_name('Match_name')
        SummarySheet = Target.copy_worksheet(ActiveWS)
        try:
            SummarySheet.title = match[:31]
        except:
            SummarySheet.title = "#####"
        # complement header
        name1_list = ('B', 'D', 'F', 'H', 'J', 'L', 'O', 'R', 'U', 'AA', 'AD', 'AF', 'AH', 'AJ', 'AL')
        name2_list = ('C', 'E', 'G', 'I', 'K', 'N', 'Q', 'T', 'W', 'AC', 'AE', 'AG', 'AI', 'AK', 'AM')
        for col1, col2 in zip(name1_list, name2_list):
            try:
                SummarySheet[col1 + '2'].value = match.split(' v ')[0]
                SummarySheet[col2 + '2'].value = match.split(' v ')[1]
            except (KeyError, AttributeError):
                continue
        for info_dict in AllResDict[match]:
            SummarySheet['A' + str(Row)].value = str(info_dict.get('Time', 'NoData'))
            for col, num in zip(['B', 'C'], range(2)):
                try:
                    SummarySheet[col + str(Row)].value = str(info_dict['Corner'][num])
                except:
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col, num in zip(['D', 'E'], range(2)):
                try:
                    SummarySheet[col + str(Row)].value = str(info_dict['Yellow'][num])
                except:
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col, num in zip(['F', 'G'], range(2)):
                try:
                    SummarySheet[col + str(Row)].value = str(info_dict['Red'][num])
                except:
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col, num in zip(['H', 'I'], range(2)):
                try:
                    SummarySheet[col + str(Row)].value = str(info_dict['Penalty'][num])
                except:
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col, num in zip(['J', 'K'], range(2)):
                try:
                    SummarySheet[col + str(Row)].value = str(info_dict['Goal'][num])
                except:
                    SummarySheet[col + str(Row)].value = 'NoData'

            for col in ['L', 'M', 'N']:
                try:
                    for key in info_dict.get('Fulltime Result'):
                        SummarySheet[col + str(Row)].value = str(info_dict['Fulltime Result'][key])
                except (KeyError, TypeError):
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col in ['O', 'P', 'Q']:
                try:
                    for key in info_dict.get('Half Time Result'):
                        SummarySheet[col + str(Row)].value = str(info_dict['Half Time Result'][key])
                except (KeyError, TypeError):
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col in ['R', 'S', 'T']:
                try:
                    for key in info_dict.get('1st Goal'):
                        SummarySheet[col + str(Row)].value = str(info_dict['1st Goal'][key])
                except (KeyError, TypeError):
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col in ['U', 'V', 'W']:
                try:
                    for key in info_dict.get('2nd Goal'):
                        SummarySheet[col + str(Row)].value = str(info_dict['2nd Goal'][key])
                except (KeyError, TypeError):
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col, num in zip(['X', 'Y', 'Z'], range(3)):
                try:
                    SummarySheet[col + str(Row)].value = str(info_dict['Alternative Match Goals'][num])
                except (KeyError, TypeError):
                    SummarySheet[col + str(Row)].value = 'NoData'
            for col, num in zip(['AA', 'AB', 'AC'], range(3)):
                try:
                    SummarySheet[col + str(Row)].value = str(info_dict['First Half Goals'][num])
                except (KeyError, TypeError):
                    SummarySheet[col + str(Row)].value = 'NoData'
            for i, j in zip([('AD', 'AE'), ('AF', 'AG'), ('AH', 'AI'), ('AJ', 'AK'), ('AL', 'AM')],(
                'Attacks', 'Dangerous Attacks', 'Possession', 'Off Target', 'On Target')):
                for col, num in zip(i, range(2)):
                    try:
                        SummarySheet[col + str(Row)].value = str(info_dict[j][num])
                    except (KeyError, TypeError):
                        SummarySheet[col + str(Row)].value = 'NoData'


            Row += 1
        try:
            for event in EventsDict[match]:
                SummarySheet['AO' + str(info_Row)].value = event
                info_Row += 1
        except KeyError:
            SummarySheet['AO' + str(info_Row)].value = 'NoData'
    # Delete tmp Sheet
    ActiveWS = Target.get_sheet_by_name('Match_name')
    Target.remove_sheet(ActiveWS)
    # Save
    Target.save('Report/Res_' + FirstTime.strftime('%Y%m%d%H%M%S') + '.xlsx')

def time_delta(time_a, time_b):
    """
    calculate time_a - time_b(hour)
    """
    time_delta = time_a - time_b
    time_delta_sec = time_delta.total_seconds()
    time_delta_h = time_delta_sec / 3600
    return time_delta_h

def e_wait(driver, element, max_time = 10):
    return_element = None
    try:
        return_element = WebDriverWait(driver, max_time).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, element))
                )
    except Exception:
        print(element + " can't get!")
        pass
    return return_element

def c_wait(driver, element, max_time = 10):
    return_element = None
    try:
        return_element = WebDriverWait(driver, max_time).until(
                EC.element_to_be_clickable((By.XPATH, element))
                )
    except Exception:
        print(element + " can't get!")
        pass
    return return_element


def x_wait(driver, element, max_time = 15):
    return_element = None
    try:
        return_element = WebDriverWait(driver, max_time).until(
                EC.presence_of_element_located((By.XPATH, element))
                )
    except Exception:
        print(element + " can't get!")
        pass
    return return_element


def v_wait(driver, element, max_time = 10):
    return_element = None
    try:
        return_element = WebDriverWait(driver, max_time).until(
                EC.visibility_of_element_located((By.XPATH, element))
                )
    except Exception:
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
    parser.add_argument('-t', '--time_h',
            action='store',
            dest='time_h',
            default=1,
            type=float,
            help="specify run time with option '-t'",
            )
    parser.add_argument('-lc', '--loop_count',
            action='store',
            dest='l_count',
            default=10,
            type=int,
            help="specify loop count with option '-lc'")

 
    args = parser.parse_args()
    return args


def main():
    print('bed365_scraper START!!')
    global url, now_scraping, AllResDict, EventsDict
    url = 'https://www.bet365.com/home/'
    driver = set_up(url)
    # Get (geme_name, X_PASS) list
    game_list = get_games(driver)
    now_scraping = crawl_tabs_setup(driver, game_list)
    AllResDict = {}
    EventsDict = {}
    while True:
        try:
            NowTime = datetime.datetime.now()
            TimeDelta = time_delta(NowTime, FirstTime)
            print('Passed Time : ' + str(TimeDelta) + " h")
            if TimeDelta >= args.time_h:
                break
            now_opening = get_games(driver)
            appending_list = get_diff(now_opening, now_scraping)
            print(appending_list)
            if len(appending_list) > 0:
                # Add adding games(now opening - now scraping)
                crawl_tabs_setup(driver, appending_list, now_scraping)
            now_scraping = now_opening
            game_list = get_games(driver)
            crawl_tabs_getdata(driver, game_list)
        except:
            driver.quit()
            driver = set_up(url)
            # Get (geme_name, X_PASS) list
            game_list = get_games(driver)
            now_scraping = crawl_tabs_setup(driver, game_list)
    driver.quit()
    ExcelWriter(AllResDict, EventsDict)
    print('END!!')


if __name__ == '__main__':
    FirstTime = datetime.datetime.now()
    args = Argument_Parser()
    main()
