import xlsxwriter
import traceback
import os
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import math
import re
from index import *
from urlextract import URLExtract
import win32com.client as win32
import pythoncom
from selenium  import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

# file_name = 'data.xlsx'

# file_name = ''
# excel_league = ''

# def get_excelLeague(value):
#     excel_league = value

# def get_fileName(value):
#     file_name = value


# def auto_fit():
#     try:
#         pwd = os.getcwd()
#         print(pwd)
#         # file_name = file_name

#         full_path = os.path.join(pwd, file_name)
#         print(full_path)
#         pythoncom.CoInitialize()
#         excel = win32.gencache.EnsureDispatch('Excel.Application')
#         wb = excel.Workbooks.Open(full_path)
#         ws = wb.Worksheets("Sheet1")
#         ws.Columns.AutoFit()
#         wb.Save()
#         excel.Application.Quit()
#         print('Done auto_fit')
#         return 1
#     except Exception as e:
#         print('Error in auto_fit', e)
#         traceback.print_exc()
#         return 0

header = {"User-Agent":'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'}


def get_teams(URL):
    print(URL,"=--------------------------------------------------------")
    # url_decision(league)

    # URL = 'https://www.rugbypass.com/premiership/'
    if URL.find('teams') == -1:
        page = requests.get(URL, headers=header)

        soup = BeautifulSoup(page.content,'html.parser')
        #count_item = soup.find(attrs={"class":"team"}).get_text()
        count_item = soup.find_all(attrs={"class":"team"})# .get_text()
        print('count_item', count_item)
        teams_list = []
        for tag in count_item:
            st = str(tag)
            if re.search('<td class="team"> <a href' , st):
                print(tag.get_text().strip())
                a = tag.find("a")
                print(a.attrs["href"])
                team_tuple = (tag.get_text().strip() , a.attrs["href"])
                print(team_tuple,"==============>>>>>>>")
                teams_list.append(team_tuple)
                # for a_elm in tag.find_all("a"):
                #     print(a_elm.attrs["href"])
        print(teams_list)
        return teams_list
    else:
        options = Options()
        options.headless = True
        driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=options)
        # driver.set_window_size(1920, 1080)
        driver.get(URL)

        teams_list = []
        elements = driver.find_elements_by_class_name('player-box-container')
        for element in elements:
            team_tuple = (element.get_attribute('href').split('teams/')[1].replace('/','').replace('-',' ').capitalize(),element.get_attribute('href'))
            print(team_tuple,"==============>>>>>>>")
            teams_list.append(team_tuple)

        return teams_list

            #print(tag)
    #total_item = count_item.split()[count_item.split().index('Menampilkan')+1]
    #divide_total = math.ceil(int(total_item)/60)
    #print(total_item)
    #return divide_total

def get_players(team_name, URL_TEAM):
    URL_Players = URL_TEAM + 'players/'
    page = requests.get(URL_Players, headers=header)
    soup = BeautifulSoup(page.content,'html.parser')
    try:
        sec = soup.find_all('section', attrs={'class': 'tournament full'})[0]
    except Exception as e:
        print(e)
        

    container = sec.find_all('div', attrs={'class': 'container'})
    clr = container[0].find('div', attrs={'class': 'clearfix'})
    clr = clr.find('div', attrs={'class': 'team-player-rankings'})
    clr = clr.find('div', attrs={'class': 'row'})
    clr = clr.find('div', attrs={'class': 'col-full rankings-players-col'})
    clr = clr.find('ul', attrs={'class': 'row team-players'})
    clr = clr.find_all('li', attrs={'class': 'player-box teams-players'})
    print(f'This Team {team_name}, Have {len(clr)} players .')
    players_list = []
    for tag in clr:
        a = tag.find("a")
        player_name = tag.attrs["data-name"]
        player_url = a.attrs["href"]
        player_tuple = (player_name, player_url)
        players_list.append(player_tuple)
    print(f'Done get_players() : {team_name}')
    return players_list


def get_player_info(player_name, URL_player):
    try:
        URL_Player = URL_player + 'statistics/'
        page = requests.get(URL_Player, headers=header)
        soup = BeautifulSoup(page.content,'html.parser')

        sec = soup.find_all('section', attrs={'class': 'team-player-menu'})[1]
        container = sec.find_all('div', attrs={'class': 'container'})
        tag = container[0].find('div', {'id': 'sub-content'})
        tag = tag.find('div', {'id': 'player-stats-details'})
        tag = tag.find('div', {'class': 'col-full all-stats-col'})
        tag = tag.find('ul', {'class': 'all-stats'})
        ul = tag.find_all('li')
        player_details_str = ''
        for li in ul:
            label = li.find('h4').get_text()
            #print(label)

            details = li.find_all('div', attrs={'class': 'all-stats-item'})
            #print(len(details))

            for det in details:
                val = det.find('div', attrs={'class': 'value'}).get_text().strip()
                #print(val)
                player_details_str += val +' '
            player_details_str += ','
        player_details_str = player_details_str[ : -2]
        print('player_details_str is ', player_details_str)
        return player_details_str
    except:
        return ',,'
        traceback.print_exc()

def get_my_data(max_teams, URL, progress_callback, progress_callback2):
    teams = get_teams(URL)
    my_data = []
    print('########## max_teams ', max_teams)
    final_max_teams = 0
    if max_teams > len(teams):
        final_max_teams = len(teams)
    else:
        final_max_teams = max_teams
    count_teams = 1
    for team in teams:
        team_name, team_url = team
        print('Start in Team : ', team_name)
        if team_name != 'Namibia' and team_name != 'Uruguay':
            player_list = get_players(team_name, team_url)
        count_player = 1
        for player in player_list:
            print(player)
            player_name, URL_player = player
            player_details_str = get_player_info(player_name, URL_player)
            player_tuple = (team_name + ',' + player_name, player_details_str)
            print(player_tuple)
            my_data.append(player_tuple)

            progress_callback.emit(count_player / len(player_list) * 100)

            count_player += 1

        progress_callback2.emit(count_teams / final_max_teams * 100)

        if count_teams == max_teams :
             break
        count_teams += 1
    print('Done get_my_data')
    return my_data

def write_excel(my_data, file_name, excel_league):
    try:
        len_all = 25
        # excel_league = excel_league

        if os.path.exists(file_name):
            os.remove(file_name)
        # Create an new Excel file and add a worksheet.
        workbook = xlsxwriter.Workbook(file_name)
        worksheet = workbook.add_worksheet()
        # Start Header

        # Add a bold format to use to highlight cells.
        print('Start Excel Header')
        border_center = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
            })
        bold_border_center = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
            })
        headers = ['Tournaments', 'Year', 'Round', 'Date', 'Team' , 'No.' , 'Player',]
        headers_cel = ['A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3']
        for hed, cel in zip(headers, headers_cel):
            worksheet.write(cel, hed, bold_border_center)
        #8
        Attack_letters = ['Po', 'Tr', 'Me', 'Ru', 'DB', 'CB', 'Pa', 'TA']
        for hed, cel in zip(Attack_letters, range(7,15)):
            worksheet.write(2, cel, hed, bold_border_center)
        #3
        Kicking_letters = ['Ki', 'CG', 'PG']
        for hed, cel in zip(Kicking_letters, range(15, 18)):
            worksheet.write(2, cel, hed, bold_border_center)
        #4
        Defence_letters = ['Ta', 'TM', 'TW', 'TC']
        for hed, cel in zip(Defence_letters, range(18, 22)):
            worksheet.write(2, cel, hed, bold_border_center)
        #3
        Discipline_letters = ['PC', 'YC', 'RC']
        for hed, cel in zip(Discipline_letters, range(22, 25)):
            worksheet.write(2, cel, hed, bold_border_center)

        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'yellow'})

        worksheet.merge_range('H2:O2', 'Attack', merge_format)
        worksheet.merge_range('P2:R2', 'Kicking', merge_format)
        worksheet.merge_range('S2:V2', 'Defence', merge_format)
        worksheet.merge_range('W2:Y2', 'Discipline', merge_format)
        worksheet.merge_range('H1:Y1', 'Player Stats', merge_format)

        print('Done Header')
        print('='*100)
        print('Start Writing Data')
        print('Data fetched is \n', my_data)

        start_col = 3
        cols = list(range(start_col, len(my_data) + start_col ))
        #rows = list(range(0,25))
        rows = list(range(0,25))
        for col in cols:
            date_time = datetime.now().strftime("%m/%d/%Y, %H:%M:%S")
            my_row = [excel_league, '2020','1', date_time ]
            f_t, s_t = my_data[col - 3]
            team, player = f_t.split(',')
            my_row.append(team)
            my_row.append(col - 2)
            my_row.append(player)

            for s in s_t.split(','):
                for i in s.split():
                    my_row.append(i)
            print('len(my_row)', len(my_row))
            print(my_row)
            cnt = 0
            #for row in rows:
            for row in range(len(my_row)):
                print('row', row, 'col', col)
                print('cnt ', cnt)
                try :
                    my_r = my_row[cnt]
                    print(my_r)
                    worksheet.write(col, row, my_r, border_center)
                except Exception as e:
                    print(e)
                    traceback.print_exc()

                cnt +=1
        print('Done writing Excel')
        workbook.close()
        # a = auto_fit()
        # print(a)
        # if a:
        #     return 1
        # else:
        #     return 0
    except Exception as e :
        print('Exception in writing fun', e)
        traceback.print_exc()
        return 0
