from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

import sys
from PyQt5.uic import loadUiType
from xlrd import *
from utility_functions import *
from my_thread_func import *
import traceback
from datetime import datetime
import requests

main_ui, _ = loadUiType('main.ui')

file_name = ''
excel_league = ''

def my_sleep(t):
    loop = QEventLoop()
    QTimer.singleShot(t * 1000, loop.quit)
    loop.exec_()

class MainApp(QMainWindow, main_ui):
    
    def __init__(self, current_tab=0):
        QMainWindow.__init__(self)

        self.threadpool = QThreadPool()
        print("Multithreading with maximum %d threads" % self.threadpool.maxThreadCount())
        self.current_tab = current_tab
        self.setupUi(self)
        self.at_start_main()

        self.Handel_Buttons()

    def at_start_main(self):
        self.btn_done.setVisible(False)

    def start_thread_write_excel(self):
        worker = Worker(self.write_excel_thread)  # Any other args, kwargs are passed to the run function
        #worker.signals.result.connect(self.deal_with_thread_get_inbox)
        worker.signals.finished.connect(self.thread_write_excel_complete)
        worker.signals.progress.connect(self.progressBar_players_fn)
        worker.signals.progress2.connect(self.progressBar_teams_fn)

        self.threadpool.start(worker)

    def write_excel_thread(self, progress_callback, progress_callback2):
        try:

            league = self.tournaments.currentText()
            
            if league == 'Premiership Rugby':
                URL = 'https://www.rugbypass.com/premiership/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Pro14':
                URL = 'https://www.rugbypass.com/pro-14/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Super Rugby Unlocked':
                URL = 'https://www.rugbypass.com/super-rugby-unlocked/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Super Rugby Aotearoa':
                URL = 'https://www.rugbypass.com/super-rugby-aotearoa/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Super Rugby Australia':
                URL = 'https://www.rugbypass.com/super-rugby-australia/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Six Nations':
                URL = 'https://www.rugbypass.com/six-nations/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Rugby Championship':
                URL = 'https://www.rugbypass.com/the-rugby-championship/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Internationals':
                URL = 'https://www.rugbypass.com/internationals/teams/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Heineken Cup':
                URL = 'https://www.rugbypass.com/european-champions-cup/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Super Rugby':
                URL = 'https://www.rugbypass.com/super-rugby/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Top 14':
                URL = 'https://www.rugbypass.com/top-14/teams/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Mitre 10 Cup':
                URL = 'https://www.rugbypass.com/mitre-10-cup/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Currie Cup':
                URL = 'https://www.rugbypass.com/currie-cup/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Challenge Cup':
                URL = 'https://www.rugbypass.com/challenge-cup/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Sevens':
                URL = 'https://www.rugbypass.com/sevens/'
                excel_league = league
                file_name = league + '.xlsx'
            elif league == 'Rugby World Cup':
                URL = 'https://www.rugbypass.com/rugby-world-cup/teams/'
                excel_league = league
                file_name = league + '.xlsx'


            # my_data = get_my_data(self.line_max_text, progress_callback, progress_callback2)

            teams = get_teams(URL)
            print(len(teams),"-----------",excel_league,"------",file_name)
            my_data = get_my_data(len(teams), URL, progress_callback, progress_callback2)
            write_xl = write_excel(my_data, file_name, excel_league)
            if write_xl:
                return 1
            else:
                return 0
        except:
            traceback.print_exc()
            return 0

    def thread_write_excel_complete(self):
        print("THREAD {thread_write_excel_complete} COMPLETE!")
        self.btn_done.setVisible(True)

        # Start Thread Send if Mode1

    # def deal_with_thread_get_inbox(self, result):
    #     print('deal_with_thread_get_inbox : ', result)
    #     is_inbox, val_inbox = result
    #     if is_inbox:
    #         print('Done get_inbox Successfully')
    #     else:
    #         print('Error in get_inbox', val_inbox)
    #         self.error_msg(val_inbox)

    def progressBar_players_fn(self, n):
        print(" done" , n , '%')
        try:
            self.progressBar_players.setValue(n)
        except:
            traceback.print_exc()

    def progressBar_teams_fn(self, n):
        print(" done" , n , '%')
        try:
            self.progressBar_teams.setValue(n)
        except:
            traceback.print_exc()

    def Handel_Buttons(self): 
        self.btn_start.clicked.connect(self.start_thread_write_excel)


    def error_msg(self, mssg):
        QMessageBox.warning(
            self, 'Error', str(mssg))

def main():
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec_()

    

if __name__ == '__main__':
    main()