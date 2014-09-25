import os.path
#from string import strip,join
#import re
import logging  
import logging.handlers
import datetime

from lynx2cfmmc import DealCMFChinaData


def setLogger(name='rebate', rootdir='.',level = logging.INFO):
    LOG_FILE = 'rebate.log'
    handler = logging.handlers.RotatingFileHandler(os.path.join( os.path.realpath(rootdir),LOG_FILE), maxBytes = 1024*1024, backupCount = 5)
    fmt = '%(asctime)s - %(filename)s:%(lineno)s - %(name)s - %(message)s'  
    formatter = logging.Formatter(fmt)
    handler.setFormatter(formatter)
    logger = logging.getLogger()
    logger.addHandler(handler)
    logger.setLevel(level)    
    return logger



def Monitor():
    import pyinotify
    fullpath = os.path.join(g_rootdir,'cmfchina')
    class MyEventHandler(pyinotify.ProcessEvent):
        findFile = False
        def process_IN_ACCESS(self, event):
            print "ACCESS event:", event.pathname
            #pass
        def process_IN_ATTRIB(self, event):
            print "ATTRIB event:", event.pathname
            #pass
        def process_IN_CLOSE_NOWRITE(self, event):
            print "CLOSE_NOWRITE event:", event.pathname
            #pass
        def process_IN_CLOSE_WRITE(self, event):
            print "CLOSE_WRITE event:", event.pathname
            if (self.findFile):
                (d,f) = os.path.split(event.pathname)
                xls = DealCMFChinaData(g_settledDate,g_account,email = g_email)
                os.remove(event.pathname)
                self.findFile = False 
                
        def process_IN_CREATE(self, event): #AccSum_20140901.xlsx
            print "CREATE event:", event.pathname
            (d,f) = os.path.split(event.pathname)
            if (len(f) == 20):
                if (f.startswith("AccSum") and f.endswith("s.xlsx")):
                    self.findFile = True 
            
        def process_IN_DELETE(self, event):
            print "DELETE event:", event.pathname
            #pass
        def process_IN_MODIFY(self, event):
            print "MODIFY event:", event.pathname
            #pass
        def process_IN_OPEN(self, event):
            print "OPEN event:", event.pathname
            #pass
            
    # watch manager
    wm = pyinotify.WatchManager()
    mask = pyinotify.IN_CREATE | pyinotify.IN_CLOSE_WRITE
    #mask =  pyinotify.ALL_EVENTS
    wm.add_watch(fullpath, mask, rec=True)
    # event handler
    eh = MyEventHandler()
    # notifier
    notifier = pyinotify.Notifier(wm, eh)
    notifier.loop()
    

def main():
    global logger, g_rootdir, g_level ,g_settledDate , g_account, g_email
    import argparse
    __author__ = 'TianJun'
    parser = argparse.ArgumentParser(description='This is a rebate script by TianJun.')
    parser.add_argument('-d','--rootdir', help='Input log file dir,default is ".".',required=False)
    parser.add_argument('-l','--level', help='Input logger level, default is INFO.',required=False)
    parser.add_argument('-D','--date', help='Input SettledDate YYYYMMDD,default is today.',required=False)
    parser.add_argument('-A','--account', help='Input account XXXXXX-000',required=False)
    parser.add_argument('-m','--email', help='Input email to send result.',required=False)    
    args = parser.parse_args()    
    if (args.rootdir):
        g_rootdir = args.rootdir
    else:
        g_rootdir = os.path.curdir 
    if (args.level):
        g_level = args.level  
    else:
        g_level = logging.INFO 
    if (args.date):
        g_settledDate = args.date
    else:
        g_settledDate = None
    if (args.account):
        g_account = args.account       
    if (args.email):
        g_email = args.email   
    logger = setLogger(rootdir=m_rootdir,level=m_level)
    
    Monitor()
    #test()
    
    return 0

if __name__ == '__main__':

    main()
