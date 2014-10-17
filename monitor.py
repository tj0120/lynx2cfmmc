import sys
import os.path
#from string import strip,join
#import re
import logging  
import logging.handlers
import datetime
import pyinotify
import time
import lynx2cfmmc
import rebatexlwt
from lynx2cfmmc import DealCMFChinaData
from rebatexlwt import TimRebate
from daemon import runner
from argparse import ArgumentParser
from multiprocessing import Process,Lock,Manager,Queue


 

class TimLogger():
    loggerFlag = False
    def __init__(self,name = u'settlement',rootdir = u'.'):
        self.rootName = name
        self.rootdir = rootdir
    def getLogger(self,name=u'',level = logging.INFO):
        if (TimLogger.loggerFlag):
            if (name):
                logger = logging.getLogger((self.rootName + u'.%s') % name)
                logger.setLevel(level) 
            else:
                logger = logging.getLogger(self.rootName) 
                logger.setLevel(level)               
            return logger
        else:
            LOG_FILE = self.rootName + u'.log'
            console = logging.StreamHandler()
            console.setLevel(logging.DEBUG)
            formatter = logging.Formatter(u'%(asctime)s - %(filename)s:%(lineno)s - %(name)s - %(message)s'  )
            handler = logging.handlers.RotatingFileHandler(os.path.join( os.path.realpath(self.rootdir),LOG_FILE), maxBytes = 1024*1024, backupCount = 5)
            handler.setFormatter(formatter)
            handler.setLevel(level)    
            if (name):
                logger = logging.getLogger((self.rootName + u'.%s') % name)
            else:
                logger = logging.getLogger(self.rootName)
            logger.addHandler(handler)
            logger.addHandler(console)
            TimLogger.loggerFlag = True
            return logger



class MyAPP():
    def __init__(self,g_rootdir, g_level):
        self.stdin_path = u'/dev/null'
        self.stdout_path = u'/dev/tty'
        self.stderr_path = u'/dev/tty'
        self.pidfile_path = u'/var/run/monitor.pid'
        self.pidfile_timeout = 5
        self.g_rootdir=g_rootdir
        self.g_level=g_level
                 
    def run(self):
        class MyEventHandler(pyinotify.ProcessEvent):
            def __init__(self,rootdir,logger,logger1,logger2):
                pyinotify.ProcessEvent.__init__(self)
                self.findFile1 = False
                self.findFile2 = False                    
                self.log = logger
                self.logger1 = logger1
                self.logger2 = logger2
                self.g_rootdir=rootdir
            def process_IN_ACCESS(self, event):
                self.log.info(u"ACCESS event:%s" % event.pathname)
            def process_IN_ATTRIB(self, event):
                self.log.info(u"ATTRIB event:%s" % event.pathname)
            def process_IN_CLOSE_NOWRITE(self, event):
                self.log.info(u"CLOSE_NOWRITE event:%s" % event.pathname)
            def process_IN_CLOSE_WRITE(self, event):
                self.log.info(u"CLOSE_WRITE event:%s" % event.pathname)
                #print(u"CLOSE_WRITE event:%s" % event.pathname)
                if (self.findFile1):
                        try:
                            (d,f) = os.path.split(event.pathname)
                            if (f.startswith(u"AccSum_") and f.endswith(u".xlsx")):
                                reload(lynx2cfmmc)
                                from lynx2cfmmc import DealCMFChinaData
                                xls = DealCMFChinaData(self.g_rootdir, mylogger = self.logger1)
                                if (xls.initFlag):
                                    xls(f[7:15],xlsfname = event.pathname )
                                os.remove(event.pathname)
                                self.findFile1 = False 
                        except Exception,e:
                            self.log.info(u"There is error:%s" % e)
                            #print(u"There is error:%s" % e)
                if (self.findFile2):
                        try :
                            (d,f) = os.path.split(event.pathname)
                            if (f.startswith(u"rebate") and f.endswith(u"q.txt")):
                                reload(rebatexlwt)
                                from rebatexlwt import TimRebate
                                xls = TimRebate(self.g_rootdir, mylogger = self.logger2)
                                if (xls.initFlag):
                                    xls(f[6:10],int(f[10]))
                                os.remove(event.pathname)
                                self.findFile2 = False          
                        except Exception,e:
                            self.log.info(u"There is error:%s" % e)
                            #print(u"There is error:%s" % e)
            def process_IN_CREATE(self, event): 
                self.log.info(u"CREATE event:%s" % event.pathname)
                #print(u"CREATE event:%s" % event.pathname)
                (d,f) = os.path.split(event.pathname)
                if (len(f) == 20):
                    if (f.startswith(u"AccSum_") and f.endswith(u".xlsx")):
                        self.findFile1 = True                         
                if (len(f) == 16):
                    if (f.startswith(u"rebate") and f.endswith(u"q.txt")):
                        self.findFile2 = True 
            def process_IN_DELETE(self, event):
                self.log.info(u"DELETE event:%s" % event.pathname)
            def process_IN_MODIFY(self, event):
                self.log.info(u"MODIFY event:%s" % event.pathname)
            def process_IN_OPEN(self, event):
                self.log.info(u"OPEN event:%s" % event.pathname)
        TLF = TimLogger(rootdir = self.g_rootdir)        
        logger = TLF.getLogger()
        logger1 = TLF.getLogger(name=u'cmfchina',level=self.g_level)
        fullpath1 = os.path.join(self.g_rootdir,u'cmfchina')
        logger2 = TLF.getLogger(name=u'rebate',level=self.g_level)
        fullpath2 = os.path.join(self.g_rootdir,u'rebate')
                  
        # watch manager
        wm = pyinotify.WatchManager()
        mask = pyinotify.IN_CREATE | pyinotify.IN_CLOSE_WRITE
        #mask =  pyinotify.ALL_EVENTS
        wm.add_watch(fullpath1, mask, rec=True)
        wm.add_watch(fullpath2, mask, rec=True)
        # event handler
        eh = MyEventHandler(self.g_rootdir,logger,logger1,logger2)
        # notifier
        notifier = pyinotify.Notifier(wm, eh)
        notifier.loop()


if __name__ == '__main__':  
    g_rootdir = r"/home/gfqhhk/doc/settlement"
    g_level = logging.INFO  
    app = MyAPP(g_rootdir, g_level)
    daemon_runner = runner.DaemonRunner(app)
    daemon_runner.do_action()    
