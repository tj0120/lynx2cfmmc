#!/usr/bin/env python
# -*- coding:utf8 -*-
#Version 6.1


import sys
reload(sys)
sys.setdefaultencoding('utf8')

import os.path
import copy
import xlrd
from xlwt import Utils
from string import strip,join
import re
import logging  
import logging.handlers
import datetime
import time
import calendar
import base64
 
def timestamp_datetime(value):
    format = '%Y-%m-%d %H:%M:%S'
    # value为传入的值为时间戳(整形)，如：1332888820
    value = time.localtime(value)
    ## 经过localtime转换后变成
    ## time.struct_time(tm_year=2012, tm_mon=3, tm_mday=28, tm_hour=6, tm_min=53, tm_sec=40, tm_wday=2, tm_yday=88, tm_isdst=0)
    # 最后再经过strftime函数转换为正常日期格式。
    dt = time.strftime(format, value)
    return dt
 
def datetime_timestamp(dt):
     #dt为字符串
     #中间过程，一般都需要将字符串转化为时间数组
     time.strptime(dt, '%Y-%m-%d %H:%M:%S')
     ## time.struct_time(tm_year=2012, tm_mon=3, tm_mday=28, tm_hour=6, tm_min=53, tm_sec=40, tm_wday=2, tm_yday=88, tm_isdst=-1)
     #将"2012-03-28 06:53:40"转化为时间戳
     s = time.mktime(time.strptime(dt, '%Y-%m-%d %H:%M:%S'))
     return int(s)


exch_name = {'CME':('CME','Chicago Mercantile Exchange Group',u'芝加哥商业交易所集团'),
             'CBOT':('CBOT','The Chicago Board of Trade',u'芝加哥期货交易所'),
             'NYME':('NYMEX','New York Mercantile Exchange',u'纽约商品交易所'),
             'NYBO':('NYBOT','New York Board of Trade',u'纽约期货交易所'),
             #'NYME':('COMEX','Commerce Exchange',u'美国（纽约）金属交易所'),
             'LME':('LME','London Metal Exchange',u'伦敦金属交易所'),
             #'NYBO':('ICE','International Exchange',u'洲际交易所'),
             'IPE':('IPE','International Petroleum Exchange',u'英国国际石油交易所'),
             'TCE':('TOCOM','Tokyo Commodity Exchange',u'东京商品交易所'),
             'SIME':('SGX','Singapore Exchange',u'新加坡交易所'),
             'MYX':('BMD','Bursa Mayaysia？',u'马来西亚衍生品交易所'),
             'CBOE':('CBOE','The Chicago Board of Trade',u'芝加哥期货交易所'), # new added, need check!
             'HKFE':('HKEX','Hong Kong Exchanges and Clear Limited',u'香港交易及结算有限公司') }


pil_vtitle = ('CMF_CODE','Market Code','Underlying','Currency','Maintenance Margin','Name','Price Per Ticker','Ticker','Multiplier')
pil_name  = {
            'LIF'       :('','CME','LI','USD',80.0,'1-MONTH LIBOR (CME)',0.0025,6.25,2500.0), 
            'JPYNF'     :('JY','CME','JPYN','USD',91.0,'JAPANESE YEN',1e-06,12.5,12500000.0),
            'COILF'     :('BRN','IPE','COIL','USD',100.0,'BRENT CRUDE OIL',0.01,10.0,1000.0),
            'FDAX'      :('FDAX','EURX','FDAX','EUR',100.0,'FDAX',0.5,12.5,25.0),
            'PAF'       :('PA','NYME','PA','USD',91.0,'PALLADIUM',0.05,5.0,100.0),
            'MSBF'      :('YK','CBOT','MSB','USD',80.0,'MINI-SIZED SOYBEANS',0.125,1.25,10.0),
            'CAC40'     :('FCE','LIFF','CAC40','EUR',100.0,'CAC40 10 EURO (FRANCE)',0.5,5.0,10.0),
            'RMBF'      :('RMB','CME','RMB','USD',91.0,'CME CHINESE RMB/USD CROSS RATE',1e-05,10.0,1000000.0),
            'EUDF'      :('','CME','EUD','USD',80.0,'3-MONTH EURO DOLLAR',0.005,12.5,2500.0),
            'SIN'       :('','SIME','SIN','USD',91.0,'CNX NIFTY INDEX',0.5,1.0,2.0),
            'MZCF'      :('YC','CBOT','MZC','USD',80.0,'MINI-SIZED CORN',0.125,1.25,10.0),
            'TNTWF'     :('TU','CBOT','TNTW','USD',91.0,'2 YRS US NOTES',0.0078125,15.625,2000.0),
            'LMALF'     :('AL','LME','AL','USD',100.0,'LME ALUMINIUM HG LME',0.5,12.5,25.0),
            'HKS'       :('','HKME','HKS','USD',80.0,'HKMEX SILVER',0.01,10.0,1000.0),
            'MJNIF'     :('','OSE','MJNI','JPY',80.0,'NIKKEI 225 MINI',5.0,500.0,100.0),
            'HSIF'      :('HSI','HKFE','HSI','HKD',80.0,'HANG SENG INDEX',1.0,50.0,50.0),
            'NZDF'      :('ND','CME','NZD','USD',91.0,'NEW ZEALAND DOLLAR',0.0001,10.0,100000.0),
            'CHH'       :('CHH','HKFE','CHH','HKD',80.0,'CSE120 INDEX FUTURES',0.5,25.0,50.0),
            'MSCITWF'   :('','SIME','MSCI','USD',91.0,'MSCI MSCI TAIWAN INDEX',1.0,100.0,100.0),
            'NAIF'      :('','CME','NAI','USD',91.0,'NASDAP NDEX',0.25,25.0,100.0),
            'SNF'       :('SN','LME','SN','USD',100.0,'TIN',5.0,25.0,5.0),
            'FCPO'      :('','KLSE','FCPO','MYR',100.0,'CRUDE PALM OIL',1.0,25.0,25.0),
            'HHIF'      :('HHI','HKFE','HHI','HKD',80.0,'H-SHARE INDEX',1.0,50.0,50.0),
            'MDJF'      :('YM','CBOT','MDJ','USD',91.0,'MINI SIZED DOW JONES IND $5',1.0,5.0,5.0),
            'NIF'       :('NI','LME','NI','USD',100.0,'PRIMARY NICKEL',5.0,30.0,6.0),
            'MDJ10'     :('','CBOT','MDJ10','USD',91.0,'MINI DOW $10',1.0,10.0,10.0),
            'COFF'      :('KC','NYBO','COF','USD',91.0,'COFFEE',0.05,18.75,375.0),
            'SIF'       :('SL','NYME','SI','USD',91.0,'SILVER',0.005,25.0,5000.0),
            'NGF'       :('NG','NYME','NG','USD',91.0,'NATURAL GAS',0.001,10.0,10000.0),
            'ZLF'       :('BO','CBOT','ZL','USD',74.1,'SOYBEANS OIL',0.01,6.0,600.0),
            'LMALFMINI' :('','SIME','AHMI','USD',91.0,'MINI-ALUMINUM SGX',0.5,2.5,5.0),
            'LSNO'      :('','LME','LSNO','USD',80.0,'LME ZN OPTION',0.5,12.5,25.0),
            'ZWF'       :('W','CBOT','ZW','USD',74.1,'WHEAT',0.25,12.5,50.0),
            'IXCA'      :('','HKME','IXCA','USD',80.0,'32 TROY OZ GOLD FUTURES',0.1,3.2,32.0),
            'UGF'       :('RB','NYME','UG','USD',91.0,'UNLEADED GASOLINE',0.0001,4.2,42000.0),
            'OJF'       :('','NYBO','OJ','USD',91.0,'O.J.FROZEN',0.05,7.5,150.0),
            'SHGZNF'    :('ZN','LME','SHGZN','USD',100.0,'SHG ZINC',0.5,2.5,5.0),
            'FTSE'      :('UKX','LIFF','FTSE','GBP',100.0,'FTSE 100',0.5,5.0,10.0),
            'EUFXF'     :('EUFX','CME','EUFX','USD',91.0,'EURO FX FUTURE',0.0001,12.5,125000.0),
            'ZI'        :('','CBOT','ZI','USD',80.0,'5000 OZ. SILVER',0.001,5.0,5000.0),
            'ZOF'       :('','CBOT','ZO','USD',74.1,'OAT',0.25,12.5,50.0),
            'COTF'      :('CT','NYBO','COT','USD',91.0,'COTTON',0.01,5.0,500.0),
            'JRUF'      :('JTU','TCE','JRU','JPY',100.0,'TOCOM RUBBER',0.1,500.0,5000.0),
            'ZGF'       :('','CBOT','ZG','USD',80.0,'CBOT 100 OZ. GOLD',0.1,10.0,100.0),
            'KPO'       :('FCPO','MYX','KPO','MYR',100.0,'CRUDE PALM OIL',1.0,25.0,25.0),
            'MCHF'      :('MCH','HKFE','MCH','HKD',80.0,'MINI H-SHARE INDEX',1.0,10.0,10.0),
            'SSI'       :('NK','SIME','SSI','JPY',91.0,'SIG NIKKEI 225 INDEX',1.0,500.0,500.0),
            'COCOF'     :('CC','NYBO','COCO','USD',91.0,'COCOA',1.0,10.0,10.0),
            'SSG'       :('TW','SIME','SSG','SGD',91.0,'MSCI SINGAPORE FREE INDEX',0.1,20.0,200.0),
            'MX'        :('','','MX','NTD',77.2,'小台指',1.0,50.0,50.0),
            'EURO/YEN'  :('','CME','RY','JPY',91.0,'EURO FX/YEN IMM',0.01,0.0,0.0),
            'TNTEF'     :('TY','CBOT','TNTE','USD',91.0,'10 YRS US NOTES',0.015625,15.625,1000.0),
            'GLD'       :('','HKFE','GLD','USD',80.0,'GOLD HK',1.0,100.0,100.0),
            'AUDF'      :('AD','CME','AUD','USD',91.0,'AUSTRALIAN DOLLAR',0.0001,10.0,100000.0),
            'VX'        :('VX','CBOE','VX','USD',91.0,'VX',0.05,50.0,1000.0),
            'USTBF'     :('','CME','USTB','USD',80.0,'3-MONTH U.S.TREASURY BILL',0.001,2.5,2500.0),
            'PEF'       :('PB','LME','PE','USD',100.0,'LEAD',0.5,12.5,25.0),
            'MYIF'      :('','CBOT','MYI','USD',80.0,'MINI-SIZED SILVER',0.001,1.0,1000.0),
            'MDJ25'     :('','CBOT','MDJ25','USD',91.0,'MINI DOW $25',1.0,25.0,25.0),
            'LCOF'      :('CL','NYME','LCO','USD',91.0,'LIGHT CRUDE OIL',0.01,10.0,1000.0),
            'SWFCF'     :('SF','CME','SWFC','USD',91.0,'SWISS FRANC',0.0001,12.5,125000.0),
            'QMF'       :('QM','NYME','QM','USD',91.0,'MINI CRUDE OIL',0.025,12.5,500.0),
            'MSPIF'     :('','CME','MSPI','USD',91.0,'S & P MID CAP 400',0.05,25.0,500.0),
            'GOLDF'     :('GC','NYME','GOLD','USD',91.0,'GOLD NYMEX',0.1,10.0,100.0),
            'EMNAIF'    :('NQ','CME','EMNAI','USD',91.0,'E-MINI NASDAQ 100 INDEX',0.25,5.0,20.0),
            'GOLDO'     :('','NYME','OG','USD',80.0,'GOLD OPTIONS',0.1,10.0,100.0),
            'SDZN'      :('','SIME','ZSMI','USD',80.0,'MINI ZINC SGX',0.5,2.5,5.0),
            'NYALF'     :('','NYME','NYAL','USD',80.0,'ALUMINIUM COMEX',0.0005,22.0,44000.0),
            'CUF'       :('CU','LME','CUA','USD',100.0,'COPPER',0.5,12.5,25.0),
            'SPIF'      :('SP','CME','SPI','USD',80.0,'S & P 500 STOCK PRICE INDEX',0.1,25.0,250.0),
            'NYDIF'     :('','NYBO','NYDI','USD',90.0,'DOLLAR INDEX',0.005,5.0,1000.0),
            'EUR'       :('EC','CME','EUR','USD',80.0,'IMM EURO DOLLARS',0.0001,0.25,2500.0),
            'XINA50'    :('XU','SIME','SA50','USD',91.0,'SGX CHINA A50 FUTURES',5.0,5.0,1.0),
            'PLF'       :('PL','NYME','PL','USD',91.0,'PLATINUM',0.1,5.0,50.0),
            'SO'        :('','NYME','SO','USD',80.0,'SILVER OPTION',0.1,5.0,50.0),
            'CUFMINI'   :('','SIME','CUMI','USD',80.0,'MINI-COPPER SGX',0.5,2.5,5.0),
            'AAF'       :('AAD','LME','AA','USD',100.0,'ALUMINIUM ALLOY',0.5,10.0,20.0),
            'CADF'      :('CD','CME','CAD','USD',91.0,'CANADIAN DOLLAR',0.0001,10.0,100000.0),
            'SUGF'      :('SB','NYBO','SUG','USD',91.0,'WORLD SUGAR NO.11',0.01,11.2,1120.0),
            'VHS'       :('VHS','HKFE','VHS','HKD',80.0,'HSI VOLATILITY INDEX FUTURES',0.01,50.0,5000.0),
            'ZSF'       :('S','CBOT','ZS','USD',74.1,'SOYBEANS',0.25,12.5,50.0),
            'DX'        :('DX','NYBO','DX','USD',91.0,'ICE US DOLLAR INDEX FUTURES',0.005,5.0,1000.0),
            'LZNO'      :('','LME','SHGZN','USD',100.0,'SHG ZINC OPTIONS',0.5,12.5,25.0),
            'CUS'       :('CUS','HKFE','CUS','CNY',80.0,'RMB CURRENCY FUTURES',0.0001,10.0,100000.0),
            'GAOF'      :('GAS','IPE','GAO','USD',100.0,'GAS OIL',0.25,25.0,100.0),
            'FGBS'      :('FGBS','EURX','FGBS','EUR',100.0,'EURO-SCHATZ FUTURES',0.005,5.0,1000.0),
            'LB'        :('','CME','LB','USD',80.0,'LUMBER CME',0.1,11.0,110.0),
            'TX'        :('','TW','TX','NTD',80.0,'TW CAPITALIZATION STOCK INDEX',1.0,200.0,200.0),
            'TNFIF'     :('FV','CBOT','TNFI','USD',91.0,'5 YRS US NOTES',0.0078125,7.8125,1000.0),
            'MYGF'      :('','CBOT','MYG','USD',80.0,'MINI-SIZED GOLD',0.1,3.32,33.2),
            'TF'        :('','TW','TF','NTD',80.0,'TW FINANCE SECTOR INDEX',0.2,200.0,1000.0),
            'TE'        :('','TW','TE','NTD',80.0,'TW ELECTRONIC SECTOR INDEX',0.05,200.0,4000.0),
            'FGBM'      :('FGBM','EURX','FGBM','EUR',100.0,'EURO-BOBL FUTURES',0.01,10.0,1000.0),
            'FGBL'      :('FGBL','EURX','FGBL','EUR',100.0,'EURO-BUND FUTURES',0.01,10.0,1000.0),
            'HOF'       :('HO','NYME','HO','USD',91.0,'HEATING OIL',0.0001,4.2,42000.0),
            'TXO'       :('','TW','TXO','NTD',80.0,'臺灣大台指期權',0.1,5.0,50.0),
            'ZMF'       :('SM','CBOT','ZM','USD',74.1,'SOYBEANS MEAL',0.1,10.0,100.0),
            'USTBOF'    :('','CBOT','USTBO','USD',91.0,'U.S.TREASURY BOND',0.015625,15.625,1000.0),
            'EMSPIF'    :('ES','CME','EMSPI','USD',91.0,'E-MINI S & P 500 INDEX FUT',0.25,12.5,50.0),
            'ZRF'       :('RICE','CBOT','ZR','USD',74.1,'ROUGH RICE',0.005,10.0,2000.0),
            'ESO'       :('ESO','CME','ESO','USD',80.0,'MINI S&P500',0.05,2.5,50.0),
            'FESX'      :('','EURX','FESX','EUR',100.0,'FESX',1.0,10.0,10.0),
            'ZCF'       :('C','CBOT','ZC','USD',74.1,'CORN',0.25,12.5,50.0),
            'BPNDF'     :('BP','CME','BPND','USD',91.0,'BRITISH POUND',0.0001,6.25,62500.0),
            'NIKF'      :('','CME','NIK','USD',80.0,'NIKKEI 225 STOCK PRICE INDEX',5.0,25.0,5.0),
            'JNIF'      :('','OSE','JNI','JPY',80.0,'OSAKA NIKKEI 225',10.0,10000.0,1000.0),
            'HGCUF'     :('HG','NYME','HGCU','USD',91.0,'HIGH GRADE COPPER',0.05,12.5,250.0),
            'MHSIO'     :('MHSIO','HKFE','MHSIO','HKD',80.0,'',1.0,10.0,10.0),
            'HSIO'      :('HSIO','HKFE','HSIO','HKD',80.0,'',1.0,50.0,50.0), 
            'QI'        :('','NYME','QI','USD',91.0,'MINI SILVER',0.0125,31.25,2500.0),
            'MHSIF'     :('MHU','HKFE','MHSI','HKD',80.0,'MINI-HANG SENG INDEX',1.0,10.0,10.0),
            'QO'        :('MGC','NYME','QO','USD',91.0,'MINI GOLD',0.25,12.5,50.0),
            'undefine01':('ND','CME','ND','USD',100.0,'Nasdaq 100',0.0,0.0,1.0),
            #'undefine02':('ES','CME','ES','USD',100.0,'S&P 500',0.0,0.0,1.0),
            'undefine03':('DD','CBOT','DD','USD',100.0,'Dow Jones',0.0,0.0,1.0),
            #'undefine04':('CHH','HKFE','','',100.0,'CES 120, 中华交易服务中国120指数期货',0.0,0.0,1.0),
            'undefine05':('BSE','HKFE','','',100.0,'SENSEX Index, 印度股市指數期貨',0.0,0.0,1.0),
            'undefine06':('CHH','HKFE','','',100.0,'CES China 120 Index, 中華120期貨',0.0,0.0,1.0),
            'undefine07':('MCX','HKFE','','',100.0,'MICEX Index, 莫斯科貨幣交易所指數期貨',0.0,0.0,1.0),
            'undefine08':('SAF','HKFE','','',100.0,'FTSE/JSE Top40, 南非領先40指數',0.0,0.0,1.0),
            'undefine09':('BOV','HKFE','','',100.0,'IBOVESPA, 巴西指數期貨',0.0,0.0,1.0),
            'undefine10':('KOSPI','KRX','','',100.0,'KOSPI 200, 韩国KOSPI 200股指',0.0,0.0,1.0),
            'undefine11':('AP','ASX','','',100.0,'ASX 200 Index Futures, 澳大利亚ASX200指数',0.0,0.0,1.0),
            'undefine12':('FESX','EURX','','',100.0,'EURO STOXX 50 Index, 欧元斯托克50',0.0,0.0,1.0),
            'undefine13':('E7','CME','','',100.0,'E-mini Euro Currencies, 小型欧元',0.0,0.0,1.0),
            'undefine14':('JT','CME','','',100.0,'E-mini Japanese Yen, 小型日元',0.0,0.0,1.0),
            'undefine15':('M6A','CME','','',100.0,'Micro AUD/USD, 微澳元/美元',0.0,0.0,1.0),
            'undefine16':('M6E','CME','','',100.0,'Micro EUR/USD, 微歐元/美元',0.0,0.0,1.0),
            'undefine17':('US','CBOT','','',100.0,'30 YrU.S. Treasury Bond Futures, 美国三十年票据',0.0,0.0,1.0),
            'undefine18':('QH','NYME','','',100.0,'E-Mini Heating Oil, 小型取暖油',0.0,0.0,1.0),
            'undefine19':('QG','NYME','','',100.0,'E-Mini Natural Gas, 小型天然气',0.0,0.0,1.0),
            'undefine20':('OJ','NYBO','','',100.0,'FCOJ, 橙汁',0.0,0.0,1.0),
            'undefine21':('YW','CBOT','','',100.0,'Mini Wheat, 小型小麦',0.0,0.0,1.0),
            'undefine22':('O','CBOT','','',100.0,'Oats, 燕麦',0.0,0.0,1.0),
            'undefine23':('DPPM','DGCX','','',100.0,'DGCX POLYPROPYLENE (PLASTICS), 聚丙烯塑料期货',0.0,0.0,1.0),
            'undefine24':('ECO','ENPAR','','',100.0,'RAPESEED, 聚丙烯塑料期货',0.0,0.0,1.0),
            'undefine25':('RS','ICE CANADA','','',100.0,'CANOLA, 油菜籽',0.0,0.0,1.0),
            'undefine26':('CTC','LCH','','',100.0,'OTC, CAPESIZE TC AVG 4 ROUTES, 航运期货',0.0,0.0,1.0),
            'undefine27':('PTC','LCH','','',100.0,'OTC, PANAMAX TC AVG 4 ROUTES, 航运期货',0.0,0.0,1.0),
            'undefine28':('FE','SIMEX','','',100.0,'OTC, IRON ORE SWAP, 铁矿石',0.0,0.0,1.0),
            'undefine29':('','','','',100.0,'',0.0,0.0,1.0),
            'undefine30':('','','','',100.0,'',0.0,0.0,1.0)
            }

settlers = {'jp':'J.P. Morgan',
            'ba':'Barclays Group',
            'rj':'RJO',
            'mr':'Marex',
            'su':'Sucden',
            'pi':'pillip',
            'gf':'GFFM',
            'MAREX':'mr',
            'PHILL':'ph',
            'GFFM':'gf',
            'BRIEN':'br',
            'SUCDN':'su',
            'DCASS':'hk' }
        
         
            
banks = {'01':('01',u'中国工商银行'),
        '02':('02',u'中国农业银行'),
        '03':('03',u'中国银行'),
        '04':('04',u'中国建设银行'),
        '05':('05',u'交通银行'),
        '99':('99',u'其它银行'),
        '001':('99','384 TO 269 | COMMISSION ADJUSTMENT'),
        '002':('99','CURRENCY TRANSFER'),
        '004':('99','ERROR TRADE | INTER CLIENT TRANSFER'),
        '005':('99','COMMISSION REBATE | HSBC USD 848162608274'),   
        '006':('99','FUND ADJUSTMENT | FEE ADJUSTMENT'),        
        '0210':('99','ICBC HKD 861-520-03021-0'),
        '0384':('99','SCB USD T S/A# 447-1-669038-4'),
        '1783':('99','SCB HKD T S/A# 447-1-780178-3'),
        '1821':('99','SCB USD T S/A# 447-1-780182-1'),
        '2193':('99','SCB CNY T S/A# 447-1-778219-3'),
        '6269':('99','SCB HKD T S/A# 447-1-669626-9'),
        '6667':('99','ICBC USD 861-530-02666-7'),
        '9367':('99','SCB HKD T C/A# 447-0-668936-7'),
        '9405':('99','SCB USD T C/A# 447-0-668940-5'),
        'FIMAT':('99','BROKER ACCOUNT-FIMAT'),
        'OTHER':('99','OTHER'),
        'SUCDN':('99','BROKER ACCOUNT-SUCDEN'),
        'ZAUTO':('99','AUTO BROKER')}

date_pat = r"\d{2} (:?JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC) \d{4}"
account_pat = r'\d{6}-\d{3}'


sc = { u'Trade Confirmation Summary':{  'Exchange':('B',),
                                        'Product':('C',),
                                        'No of Lots(buy)':('D',),
                                        'No of Lots(sell)':('E',),
                                        'Average Trading price(buy)':('F',),
                                        'Average Trading price(sell)':('G',)},
       u'Open Position Summary'     :{  'Exchange':('H',),
                                        'Product':('I',),
                                        'Prompt Date':('J',),
                                        'No of Lots(buy)':('K',),
                                        'No of Lots(sell)':('L',),
                                        'Average Trading price(buy)':('M',),
                                        'Average Trading price(sell)':('N',)},
       u'Unsettled Closed Position Summary'     :{  'Exchange':('O',),
                                        'Product':('P',),
                                        'Prompt Date':('Q',),
                                        'No of Lots(buy)':('R',),
                                        'No of Lots(sell)':('S',),
                                        'Average Trading price(buy)':('T',),
                                        'Average Trading price(sell)':('U',)},
       u'Closed Position Summary'   :{  'Exchange':('V',),
                                        'Product':('W',),
                                        'No of Lots Closed':('X',),
                                        'Prompt Date':('Y',),
                                        'Currency':('Z',),
                                        'Average Trading price(buy)':('AA',),
                                        'Average Trading price(sell)':('AB',),
                                        'Net PxlsSheet = self.xlsBook.sheet_by_name(rofit':('AC',)},
       u'Fund Movement'             :{  'Currency':('AD',),
                                        'Exchange Rate':('AE',),
                                        'Net Fund Movement':('AF',)},
       u'Account Summary'           :{  'Currency':('AG',),
                                        'Exchange Rate':('AH',),
                                        'Open Balance':('AI',),
                                        'Deposit / (Withdrawal)':('AJ',),
                                        'Commission':('AK',),
                                        'Fees & Levy':('AL',),
                                        'Trading P/(L)':('AM',),
                                        'Accrued Interest':('AN',),
                                        'Closing Balance':('AO',),
                                        'Gross Floating P/(L)':('AP',),
                                        'Unapproved P/(L)':('AQ',),
                                        'Equity':('AR',),
                                        'Initial Margin':('AS',),
                                        'Maintenance Margin':('AT',),
                                        'Margin (Call)/Excess':('AU',)},
       u'Trade Confirmation Full Details': {
                                        'Trade Date':('AV',),
                                        'Exchange':('AW',), 
                                        'Trade Ref.':('AX',), #新增 970-20140828-0000089 
                                        'Broker':('AY',),
                                        'Open or Close':('AZ',), #新增 * 开仓 # 平仓
                                        'No of Lots(buy)':('BA',),
                                        'No of Lots(sell)':('BB',),
                                        'Description':('BC',),
                                        'Trade Price/Premium':('BD',),
                                        'Currency1':('BE',),
                                        'Commission':('BF',),
                                        'Currence2':('BG',),
                                        'Market Charges and Fees':('BH',)},
       u'Open Position Full Details':{
                                        'Trade Date':('BI',),
                                        'Exchange':('BJ',),
                                        'Trade Ref.':('BK',), #新增 970-20140828-0000089 
                                        'Broker':('BL',),
                                        'No of Lots(buy)':('BM',),
                                        'No of Lots(sell)':('BN',),
                                        'Description':('BO',),
                                        'Trade Price/Premium':('BP',),
                                        'Closing Price':('BQ',),
                                        'Currency':('BR',),
                                        'Gross Floating P/(L)':('BS',),
                                        'Unapproved P/(L)':('BT',) },
       u'Unsettled Closed Position Full Details':{
                                        'Trade Date':('BU',),
                                        'Exchange':('BV',),
                                        'Trade Ref.':('BW',), #新增 970-20140828-0000089 
                                        'Broker':('BX',),
                                        'No of Lots(buy)':('BY',),
                                        'No of Lots(sell)':('BZ',),
                                        'Description':('CA',),
                                        'Trade Price/Premium':('CB',),
                                        'Closing Price':('CC',),
                                        'Currency':('CD',),
                                        'Gross Floating P/(L)':('CE',),
                                        'Unapproved P/(L)':('CF',) },
       u'Closed Position Full Details':{
                                        'Open or Close':('CG',), #新增 * 开仓 # 平仓
                                        'Trade Date':('CH',),
                                        'Exchange':('CI',),
                                        'Trade Ref.':('CJ',), #新增 970-20140828-0000089 
                                        'No of Lots(buy)':('CK',),
                                        'No of Lots(sell)':('CL',),
                                        'Description':('CM',),
                                        'Trade Price/ Permium (Buy)':('CN',),
                                        'Trade Price/ Permium (Sell)':('CO',),
                                        'Currency':('CP',),
                                        'Gross Profit/(Loss)':('CQ',) },
       u'Fund Movement Full Details':{
                                        'Currency':('CR',),
                                        'Transaction Type':('CS',),
                                        'Bank':('CT',),
                                        'Exchange Rate':('CU',),
                                        'Withdrawal':('CV',),
                                        'Deposit':('CW',) } }


cashtype  =  {
            'WITHDRAW MONEY':('O','00001'),
            'DEPOSIT IN BANK':('I','00002'),
            'BROKER ADJUSTMENT':('I','00003'),
            'CCY TRANSFER':('O','00004'),
            'INTER CLIENT TRANSFER':('','00005'),
            'COMMISSION ADJUSTMENT':('','00006'),
            'COMMISSION REBATE':('I','00007'),
            'FEE ADJUSTMENT':('','00008'),
            'FUND ADJUSTMENT':('I','00009'),
            'FUND OUT FOR PHYSICAL DELIVERY':('O','00010'),
            'DELIVERY FEE':('O','00011'),
            'FUND IN FOR PHYSCIAL DELIVERY':('I','00012'),
            'INTER-CURRENCY TRANSFER':('I','00100'),
            'ERROR TRADE':('O','00101'),
            'SYSTEM TRANSFER B/F':('I','Z0001'),
            'AUTO BROKER':('I','ZAUTO'),
            }

LOG_FILE = 'cmf.log'
handler = logging.handlers.RotatingFileHandler(os.path.join( os.path.realpath(os.path.curdir),LOG_FILE), maxBytes = 1024*1024, backupCount = 5)
fmt = '%(asctime)s - %(filename)s:%(lineno)s - %(name)s - %(message)s'  
formatter = logging.Formatter(fmt)
handler.setFormatter(formatter)
logger = logging.getLogger('cmf')
logger.addHandler(handler)
#logger.setLevel(logging.DEBUG)
logger.setLevel(logging.ERROR)


mail_from = ''
mail_server = ''
mail_id = ''
mail_pw = ''

class DealCMFChinaData(object):
    def __init__(self,p_date,account=None,xlsfname=None,email=None): # 带上参数即可只输出单一指定帐户的数据
        self.sheet_name = u'Account Summary'
        self.flagCompany = '0001'
        self.dateOfFileName = p_date
        self.sendEMail = email
        if (xlsfname):
            self.xlsFName = xlsfname
        else:
            self.xlsFName = u'AccSum_%s.xlsx' % self.dateOfFileName
        self.dirname = os.path.realpath(os.path.curdir)
        self.bookfilename = os.path.join( self.dirname,self.xlsFName )
        if (not os.path.exists(self.bookfilename)):
            print('Lynx export Xls File:%s not existed!' % self.xlsFName)    
            print('usage: sp2cmf.py [-h] [-d DATE] [-a ACCOUNT] [-f XLSFNAME] [-m EMAIL]')    
            sys.exit(0)    
        self.initLIST()
        self.initXLS()
        self.run(account)
            
    def initLIST(self):
        self.accountList = []
        self.openPositions = []
        self.unsettledClosedPositions = []
        self.fundMovement = []
        self.dailyAccountSummary = []
        self.closedPositionSummary = []
        self.tradeConfirmationSummary = []
        self.currencyRate = []
        self.exchangeRecord = []
        self.delivtailsRecord = []
        self.txtFiles = []
        self.tradeConfirmationFullDetails = []
        self.openPositionFullDetails = []
        self.unsettledClosedPositionFullDetails = []
        self.closedPositionFullDetails = []
        self.fundMovementFullDetails = []



    def initXLS(self):
        self.xlsBook = xlrd.open_workbook(self.bookfilename)
        self.xlsSheet = self.xlsBook.sheet_by_name(self.sheet_name)
        self.lastRow = self.xlsSheet.nrows - 1
        self.lastCol = self.xlsSheet.ncols - 1
        
    
    def createDir(self):
        if not os.path.exists(self.dateOfFileName):
            os.makedirs(self.dateOfFileName)
      
    def run(self,account=None):
        self.readXLS(account)
        self.dealCR() #生成汇率表
        self.dealExchangRec() #分离汇总记录
        self.lmeUCP() # 计算LME平仓未到期合约的总盈亏
        self.createDir() #建立输出目录
        self.writeTXT()
        self.createZipFile()
        self.sendMail()
        
    def writeTXT(self):
        self.cusfund() 
        #帐户资金数据，需要特别计算LME的情况 
        #按美元计算的账户市值 客户的基础货币多为港币，但也有指定为美元的，经商量，基金帐户的基础货币为美元，所以不存在问题。
        self.customer() # 根据新文档增加的，手工填写数据
        self.fundchg() 
        self.exchange() 
        self.trddata() 
        self.optdata() 
        #下一步完善期权部分
        self.holddata() 
        # 
        self.liquiddetails() 
        #
        self.holddetails() 
        #LME 平仓未到期合约也填于此
        self.delivtails() 
        #没数据可填
        
    def readXLS(self,account=None):
        if (type(account)==int):
            self.accountList = self.getAccountList(limit = account)
        else:
            self.accountList = self.getAccountList(account = account)
        for (acc,row) in self.accountList:
            logger.info('%s --- %s' % (acc,row))
        logger.info('###########openPositions############')
        self.openPositions = self.getXlsFields(u'Open Position Summary','H')
        logger.info('###########UnsettledClosedPositions############')
        self.unsettledClosedPositions = self.getXlsFields(u'Unsettled Closed Position Summary','O')
        logger.info('###########fundMovement############')
        self.fundMovement = self.getXlsFields(u'Fund Movement','AD')
        logger.info('###########dailyAccountSummary############')
        self.dailyAccountSummary = self.getXlsFields(u'Account Summary','AG')
        logger.info('###########closedPositionSummary############')
        self.closedPositionSummary = self.getXlsFields(u'Closed Position Summary','V')
        logger.info('###########tradeConfirmationSummary############')
        self.tradeConfirmationSummary = self.getXlsFields(u'Trade Confirmation Summary','B')
        logger.info('###########tradeConfirmationFullDetails############')
        self.tradeConfirmationFullDetails = self.getXlsFields(u'Trade Confirmation Full Details','AV')
        logger.info('###########openPositionFullDetails############')
        self.openPositionFullDetails = self.getXlsFields(u'Open Position Full Details','BI')
        logger.info('###########unsettledClosedPositionFullDetails############')
        self.unsettledClosedPositionFullDetails = self.getXlsFields(u'Unsettled Closed Position Full Details','BU')
        logger.info('###########closedPositionFullDetails############')
        self.closedPositionFullDetails = self.getXlsFields(u'Closed Position Full Details','CH')
        logger.info('###########fundMovementFullDetails############')
        self.fundMovementFullDetails = self.getXlsFields(u'Fund Movement Full Details','CR')


    def getAccountList(self,limit=65535,account=None): # 带上参数即可只输出单一指定帐户的数据，　也可以限制只输出多少个帐户的数据（account不指定的情况下）
        #accountList = []
        #r = self.xlsSheet.col_values(0)
        #for f in r: 
        #    mat=re.match(account_pat,str(f.Value))
        #    if (mat is not None):
        #        accountList.append(f)
        #return accountList
        accountList = []
        lt = limit
        if (account):
            lt = 65535
        cs = self.xlsSheet.col_values(0)
        for r in range(2,self.lastRow): 
            mat=re.match(account_pat,str(cs[r]))
            if (mat is not None):
                if (account):
                    if (account == str(cs[r])):
                        accountList.append((cs[r],r))
                else:
                    accountList.append((cs[r],r))
                lt = lt - 1
            if (not lt):
                break
        return accountList
        

    def getXlsFields(self,p_key,C_NotNull=None): #C_NotNull is Col if not allow NULL
        rsss = []
        for (acc,row0) in self.accountList:
            row = row0
            rss = []
            ops = sc[p_key]
            mk = ops.keys()
            while (True):
                rs = {}
                if (C_NotNull):
                    if (self.xlsSheet.cell(row,Utils.col_by_name(C_NotNull)).value):
                        for k in mk:
                            rs[k]=self.xlsSheet.cell(row,Utils.col_by_name(ops[k][0])).value
                        rss.append(rs)
                else:
                    for k in mk:
                        rs[k]=self.xlsSheet.cell(row,Utils.col_by_name(ops[k][0])).value
                    rss.append(rs)
                row = row + 1
                if row > self.lastRow:
                    break
                c = self.xlsSheet.cell(row,0).value
                if (c):
                    break
            if (rss):
                rsss.append((acc,rss))
                logger.debug('%s\n%s' % (acc,rss))
        return rsss

    def getCurrencyField(self,currency=None,product=None,exchange=None): 
        if (currency):
            rt = currency
        else:
            if (product):
                rt = self.getProductCurrence(product)
            else:
                logger.error('getCurrencyField error! currency=%s' % (currency,rt))
                rt = ''
        return rt
        
    '''
    def getCurrencyFieldOLD(self,currency=None,product=None,exchange=None): # 表格要求　区分　ＵＳＤ和 USD-LME
        if (currency):
            rt = currency
        if (product):
            rt = self.getProductCurrence(product)
        else:
            if (exchange=='LME'):
               rt = "USD-LME"
        if (rt == "USD-LME" and currency <> 'USD'):
            logger.error('getCurrencyField error! currency=%s and result=%s' % (currency,rt))
        return rt
    '''

    def getSettlerName(self,p_v):
        m_settler = strip(p_v)
        if settlers.has_key(m_settler):
            rt = settlers[m_settler]
        else:
            rt = ''
            logger.error('SettlerName error! value=%s' % m_settler)
        return rt
    
    def getBankCode(self,p_v):
        m_bank = strip(p_v)
        if banks.has_key(m_bank):
            rt = banks[m_bank][0]
        else:
            rt = '99'
            logger.error('getBankCode error! value=%s' % m_bank)
        return rt
        
    def getSDateField(self,p_date=''):
        m_d = self.dateOfFileName
        m_date = strip(p_date)
        if (m_date):
            m_d = m_date
        if (len(m_d) > 8):  #可能带入的已经是　格式：YYYY-MM-DD
            d = m_d
        else:
            #结算日期   Date    N   格式：YYYY-MM-DD
            d = '-'.join((m_d[:4],m_d[4:6],m_d[6:]))
        return d

    def getTradeRefField(self,p_date):
        if (p_date):
            if (p_date[0] == '*'):
                return p_date[1:]
            else:
                return p_date
        else:
            logger.error('getTradeRefField error! value=None')
            return ''

    def getDescriptionField(self,p_product,p_date):
        m_date = strip(p_date)
        if (not m_date):
            return ''
        if ('-' in m_date):
            t_d = m_date.split('-')
        elif ('/' in m_date):
            t_d = m_date.split('/')
        else:
            dl = len(m_date)
            if ( dl == 6):  #格式：YYYYMM
                t_d = (int(m_date[:4]),int(m_date[4:]),0)
            elif (dl == 8):
                #结算日期   Date    N   格式：YYYYMMDD
                t_d = (int(m_date[:4]),int(m_date[4:6]),int(m_date[6:]))
            else:
                t_d = (0,0,0)
        rt = ''
        t_d_s = (("%04i" % t_d[0])[2:],"%02i" % t_d[1],"%02i" % t_d[2])
        if (self.getProductExchange(p_product)=='LME'):
            rt = t_d_s[0]+t_d_s[1]+t_d_s[2]
        else:
            rt = t_d_s[0]+t_d_s[1]
            #if (t_d[2]):
            #    rt = t_d_s[0]+t_d_s[1]+t_d_s[2]
            #else:
            #    rt = t_d_s[0]+t_d_s[1]
        return self.getProduct(p_product)+rt
        
    def getPromptDateField(self,p_date,p_product=None):
        m_dfstr = "%04i-%02i-%02i"
        m_date = strip(p_date)
        if (not m_date):
            return ''
        if ('-' in m_date):
            t_d = m_date.split('-')
        elif ('/' in m_date):
            t_d = m_date.split('/')
        else:
            dl = len(m_date)
            if ( dl == 6):  #格式：YYYYMM
                t_d = (int(m_date[:4]),int(m_date[4:]),0)
            elif (dl == 8):
                #结算日期   Date    N   格式：YYYYMMDD
                t_d = (int(m_date[:4]),int(m_date[4:6]),int(m_date[6:]))
            else:
                t_d = (0,0,0)
        rt = ''
        if (p_product):
            if (self.getProductExchange(p_product)=='LME'):
                rt = m_dfstr % t_d
            else:
                if (t_d[2]):
                    rt = m_dfstr % t_d
                else:
                    ld = calendar.monthrange(t_d[0],t_d[1])
                    ldw = (ld[1]-1 + ld[0]) % 7
                    if (ldw > 4):
                        rt = m_dfstr % (t_d[0],t_d[1],ld[1]-(ldw - 4))
                    else:
                        rt = m_dfstr % (t_d[0],t_d[1],ld[1])
        else:
            rt = m_dfstr % t_d
        return rt


    def dealCR(self): # 计算汇率表
        self.currencyRate={}
        for (c,rs) in self.dailyAccountSummary:
            for r in rs:
                if (not self.currencyRate.has_key(strip(r['Currency']))):
                    self.currencyRate[strip(r['Currency'])] = r['Exchange Rate']
    
    def splitDescription(self,p_v): # 新的ＸＬＳ文件定义了Description字段，把合约和到期日合并在一起，需要分开
        # 错误修改，　需要考虑　期权的情况，格式为　ESO / 201409 / Put 1,870.00
        (prod,prompt)=('','')
        try :
            rt = p_v.split('/')
            if (len(rt)==2):
                (prod,prompt)= rt
            if (len(rt)==3):
                (prod,prompt,price)= rt
                (m_type,m_price) = strip(price).split(' ')
        except Exception,e:
            logger.error('splitDescription error! value=%s' % p_v)
        return (strip(prod),strip(prompt))
    
    def spliteDateTime(self,p_v): # 新的ＸＬＳ文件定义了TradeDate字段，把日期和时间合并在一起，需要分开
        (m_date,m_time) = ('0000-00-00','00:00:00')
        if (type(p_v) == float): # 发现新表会出现为linux时间的数据，要转换
            m_v_t = xlrd.xldate_as_tuple(p_v, 0)
            m_date = "%04i-%02i-%02i" % m_v_t[:3]
            m_time = "%02i:%02i:%02i" % m_v_t[3:]
        else:
            m_v = strip(p_v)
            try :
                (m_date,m_time)=m_v.split(' ')
            except Exception,e:
                logger.error('spliteDateTime error! value=%s, msg:%s' % (p_v,e))
        return (m_date,m_time)

    def getTimeUTC8(self,p_v): #得到北京时间
        return p_v

    def dealExchangRec(self):
        m_fmf = self.fundMovementFullDetails
        ret=[]
        for (c,rs) in m_fmf:
            m_er = []
            for r in rs:
                if ('CCY TRANSFER' == strip(r['Transaction Type'])):
                    m_er.append(r)
            for r2 in m_er:
                rs.remove(r2)
            ret.append((c,m_er)) 
        self.exchangeRecord = ret
            
    def lmeUCP(self):
        ucp = self.unsettledClosedPositions
        ret={}
        for (c,rs) in ucp:
            ucppl = {}
            for r in rs:
                if (ucppl.has_key(strip(r['Product']))):
                    ucppl[strip(r['Product'])]= ucppl[strip(r['Product'])] + r['No of Lots(buy)'] * (r['Average Trading price(sell)'] - r['Average Trading price(buy)'])
                else:
                    ucppl[strip(r['Product'])]= r['No of Lots(buy)'] * (r['Average Trading price(sell)'] - r['Average Trading price(buy)'])
            ret[c] = ucppl
        self.unsettledClosedPositionPL = ret
    
    def getlmeUCPPL(self,acc): #需要完善，确保只有ＬＭＥ的数据
        if (self.unsettledClosedPositionPL.has_key(acc)):
            pls = self.unsettledClosedPositionPL[acc]
            return sum(pls.values())
        else:
            return 0.00
    
    def getFieldString(self,value,width=4,precision=2):
        if (isinstance(value,str)):
            return strip(value)
        if (isinstance(value,int)):
            return "%*.*f" % (width,precision,value)
        if (isinstance(value,float)):
            return "%*.*f" % (width,precision,value)
        return strip(str(value))

    def getProductMulti(self,prod):
        m_prod = strip(str(prod))
        if (pil_name.has_key(m_prod)):
            return dict(zip(pil_vtitle,pil_name[m_prod]))['Multiplier']
        else:
            logger.error('ticker_name has no key:%s' % m_prod)
            return 1
   
    def getProduct(self,prod):
        m_prod = strip(str(prod))
        if (pil_name.has_key(m_prod)):
            rt = dict(zip(pil_vtitle,pil_name[m_prod]))['CMF_CODE']
            if (not rt):
                logger.error('Products has None key, Please correct it!:%s' % m_prod)
                rt = dict(zip(pil_vtitle,pil_name[m_prod]))['Underlying']
            return rt
        else:
            logger.error('Products has no key:%s' % m_prod)
            return ''


    def getProductCurrence(self,prod):
        m_prod = strip(str(prod))
        if (pil_name.has_key(m_prod)):
            #if (pil_name[m_prod][1] == 'LME'):
            #    return "USD-LME"
            #else:
            #    return pil_name[m_prod][3]
            return pil_name[m_prod][3]
        else:
            logger.error('ProductsCurrence has no key:%s' % m_prod)
            return ''
    '''
    def getProductCurrenceOLD(self,prod):
        m_prod = strip(str(prod))
        if (pil_name.has_key(m_prod)):
            if (pil_name[m_prod][1] == 'LME'):
                return "USD-LME"
            else:
                return pil_name[m_prod][3]
        else:
            logger.error('ProductsCurrence has no key:%s' % m_prod)
            return ''
    '''
    
    def getProductExchange(self,prod):
        m_prod = strip(str(prod))
        if (pil_name.has_key(m_prod)):
            return self.getExchName(pil_name[m_prod][1])
        else:
            logger.error('ProductExchange has no key:%s' % m_prod)
            return 'UN'
 
        
    def getExchName(self,exch):
        m_exch = strip(str(exch))
        if (exch_name.has_key(m_exch)):
            return exch_name[m_exch][0]
        else:
            logger.error('Exchanges has no key:%s' % m_exch)
            return ''
    
    def getFileName(self,fnfmt):
        fn = '%s%s_f%s.txt' % (self.flagCompany,fnfmt,self.dateOfFileName)
        ffn = os.path.join(self.dateOfFileName,fn)
        return ffn 
        
    def customer(self): #客户基本资料数据文件
        #Fees & Levy 不太清楚填哪
        fn = self.getFileName('customer')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.info('begin deal customer')
        for (c,rs) in self.dailyAccountSummary:
            rs_idx = 1
            rs_cc = len(rs)
            for r in rs:
                fields=[]
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                fields.append('招商基金管理有限公司') #客户名称  char(60)    N
                fields.append('') #身份证  char(40)    Y 自然人客户的身份证号码。如果是法人客户，此字段为空
                fields.append(c) #客户内部资金账户  char(18)    N
                fields.append('') #客户统一开户编码 char(8)     N   客户在统一开户系统中的编码
                fields.append('深圳') #所在地  char(40)    N
                fields.append('深圳深南大道7088号') #通讯地址  char(100)    N
                fields.append('518040') #邮政编码  char(6)    N
                fields.append('008675583073042') #联系电话  char(60)    N
                fields.append('a') #开户和销户标志   char    N 已开户的客户 a，当日销户的客户 d
                fields.append('1') #客户类型  char   N 0：自然人；1：法人；9：其它
                fields.append('2014-09-01') #客户开户日期 Date    N 客户开户并生效的实际日期，（并非预开户日期），格式：YYYY-MM-DD
                fields.append('cmfchina0001') #组织机构代码证号  char(40)    Y 法人客户的组织机构代码证号。如果是自然人客户，此字段为空
                fields.append('I20140001') #营业执照号  char(40)    Y 法人客户的营业执照号。如果是自然人客户，此字段为空
                fields.append('王立立') #开户授权人名称  char(40)    Y 开户授权人的姓名。如果系统中暂时无此字段，可以为空
                fields.append('') #开户授权人身份证  char(40)    Y 开户授权人的身份证号码。如果系统中暂时无此字段，可以为空
                fields.append('a') #客户名称  char(40)    N a表示活跃客户，d表示休眠客户
                #客户基本资料里面各主要字段（如：客户名称，所在地，通讯地址，联系电话）如果内容出现半角@字符，一律替换为 &at;
                #2012-01-04@胜利公司@@0001@12345678@上海市@上海市浦电路700号@200122@02168400901@a@N@1@2011-11-10@710685288@985321456221@王小二@110107196910275011@a
                fs = [self.getFieldString(i) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
                rs_idx= rs_idx + 1
        f.close()
        logger.info('end deal cusfund')
        
    def cusfund(self): #客户基本资金数据文件
        fn = self.getFileName('cusfund')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.info('begin deal cusfund')
        for (c,rs) in self.dailyAccountSummary:
            rs_idx = 1
            rs_cc = len(rs)
            for r in rs:
                fields=[]
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                fields.append(c) #客户内部资金账户  char(18)    N
                fields.append('') #客户统一开户编码 char(8) Y   客户在统一开户系统中的编码
                m_isUSD = r['Currency']=='USD'
                fields.append(r['Currency']) #币种    Char(7) N   按照ISO 4217; 例如:USD; JPY，如为LME美元则填写USD-LME
                if (rs_idx < rs_cc):
                    ccc = 'N'
                else:
                    ccc = 'Y'
                fields.append(ccc) #是否为基准货币 Char(1) N   Y- 该条数据为折合成基准货币的资金数据 N-该条数据为分币种资金数据（含美元）
                fields.append(r['Open Balance']) #上日结存（逐笔对冲）    Number(14,2)    Y
                fee = -(r['Commission']+r['Fees & Levy'])
                fields.append(fee) #手续费 Number(14,2)    N   包括当日所有产生佣金、手续费、协会费用等
                fields.append('0.00') #汇兑手续费    Number(14,2)    N   当日换汇所产生的手续费
                fields.append(r['Trading P/(L)']) #平仓盈亏（逐笔对冲）   Number(14,2)    N   逐笔对冲下的平仓盈亏包括当日所有平仓合约及LME到期合约的盈亏
                fields.append(r['Deposit / (Withdrawal)']) #出入金 Number(14,2)    N   当日资金出入、资金调整总额及交割货款
                fields.append('0.00') #权利金  Number(14,2)    N
                fields.append(r['Closing Balance']) #当日结存（逐笔对冲） Number(14,2)    N   期末结存=期初结存-手续费-汇兑手续费+平仓盈亏+出入金+权利金
                m_pl = r['Gross Floating P/(L)']
                if (m_isUSD):
                    m_pl = m_pl - self.getlmeUCPPL(c)
                fields.append(m_pl) #浮动盈亏（逐笔对冲）不包含LME未到账盈亏  Number(14,2)    N   逐笔对冲方式计算的浮动盈亏  浮动盈亏=∑（结算价-开仓价）*持仓量*单位
                #此处要判断是否美元帐户，是的话再看有没有平仓未到期合约
                m_eq = r['Equity']
                if (m_isUSD):
                    m_eq = m_eq - self.getlmeUCPPL(c)
                fields.append(m_eq) #期末权益（不含期权、LME） Number(14,2)    N   期末权益=期末结存+浮动盈亏
                fields.append('0.00') #期权市值 Number(14,2)    N   期权市值=∑权利金结算价*持仓量*单位
                m_wdzyy = 0.00
                if (m_isUSD):
                    m_wdzyy = self.getlmeUCPPL(c)
                fields.append(m_wdzyy) #LME的未到帐盈亏（LME变动保证金） Number(14,2)    N   LME的未到帐盈亏=∑(结算价-开仓价)*持仓量*单位
                fields.append('0.00') #LME贴现后的变动保证金 Number(14,2)    N   LME贴现后的变动保证金值 
                fields.append(-r['Initial Margin']) #客户初始保证金    Number(14,2)    N
                fields.append('0.00') #质押资金 Number(14,2)    N   包括所有质押品的市值
                fields.append(-r['Maintenance Margin']) #客户维持保证金    Number(14,2)    N   一般情况是初始保证金的80%
                fields.append(-r['Initial Margin']) #上手维持保证金  Number(14,2)    Y   待定
                m_kyzj = m_eq + r['Initial Margin']
                if (m_isUSD):
                    m_kyzj = m_kyzj + self.getlmeUCPPL(c)
                fields.append(m_kyzj) #可用资金 Number(14,2)    N   可用资金=期末权益-初始保证金+ LME的未到帐盈亏
                m_mc = 0.00 
                if (r['Margin (Call)/Excess'] < 0):
                    m_mc = r['Margin (Call)/Excess']
                fields.append(m_mc) #需追加保证金 Number(14,2)    N   追加资金=客户维持保证金-期末权益
                fields.append('0.00') #风险度  Number(14,2)    Y   风险度=客户维持保证金/账户市值 去掉百分号，如46表示46％
        
                fields.append(r['Equity']) #账户市值    Number(14,2)    N   账户市值=期末权益+期权市值+LME变动保证金
                m_ec = r['Exchange Rate']
                #if (not m_isUSD):
                #    #m_ec = r['Exchange Rate'] #如何计算美元的汇率？
                #    m_ec = 0.00
                fields.append(m_ec) #汇率 Number(14,8)    N   各币种转换成美元的汇率, 同王立立商量了，填成换基础货币的汇率
                fields.append(r['Equity']) #按美元计算的账户市值  Number(14,2)    N   各币种按汇率计算后的账户市值。  同王立立商量了，填成换基础货币的汇率
                fields.append('0.00') #其它特殊资金   Number(14,2)    Y   例如：交割货款等其他引起资金变动的金额合计值
                # 2012-01-04@0001@12345678@JPY@N@2000000.00@0.00@0.00@0.00@0.00@0.00@2000000.00@0.00@200000.00@0.00@0.00@0.00@0.00@0.00@0.00@0.00@2000000.00@0.00@0@2000000.00@0.0121@24200@0.00
                fs = [self.getFieldString(i) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
                rs_idx= rs_idx + 1
        f.close()
        logger.info('end deal cusfund')
        
    def fundchg(self): #客户出入金记录文件
        # 差一个汇畜率不知填在哪
        fn = self.getFileName('fundchg')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.info('begin deal fundchg')
        for (c,rs) in self.fundMovementFullDetails:
            for r in rs:
                fields=[]
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                fields.append(c) #客户内部资金账户  char(18)    N
                fields.append('') #客户统一开户编码 char(8) Y   客户在统一开户系统中的编码
                fields.append('C') #流水描述    Char(1) N   出入金Debit/Credit-C；汇兑Exchange-E；资金调整Adjustment-A；交割货款Delivery-D
                fields.append(r['Withdrawal']) #入金   Number(14,2)    N
                fields.append(r['Deposit']) #出金   Number(14,2)    N
                fields.append(r['Currency']) #币种    char(7) N   按照ISO 4217; 例如:USD; JPY，如为LME美元则填写USD-LME
                fields.append('99') #客户期货结算账户银行统一标识 char(2) Y   具体见数据字典银行编号部分
                fields.append('') #客户期货结算账户 char(22)    Y   人民币或外币结算账户
                fields.append('W') #客户本外币账户标识   char    Y   人民币账户-L 境内外币期货结算账户-W
                fields.append(self.getBankCode(r['Bank'])) #公司保证金专用账户银行统一标识    char(2) Y   具体见数据字典银行编号部分
                fields.append('') #期货公司在境内结算银行开立境内保证金专户 char(22)    Y   人民币保证金专户及外汇专户
                fields.append('F') #公司本外币账户标识   char    Y   人民币账户-L 境外外币保证金账户-F 境内外币保证金账户-J
                #2012-01-04@0001@12345678@C@0.00@500000.00@USD@03@1234567890123@W@03@8888888888888@J
                fs = [self.getFieldString(i) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
        f.close()
        logger.info('end deal fundchg')

    '''
    def fundchgOLD(self): #客户出入金记录文件 原三个字段的项，弃用
        # 差一个汇畜率不知填在哪
        fn = self.getFileName('fundchg')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal fundchgOLD')
        for (c,rs) in self.fundMovement:
            for r in rs:
                fields=[]
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                fields.append(c) #客户内部资金账户  char(18)    N
                fields.append('0.00') #客户统一开户编码 char(8) Y   客户在统一开户系统中的编码
                fields.append('C') #流水描述    Char(1) N   出入金Debit/Credit-C；汇兑Exchange-E；资金调整Adjustment-A；交割货款Delivery-D
                if (r['Net Fund Movement'] > 0):
                    fields.append(r['Net Fund Movement']) #入金   Number(14,2)    N
                    fields.append('0.00') #出金   Number(14,2)    N
                else:
                    fields.append('') #入金   Number(14,2)    N
                    fields.append(r['Net Fund Movement']) #出金   Number(14,2)    N
                fields.append(r['Currency']) #币种    char(7) N   按照ISO 4217; 例如:USD; JPY，如为LME美元则填写USD-LME
                fields.append('99') #客户期货结算账户银行统一标识 char(2) Y   具体见数据字典银行编号部分
                fields.append('0.00') #客户期货结算账户 char(22)    Y   人民币或外币结算账户
                fields.append('W') #客户本外币账户标识   char    Y   人民币账户-L 境内外币期货结算账户-W
                fields.append('99') #公司保证金专用账户银行统一标识    char(2) Y   具体见数据字典银行编号部分
                fields.append('0.00') #期货公司在境内结算银行开立境内保证金专户 char(22)    Y   人民币保证金专户及外汇专户
                fields.append('F') #公司本外币账户标识   char    Y   人民币账户-L 境外外币保证金账户-F 境内外币保证金账户-J
                #2012-01-04@0001@12345678@C@0.00@500000.00@USD@03@1234567890123@W@03@8888888888888@J
                fs = [self.getFieldString(i) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
        f.close()
        logger.info('end deal fundchgOLD')
    '''

    def exchange(self): #客户汇兑明细文件
        fn = self.getFileName('exchange')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal exchange')
        for (c,rs) in self.exchangeRecord:
            m_DepositeRec = None
            m_Withdrawal = None
            m_isTWO = False
            for r in rs:
                if (not m_isTWO):   # 汇兑是一对对出现的， 所以两条记录要一块处理， 这也是非常复杂而容易出错的地方
                    if (r['Withdrawal'] > 0):
                        m_Withdrawal = copy.deepcopy(r)
                    else:
                        m_DepositeRec = copy.deepcopy(r)
                    m_isTWO = not m_isTWO
                    continue
                else:
                    if (r['Withdrawal'] > 0):
                        m_Withdrawal = copy.deepcopy(r)
                    else:
                        m_DepositeRec = copy.deepcopy(r)
                    m_isTWO = not m_isTWO       
                        
                fields=[]
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                fields.append(c) #   客户内部资金账户   Char(18)    N            
                fields.append('') #   客户统一开户编码   char(8) Y   客户在统一开户系统中的编码      
                fields.append(self.getSDateField()) #   成交日期  Date    N   成交日期， 格式(form)：YYYY-MM-DD           
                fields.append(m_DepositeRec['Deposit']) #   兑换为目标币种后的金额   Number(14,2)    N         
                fields.append(m_DepositeRec['Currency']) #   目标币种  Char(3) N   按照ISO 4217; 例如:USD; JPY         
                fields.append(m_Withdrawal['Currency']) #   原币种   Char(3) N   按照ISO 4217; 例如:USD; JPY     
                m_rate =   m_Withdrawal['Exchange Rate']/m_DepositeRec['Exchange Rate']   
                fields.append("%14.8f" % m_rate) #   当笔兑换汇率    Number(14,8)    N   原币种对目标币种的汇率。 
                fields.append("%14.2f" % (m_Withdrawal['Withdrawal']*m_rate-m_DepositeRec['Deposit'])) #   汇兑手续费 Number(14,2)    N   当日换汇所产生的手续费，币种同目标币种         
                #2012-01-04@0001@12345678@2012-01-03@2000000.00@USD@JPY@0.0121@10.00        
                fs = [strip(str(i)) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
        f.close()
        logger.debug('end deal exchange')
        

    def trddata(self): #成交明细文件
        #   成交额 填写的是买成交额
        fn = self.getFileName('trddata')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal trddata')
        for (c,rs) in self.tradeConfirmationFullDetails:
            for r in rs:
                fields=[]
                (m_prod,m_prompt) = self.splitDescription(r['Description'])
                (m_date,m_time) = self.spliteDateTime(r['Trade Date'])
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                fields.append(self.getSDateField(m_date)) #   成交日期    Date    N   格式：YYYY-MM-DD
                fields.append(self.getPromptDateField(m_prompt,m_prod)) #   到期日 Date    Y   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                fields.append(c) #   客户内部资金账户   char(18)    N
                fields.append('') #   客户统一开户编码  char(8) Y   客户在统一开户系统中的编码
                #fields.append(r['Trade Ref.']) #   成交流水号 Char(16)    N   公司交易系统发布的成交序列号
                fields.append(self.getTradeRefField(r['Trade Ref.']))
                fields.append(self.getProductCurrence(m_prod)) #   币种 Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                fields.append(self.getExchName(r['Exchange'])) #   交易所  Char(5) N   具体见数据字典交易所名称部分
                fields.append(self.getProduct(m_prod)) #   品种 Char(20)    N   具体见数据字典品种名称部分 
                fields.append(self.getDescriptionField(m_prod , m_prompt)) #   合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                fields.append(self.getOpenOrClose(r['Open or Close'])) #   开平标志 char    N   开仓-O，平仓-L
                fields.append(r['No of Lots(buy)']) #   买成交量    Number(10)  N   单位：手
                fields.append(r['No of Lots(sell)']) #   卖成交量  Number(10)  N   单位：手
                fields.append(r['Trade Price/Premium']) #   成交价  Number(14,7)    N   
                m_cje = ((r['No of Lots(buy)']+r['No of Lots(sell)']) * r['Trade Price/Premium']) * self.getProductMulti(m_prod)
                fields.append(m_cje) #   成交额    Number(14,2)    N
                fields.append(m_time) #   成交时间  char(10)    N   格式：hh:mm:ss 北京时间（24小时制）
                fields.append(r['Commission']+r['Market Charges and Fees']) #   手续费   Number(14,2)    N   所有的手续费用，包括交易所、协会、上手清算机构的手续费
                #可以直接加总
                #    因为只有香港市场的'Market Charges and Fees'非0
                fields.append('0.00') #   权利金   Number(14,2)    Y   期权的成交价*成交数量
                fields.append('') #   期权类型  Char(1) Y   ”C” –看涨期权,”P”-看跌期权
                fields.append('0.00') #   执行价   Number(14,7)    Y
                fields.append(self.getSettlerName(r['Broker'])) #   上手清算机构代码  char(2) Y
                fields.append('') #   境内期货公司的账号 Char(10)    Y   即境内期货公司在上手清算公司的账号
                fields.append(r['Currency1']) #  交易费用币种 Char(7) N   按照 ISO 4217;   根据新20140922号的邮件回复，要求增加这一字段
                #2012-01-04@2012-01-03@2012-03-07@0001@12345678@2011050400000134@USD@CME@CA@ CA1203@L@0@1@4.24@106.04@15:04:010@15.00@@@@jp@12345
                fs = [strip(str(i)) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
        f.close()
        logger.debug('end deal trddata')

    def getOpenOrClose(self,p_v):
        m_v = strip(p_v)
        if (m_v == '*'):
            return  'O'
        elif (m_v ==  '#'):
            return 'L'
        else:
            return 'O'

    def trddataSummary(self): #成交明细文件 合并形式的，　弃用
        #   成交额 填写的是买成交额 不对
        fn = self.getFileName('trddata')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal trddataSummary')
        for (c,rs) in self.tradeConfirmationSummary:
            for r in rs:
                if (r['No of Lots(buy)'] > 0):
                    fields=[]
                    fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                    fields.append(self.getSDateField()) #   成交日期    Date    N   格式：YYYY-MM-DD
                    m_promptdate = '00000000'  #r['Prompt Date']
                    fields.append(self.getSDateField(m_promptdate)) #   到期日 Date    Y   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                    fields.append(c) #   客户内部资金账户   char(18)    N
                    fields.append('') #   客户统一开户编码  char(8) Y   客户在统一开户系统中的编码
                    fields.append('') #   成交流水号 Char(16)    N   公司交易系统发布的成交序列号
                    fields.append(self.getProductCurrence(r['Product'])) #   币种 Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                    fields.append(self.getExchName(r['Exchange'])) #   交易所  Char(5) N   具体见数据字典交易所名称部分
                    fields.append(self.getProduct(r['Product'])) #   品种 Char(20)    N   具体见数据字典品种名称部分 
                    fields.append(self.getProduct(r['Product'])+ self.getSDateField(m_promptdate)) #   合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                    fields.append('O') #   开平标志 char    N   开仓-O，平仓-L
                    fields.append(r['No of Lots(buy)']) #   买成交量    Number(10)  N   单位：手
                    fields.append('0.00') #   卖成交量  Number(10)  N   单位：手
                    fields.append(r['Average Trading price(buy)']) #   成交价  Number(14,7)    N   
                    m_cje = r['No of Lots(buy)']*r['Average Trading price(buy)']* self.getProductMulti(r['Product'])
                    fields.append(m_cje) #   成交额    Number(14,2)    N
                    fields.append('00:00:00') #   成交时间  char(10)    N   格式：hh:mm:ss 北京时间（24小时制）
                    fields.append('0.00') #   手续费   Number(14,2)    N   所有的手续费用，包括交易所、协会、上手清算机构的手续费
                    fields.append('') #   权利金   Number(14,2)    Y   期权的成交价*成交数量
                    fields.append('') #   期权类型  Char(1) Y   ”C” –看涨期权,”P”-看跌期权
                    fields.append('') #   执行价   Number(14,7)    Y
                    fields.append('') #   上手清算机构代码  char(2) Y
                    fields.append('') #   境内期货公司的账号 Char(10)    Y   即境内期货公司在上手清算公司的账号
                    #2012-01-04@2012-01-03@2012-03-07@0001@12345678@2011050400000134@USD@CME@CA@ CA1203@L@0@1@4.24@106.04@15:04:010@15.00@@@@jp@12345
                    fs = [strip(str(i)) for i in fields]
                    ln = '@'.join(fs)
                    f.write(ln+'\n')
                if (r['No of Lots(sell)'] > 0):
                    fields=[]
                    fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                    fields.append(self.getSDateField()) #   成交日期    Date    N   格式：YYYY-MM-DD
                    m_promptdate = '00000000'  #r['Prompt Date']
                    fields.append(self.getSDateField(m_promptdate)) #   到期日 Date    Y   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                    fields.append(c) #   客户内部资金账户   char(18)    N
                    fields.append('') #   客户统一开户编码  char(8) Y   客户在统一开户系统中的编码
                    fields.append('') #   成交流水号 Char(16)    N   公司交易系统发布的成交序列号
                    fields.append(self.getProductCurrence(r['Product'])) #   币种 Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                    fields.append(self.getExchName(r['Exchange'])) #   交易所  Char(5) N   具体见数据字典交易所名称部分
                    fields.append(self.getProduct(r['Product'])) #   品种 Char(20)    N   具体见数据字典品种名称部分 
                    fields.append(self.getProduct(r['Product'])+ self.getSDateField(m_promptdate)) #   合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                    fields.append('L') #   开平标志 char    N   开仓-O，平仓-L
                    fields.append('0.00') #   买成交量  Number(10)  N   单位：手
                    fields.append(r['No of Lots(sell)']) #   卖成交量   Number(10)  N   单位：手
                    fields.append(r['Average Trading price(sell)']) #   成交价 Number(14,7)    N   
                    m_cje = r['No of Lots(sell)']*r['Average Trading price(sell)']* self.getProductMulti(r['Product'])
                    fields.append(m_cje) #   成交额    Number(14,2)    N
                    fields.append('00:00:00') #   成交时间  char(10)    N   格式：hh:mm:ss 北京时间（24小时制）
                    fields.append('0.00') #   手续费   Number(14,2)    N   所有的手续费用，包括交易所、协会、上手清算机构的手续费
                    fields.append('') #   权利金   Number(14,2)    Y   期权的成交价*成交数量
                    fields.append('') #   期权类型  Char(1) Y   ”C” –看涨期权,”P”-看跌期权
                    fields.append('') #   执行价   Number(14,7)    Y
                    fields.append('') #   上手清算机构代码  char(2) Y
                    fields.append('') #   境内期货公司的账号 Char(10)    Y   即境内期货公司在上手清算公司的账号
                    #2012-01-04@2012-01-03@2012-03-07@0001@12345678@2011050400000134@USD@CME@CA@ CA1203@L@0@1@4.24@106.04@15:04:010@15.00@@@@jp@12345
                    fs = [strip(str(i)) for i in fields]
                    ln = '@'.join(fs)
                    f.write(ln+'\n')
        f.close()
        logger.debug('end deal trddataSummary')

    def optdata(self): #期权行权明细文件
        fn = self.getFileName('optdata')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal optdata')
        m_tcfds = self.tradeConfirmationFullDetails
        m_otcfds = [] # 此处还需要过滤只是期权的交易记录
        for (c,rs) in m_otcfds:
            for r in rs:
                fields=[]
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                fields.append('') #成交日期 Date    N   成交日期， 格式(form)：YYYY-MM-DD
                fields.append('') #到期日  Date    N   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                fields.append(c) #客户内部资金账户  char(18)    N
                fields.append(c) #客户统一开户编码  char(8) N   客户在统一开户系统中的编码
                fields.append('') #成交流水号    Char(16)    N   公司交易系统发布的成交序列号
                fields.append('') #币种   Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                fields.append('') #交易所  Char(5) N   具体见数据字典交易所名称部分
                fields.append('') #品种   Char(20)    N   具体见数据字典品种名称部分
                fields.append('') #合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                fields.append('') #买成交量 Number(10)  N
                fields.append('') #卖成交量 Number(10)  N
                fields.append('') #成交时间 Char(10)    N   格式：hh:mm:ss 北京时间（24小时制）
                fields.append('') #手续费  Number(14,2)    N   所有的手续费用，包括交易所、协会、上手清算机构的手续费
                fields.append('') #权利金单价    Number(14,7)    N   期权开仓时的权利金单价
                fields.append('') #期权类型 Char(1) N   ”C” –看涨期权,”P”-看跌期权
                fields.append('') #执行价  Number(14,7)    N
                fields.append('') #上手清算机构代码 char(2) Y
                fields.append('') #境内期货公司的账号    Char(10)    Y   即境内期货公司在上手清算公司的账号
                #2012-01-04@2012-01-03@2012-03-07@0001@12345678@2011010200000666@USD@CBOT@CO@CO1203@450@0@22:05:012@-1831.50@-45000.00@P@5.50@jp@12345
                fs = [strip(str(i)) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
        f.close()
        logger.debug('end deal optdata')

    def dealOpenPositionSummary(self,p_openPositionFullDetails):
        rsss = []
        for (c,rs) in p_openPositionFullDetails:
            rss = []
            rt0 = []
            rt1 = {}
            for r in rs:
                (m_prod,m_prompt) = self.splitDescription(r['Description'])
                (m_date,m_time) = self.spliteDateTime(r['Trade Date'])
                rs0 = {}
                if (r['No of Lots(buy)'] > 0):
                    m_flag = 'buy' 
                else:
                    m_flag = 'sell' 
                if ((m_prod,m_prompt) not in rt0):
                    rt0.append((m_prod,m_prompt))
                if (rt1.has_key((m_prod,m_prompt,m_flag))):
                    rs0 = rt1[(m_prod,m_prompt,m_flag)] 
                rs0['Prompt Date'] = m_prompt
                rs0['Product'] = m_prod
                rs0['Exchange'] = r['Exchange']
                nn = 0.0
                nnn0 = 0.0
                nnn = 0.0
                if (m_flag == 'buy'):
                    if (rs0.has_key('No of Lots(buy)')):
                        nn = r['No of Lots(buy)']
                        nnn0 = rs0['No of Lots(buy)']
                        nnn = nnn0 + r['No of Lots(buy)']
                    else:
                        nn = r['No of Lots(buy)']
                        nnn = nn
                    rs0['No of Lots(buy)'] = nnn
                    rs0['No of Lots(sell)'] = 0.0
                else:
                    if (rs0.has_key('No of Lots(sell)')):
                        nn = r['No of Lots(sell)']
                        nnn0 = rs0['No of Lots(sell)'] 
                        nnn = nnn0 + r['No of Lots(sell)']
                    else:
                        nn = r['No of Lots(sell)']
                        nnn = nn
                    rs0['No of Lots(sell)'] = nnn
                    rs0['No of Lots(buy)'] = 0.0


                if (rs0.has_key('Gross Floating P/(L)')):
                    rs0['Gross Floating P/(L)'] = rs0['Gross Floating P/(L)'] + r['Gross Floating P/(L)']
                else:
                    rs0['Gross Floating P/(L)'] = r['Gross Floating P/(L)']

                if (rs0.has_key('Trade Price/Premium')):
                    rs0['Trade Price/Premium'] = (rs0['Trade Price/Premium'] * nnn0 + r['Trade Price/Premium'] * nn)/nnn
                else:
                    rs0['Trade Price/Premium'] = r['Trade Price/Premium']
                
                rs0['Closing Price'] = r['Closing Price']
                rt1[(m_prod,m_prompt,m_flag)] = rs0
                    
            for (m_prod,m_prompt) in rt0:
                if rt1.has_key((m_prod,m_prompt,'buy')):
                    rss.append(rt1[(m_prod,m_prompt,'buy')])  
                if rt1.has_key((m_prod,m_prompt,'sell')):
                    rss.append(rt1[(m_prod,m_prompt,'sell')])  
            rsss.append((c,rss))
        return rsss

    def holddata(self): #持仓数据文件
        # 今结算价 填写的是卖平均价
        fn = self.getFileName('holddata')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal holddata')
        m_openPositions = self.dealOpenPositionSummary(self.openPositionFullDetails)
        for (c,rs) in m_openPositions:
            for r in rs:
                    fields=[]
                    fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                    fields.append(self.getPromptDateField(r['Prompt Date'],r['Product'])) #到期日    Date    N   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                    fields.append(c) #客户内部资金账户  char(18)    N
                    fields.append('') #客户统一开户编码 char(8) N   客户在统一开户系统中的编码
                    fields.append(self.getProductCurrence(r['Product'])) #币种    Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                    fields.append(self.getExchName(r['Exchange'])) #交易所 Char(5) N   具体见数据字典交易所名称部分
                    fields.append(self.getProduct(r['Product'])) #品种    Char(20)    N   具体见数据字典品种名称部分
                    fields.append(self.getDescriptionField(r['Product'],r['Prompt Date'])) #合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                    fields.append(r['No of Lots(buy)']) #买持仓量   Number(10)  N
                    fields.append(r['No of Lots(sell)']) #卖持仓量 Number(10)  N
                    fields.append(r['Gross Floating P/(L)']) #持仓盈亏(逐笔对冲)   Number(14,2)    N
                    #ccjj = r['No of Lots(buy)']*r['Average Trading price(buy)']+r['No of Lots(sell)']*r['Average Trading price(sell)']
                    fields.append(r['Trade Price/Premium']) #持仓均价    Number(14,7)    Y
                    fields.append(r['Closing Price']) #今结算价 Number(14,7)    N
                    fields.append('') #期权市值 Number(14,2)    Y
                    fields.append('') #期权类型 Char(1) Y   ”C” –看涨期权,”P”-看跌期权
                    fields.append('') #执行价  Number(14,7)    Y
                    fields.append('') #上手清算机构代码 char(2) Y
                    fields.append('') #境内期货公司的账号    Char(10)    Y   即境内期货公司在上手清算公司的账号
                    #2012-01-04@2012-03-07@0001@12345678@USD@CBOT@BO@BO1203@ 0@1@168.00@51.79@51.52@@@@jp@12345
                    fs = [strip(str(i)) for i in fields]
                    ln = '@'.join(fs)
                    f.write(ln+'\n')
        f.close()
        logger.debug('end deal holddata')

    def holddataSummary(self): #持仓数据文件 弃用
        # 今结算价 填写的是卖平均价
        fn = self.getFileName('holddata')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal holddataSummary')
        for (c,rs) in self.openPositions:
            for r in rs:
                if (r['No of Lots(buy)']>0):
                    fields=[]
                    fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                    fields.append(self.getSDateField(r['Prompt Date'])) #到期日    Date    N   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                    fields.append(c) #客户内部资金账户  char(18)    N
                    fields.append('') #客户统一开户编码 char(8) N   客户在统一开户系统中的编码
                    fields.append(self.getProductCurrence(r['Product'])) #币种    Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                    fields.append(self.getExchName(r['Exchange'])) #交易所 Char(5) N   具体见数据字典交易所名称部分
                    fields.append(self.getProduct(r['Product'])) #品种    Char(20)    N   具体见数据字典品种名称部分
                    fields.append(self.getProduct(r['Product'])+r['Prompt Date']) #合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                    fields.append(r['No of Lots(buy)']) #买持仓量   Number(10)  N
                    fields.append('0.00') #卖持仓量 Number(10)  N
                    fields.append('0.00') #持仓盈亏(逐笔对冲)   Number(14,2)    N
                    #ccjj = r['No of Lots(buy)']*r['Average Trading price(buy)']+r['No of Lots(sell)']*r['Average Trading price(sell)']
                    fields.append(r['Average Trading price(buy)']) #持仓均价    Number(14,7)    Y
                    fields.append('0.00') #今结算价 Number(14,7)    N
                    fields.append('') #期权市值 Number(14,2)    Y
                    fields.append('') #期权类型 Char(1) Y   ”C” –看涨期权,”P”-看跌期权
                    fields.append('') #执行价  Number(14,7)    Y
                    fields.append('') #上手清算机构代码 char(2) Y
                    fields.append('') #境内期货公司的账号    Char(10)    Y   即境内期货公司在上手清算公司的账号
                    #2012-01-04@2012-03-07@0001@12345678@USD@CBOT@BO@BO1203@ 0@1@168.00@51.79@51.52@@@@jp@12345
                    fs = [strip(str(i)) for i in fields]
                    ln = '@'.join(fs)
                    f.write(ln+'\n')
                if (r['No of Lots(sell)']>0):
                    fields=[]
                    fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                    fields.append(self.getSDateField(r['Prompt Date'])) #到期日    Date    N   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                    fields.append(c) #客户内部资金账户  char(18)    N
                    fields.append('') #客户统一开户编码 char(8) N   客户在统一开户系统中的编码
                    fields.append(self.getProductCurrence(r['Product'])) #币种    Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                    fields.append(self.getExchName(r['Exchange'])) #交易所 Char(5) N   具体见数据字典交易所名称部分
                    fields.append(self.getProduct(r['Product'])) #品种    Char(20)    N   具体见数据字典品种名称部分
                    fields.append(self.getProduct(r['Product'])+r['Prompt Date']) #合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                    fields.append('0.00') #买持仓量 Number(10)  N
                    fields.append(r['No of Lots(sell)']) #卖持仓量  Number(10)  N
                    fields.append('0.00') #持仓盈亏(逐笔对冲)   Number(14,2)    N
                    #ccjj = r['No of Lots(buy)']*r['Average Trading price(buy)']+r['No of Lots(sell)']*r['Average Trading price(sell)']
                    fields.append(r['Average Trading price(sell)']) #持仓均价   Number(14,7)    Y
                    fields.append('0.00') #今结算价 Number(14,7)    N
                    fields.append('') #期权市值 Number(14,2)    Y
                    fields.append('') #期权类型 Char(1) Y   ”C” –看涨期权,”P”-看跌期权
                    fields.append('') #执行价  Number(14,7)    Y
                    fields.append('') #上手清算机构代码 char(2) Y
                    fields.append('') #境内期货公司的账号    Char(10)    Y   即境内期货公司在上手清算公司的账号
                    #2012-01-04@2012-03-07@0001@12345678@USD@CBOT@BO@BO1203@ 0@1@168.00@51.79@51.52@@@@jp@12345
                    fs = [strip(str(i)) for i in fields]
                    ln = '@'.join(fs)
                    f.write(ln+'\n')
        f.close()
        logger.debug('end deal holddataSummary')


    def liquiddetails(self): #平仓明细文件
        #平仓盈亏 填写 的是平均成交卖价
        fn = self.getFileName('liquiddetails')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal liquiddetails')
        for (c,rs) in self.closedPositionFullDetails:
            m_preRec1 = None
            m_preRec2 = None
            m_isTWO = False
            for r in rs:
                if (not m_isTWO):   # 平仓是一对对出现的， 所以两条记录要一块处理， 这也是非常复杂而容易出错的地方
                    m_preRec1 = copy.deepcopy(r)
                    m_isTWO = not m_isTWO
                    continue
                else:
                    m_preRec2 = copy.deepcopy(r)
                    m_isTWO = not m_isTWO
                if ('#' in m_preRec1['Open or Close']):
                    m_CRec = m_preRec1
                    m_ORec = m_preRec2
                else:
                    m_CRec = m_preRec2
                    m_ORec = m_preRec1
                    
                fields=[]
                (m_prod,m_prompt) = self.splitDescription(m_CRec['Description'])
                (m_date,m_time) = self.spliteDateTime(m_CRec['Trade Date'])
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                fields.append(self.getSDateField(m_date)) #成交日期   Date    N   格式：YYYY-MM-DD
                fields.append(self.getPromptDateField(m_prompt,m_prod)) #到期日    Date    N   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                fields.append(c) #客户内部资金账户  char(18)    N
                fields.append('') #客户统一开户编码 char(8) Y   客户在统一开户系统中的编码
                fields.append(self.getCurrencyField(currency = m_CRec['Currency'],product = m_prod)) #币种    Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                fields.append(self.getExchName(m_CRec['Exchange'])) #交易所 Char(5) N   具体见数据字典交易所名称部分
                fields.append(self.getProduct(m_prod)) #品种    Char(20)    N   具体见数据字典品种名称部分
                fields.append(self.getDescriptionField(m_prod , m_prompt)) #合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                #fields.append(m_CRec['Trade Ref.']) #成交流水号    char(16)    N   公司交易系统发布的成交序列号
                fields.append(self.getTradeRefField(r['Trade Ref.']))
                fields.append(m_CRec['No of Lots(buy)']) #买量   Number(10)  N 
                fields.append(m_CRec['No of Lots(sell)']) #卖量   Number(10)  N
                if (m_CRec['No of Lots(buy)'] > 0):
                    cjj = m_CRec['Trade Price/ Permium (Buy)']
                else:
                    cjj = m_CRec['Trade Price/ Permium (Sell)']
                fields.append(cjj) #成交价    Number(14,7)    N   平仓时的成交价
                fields.append(m_CRec['Gross Profit/(Loss)'] + m_ORec['Gross Profit/(Loss)']) #平仓盈亏(逐笔对冲)  Number(14,2)    N   逐笔对冲方式的平仓盈亏，包括LME已到账盈亏
                fields.append('') #权利金  Number(14,2)    Y   期权的平仓价格*成交数量，若行权或放弃则权利金为0
                fields.append('') #期权类型 Char(1) Y   ”C” –看涨期权,”P”-看跌期权
                fields.append('') #执行价  Number(14,7)    Y
                (m_date2,m_time2) = self.spliteDateTime(m_ORec['Trade Date'])
                m_ykcrq = m_date2
                m_ycjls = m_ORec['Trade Ref.']
                if (m_ORec['No of Lots(buy)'] > 0):
                    m_ykcj = m_ORec['Trade Price/ Permium (Buy)']
                else:
                    m_ykcj = m_ORec['Trade Price/ Permium (Sell)']
                fields.append(m_ykcrq) #原开仓日期  Date    N   格式：YYYY-MM-DD 
                #fields.append(m_ycjls) #原成交流水    char(16)    N   对应开仓的成交序列号
                fields.append(self.getTradeRefField(m_ycjls))
                fields.append(m_ykcj) #原开仓价    Number(14,7)    N   对应开仓的成交价
                fields.append('') #上手清算机构代码 char(2) Y
                fields.append('') #境内期货公司的账号    Char(10)    Y   即境内期货公司在上手清算公司的账号
                #2012-01-04@2012-01-03@2012-03-07@0001@12345678@USD@CBOT@BO@BO1203@2011021000000006@0@1@51.88@-48.00@@@@2011-12-19@2011021000000001@51.79@jp@12345
                fs = [strip(str(i)) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
        f.close()
        logger.debug('end deal liquiddetails')

    '''
    def liquiddetailsOLD(self): #平仓明细文件
        #平仓盈亏 填写 的是平均成交卖价
        fn = self.getFileName('liquiddetails')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal liquiddetails')
        for (c,rs) in self.closedPositionFullDetails:
            m_preRec = None
            for r in rs:
                if ('*' in r['Open or Close']):
                    m_preRec = copy.deepcopy(r)
                fields=[]
                (m_prod,m_prompt) = self.splitDescription(r['Description'])
                (m_date,m_time) = self.spliteDateTime(r['Trade Date'])
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                fields.append(self.getSDateField(m_date)) #成交日期   Date    N   格式：YYYY-MM-DD
                fields.append(self.getPromptDateField(m_prompt,m_prod)) #到期日    Date    N   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                fields.append(c) #客户内部资金账户  char(18)    N
                fields.append('') #客户统一开户编码 char(8) Y   客户在统一开户系统中的编码
                fields.append(self.getCurrencyField(currency = r['Currency'],product = m_prod)) #币种    Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                fields.append(self.getExchName(r['Exchange'])) #交易所 Char(5) N   具体见数据字典交易所名称部分
                fields.append(self.getProduct(m_prod)) #品种    Char(20)    N   具体见数据字典品种名称部分
                fields.append(self.getProduct(m_prod) + m_prompt) #合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                fields.append(r['Trade Ref.']) #成交流水号    char(16)    N   公司交易系统发布的成交序列号
                fields.append(r['No of Lots(buy)']) #买量   Number(10)  N 
                fields.append(r['No of Lots(sell)']) #卖量   Number(10)  N
                if (r['No of Lots(buy)'] > 0):
                    cjj = r['Trade Price/ Permium (Buy)']
                else:
                    cjj = r['Trade Price/ Permium (Sell)']
                fields.append(cjj) #成交价    Number(14,7)    N   平仓时的成交价
                fields.append(r['Gross Profit/(Loss)']) #平仓盈亏(逐笔对冲)  Number(14,2)    N   逐笔对冲方式的平仓盈亏，包括LME已到账盈亏
                fields.append('') #权利金  Number(14,2)    Y   期权的平仓价格*成交数量，若行权或放弃则权利金为0
                fields.append('') #期权类型 Char(1) Y   ”C” –看涨期权,”P”-看跌期权
                fields.append('') #执行价  Number(14,7)    Y
                m_ykcrq = ''
                m_ycjls = ''
                m_ykcj = ''
                if ('#' in r['Open or Close']):
                    if (m_preRec):
                        (m_date2,m_time2) = self.spliteDateTime(m_preRec['Trade Date'])
                        m_ykcrq = m_date2
                        m_ycjls = m_preRec['Trade Ref.']
                        if (m_preRec['No of Lots(buy)'] > 0):
                            m_ykcj = m_preRec['Trade Price/ Permium (Buy)']
                        else:
                            m_ykcj = m_preRec['Trade Price/ Permium (Sell)']
                fields.append(m_ykcrq) #原开仓日期  Date    N   格式：YYYY-MM-DD 
                fields.append(m_ycjls) #原成交流水    char(16)    N   对应开仓的成交序列号
                fields.append(m_ykcj) #原开仓价    Number(14,7)    N   对应开仓的成交价
                fields.append('') #上手清算机构代码 char(2) Y
                fields.append('') #境内期货公司的账号    Char(10)    Y   即境内期货公司在上手清算公司的账号
                #2012-01-04@2012-01-03@2012-03-07@0001@12345678@USD@CBOT@BO@BO1203@2011021000000006@0@1@51.88@-48.00@@@@2011-12-19@2011021000000001@51.79@jp@12345
                fs = [strip(str(i)) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
        f.close()
        logger.debug('end deal liquiddetails')
    '''

    def liquiddetailsSummary(self): #平仓明细文件 弃用
        #平仓盈亏 填写 的是平均成交卖价
        fn = self.getFileName('liquiddetails')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal liquiddetailsSummary')
        for (c,rs) in self.closedPositionSummary:
            for r in rs:
                fields=[]
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                fields.append(self.getSDateField()) #成交日期   Date    N   格式：YYYY-MM-DD
                fields.append(self.getSDateField(r['Prompt Date'])) #到期日    Date    N   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                fields.append(c) #客户内部资金账户  char(18)    N
                fields.append('') #客户统一开户编码 char(8) Y   客户在统一开户系统中的编码
                fields.append(r['Currency']) #币种    Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                fields.append(self.getExchName(r['Exchange'])) #交易所 Char(5) N   具体见数据字典交易所名称部分
                fields.append(self.getProduct(r['Product'])) #品种    Char(20)    N   具体见数据字典品种名称部分
                fields.append(self.getProduct(r['Product'])+self.getSDateField(r['Prompt Date'])) #合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                fields.append('') #成交流水号    char(16)    N   公司交易系统发布的成交序列号
                fields.append('0.00') #买量   Number(10)  N 
                fields.append(r['No of Lots Closed']) #卖量   Number(10)  N
                fields.append(r['Average Trading price(sell)']) #成交价    Number(14,7)    N   平仓时的成交价
                #m_np = r['No of Lots Closed']*(r['Average Trading price(sell)']-r['Average Trading price(buy)'])
                fields.append(r['Net Profit']) #平仓盈亏(逐笔对冲)  Number(14,2)    N   逐笔对冲方式的平仓盈亏，包括LME已到账盈亏
                fields.append('') #权利金  Number(14,2)    Y   期权的平仓价格*成交数量，若行权或放弃则权利金为0
                fields.append('') #期权类型 Char(1) Y   ”C” –看涨期权,”P”-看跌期权
                fields.append('') #执行价  Number(14,7)    Y
                fields.append('0000-00-00') #原开仓日期  Date    N   格式：YYYY-MM-DD 
                fields.append('') #原成交流水    char(16)    N   对应开仓的成交序列号
                fields.append(r['Average Trading price(buy)']) #原开仓价    Number(14,7)    N   对应开仓的成交价
                fields.append('') #上手清算机构代码 char(2) Y
                fields.append('') #境内期货公司的账号    Char(10)    Y   即境内期货公司在上手清算公司的账号
                #2012-01-04@2012-01-03@2012-03-07@0001@12345678@USD@CBOT@BO@BO1203@2011021000000006@0@1@51.88@-48.00@@@@2011-12-19@2011021000000001@51.79@jp@12345
                fs = [strip(str(i)) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
        f.close()
        logger.debug('end deal liquiddetailsSummary')


    def holddetails(self): #持仓明细文件
        fn = self.getFileName('holddetails')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal holddetails')
        for (c,rs) in self.openPositionFullDetails:
            for r in rs:
                fields=[]
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                (m_prod,m_prompt) = self.splitDescription(r['Description'])
                m_prodcmf = self.getProduct(m_prod) # m_prod2 是格式化的
                fields.append(self.getPromptDateField(m_prompt,m_prod)) #到期日    Date    N   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                (m_date,m_time) = self.spliteDateTime(r['Trade Date'])
                fields.append(m_date) #成交日期 Date    N   当笔持仓的开仓日期， 格式(form)：YYYY-MM-DD
                fields.append(c) #客户内部资金账户  char(18)    N   
                fields.append('') #客户统一开户编码 char(8) Y   客户在统一开户系统中的编码
                m_currency = self.getCurrencyField(r['Currency'],product=m_prod)
                fields.append(m_currency) #币种    Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                fields.append(self.getExchName(r['Exchange'])) #交易所 Char(5) N   具体见数据字典交易所名称部分
                fields.append(m_prodcmf) #品种    Char(20)    N   具体见数据字典品种名称部分
                fields.append(self.getDescriptionField(m_prod , m_prompt)) #合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                #fields.append(r['Trade Ref.']) #成交流水号    Char(16)    N   公司交易系统发布的成交序列号
                fields.append(self.getTradeRefField(r['Trade Ref.']))
                
                fields.append(r['No of Lots(buy)']) #买持仓量   Number(10)  N
                fields.append(r['No of Lots(sell)']) #卖持仓量  Number(10)  N
                fields.append(r['Trade Price/Premium']) #开仓价 Number(14,7)    N
                fields.append(r['Closing Price']) #今结算价   Number(14,7)    N
                #if (r['No of Lots(buy)'] > 0):
                #   m_pl = (r['Closing Price'] - r['Trade Price/Premium']) * r['No of Lots(buy)']
                #else:
                #   m_pl = (r['Trade Price/Premium'] - r['Closing Price']) * r['No of Lots(sell)']
                fields.append(r['Gross Floating P/(L)']) #持仓盈亏(逐笔对冲) Number(14,2)    N   逐笔对冲方式的持仓盈亏（包括LME未到期合约）若为期权持仓，则盈亏为期权市值
                
                fields.append('') #期权市值 Number(14,2)    Y   期权的结算价*持仓量
                fields.append('') #期权类型 Char(1) Y   ”C” –看涨期权,”P”-看跌期权
                fields.append('') #执行价  Number(14,7)    Y
                fields.append(self.getSettlerName(r['Broker'])) #上手清算机构代码 char(2) Y
                fields.append('') #境内期货公司的账号    Char(10)    Y   即境内期货公司在上手清算公司的账号
                #2012-01-04@2012-03-07@2011-12-08@0001@12345678@USD@CBOT@BO@BO1203@2011021000000066@1@0@56.47@51.86@-4284@@@@jp@12345
                #import copy
                #fields2 = copy.deepcopy(fields)
                fs = [strip(str(i)) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
        self.holddetails_unsettledClosedPositions(f)
        f.close()
        logger.debug('end deal holddetails')

    def holddetails_unsettledClosedPositions(self,p_f): #持仓明细文件　增加ＬＭＥ平仓未到期合约
        f = p_f
        logger.debug('begin deal holddetails_unsettledClosedPositions')
        for (c,rs) in self.unsettledClosedPositionFullDetails:
            for r in rs:
                fields=[]
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                (m_prod,m_prompt) = self.splitDescription(r['Description'])
                m_prodcmf = self.getProduct(m_prod) # m_prod2 是格式化的
                fields.append(self.getPromptDateField(m_prompt,m_prod)) #到期日    Date    N   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                (m_date,m_time) = self.spliteDateTime(r['Trade Date'])
                fields.append(m_date) #成交日期 Date    N   当笔持仓的开仓日期， 格式(form)：YYYY-MM-DD
                fields.append(c) #客户内部资金账户  char(18)    N   
                fields.append('') #客户统一开户编码 char(8) Y   客户在统一开户系统中的编码
                m_currency = self.getCurrencyField(r['Currency'],product=m_prod)
                fields.append(m_currency) #币种    Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                fields.append(self.getExchName(r['Exchange'])) #交易所 Char(5) N   具体见数据字典交易所名称部分
                fields.append(m_prodcmf) #品种    Char(20)    N   具体见数据字典品种名称部分
                fields.append(self.getDescriptionField(m_prod , m_prompt)) #合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                #fields.append(r['Trade Ref.']) #成交流水号    Char(16)    N   公司交易系统发布的成交序列号
                fields.append(self.getTradeRefField(r['Trade Ref.']))
                
                fields.append(r['No of Lots(buy)']) #买持仓量   Number(10)  N
                fields.append(r['No of Lots(sell)']) #卖持仓量  Number(10)  N
                fields.append(r['Trade Price/Premium']) #开仓价 Number(14,7)    N
                fields.append(r['Closing Price']) #今结算价   Number(14,7)    N
                #if (r['No of Lots(buy)'] > 0):
                #   m_pl = (r['Closing Price'] - r['Trade Price/Premium']) * r['No of Lots(buy)']
                #else:
                #   m_pl = (r['Trade Price/Premium'] - r['Closing Price']) * r['No of Lots(sell)']
                fields.append(r['Gross Floating P/(L)']) #持仓盈亏(逐笔对冲) Number(14,2)    N   逐笔对冲方式的持仓盈亏（包括LME未到期合约）若为期权持仓，则盈亏为期权市值
                
                fields.append('') #期权市值 Number(14,2)    Y   期权的结算价*持仓量
                fields.append('') #期权类型 Char(1) Y   ”C” –看涨期权,”P”-看跌期权
                fields.append('') #执行价  Number(14,7)    Y
                fields.append(self.getSettlerName(r['Broker'])) #上手清算机构代码 char(2) Y
                fields.append('') #境内期货公司的账号    Char(10)    Y   即境内期货公司在上手清算公司的账号
                #2012-01-04@2012-03-07@2011-12-08@0001@12345678@USD@CBOT@BO@BO1203@2011021000000066@1@0@56.47@51.86@-4284@@@@jp@12345
                #import copy
                #fields2 = copy.deepcopy(fields)
                fs = [strip(str(i)) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
        logger.debug('end deal holddetails_unsettledClosedPositions')


    def holddetailsSummary(self): #持仓明细文件，原使用合并的记录，暂放弃
        fn = self.getFileName('holddetails')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal holddetailsSummary')
        for (c,rs) in self.unsettledClosedPositions:
            for r in rs:
                fields=[]
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                fields.append(self.getSDateField(r['Prompt Date'])) #到期日    Date    N   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                fields.append('') #成交日期 Date    N   当笔持仓的开仓日期， 格式(form)：YYYY-MM-DD
                fields.append(c) #客户内部资金账户  char(18)    N   
                fields.append('') #客户统一开户编码 char(8) Y   客户在统一开户系统中的编码
                
                fields.append(self.getProductCurrence(r['Product'])) #币种    Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                fields.append(self.getExchName(r['Exchange'])) #交易所 Char(5) N   具体见数据字典交易所名称部分
                fields.append(self.getProduct(r['Product'])) #品种    Char(20)    N   具体见数据字典品种名称部分
                fields.append(self.getProduct(r['Product'])+self.getSDateField(r['Prompt Date'])) #合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                fields.append('') #成交流水号    Char(16)    N   公司交易系统发布的成交序列号
                
                fields.append(r['No of Lots(buy)']) #买持仓量   Number(10)  N
                fields.append(r['No of Lots(sell)']) #卖持仓量  Number(10)  N
                fields.append(r['Average Trading price(buy)']) #开仓价 Number(14,7)    N
                fields.append(r['Average Trading price(sell)']) #今结算价   Number(14,7)    N
                m_pl = (r['Average Trading price(sell)'] - r['Average Trading price(buy)'])* r['No of Lots(sell)']
                fields.append(m_pl) #持仓盈亏(逐笔对冲) Number(14,2)    N   逐笔对冲方式的持仓盈亏（包括LME未到期合约）若为期权持仓，则盈亏为期权市值
                
                fields.append('') #期权市值 Number(14,2)    Y   期权的结算价*持仓量
                fields.append('') #期权类型 Char(1) Y   ”C” –看涨期权,”P”-看跌期权
                fields.append('') #执行价  Number(14,7)    Y
                fields.append('') #上手清算机构代码 char(2) Y
                fields.append('') #境内期货公司的账号    Char(10)    Y   即境内期货公司在上手清算公司的账号
                #2012-01-04@2012-03-07@2011-12-08@0001@12345678@USD@CBOT@BO@BO1203@2011021000000066@1@0@56.47@51.86@-4284@@@@jp@12345
                #import copy
                #fields2 = copy.deepcopy(fields)
                fs = [strip(str(i)) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
        f.close()
        logger.debug('end deal holddetailsSummary')

    def delivtails(self): #交割明细文件
        fn = self.getFileName('delivtails')
        self.txtFiles.append(fn)
        f = open(fn,'w+')
        logger.debug('begin deal delivtails')
        for (c,rs) in self.delivtailsRecord:
            for r in rs:
                fields=[]
                fields.append(self.getSDateField()) #结算日期   Date    N   格式：YYYY-MM-DD
                fields.append('') #品种   Char(20)    N   具体见数据字典品种名称部分
                fields.append('') #交割日期 Date    N   格式(form)：YYYY-MM-DD
                fields.append('') #到期日  Date    N   LME合约填写到期日，其他交易所合约填写最后交易日格式：YYYY-MM-DD
                fields.append('') #客户内部资金账户 char(18)    N
                fields.append('') #客户统一开户编码 char(8) Y   客户在统一开户系统中的编码
                fields.append('') #币种   Char(7) N   按照 ISO 4217; E.g:USD; JPY，如为LME美元则填写USD-LME
                fields.append('') #交易所  Char(5) N   具体见数据字典交易所名称部分
                fields.append('') #品种   Char(20)    N   具体见数据字典品种名称部分
                fields.append('') #合约描述 Char(40)    N   品种+到期时间。如玉米11年12月的合约为CO1112,LME铜11年12月3日的合约为CA111203
                fields.append('') #买交割量 Number(10)  N
                fields.append('') #卖交割量 Number(10)  N
                fields.append('') #交割价  Number(14,7)    N
                fields.append('') #交割价  Number(14,7)    N
                fields.append('') #上手清算机构代码 char(2) Y
                fields.append('') #境内期货公司的账号    Char(10)    Y   即境内期货公司在上手清算公司的账号
                #2012-01-04@2012-01-04@2012-01-04@0001@12345678@USD@LME@CA@CA20110104@1@0@88.47@5.12@jp@12345
                fs = [strip(str(i)) for i in fields]
                ln = '@'.join(fs)
                f.write(ln+'\n')
        f.close()
        logger.debug('end deal delivtails')
    
    def createZipFile(self):
        import zipfile
        logger.debug('begin createZipFile')
        zipFile = zipfile.ZipFile(self.dateOfFileName+'.zip','w')
        for f in self.txtFiles:
            zipFile.write(f)
        zipFile.close()
        logger.debug('end createZipFile')
    
    def sendMail(self):
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart
        import smtplib
        if (not self.sendEMail):
            return
        logger.debug('begin sendMail')
        #创建一个带附件的实例
        msg = MIMEMultipart()
        #构造附件1
        fn = self.dateOfFileName+'.zip'
        att1 = MIMEText(open(fn, 'rb').read(), 'base64', 'gb2312')
        att1["Content-Type"] = 'application/octet-stream'
        att1["Content-Disposition"] = 'attachment; filename="%s"' % fn #这里的filename可以任意写，写什么名字，邮件中显示什么名字
        msg.attach(att1)
        #加邮件头
        msg['to'] = self.sendEMail
        msg['from'] = base64.decodestring(mail_from) 
        msg['subject'] = 'GF Statement report: %s' % self.dateOfFileName
        #发送邮件
        try:
            server = smtplib.SMTP(base64.decodestring(mail_server))
            server.ehlo()
            server.starttls()
            server.login(base64.decodestring(mail_id),base64.decodestring(mail_pw)) #XXX为用户名，XXXXX为密码
            server.sendmail(msg['from'], msg['to'],msg.as_string())
            server.quit()
        except Exception, e:  
            logger.error("Send Email Failed:%s" % str(e))
        logger.debug('end sendMail')

def main():
    import argparse
    __author__ = 'TianJun'
    parser = argparse.ArgumentParser(description='This is a Lynx2CMF script by TianJun.')
    parser.add_argument('-d','--date', help='Input SettledDate YYYYMMDD,default is today.',required=False)
    parser.add_argument('-a','--account', help='Input account XXXXXX-000',required=False)
    parser.add_argument('-f','--xlsfname', help='Input lynx export xls file name X.XLS,default:AccSum_YYYYMMDD.xlsx',required=False)
    parser.add_argument('-m','--email', help='Input email to send result.',required=False)
    
    args = parser.parse_args()
    m_settledDate = datetime.datetime.now().strftime('%Y%m%d')
    m_account = None
    m_xlsfname = None
    m_email = None
    if (args.date):
        m_settledDate = args.date
    if (args.account):
        m_account = args.account     
    if (args.xlsfname):
        m_xlsfname = args.xlsfname     
    if (args.email):
        m_email = args.email     
    logger.info('Start...')
    cmf =  DealCMFChinaData(m_settledDate,m_account,m_xlsfname,m_email) # 带上参数即可只输出单一指定帐户的数据
    logger.info('...End.')


if __name__ == '__main__':
    main()

