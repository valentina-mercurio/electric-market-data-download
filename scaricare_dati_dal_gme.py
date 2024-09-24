from selenium.webdriver import Chrome
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import calendar
from datetime import datetime
from dateutil.relativedelta import relativedelta
from time import sleep
import os
import urllib
import glob
import xml.etree.ElementTree as ET
from itertools import islice
import csv
import pandas as pd
import zipfile
import shutil
import pypyodbc as odbc
import numpy as np
import re
import mysql
import mysql.connector
from mysql.connector import Error
import pandas as pd
import pymysql
import sqlalchemy
from sqlalchemy import create_engine
import openpyxl as opx

def MGP_prezzi(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.CLASS_NAME, 'ContentPlaceHolder1_MenuDownload1_MenuMGP_1').click()
    driver.find_element(By.ID, "ContentPlaceHolder1_MenuDownload1_SubMenuMGPn4").click()
    
    dati_mancantiMGP_prezzi=[]
    for month in range(1,13):
        #inserire la data  di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMGP_prezzi.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()
            
    if len(dati_mancantiMGP_prezzi)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancanti relativi ai prezzi del MGP sono: ")
                    for elem in dati_mancantiMGP_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')                
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancanti relativi ai prezzi del MGP sono: ")
                    for elem in dati_mancantiMGP_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')
            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        

def MGP_quantità(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.CLASS_NAME, 'ContentPlaceHolder1_MenuDownload1_MenuMGP_1').click()   
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_SubMenuMGPn9').click()

    dati_mancantiMGP_quantità=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.ID, 'ContentPlaceHolder1_tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.ID, 'ContentPlaceHolder1_tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMGP_quantità.append(month) 
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()
    if len(dati_mancantiMGP_quantità)!=0:
    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancanti relativi alle quantità del MGP sono: ")
                    for elem in dati_mancantiMGP_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')              
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancanti relativi alle quantità del MGP sono: ")
                    for elem in dati_mancantiMGP_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')
            return True            
                                                
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        

def MGP_transiti(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.CLASS_NAME, 'ContentPlaceHolder1_MenuDownload1_MenuMGP_1').click()    
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_SubMenuMGPn7').click()
    
    dati_mancantiMGP_transiti=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.ID, 'ContentPlaceHolder1_tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMGP_transiti.append(month) 
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()
    
    if len(dati_mancantiMGP_transiti)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancanti relativi ai transiti del MGP sono: ")
                    for elem in dati_mancantiMGP_transiti:
                        f.write("%s, " % elem)
                    f.write('\n')             
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancanti relativi ai transiti del MGP sono: ")
                    for elem in dati_mancantiMGP_transiti:
                        f.write("%s, " % elem)
                    f.write('\n')            
              
            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
       
def MI1_prezzi(driver, day, year, nome_file_mancanti):    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI1').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n2"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI1_prezzi=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
            
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI1_prezzi.append(month)       
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI1_prezzi)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per i prezzi del M1 sono: ")
                    for elem in dati_mancantiMI1_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')             
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per i prezzi del M1 sono: ")
                    for elem in dati_mancantiMI1_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')            
                   
            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
      
def MI1_quantità(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI1').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n4"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI1_quantità=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI1_quantità.append(month) 
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI1_quantità)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per le quantità del M1 sono: ")
                    for elem in dati_mancantiMI1_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')             
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per le quantità del M1 sono: ")
                    for elem in dati_mancantiMI1_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')          
                                
            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
 
def MI2_prezzi(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()   
    driver.find_element(By.LINK_TEXT, 'MI2').click()
    driver.find_element(By.XPATH,'//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n7"]/td/table/tbody/tr/td/a').click()
    
    
    dati_mancantiMI2_prezzi=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI2_prezzi.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI2_prezzi)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per i prezzi del M2 sono: ")
                    for elem in dati_mancantiMI2_prezzi:
                        f.write("%s, " % elem)                
                    f.write('\n')            
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per i prezzi del M2 sono: ")
                    for elem in dati_mancantiMI2_prezzi:
                        f.write("%s, " % elem)                
                    f.write('\n')            
            
            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
 
def MI2_quantità(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI2').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n9"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI2_quantità=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI2_quantità.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI2_quantità)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per le quantità del M2 sono: ")
                    for elem in dati_mancantiMI2_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per le quantità del M2 sono: ")
                    for elem in dati_mancantiMI2_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')           
            
            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
   
def MI3_prezzi(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI3').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n12"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI3_prezzi=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI3_prezzi.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI3_prezzi)!=0:
    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per i prezzi del M3 sono: ")
                    for elem in dati_mancantiMI3_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per i prezzi del M3 sono: ")
                    for elem in dati_mancantiMI3_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
        
def MI3_quantità(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI3').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n14"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI3_quantità=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI3_quantità.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI3_quantità)!=0:
    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per le quantità del M3 sono: ")
                    for elem in dati_mancantiMI3_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per le quantità del M3 sono: ")
                    for elem in dati_mancantiMI3_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
  
def MI4_prezzi(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI4').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n17"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI4_prezzi=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI4_prezzi.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI4_prezzi)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per i prezzi del M4 sono: ")
                    for elem in dati_mancantiMI4_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per i prezzi del M4 sono: ")
                    for elem in dati_mancantiMI4_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')            

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        

def MI4_quantità(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI4').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n19"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI4_quantità=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI4_quantità.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI4_quantità)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per le quantità del M4 sono: ")
                    for elem in dati_mancantiMI4_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per le quantità del M4 sono: ")
                    for elem in dati_mancantiMI4_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')              

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
   
def MI5_prezzi(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI5').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n22"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI5_prezzi=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI5_prezzi.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI5_prezzi)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per i prezzi del M5 sono: ")
                    for elem in dati_mancantiMI5_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per i prezzi del M5 sono: ")
                    for elem in dati_mancantiMI5_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')              

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        

def MI5_quantità(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI5').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n24"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI5_quantità=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI5_quantità.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
   
    if len(dati_mancantiMI5_quantità)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per le quantità del M5 sono: ")
                    for elem in dati_mancantiMI5_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per le quantità del M5 sono: ")
                    for elem in dati_mancantiMI5_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')            

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
       
def MI6_prezzi(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI6').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n27"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI6_prezzi=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI6_prezzi.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI6_prezzi)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per i prezzi del M6 sono: ")
                    for elem in dati_mancantiMI6_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per i prezzi del M6 sono: ")
                    for elem in dati_mancantiMI6_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
      
def MI6_quantità(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI6').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n29"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI6_quantità=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI6_quantità.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()            
    
    if len(dati_mancantiMI6_quantità)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per le quantità del M6 sono: ")
                    for elem in dati_mancantiMI6_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per le quantità del M6 sono: ")
                    for elem in dati_mancantiMI6_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')            

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
    
def MI7_prezzi(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI7').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n32"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI7_prezzi=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI7_prezzi.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI7_prezzi)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per i prezzi del M7 sono: ")
                    for elem in dati_mancantiMI7_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per i prezzi del M7 sono: ")
                    for elem in dati_mancantiMI7_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')              

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
      
def MI7_quantità(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MI7').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMI1n34"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI7_quantità=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI7_quantità.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
            
    if len(dati_mancantiMI7_quantità)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per le quantità del M7 sono: ")
                    for elem in dati_mancantiMI7_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per le quantità del M7 sono: ")
                    for elem in dati_mancantiMI7_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')            

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
      
def MI_A1_prezzi(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MERCATO INFRAGIORNALIERO').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMercatoInfragiornalieron2"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI_A1_prezzi=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI_A1_prezzi.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI_A1_prezzi)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per i prezzi del MI-A1 sono: ")
                    for elem in dati_mancantiMI_A1_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per i prezzi del MI-A1 sono: ")
                    for elem in dati_mancantiMI_A1_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
    
def MI_A1_quantità(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MERCATO INFRAGIORNALIERO').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMercatoInfragiornalieron3"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI_A1_quantità=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI_A1_quantità.append(month) 
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI_A1_quantità)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per le quantità del MI-A1 sono: ")
                    for elem in dati_mancantiMI_A1_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per le quantità del MI-A1 sono: ")
                    for elem in dati_mancantiMI_A1_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
          
def MI_A2_prezzi(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.ID, 'ContentPlaceHolder1_MenuDownload1_MenuMI').click()    
    driver.find_element(By.LINK_TEXT, 'MERCATO INFRAGIORNALIERO').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMercatoInfragiornalieron6"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI_A2_prezzi=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI_A2_prezzi.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI_A2_prezzi)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per i prezzi del MI-A2 sono: ")
                    for elem in dati_mancantiMI_A2_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per i prezzi del MI-A2 sono: ")
                    for elem in dati_mancantiMI_A2_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')            

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
    
def MI_A2_quantità(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.LINK_TEXT, 'MERCATO INFRAGIORNALIERO').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMercatoInfragiornalieron7"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI_A2_quantità=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI_A2_quantità.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI_A2_quantità)!=0:
    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per le quantità del MI-A2 sono: ")
                    for elem in dati_mancantiMI_A2_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per le quantità del MI-A2 sono: ")
                    for elem in dati_mancantiMI_A2_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
      
def MI_A3_prezzi(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.LINK_TEXT, 'MERCATO INFRAGIORNALIERO').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMercatoInfragiornalieron10"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI_A3_prezzi=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI_A3_prezzi.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI_A3_prezzi)!=0:
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per i prezzi del MI-A3 sono: ")
                    for elem in dati_mancantiMI_A3_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per i prezzi del MI-A3 sono: ")
                    for elem in dati_mancantiMI_A3_prezzi:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
        
def MI_A3_quantità(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.LINK_TEXT, 'MERCATO INFRAGIORNALIERO').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMercatoInfragiornalieron11"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI_A3_quantità=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI_A3_quantità.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        

    if len(dati_mancantiMI_A3_quantità)!=0:    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati per le quantità del MI-A3 sono: ")
                    for elem in dati_mancantiMI_A3_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati per le quantità del MI-A3 sono: ")
                    for elem in dati_mancantiMI_A3_quantità:
                        f.write("%s, " % elem)
                    f.write('\n')              

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
      
def MI_XBID_1FASE(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.LINK_TEXT, 'MERCATO INFRAGIORNALIERO').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMercatoInfragiornalieron14"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI_XBID_1FASE=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI_XBID_1FASE.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI_XBID_1FASE)!=0:    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati relativi agli esiti del MI-XBID per la prima fase sono: ")
                    for elem in dati_mancantiMI_XBID_1FASE:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati relativi agli esiti del MI-XBID per la prima fase sono: ")
                    for elem in dati_mancantiMI_XBID_1FASE:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
          
def MI_XBID_2FASE(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.LINK_TEXT, 'MERCATO INFRAGIORNALIERO').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMercatoInfragiornalieron14"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI_XBID_2FASE=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI_XBID_2FASE.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI_XBID_2FASE)!=0:    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati relativi agli esiti del MI-XBID per la seconda fase sono: ")
                    for elem in dati_mancantiMI_XBID_2FASE:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati relativi agli esiti del MI-XBID per la seconda fase sono: ")
                    for elem in dati_mancantiMI_XBID_2FASE:
                        f.write("%s, " % elem)
                    f.write('\n')            

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
        
def MI_XBID_3FASE(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.LINK_TEXT, 'MERCATO INFRAGIORNALIERO').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMercatoInfragiornalieron13"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMI_XBID_3FASE=[]

    for month in range(1,13):
        #inserire la data di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMI_XBID_3FASE.append(month) 
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMI_XBID_3FASE)!=0:    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati relativi agli esiti del MI-XBID per la terza fase sono: ")
                    for elem in dati_mancantiMI_XBID_3FASE:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati relativi agli esiti del MI-XBID per la terza fase sono: ")
                    for elem in dati_mancantiMI_XBID_3FASE:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
         
def MPEG(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.CLASS_NAME, 'ContentPlaceHolder1_MenuDownload1_MenuMPEG_3').click()
    driver.find_element(By.LINK_TEXT, 'MPEG').click()
    
    dati_mancantiMPEG=[]

    for month in range(1,13):
        #inserire la data  di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMPEG.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()            
    
    if len(dati_mancantiMPEG)!=0:    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati relativi agli esiti del MPEG sono: ")
                    for elem in dati_mancantiMPEG:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati relativi agli esiti del MPEG sono: ")
                    for elem in dati_mancantiMPEG:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
            

def MSD_EX_ANTE(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.CLASS_NAME, 'ContentPlaceHolder1_MenuDownload1_MenuMSD_3').click()
    driver.find_element(By.LINK_TEXT, 'MSD ex-ante').click()
    
    dati_mancantiMSD_EX_ANTE=[]

    for month in range(1,13):
        #inserire la data  di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMSD_EX_ANTE.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()            
    
    if len(dati_mancantiMSD_EX_ANTE)!=0:    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati relativi agli esiti del MSD ex ante sono: ")
                    for elem in dati_mancantiMSD_EX_ANTE:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati relativi agli esiti del MSD ex ante sono: ")
                    for elem in dati_mancantiMSD_EX_ANTE:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
                

def MSD_EX_POST(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.CLASS_NAME, 'ContentPlaceHolder1_MenuDownload1_MenuMSD_3').click()
    driver.find_element(By.LINK_TEXT, 'MSD ex-post').click()
    
    dati_mancantiMSD_EX_POST=[]

    for month in range(1,13):
        #inserire la data  di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMSD_EX_POST.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()            
    
    if len(dati_mancantiMSD_EX_POST)!=0:    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati relativi agli esiti del MSD ex post sono: ")
                    for elem in dati_mancantiMSD_EX_POST:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati relativi agli esiti del MSD ex post sono: ")
                    for elem in dati_mancantiMSD_EX_POST:
                        f.write("%s, " % elem)
                    f.write('\n')            

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
        
            
def MSD_MB_PRELIMINARI_TOTALI(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.CLASS_NAME, 'ContentPlaceHolder1_MenuDownload1_MenuMSD_3').click()
    driver.find_element(By.LINK_TEXT, 'MB preliminari').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMSDn2"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMSD_MB_PRELIMINARI_TOTALI=[]

    for month in range(1,13):
        #inserire la data  di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMSD_MB_PRELIMINARI_TOTALI.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMSD_MB_PRELIMINARI_TOTALI)!=0:    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati relativi agli esiti del MSD, in particolare i totali del MB preliminari, sono: ")
                    for elem in dati_mancantiMSD_MB_PRELIMINARI_TOTALI:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati relativi agli esiti del MSD, in particolare i totali del MB preliminari, sono: ")
                    for elem in dati_mancantiMSD_MB_PRELIMINARI_TOTALI:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
           

def MSD_MB_PRELIMINARI_RISERVA_SECONDARIA(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.CLASS_NAME, 'ContentPlaceHolder1_MenuDownload1_MenuMSD_3').click()
    driver.find_element(By.LINK_TEXT, 'MB preliminari').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMSDn3"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMSD_MB_PRELIMINARI_RISERVA_SECONDARIA=[]

    for month in range(1,13):
        #inserire la data  di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMSD_MB_PRELIMINARI_RISERVA_SECONDARIA.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMSD_MB_PRELIMINARI_RISERVA_SECONDARIA)!=0:    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati relativi agli esiti del MSD, in particolare la riserva secondaria del MB preliminari, sono: ")
                    for elem in dati_mancantiMSD_MB_PRELIMINARI_RISERVA_SECONDARIA:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati relativi agli esiti del MSD, in particolare la riserva secondaria del MB preliminari, sono: ")
                    for elem in dati_mancantiMSD_MB_PRELIMINARI_RISERVA_SECONDARIA:
                        f.write("%s, " % elem)
                    f.write('\n')            

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
                       

def MSD_MB_PRELIMINARI_ALTRI_SERVIZI(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.CLASS_NAME, 'ContentPlaceHolder1_MenuDownload1_MenuMSD_3').click()
    driver.find_element(By.LINK_TEXT, 'MB preliminari').click()
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_MenuDownload1_SubMenuMSDn4"]/td/table/tbody/tr/td/a').click()
    
    dati_mancantiMSD_MB_PRELIMINARI_ALTRI_SERVIZI=[]

    for month in range(1,13):
        #inserire la data  di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMSD_MB_PRELIMINARI_ALTRI_SERVIZI.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMSD_MB_PRELIMINARI_ALTRI_SERVIZI)!=0:    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati relativi agli esiti del MSD, in particolare gli altri servizi MB preliminari, sono: ")
                    for elem in dati_mancantiMSD_MB_PRELIMINARI_ALTRI_SERVIZI:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati relativi agli esiti del MSD, in particolare gli altri servizi MB preliminari, sono: ")
                    for elem in dati_mancantiMSD_MB_PRELIMINARI_ALTRI_SERVIZI:
                        f.write("%s, " % elem)
                    f.write('\n')            

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
              

def MTE(driver, day, year, nome_file_mancanti):
    
    #prendere dati
    driver.find_element(By.CLASS_NAME, 'ContentPlaceHolder1_MenuDownload1_MenuMGP_1').click()
    driver.find_element(By.LINK_TEXT, 'MTE').click()
    
    dati_mancantiMTE=[]

    for month in range(1,13):
        #inserire la data  di inizio
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').send_keys(str(day)+"/"+str(month)+"/"+year)
        
        #gestire la lunghezza dei vari mesi
        count=calendar.monthrange(2022,month)
        day_count=count[1]
        
        #inserire la data di fine
        driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').send_keys(str(day+day_count-1)+"/"+str(month)+"/"+year)
        #scaricare dati    
        try:
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$btnScarica').click()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
        except:
            #tenere traccia dei dati mancanti per ogni mese
            dati_mancantiMTE.append(month)
            try:            
                obj = driver.switch_to.alert 
                obj.accept()
            except:
                continue            
        finally:
            #svuotare
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStart').clear()
            driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$tbDataStop').clear()        
    
    if len(dati_mancantiMTE)!=0:    
        try:
            if os.path.isfile(nome_file_mancanti):
                with open(nome_file_mancanti, 'a') as f:
                    f.write("I dati mancati relativi agli esiti del MTE sono: ")
                    for elem in dati_mancantiMTE:
                        f.write("%s, " % elem)
                    f.write('\n')           
            else:
                with open(nome_file_mancanti, 'w') as f:
                    f.write("I dati mancati relativi agli esiti del MTE sono: ")
                    for elem in dati_mancantiMTE:
                        f.write("%s, " % elem)
                    f.write('\n')             

            return True
                    
        except FileNotFoundError:
            print("Impossibile aprire il file: file non trovato o non esiste.")
            return False
                
        except PermissionError:
            print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
            return False
        
        except IsADirectoryError:
            print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
            return False
          

def spostare_files(src, dest, extension):
    
    files = os.listdir(src)

    for f in files:
        if f.endswith(extension):
            src_path = os.path.join(src, f)
            dst_path = os.path.join(dest, f)
            shutil.move(src_path, dst_path)            
    
def estrarre_files(dest,extension):
    
    os.chdir(dest) # change directory from working dir to dir with files
    
    for item in os.listdir(dest): # loop through items in dir
            if item.endswith(extension): # check for ".zip" extension
                    file_name = os.path.abspath(item) # get full path of files
                    zip_ref = zipfile.ZipFile(file_name) # create zipfile object
                    zip_ref.extractall(dest) # extract file to dir
                    zip_ref.close() # close file
                    os.remove(file_name) # delete zipped file                      
       
def convertitore_xml_csv(directory_xml, valore_mancante, livello):
    
    #gestire eventuali eccezioni
    try:    
        #avere i nomi dei files all'interno della directory
        filenames = glob.glob(directory_xml)
        
        #scorrere i nomi dei files
        for filename in filenames:
            #aprire il file
            with open(filename, 'r', encoding="utf-8") as content:
                #analizzare l'albero
                tree=ET.parse(content)
                #avere la radice
                root=tree.getroot()
                #inizializzare un dizionario che conterrà le chiavi corrispondendi alla futura variabile del dataset e la lista corrispondente ai relativi valori
                data={}
                #inizializzare una variabile utile a stabilire il tag relativo al campo da cui prendere i dati
                campo=[]
                #inizializzare una variabile che tenga traccia delle chiavi e quindi delle future variabili del dataset
                chiavi=[]
                #scorrere i figli
                for child in root:
                    #assicurarsi che non sia presente per evitare ripetizioni
                    if child.tag not in campo:
                        #tenere traccia di questi                    
                        campo.append(child.tag)
                        #scorrere i figli dei figli 
                        for child1 in child:
                            #assicurarsi che non sia presente, essendo un dizionario
                            if child1.tag not in data:
                                #aggiungerlo al dizionario
                                data[child1.tag]=[]
                
                #essendo innestati in modo tale che tutti i figli, tranne il primo che viene eliminato, siano strutturati allo stesso modo, si terrà traccia di questo per cercare i dati solo all'interno di quelli che contengono l'informazione valida (infatti il primo non si ocnsidera)            
                campo_da_cercare=str(campo[livello])
                
                #anche la prima chiave viene eliminata non essendo contenitrice di informazione utile
                del data[next(islice(data,0,None))]
                
                #scorrere le chiavi nel dizionario
                for chiave in data.keys():
                    #si trasforma di volta in volta la chiave in stringa per poterla usare come tale
                    chiave_da_cercare=str(chiave)
                    #si collezionano tutte le chiavi come stringhe in una lista
                    chiavi.append(chiave)
                    #si scorrono tutti i tag all'interno del campo ricercato
                    for tag1 in root.findall(campo_da_cercare):
                        #gestire eventuali eccezioni
                        try:
                            #si utilizza una variabile per salvare il dato contenuto all'interno di quel tag
                            valore=tag1.find(chiave_da_cercare).text
                            #lo si appende alla lista relativa alla chiave
                            data[chiave].append(valore)
                        except:
                            #se il valore è mancante lo si sostituisce con un elemento scelto 
                            data[chiave].append(valore_mancante)                               
                
        
            #si crea il nuovo nome, quindi quello del file .csv che si andrà a creare             
            filename_fin=filename.replace("xml","csv")
            #si apre il file in modalità scrittura
        
            with open(filename_fin, 'w', newline='') as csvfile:
                #si crea l'oggetto che andrà a convertire i dati in stringhe delimitate dal simbolo scelto
                writer = csv.writer(csvfile, delimiter='|')
                #scrivere i valori sul file
                writer.writerows(data.values())
                    
            #leggere il file .csv con pandas
            df=pd.read_csv(filename_fin, sep='|', header=None)
            
            #inserire la colonna delle variabili
            elem1=[]
            elem2=[]
            elem3=[]
            elem4=[]
            for i in range(len(chiavi)):
                val=re.split("_|;", chiavi[i])
                    
                if len(val)==1:
                    if val[0].isupper() or val[0]=="Data" or val[0]=="Mercato" or val[0]=="Ora" or val[0]=="A" or val[0]=="Da":
                        elem1.append(val[0])
                        elem2.append(valore_mancante)
                        elem3.append(valore_mancante)
                    else:
                        elem1.append(valore_mancante)
                        elem2.append(val[0])
                        elem3.append(valore_mancante)                        
                elif len(chiavi[i].split("_"))==2 :
                    elem1.append(val[0])
                    elem2.append(val[1])
                    elem3.append(valore_mancante)
                elif len(chiavi[i].split("_"))==3 :
                    elem1.append(val[0])
                    elem2.append(val[1])
                    elem3.append(val[2])                
                    
            df.insert(len(df.columns), "Zona", elem1)
            df.insert(len(df.columns), "Tipologia", elem2)
            df.insert(len(df.columns), "Tipo", elem3)
            
            #trasformarlo in .csv che abbia come separatore il ;
            df.to_csv(filename_fin, sep=";", index=False, index_label=False)
            
        return True
            
    except FileNotFoundError:
        print("Impossibile aprire il file: file non trovato o non esiste.")
        return False
            
    except PermissionError:
        print("Impossibile aprire il file: non si hanno i requisiti di accesso adeguati.")
        return False
    
    except IsADirectoryError:
        print("Impossibile aprire il file: operazione richiesta su una cartella e non su un file.")
        return False
    
def create_server_connection(host_name, user_name, user_password):
    connection = None
    try:
        connection = mysql.connector.connect(
            host=host_name,
            user=user_name,
            passwd=user_password
        )
        print("MariaDB Database connection successful")
    except Error as err:
        print(f"Error: '{err}'")
    return connection 

def create_database(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        print("Database created successfully")
    except Error as err:
        print(f"Error: '{err}'")
        
        
def create_db_connection(host_name, user_name, user_password, db_name):
    connection = None
    try:
        connection = mysql.connector.connect(
        host=host_name,
        user=user_name,
        passwd=user_password,
        database=db_name
    )
        print("MariaDB Database connection successful")
    except Error as err:
        print(f"Error: '{err}'")
    return connection      


def creare_tabella(connection, nome_tabella):
    
    mycursor = connection.cursor()
    
    mycursor.execute(("""CREATE TABLE %s 
                        (Id INT AUTO_INCREMENT PRIMARY KEY, 
                        Valore FLOAT(20,10) NOT NULL, 
                        Zona VARCHAR(10) NOT NULL,
                        Tipologia CHAR(255), 
                        Mercato CHAR(30) NOT NULL, 
                        Tipo CHAR(5), 
                        Ora FLOAT(1,0) NOT NULL, 
                        Giorno INT(2) NOT NULL, 
                        Mese INT(2) NOT NULL, 
                        Anno YEAR(4) NOT NULL, 
                        Da VARCHAR(10), 
                        A VARCHAR(10) )""") %(nome_tabella))
    print("Tabella creata con successo.")
        
def import_data_to_db(directory_csv, engine, cursor, valore_mancante, nome_tabella):
    
    filenames = glob.glob(directory_csv)
    
    
    for filename in filenames:
        df=pd.read_csv(filename, sep=';', header=0)
        record={"Mercato": [], "Ora": [], "Giorno": [], "Mese": [], "Anno": [], "Valore": [], "Zona": [], "Tipologia": [], "Tipo": [],  "Da": [], "A":[]}
        
        if "MPEG" in filename:
            continue
        elif "MTE" in filename:
            continue    
        elif "XBID" in filename:
            continue
        
        elif "Transiti" in filename:
            
            for row in df.itertuples(index=True, name=None):
                for i in range(1,len(row)-3):
                    #uso questa espressione di row[0] perché essendo visualizzati anche gli indici, sfrutto questa opportunità
                    if row[0]==0:
                        num=str(row[i])                                        
                        record["Giorno"].append(str(num[6:8]))
                        record["Mese"].append(str(num[4:6]))
                        record["Anno"].append(str(num[0:4]))
                    elif row[0]==1:
                        record["Mercato"].append(str(row[i])) 
                    elif row[0]==2:                    
                        record['Ora'].append(str(row[i]))                        
                    elif row[0]==3:
                        record["Da"].append(str(row[i]))                            
                    elif row[0]==4:
                        record["A"].append(str(row[i]))
                    else:
                        record["Valore"].append(row[i])
                        record["Tipologia"].append("Transiti")
                        
            lungh=len(record["Mercato"])
            if lungh==len(record["Ora"]) and lungh==len(record["Giorno"]) and lungh==len(record["Mese"]) and lungh==len(record["Anno"]):
                lunghezza_voluta=len(record["Valore"])
                molt=int(lunghezza_voluta/lungh)
                record["Mercato"]=record["Mercato"]*molt
                record["Ora"]=record["Ora"]*molt
                record["Giorno"]=record["Giorno"]*molt
                record["Mese"]=record["Mese"]*molt
                record["Anno"]=record["Anno"]*molt  
                
            DF1 = pd.DataFrame(dict([(key, (pd.Series(value))) for key, value in record.items()]))
            DF1.to_sql(name=nome_tabella, con=engine, if_exists='append', index=False)    
               
            
        elif "Prezzi" in filename:
            
            for row in df.itertuples(index=True, name=None):
                for i in range(1,len(row)-3): 
                    
                    if row[0]==0:
                        num=str(row[i])                                        
                        record["Giorno"].append(str(num[6:8]))
                        record["Mese"].append(str(num[4:6]))
                        record["Anno"].append(str(num[0:4]))
                    elif row[0]==1:
                        record["Mercato"].append(str(row[i])) 
                    elif row[0]==2:                    
                        record['Ora'].append(str(row[i]))                        
                    else:
                        record["Zona"].append(str(row[len(row)-3]))
                        record["Tipologia"].append("Prezzi")
                        record["Tipo"].append(valore_mancante)
                        record["Valore"].append(row[i])
                            
            lungh=len(record["Mercato"])
            if lungh==len(record["Ora"]) and lungh==len(record["Giorno"]) and lungh==len(record["Mese"]) and lungh==len(record["Anno"]):
                lunghezza_voluta=len(record["Valore"])
                molt=int(lunghezza_voluta/lungh)
                record["Mercato"]=record["Mercato"]*molt
                record["Ora"]=record["Ora"]*molt
                record["Giorno"]=record["Giorno"]*molt
                record["Mese"]=record["Mese"]*molt
                record["Anno"]=record["Anno"]*molt 
                
            DF1 = pd.DataFrame(dict([(key, (pd.Series(value))) for key, value in record.items()]))
            DF1.to_sql(name=nome_tabella, con=engine, if_exists='append', index=False)  
            
        else:
            
            for row in df.itertuples(index=True, name=None):
                for i in range(1,len(row)-3):
                          
                    if row[0]==0:
                        num=str(row[i])                                        
                        record["Giorno"].append(str(num[6:8]))
                        record["Mese"].append(str(num[4:6]))
                        record["Anno"].append(str(num[0:4]))
                    elif row[0]==1:
                        record["Mercato"].append(str(row[i]))
                    elif row[0]==2:                    
                        record['Ora'].append(str(row[i]))
                    else:
                        record["Valore"].append(row[i])
                        record["Zona"].append(str(row[len(row)-3]))
                        record["Tipologia"].append(str(row[len(row)-2]))
                        record["Tipo"].append(str(row[len(row)-1]))
                        
            lungh=len(record["Mercato"])
            if lungh==len(record["Ora"]) and lungh==len(record["Giorno"]) and lungh==len(record["Mese"]) and lungh==len(record["Anno"]):
                lunghezza_voluta=len(record["Valore"])
                molt=int(lunghezza_voluta/lungh)
                record["Mercato"]=record["Mercato"]*molt
                record["Ora"]=record["Ora"]*molt
                record["Giorno"]=record["Giorno"]*molt
                record["Mese"]=record["Mese"]*molt
                record["Anno"]=record["Anno"]*molt  
            
            DF1 = pd.DataFrame(dict([(key, (pd.Series(value))) for key, value in record.items()]))
            DF1.to_sql(name=nome_tabella, con=engine, if_exists='append', index=False) 
           
            
        with open("file_con_problemi.txt", "a") as f:
            f.write(filename)
            f.write(str(record))
            f.write("\n")            
            f.write("\n")            
            
            
        
def main():
    
    #GME_LINK="https://www.mercatoelettrico.org/It/download/DatiStorici.aspx"
    
    #chrome_driver= ChromeDriverManager().install() #installare automaticamente il software ChromeDriver
    ##software installato sul pc che serve per pilotare il browser
    #driver=Chrome(service=Service(chrome_driver)) #servizio che fa uso di questo software
    ##assegnare degli ordini al browser
    #driver.maximize_window() #massimizzare la grandezza della finestra
    #driver.get(GME_LINK) #aprire pagina web
    
    ##accettare condizioni
    #accetto_condizioni1=driver.find_element(By.ID, 'ContentPlaceHolder1_CBAccetto1').click()
    #accetto_condizioni2=driver.find_element(By.ID, 'ContentPlaceHolder1_CBAccetto2').click()
    #accetto_condizioni_fin=driver.find_element(By.ID, 'ContentPlaceHolder1_Button1').click()
    
    #day=1
    #year='2020'
    #nome_file_mancanti='dati_mancanti_2020.txt'
    
    #MGP_prezzi(driver, day, year, nome_file_mancanti)    
    #MGP_quantità(driver, day, year, nome_file_mancanti)
    #MGP_transiti(driver, day, year, nome_file_mancanti)
    
    #MI1_prezzi(driver, day, year, nome_file_mancanti)
    #MI1_quantità(driver, day, year, nome_file_mancanti) 
    #MI2_prezzi(driver, day, year, nome_file_mancanti)
    #MI2_quantità(driver, day, year, nome_file_mancanti)
    #MI3_prezzi(driver, day, year, nome_file_mancanti)
    #MI3_quantità(driver, day, year, nome_file_mancanti)
    #MI4_prezzi(driver, day, year, nome_file_mancanti)
    #MI4_quantità(driver, day, year, nome_file_mancanti)
    #MI5_prezzi(driver, day, year, nome_file_mancanti)
    #MI5_quantità(driver, day, year, nome_file_mancanti)
    #MI6_prezzi(driver, day, year, nome_file_mancanti)
    #MI6_quantità(driver, day, year, nome_file_mancanti)
    #MI7_prezzi(driver, day, year, nome_file_mancanti)
    #MI7_quantità(driver, day, year, nome_file_mancanti)
    #MI_A1_prezzi(driver, day, year, nome_file_mancanti)
    #MI_A1_quantità(driver, day, year, nome_file_mancanti)
    #MI_A2_prezzi(driver, day, year, nome_file_mancanti)
    #MI_A2_quantità(driver, day, year, nome_file_mancanti)
    #MI_A3_prezzi(driver, day, year, nome_file_mancanti)
    #MI_A3_quantità(driver, day, year, nome_file_mancanti)
    #MI_XBID_1FASE(driver, day, year, nome_file_mancanti)
    #MI_XBID_2FASE(driver, day, year, nome_file_mancanti)
    #MI_XBID_3FASE(driver, day, year, nome_file_mancanti)
    
    #MPEG(driver, day, year, nome_file_mancanti)

    #MSD_EX_ANTE(driver, day, year, nome_file_mancanti)
    #MSD_EX_POST(driver, day, year, nome_file_mancanti)
    #MSD_MB_PRELIMINARI_ALTRI_SERVIZI(driver, day, year, nome_file_mancanti)
    #MSD_MB_PRELIMINARI_RISERVA_SECONDARIA(driver, day, year, nome_file_mancanti)
    #MSD_MB_PRELIMINARI_TOTALI(driver, day, year, nome_file_mancanti)
    #MTE(driver, day, year, nome_file_mancanti)
    
    #extension='.zip'
    #src = r'C:\Users\valentina.mercurio\Downloads'
    #dest = r'C:\Users\valentina.mercurio\Desktop\Progetto\FILES2020'    
    #spostare_files(src, dest, extension)
    
    #estrarre_files(dest,extension) 
        
    #directory_xml=dest+"\*.xml"
    ##impostare il livello a cui sono presenti i figli da considerare e da cui prendere i dati
    #livello=1
    ##gestire gli eventuali valori mancanti
    #valore_mancante=None    
    #print(convertitore_xml_csv(directory_xml, valore_mancante, livello))
    
    #directory_csv=dest+"\*.csv" 
    
    #connection = create_server_connection("localhost", "root", "root")
    #cursor = connection.cursor()    

    #connection = create_db_connection("localhost", "root", 'root', 'dati_energia_elettrica') # Connect to the Database
    #engine = create_engine('mysql+pymysql://root:root@127.0.0.1/dati_energia_elettrica')
    
    #cursor = connection.cursor()             
    ##create_database_query="CREATE DATABASE DATI_ENERGIA_ELETTRICA"
    ##create_database(connection, create_database_query)
    
    #engine = create_engine('mysql+pymysql://root:root@127.0.0.1/dati_energia_elettrica')
    
    #nome_tabella='ee2020'   
    ##creare_tabella(connection, nome_tabella)
    
    #import_data_to_db(directory_csv, engine, cursor, valore_mancante, nome_tabella)
    
main()