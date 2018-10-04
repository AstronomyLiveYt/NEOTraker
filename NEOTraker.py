from tkinter import *
from tkinter import filedialog
import ephem
import math
import os
import sys
import time
import datetime
import re
import json
import geocoder
import serial
import win32com.client
from urllib.request import urlopen

class trackSettings:
    filetype = 'HORIZONS'
    orbitFile = ''
    tracking = False
    NSoffset = 0
    EWoffset = 0
    FileSelected = False
    telescopetype = 'LX200'
    Lat = 0.0
    Lon = 0.0


class buttons:
    def __init__(self, master):
        topframe = Frame(master)
        master.winfo_toplevel().title("NEOTraker")
        topframe.pack(side=TOP)
        bottomframe = Frame(master)
        bottomframe.pack(side=BOTTOM)
        self.menu = Menu(master)
        master.config(menu=self.menu)
              
        self.startButton = Button(topframe, text='Start/Stop Tracking', command=self.setTracking)
        self.startButton.grid(row=5, column = 2)
        self.northButton = Button(topframe, text='N', command=self.goNorth)
        self.northButton.grid(row=1, column=2)
        self.resetButton = Button(topframe, text='Reset', command=self.goReset)
        self.resetButton.grid(row=2, column=2)
        self.westButton = Button(topframe, text='W', command=self.goWest)
        self.westButton.grid(row=2, column=3)
        self.eastButton = Button(topframe, text='E', command=self.goEast)
        self.eastButton.grid(row=2, column=1)
        self.southButton = Button(topframe, text='S', command=self.goSouth)
        self.southButton.grid(row=3, column=2)
        
        self.entryNorth = Entry(topframe)
        self.entrySouth = Entry(topframe)
        self.entryEast = Entry(topframe)
        self.entryWest = Entry(topframe)
        self.entryNorth.grid(row=0, column = 2)
        self.entryEast.grid(row=2, column = 0)
        self.entrySouth.grid(row=4, column = 2)
        self.entryWest.grid(row=2, column = 4)
        self.entryNorth.insert(0, 0)
        self.entryEast.insert(0, 0)
        self.entrySouth.insert(0, 0)
        self.entryWest.insert(0, 0)
        
        self.labelLat = Label(topframe, text='Latitude (N+)')
        self.labelLat.grid(row=0, column = 5)
        self.labelLon = Label(topframe, text='Longitude (E+)')
        self.labelLon.grid(row=1, column = 5)
        self.entryLat = Entry(topframe)
        self.entryLat.grid(row = 0, column = 6)
        self.entryLon = Entry(topframe)
        self.entryLon.grid(row = 1, column = 6)
        try:
            config = open('config.txt', 'r')
            clines = [line.rstrip('\n') for line in config]
            trackSettings.filetype = str(clines[0])
            trackSettings.telescopetype = str(clines[1])
            trackSettings.Lat = float(clines[3])
            trackSettings.Lon = float(clines[4])
            config.close()
        except:
            print('Config file not present or corrupted.')
        try:
            geolocation = geocoder.ip('me')
            self.entryLat.insert(0, geolocation.latlng[0])
            self.entryLon.insert(0, geolocation.latlng[1])
        except:
            self.entryLat.insert(0, trackSettings.Lat)
            self.entryLon.insert(0, trackSettings.Lon)
        self.comLabel = Label(topframe, text='COM Port')
        self.comLabel.grid(row = 2, column = 5)
        self.entryCom = Entry(topframe)
        self.entryCom.grid(row = 2, column = 6)
        try:
            self.entryCom.insert(0, clines[2])
        except:
            self.entryCom.insert(0, 0)
        
        
        self.fileMenu = Menu(self.menu)
        self.menu.add_cascade(label='File', menu=self.fileMenu)
        self.fileMenu.add_command(label='Select Orbit Data File...', command=self.filePicker)
        self.fileMenu.add_separator()
        self.fileMenu.add_command(label='Exit and Save Configuration', command=self.exitProg)
        
        self.typeMenu = Menu(self.menu)
        self.menu.add_cascade(label='File Type', menu=self.typeMenu)
        self.typeMenu.add_command(label='JPL HORIZONS', command=self.setHorizons)
        self.typeMenu.add_command(label='FindOrb', command=self.setFindOrb)
        
        self.telescopeMenu = Menu(self.menu)
        self.menu.add_cascade(label='Telescope Type', menu=self.telescopeMenu)
        self.telescopeMenu.add_command(label='LX200 Classic', command=self.setLX200)
        self.telescopeMenu.add_command(label='ASCOM', command=self.setASCOM)
    
    def setTracking(self):            
        if trackSettings.tracking is False and trackSettings.FileSelected is True:
            trackSettings.tracking = True
            #Connect either by LX200 or ASCOM
            if trackSettings.telescopetype == 'LX200':
                try:
                    self.comport = str('COM'+str(self.entryCom.get()))
                    self.ser = serial.Serial(self.comport, 9600)
                    self.ser.write(str.encode(':U#'))
                    self.serialconnected = True
                except:
                    print('Failed to connect on ' + self.comport)
                    trackSettings.tracking = False
                    return
            elif trackSettings.telescopetype == 'ASCOM':
                self.x = win32com.client.Dispatch("ASCOM.Utilities.Chooser")
                self.x.DeviceType = 'Telescope'
                driverName=self.x.Choose("None")
                self.tel=win32com.client.Dispatch(driverName)
                if self.tel.Connected:
                    print("Telescope was already connected")
                else:
                    self.tel.Connected = True
                    if self.tel.Connected:
                        print("Connected to telescope now")
                    else:
                        print("Unable to connect to telescope, expect exception")
            observer = ephem.Observer()
            if trackSettings.filetype == 'HORIZONS':
                with open(trackSettings.orbitFile) as f:
                    lines = [line.rstrip('\n') for line in f]
                for idx, line in enumerate(lines):
                    if "$$SOE" in line:
                        #nameline = lines[29].split('(')[1].split(')')[0]
                        targetname = str('target')
                        line1 = lines[idx+1]
                        line2 = lines[idx+2]
                        line3 = lines[idx+3]
                        line4 = lines[idx+4]
                        line5 = lines[idx+5] 
                        linesplit1 = line1.split(' ')
                        self.dateline = float(linesplit1[0]) - 2415020
                        observer.date = self.dateline
                        datesplit = str(observer.date).split('/')
                        year = datesplit[0]
                        month = datesplit[1]
                        day = float(datesplit[2].split(' ')[0])
                        fractionday = str('0.'+str(linesplit1[0].split('.')[1]))
                        fractionday = float(fractionday) - 0.5
                        if fractionday < 0:
                            fractionday = 1 + fractionday
                        day = day + fractionday
                        xephemdate = str(str(month) + '/' + str(day) + '/' + str(year))
                        self.ec = float(line2[4:26])
                        self.qr = float(line2[30:52])
                        self.inc = float(line2[56:78])

                        self.om = float(line3[4:26])
                        self.w = float(line3[30:52])
                        self.tp = float(line3[56:78])

                        self.n = float(line4[4:26])
                        self.ma = float(line4[30:52])
                        self.ta = float(line4[56:78])

                        self.a = float(line5[4:26])
                        self.ad = float(line5[30:52])
                        self.pr = float(line5[56:78])
                        if self.ec<1:
                            self.xephem = str(targetname + ',' + 'e' + ',' + str(self.inc) + ',' + str(self.om) + ',' + str(self.w) + ',' + str(self.a) + ',' + str(self.n) + ',' + str(self.ec) + ',' + str(self.ma) + ',' + xephemdate + ',' + '2000' + ',' + 'g  6.5,4.0')
                        else:
                            self.dateline = float(self.tp) - 2415020
                            observer.date = self.dateline
                            datesplit = str(observer.date).split('/')
                            year = datesplit[0]
                            month = datesplit[1]
                            day = float(datesplit[2].split(' ')[0])
                            fractionday = str('0.'+str(linesplit1[0].split('.')[1]))
                            fractionday = float(fractionday) - 0.5
                            if fractionday < 0:
                                fractionday = 1 + fractionday
                            day = day + fractionday
                            xephemdate = str(str(month) + '/' + str(day) + '/' + str(year))
                            self.xephem = str(targetname + ',' + 'h' + ',' + xephemdate + ',' + str(self.inc) + ',' + str(self.om) + ',' + str(self.w) + ',' + str(self.ec) + ',' + str(self.qr) + ',' + '2000' + ',' + 'g  6.5,4.0')
            elif trackSettings.filetype == 'FindOrb':
                with open(trackSettings.orbitFile) as f:
                    lines = [line.rstrip('\n') for line in f]
                for idx, line in enumerate(lines):
                    if "Epoch" in line:
                        targetname = str('Target')
                        line0 = lines[idx-2]
                        line1 = lines[idx]
                        line2 = lines[idx+2]
                        line3 = lines[idx+4]
                        line4 = lines[idx+6]
                        line5 = lines[idx+8]
                        linesplit1 = line1.split('JDT ')
                        linesplit1 = linesplit1[1].split(' ')[0]
                        self.dateline = float(linesplit1) - 2415020
                        observer.date = self.dateline
                        datesplit = str(observer.date).split('/')
                        year = datesplit[0]
                        month = datesplit[1]
                        day = float(datesplit[2].split(' ')[0])
                        fractionday = str('0.'+str(linesplit1.split('.')[1]))
                        fractionday = float(fractionday) - 0.5
                        if fractionday < 0:
                            fractionday = 1 + fractionday
                        day = day + fractionday
                        xephemdate = str(str(month) + '/' + str(day) + '/' + str(year))
                        linesplit2 = line5[1:14]
                        self.ec = float(linesplit2)
                        if self.ec<1:
                            linesplit2 = line3[1:14]
                            self.n = float(linesplit2)
                            linesplit2 = line3[25:35]
                            self.w = float(linesplit2)
                            linesplit2 = line4[1:14]
                            self.a = float(linesplit2)
                            linesplit2 = line4[25:35]
                            self.om = float(linesplit2)
                            linesplit2 = line5[25:35]
                            self.inc = float(linesplit2)
                            linesplit2 = line2[1:11]
                            self.ma = float(linesplit2)
                            self.xephem = str(targetname + ',' + 'e' + ',' + str(self.inc) + ',' + str(self.om) + ',' + str(self.w) + ',' + str(self.a) + ',' + str(self.n) + ',' + str(self.ec) + ',' + str(self.ma) + ',' + xephemdate + ',' + '2000' + ',' + 'g  6.5,4.0')
                        else:
                            self.tp = line0.split('JD')[1].split(')')[0]
                            self.dateline = float(self.tp) - 2415020
                            observer.date = self.dateline
                            datesplit = str(observer.date).split('/')
                            year = datesplit[0]
                            month = datesplit[1]
                            day = float(datesplit[2].split(' ')[0])
                            fractionday = str('0.'+str(self.tp.split('.')[1]))
                            fractionday = float(fractionday) - 0.5
                            if fractionday < 0:
                                fractionday = 1 + fractionday
                            day = day + fractionday
                            xephemdate = str(str(month) + '/' + str(day) + '/' + str(year))
                            linesplit2 = line3[25:35]
                            self.w = float(linesplit2)
                            linesplit2 = line4[25:35]
                            self.om = float(linesplit2)
                            linesplit2 = line5[25:35]
                            self.inc = float(linesplit2)
                            linesplit2 = line2[1:14]
                            self.qr = float(linesplit2)
                            self.xephem = str(targetname + ',' + 'h' + ',' + xephemdate + ',' + str(self.inc) + ',' + str(self.om) + ',' + str(self.w) + ',' + str(self.ec) + ',' + str(self.qr) + ',' + '2000' + ',' + 'g  6.5,4.0')
                            print(self.xephem)
            self.firstslew = True
            self.doTracking()
        else:
            if trackSettings.telescopetype == 'LX200' and self.serialconnected is True:
                self.ser.write(str.encode(':Q#'))
                self.ser.write(str.encode(':U#'))
                self.ser.close()
                self.serialconnected = False
            elif trackSettings.telescopetype == 'ASCOM':
                self.tel.Connected = False
            trackSettings.tracking = False
            
    def rad_to_sexagesimal(self):
        self.radeg = math.degrees(self.radra)
        self.decdeg = math.degrees(self.raddec)
        self.ra_h = math.trunc((self.radeg)/15)
        self.ra_m = math.trunc((((self.radeg)/15) - self.ra_h)*60)
        self.ra_s = (((((self.radeg)/15) - self.ra_h)*60) - self.ra_m)*60
        
        self.dec_d = math.trunc(self.decdeg)
        self.dec_m = math.trunc((abs(self.decdeg) - abs(self.dec_d))*60)
        self.dec_s = (((abs(self.decdeg) - abs(self.dec_d))*60) - abs(self.dec_m))*60
    
    def doTracking(self):
        if trackSettings.tracking is True:
            observer = ephem.Observer()
            d = datetime.datetime.utcnow()
            observer.date = d
            target = ephem.readdb(self.xephem)
            observer.lat = str(self.entryLat.get())
            observer.lon = str(self.entryLon.get())
            observer.elevation = 0
            observer.pressure = 1013
            target.compute(observer)
            targetra = target.ra
            targetdec = target.dec
            targetdec = targetdec + math.radians((trackSettings.NSoffset/3600))
            targetra = targetra + math.radians((trackSettings.EWoffset/3600))
            self.radra = targetra
            self.raddec = targetdec
            self.rad_to_sexagesimal()
            targetcoord = str(str(self.ra_h)+':'+str(self.ra_m)+':'+"{0:.2f}".format(round(self.ra_s,2))+' '+str(self.dec_d)+':'+str(self.dec_m)+':'+"{0:.2f}".format(round(self.dec_s,2)))   
            if target.alt < 0:
                print('Object below the horizon, stopping tracking')
                self.setTracking()
                time.sleep(1)
            if trackSettings.telescopetype == 'LX200' and target.alt > 0:
                targetcoordra = str(':Sr ' + str(self.ra_h)+':'+str(self.ra_m)+':'+str(int(self.ra_s))+'#')
                targetcoorddec = str(':Sd ' + str(self.dec_d)+'*'+str(self.dec_m)+':'+str(int(self.dec_s))+'#')
                print(targetcoordra + ' ' + targetcoorddec)
                self.ser.write(str.encode(targetcoordra))
                self.ser.write(str.encode(targetcoorddec))
                self.ser.write(str.encode(':MS#'))
            elif trackSettings.telescopetype == 'ASCOM' and target.alt > 0:
                if self.firstslew is True:
                    observer.date = (d + datetime.timedelta(seconds=1))
                    target.compute(observer)
                    targetra2 = target.ra
                    targetdec2 = target.dec
                    targetdec2 = targetdec2 + math.radians((trackSettings.NSoffset/3600))
                    targetra2 = targetra2 + math.radians((trackSettings.EWoffset/3600))
                    rarate = (math.degrees(targetra2 - targetra))*math.cos(targetdec2)
                    decrate = math.degrees(targetdec2 - targetdec)
                    self.tel.MoveAxis(0, rarate)
                    self.tel.MoveAxis(1, decrate)
                    targetrahours = float((math.degrees(targetra2)/15))
                    print('Slewing to RA hours: '+ str("{0:.4f}".format(round(targetrahours,4))) + ' Dec degrees: ' + str(float(math.degrees(targetdec))))
                    self.tel.SlewToCoordinates(targetrahours,float(math.degrees(targetdec)))
                else:
                    observer.date = (d + datetime.timedelta(seconds=1))
                    target.compute(observer)
                    targetra2 = target.ra
                    targetdec2 = target.dec
                    targetdec2 = targetdec2 + math.radians((trackSettings.NSoffset/3600))
                    targetra2 = targetra2 + math.radians((trackSettings.EWoffset/3600))
                    rarate = (math.degrees(targetra2 - targetra))*math.cos(targetdec2)
                    decrate = math.degrees(targetdec2 - targetdec)
                    print('RA Rate: ' + str(math.degrees(targetra2 - targetra)*3600*60*60*math.cos(targetdec2)) + ' arcseconds per hour.  Dec Rate: ' + str(math.degrees(targetdec2 - targetdec)*3600*60*60) + ' arcseconds per hour.')
                    self.tel.MoveAxis(0, rarate)
                    self.tel.MoveAxis(1, decrate)
            #print(targetcoord, end='\r')
            self.firstslew = False
            if trackSettings.telescopetype == 'LX200':
                root.after(10,self.doTracking)      
            elif trackSettings.telescopetype == 'ASCOM':
                root.after(1000,self.doTracking)      
    
    def goNorth(self):
        trackSettings.NSoffset = float(self.entryNorth.get())
        trackSettings.NSoffset += 1
        self.entryNorth.delete(0, END)
        self.entryNorth.insert(0, trackSettings.NSoffset)
        self.entrySouth.delete(0, END)
        self.entrySouth.insert(0, (-1*trackSettings.NSoffset))
        
    def goWest(self):
        trackSettings.EWoffset = float(self.entryWest.get())*-1
        trackSettings.EWoffset -= 1
        self.entryEast.delete(0, END)
        self.entryEast.insert(0, trackSettings.EWoffset)
        self.entryWest.delete(0, END)
        self.entryWest.insert(0, (-1*trackSettings.EWoffset))
        
    def goSouth(self):
        trackSettings.NSoffset = float(self.entrySouth.get())*-1
        trackSettings.NSoffset -= 1
        self.entryNorth.delete(0, END)
        self.entryNorth.insert(0, trackSettings.NSoffset)
        self.entrySouth.delete(0, END)
        self.entrySouth.insert(0, (-1*trackSettings.NSoffset))
        
    def goEast(self):
        trackSettings.EWoffset = float(self.entryEast.get())
        trackSettings.EWoffset += 1
        self.entryEast.delete(0, END)
        self.entryEast.insert(0, trackSettings.EWoffset)
        self.entryWest.delete(0, END)
        self.entryWest.insert(0, (-1*trackSettings.EWoffset))
        
    def goReset(self):
        trackSettings.EWoffset = 0
        trackSettings.NSoffset = 0
        self.entryNorth.delete(0, END)
        self.entryNorth.insert(0, trackSettings.NSoffset)
        self.entrySouth.delete(0, END)
        self.entrySouth.insert(0, trackSettings.NSoffset)
        self.entryEast.delete(0, END)
        self.entryEast.insert(0, trackSettings.EWoffset)
        self.entryWest.delete(0, END)
        self.entryWest.insert(0, trackSettings.EWoffset)
    
    def setHorizons(self):
        trackSettings.filetype = 'HORIZONS'
    
    def setFindOrb(self):
        trackSettings.filetype = 'FindOrb'
    
    def setLX200(self):
        trackSettings.telescopetype = 'LX200'
        
    def setASCOM(self):
        trackSettings.telescopetype = 'ASCOM'
    
    def filePicker(self):
        trackSettings.orbitFile = filedialog.askopenfilename(initialdir = ".",title = "Select file",filetypes = (("text files","*.txt"),("all files","*.*")))
        trackSettings.FileSelected = True
        print(trackSettings.orbitFile)
        
    def exitProg(self):
        config = open('config.txt','w')
        config.write(str(trackSettings.filetype)+'\n')
        config.write(str(trackSettings.telescopetype)+'\n')
        config.write(str(self.entryCom.get()) + '\n')
        config.write(str(self.entryLat.get())+'\n')
        config.write(str(self.entryLon.get()))
        config.close()
        exit()

root = Tk()
b = buttons(root)
root.mainloop()
