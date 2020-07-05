from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox, QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QMainWindow, QTableWidgetItem, QVBoxLayout, QHBoxLayout, QProgressBar
from PyQt5.QtCore import Qt
import os
import pandas as pd
import numpy as np
import xml.dom.minidom 
import zipfile


messageRows = ""
sitelistErrorDetail = ""
numRowsFault = ""
sitelist = pd.DataFrame()
global_string = ""
sitesChk = True
sectorsChk = True
fullName = ""
name = ""
listval = 0

class Ui_MainWindow(QMainWindow):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(900, 600)
        self.setWindowIcon(QtGui.QIcon('tower.png'))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.labelInfo = QtWidgets.QLabel(self.centralwidget)
        self.labelInfo.setGeometry(QtCore.QRect(9, 9, 550, 16))
        self.labelInfo.setObjectName("labelInfo")
        self.table1 = QtWidgets.QTableWidget(self.centralwidget)
        self.table1.setGeometry(QtCore.QRect(30, 30, 814, 55))
        self.table1.setObjectName("table1")
        self.table1.setColumnCount(8)
        self.table1.setRowCount(1)
        self.table1.verticalHeader().hide()
        table_data = ['1', '1A', 'name', 'GSM', '120', '-30.12345', '30.12345', '65']
        self.table1.setHorizontalHeaderLabels(['SITEID', 'CELLID', 'SITENAME', 'TECH', 'AZIMUTH', 'LAT', 'LON', 'BW'])
        for item in range(0, 8):
            self.table1.setItem(0,item,QTableWidgetItem(str(table_data[item])))
        self.labelImport = QtWidgets.QLabel(self.centralwidget)
        self.labelImport.setGeometry(QtCore.QRect(13, 222, 98, 23))
        self.labelImport.setObjectName("labelImport")
        self.chkSites = QtWidgets.QCheckBox(self.centralwidget)
        self.chkSites.setGeometry(QtCore.QRect(11, 250, 301, 21))
        self.chkSites.setObjectName("chkSites")
        self.chkSites.setChecked(True)
        self.chkSectors = QtWidgets.QCheckBox(self.centralwidget)
        self.chkSectors.setGeometry(QtCore.QRect(11, 290, 261, 21))
        self.chkSectors.setObjectName("chkSectors")
        self.chkSectors.setChecked(True)
        self.labelExport = QtWidgets.QLabel(self.centralwidget)
        self.labelExport.setGeometry(QtCore.QRect(12, 321, 97, 23))
        self.labelExport.setObjectName("labelExport")
        self.btnImport = QtWidgets.QPushButton(self.centralwidget)
        self.btnImport.setGeometry(QtCore.QRect(340, 220, 85, 23))
        self.btnImport.setObjectName("btnImport")
        self.btnExport = QtWidgets.QPushButton(self.centralwidget)
        self.btnExport.setGeometry(QtCore.QRect(340, 321, 85, 23))
        self.btnExport.setObjectName("btnExport")
        self.btnCreate = QtWidgets.QPushButton(self.centralwidget)
        self.btnCreate.setEnabled(False)
        self.btnCreate.setGeometry(QtCore.QRect(11, 410, 101, 26))
        self.btnCreate.setObjectName("btnCreate")
        self.pbar = QtWidgets.QProgressBar(self.centralwidget)
        self.pbar.setGeometry(QtCore.QRect(200, 410, 500, 25))  
        self.pbar.setTextVisible(True)
        self.labelProgress = QtWidgets.QLabel(self.centralwidget)
        self.labelProgress.setGeometry(QtCore.QRect(300, 380, 300, 25))
        self.labelProgress.setObjectName("labelProgress")        
        self.line1 = QtWidgets.QFrame(self.centralwidget)
        self.line1.setGeometry(QtCore.QRect(10, 190, 781, 16))
        self.line1.setFrameShape(QtWidgets.QFrame.HLine)
        self.line1.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line1.setObjectName("line1")
        self.line2 = QtWidgets.QFrame(self.centralwidget)
        self.line2.setGeometry(QtCore.QRect(10, 370, 781, 16))
        self.line2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line2.setObjectName("line2")
        self.importLoc = QtWidgets.QLineEdit(self.centralwidget)
        self.importLoc.setGeometry(QtCore.QRect(115, 220, 211, 23))
        self.importLoc.setAutoFillBackground(True)
        self.importLoc.setText("")
        self.importLoc.setObjectName("importLoc")
        self.exportLoc = QtWidgets.QLineEdit(self.centralwidget)
        self.exportLoc.setGeometry(QtCore.QRect(115, 321, 211, 23))
        self.exportLoc.setAutoFillBackground(False)
        self.exportLoc.setText("")
        self.exportLoc.setObjectName("exportLoc")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 24))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionOpen = QtWidgets.QAction(MainWindow)
        self.actionOpen.setObjectName("actionOpen")
        self.actionExport = QtWidgets.QAction(MainWindow)
        self.actionExport.setObjectName("actionExport")
        self.actionExit = QtWidgets.QAction(MainWindow)
        self.actionExit.setObjectName("actionExit")
        self.actionAbout = QtWidgets.QAction(MainWindow)
        self.actionAbout.setObjectName("actionAbout")
        self.menuFile.addAction(self.actionOpen)
        self.menuFile.addAction(self.actionExport)
        self.menuFile.addSeparator()
        self.menuFile.addAction(self.actionExit)
        self.menuHelp.addAction(self.actionAbout)
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "KML Creator 2.0"))
        
        self.labelInfo.setText(_translate("MainWindow", "When Importing a file, please ensure the column headings are as follows (not case sensitive):"))
        self.labelImport.setText(_translate("MainWindow", "Import Location:"))
        self.chkSites.setStatusTip(_translate("MainWindow", "Select if Site waypoints need to be created"))
        self.chkSites.setText(_translate("MainWindow", "Create Sites (Only unique siteids will be used)"))
        self.chkSectors.setStatusTip(_translate("MainWindow", "Select if Sector polygons need to be created"))
        self.chkSectors.setText(_translate("MainWindow", "Create Sectors (Grouped by technology"))
        self.labelExport.setText(_translate("MainWindow", "Export Location:"))
        self.btnImport.setStatusTip(_translate("MainWindow", "Import CSV file from"))
        self.btnImport.setText(_translate("MainWindow", "Browse"))
        self.btnExport.setStatusTip(_translate("MainWindow", "Export KML to"))
        self.btnExport.setText(_translate("MainWindow", "Browse"))
        self.btnCreate.setStatusTip(_translate("MainWindow", "Create KML file, only available if import and export locations are selected"))
        self.btnCreate.setText(_translate("MainWindow", "Create KML"))
        self.labelProgress.setText(_translate("MainWindow", "Progress Info"))
        self.importLoc.setStatusTip(_translate("MainWindow", "Location to import CSV from"))
        self.exportLoc.setStatusTip(_translate("MainWindow", "Location to Save KML in"))
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.menuHelp.setTitle(_translate("MainWindow", "Help"))
        self.actionOpen.setText(_translate("MainWindow", "Open"))
        self.actionExport.setText(_translate("MainWindow", "Export"))
        self.actionExit.setText(_translate("MainWindow", "Exit"))
        self.actionAbout.setText(_translate("MainWindow", "About"))
        self.chkSites.stateChanged.connect(self.sitesChecked)
        self.chkSectors.stateChanged.connect(self.sectorsChecked)
        self.actionExit.triggered.connect(self.close_application)
        self.btnImport.clicked.connect(self.file_open)
        self.actionOpen.triggered.connect(self.file_open)
        self.btnExport.clicked.connect((self.file_save))
        self.actionExport.triggered.connect((self.file_save))
        self.btnCreate.clicked.connect(self.exportKml)
        
    def sitesChecked(self):
        global sitesChk
        sitesChk = self.chkSites.isChecked()
        
    def sectorsChecked(self):
        global sectorsChk
        sectorsChk = self.chkSectors.isChecked()

    def close_application(self):
        box = QMessageBox()
        box.setIcon(QMessageBox.Question)
        box.setWindowTitle("Quit")
        box.setText("Are you sure you want to quit?")
        box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        box.setDefaultButton(QMessageBox.No)
        buttonYes = box.button(QMessageBox.Yes)
        buttonYes.setText("Yes")
        buttonNo = box.button(QMessageBox.No)
        buttonNo.setText("No")
        box.exec_()
        if box.clickedButton() == buttonYes:
            sys.exit()        


      
    
       
    def file_open(self):
        try:
            fileLocation = QFileDialog.getOpenFileName(self, "Open File location", "", "All Files(*)")
            self.importLoc.setText(fileLocation[0])
            self.importLoc.setAlignment(Qt.AlignRight)
            self.checkFileValid()
        except:
            box = QMessageBox()
            box.setIcon(QMessageBox.Information)
            box.setWindowTitle("Error Opening File")
            box.setText("There was an error loading the file.\nPlease try again")
            box.exec_()
            
            
                
        if self.importLoc.text() and self.exportLoc.text():
            self.btnCreate.setEnabled(True)
            


    def file_save(self):
        global fullName
        global name
        try:
            name, extention = QFileDialog.getSaveFileName(self, "Save file as...", "", ".kml")
            fullName = name + extention
            file = open(fullName, 'w')
            
            self.exportLoc.setText(name + extention)
            self.exportLoc.setAlignment(Qt.AlignRight)
            if self.importLoc.text() and self.exportLoc.text():
                self.btnCreate.setEnabled(True)
        except IOError:
            pass


    def checkFileValid(self):
        global sitelist
        try:
            # sitelist=pd.read_csv(self.importLoc.text())
            sitelist=pd.read_excel(self.importLoc.text())
        except:
            # sitelist=pd.read_excel(self.importLoc.text()) 
            sitelist=pd.read_csv(self.importLoc.text())   
        sitelist=sitelist.replace(regex=['&'], value='&amp;')
        sitelist.columns = map(str.lower, sitelist.columns)
        sitelist['lat'] = pd.to_numeric(sitelist['lat'], errors='coerce')
        sitelist['lon'] = pd.to_numeric(sitelist['lon'], errors='coerce')
        sitelist['azimuth'] = pd.to_numeric(sitelist['azimuth'], errors='coerce')
        sitelist['bw'] = pd.to_numeric(sitelist['bw'], errors='coerce')
        global sitelistErrorDetail
        global numRowsFault
        sitelistErrorDetail = sitelist[sitelist.index.isin(sitelist.index[sitelist.isnull().any(axis=1)])]
        global messageRows
        numRows = sitelist.shape[0]
        numRowsFault = sitelist.index[sitelist.isnull().any(axis=1)].shape[0]
        messageRows = "You have "+ str(numRowsFault) + " rows with incorrectly formatted data, out of " + str(numRows) + " rows of data.\nYou can re-import a file, or continue with the export and exclude the below rows.\nIncorrectly formatted fields are indicated by NaN below:"
        if numRowsFault != 0:
            self.confirm = confirmInput()
        sitelist = sitelist.dropna()
        
    def exportKml(self):
        # Remove duplicate site entries
        sites=sitelist.drop_duplicates('siteid')[['siteid', 'sitename','lat','lon']]
        # Create technology lists
        techs=sitelist.drop_duplicates("tech")[['tech']]
        # Create image length list to append to tech list, starting from largest to smallest
        self.pbar.setMaximum(len(techs)+1)
        i=0
        tech_index = []
        for item in techs.iterrows():
            tech_index.append(0.003 - i*0.0003)
            i+=1
        # Add image length list to techs list
        techs['img_len'] = tech_index
        # Create colour list to append to tech list
        k=0
        colour_index = []
        for item in techs.iterrows():
            colour_index.append(k + 70)
            k+=1
        # Add colour list to tech list
        techs['color'] = colour_index
        # Create default XML script for first part of XML file
        part1 = '<?xml version="1.0" encoding="UTF-8"?><kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2" xmlns:kml="http://www.opengis.net/kml/2.2" xmlns:atom="http://www.w3.org/2005/Atom"><Document><name>Exported KML</name><open>1</open><Style id="s_ylw-pushpin6"><IconStyle><scale>1.1</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>cc00ff00</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><Style id="s_ylw-pushpin60"><IconStyle><scale>1.1</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>ccffff55</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><Style id="s_ylw-pushpin_hl1"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>cc0055ff</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><Style id="s_ylw-pushpin61"><IconStyle><scale>1.1</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>cc7f00ff</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><Style id="s_ylw-pushpin_hl3"><IconStyle><scale>0.472727</scale><Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle_highlight.png</href></Icon></IconStyle><LabelStyle><scale>0.9</scale></LabelStyle><LineStyle><color>cc00ff00</color><width>1.5</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><StyleMap id="m_ylw-pushpin72"><Pair><key>normal</key><styleUrl>#s_ylw-pushpin63</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#s_ylw-pushpin_hl1</styleUrl></Pair></StyleMap><Style id="s_ylw-pushpin"><IconStyle><scale>0.4</scale><Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href></Icon></IconStyle><LabelStyle><scale>0.9</scale></LabelStyle><LineStyle><color>cc00ff00</color><width>1.5</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><Style id="s_ylw-pushpin1"><IconStyle><scale>1.1</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>cc00ffff</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><Style id="s_ylw-pushpin62"><IconStyle><scale>1.1</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>cc0000ff</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><StyleMap id="m_ylw-pushpin3"><Pair><key>normal</key><styleUrl>#s_ylw-pushpin</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#s_ylw-pushpin_hl3</styleUrl></Pair></StyleMap><StyleMap id="m_ylw-pushpin76"><Pair><key>normal</key><styleUrl>#s_ylw-pushpin61</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#s_ylw-pushpin_hl14</styleUrl></Pair></StyleMap><Style id="s_ylw-pushpin_hl10"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>cc00ff00</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><StyleMap id="m_ylw-pushpin74"><Pair><key>normal</key><styleUrl>#s_ylw-pushpin64</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#s_ylw-pushpin_hl13</styleUrl></Pair></StyleMap><Style id="s_ylw-pushpin_hl11"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>ccffff55</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><Style id="s_ylw-pushpin_hl12"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>cc0000ff</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><Style id="s_ylw-pushpin_hl13"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>cc7f0000</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><StyleMap id="m_ylw-pushpin75"><Pair><key>normal</key><styleUrl>#s_ylw-pushpin60</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#s_ylw-pushpin_hl11</styleUrl></Pair></StyleMap><Style id="s_ylw-pushpin_hl14"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>cc7f00ff</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><StyleMap id="m_ylw-pushpin71"><Pair><key>normal</key><styleUrl>#s_ylw-pushpin1</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#s_ylw-pushpin_hl</styleUrl></Pair></StyleMap><Style id="s_ylw-pushpin63"><IconStyle><scale>1.1</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>cc0055ff</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><StyleMap id="m_ylw-pushpin73"><Pair><key>normal</key><styleUrl>#s_ylw-pushpin62</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#s_ylw-pushpin_hl12</styleUrl></Pair></StyleMap><Style id="s_ylw-pushpin64"><IconStyle><scale>1.1</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>cc7f0000</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><StyleMap id="m_ylw-pushpin70"><Pair><key>normal</key><styleUrl>#s_ylw-pushpin6</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#s_ylw-pushpin_hl10</styleUrl></Pair></StyleMap><Style id="s_ylw-pushpin_hl"><IconStyle><scale>1.3</scale><Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon><hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/></IconStyle><LineStyle><color>cc00ffff</color><width>2</width></LineStyle><PolyStyle><fill>0</fill></PolyStyle></Style><Folder><name>Operational</name><open>1</open>'
        part2 = '<Folder><name>Sitename</name><open>1</open>'
        part3 = '<Placemark><name>'
        part4 = '</name><description><![CDATA[<br><br><br>    <table border="1" padding="0">]]></description><styleUrl>#m_ylw-pushpin3</styleUrl><Point><extrude>1</extrude><altitudeMode>relativeToGround</altitudeMode><coordinates>'
        part5 = '</coordinates></Point></Placemark>'
        part6 = '</Folder>'
        part7 = '</Folder></Document></kml>'
        secpart1 = '<Folder><name>'
        secpart2 = '</name>'
        secpart3 = '<Placemark><name>'
        secpart4 = '</name><styleUrl>#m_ylw-pushpin'
        secpart5 = '</styleUrl><Polygon><extrude>1</extrude><tessellate>1</tessellate><outerBoundaryIs><LinearRing><coordinates>'
        secpart6 = '</coordinates></LinearRing></outerBoundaryIs></Polygon></Placemark>'
        secpart7 = '</Folder>'

        listval = 1
        for index, data2 in techs.iterrows():
            self.labelProgress.setText("Progress Info - Coverting " + data2['tech'] + " Cells")
            self.pbar.setFormat(str(round((listval)/(len(techs)+1)*100,0))+ "%")
            self.pbar.setValue(listval)
            listval += 1
            placemark_str = ''
            for index, data in sites.iterrows():
                part3aa = [str(data['siteid']), data['sitename']]
                part3a = "-".join(part3aa)
                part4aa = [str(data['lon']), str(data['lat'])]
                part4a = ",".join(part4aa)
                placemark_list = [placemark_str, part3, part3a, part4, part4a, part5]
                placemark_str = "".join(placemark_list)  
            sector_str = ''
            for index, data in sitelist[sitelist.tech == data2['tech']].iterrows():
                sector_split = []
                azimuth_split = []
                for  i in np.arange(0,data['bw'], 15):
                    azimuth_split.append(data['azimuth'] - data['bw']/2 +(i))
                azimuth_split.append(data['azimuth'] + data['bw']/2)
                if data['bw'] != 360:
                    sector_split.append(data['lon'])
                    sector_split.append(data['lat'])
                    sector_split.append(0)
                for i in azimuth_split:
                    sector_split.append(data['lon'] + np.sin(np.radians(i))*data2['img_len'])
                    sector_split.append(data['lat'] + np.cos(np.radians(i))*data2['img_len'])
                    sector_split.append(0)
    
                if data['bw'] != 360:
                    sector_split.append(data['lon'])
                    sector_split.append(data['lat'])
                    sector_split.append(0)
                sector_split = str(sector_split).replace('[','').replace('[','').replace(' ','')
                sector_list = [sector_str, secpart3, str(data['cellid']), secpart4, str(data2['color']), secpart5, str(sector_split), secpart6]
                sector_str = "".join(sector_list)
            self.labelProgress.setText("Progress Info - Compiling final file")
            sector_list2 = [secpart1, str(data['tech']), secpart2, sector_str, secpart7]
            sector_str = "".join(sector_list2)
            global global_string
            global_stringList = [global_string, sector_str]
            global_string = "".join(global_stringList)
                     
        if sitesChk and sectorsChk:
            placemark_strList = [part1, part2, placemark_str, part6, global_string, part7]
            placemark_str = "".join(placemark_strList)
        elif sitesChk and not sectorsChk:
            placemark_strList = [part1, part2, placemark_str, part6, part7]
            placemark_str = "".join(placemark_strList)
        elif sectorsChk and not sitesChk:
            placemark_strList = [part1, global_string, part7]
            placemark_str = "".join(placemark_strList)
        else:
            placemark_strList = [part1, part7]
            placemark_str = "".join(placemark_strList)
        
        self.labelProgress.setText("Progress Info - Writing File to disk")
        # use xml.dom.minidom to parseString to xml default formatting
        placemark_str = xml.dom.minidom.parseString(placemark_str).toprettyxml()

        
        file = open(fullName, 'w')
        file.write(placemark_str)  # python will convert \n to os.linesep
        file.close()  # you can omit in most cases as the destructor will call it
        kml_zip = zipfile.ZipFile(name + ".kmz", 'w', compression=zipfile.ZIP_DEFLATED)
        kml_zip.write(fullName)
        kml_zip.close()
        os.remove(fullName)
        self.pbar.setFormat(str(round((listval)/(len(techs)+1)*100,0))+ "%")
        self.pbar.setValue(listval)   
        self.labelProgress.setText("Progress Info - Complete")     
    
                  
    
class confirmInput(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Some errors from Input File")
        self.setGeometry(80, 80, 300, 200)       
        self.UI()
        self.show()
        
    def UI(self):
        self.mainLayout = QVBoxLayout()
        self.messageLayout = QVBoxLayout()
        self.btnLayout = QHBoxLayout()
        self.detailMessage = QHBoxLayout()
        self.setLayout(self.mainLayout)  
        self.textInfo = QtWidgets.QLabel(messageRows) 
        self.mainLayout.addWidget(self.textInfo) 
        self.mainLayout.addLayout(self.messageLayout)
        self.mainLayout.addLayout(self.btnLayout)
        self.mainLayout.addLayout(self.detailMessage)
        self.btnLayout.addStretch()
        self.btnOk = QtWidgets.QPushButton("Ok")
        self.btnLayout.addWidget(self.btnOk)
        self.btnLayout.addStretch()
        self.infoText = QtWidgets.QTextEdit()
        self.detailMessage.addWidget(self.infoText)  
        self.infoText.setText(str(sitelistErrorDetail))
        self.btnOk.clicked.connect(self.okBtn)

    
    def okBtn(self):
        self.hide()

        
        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    MainWindow.setWindowIcon(QtGui.QIcon('tower.png'))
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

