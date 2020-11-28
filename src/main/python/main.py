# python -m fbs freeze --debug
# fbs freeze
from fbs_runtime.application_context.PyQt5 import ApplicationContext
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5 import uic


import sys
import os
import cv2

import openpyxl

import time

if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app 
    # path into variable _MEIPASS'.
    dir_path = sys._MEIPASS
else:
    dir_path = os.path.dirname(os.path.abspath(__file__))

book = openpyxl.load_workbook(dir_path+'/sample.xlsx')

sheet = book.active


class GUI(QMainWindow):
    def __init__(self,appctxt):
        self.dir_path = dir_path
            # self.dir_path =os.path.dirname(os.path.realpath(__file__))
        # self.dir_path = sys.argv[1:][0] #os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
        super(GUI,self).__init__()
        
        uic.loadUi(self.dir_path+'/main.ui', self) # Load the .ui file
        # self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.appctxt = appctxt
        self.arrangedI = 0
        
        self.cameraButton1.clicked.connect(self.runCamera1)
        # self.cameraButton2.clicked.connect(self.runCamera2)
        # self.cameraButton3.clicked.connect(self.runCamera3)
        # self.cameraButton4.clicked.connect(self.runCamera4)
        self.frontCombo.currentIndexChanged.connect(self.selectionchange)
        self.backCombo.currentIndexChanged.connect(self.selectionchange)
        self.rightCombo.currentIndexChanged.connect(self.selectionchange)
        self.leftCombo.currentIndexChanged.connect(self.selectionchange)

        self.captureImages.clicked.connect(self.talkCap)
        self.closeButtun.clicked.connect(self.close)
        
        self.selectionchange()
        
        self.show()
        self.exit_code = self.appctxt.app.exec_()      # 2. Invoke appctxt.app.exec_()

    def selectionchange(self):
        self.currentFront = int(self.frontCombo.currentText())
        self.currentBack = int(self.backCombo.currentText())
        self.currentRight = int(self.rightCombo.currentText())
        self.currentLeft = int(self.leftCombo.currentText())

    def close(self):
        try:
            self.closeAllCameras()
            self.thCamera1.ThreadCameraVideoIsRun = False
            self.thCamera2.ThreadCameraVideoIsRun = False
        except:
            pass
        QCoreApplication.exit(0)

    @pyqtSlot(QImage)
    def setImageVideo1(self, image):
        source = QPixmap.fromImage(image)
        self.cameraScreen1.setPixmap(source)

    @pyqtSlot(QImage)
    def setImageVideo2(self, image):
        source = QPixmap.fromImage(image)
        self.cameraScreen2.setPixmap(source)

    @pyqtSlot(QImage)
    def setImageVideo3(self, image):
        source = QPixmap.fromImage(image)
        self.cameraScreen3.setPixmap(source)

    @pyqtSlot(QImage)
    def setImageVideo4(self, image):
        source = QPixmap.fromImage(image)
        self.cameraScreen4.setPixmap(source)

    

    def closeAllCameras(self):
        
        try:
            self.thCamera1.ThreadCameraVideoIsRun = False
        except:
            pass
            
        try:
            self.thCamera2.ThreadCameraVideoIsRun = False
        except:
            pass
            
        try:
            self.thCamera3.ThreadCameraVideoIsRun = False
        except:
            pass
            
        try:
            self.thCamera4.ThreadCameraVideoIsRun = False
        except:
            pass

        try:
            self.thCamera1.cap.release()
        except:
            pass
        
        try:
            self.thCamera2.cap.release()
        except:
            pass
        
        try:
            self.thCamera3.cap.release()
        except:
            pass
        
        try:
            self.thCamera4.cap.release()
        except:
            pass
        
        

    def talkAllMeasureMent(self):
        self.nameOfHumanValue = str(self.nameOfHuman.text()).replace(" ", "")
        self.numValue_1 = self.num_1.text()
        self.numValue_2 = self.num_2.text()
        self.numValue_3 = self.num_3.text()
        self.numValue_4 = self.num_4.text()
        self.numValue_5 = self.num_5.text()
        self.numValue_6 = self.num_6.text()
        self.numValue_7 = self.num_7.text()
        self.numValue_8 = self.num_8.text()
        self.numValue_9 = self.num_9.text()
        self.numValue_10 = self.num_10.text()
        self.numValue_11 = self.num_11.text()
        self.numValue_12 = self.num_12.text()
        self.numValue_13 = self.num_13.text()
        self.numValue_14 = self.num_14.text()
        self.numValue_15 = self.num_15.text()
        self.numValue_16 = self.num_16.text()
        self.numValue_17 = self.num_17.text()
        self.numValue_18 = self.num_18.text()
        self.numValue_19 = self.num_19.text()
        self.numValue_20 = self.num_20.text()
        self.numValue_21 = self.num_21.text()
        self.numValue_22 = self.num_22.text()
        self.numValue_23 = self.num_23.text()
        self.numValue_24 = self.num_24.text()
        self.numValue_25 = self.num_25.text()
        self.numValue_26 = self.num_26.text()
        self.numValue_27 = self.num_27.text()
        self.numValue_28 = self.num_28.text()
        self.numValue_29 = self.num_29.text()
        self.numValue_30 = self.num_30.text()
        self.numValue_31 = self.num_31.text()
        self.numValue_32 = self.num_32.text()
        self.numValue_33 = self.num_33.text()

        rows = (
            (self.nameOfHumanValue,self.numValue_1, self.numValue_2, self.numValue_3,self.numValue_4,self.numValue_5,self.numValue_6,self.numValue_7,self.numValue_8,self.numValue_9,self.numValue_10,
             self.numValue_11, self.numValue_12, self.numValue_13,self.numValue_14,self.numValue_15,self.numValue_16,self.numValue_17,self.numValue_18,self.numValue_19,self.numValue_20,
             self.numValue_21, self.numValue_22, self.numValue_23,self.numValue_24,self.numValue_25,self.numValue_26,self.numValue_27,self.numValue_28,self.numValue_29,self.numValue_30,
             self.numValue_31,self.numValue_32,self.numValue_33),

        )
        for row in rows:
            # print(row)
            sheet.append(row)

        
        book.save(self.dir_path+'/sample.xlsx')


    def talkCap(self):
        self.talkAllMeasureMent()

        self.thCamera1.talkCapNow = True
        self.thCamera2.talkCapNow = True
        self.thCamera3.talkCapNow = True
        self.thCamera4.talkCapNow = True

    def runCamera1(self):
        self.closeAllCameras()
        
        QThread.msleep(2000)
        self.thCamera1 = ThreadCameraVideo(self,self.currentFront,False,self,'front')
        self.thCamera1.changePixmap.connect(self.setImageVideo1)
        self.thCamera1.start()
        
        self.thCamera2 = ThreadCameraVideo(self,self.currentBack,False,self,'back')
        self.thCamera2.changePixmap.connect(self.setImageVideo2)
        self.thCamera2.start()
        

        self.thCamera3 = ThreadCameraVideo(self,self.currentLeft,False,self,'left')
        self.thCamera3.changePixmap.connect(self.setImageVideo3)
        self.thCamera3.start()

        
        self.thCamera4 = ThreadCameraVideo(self,self.currentRight,False,self,'right')
        self.thCamera4.changePixmap.connect(self.setImageVideo4)
        self.thCamera4.start()



    
        




class ThreadCameraVideo(QThread):
    changePixmap = pyqtSignal(QImage)
    ThreadCameraVideoIsRun = True
    

    def __init__(self,window,cam,talkCapNow,mainGUI,arrangedI):
        super(ThreadCameraVideo,self).__init__(window)
        self.window = mainGUI
        self.cam = cam
        self.arrangedI = arrangedI
        self.talkCapNow = talkCapNow
        print(self.talkCapNow)

    
    

    
    def run(self):
        self.cap = cv2.VideoCapture(self.cam)
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 800)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT,600)
        while self.ThreadCameraVideoIsRun:
            
            ret, sample_frame = self.cap.read()
            if ret:
                # frame = cv2.flip(sample_frame, 2)
                dim = (210,210)
                sample_frame = cv2.rotate(sample_frame, cv2.ROTATE_90_CLOCKWISE) 
                frame = cv2.resize(sample_frame, dim)

                rgbImage = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                h, w, ch = rgbImage.shape
                bytesPerLine = ch * w
                convertToQtFormat = QImage(rgbImage.data, w, h, bytesPerLine, QImage.Format_RGB888)
                self.changePixmap.emit(convertToQtFormat)
                if self.talkCapNow:
                    if not os.path.exists(self.window.dir_path+"/"+self.window.nameOfHumanValue):
                        os.makedirs(self.window.dir_path+"/"+self.window.nameOfHumanValue)
                    
                    cv2.imwrite(self.window.dir_path+"/"+self.window.nameOfHumanValue+"/"+self.arrangedI+".jpg", sample_frame)

                    self.talkCapNow = False

         
        self.cap.release()
        
if __name__ == '__main__':
    appctxt = ApplicationContext()       # 1. Instantiate ApplicationContext
    mainApp = GUI(appctxt)
    print("OK")
    sys.exit(mainApp.exit_code)
    