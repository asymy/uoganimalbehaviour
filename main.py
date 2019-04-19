# ! /usr/bin/python
#
#
# Software for analysing the behaviour of animals, specifically 
# licking and biting from pre-recorded videos.
#
# Alison Symon for the spinal cord group at the University of Glasgow
# 2019
#
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 2 of the License, or
# (at your option) any later version.
#
#


import sys
import os.path
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtWidgets import QMainWindow, QWidget, QFrame, QSlider, QHBoxLayout, QPushButton, \
    QVBoxLayout, QAction, QFileDialog, QApplication, QLineEdit, QLabel
import vlc

class Player(QMainWindow):
    """A simple Media Player using VLC and Qt
    """
    def __init__(self, master=None):
        QMainWindow.__init__(self, master)
        self.setWindowTitle("Spinal Group Animal Behaviour Analyser")

        # creating a basic vlc instance
        self.instance = vlc.Instance()
        # creating an empty vlc media player
        self.mediaplayer = self.instance.media_player_new()

        self.createUI()
        self.isPaused = False
        self.isLicking = False
        self.isBiting = False

        self.startLickTime = []
        self.stopLickTime = []
        self.startBiteTime = []
        self.stopBiteTime = []

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_L:
            self.LickStartStop()
        elif e.key() == Qt.Key_B:
            self.BiteStartStop()
        elif e.key() == Qt.Key_Equal:
            self.resetSpeed()
        elif e.key() == Qt.Key_BracketLeft:
            self.decreaseSpeed()
        elif e.key() == Qt.Key_BracketRight:
            self.increaseSpeed()

    def createUI(self):
        """Set up the user interface, signals & slots
        """
        self.widget = QWidget(self)
        self.setCentralWidget(self.widget)

        # In this widget, the video will be drawn
        if sys.platform == "darwin": # for MacOS
            from PyQt5.QtWidgets import QMacCocoaViewContainer	
            self.videoframe = QMacCocoaViewContainer(0)
        else:
            self.videoframe = QFrame()
        self.palette = self.videoframe.palette()
        self.palette.setColor (QPalette.Window,
                               QColor(0,0,0))
        self.videoframe.setPalette(self.palette)
        self.videoframe.setAutoFillBackground(True)

        self.positionslider = QSlider(Qt.Horizontal, self)
        self.positionslider.setToolTip("Position")
        self.positionslider.setMaximum(1000)
        self.positionslider.sliderMoved.connect(self.setPosition)

        self.videocontolbox = QHBoxLayout()
        self.playbutton = QPushButton("Play")
        self.videocontolbox.addWidget(self.playbutton)
        self.playbutton.clicked.connect(self.PlayPause)

        self.stopbutton = QPushButton("Stop")
        self.videocontolbox.addWidget(self.stopbutton)
        self.stopbutton.clicked.connect(self.Stop)

        self.videocontolbox.addStretch(1)

        self.speedslider = QSlider(Qt.Horizontal, self)
        self.speedslider.setMaximum(200)
        self.speedslider.setMinimum(50)
        self.speedslider.setValue(self.mediaplayer.get_rate()*100)
        self.speedslider.setToolTip("Speed")
        self.videocontolbox.addWidget(self.speedslider)
        self.speedslider.valueChanged.connect(self.setSpeed)

        self.volumeslider = QSlider(Qt.Horizontal, self)
        self.volumeslider.setMaximum(100)
        self.volumeslider.setValue(self.mediaplayer.audio_get_volume())
        self.volumeslider.setToolTip("Volume")
        self.videocontolbox.addWidget(self.volumeslider)
        self.volumeslider.valueChanged.connect(self.setVolume)

        self.LickBehaviour = QHBoxLayout()
        self.licktoggle = QPushButton("Start Lick")
        self.LickBehaviour.addWidget(self.licktoggle)
        self.licktoggle.clicked.connect(self.LickStartStop)
        self.LickBehaviour.addStretch(1)

        self.BiteBehaviour = QHBoxLayout()
        self.bitetoggle = QPushButton("Start Bite")
        self.BiteBehaviour.addWidget(self.bitetoggle)
        self.bitetoggle.clicked.connect(self.BiteStartStop)
        self.BiteBehaviour.addStretch(1)

        self.InfoSave = QHBoxLayout()
        self.InfoSave.addStretch(20)
        self.IDLabel = QLabel(self)
        self.IDLabel.setText('Animal ID: ')
        self.InfoSave.addWidget(self.IDLabel)
        self.animalID = QLineEdit(self)
        self.animalID.setText("")
        self.InfoSave.addWidget(self.animalID)
        self.InfoSave.addStretch(1)

        self.DateLabel = QLabel(self)
        self.DateLabel.setText('Date: ')
        self.InfoSave.addWidget(self.DateLabel)
        self.Date = QLineEdit(self)
        self.Date.setText("")
        self.InfoSave.addWidget(self.Date)

        self.InfoSave.addStretch(1)
        self.savebutton = QPushButton("Save")
        self.savebutton.clicked.connect(self.SaveData)
        self.InfoSave.addWidget(self.savebutton)

        self.vboxlayout = QVBoxLayout()
        self.vboxlayout.addWidget(self.videoframe)
        self.vboxlayout.addWidget(self.positionslider)
        self.vboxlayout.addLayout(self.videocontolbox)
        self.vboxlayout.addLayout(self.LickBehaviour)
        self.vboxlayout.addLayout(self.BiteBehaviour)
        self.vboxlayout.addLayout(self.InfoSave)
        self.widget.setLayout(self.vboxlayout)

        open = QAction("&Open", self)
        open.triggered.connect(self.OpenFile)
        exit = QAction("&Exit", self)
        exit.triggered.connect(sys.exit)
        menubar = self.menuBar()
        filemenu = menubar.addMenu("&File")
        filemenu.addAction(open)
        filemenu.addSeparator()
        filemenu.addAction(exit)

        self.timer = QTimer(self)
        self.timer.setInterval(200)
        self.timer.timeout.connect(self.updateUI)

    def PlayPause(self):
        """Toggle play/pause status
        """
        if self.mediaplayer.is_playing():
            self.mediaplayer.pause()
            self.playbutton.setText("Play")
            self.isPaused = True
        else:
            if self.mediaplayer.play() == -1:
                self.OpenFile()
                return
            self.mediaplayer.play()
            self.playbutton.setText("Pause")
            self.timer.start()
            self.isPaused = False

    def Stop(self):
        """Stop player
        """
        self.mediaplayer.stop()
        self.playbutton.setText("Play")

    def OpenFile(self, filename=None):
        """Open a media file in a MediaPlayer
        """
        filename = QFileDialog.getOpenFileName(self, "Open File", os.path.expanduser('~'))[0]
        if not filename:
            return

        # create the media
        if sys.version < '3':
            filename = unicode(filename)
        self.media = self.instance.media_new(filename)
        # put the media in the media player
        self.mediaplayer.set_media(self.media)

        # parse the metadata of the file
        self.media.parse()
        # set the title of the track as window title
        # self.setWindowTitle(self.media.get_meta(0))

        # the media player has to be 'connected' to the QFrame
        # (otherwise a video would be displayed in it's own window)
        # this is platform specific!
        # you have to give the id of the QFrame (or similar object) to
        # vlc, different platforms have different functions for this
        if sys.platform.startswith('linux'): # for Linux using the X Server
            self.mediaplayer.set_xwindow(self.videoframe.winId())
        elif sys.platform == "win32": # for Windows
            self.mediaplayer.set_hwnd(self.videoframe.winId())
        elif sys.platform == "darwin": # for MacOS
            self.mediaplayer.set_nsobject(int(self.videoframe.winId()))
        self.PlayPause()


    def setVolume(self, Volume):
        """Set the volume
        """
        self.mediaplayer.audio_set_volume(Volume)
        self.volumeslider.setValue(self.mediaplayer.audio_get_volume())

    def increaseSpeed(self):
        "increase the speed"
        current = self.mediaplayer.get_rate()*100
        new = current + 25
        if new > 200:
            new = 200
        self.setSpeed(new)

    def decreaseSpeed(self):
        "decrease the speed"
        current = self.mediaplayer.get_rate()*100
        new = current - 25
        if new < 50:
            new = 50
        self.setSpeed(new)

    def resetSpeed(self):
        self.setSpeed(100)


    def setSpeed(self, Speed):
        """Set the playback speed
        """
        self.mediaplayer.set_rate((Speed/100))
        self.speedslider.setValue(self.mediaplayer.get_rate()*100)

    def setPosition(self, position):
        """Set the position
        """
        # setting the position to where the slider was dragged
        self.mediaplayer.set_position(position / 1000.0)
        # the vlc MediaPlayer needs a float value between 0 and 1, Qt
        # uses integer variables, so you need a factor; the higher the
        # factor, the more precise are the results
        # (1000 should be enough)

    def LickStartStop(self):
        """Toggle state of licking behaviour
        """
        if self.mediaplayer.is_playing():
            if self.isLicking == False:
                self.isLicking = True
                self.licktoggle.setText("Stop Lick")
                self.startLickTime.append(self.mediaplayer.get_position()*1000)
            else:
                self.isLicking = False
                self.stopLickTime.append(self.mediaplayer.get_position()*1000)
                self.licktoggle.setText("Start Lick")

    def BiteStartStop(self):
        """Toggle state of licking behaviour
        """
        if self.mediaplayer.is_playing():
            if self.isBiting == False:
                self.isBiting = True
                self.bitetoggle.setText("Stop Bite")
                self.startBiteTime.append(self.mediaplayer.get_position()*1000)
            else:
                self.isBiting = False
                self.stopBiteTime.append(self.mediaplayer.get_position()*1000)
                self.bitetoggle.setText("Start Bite")

    def SaveData(self):
        dest_filename = self.animalID.text() + '_' + self.Date.text() + '.xlsx'
        savefilename, _ = QFileDialog.getSaveFileName(self, 'Save File', os.path.expanduser('~') + '\\' + dest_filename, '*.xlsx')
        if savefilename:
            print(savefilename)
            wb = Workbook()
            ws1 = wb.active
            ws1.title = "summary"
            ws2 = wb.create_sheet(title="LickData")
            ws3 = wb.create_sheet(title="BiteData")
            if len(self.startLickTime) > 1:
                _ = ws2.cell(column=1, row=1, value="Occurence")
                _ = ws2.cell(column=2, row=1, value="Start Time")
                _ = ws2.cell(column=3, row=1, value="Stop Time")
                _ = ws2.cell(column=4, row=1, value="Duration")

                for row in range(len(self.startLickTime)):
                    _ = ws2.cell(column=1, row=row+2, value=row)
                    _ = ws2.cell(column=2, row=row+2, value=self.startLickTime[row])
                    _ = ws2.cell(column=3, row=row+2, value=self.stopLickTime[row])
                    _ = ws2.cell(column=4, row=row+2, value='=C{0}-B{0}'.format(row+2))
                _ = ws2.cell(column=4, row=row+3, value='=SUM(D2:D{0})'.format(row+2))
            if len(self.startBiteTime) > 1:
                _ = ws3.cell(column=1, row=1, value="Occurence")
                _ = ws3.cell(column=2, row=1, value="Start Time")
                _ = ws3.cell(column=3, row=1, value="Stop Time")
                _ = ws3.cell(column=4, row=1, value="Duration")

                for row in range(len(self.startBiteTime)):
                    _ = ws3.cell(column=1, row=row+2, value=row)
                    _ = ws3.cell(column=2, row=row+2, value=self.startBiteTime[row])
                    _ = ws3.cell(column=3, row=row+2, value=self.stopBiteTime[row])
                    _ = ws3.cell(column=4, row=row+2, value='=C{0}-B{0}'.format(row+2))
                _ = ws3.cell(column=4, row=row+3, value='=SUM(D2:D{0})'.format(row+2))

            wb.save(savefilename)
        else:
            return

    def updateUI(self):
        """updates the user interface"""
        # setting the slider to the desired position
        
        self.positionslider.setValue(self.mediaplayer.get_position() * 1000)

        if not self.mediaplayer.is_playing():
            # no need to call this function if nothing is played
            self.timer.stop()
            if not self.isPaused:
                # after the video finished, the play button stills shows
                # "Pause", not the desired behavior of a media player
                # this will fix it
                self.Stop()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    player = Player()
    player.show()
    player.resize(1280, 960)
    if sys.argv[1:]:
        player.OpenFile(sys.argv[1])
    sys.exit(app.exec_())