import os
import win32clipboard as cb
from sys import argv, exit
from datetime import datetime
from PIL import Image, ImageGrab
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QAction, QLabel, QSizePolicy
from PyQt5.QtGui import QPainter, QBrush, QColor
from PyQt5.QtCore import QPoint, Qt, QRect
from io import BytesIO

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setGeometry(200, 200, 350, 200)
        self.setWindowTitle("ScreenShoter")
        self.setStyleSheet("""
        QMainWindow {
            background: #2D3142;
            color: #BFC0C0;
            }
        QTabBar::close-button {
            background: #2D3142;
            color: #BFC0C0;
            }
        QPushButton {
            background: #2D3142;
            color: #BFC0C0;
            border: 2px solid #BFC0C0;
            font-size: 18px;
            }
        QLabel {
            background: #2D3142;
            color: #BFC0C0;
            font-size: 14px;
            }
        QMenuBar {
            background: #2D3142;
            color: #BFC0C0;
            }
        QMenuBar::item:selected {
            background: #2D3100;
            }
        QMenu {
            background: #2D3142;
            color: #BFC0C0;
            }
        QMenu::item {
            background-color: transparent;
            }
        """)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint)
        
        # screenshot button
        self.button = QPushButton("Shot!", self)
        self.button.setGeometry(5, 30, self.geometry().width() - 10, 50)
        self.button.clicked.connect(self.screenshot)
        self.button.setFlat(True)

        self.label = QLabel('Instrukcja: \nPo kliknięciu na przycisk "Shot!":\
             \n- fragment ekranu (zaznaczenie lewym przyciskiem)\
             \n- zrzut całego ekranu (prawy przycisk)\
             \n- anulowanie (Escape)', self)
        self.label.setGeometry(5, 100, self.geometry().width(), 90)
        self.label.setWordWrap(True)
        # cursor
        QApplication.setOverrideCursor(Qt.ArrowCursor)

        # menubar
        openExplorerAction = QAction("&Otwórz lokalizację zrzutów ekranu", self)
        openExplorerAction.triggered.connect(self.openExplorer)

        self.mainMenu = self.menuBar()
        self.fileMenu = self.mainMenu.addMenu('&Plik')
        self.fileMenu.addAction(openExplorerAction)

        # show all the widgets
        self.show()

    def screenshot(self):
        self.hide()
        self.window_02 = Window(self)
        self.window_02.show()
        self.window_02.raise_()
        self.window_02.activateWindow()
    
    def openExplorer(self):
        os.startfile(os.getcwd() + "/screens/", 'open')


class Window(QMainWindow):
    def __init__(self, parent):
        super(Window, self).__init__(parent)

        self.begin = QPoint()
        self.end = QPoint()
        self.pressed = False
        self.pressedRight = False

        # hide title bar
        self.setWindowFlag(Qt.FramelessWindowHint)

        # this will hide the app from task bar 
        self.setWindowFlag(Qt.Tool)

        self.setAttribute(Qt.WA_QuitOnClose, True)
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        
        # transparency
        self.setWindowOpacity(0.4)
        
        # move to corner
        self.move(0, 0)

        # cursor
        QApplication.setOverrideCursor(Qt.CrossCursor)

        # show all the widgets
        self.show()

        # set size
        self.taskBarSize = 32
        self.allScreens = QApplication.desktop().geometry()
        widgetSize = self.allScreens.adjusted(0, 0, 0, self.taskBarSize)
        self.setGeometry(widgetSize)

    def closeAndReturn(self):
        self.close()
        self.parent().show()
        # cursor
        QApplication.setOverrideCursor(Qt.ArrowCursor)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            # self.close()
            self.closeAndReturn()
            print("Escape pressed")
        else:
            print("Another key pressed")
    
    def paintEvent(self, event):
        qp = QPainter(self)
        br = QBrush(QColor(255,255,255,255))
        qp.setBrush(br)
        qp.drawRect(QRect(self.begin, self.end))

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.begin = event.pos()
            self.end = event.pos()
            self.pressed = True
            self.update()
    
    def mouseMoveEvent(self, event):
        if self.pressed == True:
            self.end = event.pos()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.end = event.pos()
            self.pressed = False
            self.update()
            self.setWindowOpacity(0)
            self.main()
        elif event.button() == Qt.RightButton:
            self.pressedRight = True
            self.x1, self.y1 = 0, 0
            self.x2, self.y2 = self.allScreens.width(), self.allScreens.height()
            self.setWindowOpacity(0)
            self.main()
            
    def send_to_clipboard(self, clip_type, data):
        cb.OpenClipboard()
        cb.EmptyClipboard()
        cb.SetClipboardData(clip_type, data)
        cb.CloseClipboard()
        
    def main(self):
        if not self.pressedRight:
            self.x1, self.y1 = self.begin.x(), self.begin.y()
            self.x2, self.y2 = self.end.x(), self.end.y()
        if self.x1 > self.x2:
            self.x1, self.x2 = self.x2, self.x1
        if self.y1 > self.y2:
            self.y1, self.y2 = self.y2, self.y1

        area = (self.x1, self.y1, self.x2, self.y2)
        print(area)

        now = datetime.now()
        # dd-mm-YY H-M-S
        dt_string = now.strftime("%d-%m-%Y %H-%M-%S")
        try:
            filepath = "screens/" + dt_string + '.png'
            self.screenshot = ImageGrab.grab(all_screens = True)
            self.screenshot = self.screenshot.crop(area)
            self.screenshot.save(filepath, 'PNG')
            self.screenshot.show()

            output = BytesIO()
            image = self.screenshot.convert("RGB").save(output, "BMP")
            data = output.getvalue()[14:]
            output.close()

            self.send_to_clipboard(cb.CF_DIB, data)
        except SystemError as identifier:
            print("Cannot save image - wrong numbers")
        self.closeAndReturn()
        # self.close()
        # self.deleteLater()

if __name__ == "__main__":
    App = QApplication(argv)
    window = MainWindow()
    exit(App.exec())
