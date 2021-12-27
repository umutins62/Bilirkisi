import sys
from PyQt5.QtWidgets import QApplication, QWidget, QGroupBox, QHBoxLayout, QGridLayout, QComboBox, QLabel, QPushButton, \
    QLineEdit, QCheckBox, QSpinBox
from PyQt5.QtCore import QSettings


class ayarlar(QWidget):
    def __init__(self):
        super().__init__()
        self.SetUI()
        self.temayukle()
        self.getsetting()


        # print(self.settings.fileName())
        # textbox
        self.email_value=self.setting_value.value('text box')
        self.sifre_value=self.setting_value1.value('text box1')
        self.gun_value=self.setting_value5.value('text box2')
        self.email.setText(self.email_value)
        self.sifre.setText(self.sifre_value)
        self.gun.setText(self.gun_value)

        # chekbox

        self.sifregoster.stateChanged.connect(self.sifregizle)
        self.sifregoster.stateChanged.connect(self.closeEvent)
        self.otyedekle.stateChanged.connect(self.gizle)
        self.otyedekle.stateChanged.connect(self.closeEvent)
        self.kayıtgoster.stateChanged.connect(self.closeEvent)

        self.gizle()

        # combobox
        self.theme_value=self.setting_value5.value('theme')
        self.theme.setCurrentText(self.theme_value)
        # form konumu
        self.move(self.setting_value9.value('window position'))

        self.sifregizle()


    def temayukle(self):
        if self.theme.currentText()=="Consol Style":
            self.setStyleSheet(open('consolstyle.qss','r').read())
        elif self.theme.currentText()=="ElegantDark":
            self.setStyleSheet(open('elegantdark.qss','r').read())
        elif self.theme.currentText()=="Ubuntu":
            self.setStyleSheet(open('ubuntu.qss','r').read())
        elif self.theme.currentText()=="MacOS":
            self.setStyleSheet(open('macos.qss','r').read())
        else:
            pass

    def getsetting(self):
        # textbox
        self.setting_value=QSettings('Set_App','vraiables')
        self.setting_value1=QSettings('Set_App','vraiables')
        self.setting_value5=QSettings('Set_App','vraiables')
        # form ayarla
        self.setting_value9=QSettings('Set_App','App1')
        # chekbox
        self.setting_value2=QSettings('A','vraiables')
        self.setting_value4=QSettings('B','vraiables')
        self.setting_value6=QSettings('C','vraiables')
        self.setting_value7=QSettings('D','vraiables')

        self.sifregoster.setChecked(self.setting_value2.value("A", False, bool))
        self.otyedekle.setChecked(self.setting_value4.value("B", False, bool))
        self.kayıtgoster.setChecked(self.setting_value6.value("C", False, bool))
        self.toplamgoster.setChecked(self.setting_value7.value("D", False, bool))
        # combobox
        self.setting_value8=QSettings('Set_App','vraiables')

    def closeEvent(self, event):
        # textbox
        self.setting_value.setValue('text box',self.email.text())
        self.setting_value1.setValue('text box1',self.sifre.text())
        self.setting_value5.setValue('text box2',self.gun.text())
        # chekbox
        self.setting_value2.setValue("A", self.sifregoster.isChecked())
        self.setting_value4.setValue("B", self.otyedekle.isChecked())
        self.setting_value6.setValue("C", self.kayıtgoster.isChecked())
        self.setting_value7.setValue("D", self.toplamgoster.isChecked())
        # combobox
        self.setting_value8.setValue('theme',self.theme.currentText())
        # form ayarla
        self.setting_value9.setValue('window position',self.pos())

    def gizle(self):
        if self.otyedekle.isChecked():
            self.gun.setVisible(True)
        else:
            self.gun.setVisible(False)

    def sifregizle(self):

        if self.sifregoster.isChecked():
            self.sifre.setEchoMode(QLineEdit.Normal)
        else:
            self.sifre.setEchoMode(QLineEdit.Password)

    def SetUI(self):
        self.setWindowTitle("Ayarlar")
        hbox=QHBoxLayout()
        grub1=QGroupBox("Veritabanı")
        grub2=QGroupBox("Uygulama")
        self.dataadet=QLabel("Kayıt Sayısını Göster")
        self.datatoplam=QLabel("Toplam Net Tutarı  Göster")

        self.theme=QComboBox()
        self.theme.addItem("Consol Style")
        self.theme.addItem("ElegantDark")
        self.theme.addItem("Ubuntu")
        self.theme.addItem("MacOS")
        self.theme.currentTextChanged.connect(self.temayukle)
        self.tema=QLabel("Tema")


        #veritabanı ayarları
        self.email=QLineEdit()
        self.email.setPlaceholderText("Email Adresini Gir")
        self.sifre=QLineEdit()
        self.sifre.setPlaceholderText("Şifreyi Gir")
        self.sil=QPushButton("Veritabanını Sil")
        self.yedekle=QPushButton("Veritabanını Yedekle")

        self.sifregoster=QCheckBox("Şifreyi Göster")
        self.kayıtgoster=QCheckBox("")
        self.toplamgoster=QCheckBox("")

        self.otyedekle=QCheckBox("Otomatik Yedekle(Gün)")
        self.gun=QLineEdit()
        self.gun.setPlaceholderText("Gün Gir")


        grid1 = QGridLayout()
        grid2 = QGridLayout()
        grub1.setLayout(grid1)
        grub2.setLayout(grid2)

        grid1.addWidget(self.email,1,0)
        grid1.addWidget(self.sifre,2,0)
        grid1.addWidget(self.sifregoster,2,1)
        grid1.addWidget(self.otyedekle,3,0)
        grid1.addWidget(self.gun,3,1)
        grid1.addWidget(self.sil,4,0)
        grid1.addWidget(self.yedekle,4,1)

        grid2.addWidget(self.dataadet,1,0)
        grid2.addWidget(self.kayıtgoster,1,1)
        grid2.addWidget(self.datatoplam,2,0)
        grid2.addWidget(self.toplamgoster,2,1)
        grid2.addWidget(self.tema,3,0)
        grid2.addWidget(self.theme,3,1)


        hbox.addWidget(grub1)
        hbox.addWidget(grub2)

        self.setLayout(hbox)
        self.show()



