# import Important modules
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic,QtGui
import sys
import qdarkstyle

##################### create pdf ########################
from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

import os


def create_pdf(full_name,cne,cin,filiere,numero_inscrit):
    # file name 
    outfile_name = full_name.split(' ')
    outfile_name = outfile_name[0] + '_' + outfile_name[1]

    # draw new pdf
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFillColorRGB(0, 0, 0)
    can.setFont("Times-Roman", 12)
    can.drawString(240, 592, full_name)
    can.drawString(250, 578, cne)
    can.drawString(220, 564, cin)
    can.drawString(305, 550, filiere)
    can.drawString(190, 536, numero_inscrit)
    can.save()

    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))
    my_file = os.path.join(THIS_FOLDER, "pdf_in.pdf")
    existing_pdf = PdfFileReader(open(my_file, "rb"))
    output = PdfFileWriter()
    # merge pdf
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    cwd = os.getcwd()
    dir = os.path.join(THIS_FOLDER,"InscriptionFiles")
    if not os.path.exists(dir):
        os.mkdir(dir)
        print("Create Directory successfully")
    else:
        print("Directory already exists")
    # output pdf to the inscrption folder
    path = THIS_FOLDER + "/InscriptionFiles/" + outfile_name + ".pdf"
    outputStream = open(path, "wb")
    output.write(outputStream)
    outputStream.close()


filiere = ['GE', 'RT', 'GP', 'AGB', 'GIM', 'GI', 'ID', 'GTE', 'TM', 'GRH', 'TGC', 'GLT', 'SE', 'LOG', 'MEEB', 'PI', 'QA', 'TIC', 'TEREE', 'ABD', 'ASR', 'EEI', 'EII', 'GL', 'LOG IND', 'LOG INT', 'PIM', 'PIP', 'MI', 'TCC', 'TVFC', 'STID', 'VPM', 'LP-MI', 'MPII', 'PP', 'GMP']


FORM_CLASS,_ = uic.loadUiType(os.path.join(os.path.dirname(__file__),'main.ui'))
#_Welcom_page_########################################################################
class MainApp(QWidget, FORM_CLASS):
    def __init__(self, parent=None):
        super(MainApp,self).__init__(parent)
        QWidget.__init__(self)
        self.setupUi(self)
        self.setupUi_s()
        self.triggedButtons()

    def setupUi_s(self):
        QApplication.processEvents()
        self.setWindowTitle('   APP for creating Pdf inscription file')
        self.setGeometry(450, 250, 487, 472)
        self.setFixedSize(487,472)
        self.setWindowIcon(QtGui.QIcon('icon.png'))
        self.comboBox_filiere.addItems(filiere)
        QApplication.processEvents()
    
    def triggedButtons(self):
        self.pushButton_submit.clicked.connect(self.submit)

    def submit(self):
        self.collect_info()
        if self.checkdata() == False:
            self.showmessage_error()
        else:
            self.checksumbit()
        
    def showmessage_error(self):
        msg = QMessageBox()
        # msg.setIcon(QMessageBox.warning)
        msg.setWindowTitle("Error Submiting")
        msg.setIcon(QMessageBox.Warning)
        msg.setText("Failed Submit")
        msg.setInformativeText("Necessary data not completed")
        msg.setDetailedText('\n'.join(self.message_data_error))
        ms = msg.exec_()

    def checkdata(self):
        check = True
        self.message_data_error = []
        if len(self.full_name) == 0 :
            self.message_data_error.append(f'Please enter full name')
            check = False
        if (' ' in self.cne) or (len(self.cne) == 0):
            self.message_data_error.append("invalide or empty CNE")
            check = False
        if (' ' in self.cin) or (len(self.cin) == 0):
            self.message_data_error.append("invalide or empty CIN")
            check = False
        if self.filiere == '...':
            self.message_data_error.append("Invalid filiere")
            check = False
        if self.num_inscrit == 0:
            self.message_data_error.append("Invalid inscrit number")
            check = False
        print(self.message_data_error)
        return check

    def checksumbit(self):
        msg = QMessageBox()
        # msg.setIcon(QMessageBox.warning)
        msg.setWindowTitle("Confirm Submiting")
        msg.setIcon(QMessageBox.Question)
        msg.setText("Are you Sure ?")
        msg.setInformativeText(f"You'll create a new inscription file for {self.full_name}")
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.buttonClicked.connect(self.checksumbitbuttons)
        ms = msg.exec_()

    def checksumbitbuttons(self,i):
        print("Button pressed is:",i.text())
        if i.text() == '&Yes':
            try:
                create_pdf(self.full_name,self.cne,self.cin,self.filiere,self.num_inscrit)
            except Exception as error:
                print("create failed")
                QMessageBox.warning(self,"Creat Failed","create pdf failed, Try again later")
                print(error)
            else:
                print("create successfly")
                QApplication.processEvents()
                QMessageBox.information(self,"Creat Succefly","create pdf succefly")
                QApplication.processEvents()
                self.openpdf(self.full_name)
                self.reload()
                QApplication.processEvents()
        else:
            pass

    def reload(self):
        self.lineEdit_full_name.setText("")
        self.lineEdit_cne.setText("")
        self.lineEdit_cin.setText("")
        self.comboBox_filiere.setCurrentText("...")
        self.lineEdit_inscrit.setValue(0)

    def collect_info(self):
        self.full_name = self.lineEdit_full_name.text().strip().upper()
        self.cne = self.lineEdit_cne.text().strip().upper()
        self.cin = self.lineEdit_cin.text().strip().upper()
        self.filiere = self.comboBox_filiere.currentText()
        self.num_inscrit = str(self.lineEdit_inscrit.value())


    def openpdf(self,full_name):
        outfile_name = full_name.split(' ')
        outfile_name = outfile_name[0] + '_' + outfile_name[1]
        THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(THIS_FOLDER, f"InscriptionFiles\{outfile_name}.pdf")
        # path = "Inscription Files\\" + outfile_name + ".pdf"
        try:
            os.system(path)
            print("opening the pdf...")
        except Exception as er:
            print("file not open because of : ")
            print(er)



def main():
    app = QApplication(sys.argv)
    widget = MainApp()
    QApplication.processEvents()
    ###########################################################   
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    widget.show()
    try:
        sys.exit(app.exec_())
    except:
        pass

if __name__ == "__main__":
    main()