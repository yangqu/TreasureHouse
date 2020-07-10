# -*-coding:utf-8-*-
import sys
import os
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import shutil
import uuid


origin_path = u'D:/设备原图/一级城市照片（20190107档期）/0107其他/'
lcd_postive_path = u'D:/设备原图/ImageAnnotation/ImageAnnotation/lcd-postive/'
lcd_negative_path = u'D:/设备原图/ImageAnnotation/ImageAnnotation/lcd-negative/'
bu_postive_path = 'building-postive/'
bu_negative_path = 'building-negative/'


# city_no,publish_date,route_no,filepath,重新识别结果,check_location_id,线上识别结果,p_check_location_id,oss_image_url


class window(QWidget):
    def __init__(self):
        super(window,self).__init__()
        self.title = "Image-classification"
        self.count = 0
        self.img_list = []
        self.get_all_img(origin_path, self.img_list)
        self.sum_img = len(self.img_list)
        self.initUI()

    def get_all_img(self, root, img_list):
        filelist = os.listdir(root)
        for file in filelist:
            file = os.path.join(root, file)
            if os.path.isfile(file):
                img_list.append(file)
            else:
                self.get_all_img(file, img_list)

    def initUI(self):
        wg1 = QWidget()
        wg2 = QWidget()
        wg3 = QWidget()
        wg4 = QWidget()

        # show images
        self.label1 = QLabel()
        self.label1.setPixmap(QPixmap(self.img_list[self.count]).scaled(400, 500))

        vbox1 = QHBoxLayout()
        vbox1.addWidget(self.label1)
        wg1.setLayout(vbox1)

        # show image count
        self.label = QLabel(wg2)
        self.label.setText("{}/{}File".format(self.count + 1, self.sum_img))

        bt2 = QPushButton(wg2)
        bt2.setText(u"下一组")
        bt2.clicked.connect(self.bt2_click)

        vbox2 = QHBoxLayout()
        vbox2.addWidget(self.label)
        vbox2.addWidget(bt2)
        wg2.setLayout(vbox2)

        self.label2 = QLabel(wg4)
        self.label2.setText(u"楼宇照片:")
        bt3 = QPushButton(wg4)
        bt3.setText(u"好照片")
        bt3.clicked.connect(self.bt3_click)

        bt4 = QPushButton(wg4)
        bt4.setText(u"坏照片")
        bt4.clicked.connect(self.bt4_click)

        vbox4 = QHBoxLayout()
        vbox4.addWidget(self.label2)
        vbox4.addWidget(bt3)
        vbox4.addWidget(bt4)
        wg4.setLayout(vbox4)

        self.label3 = QLabel(wg3)
        self.label3.setText(u"LCD照片:")
        bt5 = QPushButton(wg3)
        bt5.setText(u"好照片")
        bt5.clicked.connect(self.bt5_click)

        bt6 = QPushButton(wg3)
        bt6.setText(u"坏照片")
        bt6.clicked.connect(self.bt6_click)

        vbox3 = QHBoxLayout()
        vbox3.addWidget(self.label3)
        # vbox3.addWidget(bt2)
        vbox3.addWidget(bt5)
        vbox3.addWidget(bt6)
        wg3.setLayout(vbox3)

        wlayout = QVBoxLayout()
        wlayout.addWidget(wg1)
        wlayout.addWidget(wg2)
        wlayout.addWidget(wg4)
        wlayout.addWidget(wg3)

        self.setLayout(wlayout)
        self.setWindowTitle("ImageClassification")
        self.show()

    def bt2_click(self):
        self.change_img()

    def bt3_click(self):
        if self.count < self.sum_img:
            shutil.move(self.img_list[self.count], os.path.join(bu_postive_path, str(uuid.uuid1()) + '.jpg'))
            self.change_img()

    def bt4_click(self):
        if self.count < self.sum_img:
            shutil.move(self.img_list[self.count], os.path.join(bu_negative_path, str(uuid.uuid1()) + '.jpg'))
            self.change_img()

    def bt5_click(self):
        if self.count < self.sum_img:
            shutil.move(self.img_list[self.count], os.path.join(lcd_postive_path, str(uuid.uuid1()) + '.jpg'))
            self.change_img()

    def bt6_click(self):
        if self.count < self.sum_img:
            shutil.move(self.img_list[self.count], os.path.join(lcd_negative_path, str(uuid.uuid1()) + '.jpg'))
            self.change_img()

    def change_img(self):
        self.count += 1
        if self.count < self.sum_img:
            print(self.img_list[self.count])
            self.label1.setPixmap(QPixmap(self.img_list[self.count]).scaled(500, 500))
            self.label.setText(u"{}/{}File".format(self.count + 1, self.sum_img))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = window()
    sys.exit(app.exec_())
