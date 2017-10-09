import sys,time
from ustum_complaints_ui import *
from picture_qrc import  *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from serial.tools.list_ports import *
import datetime
import sqlite3
import uuid
import xlwt
import shutil



class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MyApp, self).__init__()
        QtWidgets.QMainWindow.__init__(self)
        self.setupUi(self)
        Ui_MainWindow.__init__(self)
        # logo
        self.setWindowIcon(QIcon(":picture/img/110.png"))
        # 默认时间戳
        self.time_stamp = datetime.datetime.now().strftime('%Y-%m-%d')
        # 连接到把数据库
        try:
            self.table_name = 'zhjk_support'
            self.data_dir = u"Y:/产品部/智慧健康产品线/技术支持/客诉资料/chipsea_support.db"
            self.connect = sqlite3.connect(self.data_dir)
            shutil.copy(self.data_dir,u'./chipsea_support.db')

        except Exception as e:
            print("无法访问数据库,请检查局域网配置")
            QMessageBox.information(self,"提示","1.访问共享盘失败,请配置共享[Y:/产品部/智慧健康产品线/技术支持/客诉资料/]!\n 2.自动连接到本地数据库,所有操作将不会同步到共享!",QMessageBox.Yes)
            self.connect = sqlite3.connect(u"./chipsea_support.db")

        self.cursor = self.connect.cursor()
        # function initial
        self.init_add_table_display()
        self.init_watch_table_display()
        self.init_edit_table_display()
        #初始化显示大小
        self.init_default_display()
        self.creat_watch_tabel_view_mecu()

        #初始化信号槽
        self.pushButton_add_database.clicked.connect(self.pushButton_add_database_clicked)
        self.watch_table_view.clicked.connect(self.on_click_watch_table_view)
        self.watch_table_view.customContextMenuRequested.connect(self.watch_table_right_clicked)
        self.watch_table_view.pressed.connect(self.watch_table_pressed)
        self.pushButton_out_excle.clicked.connect(self.pushButton_out_excle_hander)
        self.pushButton_save_watch_table.clicked.connect(self.on_click_save_watch_table_view)        #combobox
        self.comboBox_person.currentIndexChanged.connect(self.refresh_view_table)
        self.comboBox_type.currentIndexChanged.connect(self.refresh_view_table)
        self.comboBox_state.currentIndexChanged.connect(self.refresh_view_table)

    def init_default_display(self):
        # size
        self.__desktop = QApplication.desktop()
        qRect = self.__desktop.screenGeometry()  # 设备屏幕尺寸
        self.resize(qRect.width() * 2 / 5, qRect.height() * 80 / 100)
        self.move(qRect.width() / 3, qRect.height() / 30)

    def init_add_table_display(self):
        pass
        self.labels_names = [
            {'FAE':"姜恒"},
            {'开始时间':"2017-09-23"},
            {'方案名称':"蓝牙厨房秤"},
            {'问题类型':"技术支持,项目,客诉,FAQ"},
            {'问题状态':"进行中"},
            {'芯片型号':"CST34M97,CSU14PD87"},
            {'产品名称':"SF-400"},
            {'当前进度':"立项中"},
            {'问题描述':"文档看不懂,找不到北"},
            {'根本原因':"软件中断处理BUG"},
            { '解决措施':"修改软件,更改文档"},
            {'代理商':"西城微科"},
            {'终端客户':"普天"},
            {'开发工具及版本':"IDE-V4.0.7"},
            {'仿真工具及版本':"ICE-LiteV0.3"},
            {'测试工具及版本':"测试架-V1.0.2"},
            {'APP及版本':"OKOK-V2.3.1.0"},
            {'客户工程师':"李工"},
            {'客户联系电话':"1388888888,0755-289777777"},
            {'标签':"BUG,芯片问题,应用问题"},
            {'业务员':"王中王"},
            {'客诉单号':"CS-AR-2017-01-05"},
            {'预计完成时间':"2017-09-39"},
            {'实际完成时间':"2018-06-93"},
            {'重要等级':"一般,重要,紧急,立即"},
            {'占用时间':"0.5小时"},
            {'存档':"Y:/zhjk/ccc"},
            {'备注':"放弃治疗"},
            {'UUID':"自动生成,无需修改"}
        ]
        # 重新生成对应字段
        self.lables = [" "]*len(self.labels_names)

        i=0
        for dict in self.labels_names:
            h = str(dict.keys());
            h = h.replace(u"dict_keys(['", "");
            h = h.replace(u"'])", "")
            self.lables[i]= h
            i = i+1
        self.lableString= str(self.lables).replace("[","")
        self.lableString= self.lableString.replace("]","")
        self.lableString= self.lableString.replace("\'","")
        # initiate
        self.add_modle = QStandardItemModel(len(self.labels_names), 1)
        self.add_table_view.setModel(self.add_modle)
        self.add_table_view.horizontalHeader().setStretchLastSection(True)
        self.add_table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 初始化表格显示
        self.add_modle.setHorizontalHeaderLabels(['相关信息/详细填写'])
        #初始化显示内容
        i = 0
        for dic in self.labels_names:
            h = str(dic.keys());h = h.replace(u"dict_keys(['","");h = h.replace(u"'])","")
            v = str(dic.values());v = v.replace(u"dict_values(['","");v = v.replace(u"'])","")
            self.add_modle.setVerticalHeaderItem(i,QStandardItem(h))
            self.add_modle.setItem(i,QStandardItem(v))
            i=i+1
            # print(dic.keys(),dic.values())

        h_count = self.add_modle.rowCount()
        index =0
        for i in range(0,h_count):
            item = self.add_modle.verticalHeaderItem(i)
            txt = item.text()
            if txt == '开始时间':
                index = i
        self.add_modle.setItem(index,QStandardItem(str(self.time_stamp)))

    def init_watch_table_combobox(self):
           try:
                cmd = "CREATE TABLE IF NOT EXISTS %s(%s)" % (self.table_name, self.lableString)
                self.cursor.execute(cmd)
                # fae
                self.comboBox_person.addItem("不限")
                cmd = "SELECT FAE FROM %s" % (self.table_name)
                self.cursor.execute(cmd)
                data = self.cursor.fetchall()
                persons = set()
                for row in data:
                    pass
                    persons.add(row)
                print(persons)
                for per in persons:
                    print(per[0])
                    self.comboBox_person.addItem(per[0])
                #state
                self.comboBox_state.addItem('不限')
                cmd = "SELECT 问题状态 FROM %s" % (self.table_name)
                self.cursor.execute(cmd)
                data = self.cursor.fetchall()
                states = set()
                for st in data:
                    pass
                    states.add(st)
                print(states)
                for st in states:
                    print(st[0])
                    self.comboBox_state.addItem(st[0])
                self.comboBox_state.setCurrentText('进行中')
                #type
                self.comboBox_type.addItem('不限')
                cmd = "SELECT 问题类型 FROM %s" % (self.table_name)
                self.cursor.execute(cmd)
                data = self.cursor.fetchall()
                types = set()
                for tp in data:
                    pass
                    types.add(tp)
                print(types)
                for tp in types:
                    print(tp[0])
                    self.comboBox_type.addItem(tp[0])
           except Exception as e:
                print("init_watch_table_display",str(e))

    def init_watch_table_display(self):

        self.init_watch_table_combobox()

        self.watch_modle = QStandardItemModel(0, len(self.labels_names))
        self.watch_table_view.setModel(self.watch_modle)
        # self.watch_table_view.horizontalHeader().setStretchLastSection(True)
        # self.watch_table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        i = 0
        for dic in self.labels_names:
            h = str(dic.keys());
            h = h.replace(u"dict_keys(['", "");
            h = h.replace(u"'])", "")
            self.watch_modle.setHorizontalHeaderItem(i, QStandardItem(h))
            i = i + 1
        # 查询数据库数据
        cmd = "SELECT %s FROM %s WHERE 问题状态 LIKE '进行中'" %(self.lableString,self.table_name)
        self.cursor.execute(cmd)
        list = self.cursor.fetchall()
        print("搜到数据")
        i=0
        for row in list:
            print(row)
            j=0
            for d in row:
                self.watch_modle.setItem(i,j,QStandardItem(d))
                j=j+1
            i = i + 1

    def init_edit_table_display(self):
        pass


    def refresh_view_table(self):
        self.watch_modle.clear()

        print("watch_table_changed")
        # c
        i = 0
        for dic in self.labels_names:
            h = str(dic.keys());h = h.replace(u"dict_keys(['", "");h = h.replace(u"'])", "")
            self.watch_modle.setHorizontalHeaderItem(i, QStandardItem(h))
            i = i + 1

        try:
            v_person= self.comboBox_person.currentText()
            if v_person== '不限':
                v_person = '%'
            else:
                v_person = '%'+v_person+'%'
            v_state = self.comboBox_state.currentText()
            if v_state== "不限":
                v_state = '%'
            else:
                v_state = '%'+v_state+'%'
            v_type = self.comboBox_type.currentText()
            if v_type=="不限":
                v_type = '%'
            else:
                v_type = '%'+v_type+'%'
        except Exception as e:
            print(str(e))

        cmd = "SELECT %s FROM %s WHERE 问题类型 LIKE '%s' AND 问题状态 LIKE '%s' AND FAE LIKE '%s'" %(self.lableString,self.table_name,v_type,v_state,v_person)
        print(cmd)
        try:
            self.cursor.execute(cmd)
        except Exception as e:
            print("refresh_view_table",str(e))
        list = self.cursor.fetchall()
        print("搜到数据")
        i = 0
        for row in list:
            print(row)
            j = 0
            for d in row:
                self.watch_modle.setItem(i, j, QStandardItem(d))
                j = j + 1
            i = i + 1


    def on_click_watch_table_view(self, model_index):
        pass
        print("add:",model_index.row(),model_index.column())
        # QMessageBox.information(self,"提示","隐藏当前列",QMessageBox.Yes|QMessageBox.No)

    def watch_table_right_clicked(self,pos):
        pass
        print("点击右键")
        self.contextMenu.move(QtGui.QCursor.pos())
        self.contextMenu.show()
        #
    def on_click_save_watch_table_view(self):
        pass
        print("save changed....")
        rows_count = self.watch_modle.rowCount()
        column_count=self.watch_modle.columnCount()
        uuid_index = 0
        for j in range(0,column_count):
            value = self.watch_modle.horizontalHeaderItem(j).text()
            if value == 'UUID':
                uuid_index = j
        print("UUID在",uuid_index,"列")

        try:
            for i in range(0,rows_count):
                for j in range(0,column_count):
                    cmd = "UPDATE %s SET %s = '%s' WHERE UUID='%s'"%(self.table_name,self.watch_modle.horizontalHeaderItem(j).text(),self.watch_modle.item(i,j).text(),self.watch_modle.item(i,uuid_index).text())
                    print(cmd)
                    self.cursor.execute(cmd)
            self.connect.commit()
            print("保存成功")
            QMessageBox.information(self, "提示","保存更改成功",QMessageBox.Yes)

        except Exception as e:
            print(str(e))
            QMessageBox.information(self, "提示", "保存更改失败", QMessageBox.Yes)



    def watch_table_pressed(self,model_index):

        print("pressed:" ,model_index.row(),model_index.column())
        self.watch_table_view_row = model_index.row()
        self.watch_table_view_column = model_index.column()


    def creat_watch_tabel_view_mecu(self):

        self.contextMenu = QMenu(self)
        self.actionA = self.contextMenu.addAction("隐藏行")
        self.actionB = self.contextMenu.addAction("隐藏列")
        self.actionA.triggered.connect(self.menu_hide_row_action_hander)
        self.actionB.triggered.connect(self.menu_hide_colum_action_hander)

    def menu_hide_row_action_hander(self):
        pass
        print("隐藏行")
        self.watch_modle.removeRow(self.watch_table_view_row)

    def menu_hide_colum_action_hander(self):
        pass
        try:
            print("隐藏列")
            value = self.watch_modle.horizontalHeaderItem(self.watch_table_view_column).text()
            if value == 'UUID':
                QMessageBox.information(self,"提示","不能删除UUID",QMessageBox.Yes)
            else:
                self.watch_modle.removeColumn(self.watch_table_view_column)
        except Exception as e:
            print(str(e))

    def pushButton_add_database_clicked(self):
        # len = uuid + length
        self.dic = {}
        try:
            leng = len(self.labels_names)
            uuid_cur = uuid.uuid1()
            for i in range(0, leng):
                item = self.add_modle.item(i, 0)
                lable = self.add_modle.verticalHeaderItem(i)
                if item:
                    self.dic[lable.text()] = item.text()
                else:
                    self.dic[lable.text()] = ""
        except Exception as e:
            print(str(e))
        self.dic["UUID"] = uuid_cur.hex
        print("pushbutton-clicked")
        try:
            cmd ="CREATE TABLE IF NOT EXISTS %s(%s)" %(self.table_name,self.lableString)
            self.cursor.execute(cmd)
        except Exception as e:
            print(str(e))
        self.inset_data(self.dic)
    def pushButton_out_excle_hander(self):
        pass
        print("点击[导出excle]")
        count_rows = self.watch_modle.rowCount()
        count_column = self.watch_modle.columnCount()

        work_book = xlwt.Workbook(encoding='utf-8')
        sheet = work_book.add_sheet("今日内容")

        for j in range(0,count_column):
            item = self.watch_modle.horizontalHeaderItem(j)
            sheet.write(0, j, item.text())

        for i in range(0,count_rows):
            for j in range(0,count_column):
                item =  self.watch_modle.item(i,j)
                sheet.write(i+1,j,item.text())
        work_book.save("今日内容-%s.xls"%(self.time_stamp))
        QMessageBox.information(self,"提示","导出成功,请到当前目录查看！",QMessageBox.Yes)


    def inset_data(self, dic):
        key   = [""]*len(dic)
        value = [""]*len(dic)
        i = 0
        for v in dic:
            key[i] = v
            value[i] = dic[v]
            i = i+1
        keys = str(key).replace("[","");
        keys = keys.replace("]","")
        values = str(value).replace("[","")
        values = values.replace("]","")
        # print(keys)
        # print(values)
        try:
            cmd = "INSERT INTO zhjk_support(%s) VALUES (%s)" % (keys, values)
            print(cmd)
            self.cursor.execute(cmd)
            self.connect.commit()
            print("插入数据成功:", dic)
            QMessageBox.information(self,"提示","提交成功!")
            self.comboBox_state.clear()
            self.comboBox_type.clear()
            self.comboBox_person.clear()
            self.init_watch_table_display()
            # self.add_modle.removeColumn(0)
            # self.add_modle.appendColumn(QStandardItem())
            # self.add_modle.setItem(0, 0, QtGui.QStandardItem(self.time_stamp))
            # self.add_modle.setHorizontalHeaderLabels(['相关信息/详细填写'])
        except Exception as e :
            print(str(e))
            QMessageBox.information(self, "提示", "提交失败!")
    def refresh_app(self):
        qApp.processEvents()
class Custum_complains(QThread):
      # const
      def  __init__(self):
          super(Custum_complains, self).__init__()
      def run(self):
          pass
          try:
              # 串口工作主流程
              """主循环"""
              while True:
                pass
                time.sleep(0.1)
          except Exception as e:
                print(str(e))

      def mainloop_app(self):
          try:
              pass
              app = QtWidgets.QApplication(sys.argv)
              window = MyApp()
              window.show()
              pass
          except Exception as e:
              print(str(e))
          finally:
              sys.exit(app.exec_())

if __name__ == "__main__":
    try:
        custum = Custum_complains()
        custum.start()
        custum.mainloop_app()
    except Exception as e:
        print(str(e))
    finally:
        pass




