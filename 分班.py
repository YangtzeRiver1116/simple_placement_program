import xlrd
import xlwt
import operator
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QPushButton, QLabel, QFrame, \
    QInputDialog, qApp, QMessageBox, QFileDialog
import sys
import time


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.title = '分班助手'
        self.initUI()
        self.row_num = 7
        self.sex_list_n = 2
        self.input_file = ''
        self.output_path = ''
        self.file_name = ''
        self.standard = 1
        self.class_all = 1
        self.output_style = '原文件格式'
        self.top_name = []
        self.data_list = []

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(100, 100, 640, 480)

        layout_all = QVBoxLayout()

        # 路径设置
        choice_Box = QGroupBox('路径设置')
        choice_Box_layout = QVBoxLayout()
        choice_input_layout = QHBoxLayout()
        choice_output_file_layout = QHBoxLayout()
        choice_output_layout = QHBoxLayout()
        button_input_data_path = QPushButton('选择', self)
        button_input_data_path.clicked.connect(self.file_choice)
        button_output_data_path = QPushButton('选择', self)
        button_output_data_path.clicked.connect(self.path_choice)
        button_output_data_name = QPushButton('输入', self)
        button_output_data_name.clicked.connect(self.set_output_file_name)
        label_input_data_path = QLabel('分班数据载入文件目录：', self)
        label_output_data_file_name = QLabel('输入保存文件名：', self)
        label_output_data_path = QLabel('分班数据输出文件目录：', self)
        self.label_input_data_show = QLabel('', self)
        self.label_input_data_show.setFrameStyle(QFrame.Panel | QFrame.Sunken)  # 面版凹陷
        self.label_output_data_show = QLabel('', self)
        self.label_output_data_show.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        self.label_output_data_file_name_show = QLabel('', self)
        self.label_output_data_file_name_show.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        choice_input_layout.addWidget(label_input_data_path)
        choice_input_layout.addWidget(self.label_input_data_show)
        choice_input_layout.addWidget(button_input_data_path)
        choice_output_file_layout.addWidget(label_output_data_file_name)
        choice_output_file_layout.addWidget(self.label_output_data_file_name_show)
        choice_output_file_layout.addWidget(button_output_data_name)
        choice_output_layout.addWidget(label_output_data_path)
        choice_output_layout.addWidget(self.label_output_data_show)
        choice_output_layout.addWidget(button_output_data_path)
        choice_Box_layout.addLayout(choice_input_layout)
        choice_Box_layout.addLayout(choice_output_file_layout)
        choice_Box_layout.addLayout(choice_output_layout)
        choice_Box.setLayout(choice_Box_layout)
        layout_all.addWidget(choice_Box)

        # 参数设置
        set_box = QGroupBox('参数设置')
        set_box_layout = QVBoxLayout()

        # 分班总数
        set_class_layout = QHBoxLayout()
        set_class_label = QLabel('请选择分班总数：', self)
        self.class_all_label = QLabel('1', self)
        self.class_all_label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        set_class_button = QPushButton('选择')
        set_class_button.clicked.connect(self.set_class_number)
        set_class_layout.addWidget(set_class_label)
        set_class_layout.addWidget(self.class_all_label)
        set_class_layout.addWidget(set_class_button)

        set_box_layout.addLayout(set_class_layout)

        # 排列标准
        set_standard_layout = QHBoxLayout()
        set_standard_label = QLabel('请选择排列标准：', self)
        self.standard_label = QLabel('', self)
        self.standard_label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        set_standard_button = QPushButton('选择')
        set_standard_button.clicked.connect(self.set_standard)
        set_standard_layout.addWidget(set_standard_label)
        set_standard_layout.addWidget(self.standard_label)
        set_standard_layout.addWidget(set_standard_button)
        set_box_layout.addLayout(set_standard_layout)

        # 输出格式
        set_output_layout = QHBoxLayout()
        set_output_label = QLabel('请选择输出形式：', self)
        self.output_ = QLabel('原文件格式', self)
        self.output_.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        set_output_button = QPushButton('选择')
        set_output_button.clicked.connect(self.set_output_standard)
        set_output_layout.addWidget(set_output_label)
        set_output_layout.addWidget(self.output_)
        set_output_layout.addWidget(set_output_button)
        set_box_layout.addLayout(set_output_layout)
        set_box.setLayout(set_box_layout)
        layout_all.addWidget(set_box)

        # 确定和退出按钮
        fin_layout = QHBoxLayout()
        help_button = QPushButton('帮助', self)
        help_button.clicked.connect(self.help_function)
        about_button = QPushButton('关于Qt', self)
        about_button.clicked.connect(self.about_qt)
        use_button = QPushButton('确定', self)
        use_button.clicked.connect(self.sure_function)
        false_button = QPushButton('取消', self)
        false_button.clicked.connect(self.false_function)
        fin_layout.addWidget(about_button)
        fin_layout.addWidget(help_button)
        fin_layout.addStretch(1)
        fin_layout.addWidget(use_button)
        fin_layout.addWidget(false_button)

        layout_all.addLayout(fin_layout)

        self.setLayout(layout_all)
        self.show()

    def set_class_number(self):
        self.class_all, ok = QInputDialog.getInt(self, '分班总班级数', '请输入分班总班级数：', 1)
        if ok:
            self.class_all_label.setText(str(self.class_all))

    def set_standard(self):

        if self.input_file == '':
            replay = QMessageBox.question(self, '没有检测到输入文件!', '操作需要获取文件标题，\n是否重新选择输入文件？',
                                          QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if replay == QMessageBox.Yes:
                self.file_choice()
            else:
                self.false_function()

        list_choice = self.top_name

        standard_str, ok = QInputDialog.getItem(self, "排列标准", "请选择排列时使用的标准：", list_choice)
        for i in range(len(list_choice)):
            if list_choice[i] == standard_str:
                self.standard = i
        if ok:
            self.standard_label.setText(standard_str)

    def set_output_file_name(self):
        self.file_name, ok = QInputDialog.getText(self, "输出文件名", "请输入输出文件名")
        for i in range(len(self.file_name)):
            if self.file_name[i] == '.':
                str_file = self.file_name[i:]
                if str_file != '.xls':
                    self.file_name = self.file_name + '.xls'
                if str_file == '.xlsx':
                    del self.file_name[-1]
                break
            elif i == len(self.file_name) - 1:
                self.file_name = self.file_name + '.xls'
                break
        if ok:
            self.label_output_data_file_name_show.setText(self.file_name)

    def set_output_standard(self):
        list_choice = ['原文件格式', '仅姓名']

        self.output_style, ok = QInputDialog.getItem(self, "输出格式", "请选择输出时使用的格式：", list_choice)
        if ok:
            self.output_.setText(self.output_style)

    def choice_sex_list_n(self):
        class_all, ok = QInputDialog.getInt(self, '性别栏手动指定（首行序列为0）', '请输入性别栏：', 1)
        if ok:
            self.sex_list_n = class_all

    def sure_function(self):

        msgBox = QMessageBox()
        msgBox.setWindowTitle('询问')
        msgBox.setIcon(QMessageBox.Question)
        msgBox.setText('是否应用设置？')
        sure = msgBox.addButton('确定', QMessageBox.AcceptRole)
        msgBox.addButton('取消', QMessageBox.RejectRole)
        msgBox.setDefaultButton(sure)
        reply = msgBox.exec()
        if reply == QMessageBox.AcceptRole:
            if self.input_file == '':
                replay = QMessageBox.question(self, '没有检测到输入文件!', '是否重新选择输入文件？',
                                              QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                if replay == QMessageBox.Yes:
                    self.file_choice()
                else:
                    self.false_function()

            if self.file_name == '':
                replay = QMessageBox.question(self, '没有检测到输出文件名!', '是否重新输入输出文件名？',
                                              QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                if replay == QMessageBox.Yes:
                    self.set_output_file_name()
                elif replay == QMessageBox.No:
                    self.false_function()

            if self.output_path == '':
                replay = QMessageBox.question(self, '没有检测到输出路径!', '是否重新选择输出路径？',
                                              QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                if replay == QMessageBox.Yes:
                    self.path_choice()
                elif reply == QMessageBox.No:
                    self.false_function()

            if self.class_all == 1:
                replay = QMessageBox.question(self, '分班总数为1!', '是否重新输入？',
                                              QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                if replay == QMessageBox.Yes:
                    self.set_class_number()
                elif replay == QMessageBox.No:
                    self.false_function()

            self.data_list = self.sort_all(self.data_list, self.standard)
            devise_fin = self.devise_all(self.data_list)
            devise_fin = self.sex_balance(devise_fin, self.class_all)
            if self.output_style == '原文件格式':
                self.write_and_save_point(self.top_name, self.class_all, devise_fin, self.output_path, self.file_name)
            elif self.output_style == '仅姓名':
                self.write_and_save_name(self.class_all, devise_fin, self.output_path, self.row_num, self.file_name)
            reply = QMessageBox.question(self, '分班完成！', '分班成功！是否再次操作？',
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.No:
                self.false_function()

    def help_function(self):
        msgBox = QMessageBox(QMessageBox.NoIcon, '帮助', '版本号：0.1.18827 \n\n声明：\n本软件为免费软件，不存在任何收费行'
                                                       '为，如果有任何利益纠葛与开发人员无关。\n\n'
                                                       '使用说明：\n本程序为协助分班软件，不能完全代替人工分班。\n程序分班'
                                                       '时需要获取分班总表目录、分班输出文件名、分班输出路径等相关信息，请'
                                                       '按设置要求进行详细的参数设置以确保分班成功。\n程序默认考虑男女比例'
                                                       '因素，默认输出文件格式为xls格式，当以‘仅姓名’方式输出时，默认每'
                                                       '行7人\n\n'
                                                       '注意：\n本程序仅用于协助分班'
                                                       '，如用于其他用途造或不当使用并造成不良后果，'
                                                       '开发人员概不负责。\n如果程序崩溃，请退出重新打开尝试并严格按照设置'
                                                       '要求设置参数\n如有任何问题或改进建议请邮件至：koolo233@163.com\n'
                                                       '如果想获取源码，请访问：'
                                                       'https://github.com/YangtzeRiver1116/simple_placement_program')
        msgBox.exec()

    def about_qt(self):
        QMessageBox.aboutQt(self, '关于Qt')

    def false_function(self):
        qApp.quit()

    def file_choice(self):
        self.input_file, file_type = QFileDialog.getOpenFileName(self, '选择文件', './', 'Excel files(*.xlsx , *.xls)')
        self.label_input_data_show.setText(self.input_file)
        self.top_name, self.data_list = self.read_excel(self.input_file)

    def path_choice(self):
        self.output_path = QFileDialog.getExistingDirectory(self, "选取文件夹", "./") + '/'
        self.label_output_data_show.setText(self.output_path)

    def read_excel(self, input_path):
        # 打开文件
        workbook = xlrd.open_workbook(str(input_path), 'wr')
        while True:
            if not workbook:
                replay = QMessageBox.critical('无法打开文件', '是否重试？',
                                              QMessageBox.Retry | QMessageBox.Abort, QMessageBox.Retry)
                if replay == QMessageBox.Retry:
                    workbook = xlrd.open_workbook(str(input_path))
            else:
                break

        # 获取sheet
        sheet_all = workbook.sheet_by_index(0)

        # 获取第一行
        name_top = sheet_all.row_values(0)

        # 所有数据
        list_all = []
        for i in range(len(sheet_all.col_values(0, start_rowx=1))-1):
            list_all.append(sheet_all.row_values(i+1))

        '''
        # 获取第三列：分类
        whether_true = sheet_all.col_values(3, start_rowx=1)
        '''
        '''
        delete_list = []
        for i in range(len(whether_true)):
            if whether_true[i] == '否':
                delete_list.append(i)
    
        for i in delete_list:
            del name[i]
            del point_all[i]
            # del sex_all[i]
        '''

        return name_top, list_all

    def sort_all(self, list_all, sort_standard):

        # 排序
        list_all = sorted(list_all, key=operator.itemgetter(sort_standard))
        return list_all

    def devise_all(self, name_all):
        batch_num = len(name_all) // self.class_all
        people_loss = len(name_all) % self.class_all
        all_class = []

        for i in range(self.class_all):
            all_class.append([])

        for i in range(batch_num):
            if i % 2 == 0:
                for c in range(self.class_all):
                    all_class[c].append(name_all[i*self.class_all + c])
            else:
                for c in range(self.class_all):
                    all_class[c].append(name_all[i*self.class_all + self.class_all - 1 - c])

        if not people_loss == 0:
            if batch_num % 2 == 1:
                for c in range(people_loss):
                    all_class[c].append(name_all[batch_num*self.class_all + c])
            else:
                for c in range(people_loss):
                    all_class[c].append(name_all[batch_num*self.class_all + people_loss - 1 - c])

        return all_class

    def write_and_save_point(self, top, sheet_num, list_data, path_name, file_name):
        # 创建
        fin = xlwt.Workbook()
        row_0 = top
        for i in range(sheet_num):
            table = fin.add_sheet('%s 班' % (i+1))
            for t in range(len(row_0)):
                table.write(0, t, row_0[t])
            for y in range(len(list_data[i])):
                for item in range(len(list_data[i][y])):
                    table.write(y+1, item, list_data[i][y][item])
        fin.save(path_name + file_name)

    def write_and_save_name(self, sheet_num, list_data, path_name, every_row_n, file_name):
        # 创建
        fin = xlwt.Workbook()
        list_all = 0
        for i in range(sheet_num):
            table = fin.add_sheet('%s 班' % (i+1))
            for y in range(len(list_data[i])):
                if y % every_row_n == 0:
                    list_all += 1
                table.write(list_all-1, y % every_row_n, list_data[i][y][0])
            list_all = 0
        fin.save(path_name + file_name)

    def sex_balance(self, list_all, class_n):
        for i in range(len(list_all[0][0])):
            if list_all[0][0][i] == '男' or list_all[0][0][i] == '女':
                self.sex_list_n = i
                break
            elif i == len(list_all[0][0])-1:
                msg = QMessageBox()
                msg.setWindowTitle('警告')
                msg.setIcon(QMessageBox.Warning)
                msg.setText('没有检测到性别栏!')
                msg.setInformativeText('是否手动指定?')
                Save = msg.addButton('是', QMessageBox.AcceptRole)
                msg.addButton('否', QMessageBox.RejectRole)
                msg.setDefaultButton(Save)
                reply = msg.exec()
                if reply == QMessageBox.AcceptRole:
                    self.choice_sex_list_n()
                elif reply == QMessageBox.RejectRole:
                    QMessageBox.critical(self, '错误', '缺少关键数据，程序5秒后关闭！')
                    time.sleep(1)
                    QMessageBox.critical(self, '错误', '缺少关键数据，程序4秒后关闭！')
                    time.sleep(1)
                    QMessageBox.critical(self, '错误', '缺少关键数据，程序3秒后关闭！')
                    time.sleep(1)
                    QMessageBox.critical(self, '错误', '缺少关键数据，程序2秒后关闭！')
                    time.sleep(1)
                    QMessageBox.critical(self, '错误', '缺少关键数据，程序1秒后关闭！')
                    time.sleep(1)
                    qApp.quit()

        class_sex = []
        man_n = 0
        woman_n = 0
        man_all = 0
        woman_all = 0
        for i in range(class_n):
            for y in range(len(list_all[i])):
                if list_all[i][y][self.sex_list_n] == '男':
                    man_n += 1
                    man_all += 1
                else:
                    woman_n += 1
                    woman_all += 1
            class_sex.append([i, woman_n, man_n])
            man_n = 0
            woman_n = 0

        for i in range(class_n):
            class_sex[i].append(class_sex[i][1] - woman_all//class_n)

        class_sex = sorted(class_sex, key=operator.itemgetter(3))

        change_class = class_n - 1
        flag = True
        for i in range(class_n):
            while 1:
                if class_sex[i][3] < 0:
                    while 1:
                        for a in range(len(list_all[class_sex[i][0]])):
                            if list_all[class_sex[i][0]][a][self.sex_list_n] != \
                                    list_all[class_sex[change_class][0]][a][self.sex_list_n] and \
                                    len(list_all[class_sex[change_class][0]]) > a and \
                                    class_sex[change_class][3] > 0 and \
                                    list_all[class_sex[i][0]][a][self.sex_list_n] == '男':
                                list_bet = list_all[class_sex[i][0]][a]
                                list_all[class_sex[i][0]][a] = list_all[class_sex[change_class][0]][a]
                                list_all[class_sex[change_class][0]][a] = list_bet
                                class_sex[i][3] += 1
                                class_sex[change_class][3] -= 1
                                flag = False
                                break
                            elif class_sex[change_class][3] <= 0:
                                change_class -= 1
                                flag = False
                                break
                        if not flag:
                            break
                elif change_class <= i:
                    break
                else:
                    break
                flag = True
        return list_all


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())
