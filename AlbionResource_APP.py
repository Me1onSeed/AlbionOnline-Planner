import sys
import pyperclip
import winsound
import win32com.client as win
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidgetItem, QCompleter, QShortcut
from PyQt5.QtGui import QIcon, QColor, QKeySequence
from PyQt5.QtCore import Qt, QDateTime, QTime, QTimer

import AlbionResource_UI  # 读取UI文件


if __name__ == '__main__':
    # 创建APP     这里的注释详见PyIUC.txt文件
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = AlbionResource_UI.Ui_MainWindow()  # AlbionResource_UI是文件名,UI_MainWindow()是属性
    ui.setupUi(MainWindow)
    MainWindow.show()


    # MapCompleter 地图名自动填充
    def fillMapName(completion):
        tc = ui.lineMap.cursor()
        extraLength = (len(completion) - len(ui.lineMap.text()))
        tc.movePosition(tc.Left)
        tc.select(tc.WordUnderCursor)
        tc.insertText(completion)
        ui.lineMap.setCursor(tc)
        ui.lineMap.setCursorPosition(extraLength)
#    ui.mapCompleter.activated.connect(fillMapName)


    # comboBox 更改资源类型
    def RTypeChange():    # 资源类型
        RType = ui.comboBoxRType.currentText()

        if RType == '能量核心（球）' or RType == '能量水晶（风）' :
            ui.labelRColor.setText("资源颜色")
            ui.comboBoxRColor.setEnabled(True)
            ui.comboBoxRColor.clear()
            ui.comboBoxRColor.addItems(["绿色", "蓝色", "紫色", "金色"])
            ui.comboBoxRLvl.clear()
            ui.comboBoxRLvl.addItem("-")
            ui.comboBoxRLvl.setDisabled(True)
            # 资源颜色 绿-金 / 资源等级 -

        elif RType == '野外宝箱':
            ui.labelRColor.setText("宝箱大小")
            ui.comboBoxRColor.setEnabled(True)
            ui.comboBoxRColor.clear()
            ui.comboBoxRColor.addItems(["小箱", "中箱", "大箱"])
            ui.comboBoxRLvl.clear()
            ui.comboBoxRLvl.addItem("")
            ui.comboBoxRLvl.setDisabled(True)
            # 宝箱大小 小-大 / 资源等级 (*idea:地图等级)

        elif RType == '采集资源(.4)':
            ui.labelRColor.setText("采集资源种类")
            ui.comboBoxRColor.setEnabled(True)
            ui.comboBoxRColor.clear()
            ui.comboBoxRColor.addItem(QIcon('icons/hide.png'), "皮")
            ui.comboBoxRColor.addItem(QIcon('icons/fiber.png'), "棉")
            ui.comboBoxRColor.addItem(QIcon('icons/wood.png'), "木")
            ui.comboBoxRColor.addItem(QIcon('icons/ore.png'), "矿")
            ui.comboBoxRColor.addItem(QIcon('icons/stone.png'), "石")
            ui.comboBoxRLvl.setEnabled(True)
            ui.comboBoxRLvl.clear()
            ui.comboBoxRLvl.addItems(["T4.4", "T5.4", "T6.4", "T7.4", "T8.4"])
            # 资源种类 不同采集资源种类 / 资源等级 4.4-8.4

        elif RType == '猛犸象':
            ui.labelRColor.setText("资源品质（颜色）")
            ui.comboBoxRColor.clear()
            ui.comboBoxRColor.addItem("-")
            ui.comboBoxRColor.setDisabled(True)
            ui.comboBoxRLvl.clear()
            ui.comboBoxRLvl.addItem("-")
            ui.comboBoxRLvl.setDisabled(True)
            # 资源品质（颜色） - / 资源等级 -

        elif RType in ('城堡', '哨站'):
            ui.labelRColor.setText("箱子颜色")
            ui.comboBoxRColor.setEnabled(True)
            ui.comboBoxRColor.clear()
            ui.comboBoxRColor.addItems(["绿色", "蓝色", "紫色", "金色"])
            ui.comboBoxRLvl.clear()
            ui.comboBoxRLvl.addItem("-")
            ui.comboBoxRLvl.setDisabled(True)
            # 箱子颜色 绿-金 / 资源等级 -

        elif RType == '领地':
            ui.labelRColor.setText("领地类型")
            ui.comboBoxRColor.setEnabled(True)
            ui.comboBoxRColor.clear()
            ui.comboBoxRColor.addItems(["农场领地", "资源领地"]) # ! 
            ui.comboBoxRLvl.clear()
            ui.comboBoxRLvl.addItem("")
            ui.comboBoxRLvl.setDisabled(True)
            # 领地类型 ... / 资源等级 (*idea:地图等级)
    ui.comboBoxRType.currentTextChanged.connect(RTypeChange)


    def timeTypeChange(): # 时间类型改变tab顺序
        ifClock = ui.tabTimeOrClock.currentIndex()
        ifLocal = ui.tabTimeZone.currentIndex()
        if ifClock == 0: # tab的Index: 0是剩余时间
            return
        elif ifClock == 1: # 1是解锁时刻
            MainWindow.setTabOrder(ui.tabTimeOrClock, ui.tabTimeZone)
            if ifLocal == 0: # utc
                MainWindow.setTabOrder(ui.tabTimeZone, ui.timeUTC)
                MainWindow.setTabOrder(ui.timeUTC, ui.ButtonAdd)
            elif ifLocal == 1: # utc+8
                MainWindow.setTabOrder(ui.tabTimeZone, ui.timeBeijing)
                MainWindow.setTabOrder(ui.timeBeijing, ui.ButtonAdd)

    ui.tabTimeOrClock.currentChanged.connect(timeTypeChange)
    ui.tabTimeZone.currentChanged.connect(timeTypeChange)

    Speak = win.Dispatch("SAPI.SpVoice") # 语音播报声音

    ## ButtonAdd 添加资源
    # 先建立timer 按下buttonAdd后 计时才会开始
    timer = QTimer()

    def addResource():
        # table中添加一行资源信息
        ui.table.insertRow(0)

        # (0,0)留给资源图标 icon 要在(0,2)资源信息确定后
        ui.table.setItem(0,0,QTableWidgetItem())
        # (0,1)留给倒计时, 要随时间更新
        ui.table.setItem(0,1,QTableWidgetItem())

        # (0,2)资源信息
        RType = ui.comboBoxRType.currentText()
        RColor = ui.comboBoxRColor.currentText()
        RLvl = ui.comboBoxRLvl.currentText()

        if RType in ('能量核心（球）', '能量水晶（风）', '城堡', '哨站'):
            color = RColor.replace('色','') # 去除了色字 '蓝' '绿'
            if RType == '能量核心（球）':
                resourceString = '球'
            elif RType == '能量水晶（风）':
                resourceString = '风'
            elif RType == '城堡':
                resourceString = '城堡'
            elif RType == '哨站':
                resourceString = '哨站'
            resourceString = color+' '+resourceString # '紫 城堡' '金 风'
        elif RType == '野外宝箱':
            resourceString = RColor # '小箱'
        elif RType == "采集资源(.4)":
            resourceString = RLvl[1:]+' '+RColor # '7.4 皮'
        else:   # 猛犸/领地
            resourceString = RType # '领地'

        ui.table.setItem(0,2,QTableWidgetItem(resourceString))
        ui.table.item(0,2).setTextAlignment(Qt.AlignCenter)

        # 资源图标 icon
        #icon = QIcon("icons/*.png")
        iconTableItem = ui.table.item(0,0)
        #iconTableItem.setIcon(icon)

        # (0,3)地图
        map = ui.lineMap.text() # idea:QCompleter
        ui.table.setItem(0,3,QTableWidgetItem(map))
        ui.table.item(0,3).setTextAlignment(Qt.AlignCenter)

        # (0,4)备注
        if RType == '领地':
            note = RColor
        else:
            note = ""
        ui.table.setItem(0,4,QTableWidgetItem(note))
        ui.table.item(0,4).setTextAlignment(Qt.AlignCenter)

        # (0,5)资源解锁utc时间
        dateTime_utc = QDateTime.currentDateTimeUtc()
        time_utc = dateTime_utc.time()
        time_local = QTime.currentTime()
        ifClock = ui.tabTimeOrClock.currentIndex()
        ifLocal = ui.tabTimeZone.currentIndex()
        if ifClock == 0: # tab的Index: 0是剩余时间
            time_remain = ui.timeTimeRemain.time()
            seconds_remain = abs(time_remain.secsTo(QTime(0,0))) # 距离解锁差多少ms
            clock_utc = time_utc.addSecs(seconds_remain)
            clock_local = time_local.addSecs(seconds_remain)
        elif ifClock == 1: # 1是解锁时刻
            if ifLocal == 0: # utc
                clock_utc = ui.timeUTC.time()
                clock_local = clock_utc.addSecs(60*60*8) # utc+8
            elif ifLocal == 1: # utc+8
                clock_local = ui.timeBeijing.time()
                clock_utc = clock_local.addSecs(-60*60*8)
            seconds_remain = time_utc.secsTo(clock_utc)
            time_remain = QTime(0,0).addSecs(seconds_remain)
        ui.table.setItem(0,5,QTableWidgetItem(clock_utc.toString())) # 以hh:mm:ss格式字符串输出utc时间
        ui.table.item(0,5).setTextAlignment(Qt.AlignCenter)

        rIndex = ui.table.rowCount() - 1 # 如果执行完上面之后有15行,说明刚添加的这个是第14个资源
        timer.start() # add资源后,timer计时才开始,表中的倒计时才开始更新
        if_time_up = 0
        return if_time_up # 输出是否到期，用于后面到期时的警报
    ui.ButtonAdd.clicked.connect(addResource)
    shortcutAdd = [None]*6  # 快捷键回车也可添加 Return是大键盘回车 Enter是小键盘回车
    shortcutAdd[0] = QShortcut(QKeySequence('Enter'), ui.timeTimeRemain)
    shortcutAdd[1] = QShortcut(QKeySequence('Enter'), ui.timeBeijing)
    shortcutAdd[2] = QShortcut(QKeySequence('Enter'), ui.timeUTC)
    shortcutAdd[3] = QShortcut(QKeySequence('Return'), ui.timeTimeRemain)
    shortcutAdd[4] = QShortcut(QKeySequence('Return'), ui.timeBeijing)
    shortcutAdd[5] = QShortcut(QKeySequence('Return'), ui.timeUTC)
    for i in range(6):
        shortcutAdd[i].activated.connect(addResource)


    # 倒计时更新
    def updateTime():
        #timer.setInterval(timerInterval) 更新频率
        rowCount = ui.table.rowCount()
        for row in range(0,rowCount-1): # 一共有rowCount行
            if not ui.table.item(row,5) == None:
                clock_utc_string = ui.table.item(row,5).text() # 把该资源在第5列存的utc时间再读取为字符串, 进而转化为QTime格式
                clock_utc = QTime.fromString(clock_utc_string,"HH:mm:ss")
                time_utc = QDateTime.currentDateTimeUtc().time() # 读取当前utc时间, 进而计算剩余时间, 进而化为QTime格式
                seconds_remain = time_utc.secsTo(clock_utc)
                if seconds_remain == 0: # 剩余时间为0时，utc时间改换，否则回到23h59min59s, 并且整行变成红色来提醒
                    ui.table.item(row,5).setText('已到期')
                    for column in range(6):
                        if not ui.table.item(row,column) == None:
                            ui.table.item(row,column).setForeground(QColor('red'))
                time_remain = QTime(0,0).addSecs(seconds_remain)
                hh = time_remain.hour()
                mm = time_remain.minute()
                ss = time_remain.second()
                if hh == 0:
                    if mm == 0:
                        tstr = str(ss)+'s '
                    else:
                        tstr = str(mm)+'min '+str(ss)+'s '
                else:
                    tstr = str(hh)+'h '+str(mm)+'min '+str(ss)+'s '
                ui.table.item(row,1).setText(tstr)
                ui.table.item(row,1).setTextAlignment(Qt.AlignCenter)
    timer.timeout.connect(updateTime)


    # 删除资源
    def delResource():
        currentRowNum = ui.table.currentRow()
        if not ui.table.item(currentRowNum,2) == None:  # 若资源类型不为空,就删掉这行
            ui.table.removeRow(currentRowNum)
    ui.ButtonDel.clicked.connect(delResource)


    # 删除全部资源
    def clearAllResource(): #!! 搞个确认弹窗
        for row in range(ui.table.rowCount()):
            ui.table.removeRow(row)
    ui.ButtonClear.clicked.connect(clearAllResource)


    # 复制到剪贴板
    shortcutClip = QShortcut(QKeySequence('Ctrl+C'),ui.table) # 快捷键ctrl+C(只在表格里有用)
    def toClipboard():
        currentRowNum = ui.table.currentRow()
        if ui.table.item(currentRowNum,2) == None: # 为空行时清空剪贴板，跳出
            pyperclip.copy('')
            return
        RInformation = ui.table.item(currentRowNum,2).text()
        # 图标及信息
        if RInformation[-2:-1] in ('城堡', '哨站'):
            copyInfor = '(城堡) '+RInformation
        elif RInformation[-1] in ('球','风'):
            if RInformation[-1] == '球':
                copyInfor2 = '核心)'
            elif RInformation[-1] == '风':
                copyInfor2 = '水晶)'
            if RInformation[0] != '金':
                copyInfor1 = '('+RInformation[0]+'色'
            else:
                copyInfor1 = '(橘色'
            copyInfor = copyInfor1+copyInfor2
        elif RInformation[-1] == '箱':
            copyInfor = RInformation
        else:
            copyInfor = RInformation
        # 地图
        RMap = ui.table.item(currentRowNum,3).text()
        copyMap = '@'+RMap
        # 剩余时间
        copyTimeRemain = ui.table.item(currentRowNum,1).text()

        copyString = copyInfor+' '+copyMap+' (计时)'+copyTimeRemain
        pyperclip.copy(copyString)
        winsound.Beep(4000,200)
    ui.ButtonClip.clicked.connect(toClipboard)
    shortcutClip.activated.connect(toClipboard)


    # 剩余时间后，自动复制至剪贴板，并发出提示
    def reminder():
        rowCount = ui.table.rowCount()
        time_utc = QDateTime.currentDateTimeUtc().time() # 读取当前utc时间, 进而计算剩余时间, 进而化为QTime格式
        for row in range(0,rowCount-1):
            clock_utc_string = ui.table.item(row,5).text() # 读取资源在第5列存的utc时间
            clock_utc = QTime.fromString(clock_utc_string,"HH:mm:ss")
            seconds_remain = time_utc.secsTo(clock_utc)
            if seconds_remain in [0, 60, 120, 300]:   # 5min 2min 1min
                if seconds_remain == 300: # 5min 1beep
                    winsound.Beep(500,1000)
                elif seconds_remain == 120: # 2min 2beep
                    winsound.Beep(500,500)
                elif seconds_remain == 60: # 1min 3beep
                    winsound.Beep(500,400)
                if not ui.table.item(row,2) == None:
                    RInformation = ui.table.item(row,2).text()
                    # 图标及信息
                    if RInformation[-2:-1] in ('城堡', '哨站'):
                        copyInfor = '(城堡) '+RInformation
                    elif RInformation[-1] in ('球','风'):
                        if RInformation[-1] == '球':
                            copyInfor2 = '核心)'
                        elif RInformation[-1] == '风':
                            copyInfor2 = '水晶)'
                        if RInformation[0] != '金':
                            copyInfor1 = '('+RInformation[0]+'色'
                        else:
                            copyInfor1 = '(橘色'
                        copyInfor = copyInfor1+copyInfor2
                    else:
                        copyInfor = RInformation
                    # 地图
                    RMap = ui.table.item(row,3).text()
                    copyMap = '@'+RMap
                    # 剩余时间
                    copyTimeRemain = ui.table.item(row,1).text()
                    copyString = copyInfor+' '+copyMap+' (计时)'+copyTimeRemain
                    pyperclip.copy(copyString)
    timer.timeout.connect(reminder)


    # 退出
    sys.exit(app.exec_())
