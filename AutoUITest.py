
"""
1, 安装python 3.9.10 (因为目前wxPython暂不支持更高版本), 并配置系统环境变量Path
2, 命令行安装以下python库
pip install xlrd==1.2.0
pip install pillow
pip install wxpython
pip install pyperclip
pip install pyautogui
pip install opencv-python -i https://pypi.tuna.tsinghua.edu.cn/simple

如果某个库安装失败,可以手动下载对应 .whl 文件, 解压缩后, 放入{python安装目录}\Lib\site-packages
"""
import time
from datetime import datetime
import os
import shutil
import xlrd
import logging
import pyautogui
import pyperclip
import difflib
import filecmp

# 初始化pyautogui
pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True

# 初始化Result文件夹
_startTime = datetime.now().strftime("%Y%m%d_%H%M%S")
_resultFolder = f"Result_{_startTime}" #Result_YYYYMMDD_HHmmss
os.makedirs(_resultFolder)

# 初始化logging
_logger = logging.getLogger()
_logger.setLevel('DEBUG')
_formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
_ch=logging.StreamHandler()
_ch.setFormatter(_formatter)
_fh=logging.FileHandler(f"{_resultFolder}\\log.txt")
_fh.setFormatter(_formatter)
_logger.addHandler(_ch)
_logger.addHandler(_fh)

# 初始化测试环境
_testFile = "test.xlsx"
_retry = 3
_precision = 0.9
_duration = 0.2

# 定位目标位置，有png定位图片，否则返回鼠标位
def getLocation(image):
    #定位png图片
    if image.endswith(".png"):
        try:
            location=pyautogui.locateCenterOnScreen(image,confidence = _precision)
            if location is None:
                _logger.error(f"未能定位到{image}")
            else:
                _logger.info(f"定位{image}成功, location = {location}")
        except Exception as e:
            _logger.error("定位图片发生异常:" + str(e))
            return None
    else: #如果没有目标图片，返回鼠标位置
        location = pyautogui.position()
        _logger.info(f"鼠标位置, position = {location}")

    return location

# 字符串 to float, 如果失败返回 0
def tryParseToFloat(value):
    try:
        return float(value)
    except:
        return 0
        
# 鼠标单击
def mouseClick(image):
    location=getLocation(image)
    if location is None:
        return False
    else:
        pyautogui.leftClick(location.x,location.y, duration=_duration) 
        return True

# 鼠标双击
def mouseDoubleClick(image):
    location=getLocation(image)
    if location is None:
        return False
    else:
        pyautogui.doubleClick(location.x,location.y,duration=_duration)
        return True 

# 鼠标按下
def mouseDown(image):
    location=getLocation(image)
    if location is None:
        return False
    else:
        pyautogui.mouseDown(location.x,location.y,duration=_duration)
        return True 

# 鼠标松开
def mouseUp(image):
    location=getLocation(image)
    if location is None:
        return False
    else:
        pyautogui.mouseUp(location.x,location.y,duration=_duration)
        return True 

# 鼠标移动至
def mouseMoveTo(image):
    location=getLocation(image)
    if location is None:
        return False
    else:
        pyautogui.moveTo(location.x,location.y,duration=_duration)
        return True 

# 鼠标偏移
def mouseMoveRel(arg):
    try:
        offset = arg.split(",")
        x = int(offset[0])
        y = int(offset[1])
        pyautogui.moveRel(x,y, duration=_duration)
    except Exception as e:
        _logger.error("参数错误" + str(e))
    return True

# 鼠标右键单击
def mouseRightClick(image):
    location=getLocation(image)
    if location is None:
        return False
    else:
        pyautogui.rightClick(location.x,location.y,duration=_duration)
        return True 

# 鼠标滚轮
def mouseScroll(scrolling):
    pyautogui.scroll(int(scrolling),duration=_duration)
    return True 

# 输入文字
def inputText(text):
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl','v')
    return True

# 键盘输入
# 逗号隔开，组合件使用+号连接
# e.g.: a,b,c,left,enter,ctrl+c,win+r 
def keyboardInput(inputs):
    inputList = inputs.split(",")
    for input in inputList:
        keys = input.split("+")
        pyautogui.hotkey(*keys)
        time.sleep(_duration)
    return True

# 匹配图片
def assertImage(image):

    return (getLocation(image) != None)

# 读取文件
def readFile(file):
    try:
        fp = open(file, "r",encoding="utf-8")
        text = fp.read().splitlines() # 读取后以行进行分割
        fp.close()
        return text
    except IOError as error:
        _logger.error (f"读取文件 {file} 失败: {error}")
        return ""

# 对比文件
def compareFile(arg):
    files = arg.split(",")
    if len(files)!=2:
        _logger.error(f"对比文件参数错误: {arg}, 跳过对比")
        return True

    file1 = files[0]
    file2 = files[1]
    res = filecmp.cmp(file1, file2, shallow=True)
    return res

# 保存失败截图
def saveFailedImage(folder, expectedImage):
    os.makedirs(folder)
    #复制期望图片
    shutil.copy(expectedImage, f"{folder}\\expected.png")
    #保存当前屏幕
    pyautogui.screenshot(f"{folder}\\actual.png")
    return

# 保存失败文件对比结果
def saveHtmlDiff(folder, arg):

    files = arg.split(",")
    file1 = files[0]
    file2 = files[1]
    lines1 = readFile(file1)
    lines2 = readFile(file2)
    diff = difflib.HtmlDiff()
    htmlDiff = diff.make_file(lines1, lines2)

    try:
        os.makedirs(folder)
        with open(f"{folder}\\diff.html", 'w', encoding="utf-8") as result_file:
            result_file.write(htmlDiff)

    except IOError as error:
        _logger.error(f"写入html文件错误：{error}")
    return

# 保存失败结果
def saveFailedResult(folder, oper, arg):
    if oper == "对比文件":
        saveHtmlDiff(folder, arg)
    else:
        saveFailedImage(folder, arg)
    return

# 执行测试单步
def testStep(oper, arg, waitSec):
    curTry = 0
    succeed = False
    while curTry <= _retry and succeed == False:
        if curTry > 0:
            _logger.info(f"重试{curTry}/{_retry}")
        # 执行步骤
        if oper == "鼠标单击": 
            succeed = mouseClick(arg)
        elif oper == "鼠标双击":
            succeed = mouseDoubleClick(arg)
        elif oper == "鼠标按下": 
            succeed = mouseDown(arg)
        elif oper == "鼠标松开": 
            succeed = mouseUp(arg)
        elif oper == "鼠标移动至": 
            succeed = mouseMoveTo(arg)
        elif oper == "鼠标偏移": 
            succeed = mouseMoveRel(arg)
        elif oper == "鼠标右键单击":
            succeed = mouseRightClick(arg)
        elif oper == "鼠标滚轮": 
            succeed = mouseScroll(arg)
        elif oper == "输入文字":
            succeed = inputText(arg)
        elif oper == "键盘输入": 
            succeed = keyboardInput(arg)
        elif oper == "匹配图片": 
            succeed = assertImage(arg)
        elif oper == "对比文件": 
            succeed = compareFile(arg)
        else:
            _logger.error(f"未定义的操作类型 {oper}")
            succeed = True
        time.sleep(waitSec)
        curTry += 1

    return succeed

# 运行测试主程序
def RunTests(caseList = None):
    _logger.info("---------- 测试开始 ----------")
    # 读取配置
    workbook = xlrd.open_workbook(_testFile)
    settingsSheet = workbook.sheets()[0]
    _retry = settingsSheet.row(1)[1].value
    _precision = settingsSheet.row(1)[2].value
    _duration = settingsSheet.row(1)[3].value
    _logger.info(f">>> 读取配置: 重试次数={_retry}, 匹配精度={_precision}, 操作用时={_duration}s")

    errorCount = 0
    warningCount = 0
    # 运行测试用例
    cases = caseList if caseList != None else  range(1, workbook.nsheets)
    for iCase in cases:
        sheet = workbook.sheets()[iCase]
        _logger.info(f"----- 开始执行用例 {sheet.name} -----")

        # 依次执行测试步骤
        for i in range(1, sheet.nrows):
            # 获取参数
            oper = sheet.row(i)[0].value    #操作类型
            arg = sheet.row(i)[1].value     #参数
            waitSec = tryParseToFloat(sheet.row(i)[2].value)    #操作后等待时间(s)
            ignoredStep = True if sheet.row(i)[3].value != "" else False    #步骤失败可忽略
            
            # 执行
            _logger.info(f"step {i}: {oper} - {arg} - 等待{waitSec}s")
            if testStep(oper,arg,waitSec) == True:
                #成功, 记录完成
                _logger.info(f"step {i} 完成")
            else:
                #失败, 记录保存现场结果
                if ignoredStep == True:
                    #如果失败可忽略，保存为Warning
                    warningCount += 1
                    output = f"{_resultFolder}\\Warning_{warningCount}"
                    _logger.error(f"step {i} 失败, 忽略, 保存结果至 {output}")
                    saveFailedResult(output, oper, arg)
            
                else:
                    #如果失败不可忽略，保存为Error
                    errorCount += 1
                    output = f"{_resultFolder}\\Error_{errorCount}"
                    _logger.error(f"step {i} 失败, 中止当前测试用例, 保存结果至 {output}")
                    saveFailedResult(output, oper, arg)
                    break
        
        _logger.info(f"-----用例 {sheet.name} 执行完毕 -----\r\n")
    
    _logger.info(f"---------- 测试结束 ----------")
    if pyautogui.confirm(text=f"执行结果：\n  {errorCount} Errors  {warningCount} Warngings \n \n 结果已保存至 {_resultFolder} \n 查看结果?", title='执行完毕', buttons=['OK', 'Cancel']) == 'OK':
        os.startfile(_resultFolder) 

    return (errorCount, warningCount)

if __name__ == "__main__":
    if pyautogui.confirm(text=u"开始前请确认：\n  - 请以管理员身份运行\n  - 如果涉及键盘操作，请关闭输入法", title='开始', buttons=['OK', 'Cancel']) == 'OK':
        RunTests()
    