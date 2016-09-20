import datetime
import requests
import re
import os
import xlsxwriter
import prettytable

#代理
HEADERS = {"User-Agent":r"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36"}

#URL
MAIN_URL              = r"http://fund.eastmoney.com/%s.html"  #主页
INDUSTRY_URL          = r"http://fund.eastmoney.com/f10/F10DataApi.aspx?type=hypz&code=%s&year=2016"  #行业配置
fundURL               = ""   #基金的URL
industryURL           = ""   #行业配置URL
moreManagerURLList    = []   #更多基金经理信息URL列表

#基金信息
fundCode              = ""   #代码
fundName              = ""   #名字
fundType              = ""   #类型
fundRisk              = ""   #风险
fundNet               = ""   #净值
fundAge               = ""   #成立时间
fundSize              = ""   #规模
rankDays              = []   #评级日期列表(海通证券，招商证券，上海证券，济安金信)
rankList              = []   #评级列表(海通证券，招商证券，上海证券，济安金信)
incomeList            = []   #收益列表(近1周，近1月，近3月，近6月，今年来，近1年，近2年，近3年)

#当前行业配置
industryList          = []   #行业类别列表
industryPercentList   = []   #占净值比例列表
industryValueList     = []   #市值(万元)列表

#基金经理
managerNameList       = []   #基金经理名字列表
managerTimeList       = []   #管理当前基金的时间列表
managerCareer         = ""   #累积的任职时间
managerCareerList     = []   #累积的任职时间列表
managerCountFund      = ""   #单个经理同时管理的基金数
managerCountFundList  = []   #所有经理同时管理的基金数列表
managerCurIncome      = ""   #当前基金总收益
managerCurIncomeList  = []   #所有经理基金总收益

#基金经理同时管理的基金信息
managerFundCode       = []   #单个经理管理的基金代码列表
managerFundCodeList   = []   #所有经理管理的基金代码列表
managerFundName       = []   #单个经理管理的基金名称列表
managerFundNameList   = []   #所有经理管理的基金名称列表
managerFundType       = []   #单个经理管理的基金类型列表
managerFundTypeList   = []   #所有经理管理的基金类型列表
managerFundTime       = []   #单个经理任职时间列表
managerFundTimeList   = []   #所有经理任职时间列表
managerFundDay        = []   #单个经理任职天数列表
managerFundDayList    = []   #所有经理任职天数列表
managerFundIncome     = []   #单个经理任职回报列表
managerFundIncomeList = []   #所有经理任职回报列表

#当前时间
starttime = datetime.datetime.now()

if __name__ == "__main__":

    fundCode = input("请输入需要爬取的基金代码：")
    fundURL = MAIN_URL % fundCode

    print("********开始爬虫********")
    fundHTML = requests.get(fundURL, headers = HEADERS)
    fundHTML.encoding = "utf-8"

    #基本信息
    fundName = re.search("FundName\">(.*?)</span>", fundHTML.text, re.S).group(1)
    fundType = re.search("ft_;pt_\d+\">(.*?)</a>", fundHTML.text, re.S).group(1)
    fundRisk = re.search("\|&nbsp;&nbsp;(.*?)</td><td>", fundHTML.text, re.S).group(1)
    fundNet = re.search("gz_gsz\">(.*?)</span>", fundHTML.text, re.S).group(1)
    fundAge = re.search("成 立 日</span>：(.*?)</td>", fundHTML.text, re.S).group(1)
    fundSize = re.search("基金规模</a>：(.*?)</td>", fundHTML.text, re.S).group(1)
    # 截取评级
    rankContent = re.search("html\">海通证券</a>(.*?)更多评级信息></a>", fundHTML.text, re.S).group(1)
    rankDays = re.findall("alignRight\">(.*?)</td>", rankContent, re.S)
    rankList = re.findall("alignRight10\">(.*?)</td>", rankContent, re.S)
    # 截取阶段涨幅
    increaseContent = re.search("typeName\">同类排名(.*?)四分位排名<div class=\"infoTips\">", fundHTML.text, re.S).group(1)
    incomeList = re.findall("Rdata\">(.*?)</div>", increaseContent, re.S)

    #行业配置
    industryURL = INDUSTRY_URL % fundCode
    industryHTML = requests.get(industryURL)
    industryContent = re.search("&nbsp;&nbsp;(.*?)&nbsp;&nbsp;", industryHTML.text, re.S).group(1) #截取最新季度信息
    industryList = re.findall("class='tol'>(.*?)</td>", industryContent, re.S)
    industryPercentList = re.findall("class='tor'>(.*?)</td><td class='tor'>", industryContent, re.S)
    industryValueList = re.findall("class='tor'>.*?</td><td class='tor'>(.*?)</td>", industryContent, re.S)

    #获取当前基金的所有基金经理
    moreManagerContent = re.search("基金经理变动一览</a>(.*?)更多", fundHTML.text, re.S).group(1)
    moreManagerURL = re.search("href=\"(.*?)\"", moreManagerContent, re.S).group(1)
    moreManagerHTML = requests.get(moreManagerURL, headers = HEADERS)
    allmanagerContent = re.search("现任基金经理简介(.*?)正文部份结束", moreManagerHTML.text, re.S).group(1) #截取所有基金经理信息
    managerNameList = re.findall("姓名：</strong><.*?>(.*?)</a></p><p>", allmanagerContent, re.S)
    moreManagerURLList = re.findall("text-decoration:none;' href=\"(.*?)\"", allmanagerContent, re.S)
    managerTimeList = re.findall("上任日期：</strong>(.*?)</p><p>", moreManagerHTML.text, re.S)

    #获取当前任职基金经理个人信息
    for i in range(len(managerNameList)):
        sigleManagerHTML = requests.get(moreManagerURLList[i], headers = HEADERS)
        sigleManagerHTML.encoding = "utf-8"
        #累计任职时间
        managerCareer = re.search("累计任职时间：</span>(.*?)<br />", sigleManagerHTML.text, re.S).group(1)
        managerCareerList.append(managerCareer)
        #同时管理的基金数
        managerCountFund = len(re.findall("name:'(.*?)'", sigleManagerHTML.text, re.S))
        managerCountFundList.append(managerCountFund)
        managerFundContent = re.search("任职回报</th>(.*?)</tbody>", sigleManagerHTML.text, re.S).group(1) #截取任职回报table
        #基金代码
        managerFundCode = re.findall(".html\">([0-9].*?)</a>", managerFundContent, re.S)
        managerFundCodeList.append(managerFundCode)
        #基金名称
        managerFundName = re.findall("tdl\">.*?>(.*?)</a>", managerFundContent, re.S)
        managerFundNameList.append(managerFundName)
        #基金类型
        managerFundType = re.findall("档案</a></td><td>(.*?)</td>", managerFundContent, re.S)
        managerFundTypeList.append(managerFundType)
        #任职时间
        managerFundTime = re.findall("档案</a></td><td>.*?</td><td>.*?</td><td>(.*?)</td><td>", managerFundContent, re.S)
        managerFundTimeList.append(managerFundTime)
        #任职天数
        managerFundDay = re.findall("~.*?</td><td>(.*?)</td>", managerFundContent, re.S)
        managerFundDayList.append(managerFundDay)
        #任职回报
        managerFundIncome = re.findall("~.*?</td><td>.*?天</td><td class=\".*?\">(.*?)</td>", managerFundContent, re.S)
        managerFundIncomeList.append(managerFundIncome)
        #当前基金总收益
        for i in range(len(managerFundCode)):
            if managerFundCode[i] == fundCode:
                managerCurIncome = managerFundIncome[i]
                managerCurIncomeList.append(managerCurIncome)
                break

    #输出爬取的信息
    print("")
    print("基金信息")
    print("基金代码：%s" % fundCode)
    print("基金名字：%s" % fundName)
    print("基金类型：%s" % fundType)
    print("基金风险：%s" % fundRisk)
    print("基金净值：%s" % fundNet)
    print("基金成立时间：%s" % fundAge)
    print("基金规模：%s" % fundSize)
    print("")
    print("行业配置")
    pt = prettytable.PrettyTable(["序号", "行业类别", "占净值比例", "市值（万元）"])
    pt.padding_width = 5
    for i in range(len(industryList)):
        pt.add_row([(i + 1), industryList[i], industryPercentList[i], industryValueList[i]])
    print(pt)
    print("")
    print("基金评级")
    pt = prettytable.PrettyTable(["评级机构", "评级日期", "评级"])
    pt.padding_width = 5
    for i in range(len(rankList)):
        pt.add_row([rankList[i], rankDays[i], rankList[i]])
    print(pt)
    print("")
    print("基金阶段涨幅")
    pt = prettytable.PrettyTable(["近1周", "近1月", "近3月", "近6月", "今年来", "近1年", "近2年", "近3年"])
    pt.padding_width = 3
    pt.add_row(incomeList)
    print(pt)
    print("")
    print("%d位基金经理" % len(managerNameList))
    for i in range(len(managerNameList)):
        print("第%d位：%s" % ((i + 1), managerNameList[i]))
        print("管理当前基金时间：%s" % managerTimeList[i])
        print("累计任职时间：%s" % managerCareerList[i])
        print("当前基金总收益：%s" % managerCurIncomeList[i])
        print("同时在管理的基金数：%s" % managerCountFundList[i])

        pt = prettytable.PrettyTable(["基金代码", "基金名称", "基金类型", "任职时间", "任职天数", "任职回报"])
        pt.padding_width = 5
        for j in range(len(managerFundCodeList[i])):
            pt.add_row([managerFundCodeList[i][j],
                        managerFundNameList[i][j],
                        managerFundTypeList[i][j],
                        managerFundTimeList[i][j],
                        managerFundDayList[i][j],
                        managerFundIncomeList[i][j]])
        print(pt)
    print("")

    file_name = input("抓取完成，输入文件名保存(不输入则保存到脚本路径)：")
    if file_name == "":
        curTime = str(datetime.datetime.now())
        curTime = re.sub(":|\.", "", curTime)
        file_name = "%s %s" % (fundName, curTime)
    savePath = os.getcwd()
    workbook = xlsxwriter.Workbook(savePath + "\\%s.xlsx" % file_name)
    print("保存到：" + savePath + "\\%s.xlsx" % file_name)

    #第一页sheet
    worksheet1 = workbook.add_worksheet("基金信息")

    titleFormat = workbook.add_format()
    titleFormat.set_bold()
    titleFormat.set_bg_color("orange")
    titleFormat.set_font_size(12)
    titleFormat.set_align("center")
    titleFormat.set_align("vcenter")
    titleFormat.set_border(1)

    contentFormat = workbook.add_format()
    contentFormat.set_bg_color("yellow")
    contentFormat.set_align("center")
    contentFormat.set_align("vcenter")

    headList = ["基金代码",
                "基金名字",
                "基金类型",
                "基金风险",
                "基金净值",
                "成立时间",
                "基金规模",
                "海通证券评级",
                "招商证券评级",
                "上海证券",
                "济安金信",
                "近1周涨幅",
                "近1月涨幅",
                "近3月涨幅",
                "近6月涨幅",
                "今年来涨幅",
                "近1年涨幅",
                "近2年涨幅",
                "近3年涨幅"]
    contentList = [fundCode,
                   fundName,
                   fundType,
                   fundRisk,
                   fundNet,
                   fundAge,
                   fundSize,
                   rankList[0],
                   rankList[1],
                   rankList[2],
                   rankList[3],
                   incomeList[0],
                   incomeList[1],
                   incomeList[2],
                   incomeList[3],
                   incomeList[4],
                   incomeList[5],
                   incomeList[6],
                   incomeList[7]]
    for i in range(len(headList)):
        worksheet1.write(0, i, headList[i], titleFormat)
        worksheet1.write(1, i, contentList[i], contentFormat)

    #第二页sheet
    worksheet2 = workbook.add_worksheet("行业配置")
    worksheet2.write(0, 0, "行业", titleFormat)
    worksheet2.write(0, 1, "比例", titleFormat)
    worksheet2.write(0, 2, "市值（万元）", titleFormat)
    for i in range(len(industryList)):
        worksheet2.write(i + 1, 0, industryList[i], contentFormat)
        worksheet2.write(i + 1, 1, industryPercentList[i], contentFormat)
        worksheet2.write(i + 1, 2, industryValueList[i], contentFormat)

    #后续sheet
    for i in range(len(managerNameList)):

        #基金经理基本信息
        headList = ["姓名", "管理当前基金起始时间", "当前基金的收益", "累积的任职时间", "同时管理的基金数"]
        contentList = [managerNameList[i],
                       managerTimeList[i],
                       managerCurIncomeList[i],
                       managerCareerList[i],
                       managerCountFundList[i]]
        worksheet = workbook.add_worksheet("第%d位基金经理" % (i + 1))
        for j in range(len(headList)):
            worksheet.write(0, j, headList[j], titleFormat)
            worksheet.write(1, j, contentList[j], contentFormat)

        #当前管理的基金信息
        headList = ["基金代码", "基金名称", "基金类型", "任职时间", "任职天数", "任职回报"]
        worksheet.write_row("A5", headList, titleFormat)
        for k in range(len(managerFundCodeList[i])):
            contentList = [managerFundCodeList[i][k],
                           managerFundNameList[i][k],
                           managerFundTypeList[i][k],
                           managerFundTimeList[i][k],
                           managerFundDayList[i][k],
                           managerFundIncomeList[i][k]]
            if managerFundCodeList[i][k] == fundCode:
                newContentFormat = workbook.add_format()
                newContentFormat.set_bg_color("red")
                newContentFormat.set_align("center")
                newContentFormat.set_align("vcenter")
                worksheet.write_row("A%d" % (6 + k), contentList, newContentFormat)
            else:
                worksheet.write_row("A%d" % (6 + k), contentList, contentFormat)

    workbook.close()

    #爬虫消耗的时间
    endtime = datetime.datetime.now()
    time = (endtime - starttime).seconds
    print("")
    print("********结束爬虫********")
    print('总耗时：%ss' % time)

