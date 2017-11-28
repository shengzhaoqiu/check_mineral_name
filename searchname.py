import xlrd
import re

# 从氧化物的字符串中返回主量元素
def mainele(oxide):
    return(re.findall(r'[a-zA-Z]+',oxide)[0])


infofile = xlrd.open_workbook('temp.xlsx')
datasheet = infofile.sheets()[0]
name = datasheet.col_values(0)
formula = datasheet.col_values(2)   #获取标准表化学成分占比列内容
# print(formula)
datafile = xlrd.open_workbook('data.xlsx')
resultsheet = datafile.sheets()[0]
analyzednametemp = resultsheet.row_values(0)     #获取分析结果中元素名称行内容
result = resultsheet.row_values(3)               #获取分析结果中元素占比内容
for x in range(4,len(analyzednametemp)):        #获取结果中第一位元素的字符串
    analyzednametemp[x] = mainele(analyzednametemp[x])



# 计算与标准表的差值
def getdif(ran_str,resu_str):
    if ran_str.isdigit():
        return abs(float(ran_str)-float(resu_str))
    if '-' in ran_str:
        rangelist = ran_str.split('-')
        rangemin = float(rangelist[0])
        rangemax = float(rangelist[1])
        if float(resu_str)>=rangemin and float(resu_str)<=rangemax:
            return 0
        if float(resu_str) < rangemin:
            return abs(rangemin-float(resu_str))
        if float(resu_str) > rangemax:
            return abs(rangemax-float(resu_str))
    if '<' in ran_str:
        if float(resu_str) <= float(ran_str[1:]):
            return 0
        else:
            return abs(float(resu_str) - float(ran_str[1:]))

# 计算与标准表的差值,从excel读取的数字是浮点型
def getdif_float(ran_str, resu):

    if resu == '' or resu == None:
        resu = 0

    if ran_str.isdigit():
        return abs(float(ran_str) - resu)
    if '-' in ran_str:
        rangelist = ran_str.split('-')
        rangemin = float(rangelist[0])
        rangemax = float(rangelist[1])
        if resu >= rangemin and resu <= rangemax:
            return 0
        if resu < rangemin:
            return abs(rangemin - resu)
        if resu > rangemax:
            return abs(rangemax - resu)
    if '.' in ran_str:
        return abs(float(ran_str) - resu)
    if '<' in ran_str:
        if resu <= float(ran_str[1:]):
            return 0
        else:
            return abs(resu - float(ran_str[1:]))


#48C
# aa = '48C/22na'
# 将检索表氧化物含量范围字符串分割
def wt2list(wt):
    alllist = []
    temp = wt.split('/')
    # print(len(temp))
    for x in range(0,len(temp)):
        index = re.search(r'(.*)([0-9])',temp[x]).span(0)
        # print(index)
        alllist.append(temp[x][0:index[1]])
        alllist.append(temp[x][index[1]:])
    # print(alllist)
        # print(index[1]
    return alllist



#构造元素名称和成分占比字典
alldict = dict()
for x in range(0,len(formula)):
    try:
        if formula[x] == '':
            alldict[name[x]] = formula[x]
            continue
        else:
            alldict[name[x]] =  wt2list(formula[x])
    except:
        print('Worng!')
        print(formula[x])
        continue
# print(alldict)
dif_count = dict()
for name_key in alldict:
    # print(name_key)
    tempcount = 0
    for i in range(1,len(alldict[name_key]),2):
        if alldict[name_key][i] in analyzednametemp:
            # print(alldict[name_key][i])
            result_name_index = analyzednametemp.index(alldict[name_key][i])
            tempcount += getdif_float(alldict[name_key][i-1],result[result_name_index])
    dif_count[name_key] = tempcount

print(dif_count)   #输出结果
print(sorted(dif_count.items(),key=lambda item:item[1]))  #将结果按差值大小降序排列
