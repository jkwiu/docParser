from docx.shared import Inches
from docx import Document
import re

# 나와야 하는 것
yes1 = 'public LoginResult InsertLoginPROEUser(string username, string newUserPwd)'
yes2 = 'public LoginResult InsertLoginPROEUser(string username)'
yes3 = 'public ResultEntry<MethodResult> InsertUser(string userInfo, string user)'
yes4 = 'public ResultEntry InsertUser(string userInfo, string user)'
yes5 = 'public ResultEntry InsertUser(string userInfo)'
yes6 = 'public ResultEntry<MethodResult> InsertUser(string userInfo)'
yes7 = 'private ResultEntry<MethodResult> CreateUser(UserInfo user)'
yes8 = 'private ResultEntry<MethodResult> CreateUser(UserInfo user)'
yes9 = 'private string returnMethodForm(string methodName, string interfaceName)'
yes10 = 'public string returnMethodForm(string methodName, string interfaceName)'
yes11 = 'public string InsertUser(string userInfo)'
yes12 = 'private string InsertUser(string userInfo)'
yes13 = 'public class UserService : SecurityContextBase'
showMe = [
    yes1,
    yes2,
    yes3,
    yes4,
    yes5,
    yes6,
    yes7,
    yes8,
    yes9,
    yes10,
    yes11,
    yes12,
    yes13,
]
# 안나와야 하는 것
no1 = 'public class UserService : SecurityContextBase'
no2 = 'namespace Park.WebService.ServiceContractImplementation'
no3 = 'using W2B.Core.Interfaces;'
no4 = 'public UserService(string userKey, string user)'
no5 = 'public UserService(string userKey, string user, int num)'
no6 = 'public UserService(string userKey)'
no7 = '/// <summary>'
noShowMe = [
    no1,
    no2,
    no3,
    no4,
    no5,
    no6,
    no7,
]

# fileName = 'UserService'
fileName = 'ParkService'

pubExpPattern = 'public [\w]+[\<\w\>]+ [\w]+[\(][\w\s\,]+[\)]'
priExpPattern = 'private [\w]+[\<\w\>]+ [\w]+[\(][\w\s\,]+[\)]'

returnExpPattern = 'public [\w]+[\<\w\>]+'
funcNameExpPattern = '[\w]+[\(][\w\s\,]+[\)]'
paramExpPattern = '[\(][\w\s\,]+[\)]'
paramSmryExpPattern = '[\/]+ [\<]param.*'
paramSmryDelExpPattern = '[\"][\w]+[\"][\>]'
summaryExpPattern = '[\/]+ [a-zA-Z\s]+.*'

# 테스트 코드
# print('--------------나와야 함------------------------')
# for i, sm in enumerate(showMe):
#     print(i+1, ': ', re.findall(pubExpPattern, sm))
#     # print(i+1, ': ', re.findall(priExpPattern, sm))
# print('--------------안 나와야 함------------------------')
# for i, nsm in enumerate(noShowMe):
#     print(i+1, ': ', re.findall(pubExpPattern, nsm))

with open(fileName+'.cs', encoding='utf8') as f:
    lines = f.read().split("\n")
# print("Number of lines is {}".format(len(lines)))


# Dictionary

dicName = ''
dicNameArr = []
summaryArr = []
paramSummaryArr = []

isFunction = False
dic = {
    # 'Name': {
    #     'Summary': '',
    #     'Returns': '',
    #     'table': [],
    # }
}
dicTbl = {
    # 'Param Name': '',
    # 'Type': '',
    # 'Param summary': '',
}

for i, line in enumerate(lines):

  # 1. 핵심 키워드를 추출한다.
  # 2. 키워드에서 Returns, Parameter, function  name을 나눈고, 해당 키워드들을 dictionary에 넣는다.
  # 3. dictionary로부터 word 문서 테이블을 만든다.
    keyword = re.findall(pubExpPattern, line)
    if len(keyword) > 0:
        isFunction = True
        # Function name 값
        funcNameKeywordArr = re.findall(funcNameExpPattern, keyword[0])
        isGgwal = True
        filteredFuncNameKeyword = ''
        for k in funcNameKeywordArr[0]:
            if k == '(':
                isGgwal = False
                continue
            if isGgwal:
                filteredFuncNameKeyword += k
        dicName = filteredFuncNameKeyword
        dicNameArr.append(dicName)
        dic[dicName] = {
            'Summary': '',
            'Returns': '',
            'table': [],
        }
        # Returns 값
        returnKeywordArr = re.findall(returnExpPattern, keyword[0])
        returnKeyword = returnKeywordArr[0].split(' ')
        # print(returnKeyword[1])
        isGguk = False
        noGguk = True
        filteredReturnKeyword = ''
        for k in returnKeyword[1]:
            if noGguk:
                filteredReturnKeyword += k
            if k == '<':
                isGguk = True
                noGguk = False
                filteredReturnKeyword = ''
                continue
            elif k == '>':
                isGguk = False
                continue
            if isGguk:
                filteredReturnKeyword += k
        dic[dicName]['Returns'] = filteredReturnKeyword
        # Parameter 값
        paramKeywordArr = re.findall(paramExpPattern, keyword[0])
        filteredParamKeywordArr = paramKeywordArr[0].replace(
            '(', '').replace(')', '').split(',')
        # print(filteredParamKeywordArr)
        for i in range(0, len(filteredParamKeywordArr)):
            dicTbl = {
                'Param Name': filteredParamKeywordArr[i].lstrip().split(' ')[1],
                'Type': filteredParamKeywordArr[i].lstrip().split(' ')[0],
                'Param summary': '',
            }
            dic[dicName]['table'].append(dicTbl)


# Parameter Summary 값
for i, line in enumerate(lines):
    if i > 30:
        paramSummaryKeywordArr = re.findall(paramSmryExpPattern, line)
        if len(paramSummaryKeywordArr) > 0:
            filteredSummaryKeywordArr = paramSummaryKeywordArr[0].replace(
                '/// ', '').replace('<param name=', '').replace('</param>', '')
            filtering = re.findall(
                paramSmryDelExpPattern, filteredSummaryKeywordArr)
            paramSummaryArr.append(
                filteredSummaryKeywordArr.replace(filtering[0], ''))

# Summary 값
for i, line in enumerate(lines):
    if i > 30:
        summaryKeyword = re.findall(summaryExpPattern, line)
        if len(summaryKeyword) > 0:
            filteredSummaryKeyword = (summaryKeyword[0].replace(
                '///  ', '')).replace('/// ', '')
            summaryArr.append(filteredSummaryKeyword)
for i in range(0, len(dicNameArr)):
    dic[dicNameArr[i]]['Summary'] = summaryArr[i]
for i in range(0, len(dicNameArr)):
    # dic[dicNameArr[i]]['table'] = paramSummaryArr[i]
    for j in range(0, len(dic[dicNameArr[i]]['table'])):
        if len(paramSummaryArr) > 0:
            # dic[dicNameArr[i]]['table'][j]['Param summary'] = paramSummaryArr[j]
            paramSummaryArr.pop(0)

print(dic)
# print(len(paramSummaryArr))
# print(paramSummaryArr)
# for i in enumerate(dic):
#     print(i)
# print(dic)


# document = Document()

# document.add_heading(fileName, 0)
# table = document.add_table(rows=4, cols=2)
# hdr_cells0 = table.rows[0].cells
# hdr_cells1 = table.rows[1].cells
# hdr_cells2 = table.rows[2].cells
# hdr_cells0[0].text = 'Name'
# hdr_cells1[0].text = 'Summary'
# hdr_cells2[0].text = 'Returns'

# t_sum_a = table.cell(3, 0)
# t_sum_b = table.cell(3, 1)
# t_sum_a.merge(t_sum_b)

# document.save(fileName+'.docx')
