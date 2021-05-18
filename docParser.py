from docx.shared import Inches
from docx import Document
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import re
import os


def make_rows_bold(*rows):
    for row in rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True


def color_row(row=0):
    'make row of cells background colored, defaults to column header row'
    row = table.rows[row]
    for cell in row.cells:
        shading_elm_2 = parse_xml(
            r'<w:shd {} w:fill="1F5C8B"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shading_elm_2)

################################################### 주의 사항 ##########################################################
# logic 처리 귀찮아서 class 밑부분의 public 시작하는 부분을 수동으로 라인넘버 할당해줘야함 ㅠㅠ 변수는 initialStartLine
################################################### 주의 사항 ##########################################################


# Config Variables
# fileName = 'UserService'
# fileName = 'TriggerActionService'
fileName = 'StatusService'
# fileName = 'ParkService'
# initialStartLine = 30
# initialStartLine = 20
initialStartLine = 6

if os.path.exists('./document/'+fileName+'.docx'):
    os.remove('./document/'+fileName+'.docx')
else:
    print('The file does not exist')

pubExpPattern = 'public [\w]+[\<\w\>]+ [\w]+[\(\w\s\,\)]+'
priExpPattern = 'private [\w]+[\<\w\>]+ [\w]+[\(][\w\s\,]+[\)]'

returnExpPattern = 'public [\w]+[\<\w\>]+'
funcNameExpPattern = '[\w]+[\(][\w\s\,]+[\)]'
paramExpPattern = '[\(][\w\s\,]+[\)]'
paramSmryExpPattern = '[\/]+ [\<]param.*'
paramSmryDelExpPattern = '[\"][\w]+[\"][\>]'
summaryExpPattern = '[\/]+ [a-zA-Z\s]+.*'

with open('./code/'+fileName+'.cs', encoding='utf8') as f:
    lines = f.read().split("\n")

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
        if dicName in dic:
            dicName += '_chk'
            dic[dicName] = {
                'Summary': '',
                'Returns': '',
                'table': [],
            }
        else:
            dic[dicName] = {
                'Summary': '',
                'Returns': '',
                'table': [],
            }
        dicNameArr.append(dicName)
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
    if i > initialStartLine:
        paramSummaryKeywordArr = re.findall(paramSmryExpPattern, line)
        if len(paramSummaryKeywordArr) > 0:
            filteredSummaryKeywordArr = paramSummaryKeywordArr[0].replace(
                '/// ', '').replace('<param name=', '').replace('</param>', '')
            filteredSummaryKeyword = re.findall(
                paramSmryDelExpPattern, filteredSummaryKeywordArr)
            paramSummaryArr.append(
                filteredSummaryKeywordArr.replace(
                    filteredSummaryKeyword[0], ''))

# Summary 값
for i, line in enumerate(lines):
    if i > initialStartLine:
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
            dic[dicNameArr[i]]['table'][j]['Param summary'] = paramSummaryArr[0]
            paramSummaryArr.pop(0)

# print(dic)


# word 문서 작성
doc = Document()

# 테이블 그리기
for name in dicNameArr:
    rowLength = len(dic[name]['table'])+4
    table = doc.add_table(rows=rowLength, cols=3, style='TableGrid')
    # 기본 테이블 설정
    table.cell(0, 1).merge(table.cell(0, 2))
    table.cell(1, 1).merge(table.cell(1, 2))
    table.cell(2, 1).merge(table.cell(2, 2))
    cell_1 = table.rows[0].cells
    cell_2 = table.rows[1].cells
    cell_3 = table.rows[2].cells
    cell_4 = table.rows[3].cells

    cell_1[0].text = "Name"
    cell_2[0].text = "Summary"
    cell_3[0].text = "Returns"
    cell_4[0].text = "Param Name"
    cell_4[1].text = "Type"
    cell_4[2].text = "Param Summary"
    make_rows_bold(table.rows[0], table.rows[1], table.rows[2], table.rows[3])

    # 값 입력
    cell_1[1].text = name
    cell_2[1].text = dic[name]['Summary']
    cell_3[1].text = dic[name]['Returns']
    for i in range(0, len(dic[name]['table'])):
        cell = table.rows[i+4].cells
        cell[0].text = dic[name]['table'][i]['Param Name']
        cell[1].text = dic[name]['table'][i]['Type']
        cell[2].text = dic[name]['table'][i]['Param summary']
        # print(dic[name]['table'][i])
    doc.add_paragraph(' ')


doc.save('./document/'+fileName+'.docx')
