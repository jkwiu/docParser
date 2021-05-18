import re


txt = 'public ResultEntry<MethodResult> IsParkAlive()'
txt2 = 'public ProcessedResultSet<UserInfo> GetUserList(int pageNumber, int pageSize, string filter, string order, int start, int limit)'
txt3 = 'public LoginResult InsertLoginPROEUser(string username, string newUserPwd)'

pubExpPattern = 'public [\w]+[\<\w\>]+ [\w]+[\(\w\s\,\)]+'

filteredTxt = re.findall(pubExpPattern, txt3)
print(filteredTxt)
