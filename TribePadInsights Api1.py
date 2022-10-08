#!/usr/bin/env python3

import requests
import pandas as pd

tokenUrl = "https://insights.tribepad.com/api/oauth/Sodexo/access_token"
tokenParams = {
    'client_id': 'insights1131105',
    'client_secret': '',
    'grant_type': 'client_credentials',
    'scope': 'insights_api'
}
response = requests.post(tokenUrl, data=tokenParams)
r = response.json()
token = "Bearer " + r["access_token"]
apiHeader = {'Authorization': token}

reportsURL = "https://insights.tribepad.com/api/oauth/Sodexo/report/list"
reports = requests.post(reportsURL, headers=apiHeader)
output = reports.json()

# reportList = pd.DataFrame.from_dict(output, orient='columns')
# reportList.to_excel("Tribepad Report List.xlsx")
# print(reportList)

baseURL = "https://insights.tribepad.com/api/oauth/Sodexo/report/"
reportID = input("What is the number of the report you want to retrieve? ")
getReportURL = baseURL + str(reportID) + "/retrieve"
outputName = "Tribepad Insights Report " + str(reportID) + ".xlsx"

result = requests.post(getReportURL, headers=apiHeader).json()

columnHeadersRaw = result['column_labels']
columnHeaders = list(columnHeadersRaw.values())
resultData = result['data']

reportDF = pd.DataFrame(resultData)
reportDF.columns = columnHeaders
df = reportDF.set_index('Permission Id')
# df.drop(df.columns[[4,5,7,9,11,13,15,17,19,21,23,25,27,29,31,33,35,37,40,41]],axis=1, inplace=True)
df.drop(["Permission System",
        "[ROLE] Job Seeker (Non Employee) [ID]",
         "[ROLE] Job Seeker (Employee) [ID]",
         "[ROLE] Resourcing Partner [ID]",
         "[ROLE] HR User (Justice Services) [ID]",
         "[ROLE] HR User [ID]",
         "[ROLE] Unbanded hiring manager [ID]",
         "[ROLE] Agency [ID]",
         "[ROLE] Sourcer [ID]",
         "[ROLE] Partnerships [ID]",
         "[ROLE] Super user [ID]",
         "[ROLE] Sports and leisure admin [ID]",
         "[ROLE] Sports & Leisure Hiring Manager [ID]",
         "[ROLE] Frontline Hiring Manager [ID]",
         "[ROLE] Hiring Manager New [ID]",
         "[ROLE] Recruiter New [ID]",
         "[ROLE] Unbanded HM Test [ID]",
         "Feature Id",
         "Feature Group Id"], axis=1, inplace=True)
df.to_excel(outputName)
print(df)
