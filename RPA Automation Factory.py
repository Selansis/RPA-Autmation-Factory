import requests
dls = "https://rpachallenge.com/assets/downloadFiles/challenge.xlsx"
resp = requests.get(dls)

output = open('test.xlsx', 'wb')
output.write(resp.content)
output.close()