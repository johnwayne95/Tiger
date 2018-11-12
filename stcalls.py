import requests
import json
import datetime
import os
import sys
# test
minimumtime = datetime.time(0, 1, 0, 0)

today = datetime.datetime.today()

thirtydays = today - datetime.timedelta(days=3)

callsfile = os.path.join(sys.path[0], 'callsdump.json')

headers = {"Accept": "application/json", "serviceTitanApiKey" : "8ae6480a-2499-4e84-bd09-7c461f033ddb", "createdAfter" : thirtydays.strftime("%Y-%m-%d"), "pageSize" : "20000"}
resp = requests.get('HTTPS://api.servicetitan.com/v1/calls?', params = headers)

with open(callsfile, "w") as write_file:
    json.dump(resp.json(), write_file, indent=4)


with open(callsfile, "r") as read_file:
    data = json.load(read_file)


calls = data['data']

for x in range(len(calls)):
    calls[x] = dict(calls[x])
    calls[x]['leadCall'] = dict(calls[x]['leadCall'])
    calltime = datetime.datetime.strptime(calls[x]['leadCall']['duration'][:8], '%H:%M:%S').time()
    status = calls[x]['leadCall']['callType']
    callid = calls[x]['leadCall']['id']
    if(calltime <= minimumtime and status == "Abandoned"):

        calls[x]['leadCall']['callType'] = 'Excused'
        calls[x]['leadCall']['reason'] = None
        calls[x]['leadCall']['recordingUrl'] = None

        #print(calls[x]['leadCall'])
        url = 'HTTPS://api.servicetitan.com/v1/calls/' + str(calls[x]['leadCall']['id'])
        header = {'Accept' : 'application/json', 'serviceTitanApiKey' : '8ae6480a-2499-4e84-bd09-7c461f033ddb'}

        r = requests.put(url, params = header, json=calls[x]['leadCall'])

        print(r.text)