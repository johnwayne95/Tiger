import requests
import json
import datetime

headers = {"Accept": "application/json", "serviceTitanApiKey" : "8ae6480a-2499-4e84-bd09-7c461f033ddb", "startsAfter" : "2018-11-11", "pageSize" : "20000"}
resp = requests.get('HTTPS://api.servicetitan.com/v1/jobs?', params = headers)

with open("jobsdump.json", "w") as write_file:
    json.dump(resp.json(), write_file)


with open("jobsdump.json", "r") as read_file:
    data = json.load(read_file)


jobs = data['data']
total = 0.0
print(type(jobs))
for x in range(len(jobs)):
    jobs[x] = dict(jobs[x])
    jobs[x]['businessUnit'] = dict(jobs[x]['businessUnit'])

    if(jobs[x]['invoice'] != None):
        jobs[x]['invoice'] = dict(jobs[x]['invoice'])

        
        testdate = datetime.date(2018, 11, 11)
        jobstart = datetime.datetime.strptime(jobs[x]['start'][:19], "%Y-%m-%dT%H:%M:%S").date()
        # if(jobstart >= testdate):
        print (jobs[x]['businessUnit']['name'])
        # print (jobs[x]['businessUnit']['id'])
        #print(jobs[x]['invoice']['total'])
        #total += jobs[x]['invoice']['total']
        for y in range(len(jobs[x]['tags'])):
            jobs[x]['tags'][y] = dict(jobs[x]['tags'][y])
            print(jobs[x]['tags'][y]['name'])
        #print(total)