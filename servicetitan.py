import requests
import json

#headers = {"Accept": "application/json", "serviceTitanApiKey" : "8ae6480a-2499-4e84-bd09-7c461f033ddb", "createdAfter" : "2018-01-01", "pageSize" : "20000"}
#resp = requests.get('HTTPS://api.servicetitan.com/v1/jobs?', params = headers)

#with open("jobsdump.json", "w") as write_file:
    #json.dump(resp.json(), write_file)


with open("jobsdump.json", "r") as read_file:
    data = json.load(read_file)


jobs = data['data']

print(type(jobs))
for x in range(len(jobs)):
    jobs[x] = dict(jobs[x])
    jobs[x]['businessUnit'] = dict(jobs[x]['businessUnit'])

    # for y in range(len(jobs[x]['estimates'])):
    #     jobs[x]['estimates'][y] = dict(jobs[x]['estimates'][y])
        
    #     for z in range(len(jobs[x]['estimates'][y]['items'])):
    #         jobs[x]['estimates'][y]['items'][z] = dict(jobs[x]['estimates'][y]['items'][z])
    #         print(jobs[x]['estimates'][y]['items'][z]['total'])

    
    print(jobs[x]['start'])
    print (jobs[x]['businessUnit']['name'])