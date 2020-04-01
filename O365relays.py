#!/usr/bin/python3
# original source from https://support.office.com/en-us/article/managing-office-365-endpoints-99cab9d4-ef59-4207-9f2b-3728eb46bf9a?ui=en-US&rs=en-US&ad=US#ID0EACAAA=4._Web_service
# modified by some dude

import json
import os
import urllib.request
import uuid
import smtplib
import csv



# helper to call the webservice and parse the response
def webApiGet(methodName, instanceName, clientRequestId):
    ws = "https://endpoints.office.com"
    requestPath = ws + '/' + methodName + '/' + instanceName + '?clientRequestId=' + clientRequestId
    request = urllib.request.Request(requestPath)
    with urllib.request.urlopen(request) as response:
        return json.loads(response.read().decode())

# questions/input/paths
print("Where would you like the client ID and latest version to be stored? Example: C:/users/john.doe/downloads/", end=' ')
datapath = input() + '/endpoints_clientid_latestversion.txt'

print("Where would you like the Office365 URL list to be stored at? Example: C:/users/john.doe/downloads/", end=' ')
url_list = input() + '/365_url_list.csv'

print("Where would you like the Office365 firewall ACL list to be stored at? Example: C:/users/john.doe/downloads/", end=' ')
ip_list = input() + '/365_ip_list.csv'

# fetch client ID and version if data exists; otherwise create new file
if os.path.exists(datapath):
    with open(datapath) as fin:
        clientRequestId = fin.readline().strip()
        latestVersion = fin.readline().strip()
else:
    clientRequestId = str(uuid.uuid4())
    latestVersion = '0000000000'
    with open(datapath, 'w') as fout:
        fout.write(clientRequestId + '\n' + latestVersion)

# call version method to check the latest version, and pull new data if version number is different
version = webApiGet('version', 'Worldwide', clientRequestId)
if version['latest'] > latestVersion:
    print('New version of Office 365 worldwide commercial service instance endpoints detected')

    # write the new version number to the data file
    with open(datapath, 'w') as fout:
        fout.write(clientRequestId + '\n' + version['latest'])

    # invoke endpoints method to get the new data
    endpointSets = webApiGet('endpoints', 'Worldwide', clientRequestId)

    # filter results for Allow and Optimize endpoints, and transform these into tuples with port and category
    flatUrls = []
    for endpointSet in endpointSets:
        if endpointSet['category'] in ('Optimize', 'Allow'):
            category = endpointSet['category']
            urls = endpointSet['urls'] if 'urls' in endpointSet else []
            tcpPorts = endpointSet['tcpPorts'] if 'tcpPorts' in endpointSet else ''
            udpPorts = endpointSet['udpPorts'] if 'udpPorts' in endpointSet else ''
            flatUrls.extend([(category, url, tcpPorts, udpPorts) for url in urls])

    flatIps = []
    for endpointSet in endpointSets:
        if endpointSet['category'] in ('Optimize', 'Allow'):
            ips = endpointSet['ips'] if 'ips' in endpointSet else []
            category = endpointSet['category']
            # IPv4 strings have dots while IPv6 strings have colons
            ip4s = [ip for ip in ips if '.' in ip]
            tcpPorts = endpointSet['tcpPorts'] if 'tcpPorts' in endpointSet else ''
            udpPorts = endpointSet['udpPorts'] if 'udpPorts' in endpointSet else ''
            flatIps.extend([(category, ip, tcpPorts, udpPorts) for ip in ip4s])

    print('IPv4 Firewall IP Address Ranges')
    print(','.join(sorted(set([ip for (category, ip, tcpPorts, udpPorts) in flatIps]))))
    # outputs the details to the ip_list variable
    with open(ip_list, 'w') as data_write:
        data_write.write(','.join(sorted(set([ip for (category, ip, tcpPorts, udpPorts) in flatIps]))))

    print('URLs for Proxy Server')
    print(','.join(sorted(set([url for (category, url, tcpPorts, udpPorts) in flatUrls]))))
    # outputs the details to the url_list variable
    with open(url_list, 'w') as data_write:
        data_write.write(','.join(sorted(set([ip for (category, ip, tcpPorts, udpPorts) in flatUrls]))))

  
    print('IPv4 Firewall IP Address Ranges')
    print(','.join(sorted(set([ip for (category, ip, tcpPorts, udpPorts) in flatIps]))))
    # outputs the details to the ip_list variable
    with open(ip_list, 'w') as data_write:
        data_write.write(','.join(sorted(set([ip for (category, ip, tcpPorts, udpPorts) in flatIps]))))

else:
    print('Office 365 worldwide commercial service, no updates detected.')

# csv file name 
filename = url_list
  
# initializing the titles and rows list 
fields = [] 
rows = [] 
  
# reading csv file 
with open(filename, 'r') as csvfile: 
    # creating a csv reader object 
    csvreader = csv.reader(csvfile) 
      
    # extracting each data row one by one 
    for row in csvreader: 
        rows.append(row) 
  
    # get total number of rows 
    print("Total no. of rows: %d"%(csvreader.line_num)) 
  

  
# printing first 5 rows 
print('\nThe URL list members are:\n') 
for row in rows[:5]: 
    # parsing each column of a row 
    for col in row: 
        print("%10s"%col), 
    print('\n') 

# csv file name for firewall IPv4
filename2 = ip_list

# initializing the titles and rows list
fields = []
rows = []

# reading csv file of IPv4 addresses
with open(filename2, 'r') as csvfile:
    # creating a csv reader object
    csvreader = csv.reader(csvfile)

    # extracting each data row one by one
    for row in csvreader:
        rows.append(row)

    # get total number of rows
    print("Total no. of rows: %d"%(csvreader.line_num))

# printing the rows
print('\nThe Firewall IP list is:\n')
for row in rows[:5]:
    # parsing each column of a row
    for col in row:
        print("%10s"%col),
    print('\n')


  
