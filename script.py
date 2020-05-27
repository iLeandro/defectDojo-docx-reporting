import mysql.connector
from mysql.connector import Error
from docxtpl import DocxTemplate, Listing
import glob, os
import json
import datetime
import itertools as IT
import sys
os.chdir("/home/dojo/templates")

try:
   mydb = mysql.connector.connect(
      host="172.18.0.2",
      user="defectdojo",
      passwd="defectdojo",
      database="defectdojo",
      port="3306"
   )
except Error as dbConnectionError:
    print("Error", dbConnectionError)
    sys.exit()

#finally:
#if (mydb.is_connected()):
#mydb.close()
#cursor.close()
#print("MySQL connection is closed")   


client_name_input = input("Client Name: ")

try:
   cursor = mydb.cursor()
   #Get ID from the given name
   sql_select_client_name_query = ("SELECT id FROM dojo_product WHERE name = %s")
   cursor.execute(sql_select_client_name_query, (client_name_input, ))
   records_client_name = cursor.fetchall()

except Error as clientNameQueryError:
    print("Error", clientNameQueryError)
    sys.exit()

#If row result is greater than 0, then the Client exists. (Here we can change it to be equal to 1 instead. DDJ doesn't let be equal names. lowercase validated)
if cursor.rowcount > 0:
   print(client_name_input + " found!")
else:
   print("Client doesn't exist. Exiting...")
   sys.exit()

#Get Product (client) ID
client_id = 0
for row in records_client_name:
   client_id = row[0]


try:
   #Get Engagements associated with the given Client Name (Client ID).
   sql_select_engagements_query = ("SELECT id,name,product_id,target_start,target_end  FROM dojo_engagement WHERE product_id = %s")
   cursor.execute(sql_select_engagements_query, (client_id, ))
   records_engagements = cursor.fetchall()

except Error as engagementsQueryError:
    print("Error", engagementsQueryError)
    sys.exit()


engagement_option = 0
if cursor.rowcount == 0:
   print("\n")
   print("This client doesn't have engagements.")
   sys.exit()
elif cursor.rowcount == 1:
   print("\n")
   print("Found {} engagement. Do you want to use it?".format(cursor.rowcount))
   
   engagement_length = cursor.rowcount
   engagement_array = [None] * engagement_length
   engagement_array[0] = 0
   number_engagement = 0
   for row in records_engagements:
      print(number_engagement, " - " ,row[1],)
      engagement_array[number_engagement] = row[0]

   print("\n")
   engagement_option = int(input("Select Option. No(0), Yes(1): "))
   
   helper_engagement = 0 #control
   helper_number_engagement = 0 #row
   for engagement in engagement_array:
      if engagement_option == helper_number_engagement:
         helper_engagement = engagement
         helper_number_engagement += 1
      else:
         helper_number_engagement += 1



print("\n")
print("Found {} engagements. Which engangement do you want to use?".format(cursor.rowcount))

engagement_length = cursor.rowcount + 1
engagement_array = [None] * engagement_length
engagement_array[0] = 0
print("0  -  All Engagements")
number_engagement = 1
for row in records_engagements:
   print(number_engagement, " - " ,row[1],)
   engagement_array[number_engagement] = row[0]
   number_engagement += 1


engagement_option = int(input("Select Option (0-" + str(number_engagement - 1) + "): "))
#Here, if 0 (All) is choosen, we will admit you wants full report, maybe future, we can see the yearly option

helper_engagement = 0 #control
helper_number_engagement = 0 #row
for engagement in engagement_array:
   if engagement_option == helper_number_engagement:
     helper_engagement = engagement
     helper_number_engagement += 1
   else:
     helper_number_engagement += 1


print("\n")
print("\n")

#if engagement_option == 0:
#  sql_select_tests_query = ("SELECT id,engagement_id,title FROM dojo_test")
#  cursor.execute(sql_select_tests_query, (engagement_option, ))
#  records_tests = cursor.fetchall()
#else:
#sql_select_tests_query = ("select dojo_test.id, engagement_id, concat( if(title is not null or title = '', concat(dojo_test.title,'(',dojo_test_type.name,')' ), dojo_test_type.name) ) as nameEST  from dojo_test INNER JOIN dojo_test_type ON dojo_test_type.id = test_type_id where engagement_id = %s")
#cursor.execute(sql_select_tests_query, (engagement_option, ))
#records_tests = cursor.fetchall()

#  if cursor.rowcount > 0:
#    #if title is empty use test type query? to present the scan type instead of the title of test
#    #what happens if we dont have scans, or only one scan
#print("Found {} tests. Which ones do you want to use?".format(cursor.rowcount))

#n = cursor.rowcount + 1
#tests_array = [None] * n
#tests_array[0] = 0
#print("0  -  All Tests")
#number_test = 1;
#for row in records_tests:
#   print(number_test, " - " ,row[2],)
#   tests_array[number_test] = row[0]
#   number_test += 1

#print("end")
#for x in range(len(tests_array)):
#    print(tests_array[x])

#scan_option = int(input("Select Option (0-" + str(number_test - 1) + "): "))
#print(type(scan_option))
#helper_test = 0 #control
#helper_number_test = 0 #row
#print("you choosed:" ,scan_option)
#print("no useful:",helper_number_test)

#print("Array 1 ID is: ",tests_array[1])

#for test in tests_array:
#   if scan_option == helper_number_test:
#     helper_test = test
#     helper_number_test += 1
     #print("here: ",helper_test)
     #print("here: ",helper_number_test)
#   else:
#     helper_number_test += 1
     #print("no:",helper_number_test)
     #print("no:",helper_test)


#print("helper test:",helper_number_test)
#print("helper test is: ",helper_test)
print("\n")
print("\n")

#active_findings_option = int(input("Show Active Findings? No(0), Yes(1): "))

print("\n")
print("\n")

#verified_findings_option = int(input("Show Verified Findings? No(0), Yes(1): "))

#print("\n")
#print("\n")

#false_positive_option = int(input("Show False Positive Findings? No(0), Yes(1): "))

#print("\n")
#print("\n")

executive_summary_option = int(input("Do you need Executive Summary? No(0), Yes(1): "))

print("\n")
print("\n")

finding_notes_option = int(input("Do you need Finding Notes? No(0), Yes(1): "))

print("\n")
print("\n")

finding_images_option = int(input("Do you need Finding Images? No(0), Yes(1): "))


print("\n")
print("\n")


templates_count = 0
for file in glob.glob("*.docx"):
    #print(file)
    templates_count += 1

print("Found {} templates, choose one:".format(templates_count))
#print("1 - XXX.docx")
#print("2 - YYY.docx")
#print("3 - ZZZ.docx")

templates_count_help = 1
for file in glob.glob("*.docx"):
    print(templates_count_help, " - " ,file)
    templates_count_help += 1
print("\n")

template_option = int(input("Select Option (1-" + str(templates_count) + "): "))


print("Generating report, please wait...")


#sql_select_findings_query = ("SELECT title,cwe,severity,description FROM dojo_finding WHERE active = %s AND verified = %s AND false_p = %s AND test_id = %s")
#sql_select_findings_query = ("SELECT auth_user.username, auth_user.first_name, auth_user.last_name, title, date, cwe, url, severity, description, mitigation, impact, active, created, cve FROM dojo_finding INNER JOIN auth_user ON auth_user.id = dojo_finding.reporter_id WHERE active = %s AND verified = %s AND false_p = %s AND test_id = %s order by severity='Info', severity='Low', severity='Medium', severity='High', severity='Critical'")
sql_select_findings_query = ("SELECT auth_user.username, auth_user.first_name, auth_user.last_name, dojo_finding.title, date, cwe, url, severity, dojo_finding.description, mitigation, impact, active, dojo_finding.created, cve FROM dojo_finding INNER JOIN auth_user ON auth_user.id = dojo_finding.reporter_id INNER JOIN dojo_test ON dojo_test.id = dojo_finding.test_id WHERE  dojo_test.engagement_id = %s order by severity='Info', severity='Low', severity='Medium', severity='High', severity='Critical'")
#data_findings = (active_findings_option, verified_findings_option, false_positive_option, helper_test)
data_findings = (helper_engagement,)

cursor.execute(sql_select_findings_query, data_findings)
row_headers=[x[0] for x in cursor.description] #this will extract row headers
records_findings = cursor.fetchall()

total_findings = cursor.rowcount
print("TOTAL FINDINGS CARALHO:",total_findings)

#result_fin = []
#columns_fin = tuple( [d[0].decode('utf8') for d in cursor.description] )
#for row in cursor:
#  result.append(dict(zip(columns, row)))
#print(result)


json_data=[]
for result in records_findings:
    json_data.append(dict(zip(row_headers,result)))

#def myconverter(o):
#    if isinstance(o, datetime.datetime):
#        return o.__str__()

y_findings = json.dumps(json_data, indent=4, sort_keys=True, default=str)
z_findings = json.loads(y_findings)

print(z_findings)
#print("\n")
#parsed_json = (json.loads(json_data))
#print(json.dumps(parsed_json, indent=4, sort_keys=True))


#####
# Leader Name
####
#sql_select_leader_query = ("select first_name,last_name,username, dojo_test.id, dojo_test.title from auth_user inner join dojo_test on auth_user.id = dojo_test.lead_id where dojo_test.id = %s")
#cursor.execute(sql_select_leader_query, (helper_test, ))
#records_leader = cursor.fetchall()

leader_name = "ascdsda"
#for row in records_leader:
#   leader_name = row[2]


###################
# ENGAGEMENT NAME #
###################
sql_select_engagement_name_query = ("select name from dojo_engagement where dojo_engagement.id = %s")
cursor.execute(sql_select_engagement_name_query, (helper_engagement, ))
records_engagement_name = cursor.fetchall()

engagement_name = ""
for row in records_engagement_name:
   engagement_name = row[0]


#################
# ALL ENDPOINTS #
#################
#sql_select_endpoints_query = ("select distinct(dojo_endpoint.host) from dojo_endpoint inner join dojo_finding_endpoints ON endpoint_id = dojo_endpoint.id INNER JOIN dojo_finding ON finding_id = dojo_finding.id INNER join dojo_test ON dojo_test.id = dojo_finding.test_id where test_id = %s")
sql_select_endpoints_query = ("select distinct(dojo_endpoint.host) from dojo_endpoint inner join dojo_finding_endpoints ON endpoint_id = dojo_endpoint.id INNER JOIN dojo_finding ON finding_id = dojo_finding.id INNER join dojo_test ON dojo_test.id = dojo_finding.test_id INNER JOIN dojo_engagement ON dojo_engagement.id = dojo_test.engagement_id WHERE engagement_id = %s")
cursor.execute(sql_select_endpoints_query, (helper_engagement, ))
#cursor.execute(sql_select_endpoints_query, (helper_test, ))
row_headers_endpoint=[x[0] for x in cursor.description] #this will extract row headers
records_endpoints = cursor.fetchall()

endpoint_length = cursor.rowcount
endpoint_array = [None] * endpoint_length
endpoint_array[0] = 0
number_endpoint = 0
for row in records_endpoints:
   endpoint_array[number_endpoint] = row[0]
   number_endpoint += 1


#json_data_endpoints=[]
#for result in records_endpoints:
#    json_data_endpoints.append(dict(zip(row_headers_endpoint,result)))

#y_endpoints = json.dumps(json_data_endpoints)
#z_endpoints = json.loads(y_endpoints)
#print(z_endpoints)


def grouper(n, iterable):
    """
    >>> list(grouper(3, 'ABCDEFG'))
    [['A', 'B', 'C'], ['D', 'E', 'F'], ['G']]
    """
    iterable = iter(iterable)
    return iter(lambda: list(IT.islice(iterable, n)), [])

#seq = [1,2,3,4,5,6,7]
#list_endpoints = list(grouper(2, z_endpoints))
list_endpoints = list(grouper(2, endpoint_array))
print(list_endpoints)

#findings_array = [[0 for x in range(3)] for y in range(cursor.rowcount)]
#numberhelper = 0
#for row in records_findings:
   #print(row[0], " - ",row[1]," - ",row[2])
#   findings_array[numberhelper][0] = row[0]
#   findings_array[numberhelper][1] = row[1]
#   findings_array[numberhelper][2] = row[2]
#   numberhelper += 1

#for x in range(len(findings_array)):
#   print(findings_array[x])


doc = DocxTemplate("/home/dojo/templates/template.docx")
context = { 'leader_name' : leader_name, 'row_contents': z_findings, 'engagement_name': engagement_name, 'row_endpoints': list_endpoints, 'total_findings': total_findings }

#'findings': [
#        {'name': finding, 'severity': severity},
#
#        {
#            'status': status,
#            'dateDiscovered': dateDiscovered,
#            'age': age,
#            'reporter': reporter,
#            'description': description
#        },
#    ],

doc.render(context, autoescape=True)
doc.save("generated_doc.docx")

print("Progress Bar")
