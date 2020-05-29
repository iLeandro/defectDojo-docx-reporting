import mysql.connector
from mysql.connector import Error
import glob, os
import json
from datetime import datetime
import time
import sys
import itertools as IT

try:
   from docxtpl import DocxTemplate, Listing
except ImportError:
	print("Missing docxtpl library. Please install it using PIP. https://pypi.org/project/docxtpl/")
	exit()

os.chdir("/home/dojo/templates")

try:
   mydb = mysql.connector.connect(
      host="172.18.0.3",
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
def grouper(n, iterable):
   iterable = iter(iterable)
   return iter(lambda: list(IT.islice(iterable, n)), [])


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
   sql_select_engagements_query = ("SELECT id,name,product_id,target_start,target_end FROM dojo_engagement WHERE product_id = %s")
   cursor.execute(sql_select_engagements_query, (client_id, ))
   records_engagements = cursor.fetchall()

except Error as engagementsQueryError:
    print("Error", engagementsQueryError)
    sys.exit()


helper_engagement = -1 #control
#client has no engagements
if cursor.rowcount == 0:
   print("\n")
   print("This client doesn't have engagements.")
   sys.exit()
#client has one engagement   
elif cursor.rowcount == 1:
   print("\n")
   print("Found {} engagement. Do you want to use it?".format(cursor.rowcount))
   
   #engagement_length = cursor.rowcount
   engagement_array = [None] * 1 #engagement_length
   engagement_array[0] = 0
   number_engagement = 0
   for row in records_engagements:
      print("1 - " ,row[1],)
      engagement_array[number_engagement] = row[0]

   print("\n")
   engagement_option = int(input("Select Option. No(0), Yes(1): "))
   
   #check if user wants to use the engagement
   if engagement_option == 0:
      sys.exit()
   else:
      #print(engagement_option, "EO")
      helper_engagement = 0 #control
      helper_number_engagement = 1 #row
      #saves the engagement ID to a variable
      for engagement in engagement_array:
         if engagement_option == helper_number_engagement:
            helper_engagement = engagement
            #print("Engagement ID: ",helper_engagement)
            helper_number_engagement += 1
         else:
            #print("falhou ",helper_number_engagement, " he ",helper_engagement)
            helper_number_engagement += 1

      sql_select_tests_query = ("SELECT dojo_test.id, engagement_id, concat( if(title is not null or title = '', concat(dojo_test.title,'(',dojo_test_type.name,')' ), dojo_test_type.name) ) as nameEST FROM dojo_test INNER JOIN dojo_engagement ON dojo_test.engagement_id = dojo_engagement.id INNER JOIN dojo_test_type ON dojo_test_type.id = test_type_id WHERE dojo_engagement.id = %s")
      cursor.execute(sql_select_tests_query, (helper_engagement, ))
      records_tests = cursor.fetchall()

      if cursor.rowcount == 0:
         print("This engagement doesn't have tests.")
         sys.exit()
      else:
         n = cursor.rowcount
         tests_array = [None] * n
         tests_array[0] = 0
         number_test = 0
         for row in records_tests:
            print(number_test+1, " - " ,row[2],)
            tests_array[number_test] = row[0]
            number_test += 1

         #print("----------")
         for x in range(len(tests_array)):
            print(tests_array[x])

         sql_select_findings_query = ("SELECT auth_user.username, auth_user.first_name, auth_user.last_name, dojo_finding.title, date, cwe, url, severity, dojo_finding.description, mitigation, impact, active, dojo_finding.created, cve FROM dojo_finding INNER JOIN auth_user ON auth_user.id = dojo_finding.reporter_id INNER JOIN dojo_test ON dojo_test.id = dojo_finding.test_id WHERE  dojo_test.engagement_id = %s order by severity='Info', severity='Low', severity='Medium', severity='High', severity='Critical'")
         data_findings = (helper_engagement,)
         cursor.execute(sql_select_findings_query, data_findings)
         row_headers=[x[0] for x in cursor.description] #this will extract row headers
         records_findings = cursor.fetchall()

         json_data=[]
         for result in records_findings:
            json_data.append(dict(zip(row_headers,result)))

         y_findings = json.dumps(json_data, indent=4, sort_keys=True, default=str)
         z_findings = json.loads(y_findings)

         print(z_findings)
         #################
         total_findings = cursor.rowcount
         print("TOTAL FINDINGS ENCONTRADAS:",total_findings)

#client has two or more engagements
else:
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

   #save engagment id upon selected option
   helper_engagement = 0 #control
   helper_number_engagement = 0 #row
   for engagement in engagement_array:
      if engagement_option == helper_number_engagement:
         helper_engagement = engagement
         print("Saving engagement ID: ",helper_engagement)
         helper_number_engagement += 1
      else:
         helper_number_engagement += 1

   if engagement_option == 0:
      sql_select_tests_query = ("SELECT dojo_test.id, engagement_id, title FROM dojo_test INNER JOIN dojo_engagement ON dojo_test.engagement_id = dojo_engagement.id WHERE product_id = %s")
      cursor.execute(sql_select_tests_query, (client_id, ))
      records_tests = cursor.fetchall()
      ##STORE ALL ENGAGEMENTS IDS?
      ##
      ## FLAGS TO KNOW SOMETHING RIGHT>HGTHT
      ##
   else:
      #sql_select_tests_query = ("select dojo_test.id, engagement_id, concat( if(title is not null or title = '', concat(dojo_test.title,'(',dojo_test_type.name,')' ), dojo_test_type.name) ) as nameEST  from dojo_test INNER JOIN dojo_test_type ON dojo_test_type.id = test_type_id where engagement_id = %s")
      sql_select_tests_query = ("SELECT dojo_test.id, engagement_id, concat( if(title is not null or title = '', concat(dojo_test.title,'(',dojo_test_type.name,')' ), dojo_test_type.name) ) as nameEST FROM dojo_test INNER JOIN dojo_engagement ON dojo_test.engagement_id = dojo_engagement.id INNER JOIN dojo_test_type ON dojo_test_type.id = test_type_id WHERE dojo_engagement.id = %s")
      cursor.execute(sql_select_tests_query, (helper_engagement, ))
      records_tests = cursor.fetchall()

      if cursor.rowcount == 0:
         print("This engagement doesn't have tests.")
         sys.exit()
      else:
         n = cursor.rowcount
         tests_array = [None] * n
         tests_array[0] = 0
         number_test = 0
         for row in records_tests:
            print(number_test+1, " - " ,row[2],)
            tests_array[number_test] = row[0]
            number_test += 1

         #print("-----------")
         for x in range(len(tests_array)):
            print(tests_array[x])

         sql_select_findings_query = ("SELECT auth_user.username, auth_user.first_name, auth_user.last_name, dojo_finding.id AS f_id, dojo_finding.title, date, cwe, url, severity, dojo_finding.description, mitigation, impact, active, dojo_finding.created, cve, refs FROM dojo_finding INNER JOIN auth_user ON auth_user.id = dojo_finding.reporter_id INNER JOIN dojo_test ON dojo_test.id = dojo_finding.test_id WHERE  dojo_test.engagement_id = %s order by severity='Info', severity='Low', severity='Medium', severity='High', severity='Critical'")
         data_findings = (helper_engagement,)
         cursor.execute(sql_select_findings_query, data_findings)
         row_headers=[x[0] for x in cursor.description] #this will extract row headers
         records_findings = cursor.fetchall()

         json_data=[]
         for result in records_findings:
            json_data.append(dict(zip(row_headers,result)))

         y_findings = json.dumps(json_data, indent=4, sort_keys=True, default=str)
         z_findings = json.loads(y_findings)

         print(z_findings)
         print("-----------------")
         #################
         #sql_select_endpoints_query = ("SELECT DISTINCT dojo_endpoint.id, dojo_endpoint.host, dojo_endpoint.protocol FROM dojo_endpoint inner join dojo_finding_endpoints ON endpoint_id = dojo_endpoint.id INNER JOIN dojo_finding ON finding_id = dojo_finding.id INNER join dojo_test ON dojo_test.id = dojo_finding.test_id INNER JOIN dojo_engagement ON dojo_engagement.id = dojo_test.engagement_id WHERE engagement_id = %s")
         sql_select_endpoints_query = ("SELECT DISTINCT dojo_endpoint.id, dojo_endpoint.host, dojo_endpoint.protocol, dojo_finding_endpoints.finding_id FROM dojo_endpoint INNER JOIN dojo_finding_endpoints ON endpoint_id = dojo_endpoint.id INNER JOIN dojo_finding ON finding_id = dojo_finding.id INNER join dojo_test ON dojo_test.id = dojo_finding.test_id INNER JOIN dojo_engagement ON dojo_engagement.id = dojo_test.engagement_id WHERE engagement_id = %s ORDER BY finding_id ASC")
         data_endpoints = (helper_engagement,)
         cursor.execute(sql_select_endpoints_query, data_endpoints)
         row_headers_endpoints=[x[0] for x in cursor.description] #this will extract row headers
         records_endpoints = cursor.fetchall()

         json_data_endpoints=[]
         for results in records_endpoints:
            json_data_endpoints.append(dict(zip(row_headers_endpoints,results)))

         y_endpoints = json.dumps(json_data_endpoints, indent=4, sort_keys=True, default=str)
         z_endpoints = json.loads(y_endpoints)

         print(z_endpoints)
         print("-------------------------------")

         flag_counter_fe = 0
         for finding in z_findings:
            for endpoint in z_endpoints:
               if finding['f_id'] == endpoint['finding_id']:
                  print("Found Finding ID: ",finding['f_id']," - ",finding['title']," = Endpoint F_ID = ",endpoint['finding_id'], " - ",endpoint['host'])
                  #list_endpoints_ungroup = str(endpoint['finding_id']) + " - " + endpoint['host']
                  #list_endpoints_group = list(grouper(2, list_endpoints_ungroup))
                  flag_counter_fe += 1

         print("FCFE: " ,flag_counter_fe)
         print(list_endpoints_group)
         total_findings = cursor.rowcount
         print("TOTAL FINDINGS ENCONTRADAS:",total_findings)





finding_notes_option = int(input("Do you need Finding Notes? No(0), Yes(1): ")) #NOT WORKING

print("\n")
print("\n")

finding_images_option = int(input("Do you need Finding Images? No(0), Yes(1): ")) #NOT WORKING

print("\n")
print("\n")


templates_count = 0
for file in glob.glob("*.docx"):
    #print(file)
    templates_count += 1

print("Found {} templates, choose one:".format(templates_count))

templates_count_help = 1
for file in glob.glob("*.docx"):
    print(templates_count_help, " - " ,file)
    templates_count_help += 1
print("\n")

template_option = int(input("Select Option (1-" + str(templates_count) + "): ")) #NOT WORKING

###############
# Leader Name #
###############
#sql_select_leader_query = ("select first_name,last_name,username, dojo_test.id, dojo_test.title from auth_user inner join dojo_test on auth_user.id = dojo_test.lead_id where dojo_test.id = %s")
sql_select_leader_query = ("SELECT first_name, last_name, username, name FROM dojo_engagement INNER JOIN auth_user ON auth_user.id = dojo_engagement.lead_id WHERE dojo_engagement.id = %s")
cursor.execute(sql_select_leader_query, (helper_engagement, ))
records_leader = cursor.fetchall()

leader_name = ""
for row in records_leader:
   leader_name = row[2]


###################
# ENGAGEMENT NAME #
###################
sql_select_engagement_name_query = ("SELECT name FROM dojo_engagement WHERE dojo_engagement.id = %s") #USE THE OTHER QUERY
cursor.execute(sql_select_engagement_name_query, (helper_engagement, ))
records_engagement_name = cursor.fetchall()

engagement_name = ""
for row in records_engagement_name:
   engagement_name = row[0]


#################
# ALL ENDPOINTS #
#################
#sql_select_endpoints_query = ("select distinct(dojo_endpoint.host) from dojo_endpoint inner join dojo_finding_endpoints ON endpoint_id = dojo_endpoint.id INNER JOIN dojo_finding ON finding_id = dojo_finding.id INNER join dojo_test ON dojo_test.id = dojo_finding.test_id where test_id = %s")
#sql_select_endpoints_query = ("select distinct dojo_endpoint.host from dojo_endpoint inner join dojo_finding_endpoints ON endpoint_id = dojo_endpoint.id INNER JOIN dojo_finding ON finding_id = dojo_finding.id INNER join dojo_test ON dojo_test.id = dojo_finding.test_id INNER JOIN dojo_engagement ON dojo_engagement.id = dojo_test.engagement_id WHERE engagement_id = %s")
sql_select_endpoints_query = ("SELECT DISTINCT SUBSTRING_INDEX(host, ':', 1) FROM dojo_endpoint INNER JOIN dojo_product ON dojo_product.id = dojo_endpoint.product_id INNER JOIN dojo_engagement ON dojo_engagement.product_id  = dojo_product.id WHERE dojo_engagement.id = %s ORDER BY SUBSTRING_INDEX(host, ':', 1) ASC")
cursor.execute(sql_select_endpoints_query, (helper_engagement, ))
#cursor.execute(sql_select_endpoints_query, (helper_test, ))
#row_headers_endpoint=[x[0] for x in cursor.description] #this will extract row headers
records_endpoints = cursor.fetchall()

endpoint_length = cursor.rowcount
endpoint_array = [None] * endpoint_length
endpoint_array[0] = 0
number_endpoint = 0
for row in records_endpoints:
   endpoint_array[number_endpoint] = row[0]
   number_endpoint += 1

print(endpoint_array)
print("---------------")
#json_data_endpoints=[]
#for result in records_endpoints:
#   json_data_endpoints.append(dict(zip(row_headers_endpoint,result)))

#y_endpoints = json.dumps(json_data_endpoints)
#z_endpoints = json.loads(y_endpoints)
#print(z_endpoints)

##CALS GROUPER
list_endpoints = list(grouper(2, endpoint_array))
print(list_endpoints)

doc = DocxTemplate("/home/dojo/templates/template.docx")
context = { 'leader_name' : leader_name, 'row_contents': z_findings, 'engagement_name': engagement_name, 'all_endpoints': list_endpoints, 'row_endpoints': z_endpoints, 'total_findings': total_findings, 'arrayTest': endpoint_array }

doc.render(context, autoescape=True)
doc.save("generated_doc.docx")


print("\n  ######################################################################################################")	
print('    #  Report Generation started:     '+ str(datetime.now()) + '                                         #')
print("    ######################################################################################################\n")

print("\n  ######################################################")
print("    #                                                    #")
print("    #   [!] Records are beeing uploaded to the template. #")
print("    #                                                    #")
print("    ######################################################")
		

def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print('\r%s |%s| %s%% %s' % (prefix, bar, percent, suffix), end = printEnd)
    # Print New Line on Complete
    if iteration == total: 
        print()

i = 1
l = 100
for i in range(l):
   printProgressBar(i + 1, l, prefix = '    Progress:', suffix = 'Complete', length = 70)
   i += 1


print("\n\n    ###################################################################################################")	
print('    #  Script ended:       '+ str(datetime.now()) + '                                                 #')
print("    ###################################################################################################\n\n")
