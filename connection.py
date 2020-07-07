#-*- coding:utf-8 -*-
import pymysql

import config


def select_sql(command):
    try:
        sqlconnection = pymysql.connect(host = config.MYSQL_IP, user=config.MYSQL_USER, password=config.MYSQL_PASSWORD, database=config.MYSQL_DB, charset = 'utf8')
        dbcursor = sqlconnection.cursor()
        dbcursor.execute(command)
        row_headers=[x[0] for x in dbcursor.description]
        data = dbcursor.fetchall()
        row_count = dbcursor.rowcount
        #return f"{data}, {row_count}, {row_headers}"
        return {'data': data, 'row_count': row_count, 'row_headers': row_headers}
		
    except Exception as inst:
        print('error: select_sql(): \n'+str(inst))

    finally:
        dbcursor.close()
        sqlconnection.close()