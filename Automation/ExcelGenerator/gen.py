import random
import xlsxwriter
import mysql.connector
import sys
from datetime import date

def main():
	
	#print( sys.argv[1] )
	workbook = xlsxwriter.Workbook( "/home/abdulla/Desktop/stm.xlsx" )
	HeadFormat = workbook.add_format( { 'bold' : True} )
	HeadFormat.set_fg_color( '#31869B' )
	HeadFormat.set_font_color( '#FFFFFF' )
	worksheet = workbook.add_worksheet( 'ASTG' )
	worksheet.write( 'A1', 'Source Database', HeadFormat )
	worksheet.write( 'B1', 'Source Schema', HeadFormat )
	worksheet.write( 'C1', 'Soure Table', HeadFormat )
	worksheet.write( 'D1', 'Source Column', HeadFormat )
	worksheet.write( 'E1', 'Source Datatype', HeadFormat )
	worksheet.write( 'F1', 'Target Database', HeadFormat )
	worksheet.write( 'G1', 'Target Schema', HeadFormat )
	worksheet.write( 'H1', 'Target Table', HeadFormat )
	worksheet.write( 'I1', 'Target Column', HeadFormat )
	worksheet.write( 'J1', 'Target Datatype', HeadFormat )
	worksheet.write( 'K1', 'Transformation Logic', HeadFormat )
	
	try :
	        mydb = mysql.connector.connect( host='localhost',  user='root',  password='password1234',  database='mydatabase' )

        #print(mydb)
        	cursor = mydb.cursor()
        	cursor.execute("SELECT SourceDatabase, SourceSchema, SourceTable, SourceColumn, SourceDatatype, SourceDatalength  FROM stm")
		i=2
        	for (SourceDatabase, SourceSchema, SourceTable,  SourceColumn, SourceDatatype, SourceDatalength ) in cursor:
                	#print(x)
			worksheet.write( 'A' + str(i), SourceDatabase )
			worksheet.write( 'B' + str(i) , SourceSchema )
	 	        worksheet.write( 'C' + str(i), SourceTable )
			worksheet.write( 'D' + str(i), SourceColumn )
			worksheet.write( 'E' + str(i), SourceDatatype )
                        worksheet.write( 'F' + str(i) , 'ASTG' )
                        worksheet.write( 'G' + str(i), 'ADMIN' )
			worksheet.write( 'H' + str(i) , 'ASTG_' + sys.argv[2] )
			worksheet.write( 'I' + str(i), SourceColumn ) #Target Column is same as Source Column
			
			if SourceDatatype == 'NVARCHAR2':
				worksheet.write( 'J' + str(i), 'CHARACTER VARYING( ' + str(SourceDatalength) + ')' )
			elif SourceDatatype == 'DATE':
				worksheet.write( 'J' + str(i), 'TIMETAMP' )
			elif SourceDatatype == 'NUMBER':
                                worksheet.write( 'J' + str(i), 'NUMERIC(' + str(SourceDatalength) + ',0)' )
			elif SourceDatatype == 'NCLOB':
				worksheet.write(  'J' + str(i), 'CHARACTER VARYING(4000)' )

			if SourceDatatype == 'NCLOB':
	                        worksheet.write( 'K' + str(i), 'Cast to Character Varying 4000 and Load to Target Table' )
			else:
				worksheet.write( 'K' + str(i), 'Straight Load' )



			i = i + 1
        	mydb.close()

	except mysql.connector.Error as err:
  		print("Something went wrong: {}".format(err))


	##############################################################################
	
	#workbook = xlsxwriter.Workbook( "/home/abdulla/Desktop/apples.xlsx" )
        worksheet = workbook.add_worksheet( 'Revision History' )
        worksheet.write( 'A1', 'Revision No', HeadFormat )
        worksheet.write( 'B1', 'Revision Date', HeadFormat )
        worksheet.write( 'C1', 'Revised By', HeadFormat  )
        worksheet.write( 'D1', 'Comments', HeadFormat  )
	worksheet.write( 'A2', '0' )
        worksheet.write( 'B2', date.today().strftime('%m/%d/%Y') )
        worksheet.write( 'C2', 'AZ' )
        worksheet.write( 'D2', 'Initial Revision' )


	###############################################################

        worksheet = workbook.add_worksheet( 'AFL' )
        worksheet.write( 'A1', 'Source Database', HeadFormat  )
        worksheet.write( 'B1', 'Source Schema', HeadFormat  )
        worksheet.write( 'C1', 'Soure Table', HeadFormat  )
        worksheet.write( 'D1', 'Source Column', HeadFormat  )
	worksheet.write( 'E1', 'Source Datatype', HeadFormat  )
        worksheet.write( 'F1', 'Target Database', HeadFormat  )
        worksheet.write( 'G1', 'Target Schema', HeadFormat  )
        worksheet.write( 'H1', 'Target Table', HeadFormat  )
        worksheet.write( 'I1', 'Target Column', HeadFormat  )
	worksheet.write( 'J1', 'Source Datatype', HeadFormat  )
        worksheet.write( 'K1', 'Transformation Logic', HeadFormat  )

	worksheet.write( 'A2', 'NA' )
        worksheet.write( 'B2', 'NA' )
        worksheet.write( 'C2', 'NA' )
        worksheet.write( 'D2', 'NA' )
	worksheet.write( 'E2', 'NA' )
        worksheet.write( 'F2', 'TD2' )
        worksheet.write( 'F2', 'ADMIN' )
        worksheet.write( 'H2', str(sys.argv[1]) )
        worksheet.write( 'I2', 'AFL_ID' )
	worksheet.write( 'J2', 'BIGINT NOT NULL' )

        worksheet.write( 'K2', 'Sequence Generated' )

        try :
                mydb = mysql.connector.connect( host='localhost',  user='root',  password='password1234',  database='mydatabase' )

        #print(mydb)I
                cursor = mydb.cursor()
         
                     
             
		cursor.execute("SELECT SourceTable, SourceColumn, SourceDatatype, SourceDatalength  FROM stm")
                i=3
                for ( SourceTable,  SourceColumn, SourceDatatype, SourceDatalength ) in cursor:
                        #print(x)
                        worksheet.write( 'A' + str(i), 'ASTG' )
                        worksheet.write( 'B' + str(i) , 'ADMIN' )
                        worksheet.write( 'C' + str(i), SourceTable )
                        worksheet.write( 'D' + str(i), SourceColumn )
                        worksheet.write( 'E' + str(i), SourceDatatype )
                        worksheet.write( 'F' + str(i) , 'AFL' )
                        worksheet.write( 'G' + str(i), 'ADMIN' )
                        worksheet.write( 'H' + str(i) , 'AFL_' + sys.argv[2] )
                        worksheet.write( 'I' + str(i), SourceColumn ) #Target Column is same as Source Column

                        if SourceDatatype == 'NVARCHAR2':
                                worksheet.write( 'J' + str(i), 'CHARACTER VARYING( ' + str(SourceDatalength) + ')' )
                        elif SourceDatatype == 'DATE':
                                worksheet.write( 'J' + str(i), 'TIMETAMP' )
                        elif SourceDatatype == 'NUMBER':
                                worksheet.write( 'J' + str(i), 'NUMERIC(' + str(SourceDatalength) + ',0)' )
                        elif SourceDatatype == 'NCLOB':
                                worksheet.write(  'J' + str(i), 'CHARACTER VARYING(4000)' )

                        worksheet.write( 'K' + str(i), 'Straight Load' )


                        i = i + 1
                mydb.close()

		worksheet.write( 'A' + str(i), 'NA' )
                worksheet.write( 'B' + str(i) , 'NA' )
                worksheet.write( 'C' + str(i), 'NA' )
                worksheet.write( 'D' + str(i), 'NA' )
                worksheet.write( 'E' + str(i), 'NA' )
                worksheet.write( 'F' + str(i) , 'AFL' )
                worksheet.write( 'G' + str(i), 'ADMIN' )
                worksheet.write( 'H' + str(i) , 'AFL_' + sys.argv[2] )
                worksheet.write( 'I' + str(i), 'LOAD_DATE' ) #Target Column is same as Source Column
                worksheet.write( 'J' + str(i), 'DATE' )
                worksheet.write( 'K' + str(i), 'Current Load Date' )

        except mysql.connector.Error as err:
                print("Something went wrong: {}".format(err))










                        

	
	workbook.close()

if __name__ == "__main__":
	main()

       

