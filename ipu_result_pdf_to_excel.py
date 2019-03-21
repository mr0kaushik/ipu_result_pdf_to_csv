import camelot
import xlsxwriter

def fileInput():
    print('Enter Input File Path (eg: D:\Web\input.pdf)') 
    input_path_with_name = input('Input Path with File: ')
    return input_path_with_name
    try: 
        fi = open (input_path_with_name, "r")
        fi.close()
    except IOError: 
        print ('\n\nERROR:\n\tThere is no file named', input_path_with_name)
        fileInput()



input_path = fileInput()
print('\nEnter Output File Path (eg: D:\Web') 
output_path = input('Output Path : ')

print('\nEnter Pages (eg: \'2, 3, 4\' OR \'2-5\' Or \'2-end\'') 
pages_num = input('Enter pages : ')

output_path_with_file = output_path + "\\"+'output.xlsx'

print('\nIt will take some time, Depending on number of tables and data it will parse.')

tables = camelot.read_pdf(input_path, pages=pages_num)


print('No. of Tables: ' + str(tables.n))

workbook = xlsxwriter.Workbook(output_path_with_file)
worksheet = workbook.add_worksheet()
#print('\nSpecify the tables: \nSuch as \'2, 5\' for excuting on tables 2-5')
#start_table, last_table = input('Tables : ').split(',')
st = 0
lt = tables.n
xrow=-1
for t in range(st,lt):
    if tables[t].parsing_report['whitespace']!=0.59:
        df=tables[t].df
        for row in range(1,len(df.index), 3):
            xrow+=1
            xcol=0
            for col in range(2,len(df.columns)):
                val = df[col][row]
                if col == 2:
                    if len(val)>2:
                        data = val.split('\n')
                        roll_number = data[0]
                        name = data[1]
                        worksheet.write(xrow, xcol, roll_number)
                        xcol+=1
                        worksheet.write(xrow, xcol, name)
                        xcol+=1
                        #print('Roll Number - ' + roll_number + "  " + 'Name - ' + name)
                if col>2 and col<23:
                    subject_data = df[col][row]
                    two_marks_data = df[col][row+1]
                    total_mark_data_with_grade = df[col][row+2]

                    if len(subject_data)!=0:
                        subject_code = subject_data.split('(')[0]
                        #print('Subject code ' + subject_code)
                        worksheet.write(xrow, xcol, subject_code)
                        xcol+=1

                    if len(two_marks_data)!=0:
                        first, second = two_marks_data.split()
                        #print('First '+ first + ' Second ' + second)
                        worksheet.write(xrow, xcol, first)
                        xcol+=1
                        worksheet.write(xrow, xcol, second)
                        xcol+=1

                    if len(total_mark_data_with_grade)!=0:
                        #print("Total Marks = "+ total_mark_data_with_grade)
                        worksheet.write(xrow, xcol, total_mark_data_with_grade)
                        xcol+=1

                if col==23 and len(val)>0:
                    #print('credit ' + val)
                    worksheet.write(xrow, xcol, val)
        xrow+=2
workbook.close()
print('\nSuccessfully Completed')
