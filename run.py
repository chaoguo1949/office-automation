import sys  
import xlrd
import os


path = '..' + os.sep + 'new' + os.sep + '1234.xlsx'
wb = xlrd.open_workbook(path)
tables = wb.sheets()
with open('object_validation.sql', 'w+') as f:

    f.write('''SET SQL_SAFE_UPDATES = 0;\ndelete from object_validation where objectName like 'EAST%';\nINSERT INTO object_validation(objectName, attributeName, displayAttributeName, selfCheck, verificationType, verificationMethod, verificationLevel, verificationRole, enabled, ediDate) VALUES \n''')
    for table in tables:
        try:
            nrows = table.nrows #行数
            ncols = table.ncols #列数
            # print(nrows, ncols)
            colnames =  table.row_values(2) #某一行数据 
            li =[]
            for rownum in range(4,nrows):

                row = table.row_values(rownum)
                # print(row)
                if 'Audit fields' in row:
                    continue
                if '' == row[0]:
                    continue
                    
                if row:
                    app = {}
                    for i in range(len(colnames)):
                        # if 'Audit fields' in row[i]:
                        #     continue
                        app[colnames[i]] = row[i]
                li.append(app)
                # print(len(li))
            
            for i in li:
                # print(i)
                if i['Mandatory'] == 'M':
                    r = True

                if i['Mandatory'] == 'O':
                    r = False

                t = i['Field type (Format)']

                if t[0].lower() == 'v':
                    j = 0

                if t[0:4].lower() == 'date':
                    j = 1

                if t[0].lower() == 'i':
                    j = 2

                if t[0:2].lower() == 'de':
                    j = 3

                if r and j == 0:
                    ret = 'required||maxLength[%s]'% t[8:-1]
                if r and j == 1:
                    ret = 'required'

                if r and j == 2:
                    try:
                        ret = 'required||maxLength[%s]'% t[4:-1]
                    except:
                        ret = 'required'

                if r and j == 3:
                    ret = 'required||decimalOrBlank'

                if not r and j == 0:
                    ret = 'maxLength[%s]' % t[8:-1]

                if not r and j == 1:
                    ret = 'default'

                if not r and j == 2:
                    try:
                        ret = 'maxLength[%s]'% t[4:-1]
                    except:
                        ret = 'default'

                if not r and j == 3:
                    ret = 'decimalOrBlank'

                w =i['Field name'].lower()
                if w == 'edidate':
                    w = 'ediDate'
                    ret = 'required'
                if w == 'etltimestamp':
                    w = 'etlTimeStamp'
                    ret = 'default'
                data = "('%s','%s','%s','1','1','1','1','%s','1','2000-12-31'),\n" % (table.name,w,i['Chinese description'], ret)
                f.write(data)


        except Exception as ret:
            print(ret)

with open('ui_resource_config.sql', 'w+') as f:

    f.write('''SET SQL_SAFE_UPDATES = 0;\ndelete from ui_resource_config where parentElementID like 'EAST%';\nINSERT INTO ui_resource_config 
 (type,parentElementID,elementID,elementName,elementType,elementCSSType,displayFormatType,displayActionType,defaultValue,remark,isEnabled)
VALUES \n''')

    for table in tables[1:]:
        try:
            nrows = table.nrows #行数
            ncols = table.ncols #列数
            # print(nrows, ncols)
            colnames =  table.row_values(2) #某一行数据 
            li =[]
            for rownum in range(4,nrows):

                row = table.row_values(rownum)
                # print(row)
                if 'Audit fields' in row:
                    continue
                if '' == row[0]:
                    continue
                if row:
                    app = {}
                    for i in range(len(colnames)):
                        app[colnames[i]] = row[i]
                li.append(app)
            
            for i in li:
                # print(i)
                if i['Mandatory'] == 'M':
                    r = 'required'
                if i['Mandatory'] == 'O':
                    r = ''
                w =i['Field name'].lower()
                if w == 'edidate':
                    w = 'ediDate'
                    r = 'required'
                if w == 'etltimestamp':
                    w = 'etlTimeStamp'
                    r = 'default'
                data = "('I','%s','%s','','input','%s','display','',null,null,1),\n" % (table.name, w, r)
                f.write(data)
        except Exception as ret:
            print(ret)

# 替换最后一行末尾,为;
path = '.'+os.sep
files = os.listdir(path)
for file in files:
    data = ''
    
    if file[-4:] == '.sql':
        # print(file)
        filename = os.path.join(path, file)
        with open(filename, 'r') as f:
            ret = f.read()
            data = ret[:-2] + ';'
        with open(filename, 'w') as f:
            f.write(data)
