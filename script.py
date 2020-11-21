import psycopg2
import xlrd
import base64
import xmlrpc.client
from datetime import datetime

class Product():
    
    def __init__(self):
        """Initializing database connection
        """
        db_params = {
            'host': 'localhost',
            'database': 'odoo',
            'user': 'ubuntu',
            'password': '1234',
            'port': '5432'
        }
        try:
            self.db = psycopg2.connect(**db_params)
            self.cursor = self.db.cursor()
        except Exception as e:
            print('Database Connection Failed: '+str(e))
            
    def odoo_connect(self):
        """Initializing odoo connection with xmlrpc
        """
        self.url = 'http://localhost:8069'
        self.db = 'odoo'
        self.user = 'admin'
        self.pwd = 'admin'
        try:
            common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(self.url))
            self.uid = common.authenticate(self.db, self.user, self.pwd, {})
            return xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(self.url))
        except Exception as e:
            print('Odoo Connection Failed: '+str(e))
            return None
            
    def get_groups(self):
        """Returns id,name of all records from product_groups
        """
        self.cursor.execute('SELECT id,name FROM product_group')
        return self.cursor.fetchall()    
        
    def insert_product(self):
        """Reads Excel file, gets group id for each group name,
           opens each image file and decodes to base64,
           appends each record to a new list, and finally
           splits the list and passes each splited list to
           create function of xmlprc.
        """
        sheet = xlrd.open_workbook('db.xlsx').sheet_by_index(0)
        vals = sheet.row_values
        nrows = sheet.nrows
        groups = self.get_groups()
        records = []
        for i in range(1, nrows):
            row = vals(i)
            for g in groups:
                if g[1] == row[3]:
                    row[3] = g[0]
                    break
            try:
                img = open('photos/'+row[4], "rb")
                encoded = base64.b64encode(img.read()).decode('utf-8')
                row[4] = encoded
            except Exception:
                row[4] = False
            
            records.append({
                'name': row[1],
                'default_code': row[0],
                'image_1920': row[4],
                'group_id': row[3],
                'part_number': row[2]
            })  
        odoo = self.odoo_connect()
        if odoo:
            step = 10    
            steps = len(records)//step
            total = 0
            nrows = str(nrows)
            print('Uploading...\n')
            if steps>0:
                count = -1
                for i in range(0, len(records), step):
                    count = count+1
                    split = records[i:i+step] if count<steps else records[i:]         
                    ids = odoo.execute_kw(self.db, self.uid, self.pwd, 'product.template', 'create',
                        [split]
                    )
                    total = total + len(ids)
                    print(str(total)+'/'+nrows)
            else:
                ids = odoo.execute_kw(self.db, self.uid, self.pwd, 'product.template', 'create',
                        [records]
                    )
                total = total + len(ids)
                print(str(total)+'/'+nrows)
            print('\nTotal '+str(total+1)+'/'+nrows)
            
        
start = datetime.now()
Product().insert_product()
print(str(datetime.now()-start))
