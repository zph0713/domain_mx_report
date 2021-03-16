#*-* coding: UTF-8
import os,sys
import io
import dns.resolver as dns_resolver
import json
import xlwt
import re
import time
from xlrd import open_workbook
import xlutils.copy



FileConfig = {
    'export_xls_file':'domainlist_mx.xls',
    'domain_list_file':'domain.list'
}

supplier_mapping = {
    'trendmicro':'TrendMicro',
    'pphosted':'Proofpoint',
    'pinpointit':'Pinpointit',
    'google':'GMail',
    'MessageLabs':'messagelabs',
    'outlook':'Office365',
    'mimecast':'Mimecast',
    'ppe-hosted':'Proofpoint',
    'fireeye':'Fireeye',
    'barracuda':'Barracuda',
    'cisco':'Cisco',
    'fortimail':'Fortinet'
}




class transfer_to_xls(object):
    def __init__(self):
        #self.xls_file_name = 'domainlist_mx.xls'
        self.xls_file_name = FileConfig['export_xls_file']
        self.sheet_name = 'MX'

    def write_to_workbook(self,final_data):
        if not os.path.isfile(self.xls_file_name):
            wb = xlwt.Workbook()
            ws = wb.add_sheet(self.sheet_name)
            wb.save(self.xls_file_name)
        rb = open_workbook(self.xls_file_name, formatting_info=True)
        wb = xlutils.copy.copy(rb)
        sheet_names = [ s.name for s in rb.sheets() ]
        if self.sheet_name in sheet_names:
            ws =  wb.get_sheet(self.sheet_name)
        else:
            ws = wb.add_sheet(self.sheet_name)
        row = 0
        ws.write(row, 0, "domain_name")
        ws.write(row, 1, "services")
        ws.write(row, 2, "supplier")
        ws.write(row, 3, "mx_record")

        for data in final_data:
            row += 1
            ws.write(row, 0, data['domain'])
            ws.write(row,1,data['services'])
            ws.write(row,2,data['supplier'])
            num = 3
            for mxs in data['mx_record']:
                ws.write(row, num, mxs)
                num += 1
        wb.save(self.xls_file_name)

class DNS_MX(object):
    def __init__(self):
        self.query_type = 'MX'
        self.mx_record = list()

    def resolver_domain(self,domain_name):
        exchange_list = list()
        try:
            rd = dns_resolver.query(domain_name,self.query_type)
            for i in rd:
                mx_exchage = i.exchange
                exchange_list.append(str(mx_exchage)[:-1])
        except Exception:
            exchange_list.append("NoAnswer")
        return self.record_mx(domain_name,exchange_list)

    def record_mx(self,domain_name,exchange_list):
        domain_name_te = domain_name.split('.')[0]
        mx_count = len(exchange_list)
        self_total=0
        trd_total=0
        supplier = ''

        if domain_name != 'NULL' and exchange_list[0] != 'NoAnswer':
            for mx in exchange_list:
                mxsplit = mx.split('.')
                if domain_name_te in mxsplit:  
                    self_total += 1
                else:
                    trd_total += 1
            if self_total == mx_count and trd_total == 0:
                services_type = 'Self Hosted'
            else:
                services_type = '3rd Party Hosted'
                supplier = ','.join(self.supplier_check(exchange_list))

        else:
            services_type = 'Unknown'

        print {'domain' :domain_name_te,'mx_record' :exchange_list,'services' :services_type,'supplier':supplier}
        self.mx_record.append({'domain' :domain_name_te,'mx_record' :exchange_list,'services' :services_type,'supplier':supplier})

    def supplier_check(self,mx_list):
        sup_list = supplier_mapping.keys()
        mx_set = list()
        for mx in mx_list:
            mxsplit = mx.split('.')
            mxc = list(set(sup_list).intersection(set(mxsplit)))
            if mxc != []:
                mx_set.append(mxc[0])
        mx_set = set(mx_set)
        mx_used = list()
        for i in mx_set:
            if i in sup_list:
                mx_used.append(supplier_mapping[i])
        if 'TrendMicro' in mx_used:
            return ['TrendMicro']
        return mx_used

    def multiple_query_thread(self):
        for domain in self.read_domain_file():
            self.resolver_domain(domain)

    def read_domain_file(self):
        domain_list = list()
        #domain_file = os.path.join('domain.list')
        domain_file = os.path.join(FileConfig['domain_list_file'])
        with io.open(domain_file,'r',encoding="utf-8") as f:
            fr = f.readlines()
            for i in fr:
                domain_n = i.strip('\n')
                domain_list.append(domain_n)
        return domain_list



if __name__ == '__main__':
    MX = DNS_MX()
    MX.multiple_query_thread()
    data_mx = MX.mx_record
   # print data_mx
    TXLS = transfer_to_xls()
    TXLS.write_to_workbook(data_mx)
