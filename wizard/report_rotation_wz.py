# -*- coding: utf-8 -*-
from osv import osv, fields
import time
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import csv, sys, os   
import base64
import unicodedata
import xlsxwriter


class report_rotation(osv.osv_memory):
    _name = 'report.rotation'
    _columns = {   
        'csv_file' : fields.binary('Csv Report File', readonly=False),
        'export_filename': fields.char('Export CSV Filename', size=128),
        'product_category_ids': fields.many2many('product.category','rotation_prod_categ_rel', 'rotation_id','category_id', 'Categorias'),
        'product_partner_ids': fields.many2many('res.partner','partner_rotation_rel', 'rotation_id','partner_id', 'Proveedores'),
        }

    def create_csv_report_rotation(self, cr, uid, ids, context=None):
        res = self.buscar_productos(cr, uid, ids, context)
        return {
            'name': 'Reporte rotacion productos',
            'view_type': 'form',
            'view_mode': 'form',
            'view_id': [res and res[1] or False],
            'res_model': 'report.rotation',
            'context': "{}",
            'type': 'ir.actions.act_window',
            'nodestroy': True,
            'target': 'new',
            'res_id': ids[0]  or False,##please replace record_id and provide the id of the record to be opened 
        }

    def ajustar_filtro(self, cat, prod):
        filter = ''
        if cat:
            filter += 'and pc.id in (%s) ' % ','.join(str(x) for x in cat) 
        if prod:
            filter += 'and pp.id in (%s)' % ','.join(str(x) for x in prod)        
        return filter

    def buscar_productos_por_partner(self, cr, uid, partners):
        if not partners:
            return []
        sql = """select product_id from product_supplierinfo 
                where name in (%s)
        """ % ''.join(str(x) for x in partners)
        cr.execute(sql)
        ids = cr.fetchall()
        if ids: 
            ids = map(lambda x:x[0], ids)
            return ids
        return ids
    
    def buscar_productos(self, cr, uid, ids, context=None):
        list_prod = []
        this = self.browse(cr, uid, ids[-1], context)
        list_prod = map(lambda x:x.id, this.product_category_ids)
        list_prov = map(lambda x:x.id, this.product_partner_ids)        
        list_prov = self.buscar_productos_por_partner(cr, uid, list_prov)
        filter = self.ajustar_filtro(list_prod, list_prov)
        sql = """
                select pp.id, pc.name, pp.default_code, pt.name ,(select name from res_partner where id = (select name from product_supplierinfo where product_id = pp.id limit 1)
                ) ,pp.min_qty , pp.max_qty
                from product_product pp 
                join product_template pt on (pp.id = pt.id)
                left join product_category pc on (pt.categ_id = pc.id)
                where pp.active = True 
                %s
              """ % filter
        cr.execute(sql)
        data = cr.fetchall()
        path = '/tmp/Reportes_emaresa.xlsx'
        workbook, worksheet = self.create_header(path)
        self.create_body(cr, uid, ids, workbook, worksheet, data, context)
        with open(path, 'r') as myfile:
            b64data = base64.b64encode(myfile.read())
        self.write(cr, uid, ids, {'csv_file':b64data}, {})
        self.write(cr, uid, ids, {'export_filename':'Reporte_Rotacion.xlsx'}, {})
        os.remove(path)
        mod_obj = self.pool.get('ir.model.data')
        res = mod_obj.get_object_reference(cr, uid, 'econube_reporte_rotacion', 'view_report_rotation')                                   
        return res

    def create_header(self, path):
        workbook = xlsxwriter.Workbook(path)
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:D', 25)
        worksheet.set_column('E:R', 15)
        bold = workbook.add_format({'bold': True, 'border': True})
        
        worksheet.write('A1', 'PROVEEDOR', bold)
        worksheet.write('B1', 'FAMILIA', bold)
        worksheet.write('C1', 'CODIGO', bold)
        worksheet.write('D1', 'DESCRIPCION', bold)
        worksheet.write('E1', 'CATEGORIA', bold)
        worksheet.write('F1', 'ANO ANTERIOR', bold)
        worksheet.write('G1', 'ANO ACTUAL', bold)
        worksheet.write('H1', 'ULTIMO 3 MES', bold)
        worksheet.write('I1', 'ULTIMO 2 MES', bold)
        worksheet.write('J1', 'ULTIMO MES', bold)
        worksheet.write('K1', 'MES SIG ANO ANTERIOR', bold)
        worksheet.write('L1', 'STOCK', bold)
        worksheet.write('M1', 'TRANSITO', bold)
        worksheet.write('N1', 'MIN', bold)
        worksheet.write('O1', 'MAX', bold)
        worksheet.write('P1', 'SUGERIDO', bold)
        worksheet.write('Q1', 'CANTIDAD', bold)
        worksheet.write('R1', 'PRECIO', bold)        
        return workbook, worksheet 

    def create_body(self, cr, uid, ids, workbook, worksheet, data, context=None):
        year = datetime.now() + relativedelta(month=1, day=1)
        year_last = datetime.now() + relativedelta(month=12, day=31)
        last_year      = datetime.now() - relativedelta(years=1,month=1, day=1)
        last_year_last = datetime.now() - relativedelta(years=1, month=12, day=31)
        last_3_month = datetime.now() - relativedelta(months=3, day=1)
        last_3_month_last = datetime.now() - relativedelta(months=3, day=31)
        last_2_month = datetime.now() - relativedelta(months=2, day=1)
        last_2_month_last = datetime.now() - relativedelta(months=2, day=31)
        actual_month = datetime.now() + relativedelta(day=1)
        actual_month_last = datetime.now() + relativedelta(day=31)
        year_lmont = datetime.now() + relativedelta(years=-1, months=+1, day = 1)
        year_lmont_last = datetime.now() + relativedelta(years=-1, months=+1, day = 31)
        product_obj = self.pool.get('product.product')
        row = 1
        bodystyle = workbook.add_format({'align': 'right'})
        for list in data:
            product = product_obj.browse(cr, uid, list[0], context)
            monto_3 = self.get_total_by_date(cr, uid, list[0], last_3_month.strftime('%Y-%m-%d'), last_3_month_last.strftime('%Y-%m-%d'), context)
            last_year_amount = self.get_total_by_date(cr, uid, list[0], year_lmont.strftime('%Y-%m-%d'), year_lmont_last.strftime('%Y-%m-%d'), context)            
            total_categ = self.get_total_move(cr, uid, actual_month.strftime('%Y-%m-%d'), actual_month_last.strftime('%Y-%m-%d'), product.id)
            worksheet.write(row, 0, self.special(list[4]))
            worksheet.write(row, 1, self.special(list[1]))
            worksheet.write(row, 2, self.special(list[2]))
            worksheet.write(row, 3, self.special(list[3]))
            worksheet.write(row, 4, self.get_category_stock(total_categ))                        
            worksheet.write(row, 5, self.get_total_by_date(cr, uid, list[0], last_year.strftime('%Y-%m-%d'), last_year_last.strftime('%Y-%m-%d'), context), bodystyle)
            worksheet.write(row, 6, self.get_total_by_date(cr, uid, list[0], year.strftime('%Y-%m-%d'), year_last.strftime('%Y-%m-%d'), context), bodystyle)
            worksheet.write(row, 7, monto_3, bodystyle)
            worksheet.write(row, 8, self.get_total_by_date(cr, uid, list[0], last_2_month.strftime('%Y-%m-%d'), last_2_month_last.strftime('%Y-%m-%d'), context), bodystyle)
            worksheet.write(row, 9, self.get_total_by_date(cr, uid, list[0], actual_month.strftime('%Y-%m-%d'), actual_month_last.strftime('%Y-%m-%d'), context), bodystyle)
            worksheet.write(row, 10, last_year_amount, bodystyle)
            worksheet.write(row, 11, product.qty_available, bodystyle)
            worksheet.write(row, 12, product.incoming_qty + product.outgoing_qty, bodystyle)
            worksheet.write(row, 13, list[5], bodystyle)
            worksheet.write(row, 14, list[6], bodystyle)
            worksheet.write(row, 15, ((monto_3 + last_year_amount ) /4 ) -product.qty_available - (product.incoming_qty + product.outgoing_qty), bodystyle)
            worksheet.write(row, 16, '', bodystyle)
            worksheet.write(row, 17, product.list_price, bodystyle)
            row +=1
        workbook.close()
        
    def get_total_by_date(self, cr, uid, product_id, date_start, date_stop, context=None):
        cr.execute(
        """
        select sum(pol.qty) from pos_order_line pol
        join pos_order po on (po.id = pol.order_id)
        where pol.product_id = %d
        and date(po.date_order) between '%s' and '%s'
        """% (product_id, date_start, date_stop) )
        total_tpv = cr.fetchone()        
        total_tpv = total_tpv[0] if total_tpv and total_tpv[0] is not None else 0                  
        cr.execute(
        """
        select sum(sol.product_uom_qty) from sale_order_line sol
        join sale_order so on (so.id = sol.order_id)
        where sol.product_id = %d
        and date(so.date_order) between '%s' and '%s'
        """ % (product_id, date_start, date_stop))
        total_order = cr.fetchone()
        total_order = total_order[0] if total_order and total_order[0] is not None else 0    
        return total_tpv + total_order   

    def special(self,valor):# saco caracteres especiales
        if valor == None:
            return ''
        return str(unicodedata.normalize('NFKD', unicode(valor)).encode('ascii','ignore'))

    def get_total_move(self, cr, uid, date, datelast, product_id):
        total = 0
        sql = """
                select case when sm.product_uom = pt.uom_id then sum(product_qty)
                        when sm.product_uom != pt.uom_id then round(sum(product_qty / pu.factor),0)
                        end as sum
                        , sp.type
                from stock_move sm  
                left join stock_picking sp on (sp.id = sm.picking_id)
                join product_template pt on (pt.id = sm.product_id)
                join product_uom pu on (pu.id = sm.product_uom)
                where sm.product_id = %d
                and sm.state = 'done' and (sp.type != 'internal' or sp.type is null)
                and DATE(sm.create_date) Between '%s' and '%s' 
                group by sp.type, sm.product_uom, pt.uom_id, sm.origin        """ % (product_id, date, datelast)
        cr.execute(sql)
        list = cr.fetchall()
        for data in list:
            if data[1] is None:
                total += data[0]
            elif data[1] == 'in':
                total += data[0]
            elif data[1] == 'out':
                total -= data[0]        
        return total
        
    def get_category_stock(self, total):
        if total > 100:
            return 'A'
        elif total >=31 and total <=100:
            return 'B'
        elif total >=11 and total <=30:
            return 'C'
        elif total >=4 and total <=10:
            return 'D'
        elif total >=0 and total <=3:
            return 'E'
        else:
            return 'E'
        
        
    def csv_file(self):
        with open(path, 'a') as myfile:            
            writer = csv.DictWriter(myfile, fieldnames=fieldnames)            
            writer.writerow({
                      'partner_id' : 'PROVEEDOR', 
                      'familia'    : 'FAMILIA',
                      'codigo'     : 'CODIGO',
                      'descripcion': 'DESCRIPCION',
                      'categoria'  : 'CATEGORIA', 
                      'last_year'  : 'AÑO ANTERIOR',
                      'year'       : 'AÑO ACTUAL',
                      'last_3_month': 'ULTIMO 3 MES',
                      'last_2_month':'ULTIMO 2 MES',
                      'month'      : 'ULTIMO MES',
                      'january'    : 'MES SIG AÑO ANTERIOR',
                      'stock'      : 'STOCK',
                      'transito'   : 'TRANSITO',
                      'min'        : 'MIN',
                      'max'        : 'MAX',
                      'sugerido'   : 'SUGERIDO',
                      'qty'        : 'CANTIDAD',
                      'price':  'PRECIO'
            })     

            for list in data:
                product = product_obj.browse(cr, uid, list[0], context)
                monto_3 = self.get_total_by_date(cr, uid, list[0], last_3_month.strftime('%Y-%m-%d'), last_3_month_last.strftime('%Y-%m-%d'), context)
                writer.writerow({
                          'partner_id' : self.special(list[4]), 
                          'familia'    : self.special(list[1]),
                          'codigo'     : self.special(list[2]),
                          'descripcion': self.special(list[3]),
                          'categoria'  :0, 
                          'last_year'  : self.get_total_by_date(cr, uid, list[0], last_year.strftime('%Y-%m-%d'), last_year_last.strftime('%Y-%m-%d'), context),
                          'year'       : self.get_total_by_date(cr, uid, list[0], year.strftime('%Y-%m-%d'), year_last.strftime('%Y-%m-%d'), context),
                          'last_3_month':monto_3,
                          'last_2_month':self.get_total_by_date(cr, uid, list[0], last_2_month.strftime('%Y-%m-%d'), last_2_month_last.strftime('%Y-%m-%d'), context),
                          'month'      : self.get_total_by_date(cr, uid, list[0], actual_month.strftime('%Y-%m-%d'), actual_month_last.strftime('%Y-%m-%d'), context),
                          'january'    : self.get_total_by_date(cr, uid, list[0], year.strftime('%Y-%m-%d'), year_last.strftime('%Y-01-31'), context),
                          'stock'      : product.qty_available,
                          'transito'   : product.incoming_qty + product.outgoing_qty,
                          'min'        : list[5],
                          'max'        : list[6],
                          'sugerido'   : (monto_3 / 3)-product.qty_available - (product.incoming_qty + product.outgoing_qty),
                          'qty'        :0,
                          'price': product.list_price
                      })     
                
report_rotation()