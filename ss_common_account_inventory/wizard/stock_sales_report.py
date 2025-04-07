import base64
from io import BytesIO
from odoo import fields, models, _, api
from odoo.tools.misc import xlwt
from xlwt import easyxf
from odoo.exceptions import ValidationError
from odoo.exceptions import UserError, ValidationError
from datetime import datetime, date, time,timedelta as dt_timedelta
import calendar
from calendar import month
from odoo.exceptions import UserError


class salesitem_monthly_report(models.TransientModel):
    _name = 'salesitem.monthly.report'
    _description = 'Sales Per Item Monthly Report'
    
    start_date = fields.Date(string='From',required="1",default=date.today())
    end_date = fields.Date(string='To ',required="1")


    product_by = fields.Selection([('all','All'),('selected','Selected')]
    											,default='all',string='Chose Product By',required="1")
    partner_by = fields.Selection([('all','All'),('selected','Selected')]
    											,default='all',string='Chose Partner By',required="1")
    product_ids = fields.Many2many('product.product',string='Product')
    partner_ids = fields.Many2many('res.partner',string='Customer')
    method_by = fields.Selection([('periode','Periodic'),('sumary','Sumary')]
    											,default='sumary',string='Data by',required="1")

    excel_file = fields.Binary('Excel File')
    durasi              = fields.Float("Durasi")    

    @api.onchange('durasi',)
    def on_durasi_create(self):
        if self.end_date != False:
            if self.durasi < 0:
                raise UserError('End date tidak boleh Back date!!') 

    def get_date_list(self):
        date_list = []
        date_range = self.end_date - self.start_date

        for day in range(date_range.days + 1):
            current_date = self.start_date + dt_timedelta(days=day)
            
            if current_date.day == 1:
                date_in_list = current_date.strftime("%Y-%m-%d")
                da_date = current_date.strftime("%B %Y %d")
                date_list.append({
                    'date': date_in_list,
                    'day':da_date,
                })

        return date_list


################## Body Excel report 1 ##################

    def create_excel_header(self,worksheet,date_list):
        main_header_style = easyxf('align: horiz left,vert center;'
                              'font:bold True,height 250;')
        sub_header = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                              'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin')
        sub_header_date = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                                    'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin',num_format_str='dd')

        worksheet.write_merge(0, 3, 0, 2, self.env.user.company_id.name or '', main_header_style)
        
        worksheet.write_merge(5,5, 0,5,'Laporan Penjualan per Item' , sub_header)
        worksheet.write_merge(6,6, 0,5,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B %Y'), sub_header)

        if self.method_by == 'sumary':
            worksheet.write_merge(8,9, 0,0,'No.' , sub_header)
            worksheet.write_merge(8,9, 1,1,'Nama Barang' , sub_header)
            worksheet.write_merge(8,9, 2,2,'Kemasan' , sub_header)
            worksheet.write(8, 3, 'Qty', sub_header)
            worksheet.write(9, 3, 'Sales Terinvoice', sub_header)
            worksheet.write(8, 4, 'HPP', sub_header)
            worksheet.write(9, 4, 'Excl', sub_header)
            worksheet.write(8, 5, 'Harga', sub_header)
            worksheet.write(9, 5, 'Produk', sub_header)
        elif self.method_by == 'periode':
            worksheet.write_merge(8,10, 0,0,'No.' , sub_header)
            worksheet.write_merge(8,10, 1,1,'Nama Barang' , sub_header)
            worksheet.write_merge(8,10, 2,2,'Kemasan' , sub_header)
            col = 2
            for d in date_list:
                col +=1
                worksheet.write_merge(8,8, col,col+2,'Periode '+d.get('day'), sub_header)
                worksheet.write(9, col, 'Qty', sub_header)
                worksheet.write(10, col, 'Sales Terinvoice', sub_header)
                worksheet.write(9, col+1, 'HPP', sub_header)
                worksheet.write(10, col+1, 'Excl', sub_header)
                worksheet.write(9, col+2, 'Harga', sub_header)
                worksheet.write(10, col+2, 'Produk', sub_header)
                worksheet.col(col).width = 70 * 70
                worksheet.col(col+1).width = 70 * 70
                worksheet.col(col+2).width = 90 * 90
                worksheet.col(col+3).width = 20 * 20
                col +=3
        worksheet.col(0).width = 70 * 30
        worksheet.col(1).width = 140 * 140
        worksheet.col(2).width = 70 * 70
        worksheet.col(3).width = 70 * 70
        worksheet.col(4).width = 70 * 70
        return worksheet


    def create_excel_value(self,worksheet,date_list):
        text_left = easyxf('font:height 200;align:vert center,horiz left;' 'borders: top thin,bottom thin,left thin, right thin')
        text_right_accounting = easyxf('font:height 200;align:vert center,horiz right;' 'borders: top thin,bottom thin,left thin, right thin',num_format_str='#,##0.00_);(#,##0.00)')
        text_center = easyxf('font:height 200;align:vert center,horiz center;' 'borders: top thin,bottom thin,left thin, right thin')

        if self.product_by == 'all':
        	product_ids = self.env['product.product'].search([('sale_ok', '=', True)],order='name asc')
        else:
        	if not self.product_ids:
        		raise UserError(_('No Product selected'))
        	product_ids = self.product_ids
        seq = 0
        seq2 = 0
        row = 10
        kemasan = ''
#####################<<<<<<<<<<<<<<<<<<<<<<<<<<
        if self.method_by == 'sumary':
            for product in product_ids:
                qty_sum = 0.00
                total_sum = 0.00
                total_debit = 0.00
                hpp_item = 0.00
                search_inv = self.env['account.move.line'].search([('journal_id', '=', 9),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
                if search_inv:
                    if search_inv.product_id:
                        if 'galon 4 liter' in product.name.lower() or 'pail' in product.name.lower():
                            if 'galon 4 liter' in product.name.lower():
                                kemasan = '4 Liter'
                            elif 'pail' in product.name.lower():
                                kemasan = 'Pail'
                            row +=1
                            seq +=1
                            qty_sum = sum(line.quantity for line in search_inv)
                            total_sum = sum(line2.price_subtotal for line2 in search_inv)
                            search_inv_hpp = self.env['account.move.line'].search([('journal_id', '=', 9),('product_id', '=', product.id),('display_type', '=', 'cogs'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
                            if search_inv_hpp:
                                for deep_line_hpp in search_inv_hpp:
                                    total_debit += deep_line_hpp.debit
                            hpp_item = total_debit/qty_sum
                            worksheet.write(row,0,seq,text_center)
                            worksheet.write(row,1,product.name,text_left)
                            worksheet.write(row,2,kemasan,text_center)
                            worksheet.write(row,3,qty_sum,text_center)
                            worksheet.write(row,4,hpp_item,text_right_accounting)
                            worksheet.write(row,5,total_sum,text_right_accounting)
            row +=3
            worksheet.write_merge(row,row, 0,2,'Kemasan  Karton' , text_left)
            for product in product_ids:
                qty_sum = 0.00
                total_sum = 0.00
                total_debit = 0.00
                hpp_item = 0.00
                search_inv = self.env['account.move.line'].search([('journal_id', '=', 9),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
                if search_inv:
                    if search_inv.product_id:
                        if 'galon 4 liter' not in product.name.lower() and 'pail' not in product.name.lower():
                            row +=1
                            seq2 +=1
                            qty_sum = sum(line.quantity for line in search_inv)
                            total_sum = sum(line2.price_subtotal for line2 in search_inv)
                            search_inv_hpp = self.env['account.move.line'].search([('journal_id', '=', 9),('product_id', '=', product.id),('display_type', '=', 'cogs'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
                            if search_inv_hpp:
                                for deep_line_hpp in search_inv_hpp:
                                    total_debit += deep_line_hpp.debit
                            hpp_item = total_debit/qty_sum
                            worksheet.write(row,0,seq2,text_center)
                            worksheet.write(row,1,product.name,text_left)
                            worksheet.write(row,2,'',text_center)
                            worksheet.write(row,3,qty_sum,text_center)
                            worksheet.write(row,4,hpp_item,text_right_accounting)
                            worksheet.write(row,5,total_sum,text_right_accounting)
#####################<<<<<<<<<<<<<<<<<<<<<<<<<<
        elif self.method_by == 'periode':
            for product in product_ids:
                qty_sum = 0.00
                total_sum = 0.00
                total_debit = 0.00
                hpp_item = 0.00
                col = 2
                search_inv = self.env['account.move.line'].search([('journal_id', '=', 9),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
                if search_inv:
                    if search_inv.product_id:
                        if 'galon 4 liter' in product.name.lower() or 'pail' in product.name.lower():
                            if 'galon 4 liter' in product.name.lower():
                                kemasan = '4 Liter'
                            elif 'pail' in product.name.lower():
                                kemasan = 'Pail'
                            row +=1
                            seq +=1
                            worksheet.write(row,0,seq,text_center)
                            worksheet.write(row,1,product.name,text_left)
                            worksheet.write(row,2,kemasan,text_center)
                            for d in date_list:
                                qty_sum = 0.00
                                total_sum = 0.00
                                total_debit = 0.00
                                hpp_item = 0.00
                                col +=1
                                date = datetime.strptime(d.get('date'), '%Y-%m-%d')
                                last_day = calendar.monthrange(date.year, date.month)[1]
                                date_end = date.replace(day=last_day)
                                search_inv2 = self.env['account.move.line'].search([('journal_id', '=', 9),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', date),('move_id.invoice_date', '<=', date_end)])
                                if search_inv2:
                                    qty_sum = sum(line.quantity for line in search_inv2)
                                    total_sum = sum(line2.price_subtotal for line2 in search_inv2)
                                    search_inv_hpp = self.env['account.move.line'].search([('journal_id', '=', 9),('product_id', '=', product.id),('display_type', '=', 'cogs'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', date),('move_id.invoice_date', '<=', date_end)])
                                    if search_inv_hpp:
                                        for deep_line_hpp in search_inv_hpp:
                                            total_debit += deep_line_hpp.debit
                                    hpp_item = total_debit/qty_sum       
                                worksheet.write(row,col,qty_sum,text_center)
                                worksheet.write(row,col+1,hpp_item,text_right_accounting)
                                worksheet.write(row,col+2,total_sum,text_right_accounting)
                                col +=3
            row +=3
            worksheet.write_merge(row,row, 0,2,'Kemasan  Karton' , text_left)
            for product in product_ids:
                qty_sum = 0.00
                total_sum = 0.00
                total_debit = 0.00
                hpp_item = 0.00
                col = 2
                search_inv = self.env['account.move.line'].search([('journal_id', '=', 9),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
                if search_inv:
                    if search_inv.product_id:
                        if 'galon 4 liter' not in product.name.lower() and 'pail' not in product.name.lower():
                            row +=1
                            seq2 +=1
                            worksheet.write(row,0,seq2,text_center)
                            worksheet.write(row,1,product.name,text_left)
                            worksheet.write(row,2,'',text_center)
                            for d in date_list:
                                qty_sum = 0.00
                                total_sum = 0.00
                                total_debit = 0.00
                                hpp_item = 0.00
                                col +=1
                                date = datetime.strptime(d.get('date'), '%Y-%m-%d')
                                last_day = calendar.monthrange(date.year, date.month)[1]
                                date_end = date.replace(day=last_day)
                                search_inv2 = self.env['account.move.line'].search([('journal_id', '=', 9),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', date),('move_id.invoice_date', '<=', date_end)])
                                if search_inv2:
                                    qty_sum = sum(line.quantity for line in search_inv2)
                                    total_sum = sum(line2.price_subtotal for line2 in search_inv2)
                                    search_inv_hpp = self.env['account.move.line'].search([('journal_id', '=', 9),('product_id', '=', product.id),('display_type', '=', 'cogs'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', date),('move_id.invoice_date', '<=',  date_end)])
                                    if search_inv_hpp:
                                        for deep_line_hpp in search_inv_hpp:
                                            total_debit += deep_line_hpp.debit
                                    hpp_item = total_debit/qty_sum
                                worksheet.write(row,col,qty_sum,text_center)
                                worksheet.write(row,col+1,hpp_item,text_right_accounting)
                                worksheet.write(row,col+2,total_sum,text_right_accounting)
                                col +=3
        return worksheet


################## Body Excel report 2 ##################

    def create_excel_header2(self,worksheet2,date_list):
        main_header_style = easyxf('align: horiz left,vert center;'
                              'font:bold True,height 250;')
        sub_header = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                              'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin')

        worksheet2.write_merge(0, 3, 0, 1, self.env.user.company_id.name or '', main_header_style)
        
        worksheet2.write_merge(5,5, 0,1,'Laporan Penjualan per Customer' , sub_header)
        worksheet2.write_merge(6,6, 0,1,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B %Y'), sub_header)

        worksheet2.write_merge(8,9, 0,0,'Nama Customer' , sub_header)
        worksheet2.col(0).width = 140 * 140
        if self.method_by == 'sumary':
            worksheet2.write_merge(8,9, 1,1,'Sales Per Customer (Rp)' , sub_header)
        elif self.method_by == 'periode':
            col = 0
            for d in date_list:
                col +=1
                worksheet2.write(8, col,'Periode '+d.get('day'), sub_header)
                worksheet2.write(9, col,'Sales Per Customer (Rp)', sub_header)
                worksheet2.col(col).width = 70 * 70
                col +=1
                worksheet2.col(col).width = 20 * 20
        return worksheet2


    def create_excel_value2(self,worksheet2,date_list):
        text_left = easyxf('font:height 200;align:vert center,horiz left;' 'borders: top thin,bottom thin,left thin, right thin')
        text_right_accounting = easyxf('font:height 200;align:vert center,horiz right;' 'borders: top thin,bottom thin,left thin, right thin',num_format_str='#,##0.00_);(#,##0.00)')
        text_center = easyxf('font:height 200;align:vert center,horiz center;' 'borders: top thin,bottom thin,left thin, right thin')

        if self.partner_by == 'all':
        	partner_ids = self.env['res.partner'].search([],order='name asc')
        else:
        	if not self.partner_ids:
        		raise UserError(_('No Customer selected'))
        	partner_ids = self.partner_ids
        row = 9
        total_sum = 0.00
        for partner in partner_ids:
            search_inv = self.env['account.move.line'].search([('journal_id', '=', 9),('move_id.partner_id', '=', partner.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
            if search_inv:
                if search_inv.product_id:
                    row +=1
                    if self.method_by == 'sumary':
                        worksheet2.write(row,0,partner.name,text_left)
                        total_sum = sum(line.price_subtotal for line in search_inv)
                        worksheet2.write(row,1,total_sum,text_right_accounting)
                    elif self.method_by == 'periode':
                        worksheet2.write(row,0,partner.name,text_left)
                        col = 0
                        for d in date_list:
                            total_sum = 0.00
                            col +=1
                            date = datetime.strptime(d.get('date'), '%Y-%m-%d')
                            last_day = calendar.monthrange(date.year, date.month)[1]
                            date_end = date.replace(day=last_day)
                            search_inv2 = self.env['account.move.line'].search([('journal_id', '=', 9),('move_id.partner_id', '=', partner.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', date),('move_id.invoice_date', '<=', date_end)])
                            if search_inv2:
                                if search_inv2.product_id:
                                    total_sum = sum(line.price_subtotal for line in search_inv2)
                                    worksheet2.write(row,col,total_sum,text_right_accounting)
                                else:
                                    worksheet2.write(row,col,total_sum,text_right_accounting)
                            else:
                                worksheet2.write(row,col,total_sum,text_right_accounting)
                            col +=1
        return worksheet2


################## Body Excel report 3 ##################

    def create_excel_header3(self,worksheet3,date_list):
        main_header_style = easyxf('align: horiz left,vert center;'
                              'font:bold True,height 250;')
        sub_header = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                              'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin')

        worksheet3.write_merge(0, 3, 0, 2, self.env.user.company_id.name or '', main_header_style)
        
        worksheet3.write_merge(5,5, 0,5,'Laporan Penjualan per UOM' , sub_header)
        worksheet3.write_merge(6,6, 0,5,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B %Y'), sub_header)

        worksheet3.write_merge(8,10, 0,0,'Nama Customer' , sub_header)
        worksheet3.col(0).width = 140 * 140
        col = 0
        if self.method_by == 'sumary':
            worksheet3.write_merge(8,8, 1,5,'Sales per UOM' , sub_header)
            worksheet3.write_merge(9,9, 1,5,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B') , sub_header)
            worksheet3.write(10, 1, 'Pail', sub_header)
            worksheet3.write(10, 2, 'Galon', sub_header)
            worksheet3.write(10, 3, '1 Liter', sub_header)
            worksheet3.write(10, 4, '500 ML', sub_header)
            worksheet3.write(10, 5, 'Pouch', sub_header)
        elif self.method_by == 'periode':
            for d in date_list:
                col += 1
                worksheet3.write_merge(8,8, col,col+4,'Sales per UOM' , sub_header)
                worksheet3.write_merge(9,9, col,col+4,'Periode '+d.get('day') , sub_header)
                worksheet3.write(10, col, 'Pail', sub_header)
                worksheet3.write(10, col+1, 'Galon', sub_header)
                worksheet3.write(10, col+2, '1 Liter', sub_header)
                worksheet3.write(10, col+3, '500 ML', sub_header)
                worksheet3.write(10, col+4, 'Pouch', sub_header)
                worksheet3.col(col).width = 45 * 45
                worksheet3.col(col+1).width = 45 * 45
                worksheet3.col(col+2).width = 45 * 45
                worksheet3.col(col+3).width = 45 * 45
                worksheet3.col(col+4).width = 45 * 45
                col += 5
                worksheet3.col(col).width = 20 * 20
        return worksheet3


    def create_excel_value3(self,worksheet3,date_list):
        text_left = easyxf('font:height 200;align:vert center,horiz left;' 'borders: top thin,bottom thin,left thin, right thin')
        text_right_accounting = easyxf('font:height 200;align:vert center,horiz right;' 'borders: top thin,bottom thin,left thin, right thin',num_format_str='#,##0.00_);(#,##0.00)')
        text_center = easyxf('font:height 200;align:vert center,horiz center;' 'borders: top thin,bottom thin,left thin, right thin')

        if self.partner_by == 'all':
        	partner_ids = self.env['res.partner'].search([],order='name asc')
        else:
        	if not self.partner_ids:
        		raise UserError(_('No Customer selected'))
        	partner_ids = self.partner_ids
        row = 10
        for partner in partner_ids:
            search_inv = self.env['account.move.line'].search([('journal_id', '=', 9),('move_id.partner_id', '=', partner.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
            if search_inv:
                if search_inv.product_id:
                    row +=1
                    worksheet3.write(row,0,partner.name,text_left)
                    if self.method_by == 'sumary':
                        tot_pail = 0
                        tot_gal = 0
                        tot_lit = 0
                        tot_500 = 0
                        tot_pouch = 0
                        for data_prod in search_inv:
                            if data_prod.product_id.name and isinstance(data_prod.product_id.name, str):
                                product_name = data_prod.product_id.name.lower()
                                if 'pail' in product_name:
                                    tot_pail += data_prod.quantity
                                elif 'galon 4 liter' in product_name:
                                    tot_gal += data_prod.quantity
                                elif 'galon 1 liter' in product_name:
                                    tot_lit += data_prod.quantity
                                elif '500 ml' in product_name:
                                    tot_500 += data_prod.quantity
                                else:
                                    tot_pouch += data_prod.quantity

                        worksheet3.write(row,1,tot_pail,text_center)
                        worksheet3.write(row,2,tot_gal,text_center)
                        worksheet3.write(row,3,tot_lit,text_center)
                        worksheet3.write(row,4,tot_500,text_center)
                        worksheet3.write(row,5,tot_pouch,text_center)
                    elif self.method_by == 'periode':
                        col = 0
                        for d in date_list:
                            tot_pail = 0
                            tot_gal = 0
                            tot_lit = 0
                            tot_500 = 0
                            tot_pouch = 0
                            col +=1
                            date = datetime.strptime(d.get('date'), '%Y-%m-%d')
                            last_day = calendar.monthrange(date.year, date.month)[1]
                            date_end = date.replace(day=last_day)
                            search_inv2 = self.env['account.move.line'].search([('journal_id', '=', 9),('move_id.partner_id', '=', partner.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', date),('move_id.invoice_date', '<=', date_end)])
                            if search_inv2:
                                for data_prod in search_inv2:
                                    if data_prod.product_id.name and isinstance(data_prod.product_id.name, str):
                                        product_name = data_prod.product_id.name.lower()
                                        if 'pail' in product_name:
                                            tot_pail += data_prod.quantity
                                        elif 'galon 4 liter' in product_name:
                                            tot_gal += data_prod.quantity
                                        elif 'galon 1 liter' in product_name:
                                            tot_lit += data_prod.quantity
                                        elif '500 ml' in product_name:
                                            tot_500 += data_prod.quantity
                                        else:
                                            tot_pouch += data_prod.quantity

                            worksheet3.write(row,col,tot_pail,text_center)
                            worksheet3.write(row,col+1,tot_gal,text_center)
                            worksheet3.write(row,col+2,tot_lit,text_center)
                            worksheet3.write(row,col+3,tot_500,text_center)
                            worksheet3.write(row,col+4,tot_pouch,text_center)
                            col +=5


################## Body Excel report 4 ##################

    def create_excel_header4(self,worksheet4,date_list):
        main_header_style = easyxf('align: horiz left,vert center;'
                              'font:bold True,height 250;')
        sub_header = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                              'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin')
        sub_header_date = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                                    'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin',num_format_str='dd')

        worksheet4.write_merge(0, 3, 0, 2, self.env.user.company_id.name or '', main_header_style)
        
        worksheet4.write_merge(5,5, 0,5,'Laporan Penjualan per Item' , sub_header)
        worksheet4.write_merge(6,6, 0,5,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B %Y'), sub_header)

        if self.method_by == 'sumary':
            worksheet4.write_merge(8,9, 0,0,'No.' , sub_header)
            worksheet4.write_merge(8,9, 1,1,'Nama Barang' , sub_header)
            worksheet4.write_merge(8,9, 2,2,'Kemasan' , sub_header)
            worksheet4.write(8, 3, 'Qty', sub_header)
            worksheet4.write(9, 3, 'Sales Terinvoice', sub_header)
            worksheet4.write(8, 4, 'HPP', sub_header)
            worksheet4.write(9, 4, 'Excl', sub_header)
            worksheet4.write(8, 5, 'Harga', sub_header)
            worksheet4.write(9, 5, 'Produk', sub_header)
        elif self.method_by == 'periode':
            worksheet4.write_merge(8,10, 0,0,'No.' , sub_header)
            worksheet4.write_merge(8,10, 1,1,'Nama Barang' , sub_header)
            worksheet4.write_merge(8,10, 2,2,'Kemasan' , sub_header)
            col = 2
            for d in date_list:
                col +=1
                worksheet4.write_merge(8,8, col,col+2,'Periode '+d.get('day'), sub_header)
                worksheet4.write(9, col, 'Qty', sub_header)
                worksheet4.write(10, col, 'Sales Terinvoice', sub_header)
                worksheet4.write(9, col+1, 'HPP', sub_header)
                worksheet4.write(10, col+1, 'Excl', sub_header)
                worksheet4.write(9, col+2, 'Harga', sub_header)
                worksheet4.write(10, col+2, 'Produk', sub_header)
                worksheet4.col(col).width = 70 * 70
                worksheet4.col(col+1).width = 70 * 70
                worksheet4.col(col+2).width = 90 * 90
                worksheet4.col(col+3).width = 20 * 20
                col +=3
        worksheet4.col(0).width = 70 * 30
        worksheet4.col(1).width = 140 * 140
        worksheet4.col(2).width = 70 * 70
        worksheet4.col(3).width = 70 * 70
        worksheet4.col(4).width = 70 * 70
        return worksheet4


    def create_excel_value4(self,worksheet4,date_list):
        text_left = easyxf('font:height 200;align:vert center,horiz left;' 'borders: top thin,bottom thin,left thin, right thin')
        text_right_accounting = easyxf('font:height 200;align:vert center,horiz right;' 'borders: top thin,bottom thin,left thin, right thin',num_format_str='#,##0.00_);(#,##0.00)')
        text_center = easyxf('font:height 200;align:vert center,horiz center;' 'borders: top thin,bottom thin,left thin, right thin')

        if self.product_by == 'all':
        	product_ids = self.env['product.product'].search([('sale_ok', '=', True)],order='name asc')
        else:
        	if not self.product_ids:
        		raise UserError(_('No Product selected'))
        	product_ids = self.product_ids
        if self.partner_by == 'all':
        	partner_ids = self.env['res.partner'].search([],order='name asc')
        else:
        	if not self.partner_ids:
        		raise UserError(_('No Customer selected'))
        	partner_ids = self.partner_ids

        seq = 0
        row = 10
        kemasan = ''
#####################<<<<<<<<<<<<<<<<<<<<<<<<<<
        if self.method_by == 'sumary':
            for partner in partner_ids:
                search_inv_partner = self.env['account.move.line'].search([('journal_id', '=', 9),('move_id.partner_id', '=', partner.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
                if search_inv_partner:
                    row +=1
                    worksheet4.write_merge(row,row, 0,2,partner.name, text_left)
                    for product in product_ids:
                        qty_sum = 0.00
                        total_sum = 0.00
                        total_debit = 0.00
                        hpp_item = 0.00
                        search_inv = self.env['account.move.line'].search([('journal_id', '=', 9),('move_id.partner_id', '=', partner.id),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
                        if search_inv:
                            if search_inv.product_id:
                                if 'galon 4 liter' in product.name.lower():
                                    kemasan = '4 Liter'
                                elif 'pail' in product.name.lower():
                                    kemasan = 'Pail'
                                row +=1
                                seq +=1
                                qty_sum = sum(line.quantity for line in search_inv)
                                total_sum = sum(line2.price_subtotal for line2 in search_inv)
                                search_inv_hpp = self.env['account.move.line'].search([('journal_id', '=', 9),('product_id', '=', product.id),('display_type', '=', 'cogs'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
                                if search_inv_hpp:
                                    for deep_line_hpp in search_inv_hpp:
                                        total_debit += deep_line_hpp.debit
                                hpp_item = total_debit/qty_sum
                                worksheet4.write(row,0,seq,text_center)
                                worksheet4.write(row,1,product.name,text_left)
                                worksheet4.write(row,2,kemasan,text_center)
                                worksheet4.write(row,3,qty_sum,text_center)
                                worksheet4.write(row,4,hpp_item,text_right_accounting)
                                worksheet4.write(row,5,total_sum,text_right_accounting)
                    row +=1
#####################<<<<<<<<<<<<<<<<<<<<<<<<<<
        elif self.method_by == 'periode':
            for partner in partner_ids:
                search_inv_partner = self.env['account.move.line'].search([('journal_id', '=', 9),('move_id.partner_id', '=', partner.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
                if search_inv_partner:
                    row +=1
                    worksheet4.write_merge(row,row, 0,2,partner.name, text_left)
                    for product in product_ids:
                        col = 2
                        search_inv = self.env['account.move.line'].search([('journal_id', '=', 9),('move_id.partner_id', '=', partner.id),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
                        if search_inv:
                            if search_inv.product_id:
                                if 'galon 4 liter' in product.name.lower():
                                    kemasan = '4 Liter'
                                elif 'pail' in product.name.lower():
                                    kemasan = 'Pail'
                                row +=1
                                seq +=1
                                worksheet4.write(row,0,seq,text_center)
                                worksheet4.write(row,1,product.name,text_left)
                                worksheet4.write(row,2,kemasan,text_center)
                                for d in date_list:
                                    qty_sum = 0.00
                                    total_sum = 0.00
                                    total_debit = 0.00
                                    hpp_item = 0.00
                                    col +=1
                                    date = datetime.strptime(d.get('date'), '%Y-%m-%d')
                                    last_day = calendar.monthrange(date.year, date.month)[1]
                                    date_end = date.replace(day=last_day)
                                    search_inv2 = self.env['account.move.line'].search([('journal_id', '=', 9),('move_id.partner_id', '=', partner.id),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', date),('move_id.invoice_date', '<=', date_end)])
                                    if search_inv2:
                                        qty_sum = sum(line.quantity for line in search_inv2)
                                        total_sum = sum(line2.price_subtotal for line2 in search_inv2)
                                        search_inv_hpp = self.env['account.move.line'].search([('journal_id', '=', 9),('move_id.partner_id', '=', partner.id),('product_id', '=', product.id),('display_type', '=', 'cogs'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', date),('move_id.invoice_date', '<=', date_end)])
                                        if search_inv_hpp:
                                            for deep_line_hpp in search_inv_hpp:
                                                total_debit += deep_line_hpp.debit
                                        hpp_item = total_debit/qty_sum
                                    worksheet4.write(row,col,qty_sum,text_center)
                                    worksheet4.write(row,col+1,hpp_item,text_right_accounting)
                                    worksheet4.write(row,col+2,total_sum,text_right_accounting)
                                    col +=3
                    row +=1
        return worksheet4


    def export_excel(self):
        if self.end_date < self.start_date:
            raise ValidationError(_('End Date must be greater than Start Date'))
        workbook = xlwt.Workbook()
        filename = 'Report_sales_SNB.xls'

        date_list = self.get_date_list()
        #report 1
        worksheet = workbook.add_sheet('Penjualan Per Item')
        for c in range(0, 100):
            worksheet.col(c).width = 140 * 30
        for c in range(0, 100):
            if c <= 7:
                worksheet.row(c).height = 250
            else:
                worksheet.row(c).height = 350
        worksheet = self.create_excel_header(worksheet,date_list)
        worksheet = self.create_excel_value(worksheet,date_list)

        #report 2
        worksheet2 = workbook.add_sheet('Penjualan Per Customer')
        for c in range(0, 100):
            worksheet2.col(c).width = 140 * 30
        for c in range(0, 100):
            if c <= 7:
                worksheet2.row(c).height = 250
            else:
                worksheet2.row(c).height = 350
        worksheet2 = self.create_excel_header2(worksheet2,date_list)
        worksheet2 = self.create_excel_value2(worksheet2,date_list)

        #report 3
        worksheet3 = workbook.add_sheet('Penjualan Per UOM')
        for c in range(0, 100):
            worksheet3.col(c).width = 140 * 30
        for c in range(0, 100):
            if c <= 7:
                worksheet3.row(c).height = 250
            else:
                worksheet3.row(c).height = 350
        worksheet3 = self.create_excel_header3(worksheet3,date_list)
        worksheet3 = self.create_excel_value3(worksheet3,date_list)

        #report 4
        worksheet4 = workbook.add_sheet('Penjualan Per Customer per Item')
        for c in range(0, 100):
            worksheet4.col(c).width = 140 * 30
        for c in range(0, 100):
            if c <= 7:
                worksheet4.row(c).height = 250
            else:
                worksheet4.row(c).height = 350
        worksheet4 = self.create_excel_header4(worksheet4,date_list)
        worksheet4 = self.create_excel_value4(worksheet4,date_list)

        fp = BytesIO()
        workbook.save(fp)
        fp.seek(0)
        excel_file = base64.encodebytes(fp.read())
        fp.close()

        self.write({'excel_file': excel_file})

        if self.excel_file:
            active_id = self.ids[0]
            return {
                'type': 'ir.actions.act_url',
                'url': 'web/content/?model=salesitem.monthly.report&download=true&field=excel_file&id=%s&filename=%s' % (
                    active_id, filename),
                'target': 'new',
            }

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
