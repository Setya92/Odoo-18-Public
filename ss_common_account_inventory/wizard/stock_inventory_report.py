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


class stockinventory_monthly_report(models.TransientModel):
    _name = 'stockinventory.monthly.report'
    _description = 'Stock Monthly Report'
    
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

#########################################################
################## Body Excel report 1 ##################

    def create_excel_header(self,worksheet):
        main_header_style = easyxf('align: horiz left,vert center;'
                              'font:bold True,height 250;')
        sub_header = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                              'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin')

        worksheet.write_merge(0, 3, 0, 2, self.env.user.company_id.name or '', main_header_style)
        
        worksheet.write_merge(5,5, 0,5,'Laporan Penjualan per Item' , sub_header)
        worksheet.write_merge(6,6, 0,5,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B %Y'), sub_header)

        worksheet.write_merge(8,9, 0,0,'No.' , sub_header)
        worksheet.write_merge(8,9, 1,1,'Nama Barang' , sub_header)
        worksheet.write_merge(8,9, 2,2,'Kemasan' , sub_header)
        worksheet.write(8, 3, 'Qty', sub_header)
        worksheet.write(9, 3, 'Sales Terinvoice', sub_header)
        worksheet.write(8, 4, 'HPP', sub_header)
        worksheet.write(9, 4, 'Excl', sub_header)
        worksheet.write(8, 5, 'Harga', sub_header)
        worksheet.write(9, 5, 'Produk', sub_header)

        worksheet.col(0).width = 70 * 30
        worksheet.col(1).width = 140 * 140
        worksheet.col(2).width = 70 * 70
        worksheet.col(3).width = 70 * 70
        worksheet.col(4).width = 70 * 70
        return worksheet


    def create_excel_value(self,worksheet):
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
        row = 9
        kemasan = ''
        qty_sum = 0.00
        total_sum = 0.00
        for product in product_ids:
            search_inv = self.env['account.move.line'].search([('journal_id', '=', 1),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
            if search_inv:
                if 'galon 4 liter' in product.name.lower() or 'pail' in product.name.lower():
                    if 'galon 4 liter' in product.name.lower():
                        kemasan = '4 Liter'
                    elif 'pail' in product.name.lower():
                        kemasan = 'Pail'
                    row +=1
                    seq +=1
                    qty_sum = sum(line.quantity for line in search_inv)
                    total_sum = sum(line2.price_subtotal for line2 in search_inv)
                    worksheet.write(row,0,seq,text_center)
                    worksheet.write(row,1,product.name,text_left)
                    worksheet.write(row,2,kemasan,text_center)
                    worksheet.write(row,3,qty_sum,text_center)
                    worksheet.write(row,4,search_inv.product_id.standard_price,text_right_accounting)
                    worksheet.write(row,5,total_sum,text_right_accounting)
        row +=3
        worksheet.write_merge(row,row, 0,2,'Kemasan  Karton' , text_left)
        for product in product_ids:
            search_inv = self.env['account.move.line'].search([('journal_id', '=', 1),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
            if search_inv:
                if 'galon 4 liter' not in product.name.lower() and 'pail' not in product.name.lower():
                    row +=1
                    seq2 +=1
                    qty_sum = sum(line.quantity for line in search_inv)
                    total_sum = sum(line2.price_subtotal for line2 in search_inv)
                    worksheet.write(row,0,seq2,text_center)
                    worksheet.write(row,1,product.name,text_left)
                    worksheet.write(row,2,'',text_center)
                    worksheet.write(row,3,qty_sum,text_center)
                    worksheet.write(row,4,search_inv.product_id.standard_price,text_right_accounting)
                    worksheet.write(row,5,total_sum,text_right_accounting)
        return worksheet


################## Body Excel report 2 ##################

    def create_excel_header2(self,worksheet2):
        main_header_style = easyxf('align: horiz left,vert center;'
                              'font:bold True,height 250;')
        sub_header = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                              'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin')

        worksheet2.write_merge(0, 3, 0, 1, self.env.user.company_id.name or '', main_header_style)
        
        worksheet2.write_merge(5,5, 0,1,'Laporan Penjualan per Customer' , sub_header)
        worksheet2.write_merge(6,6, 0,1,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B %Y'), sub_header)

        worksheet2.write(8, 0, 'Nama Customer', sub_header)
        worksheet2.write(8, 1, 'Sales Per Customer (Rp)', sub_header)

        worksheet2.col(0).width = 140 * 140
        worksheet2.col(1).width = 70 * 70
        return worksheet2


    def create_excel_value2(self,worksheet2):
        text_left = easyxf('font:height 200;align:vert center,horiz left;' 'borders: top thin,bottom thin,left thin, right thin')
        text_right_accounting = easyxf('font:height 200;align:vert center,horiz right;' 'borders: top thin,bottom thin,left thin, right thin',num_format_str='#,##0.00_);(#,##0.00)')
        text_center = easyxf('font:height 200;align:vert center,horiz center;' 'borders: top thin,bottom thin,left thin, right thin')

        if self.partner_by == 'all':
        	partner_ids = self.env['res.partner'].search([],order='name asc')
        else:
        	if not self.partner_ids:
        		raise UserError(_('No Customer selected'))
        	partner_ids = self.partner_ids
        row = 8
        total_sum = 0.00
        for partner in partner_ids:
            search_inv = self.env['account.move.line'].search([('journal_id', '=', 1),('move_id.partner_id', '=', partner.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
            if search_inv:
                row +=1
                total_sum = sum(line.price_subtotal for line in search_inv)
                worksheet2.write(row,0,partner.name,text_left)
                worksheet2.write(row,1,total_sum,text_right_accounting)
        return worksheet2


################## Body Excel report 3 ##################

    def create_excel_header3(self,worksheet3):
        main_header_style = easyxf('align: horiz left,vert center;'
                              'font:bold True,height 250;')
        sub_header = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                              'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin')

        worksheet3.write_merge(0, 3, 0, 2, self.env.user.company_id.name or '', main_header_style)
        
        worksheet3.write_merge(5,5, 0,5,'Laporan Penjualan per UOM' , sub_header)
        worksheet3.write_merge(6,6, 0,5,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B %Y'), sub_header)

        worksheet3.write_merge(8,10, 0,0,'Nama Customer' , sub_header)
        worksheet3.write_merge(8,8, 1,5,'Sales per UOM' , sub_header)
        worksheet3.write_merge(9,9, 1,5,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B') , sub_header)
        worksheet3.write(10, 1, 'Pail', sub_header)
        worksheet3.write(10, 2, 'Galon', sub_header)
        worksheet3.write(10, 3, '1 Liter', sub_header)
        worksheet3.write(10, 4, '500 ML', sub_header)
        worksheet3.write(10, 5, 'Pouch', sub_header)

        col=5
        col_head=5
        worksheet3.col(0).width = 140 * 140
        worksheet3.col(1).width = 45 * 45
        worksheet3.col(2).width = 45 * 45
        worksheet3.col(3).width = 45 * 45
        worksheet3.col(4).width = 45 * 45
        worksheet3.col(5).width = 45 * 45
        return worksheet3


    def create_excel_value3(self,worksheet3):
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
            tot_pail = 0
            tot_gal = 0
            tot_lit = 0
            tot_500 = 0
            tot_pouch = 0
            search_inv = self.env['account.move.line'].search([('journal_id', '=', 1),('move_id.partner_id', '=', partner.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
            if search_inv:
                row +=1
                for data_prod in search_inv:
                    if isinstance(data_prod.product_id.name, str):
                        if 'pail' in data_prod.product_id.name.lower():
                            tot_pail += data_prod.quantity
                        elif 'galon 4 liter' in data_prod.product_id.name.lower():
                            tot_gal += data_prod.quantity
                        elif 'galon 1 liter' in data_prod.product_id.name.lower():
                            tot_lit += data_prod.quantity
                        elif '500 ml' in data_prod.product_id.name.lower():
                            tot_500 += data_prod.quantity
                        else:
                            tot_pouch += data_prod.quantity
                worksheet3.write(row,0,partner.name,text_left)
                worksheet3.write(row,1,tot_pail,text_center)
                worksheet3.write(row,2,tot_gal,text_center)
                worksheet3.write(row,3,tot_lit,text_center)
                worksheet3.write(row,4,tot_500,text_center)
                worksheet3.write(row,5,tot_pouch,text_center)
#########################################################
#########################################################


#########################################################
################## Body Excel report 1 ##################

    def create_excel_header_periodic(self,worksheet):
        main_header_style = easyxf('align: horiz left,vert center;'
                              'font:bold True,height 250;')
        sub_header = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                              'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin')

        worksheet.write_merge(0, 3, 0, 2, self.env.user.company_id.name or '', main_header_style)
        
        worksheet.write_merge(5,5, 0,5,'Laporan Penjualan per Item' , sub_header)
        worksheet.write_merge(6,6, 0,5,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B %Y'), sub_header)

        worksheet.write_merge(8,9, 0,0,'No.' , sub_header)
        worksheet.write_merge(8,9, 1,1,'Nama Barang' , sub_header)
        worksheet.write_merge(8,9, 2,2,'Kemasan' , sub_header)
        worksheet.write(8, 3, 'Qty', sub_header)
        worksheet.write(9, 3, 'Sales Terinvoice', sub_header)
        worksheet.write(8, 4, 'HPP', sub_header)
        worksheet.write(9, 4, 'Excl', sub_header)
        worksheet.write(8, 5, 'Harga', sub_header)
        worksheet.write(9, 5, 'Produk', sub_header)

        worksheet.col(0).width = 70 * 30
        worksheet.col(1).width = 140 * 140
        worksheet.col(2).width = 70 * 70
        worksheet.col(3).width = 70 * 70
        worksheet.col(4).width = 70 * 70
        return worksheet


    def create_excel_value_periodic(self,worksheet):
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
        row = 9
        kemasan = ''
        qty_sum = 0.00
        total_sum = 0.00
        for product in product_ids:
            search_inv = self.env['account.move.line'].search([('journal_id', '=', 1),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
            if search_inv:
                if 'galon 4 liter' in product.name.lower() or 'pail' in product.name.lower():
                    if 'galon 4 liter' in product.name.lower():
                        kemasan = '4 Liter'
                    elif 'pail' in product.name.lower():
                        kemasan = 'Pail'
                    row +=1
                    seq +=1
                    qty_sum = sum(line.quantity for line in search_inv)
                    total_sum = sum(line2.price_subtotal for line2 in search_inv)
                    worksheet.write(row,0,seq,text_center)
                    worksheet.write(row,1,product.name,text_left)
                    worksheet.write(row,2,kemasan,text_center)
                    worksheet.write(row,3,qty_sum,text_center)
                    worksheet.write(row,4,search_inv.product_id.standard_price,text_right_accounting)
                    worksheet.write(row,5,total_sum,text_right_accounting)
        row +=3
        worksheet.write_merge(row,row, 0,2,'Kemasan  Karton' , text_left)
        for product in product_ids:
            search_inv = self.env['account.move.line'].search([('journal_id', '=', 1),('product_id', '=', product.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
            if search_inv:
                if 'galon 4 liter' not in product.name.lower() and 'pail' not in product.name.lower():
                    row +=1
                    seq2 +=1
                    qty_sum = sum(line.quantity for line in search_inv)
                    total_sum = sum(line2.price_subtotal for line2 in search_inv)
                    worksheet.write(row,0,seq2,text_center)
                    worksheet.write(row,1,product.name,text_left)
                    worksheet.write(row,2,'',text_center)
                    worksheet.write(row,3,qty_sum,text_center)
                    worksheet.write(row,4,search_inv.product_id.standard_price,text_right_accounting)
                    worksheet.write(row,5,total_sum,text_right_accounting)
        return worksheet


################## Body Excel report 2 ##################

    def create_excel_header_periodic2(self,worksheet2):
        main_header_style = easyxf('align: horiz left,vert center;'
                              'font:bold True,height 250;')
        sub_header = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                              'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin')

        worksheet2.write_merge(0, 3, 0, 1, self.env.user.company_id.name or '', main_header_style)
        
        worksheet2.write_merge(5,5, 0,1,'Laporan Penjualan per Customer' , sub_header)
        worksheet2.write_merge(6,6, 0,1,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B %Y'), sub_header)

        worksheet2.write(8, 0, 'Nama Customer', sub_header)
        worksheet2.write(8, 1, 'Sales Per Customer (Rp)', sub_header)

        worksheet2.col(0).width = 140 * 140
        worksheet2.col(1).width = 70 * 70
        return worksheet2


    def create_excel_value_periodic2(self,worksheet2):
        text_left = easyxf('font:height 200;align:vert center,horiz left;' 'borders: top thin,bottom thin,left thin, right thin')
        text_right_accounting = easyxf('font:height 200;align:vert center,horiz right;' 'borders: top thin,bottom thin,left thin, right thin',num_format_str='#,##0.00_);(#,##0.00)')
        text_center = easyxf('font:height 200;align:vert center,horiz center;' 'borders: top thin,bottom thin,left thin, right thin')

        if self.partner_by == 'all':
        	partner_ids = self.env['res.partner'].search([],order='name asc')
        else:
        	if not self.partner_ids:
        		raise UserError(_('No Customer selected'))
        	partner_ids = self.partner_ids
        row = 8
        total_sum = 0.00
        for partner in partner_ids:
            search_inv = self.env['account.move.line'].search([('journal_id', '=', 1),('move_id.partner_id', '=', partner.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
            if search_inv:
                row +=1
                total_sum = sum(line.price_subtotal for line in search_inv)
                worksheet2.write(row,0,partner.name,text_left)
                worksheet2.write(row,1,total_sum,text_right_accounting)
        return worksheet2


################## Body Excel report 3 ##################

    def create_excel_header_periodic3(self,worksheet3):
        main_header_style = easyxf('align: horiz left,vert center;'
                              'font:bold True,height 250;')
        sub_header = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                              'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin')

        worksheet3.write_merge(0, 3, 0, 2, self.env.user.company_id.name or '', main_header_style)
        
        worksheet3.write_merge(5,5, 0,5,'Laporan Penjualan per UOM' , sub_header)
        worksheet3.write_merge(6,6, 0,5,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B %Y'), sub_header)

        worksheet3.write_merge(8,10, 0,0,'Nama Customer' , sub_header)
        worksheet3.write_merge(8,8, 1,5,'Sales per UOM' , sub_header)
        worksheet3.write_merge(9,9, 1,5,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B') , sub_header)
        worksheet3.write(10, 1, 'Pail', sub_header)
        worksheet3.write(10, 2, 'Galon', sub_header)
        worksheet3.write(10, 3, '1 Liter', sub_header)
        worksheet3.write(10, 4, '500 ML', sub_header)
        worksheet3.write(10, 5, 'Pouch', sub_header)

        col=5
        col_head=5
        worksheet3.col(0).width = 140 * 140
        worksheet3.col(1).width = 45 * 45
        worksheet3.col(2).width = 45 * 45
        worksheet3.col(3).width = 45 * 45
        worksheet3.col(4).width = 45 * 45
        worksheet3.col(5).width = 45 * 45
        return worksheet3


    def create_excel_value_periodic3(self,worksheet3):
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
            tot_pail = 0
            tot_gal = 0
            tot_lit = 0
            tot_500 = 0
            tot_pouch = 0
            search_inv = self.env['account.move.line'].search([('journal_id', '=', 1),('move_id.partner_id', '=', partner.id),('display_type', '=', 'product'),('move_id.state', '=', 'posted'),('move_id.invoice_date', '>=', self.start_date),('move_id.invoice_date', '<=', self.end_date)])
            if search_inv:
                row +=1
                for data_prod in search_inv:
                    if isinstance(data_prod.product_id.name, str):
                        if 'pail' in data_prod.product_id.name.lower():
                            tot_pail += data_prod.quantity
                        elif 'galon 4 liter' in data_prod.product_id.name.lower():
                            tot_gal += data_prod.quantity
                        elif 'galon 1 liter' in data_prod.product_id.name.lower():
                            tot_lit += data_prod.quantity
                        elif '500 ml' in data_prod.product_id.name.lower():
                            tot_500 += data_prod.quantity
                        else:
                            tot_pouch += data_prod.quantity
                worksheet3.write(row,0,partner.name,text_left)
                worksheet3.write(row,1,tot_pail,text_center)
                worksheet3.write(row,2,tot_gal,text_center)
                worksheet3.write(row,3,tot_lit,text_center)
                worksheet3.write(row,4,tot_500,text_center)
                worksheet3.write(row,5,tot_pouch,text_center)
#########################################################
#########################################################


    def export_excel(self):
        if self.end_date < self.start_date:
            raise ValidationError(_('End Date must be greater than Start Date'))
        workbook = xlwt.Workbook()
        filename = 'Report_sales_SNB.xls'

        #report 1
        worksheet = workbook.add_sheet('Penjualan Per Item')
        for c in range(0, 100):
            worksheet.col(c).width = 140 * 30
        for c in range(0, 100):
            if c <= 7:
                worksheet.row(c).height = 250
            else:
                worksheet.row(c).height = 350
        worksheet = self.create_excel_header(worksheet)
        worksheet = self.create_excel_value(worksheet)

        #report 2
        worksheet2 = workbook.add_sheet('Penjualan Per Customer')
        for c in range(0, 100):
            worksheet2.col(c).width = 140 * 30
        for c in range(0, 100):
            if c <= 7:
                worksheet2.row(c).height = 250
            else:
                worksheet2.row(c).height = 350
        worksheet2 = self.create_excel_header2(worksheet2)
        worksheet2 = self.create_excel_value2(worksheet2)

        #report 3
        worksheet3 = workbook.add_sheet('Penjualan Per UOM')
        for c in range(0, 100):
            worksheet3.col(c).width = 140 * 30
        for c in range(0, 100):
            if c <= 7:
                worksheet3.row(c).height = 250
            else:
                worksheet3.row(c).height = 350
        worksheet3 = self.create_excel_header3(worksheet3)
        worksheet3 = self.create_excel_value3(worksheet3)

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
