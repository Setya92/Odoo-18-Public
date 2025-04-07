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


class stock_management_report(models.TransientModel):
    _name = 'stock.management.report'
    _description = 'Report Stock In Out'
    
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

    stock_mode = fields.Selection(
        selection=[
            ('mo', 'Manufacture'),
            ('so', 'Sales'),
            ('po', 'Purchase'),
            ('ca', 'Cash Advance'),
            ],
        string='Stock Report Source')

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
                da_date = current_date.strftime("%B %Y")
                date_list.append({
                    'date': date_in_list,
                    'day':da_date,
                })

        return date_list


################## Body Excel report Manufacture ##################

    def create_excel_value_manufacture(self,worksheet,date_list):
        text_left = easyxf('font:height 200;align:vert center,horiz left;' 'borders: top thin,bottom thin,left thin, right thin')
        text_right_accounting = easyxf('font:height 200;align:vert center,horiz right;' 'borders: top thin,bottom thin,left thin, right thin',num_format_str='#,##0.00_);(#,##0.00)')
        main_header_style = easyxf('align: horiz left,vert center;'
                              'font:bold True,height 250;')
        text_date = easyxf('font:height 200;align:vert center,horiz center;' 'borders: top thin,bottom thin,left thin, right thin',num_format_str='dd-mm-yyyy')
        text_center = easyxf('font:height 200;align:vert center,horiz center;' 'borders: top thin,bottom thin,left thin, right thin')
        sub_header = easyxf('pattern: pattern solid, fore_colour gray25;align: horiz center,vert center;'
                              'font: colour black, bold True,height 150;' 'borders: top thin,bottom thin,left thin, right thin')
        text_right_accounting = easyxf('font:height 200;align:vert center,horiz right;' 'borders: top thin,bottom thin,left thin, right thin',num_format_str='#,##0.00_);(#,##0.00)')

        if self.product_by == 'all':
        	product_ids = self.env['product.product'].search([('sale_ok', '=', True)],order='name asc')
        else:
        	if not self.product_ids:
        		raise UserError(_('No Product selected'))
        	product_ids = self.product_ids

        worksheet.write_merge(0, 3, 0, 2, self.env.user.company_id.name or '', main_header_style)
        worksheet.write_merge(5,5, 0,5,'Laporan Manufacture' , sub_header)
        worksheet.write_merge(6,6, 0,5,'Periode '+self.start_date.strftime('%d')+' - '+self.end_date.strftime('%d %B %Y'), sub_header)
        row = 7
        row_head = 0
        sum_price = 0.0
        sum_price2 = 0.0
#####################<<<<<<<<<<<<<<<<<<<<<<<<<<
        if self.method_by == 'sumary':
            for product in product_ids:
                search_mo = self.env['mrp.production'].search([('product_id', '=', product.id),('date_planned_start', '>=', self.start_date),('date_planned_start', '<=', self.end_date)])
                if search_mo:
                    row +=1
                    worksheet.write(row,0,'Selected Item :',text_left)
                    worksheet.write(row,1,product.name,text_left)
                    row +=1
                    worksheet.write_merge(row,row+2, 0,0,'Nomor MO' , text_center)
                    worksheet.write_merge(row,row+2, 1,1,'Tanggal Request', text_center)
                    worksheet.write_merge(row,row+2, 2,2,'Kode Nama Bahan Baku' , text_center)
                    worksheet.write_merge(row,row, 3,5,'MO Created' , text_center)
                    worksheet.write_merge(row+1,row+2, 3,3,'Qty' , text_center)
                    worksheet.write_merge(row+1,row+1, 4,5,'Value' , text_center)
                    worksheet.write(row+2,4,'Per Unit',text_center)
                    worksheet.write(row+2,5,'Total',text_center)
                    worksheet.write_merge(row,row+2, 6,6,'Tanggal Finished', text_center)
                    worksheet.write_merge(row,row, 7,9,'MO Done' , text_center)
                    worksheet.write_merge(row+1,row+2, 7,7,'Qty' , text_center)
                    worksheet.write_merge(row+1,row+1,8,9,'Value' , text_center)
                    worksheet.write(row+2,8,'Per Unit',text_center)
                    worksheet.write(row+2,9,'Total',text_center)
                    row +=2
                    for line in search_mo:
                        row_head = row+1
                        for line_mo in line.move_raw_ids:
                            row +=1
                            sum_price = line_mo.product_uom_qty * line_mo.product_id.standard_price
                            worksheet.write(row,2,line_mo.product_id.name,text_left)
                            worksheet.write(row,3,line_mo.product_uom_qty,text_left)
                            worksheet.write(row,4,line_mo.product_id.standard_price,text_left)
                            worksheet.write(row,5,sum_price,text_left)
                            if line.state == 'done':
                                sum_price2 = line_mo.quantity_done * line_mo.product_id.standard_price
                                worksheet.write(row,7,line_mo.quantity_done,text_left)
                                worksheet.write(row,8,line_mo.product_id.standard_price,text_left)
                                worksheet.write(row,9,sum_price2,text_left)
                            else:
                                worksheet.write(row,7,0.0,text_left)
                                worksheet.write(row,8,0.0,text_left)
                                worksheet.write(row,9,0.0,text_left)
                        if line.state == 'done':
                            worksheet.write_merge(row_head,row, 0,0,line.name , text_center)
                            worksheet.write_merge(row_head,row, 1,1,line.date_planned_start , text_date)
                            worksheet.write_merge(row_head,row, 6,6,line.date_finished , text_date)
                        else:
                            worksheet.write_merge(row_head,row, 0,0,line.name , text_center)
                            worksheet.write_merge(row_head,row, 1,1,line.date_planned_start , text_date)
                            worksheet.write_merge(row_head,row, 6,6,' ', text_date)
                    row += 2
        return worksheet



    def export_excel(self):
        if self.end_date < self.start_date:
            raise ValidationError(_('End Date must be greater than Start Date'))
        workbook = xlwt.Workbook()
        date_list = self.get_date_list()
        if self.stock_mode == 'mo':
            filename = 'Report_manufacture_SNB.xls'
            worksheet = workbook.add_sheet('Manufacture Per Item')
            worksheet = self.create_excel_value_manufacture(worksheet,date_list)
        elif self.stock_mode == 'so':
            filename = 'Report_sales_SNB.xls'
            worksheet = workbook.add_sheet('Penjualan Per Item')
            worksheet = self.create_excel_header(worksheet,date_list)
        elif self.stock_mode == 'po':
            filename = 'Report_purchase_SNB.xls'
            worksheet = workbook.add_sheet('Pembelian Per Item')
            worksheet = self.create_excel_header(worksheet,date_list)
        elif self.stock_mode == 'ca':
            filename = 'Report_cashadvance_SNB.xls'
            worksheet = workbook.add_sheet('Cashadvance Data List')
            worksheet = self.create_excel_header(worksheet,date_list)


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
                'url': 'web/content/?model=stock.management.report&download=true&field=excel_file&id=%s&filename=%s' % (
                    active_id, filename),
                'target': 'new',
            }

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
