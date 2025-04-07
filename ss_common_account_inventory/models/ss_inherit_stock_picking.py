
from collections import defaultdict
from datetime import timedelta
from operator import itemgetter

from odoo import _, api, Command, fields, models
from odoo.exceptions import UserError
from odoo.osv import expression
from odoo.tools.float_utils import float_compare, float_is_zero, float_round
from odoo.tools.misc import clean_context, OrderedSet, groupby

class ssStockPicking(models.Model):
    _inherit = 'stock.picking'
    
    coa_revisi_id = fields.Many2one('account.account', string='account',)
    type_transfer_stock = fields.Selection(
        selection=[
            ('adjustment', 'Adjustment'),
            ('sales', 'Sales'),
            ('purchase', 'Purchase'),
            ('manufacture', 'Manufacture'),
            ],
        string='Transfer Type',related='picking_type_id.type_transfer_stock')
    active = fields.Boolean(default=True, help="Set active to false to hide the Journal without removing it.")

    def button_validate(self):

    def set_coa_for_adjustment(self):
        for doc in self:
            for line_stock in doc.move_ids_without_package:
                coa_stock_out = line_stock.product_id.categ_id.property_stock_account_output_categ_id.id
                coa_stock_in = line_stock.product_id.categ_id.property_stock_account_input_categ_id.id
                search_journal = self.env['account.move'].search([('stock_move_id', '=', line_stock.id)])
                if search_journal:
                    for line_journal in search_journal.line_ids:
                        if line_journal.account_id.id == coa_stock_out or line_journal.account_id.id == coa_stock_in :
                            if doc.type_transfer_stock == 'adjustment':
                                line_journal.account_id = doc.coa_revisi_id.id
                            elif doc.type_transfer_stock == 'sales':
                                line_journal.account_id = doc.picking_type_id.coa_revisi_id.id
                            elif doc.type_transfer_stock == 'manufacture':
                                line_journal.account_id = doc.picking_type_id.coa_revisi_id.id

    def set_active_inactive(self):
        self.active = not self.active

class ssPickingType(models.Model):
    _inherit = "stock.picking.type"

    coa_revisi_id = fields.Many2one('account.account', string='account',)
    type_transfer_stock = fields.Selection(
        selection=[
            ('adjustment', 'Adjustment'),
            ('sales', 'Sales'),
            ('purchase', 'Purchase'),
            ('manufacture', 'Manufacture'),
            ],
        string='Transfer Type')

# class StockMove(models.Model):
#     _inherit = "stock.move"

#     def _action_assign(self, force_qty=False):
#         res = super(StockMove, self)._action_assign()
#         for line in 
  