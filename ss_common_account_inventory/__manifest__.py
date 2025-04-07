# -*- coding: utf-8 -*-
##############################################################################
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) Sitaram Solutions (<https://sitaramsolutions.in/>).
#
#    For Module Support : info@sitaramsolutions.in  or Skype : contact.hiren1188
#
##############################################################################

{
    'name': 'SS Common Addons For Inventory add account',
    'version': '18.0.0.0',
    'category': 'Inventory',
    "license": "OPL-1",
    'summary': 'odoo apps for setting account on inventory transfer',
    'description': """
        
""",
    "price": 0,
    "currency": 'EUR',
    'author': 'Setya',
    'depends': ['base','stock','stock_account','account','mrp'],
    'data': [
             'security/ir.model.access.csv',
             'views/ss_inherit_stock_picking.xml',
             'wizard/stock_sales_report.xml',
             'wizard/stock_management_report.xml',
    ],
    'website':'',
    'installable': True,
    'auto_install': False,
    'live_test_url':'',
}

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
