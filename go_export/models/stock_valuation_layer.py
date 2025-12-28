from odoo import fields, models, api


class StockValuationLayer(models.Model):
    _inherit = 'stock.valuation.layer'

    cost_line_id = fields.Many2one('stock.landed.cost.lines', string='Landed Cost Line',store=True)