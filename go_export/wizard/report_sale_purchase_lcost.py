from odoo import fields, models, api
from io import BytesIO
import xlsxwriter
import base64
import json
from datetime import datetime
class Report(models.TransientModel):
    _name = 'report.sale.purchase.lcost'
    _description = 'Sale Purchase Landed Cost Report Wizard'

    partner_ids = fields.Many2many(
        'res.partner',
        string='Customers',
    )
    date = fields.Date(string='Start Date')


    def action_xlsx_report(self):
        domain = [('state', 'in', ['sale', 'done'])]
        if self.partner_ids:
            domain += [('partner_id', 'in', self.partner_ids.ids)]
        if self.date:
            domain += [('date_order', '>=', self.date)]

        report_data = []
        all_landed_cost_names = set()
        # print("domain ====== ", domain)
        orders = self.env['sale.order'].search_read(domain,['partner_id','picking_ids','order_line','date_order'])
        for order in orders:
            customer_name= order['partner_id'][1]
            order_date = order['date_order'].date()
            order['picking_ids'] = self.env['stock.picking'].browse(order['picking_ids'])
            for pick in order['picking_ids']:
                moves_with_lots = pick.move_ids.filtered(lambda p: p.lot_ids)
                for move in moves_with_lots:
                    for lot in move.lot_ids:
                        move_valuation_lot = self.env['stock.valuation.layer'].search([('lot_id', '=', lot.id)])
                        # print("len", len(move_purchase))
                        serial_data = {
                            'customer_name': customer_name,
                            'order_date': order_date,
                            'product_name': move.product_id.name,
                            'lot_name': lot.name,
                            'purchase_cost': 0,
                            'landed_costs': {},
                            'sale_price': 0,
                            'analytic_accounts': '',
                        }
                        for po in move_valuation_lot.filtered(lambda p: p.stock_move_id.purchase_line_id and not p.stock_landed_cost_id):
                            serial_data['purchase_cost'] += abs(po.value)
                        for landed_cost in move_valuation_lot.filtered(lambda p: p.stock_landed_cost_id and p.cost_line_id):
                            print("landed_cost.cost_line_id.name", landed_cost.cost_line_id.read())
                            all_landed_cost_names.add(landed_cost.cost_line_id.product_id.name)
                            if landed_cost.cost_line_id.name not in serial_data['landed_costs']:
                                serial_data['landed_costs'][landed_cost.cost_line_id.name] = 0
                            serial_data['landed_costs'][landed_cost.cost_line_id.name] += abs(landed_cost.value)

                        sale_price = 0
                        for sale_line in order['order_line']:
                            sale_lin_obj = self.env['sale.order.line'].browse(sale_line)
                            if sale_lin_obj.analytic_distribution:
                                account_names = ''
                                for account_id,percentage in sale_lin_obj.analytic_distribution.items():
                                    analytic_account = self.env['account.analytic.account'].browse(int(account_id))
                                    print(f"{account_id}, Name: {analytic_account.name}, Percentage: {percentage}%")
                                    account_names += f"{analytic_account.name} | {percentage} %, "

                                serial_data['analytic_accounts'] = account_names.rstrip(', ')

                            if sale_lin_obj.product_id.id == move.product_id.id:
                                sale_price= sale_lin_obj.price_total
                        serial_data['sale_price'] = sale_price
                        report_data.append(serial_data)
                        print("purchase data", serial_data)

        print("all_landed_cost_names", all_landed_cost_names)
        sorted_landed_cost_names = sorted(list(all_landed_cost_names))
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Sale Purchase Landed Cost Report')
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        currency_format = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        percentage_format = workbook.add_format({'num_format': '0.00%', 'border': 1})
        text_format = workbook.add_format({'border': 1})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1})
        print("sorted_landed_cost_names", sorted_landed_cost_names)
        headers = ['Customer Name', 'Order Date', 'Product Name', 'Serial Number','Analytic Account', 'Purchase Cost'] + sorted_landed_cost_names + ['Total Landed Cost', 'Sale Price', 'Gross Profit', 'Gross Profit %']
        for col , header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        row = 1
        for serial in report_data:
            col = 0
            worksheet.write(row, col, serial['customer_name'], text_format)
            col += 1
            worksheet.write(row, col, serial['order_date'], date_format)
            col += 1
            worksheet.write(row, col, serial['product_name'], text_format)
            col += 1
            worksheet.write(row, col, serial['lot_name'], text_format)
            col += 1
            worksheet.write(row, col, serial['analytic_accounts'], text_format)
            col += 1
            worksheet.write(row, col, serial['purchase_cost'], currency_format)
            col += 1
            total_landed_cost = 0
            for lc_name in sorted_landed_cost_names:
                lc_value = serial['landed_costs'].get(lc_name, 0)
                worksheet.write(row, col, lc_value, currency_format)
                total_landed_cost += lc_value
                col += 1
            worksheet.write(row, col, total_landed_cost, currency_format)
            col += 1
            worksheet.write(row, col, serial['sale_price'], currency_format)
            col += 1
            # print("total_landed_cost", total_landed_cost)
            # print("serial['sale_price']", serial['sale_price'])
            # print("serial['purchase_cost']", serial['purchase_cost'])
            gross_profit = serial['sale_price'] - (serial['purchase_cost'] + total_landed_cost)
            # print("gross_profit", gross_profit)
            worksheet.write(row, col, gross_profit, currency_format)
            col += 1
            gross_profit_pct = gross_profit / serial['sale_price'] if serial['sale_price'] != 0 else 0
            worksheet.write(row, col, gross_profit_pct, percentage_format)
            row += 1

        for i,header in enumerate(headers):
            worksheet.set_column(i, i, 20)

        workbook.close()
        output.seek(0)
        filename = 'Serial_Profit_Report_%s.xlsx' % datetime.now().strftime('%Y%m%d_%H%M%S')
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'type': 'binary',
            'datas': base64.b64encode(output.read()),
            'store_fname': filename,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/%s?download=true' % attachment.id,
            'target': 'self',
        }