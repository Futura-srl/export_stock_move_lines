from odoo import models
from datetime import datetime, timedelta
import base64
import io
import csv

class StockMoveLineUpdate(models.Model):
    _inherit = "stock.move.line"


    def export_stock_move_lines_csv(self):
        # Calcola la data di tre giorni fa
        three_days_ago = datetime.now() - timedelta(days=3)
        today = str(datetime.now())

        # Cerca i record degli ultimi tre giorni in stock.move.line
        stock_moves = self.env['stock.move.line'].search([('date', '>=', three_days_ago.strftime('%Y-%m-%d')), ('branch_id', '=', 1)])

        # Costruisci il contenuto del file CSV in memoria
        csv_content = io.StringIO()
        csv_writer = csv.writer(csv_content)
        csv_writer.writerow(['Id', 'Batch', 'Branch', 'Company', 'Created by', 'Date', 'Destination location', 'Destination Package', 'Done', 'From', 'From Owner', 'Last update by', 'Last update on', 'Lot/Serial number', 'Product', 'Reference', 'Source location', 'Source package', 'Status', 'To', 'Unit of misure'])  # Intestazione del CSV

        for stock_move in stock_moves:
            id = self.env['stock.move.line'].browse(stock_move.id)
            record_id = int(str(id).split('(')[1].split(',')[0])
            
            batch_transfer = self.env['stock.picking.batch'].browse(stock_move.batch_id.id)
            location = self.env['stock.location'].browse(stock_move.picking_location_dest_id.id)
            company = self.env['res.company'].browse(stock_move.company_id.id)
            branch = self.env['res.branch'].browse(stock_move.branch_id.id)
            created_by = self.env['res.branch'].browse(stock_move.create_uid.id)
            destination_packages = self.env['stock.move.line'].browse(stock_move.result_package_id.id)
            from_name = self.env['stock.location'].browse(stock_move.location_id.id)
            from_owner = self.env['res.partner'].browse(stock_move.owner_id.id)
            write = self.env['res.users'].browse(stock_move.write_uid.id)
            source_location = self.env['stock.location'].browse(stock_move.picking_location_id.id)
            source_package = self.env['stock.location'].browse(stock_move.package_id.id)
            to = self.env['stock.location'].browse(stock_move.location_dest_id.id)
            unit = self.env['uom.uom'].browse(stock_move.product_uom_id.id)
            lot = self.env['stock.lot'].browse(stock_move.lot_id.id)
            product = self.env['product.product'].browse(stock_move.product_id.id)
            
            csv_writer.writerow([record_id, batch_transfer, branch.name, company, created_by, stock_move.date, location.display_name, destination_packages, stock_move.qty_done, from_name.display_name, from_owner.name, write.name, stock_move.write_date, lot.name, product.display_name, stock_move.reference, source_location.display_name, source_package, stock_move.state, to.display_name, unit.name])

        # Crea l'allegato in Odoo
        attachment_values = {
            'name': 'stock_move.csv',
            'datas': base64.encodebytes(csv_content.getvalue().encode()).decode(),
            'res_model': self._name,
            'res_id': self.id,
            'type': 'binary',
        }
        attachment = self.env['ir.attachment'].create(attachment_values)

        # Invia l'email con l'allegato
        mail_values = {
            'subject': 'Movimentazioni stock Tito Scalo dal ' + three_days_ago.strftime('%d-%m-%Y') + ' al ' + today.strftime('%d-%m-%Y'),
            'email_from': 'noreply@futurasl.com',
            'email_to': 'luca.cocozza@futurasl.com',
            'body_html': '<p>Questa è una email con il file CSV allegato.</p>',
            'attachment_ids': [(4, attachment.id)],  # Aggiungi l'allegato all'email
        }

        # Crea e invia l'email utilizzando il metodo create di mail.mail
        mail = self.env['mail.mail'].sudo().create(mail_values)
        mail.send()
 