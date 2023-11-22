from odoo import models
from datetime import datetime, timedelta
import base64
import io
import xlsxwriter

class FleetFieldsUpdate(models.Model):
    _name = "stock.export"
    _description = "stock move line export"

    def export_stock_move_lines_xlsx(self):
        # Calcola la data di tre giorni fa
        three_days_ago = datetime.now() - timedelta(days=3)
        today = datetime.now()

        # Cerca i record degli ultimi tre giorni in stock.move.line
        stock_moves = self.env['stock.move.line'].search([('date', '>=', three_days_ago.strftime('%Y-%m-%d')), ('branch_id', '=', 1)])

        # Costruisci il contenuto del file XLSX in memoria
        xlsx_content = io.BytesIO()
        workbook = xlsxwriter.Workbook(xlsx_content)
        worksheet = workbook.add_worksheet()

        headers = ['Id', 'Batch', 'Branch', 'Company', 'Created by', 'Date', 'Destination location', 'Destination Package', 'Done', 'From', 'From Owner', 'Last update by', 'Last update on', 'Lot/Serial number', 'Product', 'Reference', 'Source location', 'Source package', 'Status', 'To', 'Unit of misure']

        # Aggiungi gli header alla prima riga
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        row = 1  # Inizia dalla seconda riga per i dati
        for stock_move in stock_moves:
            id = self.env['stock.move.line'].browse(stock_move.id)
            record_id = int(str(id).split('(')[1].split(',')[0])
            
            #batch_transfer = self.env['stock.picking.batch'].browse(stock_move.batch_id.id)
            location = self.env['stock.location'].browse(stock_move.picking_location_dest_id.id)
            company = self.env['res.company'].browse(stock_move.company_id.id)
            branch = self.env['res.branch'].browse(stock_move.branch_id.id)
            destination_packages = self.env['stock.quant.package'].browse(stock_move.result_package_id.id)
            create_uid = self.env['res.users'].browse(stock_move.create_uid.id)
            write_uid = self.env['res.users'].browse(stock_move.write_uid.id)
            from_name = self.env['stock.location'].browse(stock_move.location_id.id)
            from_owner = self.env['res.partner'].browse(stock_move.owner_id.id)
            source_location = self.env['stock.location'].browse(stock_move.picking_location_id.id)
            source_package = self.env['stock.quant.package'].browse(stock_move.package_id.id)
            to = self.env['stock.location'].browse(stock_move.location_dest_id.id)
            unit = self.env['uom.uom'].browse(stock_move.product_uom_id.id)
            lot = self.env['stock.lot'].browse(stock_move.lot_id.id)
            product = self.env['product.product'].browse(stock_move.product_id.id)

            
            worksheet.write(row, 0, str(record_id))
            worksheet.write(row, 1, str(stock_move.batch_id.id))
            worksheet.write(row, 2, str(branch.name))
            worksheet.write(row, 3, str(company.name))
            worksheet.write(row, 4, str(create_uid.name))
            worksheet.write(row, 5, str(stock_move.date))
            worksheet.write(row, 6, str(location.display_name))
            worksheet.write(row, 7, str(destination_packages.name))
            worksheet.write(row, 8, str(stock_move.qty_done))
            worksheet.write(row, 9, str(from_name.display_name))
            worksheet.write(row, 10, str(from_owner.name))
            worksheet.write(row, 11, str(write_uid.name))
            worksheet.write(row, 12, str(stock_move.write_date))
            worksheet.write(row, 13, str(lot.name))
            worksheet.write(row, 14, str(product.display_name))
            worksheet.write(row, 15, str(stock_move.reference))
            worksheet.write(row, 16, str(source_location.display_name))
            worksheet.write(row, 17, str(source_package.name))
            worksheet.write(row, 18, str(stock_move.state))
            worksheet.write(row, 19, str(to.display_name))
            worksheet.write(row, 20, str(unit.name))


            row += 1  # Passa alla riga successiva per il prossimo stock_move

        workbook.close()

        # Imposta l'allegato in Odoo come file XLSX
        xlsx_content.seek(0)
        attachment_values = {
            'name': 'stock_move.xlsx',
            'datas': base64.encodebytes(xlsx_content.getvalue()).decode(),
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
            'body_html': '<p>In allegato file .XLSX con le movimentazioni degli ultimi 3 giorni.</p>',
            'attachment_ids': [(4, attachment.id)],  # Aggiungi l'allegato all'email
        }

        # Crea e invia l'email utilizzando il metodo create di mail.mail
        mail = self.env['mail.mail'].sudo().create(mail_values)
        mail.send()




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
 