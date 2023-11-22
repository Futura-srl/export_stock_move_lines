from odoo import models
from datetime import datetime, timedelta
import base64
import io, logging
import xlsxwriter

_logger = logging.getLogger(__name__)

class StockMoveLineExport(models.Model):
    _name = "stock.export"
    _description = "stock move line export"

    def export_inventory_xlsx(self):

        # Cerca i record degli ultimi tre giorni in stock.move.line
        stock_inventory = self.env['stock.quant'].search(['|', ('location_id', 'ilike', "TITO/IN"), ('location_id', 'ilike', "TITO/ST")])

        
        # Costruisci il contenuto del file XLSX in memoria
        xlsx_content = io.BytesIO()
        workbook = xlsxwriter.Workbook(xlsx_content)
        worksheet = workbook.add_worksheet()

        headers = ['Articolo', "Lotto", "Hu", "Qty"]

        # Aggiungi gli header alla prima riga
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        row = 1  # Inizia dalla seconda riga per i dati
        for record in stock_inventory:

            _logger.info(record.product_id)
            _logger.info(record.lot_id)
            _logger.info(record.package_id)
            _logger.info(record.quantity)

            worksheet.write(row, 0, str(record.product_id))
            worksheet.write(row, 1, str(record.lot_id))
            worksheet.write(row, 2, str(record.package_id))
            worksheet.write(row, 3, str(record.quantity))

            row += 1  # Passa alla riga successiva per il prossimo stock_move

        workbook.close()

        # Imposta l'allegato in Odoo come file XLSX
        xlsx_content.seek(0)
        attachment_values = {
            'name': 'stock_inventory.xlsx',
            'datas': base64.encodebytes(xlsx_content.getvalue()).decode(),
            'res_model': self._name,
            'res_id': self.id,
            'type': 'binary',
        }
        attachment = self.env['ir.attachment'].create(attachment_values)

        # Invia l'email con l'allegato
        mail_values = {
            'subject': 'Inventario Ferrero Tito Scalo',
            'email_from': 'noreply@futurasl.com',
            'email_to': 'dati+stocktito@svcfutura.cloud',
            'body_html': "<p>In allegato l'inventario del magazzino Ferrero di Tito Scalo (PZ).</p>",
            'attachment_ids': [(4, attachment.id)],  # Aggiungi l'allegato all'email
        }

        # Crea e invia l'email utilizzando il metodo create di mail.mail
        mail = self.env['mail.mail'].sudo().create(mail_values)
        mail.send()
