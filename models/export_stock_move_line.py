from odoo import models
from datetime import datetime, timedelta
import base64
import io, logging
import xlsxwriter

_logger = logging.getLogger(__name__)

class StockMoveLineExport(models.Model):
    _name = "stock.export"
    _description = "stock move line export"

    def export_stock_move_lines_Ferrero_Tito_Scalo_xlsx(self):
        # Calcola la data di tre giorni fa
        three_days_ago = datetime.now() - timedelta(days=3)
        three_days_ago = three_days_ago.replace(hour=0, minute=0, second=0, microsecond=0)
        today = datetime.now()

        # Cerca i record degli ultimi tre giorni in stock.move.line
        stock_moves = self.env['stock.move.line'].search([('date', '>=', three_days_ago.strftime('%Y-%m-%d')), ('branch_id', '=', 1)])

        _logger.info(three_days_ago)
        
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
            'email_to': 'dati+stocktito@svcfutura.cloud',
            'body_html': '<p>In allegato file .XLSX con le movimentazioni degli ultimi 3 giorni.</p>',
            'attachment_ids': [(4, attachment.id)],  # Aggiungi l'allegato all'email
        }

        # Crea e invia l'email utilizzando il metodo create di mail.mail
        mail = self.env['mail.mail'].sudo().create(mail_values)
        mail.send()



    # Esempio di funzione per creare file CSV.
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
            'email_to': 'dati+stocktito@svcfutura.cloud',
            'body_html': '<p>Questa è una email con il file CSV allegato.</p>',
            'attachment_ids': [(4, attachment.id)],  # Aggiungi l'allegato all'email
        }

        # Crea e invia l'email utilizzando il metodo create di mail.mail
        mail = self.env['mail.mail'].sudo().create(mail_values)
        mail.send()



    # Funzione per esportare l'inventario di Tito Scalo
    def export_inventory_Ferrero_Tito_Scalo_xlsx(self):
        today = datetime.now().strftime('%d_%m_%Y')
        export_datetime = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

        # Cerca i record degli ultimi tre giorni in stock.move.line
        stock_inventory = self.env['stock.quant'].search(['|', ('location_id', 'ilike', "TITO/IN"), ('location_id', 'ilike', "TITO/ST")])

        
        # Costruisci il contenuto del file XLSX in memoria
        xlsx_content = io.BytesIO()
        workbook = xlsxwriter.Workbook(xlsx_content)
        worksheet = workbook.add_worksheet()

        headers = ["Articolo", "Descrizione", "Lotto", "Hu", "Quantità", "Estratto il"]

        # Aggiungi gli header alla prima riga
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        row = 1  # Inizia dalla seconda riga per i dati
        for record in stock_inventory:

            product = self.env['product.product'].browse(record.product_id.id)
            lot = self.env['stock.lot'].browse(record.lot_id.id)
            package = self.env['stock.quant.package'].browse(record.package_id.id)
            
            _logger.info(product.barcode)
            _logger.info(product.name)
            _logger.info(lot.name)
            _logger.info(package.name)
            _logger.info(record.quantity)
            _logger.info(export_datetime)

            worksheet.write(row, 0, str(product.barcode))
            worksheet.write(row, 1, str(product.name))
            worksheet.write(row, 2, str(lot.name))
            worksheet.write(row, 3, str(package.name))
            worksheet.write(row, 4, str(record.quantity))
            worksheet.write(row, 5, str(export_datetime))

            row += 1  # Passa alla riga successiva per il prossimo stock_move

        workbook.close()

        # Imposta l'allegato in Odoo come file XLSX
        xlsx_content.seek(0)
        attachment_values = {
            'name': 'Inventario_Tito_Scalo_' + today + '.xlsx',
            'datas': base64.encodebytes(xlsx_content.getvalue()).decode(),
            'res_model': self._name,
            'res_id': self.id,
            'type': 'binary',
        }
        attachment = self.env['ir.attachment'].create(attachment_values)

        # Invia l'email con l'allegato
        mail_values = {
            'subject': 'Inventario Ferrero Tito Scalo del ' + export_datetime,
            'email_from': 'noreply@futurasl.com',
            'email_to': 'antonio.croglia@ferrero.com',
            'email_cc': 'domenico.gala@futurasl.com, michele.divincenzo@futurasl.com, luca.cocozza@futurasl.com, fabio.righini@futurasl.com',
            'reply_to': 'domenico.gala@futurasl.com, michele.divincenzo@futurasl.com',
            'body_html': "<p>Salve,</br>in allegato copia inventario del magazzino Ferrero di Tito Scalo (PZ).</br></br>Futura S.p.A.</p>",
            'attachment_ids': [(4, attachment.id)],  # Aggiungi l'allegato all'email
        }

        # Crea e invia l'email utilizzando il metodo create di mail.mail
        mail = self.env['mail.mail'].sudo().create(mail_values)
        mail.send()
                # Funzione per esportare l'inventario di Tito Scalo
    def export_daily_inventory_Ferrero_Tito_Scalo_xlsx(self):
        today = datetime.now().strftime('%d_%m_%Y')
        export_datetime = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

        # Cerca i record degli ultimi tre giorni in stock.move.line
        stock_inventory = self.env['stock.quant'].search(['|', ('location_id', 'ilike', "TITO/IN"), ('location_id', 'ilike', "TITO/ST")])

        
        # Costruisci il contenuto del file XLSX in memoria
        xlsx_content = io.BytesIO()
        workbook = xlsxwriter.Workbook(xlsx_content)
        worksheet = workbook.add_worksheet()

        headers = ["Articolo", "Descrizione", "Lotto", "Hu", "Quantità", "Estratto il"]

        # Aggiungi gli header alla prima riga
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        row = 1  # Inizia dalla seconda riga per i dati
        for record in stock_inventory:

            product = self.env['product.product'].browse(record.product_id.id)
            lot = self.env['stock.lot'].browse(record.lot_id.id)
            package = self.env['stock.quant.package'].browse(record.package_id.id)
            
            _logger.info(product.barcode)
            _logger.info(product.name)
            _logger.info(lot.name)
            _logger.info(package.name)
            _logger.info(record.quantity)
            _logger.info(export_datetime)

            worksheet.write(row, 0, str(product.barcode))
            worksheet.write(row, 1, str(product.name))
            worksheet.write(row, 2, str(lot.name))
            worksheet.write(row, 3, str(package.name))
            worksheet.write(row, 4, str(record.quantity))
            worksheet.write(row, 5, str(export_datetime))

            row += 1  # Passa alla riga successiva per il prossimo stock_move

        workbook.close()

        # Imposta l'allegato in Odoo come file XLSX
        xlsx_content.seek(0)
        attachment_values = {
            'name': 'Inventario_Tito_Scalo_' + today + '.xlsx',
            'datas': base64.encodebytes(xlsx_content.getvalue()).decode(),
            'res_model': self._name,
            'res_id': self.id,
            'type': 'binary',
        }
        attachment = self.env['ir.attachment'].create(attachment_values)

        # Invia l'email con l'allegato
        mail_values = {
            'subject': 'Inventario Ferrero Tito Scalo del ' + export_datetime,
            'email_from': 'noreply@futurasl.com',
            'email_to': 'domenico.gala@futurasl.com, michele.divincenzo@futurasl.com',
            'email_cc': ', luca.cocozza@futurasl.com, fabio.righini@futurasl.com',
            'reply_to': 'domenico.gala@futurasl.com, michele.divincenzo@futurasl.com',
            'body_html': f"<p>Salve,</br>in allegato copia inventario del magazzino Ferrero di Tito Scalo (PZ) aggiornato al {export_datetime}.</br></br>Futura S.p.A.</p>",
            'attachment_ids': [(4, attachment.id)],  # Aggiungi l'allegato all'email
        }

        # Crea e invia l'email utilizzando il metodo create di mail.mail
        mail = self.env['mail.mail'].sudo().create(mail_values)
        mail.send()


    def check(self):
        def last_weekday_of_month(year, month, start_weekday, end_weekday):
            # Trova l'ultimo giorno del mese
            last_day_of_month = datetime(year, month+1, 1)
            last_day_of_month = last_day_of_month.replace(day=1) - timedelta(days=1)
    
            # Lista dei giorni candidati: 27, 28, 29, 30, 31 (se presente)
            candidate_days = [27, 28, 29, 30, 31] if month in [1, 3, 5, 7, 8, 10, 12] else [27, 28, 29, 30]
    
            # Itera all'indietro fino a trovare il giorno richiesto
            while last_day_of_month.day not in candidate_days or last_day_of_month.weekday() not in range(start_weekday, end_weekday + 1):
                last_day_of_month -= timedelta(days=1)
    
            return last_day_of_month
    
        # Definisci l'anno e il mese
        test_year = 2024
        test_month = 8  #  1 per gennaio, 2 per febbraio, ...
    
        # Definisci gli indici per lunedì e sabato
        start_weekday = 0  # Lunedì
        end_weekday = 5    # Sabato
    
        # Trova l'ultimo giorno del mese specificato che ricade tra lunedì e sabato
        last_day = last_weekday_of_month(test_year, test_month, start_weekday, end_weekday)
    
        # Verifica se last_day corrisponde al giorno odierno
        today = datetime.now().date()
    
        print("L'ultimo giorno del mese corrente che ricade tra lunedì e sabato è:", last_day.strftime('%Y-%m-%d'))
        if last_day.date() == today:
            print("L'ultimo giorno del mese corrente corrisponde al giorno odierno:", today.strftime('%Y-%m-%d'))
            return True
        else:
            print("L'ultimo giorno del mese corrente non corrisponde al giorno odierno:", today.strftime('%Y-%m-%d'))
            return False
        self.last_weekday_of_month()

    
    def export_pallet_in_fepz(self):
        result = self.check()
        if result == False:
            _logger.info("La data non è giusta")
        else:
            _logger.info("La data è giusta")
            
            # Ottieni la data odierna
            today_datetime = datetime.now()
        
            # Ottieni la data odierna in formato 'dd_mm_YYYY'
            today = today_datetime.strftime('%d_%m_%Y')
        
            # Ottieni il mese corrente e l'anno corrente
            current_month = today_datetime.strftime('%m')
            current_year = today_datetime.strftime('%Y')
        
            # Log delle informazioni
            _logger.info(today)
            _logger.info(current_month)
            _logger.info(current_year)
        
            # Ottieni il primo giorno del mese corrente
            first_date = today_datetime.replace(day=1,hour=00, minute=00, second=00)
        
            # Ottieni l'ultimo giorno del mese corrente
            # Imposta prima l'ultimo giorno al primo giorno del mese successivo
            # quindi sottrai un giorno per ottenere l'ultimo giorno del mese corrente
            last_date = today_datetime.replace(day=1, month=today_datetime.month + 1)
            last_date = (last_date - timedelta(days=1)).replace(hour=23, minute=59, second=59)
        
            # Log delle informazioni
            _logger.info(first_date)
            _logger.info(last_date)
    
            # Cerca i record degli ultimi tre giorni in stock.move.line
            stock_inventory = self.env['stock.move.line'].search([('location_id', 'ilike', "TITO/IN"), ('date', '>=', first_date), ('date', '<=', last_date)])
    
            
            # Costruisci il contenuto del file XLSX in memoria
            xlsx_content = io.BytesIO()
            workbook = xlsxwriter.Workbook(xlsx_content)
            worksheet = workbook.add_worksheet()
    
            headers = ["Articolo", "Descrizione", "Lotto", "Hu", "Quantità"]
    
            # Aggiungi gli header alla prima riga
            for col, header in enumerate(headers):
                worksheet.write(0, col, header)
    
            row = 1  # Inizia dalla seconda riga per i dati
            for record in stock_inventory:
    
                product = self.env['product.product'].browse(record.product_id.id)
                lot = self.env['stock.lot'].browse(record.lot_id.id)
                package = self.env['stock.quant.package'].browse(record.package_id.id)
                
                _logger.info(product.barcode)
                _logger.info(product.name)
                _logger.info(lot.name)
                _logger.info(package.name)
                _logger.info(record.qty_done)
    
                worksheet.write(row, 0, str(product.barcode))
                worksheet.write(row, 1, str(product.name))
                worksheet.write(row, 2, str(lot.name))
                worksheet.write(row, 3, str(package.name))
                worksheet.write(row, 4, str(record.qty_done))
    
                row += 1  # Passa alla riga successiva per il prossimo stock_move
    
            workbook.close()
    
            # Imposta l'allegato in Odoo come file XLSX
            xlsx_content.seek(0)
            attachment_values = {
                'name': 'Bancali ingressati nel mese.xlsx',
                'datas': base64.encodebytes(xlsx_content.getvalue()).decode(),
                'res_model': self._name,
                'res_id': self.id,
                'type': 'binary',
            }
            attachment = self.env['ir.attachment'].create(attachment_values)
    
            # Invia l'email con l'allegato
            mail_values = {
                'subject': 'Bancali ingressati nel mese ' + current_month + "/" + current_year,
                'email_from': 'noreply@futurasl.com',
                'email_to': 'antonio.croglia@ferrero.com',
                'email_cc': 'domenico.gala@futurasl.com, michele.divincenzo@futurasl.com, luca.cocozza@futurasl.com, fabio.righini@futurasl.com',
                'reply_to': 'domenico.gala@futurasl.com, michele.divincenzo@futurasl.com',
                'body_html': f"<p>Salve,</br>in allegato bancali ingressati nel mese corrente nel magazzino Ferrero di Tito Scalo (PZ) aggiornato al {last_date.strftime('%d/%m/%Y')}.</br></br>Futura S.p.A.</p>",
                'attachment_ids': [(4, attachment.id)],  # Aggiungi l'allegato all'email
            }
    
            # Crea e invia l'email utilizzando il metodo create di mail.mail
            mail = self.env['mail.mail'].sudo().create(mail_values)
            mail.send()
        
        
