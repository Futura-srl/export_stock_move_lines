from odoo import models
from datetime import datetime, timedelta
import base64
import io, logging
import xlsxwriter
import pytz


_logger = logging.getLogger(__name__)

class GtmsTripExport(models.Model):
    _name = "gtms.export"
    _description = "Gtms trip export"

    def export_gtms_trip_xlsx(self):
        # Calcola la data di tre giorni fa
        three_days_ago = datetime.now() - timedelta(days=3)
        three_days_ago = three_days_ago.replace(hour=0, minute=0, second=0, microsecond=0)
        today = datetime.now()

        # Cerca i record degli ultimi tre giorni in stock.move.line
        gtms_trips = self.env['gtms.trip'].search([('competence_date', '>=', three_days_ago),('competence_date', '<=', today)])
        _logger.info(gtms_trips)
        _logger.info(three_days_ago)
        
        # Costruisci il contenuto del file XLSX in memoria
        xlsx_content = io.BytesIO()
        workbook = xlsxwriter.Workbook(xlsx_content)
        worksheet = workbook.add_worksheet()

        headers = ['Id', 'Codice Viaggio', 'Trip Type', 'Source Document', 'From', 'Datetime start pianificato', 'To', 'Datetime end pianificato', 'Organization', 'N Stops', 'Datetime start sondaggio', 'Datetime end sondaggio', 'Vehicle', 'ID Driver', 'Driver', 'ID Learning Driver', 'Driver learning', 'Modalità pagamento', 'State']

        # Aggiungi gli header alla prima riga
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        row = 1  # Inizia dalla seconda riga per i dati
        for gtms_trip_id in gtms_trips:
            gtms_trip = self.env['gtms.trip'].search_read([('id', '=', gtms_trip_id.id)],['id','name','trip_type_id','source_document','from_address_partner_id','to_address_partner_id','first_stop_planned_at','last_stop_planned_at','organization_id','number_of_stops','trip_start_from_survey','trip_end_from_survey', 'current_fleet_id', 'all_drivers_ids', 'drivers_payment', 'state'], limit=1, order="id asc")
            # id = self.env['gtms.trip'].browse(gtms_trip.id)
            # record_id = int(str(id).split('(')[1].split(',')[0])
            _logger.info(gtms_trip)
            id = gtms_trip[0]['id']
            if gtms_trip[0]['name'] != False:
                trip_name = gtms_trip[0]['name']
            else:
                trip_name = ''
                
            if gtms_trip[0]['trip_type_id'] != False:
                trip_type_id = gtms_trip[0]['trip_type_id'][1]
            else:
                trip_type_id = ''
            if gtms_trip[0]['source_document'] != False:
                source_document = gtms_trip[0]['source_document']
            else:
                source_document = ''
            if gtms_trip[0]['from_address_partner_id'] != False:
                from_address_partner_id = gtms_trip[0]['from_address_partner_id'][1]
            else:
                from_address_partner_id = ''
            if gtms_trip[0]['to_address_partner_id'] != False:
                to_address_partner_id = gtms_trip[0]['to_address_partner_id'][1]
            else:
                to_address_partner_id = ''

            if gtms_trip[0]['first_stop_planned_at'] != False:
                first_stop_planned_at = gtms_trip[0]['first_stop_planned_at']
                utc_dt = pytz.utc.localize(first_stop_planned_at)
                first_stop_planned_at = utc_dt.astimezone(pytz.timezone('Europe/Rome'))
            else:
                first_stop_planned_at = ''
            if gtms_trip[0]['last_stop_planned_at'] != False:
                last_stop_planned_at = gtms_trip[0]['last_stop_planned_at']
                utc_dt = pytz.utc.localize(last_stop_planned_at)
                last_stop_planned_at = utc_dt.astimezone(pytz.timezone('Europe/Rome'))
            else:
                last_stop_planned_at = ''

            if gtms_trip[0]['organization_id'] != False:
                organization_id = gtms_trip[0]['organization_id'][1]
            else:
                organization_id = ''
            number_of_stops = gtms_trip[0]['number_of_stops']

            if gtms_trip[0]['trip_start_from_survey'] != False:
                trip_start_from_survey = gtms_trip[0]['trip_start_from_survey']
                utc_dt = pytz.utc.localize(trip_start_from_survey)
                trip_start_from_survey = utc_dt.astimezone(pytz.timezone('Europe/Rome'))
            else:
                trip_start_from_survey = ''
            if gtms_trip[0]['trip_end_from_survey'] != False:
                trip_end_from_survey = gtms_trip[0]['trip_end_from_survey']
                utc_dt = pytz.utc.localize(trip_end_from_survey)
                trip_end_from_survey = utc_dt.astimezone(pytz.timezone('Europe/Rome'))
            else:
                trip_end_from_survey = ''                
            
            if gtms_trip[0]['current_fleet_id'] != False:
                current_fleet_id = gtms_trip[0]['current_fleet_id'][1].split('/')[2]
            else:
                current_fleet_id = ''
            all_drivers_ids = gtms_trip[0]['all_drivers_ids'] 
            if len(gtms_trip[0]['all_drivers_ids']) == 1:
                driver_1_id = gtms_trip[0]['all_drivers_ids'][0]
                driver_1 = self.env['res.partner'].search_read([('id', '=', driver_1_id)], ['name'])[0]['name']
                driver_2 = ''
            elif len(gtms_trip[0]['all_drivers_ids']) == 2:
                driver_1_id = gtms_trip[0]['all_drivers_ids'][0]
                driver_1 = self.env['res.partner'].search_read([('id', '=', driver_1_id)], ['name'])[0]['name']
                driver_2_id = gtms_trip[0]['all_drivers_ids'][1]
                driver_2 = self.env['res.partner'].search_read([('id', '=', driver_2_id)], ['name'])[0]['name']
            else:
                driver_1 = ''
                driver_2 = ''
                driver_1_id = ''
                driver_2_id = ''
            drivers_payment = gtms_trip[0]['drivers_payment']
            state = gtms_trip[0]['state']

            
            _logger.info(id)
            _logger.info(trip_name)
            _logger.info(trip_type_id)
            _logger.info(source_document)
            _logger.info(from_address_partner_id)
            _logger.info(first_stop_planned_at)
            _logger.info(to_address_partner_id)
            _logger.info(last_stop_planned_at)
            _logger.info(organization_id)
            _logger.info(number_of_stops)
            _logger.info(trip_start_from_survey)
            _logger.info(trip_end_from_survey)
            _logger.info(organization_id)
            _logger.info(current_fleet_id)
            _logger.info("Strampa driver")
            _logger.info(all_drivers_ids)
            _logger.info(driver_1_id)
            _logger.info(driver_1)
            _logger.info(driver_2_id)
            _logger.info(driver_2)
            _logger.info(drivers_payment)
            _logger.info(state)
            

            
            worksheet.write(row, 0, str(id))
            worksheet.write(row, 1, str(trip_name))
            worksheet.write(row, 2, str(trip_type_id))
            worksheet.write(row, 3, str(source_document))
            worksheet.write(row, 4, str(from_address_partner_id))
            worksheet.write(row, 5, str(first_stop_planned_at))
            worksheet.write(row, 6, str(to_address_partner_id))
            worksheet.write(row, 7, str(last_stop_planned_at))
            worksheet.write(row, 8, str(organization_id))
            worksheet.write(row, 9, str(number_of_stops))
            worksheet.write(row, 10, str(trip_start_from_survey))
            worksheet.write(row, 11, str(trip_end_from_survey))
            worksheet.write(row, 12, str(current_fleet_id))
            worksheet.write(row, 13, str(driver_1_id))
            worksheet.write(row, 14, str(driver_1))
            worksheet.write(row, 15, str(driver_2_id))
            worksheet.write(row, 16, str(driver_2))
            worksheet.write(row, 17, str(drivers_payment))
            worksheet.write(row, 18, str(state))



            row += 1  # Passa alla riga successiva per il prossimo stock_move

        workbook.close()

        # Imposta l'allegato in Odoo come file XLSX
        xlsx_content.seek(0)
        attachment_values = {
            'name': 'gtms_trip.xlsx',
            'datas': base64.encodebytes(xlsx_content.getvalue()).decode(),
            'res_model': self._name,
            'res_id': self.id,
            'type': 'binary',
        }
        attachment = self.env['ir.attachment'].create(attachment_values)

        # Invia l'email con l'allegato
        mail_values = {
            'subject': 'Gtms trip dal ' + three_days_ago.strftime('%d-%m-%Y') + ' al ' + today.strftime('%d-%m-%Y'),
            'email_from': 'noreply@futurasl.com',
            'email_to': 'dati+gtmstrip@svcfutura.cloud',
            'body_html': '<p>In allegato file .XLSX con i viaggi degli ultimi 3 giorni.</p>',
            'attachment_ids': [(4, attachment.id)],  # Aggiungi l'allegato all'email
        }

        # Crea e invia l'email utilizzando il metodo create di mail.mail
        mail = self.env['mail.mail'].sudo().create(mail_values)
        mail.send()

