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

    def export_gtms_trip_xlsx(self, giorni):
        # Calcola la data di tre giorni fa
        three_days_ago = datetime.now() - timedelta(days=giorni)
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

        headers = ['Id', 'Codice Viaggio', 'Trip Type', 'Source Document', 'From', 'Datetime start pianificato', 'To', 'Datetime end pianificato', 'Organization', 'N Stops', 'Datetime start sondaggio', 'Datetime end sondaggio', 'Vehicle', 'Vehicle ID', 'ID Driver', 'Driver', 'Pwork Az 1', 'Pwork Dip 1', 'ID Learning Driver', 'Driver learning', 'Pwork Az Learning', 'Pwork Dip Learning', 'Modalità pagamento', 'State', 'Distance Expected', 'Total Sales', 'Total Purchases', 'Cliente', 'Vettore', 'Numero OC', 'Numero Ddt', 'Anticipo Partenza', 'Anticipo Partenza Extra', 'Totale Documenti', 'Totale documenti fornitore', 'Totale documenti fornitore completi', 'Km cliente', 'Città partenza', 'CAP partenza', 'Stato partenza', 'Provincia partenza', 'Indirizzo di partenza', 'Città arrivo', 'CAP arrivo', 'Stato arrivo', 'Provincia arrivo', 'Indirizzo di arrivo', 'KM GPS']

        # Aggiungi gli header alla prima riga
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        row = 1  # Inizia dalla seconda riga per i dati
        for gtms_trip_id in gtms_trips:
            gtms_trip = self.env['gtms.trip'].search_read([('id', '=', gtms_trip_id.id)],['id','name','trip_type_id','source_document','from_address_partner_id','to_address_partner_id','first_stop_planned_at','last_stop_planned_at','organization_id','number_of_stops','trip_start_from_survey','trip_end_from_survey', 'current_fleet_id', 'all_drivers_ids', 'drivers_payment', 'state', 'distance_expected', 'total_sales', 'total_purchases', 'related_customer_id', 'related_supplier_id', 'oc_number', 'ddt_number', 'anticipo_partneza', 'anticipo_partenza_extra', 'total_documents', 'total_supplier_documents', 'total_complete_supplier_documents', 'km_from_customer', 'from_city', 'from_zip', 'from_country_id', 'from_state_id', 'from_street', 'to_city', 'to_zip', 'to_country_id', 'to_state_id', 'to_street', 'gps_km'], limit=1, order="id asc")
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
            # Id veicolo odoo
            if gtms_trip[0]['current_fleet_id'] != False:
                current_fleet_id = gtms_trip[0]['current_fleet_id'][1].split('/')[2]
            else:
                current_fleet_id = ''
            # Targa veicolo
            if gtms_trip[0]['current_fleet_id'] != False:
                current_fleet_license_plate = gtms_trip[0]['current_fleet_id'][0]
            else:
                current_fleet_license_plate = ''
            all_drivers_ids = gtms_trip[0]['all_drivers_ids'] 
            if len(gtms_trip[0]['all_drivers_ids']) == 1:
                driver_1_id = gtms_trip[0]['all_drivers_ids'][0]
                driver_1 = self.env['res.partner'].search_read([('id', '=', driver_1_id)], ['name'])[0]['name']
                driver_2_id = ''
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
            if gtms_trip[0]['distance_expected'] != False:
                distance_expected = gtms_trip[0]['distance_expected']
            else:
                distance_expected = ''
            if gtms_trip[0]['total_sales'] != False:
                total_sales = gtms_trip[0]['total_sales']
            else:
                total_sales = ''
            if gtms_trip[0]['total_purchases'] != False:
                total_purchases = gtms_trip[0]['total_purchases']
            else:
                total_purchases = ''
            if gtms_trip[0]['related_customer_id'] != False:
                related_customer_id = gtms_trip[0]['related_customer_id'][1]
            else:
                related_customer_id = ''
            if gtms_trip[0]['related_supplier_id'] != False:
                related_supplier_id = gtms_trip[0]['related_supplier_id'][1]
            else:
                related_supplier_id = ''
            if gtms_trip[0]['oc_number'] != False:
                oc_number = gtms_trip[0]['oc_number']
            else:
                oc_number = ''
            if gtms_trip[0]['ddt_number'] != False:
                ddt_number = gtms_trip[0]['ddt_number']
            else:
                ddt_number = ''
            if gtms_trip[0]['anticipo_partneza'] != False:
                anticipo_partneza = gtms_trip[0]['anticipo_partneza']
            else:
                anticipo_partneza = ''
            if gtms_trip[0]['anticipo_partenza_extra'] != False:
                anticipo_partenza_extra = gtms_trip[0]['anticipo_partenza_extra']
            else:
                anticipo_partenza_extra = ''
            if gtms_trip[0]['total_documents'] != False:
                total_documents = gtms_trip[0]['total_documents']
            else:
                total_documents = ''
            if gtms_trip[0]['total_supplier_documents'] != False:
                total_supplier_documents = gtms_trip[0]['total_supplier_documents']
            else:
                total_supplier_documents = ''
            if gtms_trip[0]['total_complete_supplier_documents'] != False:
                total_complete_supplier_documents = gtms_trip[0]['total_complete_supplier_documents']
            else:
                total_complete_supplier_documents = ''
            if gtms_trip[0]['km_from_customer'] != False:
                km_from_customer = gtms_trip[0]['km_from_customer']
            else:
                km_from_customer = ''
            if gtms_trip[0]['from_city'] != False:
                from_city = gtms_trip[0]['from_city']
            else:
                from_city = ''
            if gtms_trip[0]['from_zip'] != False:
                from_zip = gtms_trip[0]['from_zip']
            else:
                from_zip = ''
            if gtms_trip[0]['from_country_id'] != False:
                from_country_id = gtms_trip[0]['from_country_id'][1]
            else:
                from_country_id = ''
            if gtms_trip[0]['from_state_id'] != False:
                from_state_id = gtms_trip[0]['from_state_id'][1]
            else:
                from_state_id = ''
            if gtms_trip[0]['from_street'] != False:
                from_street = gtms_trip[0]['from_street']
            else:
                from_street = ''
            if gtms_trip[0]['to_city'] != False:
                to_city = gtms_trip[0]['to_city']
            else:
                to_city = ''
            if gtms_trip[0]['to_zip'] != False:
                to_zip = gtms_trip[0]['to_zip']
            else:
                to_zip = ''
            if gtms_trip[0]['to_country_id'] != False:
                to_country_id = gtms_trip[0]['to_country_id'][1]
            else:
                to_country_id = ''
            if gtms_trip[0]['to_state_id'] != False:
                to_state_id = gtms_trip[0]['to_state_id'][1]
            else:
                to_state_id = ''
            if gtms_trip[0]['to_street'] != False:
                to_street = gtms_trip[0]['to_street']
            else:
                to_street = ''
            if gtms_trip[0]['gps_km'] != False:
                gps_km = gtms_trip[0]['gps_km']
            else:
                gps_km = ''

            
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
            _logger.info(current_fleet_license_plate)
            _logger.info(current_fleet_id)
            _logger.info("Strampa driver")
            _logger.info(all_drivers_ids)
            _logger.info(driver_1_id)
            _logger.info(driver_1)
            _logger.info(driver_2_id)
            _logger.info(driver_2)
            _logger.info(drivers_payment)
            _logger.info(state)
            _logger.info(distance_expected)
            _logger.info(total_sales)
            _logger.info(total_purchases)
            _logger.info(related_customer_id)
            _logger.info(related_supplier_id)
            _logger.info(oc_number)
            _logger.info(ddt_number)
            _logger.info(anticipo_partneza)
            _logger.info(anticipo_partenza_extra)
            _logger.info(total_documents)
            _logger.info(total_supplier_documents)
            _logger.info(total_complete_supplier_documents)
            _logger.info(km_from_customer)
            _logger.info(from_city)
            _logger.info(from_zip)
            _logger.info(from_country_id)
            _logger.info(from_state_id)
            _logger.info(from_street)
            _logger.info(to_city)
            _logger.info(to_zip)
            _logger.info(to_country_id)
            _logger.info(to_state_id)
            _logger.info(to_street)
            _logger.info(gps_km)

            id_az_pwork_1 = ''
            id_dip_pwork_1 = ''
            id_az_pwork_2 = ''
            id_dip_pwork_2 = ''

            # STAMPO ID HR.EMPLOYEE PWORK ASSOCIATO AL RES.PARTNER
            if driver_1_id != '':
                employee_id_1 = self.env['hr.employee'].search_read([('address_home_id', '=', driver_1_id)], ['id'])
                _logger.info(employee_id_1)
                for employee_1 in employee_id_1:
                    employee_contract_id_1 = self.env['hr.contract'].search_read([('employee_id', '=', employee_1['id']),('date_start', '<=', first_stop_planned_at), '|', ('date_end', '=', False), ('date_end', '>=', first_stop_planned_at)],['id'])
                    _logger.info("STAMPO contratto attivo")
                    if employee_contract_id_1 != []:
                        _logger.info(employee_contract_id_1)
                        _logger.info(employee_contract_id_1[0]['id'])
                        _logger.info("STAMPO id dipendente")
                        _logger.info(employee_1['id'])
                        # Cerco id_dip e id_az pwork
                        id_az_pwork_1 = self.env['hr.employee'].search_read([('id', '=', employee_1['id'])],['pwork_azienda_id'])[0]['pwork_azienda_id']
                        id_dip_pwork_1 = self.env['hr.employee'].search_read([('id', '=', employee_1['id'])],['pwork_dipendente_id'])[0]['pwork_dipendente_id']
                        _logger.info("STAMPO ID AZIENDA PWORK DIPENDENTE 1")
                        _logger.info(id_az_pwork_1)
                        _logger.info("STAMPO ID DIPENDENTI PWORK DIPENDENTE 1")
                        _logger.info(id_dip_pwork_1)
                        continue
                    else:
                        _logger.info("SETTO FALSE")
                        id_az_pwork_1 = ''
                        id_dip_pwork_1 = ''

            if driver_2_id != '':
                employee_id_2 = self.env['hr.employee'].search_read([('address_home_id', '=', driver_2_id)], ['id'])
                _logger.info(employee_id_2)
                for employee_2 in employee_id_2:
                    _logger.info("STAMPO ID 2")
                    _logger.info(driver_2_id)
                    employee_contract_id_2 = self.env['hr.contract'].search_read([('employee_id', '=', employee_2['id']),('date_start', '<=', first_stop_planned_at), '|', ('date_end', '=', False), ('date_end', '>=', first_stop_planned_at)],['id'])
                    _logger.info("STAMPO contratto attivo")
                    if employee_contract_id_2 != []:
                        _logger.info(employee_contract_id_2)
                        _logger.info(employee_contract_id_2[0]['id'])
                        _logger.info("STAMPO id dipendente")
                        _logger.info(employee_2['id'])
                        # Cerco id_dip e id_az pwork
                        id_az_pwork_2 = self.env['hr.employee'].search_read([('id', '=', employee_2['id'])],['pwork_azienda_id'])[0]['pwork_azienda_id']
                        id_dip_pwork_2 = self.env['hr.employee'].search_read([('id', '=', employee_2['id'])],['pwork_dipendente_id'])[0]['pwork_dipendente_id']
                        _logger.info("STAMPO ID AZIENDA PWORK DIPENDENTE 2")
                        _logger.info(id_az_pwork_2)
                        _logger.info("STAMPO ID DIPENDENTI PWORK DIPENDENTE 2")
                        _logger.info(id_dip_pwork_2)
                        break
                    else:
                        _logger.info("SETTO FALSE")
                        id_az_pwork_2 = ''
                        id_dip_pwork_2 = ''
            

            
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
            worksheet.write(row, 12, str(current_fleet_license_plate))
            worksheet.write(row, 13, str(current_fleet_id))
            worksheet.write(row, 14, str(driver_1_id))
            worksheet.write(row, 15, str(driver_1))
            worksheet.write(row, 16, str(id_az_pwork_1))
            worksheet.write(row, 17, str(id_dip_pwork_1))
            worksheet.write(row, 18, str(driver_2_id))
            worksheet.write(row, 19, str(driver_2))
            worksheet.write(row, 20, str(id_az_pwork_2))
            worksheet.write(row, 21, str(id_dip_pwork_2))
            worksheet.write(row, 22, str(drivers_payment))
            worksheet.write(row, 23, str(state))
            worksheet.write(row, 24, str(distance_expected))
            worksheet.write(row, 25, str(total_sales))
            worksheet.write(row, 26, str(total_purchases))
            worksheet.write(row, 27, str(related_customer_id))
            worksheet.write(row, 28, str(related_supplier_id))
            worksheet.write(row, 29, str(oc_number))
            worksheet.write(row, 30, str(ddt_number))
            worksheet.write(row, 31, str(anticipo_partneza))
            worksheet.write(row, 32, str(anticipo_partenza_extra))
            worksheet.write(row, 33, str(total_documents))
            worksheet.write(row, 34, str(total_supplier_documents))
            worksheet.write(row, 35, str(total_complete_supplier_documents))
            worksheet.write(row, 36, str(km_from_customer))
            worksheet.write(row, 37, str(from_city))
            worksheet.write(row, 38, str(from_zip))
            worksheet.write(row, 39, str(from_country_id))
            worksheet.write(row, 40, str(from_state_id))
            worksheet.write(row, 41, str(from_street))
            worksheet.write(row, 42, str(to_city))
            worksheet.write(row, 43, str(to_zip))
            worksheet.write(row, 44, str(to_country_id))
            worksheet.write(row, 45, str(to_state_id))
            worksheet.write(row, 46, str(to_street))
            worksheet.write(row, 47, str(gps_km))


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
            'body_html': f'<p>In allegato file .XLSX con i viaggi degli ultimi {giorni} giorni.</p>',
            'attachment_ids': [(4, attachment.id)],  # Aggiungi l'allegato all'email
        }

        # Crea e invia l'email utilizzando il metodo create di mail.mail
        mail = self.env['mail.mail'].sudo().create(mail_values)
        mail.send()

