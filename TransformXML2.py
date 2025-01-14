import paramiko
import os
import sys
import glob
import shutil
import re
import time
import itertools
import logging
import smtplib
import pandas as pd
from lxml import etree
from deepdiff import DeepDiff
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

# Konfigūracijos
SFTP_CONFIG = {
    'HOST': 's.judu.lt',
    'PORT': 22,
    'USERNAME': 'ridango',
    'PASSWORD': 'zYi%Eft5X~C+C.A?w"Ze',
    'REMOTE_PATH': '/mnt/',
    'ARCHIVE_PATH': '/mnt/Archive/'
}

SFTP_ODOO_CONFIG = {
    'HOST': 'pagalba.judu.lt',
    'PORT': 22,
    'USERNAME': 'exportuser',
    'PASSWORD': '7yHbplT8p6bGPsGYuXSe',
    'REMOTE_PATH': '/opt/odoo/addons_rdail/Eksportai/'
}

EMAIL_CONFIG = {
    'SMTP_SERVER': "smtp.sisp.lt",
    #'SMTP_SERVER': "mail.sisp.lt",
    'FROM': "xml-orderiai@judu.lt",
    'TO': "xml-orderiai@judu.lt",
    'SUBJECT': "Klaida vykdant TransformXML Operacijas",
    'EMAIL_BODY' : "Informuojame, kad vykdant TransformXML operacijas, įvyko neplanuota klaida. Pridėtas failas su klaidomis logfile.log"
}

# Aplankai
archive_directory = 'Archyvas/'
updated_directory = 'Pakeisti_failai/'
retailer_directory = updated_directory + 'Platintojai'

transcription_map = {
    # Russian characters
    
    'а': 'a', 'е': 'e', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'ё': 'yo',
    'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm',
    'н': 'n', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
    'ф': 'f', 'х': 'h', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'sch', 'ъ': '',
    'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya',

    # Bulgarian characters
    'ьо': 'yo', 'ъа': 'a',

    # Ukrainian characters
    'ґ': 'g', 'є': 'ye', 'і': 'i', 'ї': 'yi',

    # Belarusian characters
    'ў': 'u',

    # Others
    'љ': 'lj', 'њ': 'nj', 'џ': 'dz',

    # Uppercase Russian characters
    'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D', 'Ё': 'Yo', 
    'Ж': 'Zh', 'З': 'Z', 'И': 'I', 'Й': 'Y', 'К': 'K', 'Л': 'L', 'М': 'M',
    'Н': 'N', 'П': 'P', 'Р': 'R', 'С': 'S', 'Т': 'T', 'У': 'U',
    'Ф': 'F', 'Х': 'H', 'Ц': 'Ts', 'Ч': 'Ch', 'Ш': 'Sh', 'Щ': 'Sch', 'Ъ': '',
    'Ы': 'Y', 'Ь': '', 'Э': 'E', 'Ю': 'Yu', 'Я': 'Ya',

    # Uppercase Bulgarian characters
    'ЬО': 'YO', 'ЪА': 'A',

    # Uppercase Ukrainian characters
    'Ґ': 'G', 'Є': 'YE', 'Ї': 'YI',

    # Uppercase Belarusian characters
    'Ў': 'U',

    # Uppercase others
    'Љ': 'LJ', 'Њ': 'NJ', 'Џ': 'DZ',

    # Additional Latin-based characters
    'ć': 'c', 'đ': 'dj', 'ě': 'e', 'ł': 'l', 'ń': 'n', 'ř': 'r', 
    'ť': 't', 'ů': 'u', 
    
    'Ć': 'C', 'Đ': 'Dj', 'Ě': 'E', 'Ł': 'L', 'Ń': 'N', 'Ř': 'R', 
    'Ť': 'T', 'Ů': 'U', 
    
    # Some common diacritics
    'â': 'a', 'ä': 'a', 'á': 'a', 'à': 'a', 'ã': 'a',
    'ê': 'e', 'ë': 'e', 'é': 'e', 'è': 'e',
    'î': 'i', 'ï': 'i', 'í': 'i', 'ì': 'i',
    'ô': 'o', 'ö': 'o', 'ó': 'o', 'ò': 'o', 'õ': 'o',
    'û': 'u', 'ü': 'u', 'ú': 'u', 'ù': 'u',
    'ñ': 'n', 'ç': 'c', 'ß': 'ss',
    
    'Â': 'A', 'Ä': 'A', 'Á': 'A', 'À': 'A', 'Ã': 'A',
    'Ê': 'E', 'Ë': 'E', 'É': 'E', 'È': 'E',
    'Î': 'I', 'Ï': 'I', 'Í': 'I', 'Ì': 'I',
    'Ô': 'O', 'Ö': 'O', 'Ó': 'O', 'Ò': 'O', 'Õ': 'O',
    'Û': 'U', 'Ü': 'U', 'Ú': 'U', 'Ù': 'U',
    'Ñ': 'N', 'Ç': 'C'
}

def custom_print(*args, **kwargs):
    """Funkcija, kuri užtikrina, kad pranešimai būtų matomi tiek konsolėje, tiek exe faile"""
    print(*args, **kwargs)
    if sys.stdout is not None:
        sys.stdout.flush()
    else:
        # If sys.stdout is None, we can log this occurrence
        logging.warning("sys.stdout is None in custom_print function")

def transcribe_russian_to_latin(text):
    result = ''.join(transcription_map.get(char, char) for char in text)
    # Palikti tik lotyniškas ir lietuviškas raides
    result = re.sub(r'[^a-zA-ZąčęėįšųūžĄČĘĖĮŠŲŪŽ0-9 -~_@^]', '', result)
    return result

def secure_sftp_connection(config):
    try:
        custom_print(f"Jungiamasi prie SFTP serverio: {config['HOST']}")
        transport = paramiko.Transport((config['HOST'], config['PORT']))
        transport.connect(username=config['USERNAME'], password=config['PASSWORD'])
        sftp = paramiko.SFTPClient.from_transport(transport)
        custom_print(f"Sėkmingai prisijungta prie SFTP serverio: {config['HOST']}")
        return sftp, transport
    except Exception as e:
        error_msg = f"SFTP prisijungimas nepavyko: {str(e)}"
        custom_print(error_msg)
        logging.error(error_msg)
        raise

def download_files(sftp, remote_path, local_path):
    try:
        custom_print(f"Pradedamas failų atsisiuntimas iš {remote_path}")
        sftp.chdir(remote_path)
        for filename in sftp.listdir():
            if filename.endswith(('.xlsx', '.xml', '.csv')):
                sftp.get(filename, os.path.join(local_path, filename))
                custom_print(f"Atsisiųsta: {filename}")
        custom_print("Failų atsisiuntimas baigtas")
    except Exception as e:
        error_msg = f"Failų atsisiuntimas nepavyko: {str(e)}"
        custom_print(error_msg)
        logging.error(error_msg)
        raise

def process_files(files, all_data):
    custom_print("Pradedamas failų apdorojimas")
    for file in files:
        custom_print(f"Apdorojamas failas: {file}")
        if file.endswith('.xml'):
            process_xml(file, all_data)
        elif file.endswith('.xlsx'):
            process_excel(file)
        elif file.endswith('.csv'):
            process_csv(file)
    custom_print("Failų apdorojimas baigtas")

def process_xml(file, all_data, error_messages):
    try:
        custom_print(f"Apdorojamas XML failas: {file}")
        parser = etree.XMLParser(recover=True)
        tree = etree.parse(file, parser=parser)
        root = tree.getroot()
        
        changes_made = False
        missing_clients_found = False
        missing_clients = []

        for client in root.iter('Client'):
            order_no = client.find('InvoiceData').get('OrderNo')
            
            if order_no is None or all_data[all_data['Ridango ID'] == int(order_no)].empty:
                missing_clients.append(client)
                missing_clients_found = True
                root.remove(client)
                continue

            order_row = all_data[all_data['Ridango ID'] == int(order_no)]
            if order_row.empty:
                continue

            user_name = str(order_row['Vartotojas'].values[0])
            company_name = str(order_row['Įmonė'].values[0])
            company_code = order_row['Įmonės kodas'].values[0]
            vat_id = str(order_row['PVM mokėtojo kodas'].values[0])

            client_name = user_name if user_name else company_name
            client_name_transcribed = transcribe_russian_to_latin(client_name)
            client.set('clientName', client_name_transcribed)

            if not user_name:
                if company_code and vat_id:
                    client.set('clientID', str(int(company_code)))
                    client.set('clientVATID', vat_id)
                else:
                    logging.error(f"Klaida: Negalima priskirti clientID ir clientVATID, nes company_code arba vat_id yra tuščias. Kliento informacija: {etree.tostring(client, encoding='utf-8').decode()}")

            invoice_no = str(order_row['Sąskaitos Nr.'].values[0])
            split_invoice_no = invoice_no.split('-')
            if len(split_invoice_no) > 1:
                split_invoice_no[1] = str(split_invoice_no[1]).zfill(10)
                invoice_no = '-'.join(split_invoice_no)
            else:
                logging.error(f"Klaida: Sąskaitos numeris neturi tinkamos formos: {invoice_no}")

            client.find('InvoiceData').set('InvoiceNo', invoice_no)
            changes_made = True

        if changes_made or missing_clients_found:
            is_web_org = "web_org" in os.path.basename(file).lower()
            target_directory = os.path.join(updated_directory, "web_org") if is_web_org else updated_directory
            
            if not os.path.exists(target_directory):
                os.makedirs(target_directory)

            updated_file_path = os.path.join(target_directory, os.path.basename(file))
            tree.write(updated_file_path, encoding='utf-8')

            if missing_clients_found and is_web_org:
                missing_clients_filename = os.path.join(target_directory, os.path.splitext(os.path.basename(file))[0] + '_missing_clients.xml')
                with open(missing_clients_filename, 'wb') as missing_clients_file:
                    for client in missing_clients:
                        missing_clients_file.write(etree.tostring(client, encoding='utf-8'))

            if not os.path.exists(archive_directory):
                os.makedirs(archive_directory)
            archive_location = os.path.join(archive_directory, os.path.basename(file))
            shutil.move(file, archive_location)

        custom_print(f"XML failas {file} sėkmingai apdorotas")
    except Exception as e:
        error_message = f"XML apdorojimas nepavyko failui {file}: {str(e)}"
        custom_print(error_message)
        logging.error(error_message)
        error_messages.append(error_message)

def process_excel(file, error_messages):
    try:
        custom_print(f"Apdorojamas Excel failas: {file}")
        df = pd.read_excel(file)
        # Excel processing logic here
        df.to_excel(file, index=False)
        custom_print(f"Excel failas {file} sėkmingai apdorotas")
    except Exception as e:
        error_message = f"Excel apdorojimas nepavyko failui {file}: {str(e)}"
        custom_print(error_message)
        logging.error(error_message)
        error_messages.append(error_message)

def process_csv(file, error_messages):
    try:
        custom_print(f"Apdorojamas CSV failas: {file}")
        df = pd.read_csv(file)
        # CSV processing logic here
        df.to_csv(file, index=False)
        custom_print(f"CSV failas {file} sėkmingai apdorotas")
    except Exception as e:
        error_message = f"CSV apdorojimas nepavyko failui {file}: {str(e)}"
        custom_print(error_message)
        logging.error(error_message)
        error_messages.append(error_message)

# Siųsti el. pašto pranešimą su klaidos informacija
def send_error_email(config, error_messages):
    try:
        custom_print("Siunčiamas klaidos pranešimas el. paštu")
        msg = MIMEMultipart()
        msg['From'] = config['FROM']
        msg['To'] = config['TO']
        msg['Subject'] = config['SUBJECT']
        
        # Pridedame el. laiško turinį
        body = "Informuojame, kad vykdant TransformXML operacijas, įvyko neplanuota klaida. Klaidos pranešimas:\n\n"
        body += "".join(error_messages) 
        body += "\n\nPridėtas failas su klaidomis logfile.log"
        
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Pridedame log failą kaip priedą
        with open('logfile.log', 'r', encoding='utf-8') as f:
            log_content = f.read()
        
        attachment = MIMEText(log_content, _charset='utf-8')
        attachment.add_header('Content-Disposition', 'attachment', filename='logfile.log')
        msg.attach(attachment)
        
        with smtplib.SMTP(config['SMTP_SERVER'], 25) as server:
            server.send_message(msg)
        custom_print('El. laiškas sėkmingai išsiųstas')
    except Exception as e:
        error_msg = f"Nepavyko išsiųsti el. laiško: {str(e)}"
        custom_print(error_msg)
        logging.error(error_msg)

        # Bandome išsaugoti klaidos pranešimą į atskirą failą
        try:
            with open('email_error.log', 'w', encoding='utf-8') as error_file:
                error_file.write(error_msg)
            custom_print("Klaidos pranešimas išsaugotas į 'email_error.log' failą")
        except Exception as save_error:
            custom_print(f"Nepavyko išsaugoti klaidos pranešimo į failą: {str(save_error)}")
def main():
    error_messages = []

    try:
        custom_print("Programos vykdymas pradėtas")
        
        # Nustatome vykdomojo failo direktoriją
        if getattr(sys, 'frozen', False):
          # Jei programa yra exe failas
            application_path = os.path.dirname(sys.executable)
        else:
            # Jei programa vykdoma kaip Python skriptas
            application_path = os.path.dirname(os.path.abspath(__file__))

        # Pakeičiame dabartinę darbinę direktoriją į programos kelią
        os.chdir(application_path)
        # Logging konfigūracija
        open('logfile.log', 'w', encoding='utf-8')
        logging.basicConfig(
            filename='logfile.log',
            level=logging.ERROR, 
            format='%(asctime)s %(levelname)s %(message)s', 
            datefmt='%Y-%m-%d %H:%M:%S',
            encoding='utf-8'
        )
        # Pridedame handler'į, kad matytume žurnalus ir konsolėje
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(formatter)
        logging.getLogger('').addHandler(console_handler)

        custom_print(f"Darbinė direktorija pakeista į: {os.getcwd()}")

        # SFTP operacijos
        custom_print("Pradedamos SFTP operacijos")
        try:
            sftp_ridango, transport_ridango = secure_sftp_connection(SFTP_CONFIG)
            download_files(sftp_ridango, SFTP_CONFIG['REMOTE_PATH'], '.')
            sftp_ridango.close()
            transport_ridango.close()
        except Exception as e:
            error_message = f"Klaida SFTP operacijose su Ridango: {str(e)}"
            logging.error(error_message)
            error_messages.append(error_message)

        try:
            sftp_odoo, transport_odoo = secure_sftp_connection(SFTP_ODOO_CONFIG)
            download_files(sftp_odoo, SFTP_ODOO_CONFIG['REMOTE_PATH'], '.')
            sftp_odoo.close()
            transport_odoo.close()
        except Exception as e:
            error_message = f"Klaida SFTP operacijose su Odoo: {str(e)}"
            logging.error(error_message)
            error_messages.append(error_message)

        custom_print("SFTP operacijos baigtos")

        # Pridedame datą prie ODOO Excel failo pavadinimo
        current_date = datetime.now().strftime("%Y%m%d")
        odoo_excel_files = glob.glob(os.path.join(application_path, '*.xlsx'))
        for odoo_file in odoo_excel_files:
            file_name, file_extension = os.path.splitext(odoo_file)
            new_file_name = f"{file_name}_{current_date}{file_extension}"
            os.rename(odoo_file, new_file_name)
            custom_print(f"ODOO Excel failo pavadinimas pakeistas į: {os.path.basename(new_file_name)}")

        # Atnaujintas failų sąrašas po pavadinimų pakeitimo
        excel_files = glob.glob(os.path.join(application_path, '*.xlsx'))
        xml_files = glob.glob(os.path.join(application_path, '*.xml'))
        csv_files = glob.glob(os.path.join(application_path, '*.csv'))
        all_files = excel_files + xml_files + csv_files

        total_files = len(all_files)

        custom_print(f"Pradėtas {total_files} failų apdorojimas...")

        # Įkelti visus Excel failus į vieną DataFrame
        dataframes = []
        for excel_file in excel_files:
            try:
                df = pd.read_excel(excel_file)
                dataframes.append(df)
            except Exception as e:
                error_message = f"Klaida skaitant .xlsx failą {excel_file}: {str(e)}"
                custom_print(error_message)
                logging.error(error_message)
                error_messages.append(error_message)


        # Sujungti visus DataFrame į vieną
        if dataframes:
            all_data = pd.concat(dataframes, ignore_index=True)
            all_data.fillna("", inplace=True)
            all_data.drop_duplicates(inplace=True)
        else:
            error_message = "Nėra .xlsx (orders) failų, kuriuos galima apdoroti. Nutraukiamas programos vykdymas."
            custom_print(error_message)
            logging.error(error_message)
            error_messages.append(error_message)
            sys.exit()

        # Apdorojame failus
        for file in all_files:
            file_extension = os.path.splitext(file)[1].lower()
            
            # Patikriname ar failo pavadinimas prasideda "Settlement Report"
            if "Settlement Report" and file.endswith('.csv'):
                if not os.path.exists(updated_directory):
                    os.makedirs(updated_directory)
                new_location = os.path.join(updated_directory, os.path.basename(file))
                shutil.move(file, new_location)
                custom_print(f"Failas {file} perkeltas į {updated_directory} be apdorojimo.")
                continue

            # Patikriname ar failo pavadinimas turi "CEMV" (nepaisant didžiųjų/mažųjų raidžių)
            if "cemv" in os.path.basename(file).lower():
                if not os.path.exists(updated_directory):
                    os.makedirs(updated_directory)
                new_location = os.path.join(updated_directory, os.path.basename(file))
                shutil.move(file, new_location)
                custom_print(f"Failas {file} perkeltas į {updated_directory} be apdorojimo.")
                continue

            # Patikriname ar tai platintojo failas arba date_usage_wallet failas
            if ("Retailer" in os.path.basename(file) and file.endswith('.xml')) or "date_usage_wallet" in os.path.basename(file):
               if not os.path.exists(retailer_directory):
                   os.makedirs(retailer_directory)
               new_location = os.path.join(retailer_directory, os.path.basename(file))
               shutil.move(file, new_location)
               custom_print(f"Failas {file} perkeltas į Platintojai direktoriją")
               continue

            # Apdorojame likusius failus
            if file_extension == '.xml':
                process_xml(file, all_data, error_messages)
            elif file_extension == '.xlsx':
                process_excel(file, error_messages)
            elif file_extension == '.csv':
                 process_csv(file, error_messages)

        # Perkeliame originalius Excel failus į "Archyvas" direktoriją
        for excel_file in excel_files:
            if not os.path.exists(archive_directory):
                os.makedirs(archive_directory)
            new_location = os.path.join(archive_directory, os.path.basename(excel_file))
            shutil.move(excel_file, new_location)
            custom_print(f"Excel failas {excel_file} perkeltas į {archive_directory}")

        # SMTP failų perkėlimas
        custom_print("Pradedamas SMTP failų perkėlimas")
        sftp_ridango, transport_ridango = secure_sftp_connection(SFTP_CONFIG)
        try:
            remote_files = sftp_ridango.listdir(SFTP_CONFIG['REMOTE_PATH'])
            for remote_file in remote_files:
                if remote_file.endswith('.xml') or remote_file.endswith('.csv'):
                    remote_file_path = os.path.join(SFTP_CONFIG['REMOTE_PATH'], remote_file)
                    archive_file_path = os.path.join(SFTP_CONFIG['ARCHIVE_PATH'], remote_file)
                    try:
                        sftp_ridango.rename(remote_file_path, archive_file_path)
                        custom_print(f"SMTP failas {remote_file} perkeltas į archyvą")
                    except IOError as e:
                        error_message = f"Klaida perkeliant SMTP failą {remote_file} į archyvą: {str(e)}"
                        custom_print(error_message)
                        logging.error(error_message)
                        error_messages.append(error_message)
        except Exception as e:
            error_message = f"Klaida perkeliant SMTP failus: {str(e)}"
            custom_print(error_message)
            logging.error(error_message)
            error_messages.append(error_message)

        finally:
            sftp_ridango.close()
            transport_ridango.close()

        custom_print("Apdorojimas sėkmingai baigtas")
    except Exception as e:
        error_message = f"Įvyko nenumatyta klaida: {str(e)}"
        custom_print(error_message)
        logging.error(error_message)
        error_messages.append(error_message)

    finally:
        if logging.getLogger().hasHandlers():
            log_handler = logging.getLogger().handlers[0]
            log_handler.close()
            logging.getLogger().removeHandler(log_handler)
        if error_messages:
            send_error_email(EMAIL_CONFIG, "\n".join(error_messages))

        custom_print("Programa užsidarys po 10 sekundžių...")
        time.sleep(10)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error_message = f"Įvyko nenumatyta klaida: {str(e)}"
        logging.error(error_message)
        custom_print(error_message)

