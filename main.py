import xmltodict, logging, requests, os, openpyxl
from datetime import datetime
from pathlib import Path


logging.basicConfig(level=logging.INFO, filename="xmlparser.log", format='%(asctime)s - %(levelname)s - %(message)s')

"""
    # * We can't remove TMS Shipments with Japan Origins Workflow because we are operating without all the information
Plant Code + SLOC Number is a unique identifier that we'll use to find the proper location reference in worldtrak. 
# XTODO: TMS Shipments Only. Get fixed addresses matrix from Tim, compare Street Number, City, State, Country and use that matrix push out the proper address
    STREET ADDRESS	CITY	            COUNTRY	        ACTUAL ADDRESS
    7826	        OsanSi, GyeonggiDo	KR	            #78-26 GYEONGGI-BAERO
    6753	        Newark	            US	            6753 MOWRY AVE
    6551	        Tracy	            US	            6551 W SCHULTE ROAD
    1201	        LIVERMORE	        US	            1201 VOYAGER STREET
    21000	        Tualatin	        US	            21000 SW 115TH AVE

# XTODO: Update Workflow to change flowlines for Scott
# TODO: Change from Japan Shipments being CSVs to XLSX files that append to the exisiting files
# XTODO: Follow up with LAM to see if they can upload the US Outbound Report to the FTP server, rather than just emailing it to us
    # * LAM is pushing these items to be uploaded. This is an action item we can take back to LAM. 
# XTODO: Tim to send list of Shippers he wants consolidated into single Japan XLSX Shipper Files
    Shipper Name	                            File Name
    Edwards Japan	                            EDWARDS
    TAICA Corporation	                        INTERNAL
    TDK CORPORATION C/O ALPS LOGISTICS	        TDK
    TDK CORPORATION	                            TDK
    TOTO LTD	                                TOTO
    AVISERVICE C/O TDK CORPORATION	            AVI
    Kyocera Corporation	                        KYOCERA
    SHINKO_ARI	                                SHINKO
    Shinko_TKK	                                SHINKO
    Shinko_AIZ	                                SHINKO
    Mitsubishi Materials  Sanda Plant	        MITSUBISHI MATERIALS
    NHK SPRING Asia Transport	                NHK
    KSA INTERNATIONAL INC	                    SHIMADZU
    NICHIAS Corporation	                        NICHIAS
    Kuroda Precision Industries  Asahi Plant	KURODA
    FERROTEC ISHIKAWA	                        FERROTEC
    KAWASAKI HEAVY INDUSTRIES, LTD	            KAWASAKI
    NHK SPRING CO,LTD	                        NHK
    Ferrotec Kansai	                            FERROTEC
    MEIDEN	                                    MEIDEN
    TOTO	                                    INTERNAL
    COORSTEK	                                INTERNAL
    EDWARDS JAPAN LIMITED	                    EDWARDS
                                                OTHER
    # TODO: Update the file to reference server locations, not local locations
    """ 

# XTODO: error handling for element of html_string not being availble after parsing


def main():
    # gather all the paths in the xmls folder into a list
    xml_files_list = [Path(f) for f in Path('xmls').glob('*.xml')]

    # loop through all the file paths in xml_files_list and call parse_xml_file() for each one
    for xml_file_path in xml_files_list:
        (parse_success, html_string, data_dict, data_points, japan_data) = parse_xml_file_to_str(fp=xml_file_path)
        if parse_success:
            if(data_points['origin_country'] == 'JP'):
                # update the japan shipment excel file with the data from the xml file
                japan_updated = japan_shipments_v2(japan_data, xml_file_path=xml_file_path)
                # send an email to sfoexports@transpak.com
                email_sent = send_email(email_body=html_string, data_dict=data_dict, data_points=data_points)
                if email_sent and japan_updated:
                    xml_file_path.rename(Path(Path.cwd(), 'finished', xml_file_path.name))
            else:
                email_sent = send_email(email_body=html_string, data_dict=data_dict, data_points=data_points)
                if email_sent:
                    xml_file_path.rename(Path(Path.cwd(), 'finished', xml_file_path.name))
        else: 
            logging.error(f'Failed to parse {xml_file_path.name}')
            send_error_email(xml_file_path.name, Exception("Error in parsing XML file"))
            break

       
    
def send_email(email_body, data_dict, data_points): 
    email_subject = f"[TESTING] " \
                f"{'EXPRESS' if data_points['service_level'] == 'EX' else ''} " \
                f"{data_dict['Bookings']['AirBooking']['ShipperAddress']['Country']['#text'] if data_dict['Bookings']['AirBooking']['ShipperAddress']['Country']['#text'] != 'US' else ''} " \
                f"{', '+data_dict['Bookings']['AirBooking']['ShipperAddress']['State'] if data_dict['Bookings']['AirBooking']['ShipperAddress']['State'] != 'CA' and data_dict['Bookings']['AirBooking']['ShipperAddress']['Country']['#text'] == 'US' else ''} " \
                f"LAM EDI Shipment " \
                f"[{data_dict['Bookings']['AirBooking']['ShipperName']}] " \
                f"[{data_points['x_number']}]" \
    
    url = "https://api2.frontapp.com/channels/cha_7gt0h/messages"
    payload = {
        "to": ["alex.clifford@transpak.com"],
        "options": {"archive": True},
        "sender_name": "noreply",
        "subject": email_subject,
        "author_id": "tea_5l4v5",
        "body": email_body
    }

    headers = {
            "Accept": "application/json",
            "Content-Type": "application/json",
            # * Replace this bearer token if you want to change the user who sends this message
            "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzY29wZXMiOlsidGltOjYwNTQ2MjUiXSwiaWF0IjoxNjYwNjk2NzQwLCJpc3MiOiJmcm9udCIsInN1YiI6ImE5YTBkNmJkMWQ3MmRiOTgyNzljIiwianRpIjoiYWY1MjcxZWQxOTdiNWNlYyJ9.GfXskctwTtJ-PfFA3w2lfp6pkqK80Vtcryq_onCAjKI"
        }
        
    try:
        response = requests.post(url, json=payload, headers=headers)
        logging.info(f"{data_points['x_number']} Email Sent Successfully [{response}]")
        return True
    except Exception as e:
        logging.error(f"{data_points['x_number']} Email Failed [{response}]: {e}")
        send_error_email(data_dict["xml file"], e)
        return False

def send_error_email(file_path, error_message):
    email_subject = f"Error while processing Lam XML file"
    
    url = "https://api2.frontapp.com/channels/cha_7gt0h/messages"
    payload = {
        "to": ["alex.clifford@transpak.com"],
        "options": {"archive": True},
        "sender_name": "noreply",
        "subject": email_subject,
        "author_id": "tea_5l4v5",
        "body": f"Error while processing XML file. File has not been processed, please process manually or allow file to be processed again.<br><br>File Path: {file_path}<br>Error Message: {error_message}"
    }

    headers = {
            "Accept": "application/json",
            "Content-Type": "application/json",
            # * Replace this bearer token if you want to change the user who sends this message
            "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzY29wZXMiOlsidGltOjYwNTQ2MjUiXSwiaWF0IjoxNjYwNjk2NzQwLCJpc3MiOiJmcm9udCIsInN1YiI6ImE5YTBkNmJkMWQ3MmRiOTgyNzljIiwianRpIjoiYWY1MjcxZWQxOTdiNWNlYyJ9.GfXskctwTtJ-PfFA3w2lfp6pkqK80Vtcryq_onCAjKI"
        }
        
    try:
        response = requests.post(url, json=payload, headers=headers)
        logging.info(f"{file_path} Error Email Sent Successfully [{response}]")
        return True
    except Exception as e:
        logging.error(f"{file_path} Error Email Failed [{response}]: {e}")
        send_error_email("Error in send_error_email()", e)
        return False

def japan_shipments_v2(data, xml_file_path): 
    """
    Check shipper name to see if it falls in a known list of shipper names
        If it does, then we need to check if the SHIPPER file exists
            If it does, then we need to append the data to the file
            If it doesn't, then we need to create the file and add the header row
        If it doesn't, then we need to check that an OTHER file exisits
            If it does, then we need to append the data to the OTHER file
            If it doesn't, then we need to create the file and add the header row
    """

    # Check if the shipper name is in the list of known shipper names
    # if it is, set Shipper Name to the correct file name
    match (data['Shipper Name'].lower()): 
        case "edwards japan":	
            data['Shipper File'] = "Edwards TMS.xlsx"
        case "taica corporation":	
            data['Shipper File'] = "Internal TMS.xlsx"
        case "tdk corporation c/o alps logistics":	
            data['Shipper File'] = "Tdk TMS.xlsx"
        case "tdk corporation":	
            data['Shipper File'] = "Tdk TMS.xlsx"
        case "toto ltd":	
            data['Shipper File'] = "Toto TMS.xlsx"
        case "aviservice c/o tdk corporation":	
            data['Shipper File'] = "Avi TMS.xlsx"
        case "kyocera corporation":	
            data['Shipper File'] = "Kyocera TMS.xlsx"
        case "shinko_ari":	
            data['Shipper File'] = "Shinko TMS.xlsx"
        case "shinko_tkk":	
            data['Shipper File'] = "Shinko TMS.xlsx"
        case "shinko_aiz":	
            data['Shipper File'] = "Shinko TMS.xlsx"
        case "mitsubishi materials sanda plant":	
            data['Shipper File'] = "Mitsubishi Materials TMS.xlsx"
        case "nhk spring asia transport":	
            data['Shipper File'] = "Nhk TMS.xlsx"
        case "ksa international inc":	
            data['Shipper File'] = "Shimadzu TMS.xlsx"
        case "nichias corporation":	
            data['Shipper File'] = "Nichias TMS.xlsx"
        case "kuroda precision industries asahi plant":	
            data['Shipper File'] = "Kuroda TMS.xlsx"
        case "ferrotec ishikawa":	
            data['Shipper File'] = "Ferrotec TMS.xlsx"
        case "kawasaki heavy industries, ltd":	
            data['Shipper File'] = "Kawasaki TMS.xlsx"
        case "nhk spring co,ltd":	
            data['Shipper File'] = "Nhk TMS.xlsx"
        case "ferrotec kansai":	
            data['Shipper File'] = "Ferrotec TMS.xlsx"
        case "meiden":	
            data['Shipper File'] = "Meiden TMS.xlsx"
        case "toto":	
            data['Shipper File'] = "Internal TMS.xlsx"
        case "coorstek":	
            data['Shipper File'] = "Internal TMS.xlsx"
        case "edwards japan limited":	
            data['Shipper File'] = "Edwards TMS.xlsx"
        case _:	
            data['Shipper File'] = "OTHER TMS.xlsx"
    
    dest_filename = Path(Path.cwd(), "JP TMS Spreadsheets", data['Shipper File'])
    if os.path.exists(dest_filename):
        try:
            wb = openpyxl.load_workbook(dest_filename)
            ws = wb.active
            ws.append([data['Date Received'], data['X Number'], data['# of Pieces'], data['Weight'], data['PO #'], data['Express?'], data['FTZ'], data['Departure Date'], data['Consignee']])
        except Exception as e:
            logging.error(f"{data['X Number']} Error on appending to to {data['Shipper File']}: {e}")
            send_error_email(xml_file_path, e)
            return False
    elif not os.path.exists(dest_filename):
        try:
            wb = openpyxl.Workbook()
            wb.save(dest_filename)
            ws = wb.active
            ws.append(['Date Received', 'X Number', '# of Pieces', 'Weight', 'PO #', 'Express?', 'FTZ', 'Departure Date', 'Consignee', 'Invoice', 'HAWB #', "Received Pre Alert"])
            ws.append([data['Date Received'], data['X Number'], data['# of Pieces'], data['Weight'], data['PO #'], data['Express?'], data['FTZ'], data['Departure Date'], data['Consignee']])
        except Exception as e:
            logging.error(f"{data['X Number']} Error on creating and appending to {data['Shipper File']}: {e}")
            send_error_email(xml_file_path, e)
            return False

    try:
        wb.save(dest_filename)
        return True
    except PermissionError as e: 
        # TODO: When a file is in use, don't move it to completed
        logging.warning(f"{data['X Number']} File {data['Shipper File']} in use, will attempt to save upon next cycle: {e}")
        send_error_email(xml_file_path, e)
        return False
    except Exception as e:
        # TODO: When a japan file is completed, move it to completed
        logging.error(f"{data['X Number']} General error on saving {data['Shipper File']}: {e}")
        send_error_email(xml_file_path, e)
        return False

def fix_addresses(address1, city, state, zip_code, country, fixed):
    # Creats a slug of what the current address is
    addslug = address1 + city + state + zip_code + country

    # uses a switch to check where the case is a known bad address and the return sends back the fixed address
    # * New Addresses can be added in by adding a new case and return
    match (addslug.lower()): 
        case "6551tracyca95377us":
            return "6551 W SCHULTE ROAD", "TRACY", "CA", "95377", "US", addslug.lower()
        case "6753newarkca94560us":
            return "6753 MOWRY AVE", "NEWARK", "CA", "94560", "US", addslug.lower()
        case "1201livermoreca94551jp":
            return "1201 VOYAGER STREET", "LIVERMORE", "CA", "94551", "US", addslug.lower()
        case "1201livermoreca94551us":
            return "1201 VOYAGER STREET", "LIVERMORE", "CA", "94551", "US", addslug.lower()
        case "21000tualatinor97062us":
            return "21000 SW 115TH AVE", "TUALATIN", "OR", "97062", "US", addslug.lower()
        case "7826osansi, gyeonggido0918145kr":
            return "#78-26 GYEONGGI-BAERO", "OSANSI, GYEONGGIDO", "KR", "18145", "KR", addslug.lower()
        case "7826osansi, gyeonggido0918145jp":
            return "#78-26 GYEONGGI-BAERO", "OSANSI,GYEONGGIDO", "KR", "18145", "KR", addslug.lower()
        case _:
            return address1, city, state, zip_code, country, ""

def parse_xml_file_to_str(fp):
    """"
    Parses an XML file (fp) into a data dictionary, pushes that data into an html-formatted string to be emailed later. 
    Also collects general useful data and japan-related shipment data into seperate dictionaries.
    Returns everything into main(). 
    
    """
    html_string = ""
    data_dict = {}
    data_points = {}
    japan_data = {}

    try:
        with open(fp) as xml_file:
            data_dict = xmltodict.parse(xml_file.read())
    except Exception as e:
        logging.error(f"{xml_file.name} Error upon reading XML Data: {e}")
        send_error_email(fp.name, e)
        return False, html_string, data_dict, data_points, japan_data

    # add number from data_dict to data_points dict
    data_dict['xml file'] = fp
    data_points['x_number'] = data_dict['Bookings']['AirBooking']['Number']

    try:
        html_string += ('Hello, <br><br>New shipment added. <br><br>')
        # # outputs the entire xml in json format
        html_string += ("<table style='width:40%' border='1'>")
        html_string += (f"<tr><td style='width:40%'>X #            </td><td>{data_dict['Bookings']['AirBooking']['Number']}</td></tr>")
        html_string += (f"<tr><td>Shipper        </td><td>{data_dict['Bookings']['AirBooking']['ShipperName']}</td></tr>")

        # call fix_addresses() to fix the addresses of Street, City, State, ZipCode and Country

        shipadd_fixed, conadd_fixed = "", ""
        (data_dict['Bookings']['AirBooking']['ShipperAddress']['Street'], 
        data_dict['Bookings']['AirBooking']['ShipperAddress']['City'],
        data_dict['Bookings']['AirBooking']['ShipperAddress']['State'],
        data_dict['Bookings']['AirBooking']['ShipperAddress']['ZipCode'],
        data_dict['Bookings']['AirBooking']['ShipperAddress']['Country']['#text'], 
        shipadd_fixed) = fix_addresses(data_dict['Bookings']['AirBooking']['ShipperAddress']['Street'], 
                                                                                            data_dict['Bookings']['AirBooking']['ShipperAddress']['City'],
                                                                                            data_dict['Bookings']['AirBooking']['ShipperAddress']['State'],
                                                                                            data_dict['Bookings']['AirBooking']['ShipperAddress']['ZipCode'],
                                                                                            data_dict['Bookings']['AirBooking']['ShipperAddress']['Country']['#text'],
                                                                                            shipadd_fixed)

        (data_dict['Bookings']['AirBooking']['ConsigneeAddress']['Street'], 
        data_dict['Bookings']['AirBooking']['ConsigneeAddress']['City'],
        data_dict['Bookings']['AirBooking']['ConsigneeAddress']['State'],
        data_dict['Bookings']['AirBooking']['ConsigneeAddress']['ZipCode'],
        data_dict['Bookings']['AirBooking']['ConsigneeAddress']['Country']['#text'],
        conadd_fixed) = fix_addresses(data_dict['Bookings']['AirBooking']['ConsigneeAddress']['Street'], 
                                                                                            data_dict['Bookings']['AirBooking']['ConsigneeAddress']['City'],
                                                                                            data_dict['Bookings']['AirBooking']['ConsigneeAddress']['State'],
                                                                                            data_dict['Bookings']['AirBooking']['ConsigneeAddress']['ZipCode'],
                                                                                            data_dict['Bookings']['AirBooking']['ConsigneeAddress']['Country']['#text'],
                                                                                            conadd_fixed)


        for index, value in enumerate(data_dict['Bookings']['AirBooking']['ShipperAddress'].items()):
            if value[0] == 'Country':
                data_points['origin_country'] = value[1]['#text']
                html_string += (f"<tr><td></td><td>{value[1]['#text']}</td></tr>")
                if value[1]['#text'] == 'JP':
                    japan_data['Shipper Name'] = data_dict["Bookings"]["AirBooking"]["ShipperName"]
                    japan_data['Date Received'] = datetime.strptime(data_dict["Bookings"]["AirBooking"]["CreatedOn"], "%Y-%m-%dT%H:%M:%S").strftime('%Y-%m-%d')
                    japan_data['X Number'] = data_dict['Bookings']['AirBooking']['Number']
                    japan_data['# of Pieces'] = data_dict['Bookings']['AirBooking']['TotalPieces']
                    japan_data['Weight'] = f"{data_dict['Bookings']['AirBooking']['TotalWeight']['#text']} {data_dict['Bookings']['AirBooking']['TotalWeight']['@Unit']}"
                    # japan_data['PO #'] = hanlded at end during POs
                    # japan_data[, 'Express'] = handled during custom data section
                    # japan_data[, 'FTZ'] = hanlded during custom data section
                    japan_data['Departure Date'] = datetime.strptime(data_dict['Bookings']['AirBooking']['EstimatedDepartureDate'], '%Y-%m-%dT%H:%M:%S').strftime('%Y-%m-%d')
                    # japan_data[, 'Consignee'] = hanlded after this block
            elif value[0] == 'ContactName':
                html_string += (f"<tr><td>Ship Contact{' ' * 3}</td><td>{value[1]}</td></tr>")
            else:
                html_string += (f"<tr><td>{' ' * 15}</td><td>{value[1]}</td></tr>")
        html_string += (f"<tr><td>Consignee      </td><td>{data_dict['Bookings']['AirBooking']['ConsigneeName']}</td></tr>")
        consignee = ''
        for index, value in enumerate(data_dict['Bookings']['AirBooking']['ConsigneeAddress'].items()):
            if value[0] == 'Country':
                html_string += (f"<tr><td>{' ' * 15}</td><td>{value[1]['#text']}</td></tr>")
                consignee += f"{value[1]['#text']}, "
            elif value[0] == 'ContactName':
                html_string += (f"<tr><td>Con Contact{' ' * 4}</td><td>{value[1]}</td></tr>")
            else:
                html_string += (f"<tr><td>{' ' * 15}</td><td>{value[1]}</td></tr>")
                consignee += f"{value[1]}, "

        japan_data['Consignee'] = consignee[:-2]

        # format date and time in pretty print
        html_string += (f"<tr><td>Est Departure  </td><td>{datetime.strptime(data_dict['Bookings']['AirBooking']['EstimatedDepartureDate'], '%Y-%m-%dT%H:%M:%S').strftime('%Y-%m-%d')}</td></tr>")

        for index, value in enumerate(data_dict['Bookings']['AirBooking']['CustomFields']['CustomField']):
            if(value['CustomFieldDefinition']['InternalName'] == 'edi_service_Level'):
                data_points['service_level'] = value['Value']
                html_string += (f"<tr><td>Service Level{' ' * 2}</td><td>{value['Value']}</td></tr>")
                # if value['Value'] == 'EX' then japan_data['Express?'] = Yes
                if value['Value'] == 'EX':
                    japan_data['Express?'] = 'EX'
                else:
                    japan_data['Express?'] = ''
            elif(value['CustomFieldDefinition']['InternalName'] == 'raterequestonly'):
                # html_string += (f"<tr><td>Quote Only{' ' * 5}</td><td>{value['Value']}</td></tr>")
                pass
            elif(value['CustomFieldDefinition']['InternalName'] == 'ftz_flag'):
                if value['Value'] == 'FTZ':
                    data_points['ftz_flag'] = 'FTZ'
                else:
                    data_points['ftz_flag'] = 'Non-FTZ'
                html_string += (f"<tr><td>FTZ{' ' * 12}</td><td>{data_points['ftz_flag']}</td></tr>")
                japan_data['FTZ'] = data_points['ftz_flag']
            elif(value['CustomFieldDefinition']['InternalName'] == 'billing_party'):
                # html_string += (f"<tr><td>Billing Party{' ' * 2}</td><td>{value['Value']}</td></tr>")
                pass
            else:
                pass

        # output TotalPieces
        html_string += (f"<tr><td>Total Pieces    </td><td>{data_dict['Bookings']['AirBooking']['TotalPieces']} pieces @ {data_dict['Bookings']['AirBooking']['TotalWeight']['#text']} {data_dict['Bookings']['AirBooking']['TotalWeight']['@Unit']}</td></tr>")
        po_numbers = []

        # if data_dict['Bookings']['AirBooking']items item is an array, loop through each item in the array and output the PO number
        Item_s = data_dict['Bookings']['AirBooking']['Items']['Item']
        if isinstance(Item_s, list):
            for value in Item_s:
                # po_numbers.append(value['SupplierPONumber'])
                html_string += (f"<tr><td>{' ' * 16}</td><td>{value['Length']['#text']} {value['Length']['@Unit']} X {value['Width']['#text']} {value['Width']['@Unit']} X {value['Height']['#text']} {value['Height']['@Unit']}, ")
                html_string += (f"{value['Weight']['#text']} {value['Weight']['@Unit']}</td></tr>")
                # split the comma separated PO numbers into a list
                po_numbers.append(value['SupplierPONumber'].split(','))
        else:
            html_string += (f"<tr><td>{' ' * 16}</td><td>{Item_s['Length']['#text']} {Item_s['Length']['@Unit']} X {Item_s['Width']['#text']} {Item_s['Width']['@Unit']} X {Item_s['Height']['#text']} {Item_s['Height']['@Unit']}, ")
            html_string += (f"{Item_s['Weight']['#text']} {Item_s['Weight']['@Unit']}</td></tr>")
            po_numbers.append(Item_s['SupplierPONumber'].split(','))

        # remove duplicates from the po_numbers list
        # hard copy into japan_data['PO #'] from po_numbers
        po_numbers = list(set(sum(po_numbers, [])))

        data_points['po_numbers'] = []
        japan_data['PO #'] = []
        string_of_po_numbers = ''

        for ele in po_numbers:
            string_of_po_numbers += f"{ele}, "
            data_points['po_numbers'].append(ele)
        string_of_po_numbers = string_of_po_numbers[:-2]
        japan_data['PO #'] = string_of_po_numbers

        # output the first PO number in pretty print
        # in order to get the formatting right, have to pop the first PO number
        # html_string += (f"<tr><td>PO Numbers{' ' * 6}</td><td>{po_numbers.pop()}</td></tr>")
        html_string += (f"<tr><td>PO Numbers{' ' * 6}</td><td>{string_of_po_numbers}</td></tr>")
        # loop through po numbers in pretty print format with a new line for each entry with the first entry titled PO Number
        # for index, value in enumerate(po_numbers):
        #     html_string += (f"<tr><td>{' ' * 16}</td><td>{value}</td></tr>")
        html_string += ("</table><br>")

        # output Success Message if this was reached
        html_string += ("Translation Successful<br>")
        # output the created date and time
        # output data_file booking airbookings CreatedOn date and time in pretty print
        html_string += ('=================<br><br>')
        html_string += (f'Uploaded on{" " * 4}: {datetime.strptime(data_dict["Bookings"]["AirBooking"]["CreatedOn"], "%Y-%m-%dT%H:%M:%S")}<br>')
        html_string += (f'Processed      : {datetime.utcnow()} (UTC)<br>')
        if(shipadd_fixed):
            html_string += (f'Original Ship Address: {shipadd_fixed}<br>')
        if(conadd_fixed):
            html_string += (f'Original Con Address: {conadd_fixed}<br>')

        return True, html_string, data_dict, data_points, japan_data

    except Exception as e:
        logging.error(f"{data_points['x_number']} Error while building email body: {e}")
        send_error_email(data_dict["xml file"], e)
        return False, html_string, data_dict, data_points, japan_data


if __name__ == '__main__':
    main()
