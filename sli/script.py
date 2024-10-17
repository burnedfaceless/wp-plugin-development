import pymysql
from openpyxl import load_workbook
from datetime import datetime

# Database connection setup
def get_db_connection():
    connection = pymysql.connect(
        host='aquacertify.cq7rdlwwg2mm.us-east-1.rds.amazonaws.com',
        user='admin',
        password='MBfFzNPkXP2BCCrCyYSgenfnk',
        db='aquacertify',
        cursorclass=pymysql.cursors.DictCursor
    )
    return connection

# Get addresses by system ID
def get_addresses_by_system_id(system_id):
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            sql = """
                SELECT id, 
                       street_address_1 AS streetAddress1, 
                       street_address_2 AS streetAddress2, 
                       city, state, zip, 
                       `120water_location_id` AS oneTwentyWaterLocationId 
                FROM sli_ga_addresses 
                WHERE system_id = %s
            """
            cursor.execute(sql, (system_id,))
            addresses = cursor.fetchall()
        return addresses
    except Exception as e:
        print(f"Error retrieving addresses: {e}")
        return None
    finally:
        connection.close()

# Get system by ID
def get_system_by_id(system_id):
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            sql = "SELECT * FROM systems WHERE id = %s"
            cursor.execute(sql, (system_id,))
            system = cursor.fetchone()
        return system
    except Exception as e:
        print(f"Error retrieving system: {e}")
        return None
    finally:
        connection.close()

# Get all values by address ID
def get_all_values_by_address_id(address_id):
    connection = get_db_connection()
    try:
        with connection.cursor() as cursor:
            sql = """
                SELECT 
                    sli_ga_service_lines.*,
                    slo.response AS serviceLineOwnershipType,
                    som.response AS systemOwnedMaterial,
                    sosm.response AS systemOwnedSpecificMaterial,
                    sob.response AS systemOwnedBasisOfClassification,
                    soid.response AS systemOwnedInstallationDate,
                    pl.response AS presenceOfLeadConnector,
                    com.response AS customerOwnedMaterial,
                    cos.response AS customerOwnedSpecificMaterial,
                    cob.response AS customerOwnedBasisOfClassification,
                    coid.response AS customerOwnedInstallationDate,
                    osc.response AS overallServiceLineClassification,
                    bt.response AS buildingType,
                    scc.response AS schoolOrChildcare
                FROM sli_ga_service_lines
                LEFT JOIN sli_ga_service_line_ownership_type_responses AS slo
                    ON sli_ga_service_lines.service_line_ownership_type_id = slo.id
                LEFT JOIN sli_ga_system_owned_material_responses AS som
                    ON sli_ga_service_lines.system_owned_material_id = som.id
                LEFT JOIN sli_ga_specific_service_line_material_responses AS sosm
                    ON sli_ga_service_lines.system_owned_specific_material_id = sosm.id
                LEFT JOIN sli_ga_basis_of_classification_responses AS sob
                    ON sli_ga_service_lines.system_owned_basis_id = sob.id
                LEFT JOIN sli_ga_service_line_installation_date_responses AS soid
                    ON sli_ga_service_lines.system_owned_installation_date_id = soid.id
                LEFT JOIN sli_ga_presence_of_lead_connector_responses AS pl
                    ON sli_ga_service_lines.presence_of_lead_connector_id = pl.id
                LEFT JOIN sli_ga_customer_owned_material_responses AS com
                    ON sli_ga_service_lines.customer_owned_material_id = com.id
                LEFT JOIN sli_ga_specific_service_line_material_responses AS cos
                    ON sli_ga_service_lines.customer_owned_specific_material_id = cos.id
                LEFT JOIN sli_ga_basis_of_classification_responses AS cob
                    ON sli_ga_service_lines.customer_owned_basis_id = cob.id
                LEFT JOIN sli_ga_service_line_installation_date_responses AS coid
                    ON sli_ga_service_lines.customer_owned_installation_date_id = coid.id
                LEFT JOIN sli_ga_overall_service_line_classification_responses AS osc
                    ON sli_ga_service_lines.overall_service_line_classification_id = osc.id
                LEFT JOIN sli_ga_building_type_responses AS bt
                    ON sli_ga_service_lines.building_type_id = bt.id
                LEFT JOIN sli_ga_school_or_childcare_facility_responses AS scc
                    ON sli_ga_service_lines.school_or_childcare_id = scc.id
                WHERE sli_ga_service_lines.sli_ga_address_id = %s
            """
            cursor.execute(sql, (address_id,))
            service_lines = cursor.fetchall()
        return service_lines
    except Exception as e:
        print(f"Error retrieving service lines: {e}")
        return None
    finally:
        connection.close()

# Main function to update the Excel file
def main():
    try:
        system_id = 3

        # Get system data
        system = get_system_by_id(system_id)
        if not system:
            print(f"No system found with ID {system_id}")
            return

        # Get addresses data
        addresses = get_addresses_by_system_id(system_id)
        if not addresses:
            print(f"No addresses found for system ID {system_id}")
            return

        # Load the existing Excel file
        workbook = load_workbook('georgia-import.xlsx')
        worksheet = workbook['GA Detailed SL Inventory']

        # Update specific cells
        worksheet['C3'] = system['name'] or 'Unknown'
        worksheet['C6'] = datetime.today().strftime('%m/%d/%Y')

        for i, address in enumerate(addresses, start=15):
            service_line = get_all_values_by_address_id(address['id'])
            if not service_line or not service_line[0]:
                print(f"No service line found for address ID: {address['id']}")
                continue

            print(address)

            service_line = service_line[0]  # Take the first service line result


            worksheet[f'B{i}'] = service_line['id'] or 'Unknown'
            worksheet[f'C{i}'] = address['streetAddress1'] or ''
            worksheet[f'D{i}'] = address['streetAddress2'] or ''
            worksheet[f'E{i}'] = address['city'] or ''
            worksheet[f'F{i}'] = address['state'] or ''
            worksheet[f'G{i}'] = address['zip'] or ''
            worksheet[f'I{i}'] = service_line['serviceLineOwnershipType'] or ''
            worksheet[f'J{i}'] = service_line['systemOwnedMaterial'] or ''
            worksheet[f'K{i}'] = service_line['systemOwnedSpecificMaterial'] or ''
            worksheet[f'L{i}'] = service_line['system_owned_material_additional_information'] or ''
            worksheet[f'M{i}'] = service_line['systemOwnedBasisOfClassification'] or ''
            worksheet[f'N{i}'] = service_line['system_owned_basis_additional_notes'] or ''
            worksheet[f'O{i}'] = service_line['systemOwnedInstallationDate'] or ''
            worksheet[f'Q{i}'] = service_line['presenceOfLeadConnector'] or ''
            worksheet[f'R{i}'] = service_line['customerOwnedMaterial'] or ''
            worksheet[f'S{i}'] = service_line['customerOwnedSpecificMaterial'] or ''
            worksheet[f'T{i}'] = service_line['customer_owned_material_additional_information'] or ''
            worksheet[f'U{i}'] = service_line['customerOwnedBasisOfClassification'] or ''
            worksheet[f'V{i}'] = service_line['customer_owned_basis_additional_notes'] or ''
            worksheet[f'W{i}'] = service_line['customerOwnedInstallationDate'] or ''
            worksheet[f'Z{i}'] = service_line['buildingType'] or ''
            worksheet[f'AB{i}'] = service_line['schoolOrChildcare'] or ''
            worksheet[f'AF{i}'] = address['oneTwentyWaterLocationId'] or ''

        # Save the updated Excel file
        workbook.save('larchmont.xlsx')
        print('File updated successfully')

    except Exception as e:
        print(f"Error updating Excel file: {e}")

if __name__ == '__main__':
    main()