from bs4 import BeautifulSoup
import requests
import xlwings as xw
import pandas as pd
import openpyxl

table = {'Наименование': [], 'Цена': [], 'Производитель': [], 'Москва': [], 'Санкт-Петербург': [], 'Новосибирск': [], 'Екатеринбург': [], 'Товары в пути': [], 'Замена от производителя': [], 'Доступно для резерва': []}

Programmable_controllers = "https://nnz-ipc.ru/catalogue/automation/controllers/"
Signal_input_output_modules = "https://nnz-ipc.ru/catalogue/automation/io/"
Signal_I_O_boards = "https://nnz-ipc.ru/catalogue/automation/io_boards/"
Modular_I_O_cages = "https://nnz-ipc.ru/catalogue/automation/extension_chassis/"
Specialized_modules = "https://nnz-ipc.ru/catalogue/automation/specialized_modules/"
Specialized_boards = "https://nnz-ipc.ru/catalogue/automation/specialized_boards/"
Operator_panels_HMI = "https://nnz-ipc.ru/catalogue/automation/hmi/"
IP_video_surveillance = "https://nnz-ipc.ru/catalogue/automation/video/"
Measurement_systems = "https://nnz-ipc.ru/catalogue/automation/remote_data_logger/"
L_card_equipment = "https://nnz-ipc.ru/catalogue/automation/lcard/"
Automation_accessories = "https://nnz-ipc.ru/catalogue/automation/automation_accesories/"
Industrial_Ethernet = "https://nnz-ipc.ru/catalogue/comm/ethernet/"
COM_Port_to_Ethernet_Converters = "https://nnz-ipc.ru/catalogue/comm/serial_to_ethernet/"
Multiport_RS_232_422_485_and_CAN_cards = "https://nnz-ipc.ru/catalogue/comm/serial/"
Wireless_access = "https://nnz-ipc.ru/catalogue/comm/wireless/"
Converters_and_repeaters = "https://nnz-ipc.ru/catalogue/comm/converters/"
Software = "https://nnz-ipc.ru/catalogue/comm/software/"
Accessories_for_switches = "https://nnz-ipc.ru/catalogue/comm/network_accesories/"
Accessories = "https://nnz-ipc.ru/catalogue/comm/accessories/"
Remote_Access_Tools = "https://nnz-ipc.ru/catalogue/dc_equipment/remote_access/"
Power_distribution = "https://nnz-ipc.ru/catalogue/dc_equipment/power_distribution/"
Dc_accessories = "https://nnz-ipc.ru/catalogue/dc_equipment/dc_accessories/"

ACS = [Programmable_controllers, Signal_input_output_modules, Signal_I_O_boards, Modular_I_O_cages, Specialized_modules, Specialized_boards, Operator_panels_HMI, IP_video_surveillance, Measurement_systems,
L_card_equipment, Automation_accessories, Industrial_Ethernet, COM_Port_to_Ethernet_Converters, Multiport_RS_232_422_485_and_CAN_cards, Wireless_access, Converters_and_repeaters, Software, Accessories_for_switches, Accessories,
Remote_Access_Tools, Power_distribution, Dc_accessories]

entry_count = 0

all_cities = set(['Москва', 'Санкт-Петербург', 'Новосибирск', 'Екатеринбург', 'Товары в пути'])
end_of_life = set(['Замена от производителя'])
all_reserved = set(['Доступно для резерва'])

def add_data(data, values, plaques):
    for i in range(len(data)):
        table[data[i]].append(values[i])
    for plaque in plaques.difference(set(data)):
        table[plaque].append('0')

def process_data(raw_data, plaques):
    if raw_data is not None:
        counts = raw_data.find_all("li")
        names = []
        values = []
        for data in counts:
            name = str(data.find("span").text)
            count = str(data.find("b").text)
            names.append(name)
            values.append(count)
        add_data(names, values, plaques)
    else:
        add_data([], [], plaques)

for ur in ACS:
    url = ur + "?pa=100"
    request = requests.get(url)

    soup = BeautifulSoup(request.text, "html.parser")

    entries = soup.find_all("div", class_ = "catcard_entry")
    if len(entries) == 0:
        break

    for entry in entries:
        entry_count = entry_count+1
        name = entry.find("meta", {"itemprop":"name"})
        vendor = entry.find("span", {"itemprop":"brand"})
        price = entry.find("div", {"class":"catcard_price_strong"})
        stock = entry.find("div", {"class":"infobox availbox_popup cs_popup"})
        reserved = entry.find("div", {"class":"infobox infobox-road availbox_popup cs_popup"})
        eol = entry.find("div", {"class":"infobox infobox-single infobox-outprod availbox_popup cs_popup"})

        process_data(reserved, all_reserved)
        process_data(eol, end_of_life)
        process_data(stock, all_cities)

        table['Наименование'].append(str(name.get("content")))
        table['Производитель'].append(str(vendor.text))
        if price != None:
            table['Цена'].append(str(price.text))
        else:
            table['Цена'].append('нет цены')

print(entry_count)
df = pd.DataFrame(table)
df.to_excel('parsing_kalek.xlsx')
