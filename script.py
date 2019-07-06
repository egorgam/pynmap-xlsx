import sys, json, time, xmltodict, subprocess, xlsxwriter


def get_addr(i):
    if isinstance(i['hostnames']['hostname'], list):
        addr = [h for h in i['hostnames']['hostname'] if h['type'] == 'user'][0]['name']

    else:
        addr = i['hostnames']['hostname']['name']

    return addr


def get_target(k, i, row, states):
    print(1)
    col = -1
    res = {'ip': i['address']['addr'], 'hostname': get_addr(i), 'portid': k['portid'], 'protocol': k['protocol'],
           'portname': k['service']['name'], 'state': k['state']['state'], 'status': 'up'}

    if k['state']['state'] == 'open': 
        states.append(res)
        
    for v in res.values():
        col += 1
        worksheet.write(row, col, v, cell_format)


def get_no_result(i, row):
    res = {'ip': i['address']['addr'], 'hostname': '', 'portid': '', 'protocol': '', 'portname': '', 'state': '', 'status': 'up'}
    col = -1
    for v in res.values():
        col += 1
        worksheet.write(row, col, v, cell_format)
    row += 1


def get_result(i, row, states):
    if isinstance(i['ports']['port'], list):
        for k in i['ports']['port']:
            get_target(k, i, row, states)
            row += 1
            
    else:
        get_target(i['ports']['port'], i, row, states)
        row += 1

    return row


def query(i, row):
    states = []
    if 'port' not in i['ports']:
        row += 1
        i['ports']['port'] = []

    first_row = row + 1
    row = get_result(i, row, states)
    res = {'ip': 'Opened ports: ' + str(len(states)), 'hostname': '', 'portid': '', 'protocol': '', 'portname': '',
           'state': '', 'status': ''}
    col = -1

    for v in res.values():
        col += 1
        worksheet.write(row, col, v, cell_format_res)

    row += 1
    worksheet.merge_range('A' + str(first_row) + ':' + 'A' + str(row - 1), i['address']['addr'], merge_format)
    worksheet.merge_range('B' + str(first_row) + ':' + 'B' + str(row - 1), get_addr(i), merge_format)
    worksheet.merge_range('G' + str(first_row) + ':' + 'G' + str(row - 1), 'up', merge_format)

    if i['ports']['port'] == []:
        worksheet.merge_range('C' + str(first_row) + ':' + 'C' + str(row - 1), '', merge_format)
        worksheet.merge_range('D' + str(first_row) + ':' + 'D' + str(row - 1), '', merge_format)
        worksheet.merge_range('E' + str(first_row) + ':' + 'E' + str(row - 1), '', merge_format)
        worksheet.merge_range('F' + str(first_row) + ':' + 'F' + str(row - 1), '', merge_format)

    return row

try:
    address = sys.argv[1]
    nmap = subprocess.check_output(['nmap', '-oX', '-', address])
except Exception as e:
    print('wrong input paremeters!')
    print("{}: {}".format(e.__class__.__qualname__, e))
    exit(10)

json_nmap_output = json.dumps(dict(xmltodict.parse(nmap)['nmaprun']), separators=(',', ': ')).replace('@', '')
dict_nmap_output = json.loads(json_nmap_output)

workbook = xlsxwriter.Workbook('nmap_' + address.replace('/','_') + '_' + str(time.strftime("%Y-%m-%d-%H.%M", time.localtime())) + '.xlsx')
worksheet = workbook.add_worksheet()

worksheet.set_column('A:A', 18)
worksheet.set_column('B:B', 40)
worksheet.set_column('D:D', 18)
worksheet.autofilter('A1:G11')

merge_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#E0E0E0'})

merge_format2 = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'bold': True,
        'border': 1,
        'bg_color': '#CCE5FF'})

cell_format = workbook.add_format({'border': 1, 'bg_color': '#E0E0E0'})
cell_format_res = workbook.add_format({'border': 1})


row = 0
col = -1
for HeadElement in ['ip', 'hostname', 'portid', 'protocol', 'portname', 'state', 'status']:
    col += 1
    worksheet.write(row, col, HeadElement, workbook.add_format({'bold': True, 'border': 1}))

row += 1

if dict_nmap_output['runstats']['hosts']['total'] == dict_nmap_output['runstats']['hosts']['down']:
    get_no_result(address, row)
else:
    if isinstance(dict_nmap_output['host'], list):
        for i in dict_nmap_output['host']:
            row = query(i, row)

    else:
        row = query(dict_nmap_output['host'], row)


s = 'ADDRS:' + str(dict_nmap_output['runstats']['hosts']['total']) + ' UNKNOWN/DOWN :' + str(dict_nmap_output['runstats']['hosts']['down']) + ' UP:' + str(dict_nmap_output['runstats']['hosts']['up'])
row += 1

worksheet.merge_range('A' + str(row) + ':' + 'G' + str(row), s, merge_format2)
print(dict_nmap_output['runstats']['hosts'])
workbook.close()
