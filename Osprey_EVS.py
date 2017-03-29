# -*- coding: cp950 -*-

import os
import re
from decimal import Decimal
from myXlsx import *

pe_criteria_csv = 'OspreyA0_170313_V01C_DLY.csv'

yield_monitor = ['DLY']
hlv_pass_criteria = False  # True = 卡兩邊
blocks = ['CPU_LOGIC', 'CPU_MBIST', 'SOC_LOGIC', 'SOC_MBIST', 'M9_LOGIC', 'M9_MBIST']
max_cpu_core = 16
max_soc_core = 4
max_m9_core = 4

ti_logic_mbist = {}
ti_block = {}
ti_block_core = {}
ti_yield = {}

all_ti = []
all_ti_hlv = []

hlv_mapping = {}  # SocSaChain_DM_HV/LV -> SocSaChain_DM_HLV
#  ==============================================================
with open(pe_criteria_csv, 'r') as lines:
    for line in lines:
        s = line.split(',')
        if len(s) < 5:
            continue
        if s[4] != '1':  # Test Flag
            continue

        yield_info = s[0].upper()
        yield_type = s[6].upper()

        block_info = s[1].upper()
        core_info = s[2].upper()
        test_instance = s[3].upper()

        if re.search(r'DC|HardIP|HIP', yield_info, re.I):
            continue

        cur_block = ''
        cur_block_core = ''
        if re.search(r'_B\d+', core_info):
            m = re.match(r'.+B(\d+)', core_info)
            if m:
                cur_block_core = m.group(1)  # print m.group(1)
        elif re.search(r'_SOC\d', core_info):
            m = re.match(r'.+SOC(\d)', core_info)
            if m:
                cur_block_core = m.group(1)  # print m.group(1)
        elif re.search(r'_M9X\d', core_info):
            m = re.match(r'.+M9X(\d)', core_info)
            if m:
                cur_block_core = m.group(1)  # print m.group(1)

        if re.search(r'CPU', block_info):
            cur_block = 'CPU_LOGIC'
            if re.search(r'BIST', block_info):
                cur_block = 'CPU_MBIST'
        elif re.search(r'SOC', block_info):
            cur_block = 'SOC_LOGIC'
            if re.search(r'BIST', block_info):
                cur_block = 'SOC_MBIST'
        elif re.search(r'M9', block_info):
            cur_block = 'M9_LOGIC'
            if re.search(r'BIST', block_info):
                cur_block = 'M9_MBIST'
        else:
            cur_block = block_info  # Repair

        if test_instance not in all_ti_hlv:
            all_ti_hlv.append(test_instance)
            ti_block[test_instance] = cur_block
            ti_yield[test_instance] = yield_type
            ti_logic_mbist[test_instance] = yield_info
            if cur_block_core != '':
                ti_block_core[test_instance] = '%s_B%d' % (cur_block, int(cur_block_core))

        tis = []
        if re.search(r'_HLV|_LHV|_MHLV|_MLHV', test_instance):
            t = test_instance.split('_')
            if re.search(r'_HLV|_LHV', test_instance):
                tis = ['_'.join(t[:-1]) + '_HV', '_'.join(t[:-1]) + '_LV']
            else:
                tis = ['_'.join(t[:-1]) + '_MHV', '_'.join(t[:-1]) + '_MLV']
        else:
            tis.append(test_instance)

        for ti in tis:
            if ti in all_ti:
                continue
            all_ti.append(ti)
            hlv_mapping[ti] = test_instance
            ti_block[ti] = cur_block
            ti_yield[ti] = yield_type
            ti_logic_mbist[ti] = yield_info
            if cur_block_core != '':
                ti_block_core[ti] = '%s_B%d' % (cur_block, int(cur_block_core))
#  ==============================================================

seq = {}  # real_lot_name -> read_point -> timestamp:fileName

for log in [f for f in os.listdir(os.curdir) if re.search(r'STDF.+\.csv', f)]:
    dum = log.replace('.csv', '').split('_')
    read_point = Decimal(dum[-1].replace('T', ''))
    # pure_lot_name = ''
    lotId = ''
    timestamp = ''

    for i in dum:
        if re.match(r'\d+\w\w\w\w\w\w\.\w+', i):  # always \d\dNxxxxx.xx, ex: 01N9XM93.B1
            # pure_lot_name = re.sub(r'^\d+|\.\w+', '', i)
            lotId = i
            # log_stage = 'FT'
        elif re.match(r'\w\w\w\w\w\w-\w+', i):
            lotId = i
            # pure_lot_name = re.sub(r'^\d+|\.\w+|-\w+', '', i)  # ex: N9XM93-12
        elif re.match(r'\d\d\d\d\d\d\d\d\d\d\d\d', i):
            timestamp = i

    if lotId not in seq.keys():
        seq[lotId] = {}
    if read_point not in seq[lotId]:
        seq[lotId][read_point] = []
    seq[lotId][read_point].append('%s:%s' % (timestamp, log))
    # print real_lot_name, pure_lot_name, read_point, timestamp
#  ==============================================================

all_wxy = {}
summary = {}

for lotId in sorted(seq.keys()):
    print lotId
    all_wxy[lotId] = {}  #

    for read_point in sorted(seq[lotId].keys()):
        print ' => Read Point T%d:' % read_point
        for timeLog in sorted(seq[lotId][read_point]):
            timestamp = timeLog.split(':')[0]
            log = timeLog.split(':')[1]
            print '  => Processing... Time Stamp:', timestamp, ', Log:', log

            with open(log, 'r') as lines:
                for line in lines:
                    s = line.split(',')

                    if len(s) < 9:
                        if not re.search(r'^#\d', line):
                            continue
                        if s[1] == '0':
                            continue
                        wxy = '%s:%s:%s' % (s[1], s[3], s[4])  # WID:X:Y
                        if wxy in all_wxy[lotId].keys():
                            if 'Bin' not in all_wxy[lotId][wxy].keys():
                                all_wxy[lotId][wxy]['Bin'] = {}
                            if read_point not in all_wxy[lotId][wxy]['Bin'].keys():
                                all_wxy[lotId][wxy]['Bin'][read_point] = s[5]
                            all_wxy[lotId][wxy]['Bin'][read_point] = s[5]
                        continue

                    test_instance = s[8].upper()

                    if test_instance not in all_ti:
                        continue

                    wxy = '%s:%s:%s' % (s[3], s[1], s[2])
                    if wxy not in all_wxy[lotId].keys():
                        all_wxy[lotId][wxy] = {}

                    test_result = s[-2][0]  # P or F

                    if test_instance not in all_wxy[lotId][wxy].keys():
                        all_wxy[lotId][wxy][test_instance] = {}

                    test_instance_hlv = hlv_mapping[test_instance]
                    if test_instance_hlv not in all_wxy[lotId][wxy].keys():
                        all_wxy[lotId][wxy][test_instance_hlv] = {}

                    if read_point not in all_wxy[lotId][wxy][test_instance].keys():
                        all_wxy[lotId][wxy][test_instance][read_point] = test_result
                    elif test_result == 'F':
                        all_wxy[lotId][wxy][test_instance][read_point] = 'F'

                    if read_point not in all_wxy[lotId][wxy][test_instance_hlv].keys():
                        all_wxy[lotId][wxy][test_instance_hlv][read_point] = test_result
                    elif hlv_pass_criteria:  # 這種寫法是賭HV LV各只出現一次 not good
                        if test_result == 'F':
                            all_wxy[lotId][wxy][test_instance_hlv][read_point] = 'F'
                    else:  # 這種寫法是賭HV LV各只出現一次 not good
                        if test_result == 'P':
                            all_wxy[lotId][wxy][test_instance_hlv][read_point] = 'P'
                        elif all_wxy[lotId][wxy][test_instance_hlv][read_point] == 'F':
                            all_wxy[lotId][wxy][test_instance_hlv][read_point] = 'F'
    #  ==============================================================
    print 'Summarizing...'
    # Summarize

    wb = Workbook()
    curr_row = curr_col = 1

    ws_all = wb.active
    ws_all.title = 'All'

    wxy_swap = {}
    all_read_point = sorted(seq[lotId].keys())
    t = ''
    for read_point in all_read_point:
        t += '_T%d' % read_point
    # csv_name = lotId + '%s_ALL.csv' % t
    xlsx_name = lotId + '%s_ALL.csv' % t

    # csv = open(csv_name, 'w')
    l = 'YieldType,Category,ByBlock,Test Instance,'
    for wxy in sorted(all_wxy[lotId].keys()):
        l += ' %s,' % wxy
        wxy_swap[wxy] = 'P'
    write_excel_header(ws_all, l + 'Result', curr_row, curr_col)
    # csv.write(l + 'Result\n')

    curr_row += 1
    curr_col = 1

    l = ',,,Bin,'
    for wxy in sorted(all_wxy[lotId].keys()):
        all_wxy[lotId][wxy]['Bin Swap'] = ''
        result = []
        for read_point in all_read_point:
            if read_point not in all_wxy[lotId][wxy]['Bin'].keys():
                result.append('N')
            else:
                result.append(all_wxy[lotId][wxy]['Bin'][read_point])
        if result.count(result[0]) == len(result):
            all_wxy[lotId][wxy]['Bin Swap'] = result[0]
        else:
            all_wxy[lotId][wxy]['Bin Swap'] = '%s' % '->'.join(result)
        l += '%s,' % '->'.join(result)
    # csv.write(l + '\n')
    write_excel_row(ws_all, l, curr_row, curr_col)

    curr_row += 1
    curr_col = 1

    for test_instance in all_ti:
        l = '%s,%s,%s,%s,' % (ti_yield[test_instance], ti_logic_mbist[test_instance], ti_block[test_instance], test_instance)

        failCount_Ti = 0  # T.I的FC
        for wxy in sorted(all_wxy[lotId].keys()):

            if test_instance not in all_wxy[lotId][wxy].keys():
                l += 'N,'
                continue
            result = []
            for read_point in all_read_point:
                if read_point not in all_wxy[lotId][wxy][test_instance].keys():
                    result.append('N')
                else:
                    result.append(all_wxy[lotId][wxy][test_instance][read_point])
            if result.count('P') == len(result) or result.count('F') == len(result):
                l += result[0] + ','  # 不要印還比較好看
            else:
                if result.count('F') > 0:
                    l += '%s,' % '_'.join(result)
                    failCount_Ti += 1
                else:
                    l += ','  # 不要印還比較好看
                if ti_yield[test_instance] == 'DLY':
                    wxy_swap[wxy] = 'F'
        # csv.write(l + '%d\n' % failCount_Ti)
        write_excel_row(ws_all, l + '%d' % failCount_Ti, curr_row, curr_col)
        curr_row += 1
        curr_col = 1

    l = ',,,Bin,'
    for wxy in sorted(all_wxy[lotId].keys()):
        result = []
        for read_point in all_read_point:
            if read_point not in all_wxy[lotId][wxy]['Bin'].keys():
                result.append('N')
            else:
                result.append(all_wxy[lotId][wxy]['Bin'][read_point])
        l += '%s,' % '->'.join(result)
    write_excel_row(ws_all, l, curr_row, curr_col)
    # csv.write(l + '\n')

    curr_row += 1
    curr_col = 1

    l = ',,,Summary,'
    for wxy in sorted(all_wxy[lotId].keys()):
        l += '%s,' % wxy_swap[wxy]
    write_excel_row(ws_all, l, curr_row, curr_col)
    # csv.write(l + '\n')

    curr_row += 1
    curr_col = 1

    l = ',,,Read Point,'
    for wxy in sorted(all_wxy[lotId].keys()):
        rp = sorted(all_wxy[lotId][wxy]['Bin'].keys())
        l += '%s,' % '_'.join(map(str, rp))
    write_excel_row(ws_all, l, curr_row, curr_col)
    # csv.write(l + '\n')
    # csv.close()

    # if os.path.exists(csv_name):
    #    os.startfile(csv_name)
    #  ==============================================================

    ws_hlv = wb.create_sheet(title='HLV')
    curr_row = curr_col = 1

    summary[lotId] = {}

    wxy_swap = {}
    wxy_t0_good = {}
    all_read_point = sorted(seq[lotId].keys())
    t = ''
    for read_point in all_read_point:
        t += '_T%d' % read_point
    # csv_name = lotId + '%s_HLV.csv' % t
    # csv = open(csv_name, 'w')

    l = 'YieldType,Category,ByBlock,Test Instance,'
    for wxy in sorted(all_wxy[lotId].keys()):
        l += ' %s,' % wxy
        wxy_swap[wxy] = 'P'
        wxy_t0_good[wxy] = 'Y'
        summary[lotId][wxy] = {}
        for gp in ['T0', 'SWAP']:
            summary[lotId][wxy][gp] = {}
            for block in blocks:
                summary[lotId][wxy][gp][block] = 0  # Lot-WID-X-Y
                for core in range(0, max_cpu_core):
                    if re.search(r'SOC|M9', block) and core > 3:
                        continue
                    summary[lotId][wxy][gp]['%s_B%d' % (block, core)] = 0

        summary[lotId][wxy]['T0']['BIN1'] = 1
        summary[lotId][wxy]['SWAP']['SWAPPED'] = 0
    write_excel_header(ws_hlv, l + 'Result', curr_row, curr_col)
    # csv.write(l + 'Result\n')

    curr_row += 1
    curr_col = 1

    l = ',,,Bin,'
    for wxy in sorted(all_wxy[lotId].keys()):
        result = []
        for read_point in all_read_point:
            if read_point not in all_wxy[lotId][wxy]['Bin'].keys():
                result.append('N')
            else:
                result.append(all_wxy[lotId][wxy]['Bin'][read_point])
        l += '%s,' % '->'.join(result)
    write_excel_row(ws_hlv, l, curr_row, curr_col)
    # csv.write(l + '\n')

    curr_row += 1
    curr_col = 1

    for test_instance in all_ti_hlv:
        l = '%s,%s,%s,%s,' % (ti_yield[test_instance], ti_logic_mbist[test_instance], ti_block[test_instance], test_instance)

        curr_block = ti_block[test_instance]
        curr_block_core = ti_block_core[test_instance]

        failCount_Ti = 0  # T.I的FC
        for wxy in sorted(all_wxy[lotId].keys()):

            if test_instance not in all_wxy[lotId][wxy].keys():
                l += 'N,'
                continue

            if ti_yield[test_instance] == 'DLY':
                if all_read_point[0] in all_wxy[lotId][wxy][test_instance].keys():
                    if all_wxy[lotId][wxy][test_instance][all_read_point[0]] == 'F':
                        wxy_t0_good[wxy] = 'N'
                        summary[lotId][wxy]['T0'][curr_block] = 1
                        summary[lotId][wxy]['T0'][curr_block_core] = 1
                        summary[lotId][wxy]['T0']['BIN1'] = 0

            result = []
            for read_point in all_read_point:
                if read_point not in all_wxy[lotId][wxy][test_instance].keys():
                    result.append('N')
                else:
                    result.append(all_wxy[lotId][wxy][test_instance][read_point])
            if result.count('P') == len(result) or result.count('F') == len(result):
                l += result[0] + ','  # 不要印還比較好看
            elif result[-1] == 'P':
                l += result[-1] + ','  # 當沒發生過
            else:
                if result.count('F') > 0:
                    l += '%s,' % '_'.join(result)
                    failCount_Ti += 1
                else:
                    l += '%s,' % '_'.join(result)
                    # l += ','  # 不要印還比較好看
                if ti_yield[test_instance] == 'DLY':  # 如果遇到N的話 比較麻煩.
                    wxy_swap[wxy] = 'F'
                    if result[-1] == 'F':
                        summary[lotId][wxy]['SWAP'][curr_block] = 1
                        summary[lotId][wxy]['SWAP'][curr_block_core] = 1
                        summary[lotId][wxy]['SWAP']['SWAPPED'] = 1
        write_excel_row(ws_hlv, l + '%d' % failCount_Ti, curr_row, curr_col)
        curr_row += 1
        curr_col = 1
        # csv.write(l + '%d\n' % failCount_Ti)

    l = ',,,Bin,'
    for wxy in sorted(all_wxy[lotId].keys()):
        result = []
        for read_point in all_read_point:
            if read_point not in all_wxy[lotId][wxy]['Bin'].keys():
                result.append('N')
            else:
                result.append(all_wxy[lotId][wxy]['Bin'][read_point])
        l += '%s,' % '->'.join(result)
    write_excel_row(ws_hlv, l, curr_row, curr_col)
    # csv.write(l + '\n')

    curr_row += 1
    curr_col = 1

    l = ',,,T0 Good,'
    for wxy in sorted(all_wxy[lotId].keys()):
        l += '%s,' % wxy_t0_good[wxy]
    write_excel_row(ws_hlv, l, curr_row, curr_col)
    # csv.write(l + '\n')

    curr_row += 1
    curr_col = 1

    l = ',,,Read Point,'
    for wxy in sorted(all_wxy[lotId].keys()):
        rp = sorted(all_wxy[lotId][wxy]['Bin'].keys())
        l += '%s,' % '_'.join(map(str, rp))
    write_excel_row(ws_hlv, l, curr_row, curr_col)
    # csv.write(l + '\n')
    # csv.close()

    # if os.path.exists(csv_name):
    #    os.startfile(csv_name)
    #  ==============================================================

    ws_summary = wb.create_sheet(title='Summary')
    ws_summary.sheet_properties.tabColor = '0070C0'
    row = col = 1

    all_read_point = sorted(seq[lotId].keys())
    t = ''
    for read_point in all_read_point:
        t += '_T%d' % read_point
    # csv_name = lotId + '%s_Summary.csv' % t
    # csv = open(csv_name, 'w')
    write_excel_header(ws_summary, ',,,,,T0 Passed,,,,,,,Swapped,,,,,,,Core', curr_row, curr_col)
    # csv.write(',,,,,T0 Passed,,,,,,,Swapped,,,,,,,Core\n')
    curr_row += 1
    curr_col = 1

    l = 'LotID,Wafer,X,Y,Bin Swap,Bin1,'
    for block in blocks:
        l += '%s,' % block
    l += 'Swap,'
    for block in blocks:
        l += '%s,' % block
    for block in blocks:
        for core in range(0, max_cpu_core):
            if re.search(r'SOC|M9', block) and core > 3:
                continue
            l += '%s_B%d,' % (block, core)
    write_excel_header(ws_summary, l, curr_row, curr_col)
    # csv.write(l + '\n')
    curr_row += 1
    curr_col = 1

    for wxy in sorted(all_wxy[lotId].keys()):
        tt = wxy.split(':')
        l = '%s,%s,%s,%s,' % (lotId, tt[0], tt[1], tt[2])
        l += '%s,' % all_wxy[lotId][wxy]['Bin Swap']
        l += '%d,' % summary[lotId][wxy]['T0']['BIN1']
        for block in blocks:
            l += '%d,' % summary[lotId][wxy]['T0'][block]  # Lot-WID-X-Y
        l += '%d,' % summary[lotId][wxy]['SWAP']['SWAPPED']
        for block in blocks:
            l += '%d,' % summary[lotId][wxy]['SWAP'][block]  # Lot-WID-X-Y
        for block in blocks:
            for core in range(0, max_cpu_core):
                if re.search(r'SOC|M9', block) and core > 3:
                    continue
                l += '%d,' % summary[lotId][wxy]['SWAP']['%s_B%d' % (block, core)]

        # csv.write(l + '\n')
        write_excel_row(ws_summary, l, curr_row, curr_col)
        curr_row += 1
        curr_col = 1

    # csv.close()

    # if os.path.exists(csv_name):
    #    os.startfile(csv_name)

    wb.save(filename=xlsx_name)
    if os.path.exists(xlsx_name):
        os.startfile(xlsx_name)
