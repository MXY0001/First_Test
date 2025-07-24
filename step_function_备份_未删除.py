import step_split
import numpy as np


def End_V(sht, num):
    col_record = sht.used_range.last_cell.column
    row_record = sht.used_range.last_cell.row
    title_record = sht[0:1, 0:col_record].value
    V_record_index = title_record.index('电压(V)')
    V_value = sht[0:row_record, V_record_index:V_record_index+1].value
    if '工步序号' in title_record:
        step_record_index = title_record.index('工步序号')
    elif '工步号' in title_record:
        step_record_index = title_record.index('工步号')
    elif '步次' in title_record:
        step_record_index = title_record.index('步次')
    elif '状态' in title_record:
        step_record_index = title_record.index('状态')
    else:
        return 0
    step_value = sht[0:row_record, step_record_index:step_record_index+1].value
    V_value_np = np.array(V_value)
    step_value_np = np.array(step_value)
    recode_1 = np.where(step_value_np == num)[0]
    index_list = list(recode_1)
    value_index_2 = V_value_np[index_list]

    return float(value_index_2[-1])


def End_I(sht, num):
    col_record = sht.used_range.last_cell.column
    row_record = sht.used_range.last_cell.row
    title_record = sht[0:1, 0:col_record].value
    V_record_index = title_record.index('电流(A)')
    V_value = sht[0:row_record, V_record_index:V_record_index+1].value
    if '工步序号' in title_record:
        step_record_index = title_record.index('工步序号')
    elif '工步号' in title_record:
        step_record_index = title_record.index('工步号')
    elif '步次' in title_record:
        step_record_index = title_record.index('步次')
    elif '状态' in title_record:
        step_record_index = title_record.index('状态')
    else:
        return 0
    step_value = sht[0:row_record, step_record_index:step_record_index+1].value
    V_value_np = np.array(V_value)
    step_value_np = np.array(step_value)
    recode_1 = np.where(step_value_np == num)[0]
    index_list = list(recode_1)
    value_index_2 = V_value_np[index_list]

    return float(value_index_2[-1])


def index_time_step(title_record):
    if '工步序号' in title_record:
        step_record_index = title_record.index('工步序号')
    elif '工步号' in title_record:
        step_record_index = title_record.index('工步号')
    elif '步次' in title_record:
        step_record_index = title_record.index('步次')
    elif '状态' in title_record:
        step_record_index = title_record.index('状态')
    else:
        return 0, 0

    if '相对时间(d.h:min:s.ms)' in title_record:
        times_record_index = title_record.index('相对时间(d.h:min:s.ms)')
    elif '相对时间(d.hh:mm:ss.sss)' in title_record:
        times_record_index = title_record.index('相对时间(d.hh:mm:ss.sss)')
    elif '记录时间(HH:mm:ss.sss)' in title_record:
        times_record_index = title_record.index('记录时间(HH:mm:ss.sss)')
    elif '记录时间(d.hh:mm:ss.sss)' in title_record:
        times_record_index = title_record.index('记录时间(d.hh:mm:ss.sss)')
    elif '相对时间(h:min:s.ms)' in title_record:
        times_record_index = title_record.index('相对时间(h:min:s.ms)')
    else:
        return 0, 0
    return step_record_index, times_record_index


def step_1(wb, w): # 循环3圈
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[12:13, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[14:15, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[12:13, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[14:15, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[12:13, energy_index:energy_index+1].value
        discharge_energy = sht_step[14:15, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '12.0')[0]
    recode_2 = np.where(step_value_np == '14.0')[0]
    recode_3 = np.where(step_value_np == '13.0')[0]
    recode_4 = np.where(step_value_np == '15.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[12:13, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '12.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[14:15, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '14.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]

    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[19:20, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[21:22, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[20:21, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[20:21, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[22:23, discharge_V:discharge_V + 1].value
        charge_end_I_10 = float(sht_step[22:23, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '19.0')
        standing_end2_V = End_V(sht_record, '21.0')
        discharge_end_V_10 = End_V(sht_record, '20.0')
        discharge_end_I_10 = End_I(sht_record, '20.0')
        charge_end_V_10 = End_V(sht_record, '22.0')
        charge_end_I_10 = End_I(sht_record, '22.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[12:13, stress_index:stress_index+1].value, sht_step[13:14, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[14:15, stress_index:stress_index+1].value, sht_step[15:16, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[12:13, stress_index:stress_index+1].value, sht_step[13:14, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[14:15, stress_index:stress_index+1].value, sht_step[15:16, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[12:13, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[14:15, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           standing_end2_V, charge_end_V_10, abs(charge_end_I_10), charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_2(wb, w): # 循环2圈
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]

    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value
    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []

    # 容量 能量
    charge_cap = abs(sht_step[8:9, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[10:11, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[8:9, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[10:11, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[8:9, energy_index:energy_index+1].value
        discharge_energy = sht_step[10:11, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '8.0')[0]
    recode_2 = np.where(step_value_np == '10.0')[0]
    recode_3 = np.where(step_value_np == '9.0')[0]
    recode_4 = np.where(step_value_np == '11.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1)-1 and float(I_value_record_1[x]) - float(I_value_record_1[x+1]) >= 1]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V!= []:
        charge_end_V = sht_step[8:9, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '8.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[10:11, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '10.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        try:
            recode_4 = np.where(step_value_np == '16.0')[0]
            index_list_3 = list(recode_4)
            value_index_3 = V_value_np[index_list_3]
            discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
        except IndexError:
            discharge_standing_end_V = value_index_3[-1]


    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[19:20, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[21:22, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[20:21, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[20:21, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[22:23, discharge_V :discharge_V + 1].value
        charge_end_I_10 = float(sht_step[22:23, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '19.0')
        standing_end2_V = End_V(sht_record, '21.0')
        discharge_end_V_10 = End_V(sht_record, '20.0')
        discharge_end_I_10 = End_I(sht_record, '20.0')
        charge_end_V_10 = End_V(sht_record, '22.0')
        charge_end_I_10 = End_I(sht_record, '22.0')
    # 膨胀力


    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[8:9, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[10:11, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'


    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                    discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
               standing_end2_V, charge_end_V_10, charge_end_I_10, charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
               V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1



def step_3_1(wb, w):
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[8:9, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[10:11, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[8:9, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[10:11, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[8:9, energy_index:energy_index+1].value
        discharge_energy = sht_step[10:11, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '8.0')[0]
    recode_2 = np.where(step_value_np == '10.0')[0]
    recode_3 = np.where(step_value_np == '9.0')[0]
    recode_4 = np.where(step_value_np == '11.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[8:9, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '8.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[10:11, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '10.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]


    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[16:17, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[17:18, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[17:18, discharge_I:discharge_I + 1].value)

    else:
        standing_end1_V = End_V(sht_record, '16.0')
        discharge_end_V_10 = End_V(sht_record, '17.0')
        discharge_end_I_10 = End_I(sht_record, '17.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[8:9, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[10:11, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = None
            discharge_max_tem = None

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           None, [None], [None], charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_5_1(wb, w): # 循环3圈
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []

    # 容量 能量
    charge_cap = abs(sht_step[12:13, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[14:15, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[12:13, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[14:15, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[12:13, energy_index:energy_index+1].value
        discharge_energy = sht_step[14:15, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '12.0')[0]
    recode_2 = np.where(step_value_np == '14.0')[0]
    recode_3 = np.where(step_value_np == '13.0')[0]
    recode_4 = np.where(step_value_np == '15.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[12:13, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '12.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[14:15, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '14.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]

    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[18:19, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[20:21, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[19:20, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[19:20, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[21:22, discharge_V:discharge_V + 1].value
        charge_end_I_10 = float(sht_step[21:22, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '18.0')
        standing_end2_V = End_V(sht_record, '20.0')
        discharge_end_V_10 = End_V(sht_record, '19.0')
        discharge_end_I_10 = End_I(sht_record, '19.0')
        charge_end_V_10 = End_V(sht_record, '21.0')
        charge_end_I_10 = End_I(sht_record, '21.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[12:13, stress_index:stress_index+1].value, sht_step[13:14, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[14:15, stress_index:stress_index+1].value, sht_step[15:16, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[12:13, stress_index:stress_index+1].value, sht_step[13:14, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[14:15, stress_index:stress_index+1].value, sht_step[15:16, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[12:13, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[14:15, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           standing_end2_V, charge_end_V_10, abs(charge_end_I_10), charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_6(wb, w):
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[4:5, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[6:7, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[4:5, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[6:7, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[4:5, energy_index:energy_index+1].value
        discharge_energy = sht_step[6:7, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '4.0')[0]
    recode_2 = np.where(step_value_np == '6.0')[0]
    recode_3 = np.where(step_value_np == '5.0')[0]
    recode_4 = np.where(step_value_np == '7.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[4:5, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '4.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[6:7, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '6.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]

    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[9:10, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[11:12, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[10:11, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[10:11, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[12:13, discharge_V:discharge_V + 1].value
        charge_end_I_10 = float(sht_step[12:13, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '9.0')
        standing_end2_V = End_V(sht_record, '11.0')
        discharge_end_V_10 = End_V(sht_record, '10.0')
        discharge_end_I_10 = End_I(sht_record, '10.0')
        charge_end_V_10 = End_V(sht_record, '12.0')
        charge_end_I_10 = End_I(sht_record, '12.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[4:5, stress_index:stress_index+1].value, sht_step[5:6, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[6:7, stress_index:stress_index+1].value, sht_step[7:8, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[4:5, stress_index:stress_index+1].value, sht_step[5:6, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[6:7, stress_index:stress_index+1].value, sht_step[7:8, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[4:5, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[6:7, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           standing_end2_V, charge_end_V_10, abs(charge_end_I_10), charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_8(wb, w): # 循环3圈
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[12:13, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[14:15, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[12:13, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[14:15, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[12:13, energy_index:energy_index+1].value
        discharge_energy = sht_step[14:15, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '12.0')[0]
    recode_2 = np.where(step_value_np == '14.0')[0]
    recode_3 = np.where(step_value_np == '13.0')[0]
    recode_4 = np.where(step_value_np == '15.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[12:13, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '12.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[14:15, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '14.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]

    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[20:21, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[22:23, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[21:22, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[21:22, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[23:24, discharge_V:discharge_V + 1].value
        charge_end_I_10 = float(sht_step[23:24, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '20.0')
        standing_end2_V = End_V(sht_record, '22.0')
        discharge_end_V_10 = End_V(sht_record, '21.0')
        discharge_end_I_10 = End_I(sht_record, '21.0')
        charge_end_V_10 = End_V(sht_record, '23.0')
        charge_end_I_10 = End_I(sht_record, '23.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[12:13, stress_index:stress_index+1].value, sht_step[13:14, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[14:15, stress_index:stress_index+1].value, sht_step[15:16, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[12:13, stress_index:stress_index+1].value, sht_step[13:14, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[14:15, stress_index:stress_index+1].value, sht_step[15:16, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[12:13, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[14:15, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           standing_end2_V, charge_end_V_10, abs(charge_end_I_10), charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_9_1(wb, w):
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[13:14, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[15:16, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[13:14, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[15:16, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[13:14, energy_index:energy_index+1].value
        discharge_energy = sht_step[15:16, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '13.0')[0]
    recode_2 = np.where(step_value_np == '15.0')[0]
    recode_3 = np.where(step_value_np == '14.0')[0]
    recode_4 = np.where(step_value_np == '16.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[13:14, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '13.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[15:16, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '15.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]


    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[21:22, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[23:24, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[22:23, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[22:23, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[24:25, discharge_V:discharge_V + 1].value
        charge_end_I_10 = float(sht_step[24:25, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '21.0')
        standing_end2_V = End_V(sht_record, '23.0')
        discharge_end_V_10 = End_V(sht_record, '22.0')
        discharge_end_I_10 = End_I(sht_record, '22.0')
        charge_end_V_10 = End_V(sht_record, '24.0')
        charge_end_I_10 = End_I(sht_record, '24.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[13:14, stress_index:stress_index+1].value, sht_step[14:15, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[15:16, stress_index:stress_index+1].value, sht_step[16:17, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[13:14, stress_index:stress_index+1].value, sht_step[14:15, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[15:16, stress_index:stress_index+1].value, sht_step[16:17, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[13:14, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[15:16, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = None
            discharge_max_tem = None

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           standing_end2_V, charge_end_V_10, charge_end_I_10, charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_10_1(wb, w):
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[12:13, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[14:15, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[12:13, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[14:15, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[12:13, energy_index:energy_index+1].value
        discharge_energy = sht_step[14:15, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '12.0')[0]
    recode_2 = np.where(step_value_np == '14.0')[0]
    recode_3 = np.where(step_value_np == '13.0')[0]
    recode_4 = np.where(step_value_np == '15.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[12:13, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '12.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[14:15, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '14.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]


    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[19:20, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[20:21, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[20:21, discharge_I:discharge_I + 1].value)

    else:
        standing_end1_V = End_V(sht_record, '19.0')
        discharge_end_V_10 = End_V(sht_record, '20.0')
        discharge_end_I_10 = End_I(sht_record, '20.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[12:13, stress_index:stress_index+1].value, sht_step[13:14, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[14:15, stress_index:stress_index+1].value, sht_step[15:16, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[12:13, stress_index:stress_index+1].value, sht_step[13:14, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[14:15, stress_index:stress_index+1].value, sht_step[15:16, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[12:13, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[14:15, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = None
            discharge_max_tem = None

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           None, None, None, charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_12(wb, w):
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[13:14, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[15:16, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[13:14, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[15:16, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[13:14, energy_index:energy_index+1].value
        discharge_energy = sht_step[15:16, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '13.0')[0]
    recode_2 = np.where(step_value_np == '15.0')[0]
    recode_3 = np.where(step_value_np == '14.0')[0]
    recode_4 = np.where(step_value_np == '16.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[13:14, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '13.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[15:16, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '15.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]


    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[20:21, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[22:23, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[21:22, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[21:22, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[23:24, discharge_V:discharge_V + 1].value
        charge_end_I_10 = float(sht_step[23:24, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '20.0')
        standing_end2_V = End_V(sht_record, '22.0')
        discharge_end_V_10 = End_V(sht_record, '21.0')
        discharge_end_I_10 = End_I(sht_record, '21.0')
        charge_end_V_10 = End_V(sht_record, '23.0')
        charge_end_I_10 = End_I(sht_record, '23.0')



    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[13:14, stress_index:stress_index+1].value, sht_step[14:15, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[15:16, stress_index:stress_index+1].value, sht_step[16:17, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[13:14, stress_index:stress_index+1].value, sht_step[14:15, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[15:16, stress_index:stress_index+1].value, sht_step[16:17, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[13:14, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[15:16, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = None
            discharge_max_tem = None

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           None, None, None, charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_13(wb, w): #
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[8:9, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[10:11, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[8:9, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[10:11, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[8:9, energy_index:energy_index+1].value
        discharge_energy = sht_step[10:11, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '8.0')[0]
    recode_2 = np.where(step_value_np == '10.0')[0]
    recode_3 = np.where(step_value_np == '9.0')[0]
    recode_4 = np.where(step_value_np == '11.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[8:9, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '8.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[10:11, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '10.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]

    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[16:17, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[18:19, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[17:18, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[17:18, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[19:20, discharge_V:discharge_V + 1].value
        charge_end_I_10 = float(sht_step[19:20, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '16.0')
        standing_end2_V = End_V(sht_record, '18.0')
        discharge_end_V_10 = End_V(sht_record, '17.0')
        discharge_end_I_10 = End_I(sht_record, '17.0')
        charge_end_V_10 = End_V(sht_record, '19.0')
        charge_end_I_10 = End_I(sht_record, '19.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[8:9, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[10:11, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           standing_end2_V, charge_end_V_10, abs(charge_end_I_10), charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_15(wb, w): #
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[8:9, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[10:11, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[8:9, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[10:11, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[8:9, energy_index:energy_index+1].value
        discharge_energy = sht_step[10:11, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '8.0')[0]
    recode_2 = np.where(step_value_np == '10.0')[0]
    recode_3 = np.where(step_value_np == '9.0')[0]
    recode_4 = np.where(step_value_np == '11.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[8:9, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '8.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[10:11, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '10.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]

    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[15:16, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[17:18, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[16:17, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[16:17, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[18:19, discharge_V:discharge_V + 1].value
        charge_end_I_10 = float(sht_step[18:19, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '15.0')
        standing_end2_V = End_V(sht_record, '17.0')
        discharge_end_V_10 = End_V(sht_record, '16.0')
        discharge_end_I_10 = End_I(sht_record, '16.0')
        charge_end_V_10 = End_V(sht_record, '18.0')
        charge_end_I_10 = End_I(sht_record, '18.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[8:9, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[10:11, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           standing_end2_V, charge_end_V_10, abs(charge_end_I_10), charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_16(wb, w): #
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[6:7, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[8:9, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[6:7, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[8:9, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[6:7, energy_index:energy_index+1].value
        discharge_energy = sht_step[8:9, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '6.0')[0]
    recode_2 = np.where(step_value_np == '8.0')[0]
    recode_3 = np.where(step_value_np == '7.0')[0]
    recode_4 = np.where(step_value_np == '9.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[6:7, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '6.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[8:9, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '8.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]

    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[9:10, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[10:11, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[10:11, discharge_I:discharge_I + 1].value)

    else:
        standing_end1_V = End_V(sht_record, '9.0')
        discharge_end_V_10 = End_V(sht_record, '10.0')
        discharge_end_I_10 = End_I(sht_record, '10.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[6:7, stress_index:stress_index+1].value, sht_step[7:8, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:11, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[6:7, stress_index:stress_index+1].value, sht_step[7:8, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:11, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[6:7, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[8:9, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           None, None, None, charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_18(wb, w): # 循环3圈
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    title_step = sht_step[0:1, 0:col_step].value
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[5:6, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[6:7, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[6:7, discharge_I:discharge_I + 1].value)

    else:
        standing_end1_V = End_V(sht_record, '5.0')
        discharge_end_V_10 = End_V(sht_record, '6.0')
        discharge_end_I_10 = End_I(sht_record, '6.0')


    return None, None, None, None, None, None, None, None, None, None, None, standing_end1_V, discharge_end_V_10, abs(
        discharge_end_I_10), None, None, None, None, None, None, None, None, None, None, None


def step_19(wb, w): #
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[8:9, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[10:11, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[8:9, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[10:11, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[8:9, energy_index:energy_index+1].value
        discharge_energy = sht_step[10:11, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '8.0')[0]
    recode_2 = np.where(step_value_np == '10.0')[0]
    recode_3 = np.where(step_value_np == '9.0')[0]
    recode_4 = np.where(step_value_np == '11.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)
    if index_list_3 == []:
        recode_4 = np.where(step_value_np == '16.0')[0]
        index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[8:9, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '8.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[10:11, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '10.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]

    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[15:16, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[16:17, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[16:17, discharge_I:discharge_I + 1].value)

    else:
        standing_end1_V = End_V(sht_record, '15.0')
        discharge_end_V_10 = End_V(sht_record, '16.0')
        discharge_end_I_10 = End_I(sht_record, '16.0')

    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[8:9, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[10:11, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           None, None, None, charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_20(wb, w): #
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[6:7, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[8:9, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[6:7, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[8:9, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[6:7, energy_index:energy_index+1].value
        discharge_energy = sht_step[8:9, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '6.0')[0]
    recode_2 = np.where(step_value_np == '8.0')[0]
    recode_3 = np.where(step_value_np == '7.0')[0]
    recode_4 = np.where(step_value_np == '9.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[6:7, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '6.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[8:9, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '8.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]

    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[13:14, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[14:15, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[14:15, discharge_I:discharge_I + 1].value)

    else:
        standing_end1_V = End_V(sht_record, '13.0')
        discharge_end_V_10 = End_V(sht_record, '14.0')
        discharge_end_I_10 = End_I(sht_record, '14.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[6:7, stress_index:stress_index+1].value, sht_step[7:8, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[6:7, stress_index:stress_index+1].value, sht_step[7:8, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[6:7, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[8:9, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           None, None, None, charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1



def step_21(wb, w):
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    title_step = sht_step[0:1, 0:col_step].value

    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[7:8, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[8:9, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[8:9, discharge_I:discharge_I + 1].value)

    else:
        standing_end1_V = End_V(sht_record, '7.0')
        discharge_end_V_10 = End_V(sht_record, '8.0')
        discharge_end_I_10 = End_I(sht_record, '8.0')


    return None, None, None, None, None, None, None, None, None, None, None, standing_end1_V, discharge_end_V_10, abs(
        discharge_end_I_10), None, None, None, None, None, None, None, None, None, None, None


def step_22(wb, w):
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[6:7, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[8:9, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[6:7, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[8:9, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[6:7, energy_index:energy_index+1].value
        discharge_energy = sht_step[8:9, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '6.0')[0]
    recode_2 = np.where(step_value_np == '8.0')[0]
    recode_3 = np.where(step_value_np == '7.0')[0]
    recode_4 = np.where(step_value_np == '9.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[6:7, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '6.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[8:9, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '8.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[6:7, stress_index:stress_index+1].value, sht_step[7:8, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[6:7, stress_index:stress_index+1].value, sht_step[7:8, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[6:7, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[8:9, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = None
            discharge_max_tem = None

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, None, None, None, None, None, None, charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_23(wb, w): # 循环3圈
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[8:9, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[10:11, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[8:9, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[10:11, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[8:9, energy_index:energy_index+1].value
        discharge_energy = sht_step[10:11, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '8.0')[0]
    recode_2 = np.where(step_value_np == '10.0')[0]
    recode_3 = np.where(step_value_np == '9.0')[0]
    recode_4 = np.where(step_value_np == '11.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[8:9, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '8.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[10:11, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '10.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[8:9, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[10:11, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = None
            discharge_max_tem = None

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, None, None, None, \
           None, None, None, charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_25(wb, w):
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[4:5, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[6:7, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[4:5, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[6:7, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[4:5, energy_index:energy_index+1].value
        discharge_energy = sht_step[6:7, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '4.0')[0]
    recode_2 = np.where(step_value_np == '6.0')[0]
    recode_3 = np.where(step_value_np == '5.0')[0]
    recode_4 = np.where(step_value_np == '7.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[4:5, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '4.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[6:7, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '6.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[4:5, stress_index:stress_index+1].value, sht_step[5:6, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[6:7, stress_index:stress_index+1].value, sht_step[7:8, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[4:5, stress_index:stress_index+1].value, sht_step[5:6, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[6:7, stress_index:stress_index+1].value, sht_step[7:8, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[4:5, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[6:7, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = None
            discharge_max_tem = None

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, None, None, None, \
           None, None, None, charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_26(wb, w): # 循环3圈
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[12:13, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[14:15, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[12:13, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[14:15, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[12:13, energy_index:energy_index+1].value
        discharge_energy = sht_step[14:15, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '12.0')[0]
    recode_2 = np.where(step_value_np == '14.0')[0]
    recode_3 = np.where(step_value_np == '13.0')[0]
    recode_4 = np.where(step_value_np == '15.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[12:13, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '12.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[14:15, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '14.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]

    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[19:20, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[21:22, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[20:21, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[20:21, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[22:23,discharge_V:discharge_V + 1].value
        charge_end_I_10 = float(sht_step[22:23, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '19.0')
        standing_end2_V = End_V(sht_record, '21.0')
        discharge_end_V_10 = End_V(sht_record, '20.0')
        discharge_end_I_10 = End_I(sht_record, '20.0')
        charge_end_V_10 = End_V(sht_record, '22.0')
        charge_end_I_10 = End_I(sht_record, '22.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[12:13, stress_index:stress_index+1].value, sht_step[13:14, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[14:15, stress_index:stress_index+1].value, sht_step[15:16, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[12:13, stress_index:stress_index+1].value, sht_step[13:14, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[14:15, stress_index:stress_index+1].value, sht_step[15:16, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[12:13, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[14:15, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           standing_end2_V, charge_end_V_10, abs(charge_end_I_10), charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_27(wb, w): # 循环3圈
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[8:9, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[10:11, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[8:9, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[10:11, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[8:9, energy_index:energy_index+1].value
        discharge_energy = sht_step[10:11, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '8.0')[0]
    recode_2 = np.where(step_value_np == '10.0')[0]
    recode_3 = np.where(step_value_np == '9.0')[0]
    recode_4 = np.where(step_value_np == '11.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[8:9, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '8.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[10:11, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '10.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]

    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[16:17, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[18:19, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[17:18, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[17:18, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[19:20, discharge_V:discharge_V + 1].value
        charge_end_I_10 = float(sht_step[19:20, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '16.0')
        standing_end2_V = End_V(sht_record, '18.0')
        discharge_end_V_10 = End_V(sht_record, '17.0')
        discharge_end_I_10 = End_I(sht_record, '17.0')
        charge_end_V_10 = End_V(sht_record, '19.0')
        charge_end_I_10 = End_I(sht_record, '19.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[8:9, stress_index:stress_index+1].value, sht_step[9:10, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[10:11, stress_index:stress_index+1].value, sht_step[11:12, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[8:9, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[10:11, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           standing_end2_V, charge_end_V_10, abs(charge_end_I_10), charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_28(wb, w):
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[19:20, cap_index:cap_index+1].value)
    discharge_cap = abs(sht_step[21:22, cap_index:cap_index+1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[19:20, energy_index:energy_index+1].value)
        discharge_energy = abs(sht_step[21:22, energy_index:energy_index+1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[19:20, energy_index:energy_index+1].value
        discharge_energy = sht_step[21:22, disenergy_index:disenergy_index+1].value

    # CC+CV 时间 容量占比

    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index+1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index+1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index+1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index+1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index+1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '19.0')[0]
    recode_2 = np.where(step_value_np == '21.0')[0]
    recode_3 = np.where(step_value_np == '20.0')[0]
    recode_4 = np.where(step_value_np == '22.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0]+1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[19:20, discharge_V:discharge_V+1].value
    else:
        charge_end_V = End_V(sht_record, '19.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[21:22, discharge_V:discharge_V+1].value
    else:
        discharge_end_V = End_V(sht_record, '21.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]

    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[26:27, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[28:29, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[27:28, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[27:28, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[29:30, discharge_V:discharge_V + 1].value
        charge_end_I_10 = float(sht_step[29:30, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '26.0')
        standing_end2_V = End_V(sht_record, '28.0')
        discharge_end_V_10 = End_V(sht_record, '27.0')
        discharge_end_I_10 = End_I(sht_record, '27.0')
        charge_end_V_10 = End_V(sht_record, '29.0')
        charge_end_I_10 = End_I(sht_record, '29.0')


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[19:20, stress_index:stress_index+1].value, sht_step[20:21, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[21:22, stress_index:stress_index+1].value, sht_step[22:23, stress_index:stress_index+1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[19:20, stress_index:stress_index+1].value, sht_step[20:21, stress_index:stress_index+1].value)
        discharge_max_stress = max(sht_step[21:22, stress_index:stress_index+1].value, sht_step[22:23, stress_index:stress_index+1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[19:20, tem_max:tem_max+1].value
        discharge_max_tem = sht_step[21:22, tem_max:tem_max+1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max+1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, standing_end1_V, discharge_end_V_10, abs(discharge_end_I_10), \
           standing_end2_V, charge_end_V_10, abs(charge_end_I_10), charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1


def step_29(wb, w):
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    title_step = sht_step[0:1, 0:col_step].value

    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # DCR
    if discharge_V != []:
        standing_end1_V = sht_step[5:6, discharge_V:discharge_V+1].value
        standing_end2_V = sht_step[7:8, discharge_V:discharge_V+1].value
        discharge_end_V_10 = sht_step[6:7, discharge_V:discharge_V + 1].value
        discharge_end_I_10 = float(sht_step[6:7, discharge_I:discharge_I + 1].value)
        charge_end_V_10 = sht_step[8:9, discharge_V:discharge_V + 1].value
        charge_end_I_10 = float(sht_step[8:9, discharge_I:discharge_I + 1].value)
    else:
        standing_end1_V = End_V(sht_record, '5.0')
        standing_end2_V = End_V(sht_record, '7.0')
        discharge_end_V_10 = End_V(sht_record, '6.0')
        discharge_end_I_10 = End_I(sht_record, '6.0')
        charge_end_V_10 = End_V(sht_record, '8.0')
        charge_end_I_10 = End_I(sht_record, '8.0')


    return None, None, None, None, None, None, None, None, None, None, None, standing_end1_V, discharge_end_V_10, abs(
        discharge_end_I_10), standing_end2_V, charge_end_V_10, abs(charge_end_I_10), None, None, None, None, None, None, None, None


def step_30(wb, w):
    try:
        sht_step = wb.sheets['工步信息']
        sht_record = wb.sheets['记录数据']
    except Exception:
        sht_step = wb.sheets[2]
        sht_record = wb.sheets[3]
    col_step = sht_step.used_range.last_cell.column
    col_record = sht_record.used_range.last_cell.column
    row_record = sht_record.used_range.last_cell.row
    title_step = sht_step[0:1, 0:col_step].value
    title_record = sht_record[0:1, 0:col_record].value

    cap_index = title_step.index('容量(Ah)')
    try:
        discharge_V = title_step.index('结束电压(V)')
    except ValueError:
        try:
            discharge_V = title_step.index('终止电压(V)')
        except ValueError:
            discharge_V = []
    try:
        discharge_I = title_step.index('结束电流(A)')
    except ValueError:
        try:
            discharge_I = title_step.index('终止电流(A)')
        except ValueError:
            discharge_I = []
    # 容量 能量
    charge_cap = abs(sht_step[12:13, cap_index:cap_index + 1].value)
    discharge_cap = abs(sht_step[14:15, cap_index:cap_index + 1].value)
    try:
        energy_index = title_step.index('净放电能量(Wh)')
        charge_energy = abs(sht_step[12:13, energy_index:energy_index + 1].value)
        discharge_energy = abs(sht_step[14:15, energy_index:energy_index + 1].value)
    except ValueError:
        energy_index = title_step.index('充电能量(Wh)')
        disenergy_index = title_step.index('放电能量(Wh)')
        charge_energy = sht_step[12:13, energy_index:energy_index + 1].value
        discharge_energy = sht_step[14:15, disenergy_index:disenergy_index + 1].value

    # CC+CV 时间 容量占比
    step_record_index, times_record_index = index_time_step(title_record)
    if step_record_index == 0:
        return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

    V_record_index = title_record.index('电压(V)')
    I_record_index = title_record.index('电流(A)')
    cap_record_index = title_record.index('容量(Ah)')

    step_value = sht_record[0:row_record, step_record_index:step_record_index + 1].value
    times_value = sht_record[0:row_record, times_record_index:times_record_index + 1].value
    V_value = sht_record[0:row_record, V_record_index:V_record_index + 1].value
    I_value = sht_record[0:row_record, I_record_index:I_record_index + 1].value
    cap_value = sht_record[0:row_record, cap_record_index:cap_record_index + 1].value

    step_value_np = np.array(step_value)
    times_value_np = np.array(times_value)
    V_value_np = np.array(V_value)
    I_value_np = np.array(I_value)
    cap_value_np = np.array(cap_value)
    recode_1 = np.where(step_value_np == '12.0')[0]
    recode_2 = np.where(step_value_np == '14.0')[0]
    recode_3 = np.where(step_value_np == '13.0')[0]
    recode_4 = np.where(step_value_np == '15.0')[0]
    index_list = list(recode_1)
    index_list_1 = list(recode_2)
    index_list_2 = list(recode_3)
    index_list_3 = list(recode_4)

    charge_times_1 = times_value_np[index_list]
    charge_times_2 = times_value_np[index_list_1]
    charge_times_3 = times_value_np[index_list_2]
    charge_times_4 = times_value_np[index_list_3]
    V_value_record_1 = V_value_np[index_list]
    I_value_record_1 = I_value_np[index_list]
    V_value_record_2 = V_value_np[index_list_1]
    I_value_record_2 = I_value_np[index_list_1]
    charge_cap_value = cap_value_np[index_list]
    charge_cap_value_1 = cap_value_np[index_list_1]
    Constant_voltage_index = [x for x in range(len(I_value_record_1)) if x < len(I_value_record_1) - 1 and float(I_value_record_1[x]) - float(I_value_record_1[x + 1]) >= 0.5]
    CC_time = charge_times_1[Constant_voltage_index[0]]
    total_time = charge_times_1[-1]
    CC_cap = charge_cap_value[Constant_voltage_index[0] + 1]

    # 静置30min后电压回弹 末端电压
    if discharge_V != []:
        charge_end_V = sht_step[12:13, discharge_V:discharge_V + 1].value
    else:
        charge_end_V = End_V(sht_record, '12.0')
    value_index_2 = V_value_np[index_list_2]
    try:
        charge_standing_end_V = [value_index_2[x] for x in range(len(index_list_2)) if float(str(charge_times_3[x]).split(':')[0]) == 0 and float(str(charge_times_3[x]).split(':')[1]) == 30][0]
    except IndexError:
        charge_standing_end_V = value_index_2[-1]
    if discharge_V != []:
        discharge_end_V = sht_step[14:15, discharge_V:discharge_V + 1].value
    else:
        discharge_end_V = End_V(sht_record, '14.0')
    value_index_3 = V_value_np[index_list_3]
    try:
        discharge_standing_end_V = [value_index_3[x] for x in range(len(index_list_3)) if float(str(charge_times_4[x]).split(':')[0]) == 0 and float(str(charge_times_4[x]).split(':')[1]) == 30][0]
    except IndexError:
        discharge_standing_end_V = value_index_3[-1]


    # 膨胀力
    if '最大压力(Kg)' in title_step:
        stress_index = title_step.index('最大压力(Kg)')
        charge_max_stress = max(sht_step[12:13, stress_index:stress_index + 1].value, sht_step[13:14, stress_index:stress_index + 1].value)
        discharge_max_stress = max(sht_step[14:15, stress_index:stress_index + 1].value, sht_step[15:16, stress_index:stress_index + 1].value)
    elif '最大压力' in title_step:
        stress_index = title_step.index('最大压力')
        charge_max_stress = max(sht_step[12:13, stress_index:stress_index + 1].value, sht_step[13:14, stress_index:stress_index + 1].value)
        discharge_max_stress = max(sht_step[14:15, stress_index:stress_index + 1].value, sht_step[15:16, stress_index:stress_index + 1].value)
    elif '压力(Kg)' in title_record:
        tem_max = title_record.index('压力(Kg)')
        charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max + 1].value)
        discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max + 1].value)
        stress_value_1 = max(charge_tem_np[index_list])
        stress_value_2 = max(charge_tem_np[index_list_1])
        stress_value_3 = max(charge_tem_np[index_list_2])
        stress_value_4 = max(charge_tem_np[index_list_3])
        charge_max_stress = max(stress_value_1, stress_value_3)
        discharge_max_stress = max(stress_value_2, stress_value_4)
    else:
        charge_max_stress = '-'
        discharge_max_stress = '-'

    # 温度
    try:
        tem_max = title_step.index('最高温度(℃)')
        charge_max_tem = sht_step[12:13, tem_max:tem_max + 1].value
        discharge_max_tem = sht_step[14:15, tem_max:tem_max + 1].value
    except ValueError:
        if '负载温度1' in title_record:
            tem_max = title_record.index('负载温度1')
        elif '负载温度1(℃)' in title_record:
            tem_max = title_record.index('负载温度1(℃)')
        elif '温度1(℃)' in title_record:
            tem_max = title_record.index('温度1(℃)')
        elif '辅助通道 TU1 T(°C)' in title_record:
            tem_max = title_record.index('辅助通道 TU1 T(°C)')
        else:
            tem_max = '-'
        if tem_max != '-':
            charge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max + 1].value)
            discharge_tem_np = np.array(sht_record[0:row_record, tem_max:tem_max + 1].value)
            charge_max_tem = max(charge_tem_np[index_list])
            discharge_max_tem = max(discharge_tem_np[index_list_1])
        else:
            charge_max_tem = '-'
            discharge_max_tem = '-'

    return charge_cap, discharge_cap, charge_energy, discharge_energy, charge_end_V, charge_standing_end_V, discharge_end_V, \
                discharge_standing_end_V, CC_time, total_time, CC_cap, None, None, None, \
           None, None, None, charge_max_tem, discharge_max_tem, charge_max_stress, discharge_max_stress,\
           V_value_record_1, V_value_record_2, charge_cap_value, charge_cap_value_1
