import os
import re
from datetime import datetime
# 新增 tkinter 匯入
import tkinter as tk
from tkinter import filedialog

def parse_xcs_size(xcs_content):
    match = re.search(r'CreateFinishedWorkpieceBox\("[^"]*",\s*(\d+\.?\d*),\s*(\d+\.?\d*),\s*(\d+\.?\d*)\)', xcs_content)
    if match:
        lpx, lpy, lpz = match.group(1), match.group(2), match.group(3)
        return lpx, lpy, lpz
    return "0", "0", "0"

def workplane_map(wp):
    return {
        'Top': 0,
        'Left': 1,
        'Back': 4,
        'Right': 3,
        'Front': 2,
        'Bottom': 5
    }.get(wp, 0)

def setup_map(a, b, c, d):
    return {
        ('0','0','0','0'): 2, # 左下
        ('0','1','0','0'): 1, # 左上
        ('1','1','0','0'): 4, # 右上
        ('1','0','0','0'): 3  # 右下
    }.get((a,b,c,d), 2)

def xcs_to_bpp_content(xcs_content, is_sp=False):
    lpx, lpy, lpz = parse_xcs_size(xcs_content)
    # SP規則：寬高各扣1
    if is_sp:
        try:
            lpx = str(float(lpx) - 1)
            lpy = str(float(lpy) - 1)
        except Exception:
            pass
    now = datetime.now()
    time_code = now.strftime('%H%M')
    header = '[HEADER]\nTYPE=BPP\nVER=150\n\n[DESCRIPTION]\n|\n\n[VARIABLES]\n'
    variables = f'PAN=LPX|{lpx}||4|\nPAN=LPY|{lpy}||4|\nPAN=LPZ|{lpz}||4|\nPAN=ORLST|"1"||3|\nPAN=SIMMETRY|1||1|\nPAN=TLCHK|0||1|\nPAN=TOOLING|""||3|\nPAN=CUSTSTR|$B$KInsiderExternalCam.Cam$V""||3|\nPAN=FCN|1.000000||2|\nPAN=XCUT|0||4|\nPAN=YCUT|0||4|\nPAN=JIGTH|0||4|\nPAN=CKOP|0||1|\nPAN=UNIQUE|0||1|\nPAN=MATERIAL|"wood"||3|\nPAN=PUTLST|""||3|\nPAN=OPPWKRS|0||1|\nPAN=UNICLAMP|0||1|\nPAN=CHKCOLL|0||1|\nPAN=WTPIANI|0||1|\nPAN=COLLTOOL|0||1|\nPAN=CALCEDTH|0||1|\nPAN=ENABLELABEL|0||1|\nPAN=LOCKWASTE|0||1|\nPAN=LOADEDGEOPT|0||1|\nPAN=ITLTYPE|0||1|\nPAN=RUNPAV|0||1|\nPAN=FLIPEND|0||1|\nPAN=ENABLEMACHLINKS|0||1|\nPAN=ENABLEPURSUITS|0||1|\nPAN=ENABLEFASTVERTBORINGS|0||1|\nPAN=FASTVERTBORINGSVALUE|0||4|\n\n[PROGRAM]\n'
    bpp_lines = []
    drill_index = 1
    current_workplane = 0
    current_setup = 2
    lines = xcs_content.splitlines()
    for line in lines:
        wp_match = re.match(r'SelectWorkplane\("([A-Za-z]+)"\);', line.strip())
        if wp_match:
            current_workplane = workplane_map(wp_match.group(1))
            continue
        setup_match = re.match(r'SetWorkpieceSetupPosition\((\d+),(\d+),(\d+),(\d+)\);', line.strip())
        if setup_match:
            current_setup = setup_map(setup_match.group(1), setup_match.group(2), setup_match.group(3), setup_match.group(4))
            continue
        drill_match = re.match(r'CreateDrill\((.*)\);', line.strip())
        if drill_match:
            params = drill_match.group(1)
            param_list = [p.strip().strip('"') for p in params.split(',')]
            name = param_list[0]
            x = param_list[1]
            y = param_list[2]
            depth = param_list[3]  # 修正：原本是 diameter，應為 depth
            diameter = param_list[4]  # 修正：原本是 depth，應為 diameter
            # 點孔/穿孔規則（型別轉換後比對）
            try:
                dia_f = float(diameter)
                dep_f = float(depth)
            except Exception:
                dia_f = diameter
                dep_f = depth
            if dia_f == 8 and dep_f == 2:
                diameter = '8.2'
            elif dia_f == 8 and dep_f == 26:
                diameter = '8.1'
                depth = '18.5'
            # SP規則：TOP面X.Y各扣0.5
            if is_sp and current_workplane == 0:
                try:
                    x = str(float(x) - 0.5)
                    y = str(float(y) - 0.5)
                except Exception:
                    pass
            if diameter == '26':
                tool = '2'
                flag = '1'
            elif diameter == '2':
                tool = '2'
                flag = '0'
            else:
                tool = '12'
                flag = '0'
            pid = f'P{1000+drill_index}'
            uniq = f'{time_code}{drill_index:04d}'
            # 決定加工型態與標籤
            if current_workplane in [1,2,3,4]:
                op_type = 'BH'
            else:
                op_type = 'BV'
            # 組成bpp行
            bpp = f'@ {op_type}, "", "", {uniq}, "", 0 : {current_workplane}, "{current_setup}", {x}, {y}, 0, {depth}, {diameter}, {flag}, -1, 32, 32, 50, 0, 45, 0, "", 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, -1, "{pid}", 0, "", "", 0, 0, 0, 0, 0, "", 0, 0, 0, 0, 0, "", "", "{op_type}", 0, 0, 0, 0, -1, 0, 0, 0'
            bpp_lines.append(bpp)
            drill_index += 1
    tail = '\n\n[VBSCRIPT]\n\n[MACRODATA]\n\n[TDCODES]\n\n[PCF]\n\n[TOOLING]\n\n[SUBPROGS]\n\n'
    return header + variables + '\n'.join(bpp_lines) + tail

def batch_convert_xcs_to_bpp(input_folder, output_folder):
    for filename in os.listdir(input_folder):
        if filename.lower().endswith('.xcs'):
            xcs_path = os.path.join(input_folder, filename)
            bpp_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '.bpp')
            with open(xcs_path, 'r', encoding='utf-8') as f:
                xcs_content = f.read()
            # 判斷SP規則
            is_sp = filename.upper().startswith('SP')
            bpp_content = xcs_to_bpp_content(xcs_content, is_sp=is_sp)
            with open(bpp_path, 'w', encoding='utf-8') as f:
                f.write(bpp_content)
    print('批次轉換完成！')

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    print('請選擇匯入（.xcs）資料夾...')
    input_folder = filedialog.askdirectory(title='選擇匯入資料夾')
    print('請選擇匯出（.bpp）資料夾...')
    output_folder = filedialog.askdirectory(title='選擇匯出資料夾')
    if input_folder and output_folder:
        batch_convert_xcs_to_bpp(input_folder, output_folder)
    else:
        print('未選擇資料夾，程式結束。') 