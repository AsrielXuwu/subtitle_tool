import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from tkinter import font as tkfont
import pandas as pd
import os
import re
import json
import zipfile
import math
import platform

# ======= 新增：尝试引入视频帧提取与渲染库 =======
try:
    import cv2
    from PIL import Image, ImageTk
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False

# ======= 新增：全局 ASS 内存暂存字典 =======
global_ass_memory_cache = {}

# ================= 预设配置文件 (跨平台兼容版) =================
# 解决 PyInstaller 在 macOS 打包后无权限在当前目录读写 json 文件的问题
def get_config_dir():
    if platform.system() == 'Windows':
        # Windows: 存在 AppData 下
        conf_dir = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'SubtitleToolbox')
    else:
        # macOS / Linux: 存在用户主目录下的隐藏文件夹
        conf_dir = os.path.join(os.path.expanduser('~'), '.subtitle_toolbox')
    os.makedirs(conf_dir, exist_ok=True)
    return conf_dir

CONFIG_DIR = get_config_dir()
PRESET_FILE_REP = os.path.join(CONFIG_DIR, "column_presets.json")
PRESET_FILE_SPLIT = os.path.join(CONFIG_DIR, "split_presets.json")
PRESET_FILE_ASS = os.path.join(CONFIG_DIR, "ass_presets.json")

DEFAULT_PRESETS_REP = ["A, B, E", "B, C, I"]
DEFAULT_PRESETS_SPLIT = ["A, B, C, D", "A, B, C, D, E"]
DEFAULT_PRESETS_ASS = {
    "默认样式": {
        "play_res_x": "1080", "play_res_y": "1920",
        "n_font": "SimHei", "n_size": "60", "n_color": "#FFFFFF", "n_out_color": "#000000",
        "n_margin_v": "20", "n_margin_lr": "20", "n_outline": "2", "n_align": "2", "n_shadow": "0", "n_bold": 0, "n_italic": 0,
        "s_font": "SimHei", "s_size": "60", "s_color": "#26E3FF", "s_out_color": "#000000",
        "s_margin_v": "850", "s_margin_lr": "20", "s_outline": "2", "s_align": "8", "s_shadow": "0", "s_bold": 0, "s_italic": 0
    }
}

def load_presets(filepath, default_data):
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return default_data.copy() if isinstance(default_data, list) else dict(default_data)
    return default_data.copy() if isinstance(default_data, list) else dict(default_data)

def save_presets_to_file(filepath, presets):
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(presets, f, ensure_ascii=False)
    except Exception as e:
        print(f"保存预设 {filepath} 失败:", e)

def col2num(col_str):
    num = 0
    for c in col_str.upper():
        if 'A' <= c <= 'Z':
            num = num * 26 + (ord(c) - ord('A')) + 1
    return num - 1

global_ass_preset_cbs = []

def get_ass_resolution(filepath):
    """解析 ASS 文件的 PlayResX 和 PlayResY"""
    rx, ry = "1080", "1920"
    if os.path.exists(filepath):
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            for line in f:
                if line.startswith('PlayResX:'): rx = line.split(':')[1].strip()
                elif line.startswith('PlayResY:'): ry = line.split(':')[1].strip()
                elif line.startswith('[Events]'): break
    return rx, ry

def update_all_ass_preset_cbs():
    keys = list(current_presets_ass.keys())
    for cb in global_ass_preset_cbs:
        cb['values'] = keys

def create_ass_preset_bar(parent, n_vars, s_vars, c_btns, preset_cb_var, res_vars=None):
    f_ps = ttk.Frame(parent)
    ttk.Label(f_ps, text="样式预设:").pack(side=tk.LEFT, padx=(0,5))
    cb = ttk.Combobox(f_ps, textvariable=preset_cb_var, values=list(current_presets_ass.keys()), width=15)
    cb.pack(side=tk.LEFT, padx=5)
    global_ass_preset_cbs.append(cb)
    
    def load_p(event=None):
        name = preset_cb_var.get()
        if name in current_presets_ass:
            d = current_presets_ass[name]
            if res_vars and len(res_vars) >= 2:
                res_vars[0].set(d.get("play_res_x", "1080"))
                res_vars[1].set(d.get("play_res_y", "1920"))
            n_vars[0].set(d.get("n_font", "SimHei")); n_vars[1].set(d.get("n_size", "60"))
            n_vars[2].set(d.get("n_color", "#FFFFFF")); n_vars[3].set(d.get("n_out_color", "#000000"))
            n_vars[4].set(d.get("n_margin_v", "20")); n_vars[5].set(d.get("n_margin_lr", "20")); n_vars[6].set(d.get("n_outline", "2"))
            n_vars[7].set(d.get("n_align", "2")); n_vars[8].set(d.get("n_shadow", "0"))
            n_vars[9].set(d.get("n_bold", 0)); n_vars[10].set(d.get("n_italic", 0))
            if len(n_vars) > 11:
                n_vars[11].set(d.get("n_alpha", "00")); n_vars[12].set(d.get("n_out_alpha", "00"))
            update_color_btn(c_btns[0], d.get("n_color", "#FFFFFF")); update_color_btn(c_btns[1], d.get("n_out_color", "#000000"))
            
            if s_vars and len(s_vars) >= 11 and len(c_btns) == 4:
                s_vars[0].set(d.get("s_font", "SimHei")); s_vars[1].set(d.get("s_size", "60"))
                s_vars[2].set(d.get("s_color", "#26E3FF")); s_vars[3].set(d.get("s_out_color", "#000000"))
                s_vars[4].set(d.get("s_margin_v", "850")); s_vars[5].set(d.get("s_margin_lr", "20")); s_vars[6].set(d.get("s_outline", "2"))
                s_vars[7].set(d.get("s_align", "8")); s_vars[8].set(d.get("s_shadow", "0"))
                s_vars[9].set(d.get("s_bold", 0)); s_vars[10].set(d.get("s_italic", 0))
                if len(s_vars) > 11:
                    s_vars[11].set(d.get("s_alpha", "00")); s_vars[12].set(d.get("s_out_alpha", "00"))
                update_color_btn(c_btns[2], d.get("s_color", "#26E3FF")); update_color_btn(c_btns[3], d.get("s_out_color", "#000000"))

    def save_p():
        name = preset_cb_var.get().strip()
        if not name: return messagebox.showwarning("提示", "请输入预设名称！")
        d = current_presets_ass.get(name, DEFAULT_PRESETS_ASS["默认样式"].copy())
        if res_vars and len(res_vars) >= 2:
            d.update({"play_res_x": res_vars[0].get(), "play_res_y": res_vars[1].get()})
        d.update({"n_font": n_vars[0].get(), "n_size": n_vars[1].get(), "n_color": n_vars[2].get(), "n_out_color": n_vars[3].get(), "n_margin_v": n_vars[4].get(), "n_margin_lr": n_vars[5].get(), "n_outline": n_vars[6].get(), "n_align": n_vars[7].get(), "n_shadow": n_vars[8].get(), "n_bold": n_vars[9].get(), "n_italic": n_vars[10].get()})
        if len(n_vars) > 11:
            d.update({"n_alpha": n_vars[11].get(), "n_out_alpha": n_vars[12].get()})
        if s_vars: 
            d.update({"s_font": s_vars[0].get(), "s_size": s_vars[1].get(), "s_color": s_vars[2].get(), "s_out_color": s_vars[3].get(), "s_margin_v": s_vars[4].get(), "s_margin_lr": s_vars[5].get(), "s_outline": s_vars[6].get(), "s_align": s_vars[7].get(), "s_shadow": s_vars[8].get(), "s_bold": s_vars[9].get(), "s_italic": s_vars[10].get()})
            if len(s_vars) > 11:
                d.update({"s_alpha": s_vars[11].get(), "s_out_alpha": s_vars[12].get()})
        current_presets_ass[name] = d; save_presets_to_file(PRESET_FILE_ASS, current_presets_ass); update_all_ass_preset_cbs(); preset_cb_var.set(name); messagebox.showinfo("提示", "样式预设保存成功！")
        
    def del_p():
        name = preset_cb_var.get().strip()
        if name in current_presets_ass:
            del current_presets_ass[name]; save_presets_to_file(PRESET_FILE_ASS, current_presets_ass); update_all_ass_preset_cbs()
            keys = list(current_presets_ass.keys()); preset_cb_var.set(keys[0] if keys else ""); load_p(); messagebox.showinfo("提示", "预设已删除！")
        else: messagebox.showwarning("提示", "未找到该预设！")

    cb.bind("<<ComboboxSelected>>", load_p); ttk.Button(f_ps, text="保存预设", command=save_p).pack(side=tk.LEFT, padx=5); ttk.Button(f_ps, text="删除选中", command=del_p).pack(side=tk.LEFT, padx=5)
    if preset_cb_var.get(): load_p()
    return f_ps

# ================= 核心功能逻辑 =================

def process_ass_merge(dir1, dir2, out_dir):
    files1 = {f for f in os.listdir(dir1) if f.lower().endswith('.ass')}
    files2 = {f for f in os.listdir(dir2) if f.lower().endswith('.ass')}
    
    common_files = files1.intersection(files2)
    if not common_files: raise ValueError("两个文件夹中没有找到同名的 ASS 文件！")
    
    os.makedirs(out_dir, exist_ok=True)
    processed_count = 0
    
    for file in common_files:
        path1 = os.path.join(dir1, file)
        path2 = os.path.join(dir2, file)
        
        with open(path1, 'r', encoding='utf-8-sig') as f: lines1 = f.read().split('\n')
        with open(path2, 'r', encoding='utf-8-sig') as f: lines2 = f.read().split('\n')
        
        def parse_ass_parts(lines):
            info, styles, events = [], [], []
            curr = "info"
            for l in lines:
                ls = l.strip()
                if ls.startswith('[V4+ Styles]'): curr = "styles"
                elif ls.startswith('[Events]'): curr = "events"
                
                if curr == "info": info.append(l)
                elif curr == "styles": styles.append(l)
                elif curr == "events": events.append(l)
            return info, styles, events
            
        info1, styles1, events1 = parse_ass_parts(lines1)
        info2, styles2, events2 = parse_ass_parts(lines2)
        
        # 1. 头部信息：取文件1的
        merged_info = info1
        
        # 2. 样式去重
        merged_styles = []
        seen_style_names = set()
        
        # 优先保留 [V4+ Styles] 和 Format 行
        for l in styles1:
            if not l.strip().startswith('Style:'):
                merged_styles.append(l)
            else:
                name = l.split('Style:')[1].split(',')[0].strip()
                if name not in seen_style_names:
                    merged_styles.append(l)
                    seen_style_names.add(name)
                    
        for l in styles2:
            if l.strip().startswith('Style:'):
                name = l.split('Style:')[1].split(',')[0].strip()
                if name not in seen_style_names:
                    merged_styles.append(l)
                    seen_style_names.add(name)
                    
        # 3. 字幕行合并：文件夹2 拼在 文件夹1 顶部
        merged_events = []
        has_events_header = False
        
        # 先加入文件2的 Dialogue (跳过 [Events] 和 Format 行)
        for l in events2:
            if l.strip().startswith('[Events]') or l.strip().startswith('Format:'):
                if not has_events_header:
                    merged_events.append(l)
                    has_events_header = True
            elif l.strip().startswith('Dialogue:'):
                merged_events.append(l)
                
        # 再加入文件1的 Dialogue
        for l in events1:
            if l.strip().startswith('[Events]') or l.strip().startswith('Format:'):
                if not has_events_header:
                    merged_events.append(l)
                    has_events_header = True
            elif l.strip().startswith('Dialogue:'):
                merged_events.append(l)

        # 组合写入
        out_content = "\n".join(merged_info) + "\n" + "\n".join(merged_styles) + "\n" + "\n".join(merged_events) + "\n"
        with open(os.path.join(out_dir, file), 'w', encoding='utf-8') as f:
            f.write(out_content)
            
        processed_count += 1
        
    return processed_count

def process_time_split(in_dir, out_norm_dir, out_scr_dir, mode="SRT"):
    files = [f for f in os.listdir(in_dir) if f.lower().endswith('.srt' if mode == "SRT" else '.ass')]
    if not files: raise ValueError(f"输入文件夹中没有找到 {mode} 文件！")
    
    os.makedirs(out_norm_dir, exist_ok=True)
    os.makedirs(out_scr_dir, exist_ok=True)
    
    processed_count = 0

    def srt_time_to_ms(time_str):
        time_str = time_str.replace(',', '.')
        parts = time_str.split(':')
        h = int(parts[0])
        m = int(parts[1])
        s, ms = parts[2].split('.')
        return (h * 3600 + m * 60 + int(s)) * 1000 + int(ms)

    def ass_time_to_ms(time_str):
        parts = time_str.split(':')
        h = int(parts[0])
        m = int(parts[1])
        s, ms = parts[2].split('.')
        return (h * 3600 + m * 60 + int(s)) * 1000 + int(ms) * 10
    
    for file in files:
        filepath = os.path.join(in_dir, file)
        base_name = os.path.splitext(file)[0]
        
        if mode == "SRT":
            blocks = parse_srt_file(filepath)
            split_idx = -1
            prev_end_ms = -1
            
            # 自动扫描分界点
            for i, block in enumerate(blocks):
                start_str, end_str = block['Timeline'].split(' --> ')
                curr_start_ms = srt_time_to_ms(start_str.strip())
                curr_end_ms = srt_time_to_ms(end_str.strip())
                
                # 如果当前行开始时间早于上一行结束时间，找到分界点
                if prev_end_ms != -1 and curr_start_ms < prev_end_ms:
                    split_idx = i
                    break
                prev_end_ms = curr_end_ms
                
            # 如果没有找到任何倒挂，默认全为对白
            if split_idx == -1: split_idx = 0
            
            norm_blocks_raw = blocks[split_idx:]
            scr_blocks_raw = blocks[:split_idx]
            
            # 重新生成 ID
            norm_blocks = [f"{i+1}\n{b['Timeline']}\n{b['Text']}\n" for i, b in enumerate(norm_blocks_raw)]
            scr_blocks = [f"{i+1}\n{b['Timeline']}\n{b['Text']}\n" for i, b in enumerate(scr_blocks_raw)]
            
            with open(os.path.join(out_norm_dir, base_name + '.srt'), 'w', encoding='utf-8') as f:
                f.write("\n".join(norm_blocks))
            with open(os.path.join(out_scr_dir, base_name + '.srt'), 'w', encoding='utf-8') as f:
                f.write("\n".join(scr_blocks))
            processed_count += 1
            
        elif mode == "ASS":
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                lines = f.read().split('\n')
            
            h_lines, s_lines, ev_lines = [], [], []
            curr = "info"
            for line in lines:
                l = line.strip()
                if l.startswith('[V4+ Styles]'): curr = "styles"
                elif l.startswith('[Events]'): curr = "events"
                
                if curr == "info": h_lines.append(line)
                elif curr == "styles": s_lines.append(line)
                elif curr == "events": ev_lines.append(line)
                
            dialogue_lines = []
            other_ev_lines = []
            for ev in ev_lines:
                if ev.startswith('Dialogue:'):
                    dialogue_lines.append(ev)
                else:
                    other_ev_lines.append(ev)
                    
            split_idx = -1
            prev_end_ms = -1
            
            # 自动扫描分界点
            for i, ev in enumerate(dialogue_lines):
                parts = ev.split(',', 9)
                if len(parts) >= 10:
                    curr_start_ms = ass_time_to_ms(parts[1].strip())
                    curr_end_ms = ass_time_to_ms(parts[2].strip())
                    
                    if prev_end_ms != -1 and curr_start_ms < prev_end_ms:
                        split_idx = i
                        break
                    prev_end_ms = curr_end_ms
            
            if split_idx == -1: split_idx = 0
            
            screen_ev = other_ev_lines + dialogue_lines[:split_idx]
            normal_ev = other_ev_lines + dialogue_lines[split_idx:]
            
            with open(os.path.join(out_norm_dir, file), 'w', encoding='utf-8') as f:
                f.write("\n".join(h_lines) + "\n" + "\n".join(s_lines) + "\n" + "\n".join(normal_ev) + "\n")
            with open(os.path.join(out_scr_dir, file), 'w', encoding='utf-8') as f:
                f.write("\n".join(h_lines) + "\n" + "\n".join(s_lines) + "\n" + "\n".join(screen_ev) + "\n")
            processed_count += 1
            
    return processed_count

def process_split(input_file, out_dir_src, out_dir_tgt, cols_list):
    if input_file.lower().endswith('.csv'): df = pd.read_csv(input_file)
    else: df = pd.read_excel(input_file)
        
    c_file, c_id, c_time = col2num(cols_list[0]), col2num(cols_list[1]), col2num(cols_list[2])
    c_texts = [col2num(c) for c in cols_list[3:]]
    max_col_needed = max([c_file, c_id, c_time] + c_texts)
    if max_col_needed >= len(df.columns): raise ValueError(f"指定的列字母超出范围！\n当前表格只有 {len(df.columns)} 列。")
       
    file_data = {}
    for index, row in df.iterrows():
        filename_val = str(row.iloc[c_file]).strip()
        if pd.isna(row.iloc[c_file]) or not filename_val or filename_val == "nan": continue
        
        # 修复：防止 Excel 将 0001 吞成 1，自动强制补齐四位数字
        if filename_val.endswith('.0'): filename_val = filename_val[:-2]
        base_name = filename_val[:-4] if filename_val.lower().endswith('.srt') else filename_val
        if base_name.isdigit(): base_name = base_name.zfill(4)
        filename_val = base_name + '.srt'
            
        id_val = str(row.iloc[c_id]).strip()
        if id_val.endswith('.0'): id_val = id_val[:-2]
        time_val = str(row.iloc[c_time]).strip()
 
        if filename_val not in file_data: file_data[filename_val] = []
        file_data[filename_val].append((id_val, time_val, row))
        
    if not file_data: raise ValueError("未找到有效的数据，请检查列名是否正确！")

    for i, c_text in enumerate(c_texts):
        lang_dir = out_dir_src if i == 0 else out_dir_tgt
        if not lang_dir: continue 
        os.makedirs(lang_dir, exist_ok=True)
 
        for filename_val, rows in file_data.items():
            
            # --- 智能排序机制 ---
            def get_sort_key(item):
                try: return (0, int(item[0]))
                except: return (1, item[0])
            rows.sort(key=get_sort_key)
            
            srt_content = []
            for id_val, time_val, row in rows:
                text_val = row.iloc[c_text]
                text_val = "" if pd.isna(text_val) else str(text_val).strip()
                srt_content.append(f"{id_val}\n{time_val}\n{text_val}\n")
            with open(os.path.join(lang_dir, filename_val), 'w', encoding='utf-8') as f:
                f.write("\n".join(srt_content))
    return len(file_data)

def process_split_mode2(input_file, out_dir_src, out_dir_tgt, cols_list, sheet_name):
    """XLSX 转 SRT 模式2：无 ID 列，按行自增，支持双目录输出"""
    if input_file.lower().endswith('.csv'): df = pd.read_csv(input_file)
    else: df = pd.read_excel(input_file, sheet_name=sheet_name if sheet_name else 0)
        
    c_file, c_time = col2num(cols_list[0]), col2num(cols_list[1])
    c_texts = [col2num(c) for c in cols_list[2:]]
    max_col_needed = max([c_file, c_time] + c_texts)
    if max_col_needed >= len(df.columns): raise ValueError("指定的列字母超出表格实际范围！")
       
    file_data = {}
    for index, row in df.iterrows():
        filename_val = str(row.iloc[c_file]).strip()
        if pd.isna(row.iloc[c_file]) or not filename_val or filename_val == "nan": continue
        
        # 修复：防止 Excel 将 0001 吞成 1，自动强制补齐四位数字
        if filename_val.endswith('.0'): filename_val = filename_val[:-2]
        base_name = filename_val[:-4] if filename_val.lower().endswith('.srt') else filename_val
        if base_name.isdigit(): base_name = base_name.zfill(4)
        filename_val = base_name + '.srt'
            
        time_val = str(row.iloc[c_time]).strip()
        id_val = str(index + 1)
 
        if filename_val not in file_data: file_data[filename_val] = []
        file_data[filename_val].append((id_val, time_val, row))
        
    if not file_data: raise ValueError("未找到有效的数据，请检查列名是否正确！")

    for i, c_text in enumerate(c_texts):
        lang_dir = out_dir_src if i == 0 else out_dir_tgt
        if not lang_dir: continue
        os.makedirs(lang_dir, exist_ok=True)
 
        for filename_val, rows in file_data.items():
            srt_content = []
            for id_val, time_val, row in rows:
                text_val = row.iloc[c_text]
                text_val = "" if pd.isna(text_val) else str(text_val).strip()
                srt_content.append(f"{id_val}\n{time_val}\n{text_val}\n")
            with open(os.path.join(lang_dir, filename_val), 'w', encoding='utf-8') as f:
                f.write("\n".join(srt_content))
    return len(file_data)

def parse_srt_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f: content = f.read().strip()
    blocks = re.split(r'\n\s*\n', content)
    parsed_data = []
    for block in blocks:
        lines = block.strip().split('\n')
        if len(lines) >= 3:
            idx = lines[0].strip()
            timeline = lines[1].strip()
            text = "\n".join(lines[2:]).strip()
            parsed_data.append({'ID': idx, 'Timeline': timeline, 'Text': text})
    return parsed_data

def process_merge(src_dir, tgt_dir, src_lang_name, tgt_lang_name, output_excel):
    src_files = {f for f in os.listdir(src_dir) if f.lower().endswith('.srt')}
    tgt_files = {f for f in os.listdir(tgt_dir) if f.lower().endswith('.srt')}
    common_files = src_files.intersection(tgt_files)
    if not common_files: raise ValueError("两个目录中没有找到同名的 SRT 文件，无法合并！")
        
    # 修复：防止文件 1, 10, 2 乱序，强制使用数字自然数大小进行排序
    def get_sort_key(f):
        b = os.path.splitext(f)[0]
        return int(b) if b.isdigit() else b

    master_data = []
    for file_name in sorted(list(common_files), key=get_sort_key):
        base_name = os.path.splitext(file_name)[0]
        # 修复：强行补充四位数字（写入 Excel 的时候保证是 0001）
        if base_name.isdigit(): base_name = base_name.zfill(4)
        
        src_path, tgt_path = os.path.join(src_dir, file_name), os.path.join(tgt_dir, file_name)
        src_data, tgt_data = parse_srt_file(src_path), parse_srt_file(tgt_path)
        
        for i, src_item in enumerate(src_data):
            tgt_text = tgt_data[i]['Text'] if i < len(tgt_data) else ""
            master_data.append({
                'Episode': base_name, 'ID': src_item['ID'], 'Time': src_item['Timeline'],
                src_lang_name: src_item['Text'], tgt_lang_name: tgt_text
            })
    pd.DataFrame(master_data).to_excel(output_excel, index=False)
    return len(common_files)

def process_merge_mode2(src_dir, tgt_dir, src_lang_name, tgt_lang_name, output_excel):
    """SRT 合并为 XLSX 模式2：无ID列，保留文件和时间轴，支持单语言合并"""
    src_files = {f for f in os.listdir(src_dir) if f.lower().endswith('.srt')} if os.path.exists(src_dir) else set()
    if not src_files: raise ValueError("源语言目录中没有找到 SRT 文件！")
        
    # 修复：防止文件 1, 10, 2 乱序，强制使用数字自然数大小进行排序
    def get_sort_key(f):
        b = os.path.splitext(f)[0]
        return int(b) if b.isdigit() else b

    master_data = []
    for file_name in sorted(list(src_files), key=get_sort_key):
        base_name = os.path.splitext(file_name)[0]
        # 修复：强行补充四位数字（写入 Excel 的时候保证是 0001）
        if base_name.isdigit(): base_name = base_name.zfill(4)
        
        src_path = os.path.join(src_dir, file_name)
        src_data = parse_srt_file(src_path)
        
        tgt_data = []
        if tgt_dir and os.path.exists(tgt_dir):
            tgt_path = os.path.join(tgt_dir, file_name)
            if os.path.exists(tgt_path):
                tgt_data = parse_srt_file(tgt_path)
        
        for i, src_item in enumerate(src_data):
            tgt_text = tgt_data[i]['Text'] if i < len(tgt_data) else ""
            row_dict = {
                'Episode': base_name, 
                'Time': src_item['Timeline'],
                src_lang_name if src_lang_name else '源语言': src_item['Text']
            }
            if tgt_dir and tgt_lang_name:
                row_dict[tgt_lang_name] = tgt_text
            master_data.append(row_dict)
            
    pd.DataFrame(master_data).to_excel(output_excel, index=False)
    return len(src_files)

def process_replace(report_file, srt_dir, out_summary, col_filename_str, col_id_str, col_text_str):
    if report_file.lower().endswith('.csv'): df = pd.read_csv(report_file)
    else: df = pd.read_excel(report_file)
        
    c_file, c_id, c_text = col2num(col_filename_str), col2num(col_id_str), col2num(col_text_str)
    if max(c_file, c_id, c_text) >= len(df.columns): raise ValueError(f"指定的列字母超出范围！")
        
    srt_cache, summary_data = {}, []
    for index, row in df.iterrows():
        filename_val, id_val, text_val = str(row.iloc[c_file]).strip(), str(row.iloc[c_id]).strip(), row.iloc[c_text]
        if id_val.endswith('.0'): id_val = id_val[:-2]
        if pd.isna(text_val) or str(text_val).strip() == "" or filename_val == "nan": continue
            
        text_val = str(text_val).strip()
        basename = os.path.basename(filename_val.replace('\\', '/'))
        if not basename.lower().endswith('.srt'): basename += '.srt'
        filepath = os.path.join(srt_dir, basename)
        if not os.path.exists(filepath): continue
            
        if basename not in srt_cache: srt_cache[basename] = parse_srt_file(filepath)
        for block in srt_cache[basename]:
            if block['ID'] == id_val:
                old_text = block['Text']
                if old_text != text_val:
                    block['Text'] = text_val
                    summary_data.append({'SRT文件名': basename, '字幕ID': id_val, '原字幕内容': old_text, '替换后新内容': text_val})
                break
                
    for basename, blocks in srt_cache.items():
        srt_content = [f"{block['ID']}\n{block['Timeline']}\n{block['Text']}\n" for block in blocks]
        with open(os.path.join(srt_dir, basename), 'w', encoding='utf-8') as f:
            f.write("\n".join(srt_content))
    if summary_data: pd.DataFrame(summary_data).to_excel(out_summary, index=False)
    return len(summary_data), len(srt_cache)

def process_zip(target_dir, output_dir, max_files):
    files = sorted([f for f in os.listdir(target_dir) if os.path.isfile(os.path.join(target_dir, f))])
    if not files: raise ValueError("目标文件夹中没有任何文件！")
    total_files = len(files)
    num_zips = math.ceil(total_files / max_files)
        
    for i in range(num_zips):
        chunk = files[i*max_files : (i+1)*max_files]
        first_name, last_name = os.path.splitext(chunk[0])[0], os.path.splitext(chunk[-1])[0]
        zip_filename = f"{first_name}.zip" if first_name == last_name else f"{first_name}_{last_name}.zip"
        with zipfile.ZipFile(os.path.join(output_dir, zip_filename), 'w', compression=zipfile.ZIP_STORED) as zf:
            for f in chunk: zf.write(os.path.join(target_dir, f), f)
    return total_files, num_zips

def ass_to_hex(ass_str):
    try:
        s = ass_str.upper().replace('&H', '')
        if len(s) >= 8: s = s[-6:]
        elif len(s) < 6: s = s.zfill(6)
        b, g, r = s[0:2], s[2:4], s[4:6]
        return f"#{r}{g}{b}"
    except: return "#FFFFFF"

# ================= ASS 相关辅助模块 =================

def srt_to_ass_time(srt_time):
    srt_time = srt_time.strip()
    time_part, ms_part = srt_time.split(',') if ',' in srt_time else (srt_time.split('.') if '.' in srt_time else (srt_time, "000"))
    parts = time_part.split(':')
    h, m, s = parts if len(parts) == 3 else ("00", parts[0], parts[1])
    return f"{int(h)}:{m}:{s}.{ms_part[:2]}"

def hex2ass_with_alpha(hex_str, alpha_str="00"):
    if not hex_str or len(hex_str) != 7: hex_str = "#FFFFFF"
    r, g, b = hex_str[1:3], hex_str[3:5], hex_str[5:7]
    a = str(alpha_str).strip().upper()
    if len(a) != 2: a = "00"
    return f"&H{a}{b}{g}{r}"

def clean_ass_text(txt):
    lines = [l.strip() for l in re.split(r'\\N|\n', txt) if l.strip()]
    return "\\N".join(lines)

def rename_style_line(style_line, new_name):
    if not style_line.startswith("Style:"): return style_line
    prefix, rest = style_line.split(":", 1)
    parts = rest.split(",")
    parts[0] = f" {new_name}"
    return f"{prefix}:{','.join(parts)}"

def replace_font_in_style(style_line, new_font):
    """直接替换样式字符串中的字体名称"""
    if not style_line.startswith("Style:"): return style_line
    prefix, rest = style_line.split(":", 1)
    parts = rest.split(",")
    if len(parts) > 1:
        parts[1] = new_font
    return f"{prefix}:{','.join(parts)}"

def build_ass_style_line(name, f_name, f_size, color, out_color, mv, mlr, outline, align="2", shadow="0", bold=0, italic=0, alpha="00", out_alpha="00"):
    b_val = "-1" if str(bold) == "1" else "0"
    i_val = "-1" if str(italic) == "1" else "0"
    c_ass = color if str(color).startswith("&H") else hex2ass_with_alpha(color, alpha)
    oc_ass = out_color if str(out_color).startswith("&H") else hex2ass_with_alpha(out_color, out_alpha)
    return f"Style: {name},{f_name},{f_size},{c_ass},&H000000FF,{oc_ass},&H00000000,{b_val},{i_val},0,0,100,100,0,0,1,{outline},{shadow},{align},{mlr},{mlr},{mv},1"

def merge_ass_dialogues(dialogues_list, filename, report_list):
    if not dialogues_list: return []
    def parse_diag(line):
        prefix, content = line.split(":", 1)
        parts = content.split(",", 9)
        return {"start": parts[1].strip(), "end": parts[2].strip(), "style": parts[3].strip(), "text": parts[9].strip(), "raw": parts, "prefix": prefix}
    
    parsed = [parse_diag(d) for d in dialogues_list]
    merged, current = [], parsed[0]
    
    for nxt in parsed[1:]:
        if current["text"] == nxt["text"] and current["style"] == nxt["style"] and current["end"] == nxt["start"]:
            report_list.append({"文件名": filename, "合并文本": current["text"], "原时间轴1": f"{current['start']} --> {current['end']}", "原时间轴2": f"{nxt['start']} --> {nxt['end']}", "合并后时间轴": f"{current['start']} --> {nxt['end']}"})
            current["end"] = nxt["end"]
            current["raw"][2] = nxt["end"]
        else:
            merged.append(current["prefix"] + ":" + ",".join(current["raw"]))
            current = nxt
            
    merged.append(current["prefix"] + ":" + ",".join(current["raw"]))
    return merged

# ================= 业务流程处理 =================

def process_srt_to_ass(input_dir, out_dir, bracket_str, regex_text, custom_style_dict, do_merge, merge_report_path, style_mode, ref_cfg):
    patterns = []
    for b in bracket_str.replace('，', ',').split(','):
        b = b.strip()
        if len(b) >= 2:
            half = len(b) // 2
            l, r = re.escape(b[:half]), re.escape(b[half:])
            patterns.append(f"{l}[\\s\\S]*?{r}")
    bracket_regex = "|".join(patterns) if patterns else ""

    replacements = []
    for line in regex_text.split('\n'):
        if '>>>' in line:
            pat, repl = line.split('>>>', 1)
            repl_python = re.sub(r'\$(\d+)', r'\\\1', repl.strip())
            replacements.append((pat.strip(), repl_python))

    files = [f for f in os.listdir(input_dir) if f.lower().endswith('.srt')]
    if not files: raise ValueError("指定的输入文件夹中没有找到 .srt 文件！")

    merge_reports = []

    for file in files:
        filepath = os.path.join(input_dir, file)
        out_path = os.path.join(out_dir, file.rsplit('.', 1)[0] + '.ass')

        blocks = parse_srt_file(filepath)
        screen_events, normal_events = [], []

        for block in blocks:
            text = block['Text']
            if bracket_regex:
                s_texts = re.findall(bracket_regex, text)
                n_text = re.sub(bracket_regex, "", text).strip()
            else:
                s_texts, n_text = [], text

            n_text = clean_ass_text(n_text)
            s_text = clean_ass_text("\\N".join(s_texts))

            for pat, repl in replacements:
                if n_text: n_text = re.sub(pat, repl, n_text)
                if s_text: s_text = re.sub(pat, repl, s_text)

            start_ass = srt_to_ass_time(block['Timeline'].split(' --> ')[0])
            end_ass = srt_to_ass_time(block['Timeline'].split(' --> ')[1])

            if s_text: screen_events.append(f"Dialogue: 0,{start_ass},{end_ass},画面字,,0,0,0,,{s_text}")
            if n_text: normal_events.append(f"Dialogue: 0,{start_ass},{end_ass},对白字幕,,0,0,0,,{n_text}")

        if style_mode == 0:
            resx = custom_style_dict.get('play_res_x', '1080')
            resy = custom_style_dict.get('play_res_y', '1920')
        else:
            resx, resy = get_ass_resolution(ref_cfg['ref_path'])
            
        srt_script_info = f"[Script Info]\nScriptType: v4.00+\nPlayResX: {resx}\nPlayResY: {resy}\n"
        srt_styles_block = "[V4+ Styles]\nFormat: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding\n"
        
        if style_mode == 0:
            d = custom_style_dict
            srt_styles_block += build_ass_style_line("对白字幕", d['n_font'], d['n_size'], d['n_color'], d['n_out_color'], d['n_margin_v'], d['n_margin_lr'], d['n_outline'], d.get('n_align','2'), d.get('n_shadow','0'), d.get('n_bold',0), d.get('n_italic',0), d.get('n_alpha','00'), d.get('n_out_alpha','00')) + "\n"
            srt_styles_block += build_ass_style_line("画面字", d['s_font'], d['s_size'], d['s_color'], d['s_out_color'], d['s_margin_v'], d['s_margin_lr'], d['s_outline'], d.get('s_align','8'), d.get('s_shadow','0'), d.get('s_bold',0), d.get('s_italic',0), d.get('s_alpha','00'), d.get('s_out_alpha','00'))
        else:
            ref_styles = scan_all_styles_from_ass(ref_cfg['ref_path'])
            n_line = rename_style_line(ref_styles.get(ref_cfg['n_style'], build_ass_style_line("对白字幕", "Arial", "60", "&H00FFFFFF", "&H00000000", "20", "20", "2")), "对白字幕")
            s_line = rename_style_line(ref_styles.get(ref_cfg['s_style'], build_ass_style_line("画面字", "Arial", "60", "&H00FFFFFF", "&H00000000", "850", "20", "2")), "画面字")
            
            if ref_cfg['font_mode'] == 1:
                n_line = replace_font_in_style(n_line, ref_cfg['override_font'])
                s_line = replace_font_in_style(s_line, ref_cfg['override_font'])
            
            srt_styles_block += n_line + "\n" + s_line

        ass_events_str = "[Events]\nFormat: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text\n"
        ass_events_str += "\n".join(screen_events) + ("\n" if screen_events else "")
        ass_events_str += "\n".join(normal_events) + "\n"

        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(f"{srt_script_info}\n{srt_styles_block}\n\n{ass_events_str}")

    if do_merge and merge_reports and merge_report_path:
        pd.DataFrame(merge_reports).to_excel(merge_report_path, index=False)

    return len(files)

def scan_all_styles_from_ass(filepath):
    """扫描 ASS 返回字典 {样式名: 原样式行内容}"""
    styles = {}
    try:
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            in_styles = False
            for line in f:
                line_s = line.strip()
                if line_s.startswith('[V4+ Styles]'): in_styles = True; continue
                if line_s.startswith('[Events]'): break
                if in_styles and line_s.startswith('Style:'):
                    name = line_s.split('Style:')[1].split(',')[0].strip()
                    styles[name] = line_s
    except: pass
    return styles

def process_ass_editor(input_dir, out_dir, mode_cfg):
    pass # Reserved, actual logic integrated into execute_ass_editor

def process_column_copy_batch(src_dir, tgt_dir, out_dir, err_rep, fmt, col_str):
    col_idx = int(col_str.split(':')[0])
    ext = '.srt' if fmt == 'SRT' else '.ass'
    tgt_files = [f for f in os.listdir(tgt_dir) if f.lower().endswith(ext)]
    if not tgt_files: raise ValueError(f"待接收数据的目标文件夹中没有 {ext} 文件！")
    
    all_errors = []
    processed_count = 0
    os.makedirs(out_dir, exist_ok=True)
    
    for file in tgt_files:
        src_file = os.path.join(src_dir, file)
        tgt_file = os.path.join(tgt_dir, file)
        out_file = os.path.join(out_dir, file)
    
        if not os.path.exists(src_file):
            all_errors.append({'文件名': file, '目标行号/时间轴': 'N/A', '目标文本': 'N/A', '错误说明': '提供数据的源文件夹中找不到同名文件'})
            with open(tgt_file, 'r', encoding='utf-8-sig') as f, open(out_file, 'w', encoding='utf-8') as out_f:
                out_f.write(f.read())
            continue

        if fmt == "SRT":
            src_blocks = parse_srt_file(src_file)
            tgt_blocks = parse_srt_file(tgt_file)

            if len(src_blocks) != len(tgt_blocks):
                all_errors.append({'文件名': file, '目标行号/时间轴': 'N/A', '目标文本': 'N/A', '错误说明': f'警告：行数不一致! 源:{len(src_blocks)}行, 目标:{len(tgt_blocks)}行'})

            out_blocks = []
            for i, t_b in enumerate(tgt_blocks):
                p = [t_b['ID'], t_b['Timeline'], t_b['Text']]
                if i < len(src_blocks):
                    s_b = src_blocks[i]
                    s_p = [s_b['ID'], s_b['Timeline'], s_b['Text']]

                    if t_b['Timeline'] != s_b['Timeline']:
                        all_errors.append({'文件名': file, '目标行号/时间轴': f"第{i+1}行 {t_b['Timeline']}", '目标文本': t_b['Text'], '错误说明': f"时间轴不一致! 源时间轴为 {s_b['Timeline']}"})
                    
                    p[col_idx] = s_p[col_idx] # 粗暴覆盖指定列
                else:
                    all_errors.append({'文件名': file, '目标行号/时间轴': f"第{i+1}行 {t_b['Timeline']}", '目标文本': t_b['Text'], '错误说明': '源文件在此行缺失，无法复制'})

                out_blocks.append(f"{p[0]}\n{p[1]}\n{p[2]}\n")

            with open(out_file, 'w', encoding='utf-8') as f: f.write("\n".join(out_blocks))
            processed_count += 1

        else: # ASS 格式
            with open(src_file, 'r', encoding='utf-8-sig') as f: src_lines = f.read().split('\n')
            with open(tgt_file, 'r', encoding='utf-8-sig') as f: tgt_lines = f.read().split('\n')

            src_diags = [l for l in src_lines if l.strip().startswith('Dialogue:')]
            out_lines = []
            diag_idx = 0
            
            tgt_diags_count = len([l for l in tgt_lines if l.strip().startswith('Dialogue:')])
            if len(src_diags) != tgt_diags_count:
                all_errors.append({'文件名': file, '目标行号/时间轴': 'N/A', '目标文本': 'N/A', '错误说明': f"Dialogue行数不一致! 源:{len(src_diags)}行, 目标:{tgt_diags_count}行"})

            for line in tgt_lines:
                if line.strip().startswith('Dialogue:'):
                    t_p = line.split(',', 9)
                    if len(t_p) >= 10:
                        if diag_idx < len(src_diags):
                            s_line = src_diags[diag_idx]
                            s_p = s_line.split(',', 9)
                            if len(s_p) >= 10:
                                t_time = f"{t_p[1]} --> {t_p[2]}"
                                s_time = f"{s_p[1]} --> {s_p[2]}"
                                if t_time != s_time:
                                    all_errors.append({'文件名': file, '目标行号/时间轴': f"第{diag_idx+1}条 {t_time}", '目标文本': t_p[9], '错误说明': f"时间轴不一致! 源时间轴为 {s_time}"})
                                
                                if col_idx < len(t_p) and col_idx < len(s_p):
                                    t_p[col_idx] = s_p[col_idx] # 粗暴覆盖指定列
                                    line = ",".join(t_p)
                        else:
                            all_errors.append({'文件名': file, '目标行号/时间轴': f"第{diag_idx+1}条", '目标文本': t_p[9] if len(t_p)>9 else 'N/A', '错误说明': '源文件在此行缺失，无法复制'})
                    out_lines.append(line)
                    diag_idx += 1
                else:
                    out_lines.append(line)

            with open(out_file, 'w', encoding='utf-8') as f: f.write("\n".join(out_lines))
            processed_count += 1

    if all_errors and err_rep: pd.DataFrame(all_errors).to_excel(err_rep, index=False)
    return processed_count, len(all_errors)

def process_srt_bilingual_split_batch(in_dir, out_dir, suffix1, suffix2):
    files = [f for f in os.listdir(in_dir) if f.lower().endswith('.srt')]
    if not files: raise ValueError("输入文件夹中没有找到 .srt 文件！")
    os.makedirs(out_dir, exist_ok=True)
    total_blocks = 0
    
    def char_profile(s):
        cjk = len(re.findall(r'[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af\u0e00-\u0e7f\u0400-\u04ff]', s))
        latin = len(re.findall(r'[a-zA-Z]', s))
        return 'C' if cjk > latin else 'L'

    for file in files:
        srt_file = os.path.join(in_dir, file)
        blocks = parse_srt_file(srt_file)
        base_name = os.path.splitext(file)[0]
        out_blocks1, out_blocks2 = [], []
        
        for block in blocks:
            lines = [l.strip() for l in block['Text'].split('\n') if l.strip()]
            text1, text2 = "", ""
            if len(lines) == 0: pass
            elif len(lines) == 1: text1, text2 = lines[0], ""
            elif len(lines) == 2: text1, text2 = lines[0], lines[1]
            elif len(lines) == 4: text1, text2 = "\n".join(lines[:2]), "\n".join(lines[2:])
            elif len(lines) == 3:
                p0, p1, p2 = char_profile(lines[0]), char_profile(lines[1]), char_profile(lines[2])
                if p1 == p0 and p1 != p2: text1, text2 = "\n".join(lines[:2]), lines[2]
                elif p1 == p2 and p1 != p0: text1, text2 = lines[0], "\n".join(lines[1:])
                else: text1, text2 = "\n".join(lines[:2]), lines[2]
            else:
                half = len(lines) // 2
                text1, text2 = "\n".join(lines[:half]), "\n".join(lines[half:])
                
            out_blocks1.append(f"{block['ID']}\n{block['Timeline']}\n{text1}\n")
            out_blocks2.append(f"{block['ID']}\n{block['Timeline']}\n{text2}\n")
            
        with open(os.path.join(out_dir, f"{base_name}_{suffix1}.srt"), 'w', encoding='utf-8') as f: f.write("\n".join(out_blocks1))
        with open(os.path.join(out_dir, f"{base_name}_{suffix2}.srt"), 'w', encoding='utf-8') as f: f.write("\n".join(out_blocks2))
        total_blocks += len(blocks)
        
    return len(files), total_blocks

def process_merge_srt_to_ass_batch(norm_dir, scr_dir, out_dir, custom_style_dict, style_mode, ref_cfg, regex_cfg=None):
    os.makedirs(out_dir, exist_ok=True)
    
    # --- 解析正则替换规则 ---
    replacements = []
    target_mode = ""
    if regex_cfg and regex_cfg.get('enable'):
        target_mode = regex_cfg.get('target', '画面字')
        for line in regex_cfg.get('text', '').split('\n'):
            if '>>>' in line:
                pat, repl = line.split('>>>', 1)
                # 自动将用户习惯的 $1, $2 转换为 Python 底层支持的 \1, \2
                repl_python = re.sub(r'\$(\d+)', r'\\\1', repl.strip())
                replacements.append((pat.strip(), repl_python))

    norm_files = {f for f in os.listdir(norm_dir) if f.lower().endswith('.srt')} if os.path.exists(norm_dir) else set()
    scr_files = {f for f in os.listdir(scr_dir) if f.lower().endswith('.srt')} if os.path.exists(scr_dir) else set()
    
    all_files = sorted(list(norm_files.union(scr_files)))
    if not all_files: raise ValueError("输入的文件夹中没有找到任何 .srt 文件！")
    
    if style_mode == 0:
        resx = custom_style_dict.get('play_res_x', '1080')
        resy = custom_style_dict.get('play_res_y', '1920')
    else:
        resx, resy = get_ass_resolution(ref_cfg['ref_path'])
        
    srt_script_info = f"[Script Info]\nScriptType: v4.00+\nPlayResX: {resx}\nPlayResY: {resy}\n"
    srt_styles_block = "[V4+ Styles]\nFormat: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding\n"
    
    if style_mode == 0:
        d = custom_style_dict
        srt_styles_block += build_ass_style_line("对白字幕", d['n_font'], d['n_size'], d['n_color'], d['n_out_color'], d['n_margin_v'], d['n_margin_lr'], d['n_outline'], d.get('n_align','2'), d.get('n_shadow','0'), d.get('n_bold',0), d.get('n_italic',0), d.get('n_alpha','00'), d.get('n_out_alpha','00')) + "\n"
        srt_styles_block += build_ass_style_line("画面字", d['s_font'], d['s_size'], d['s_color'], d['s_out_color'], d['s_margin_v'], d['s_margin_lr'], d['s_outline'], d.get('s_align','8'), d.get('s_shadow','0'), d.get('s_bold',0), d.get('s_italic',0), d.get('s_alpha','00'), d.get('s_out_alpha','00'))
    else:
        ref_styles = scan_all_styles_from_ass(ref_cfg['ref_path'])
        n_line = rename_style_line(ref_styles.get(ref_cfg['n_style'], build_ass_style_line("对白字幕", "Arial", "60", "&H00FFFFFF", "&H00000000", "20", "20", "2")), "对白字幕")
        s_line = rename_style_line(ref_styles.get(ref_cfg['s_style'], build_ass_style_line("画面字", "Arial", "60", "&H00FFFFFF", "&H00000000", "850", "20", "2")), "画面字")
        
        if ref_cfg['font_mode'] == 1:
            n_line = replace_font_in_style(n_line, ref_cfg['override_font'])
            s_line = replace_font_in_style(s_line, ref_cfg['override_font'])
            
        srt_styles_block += n_line + "\n" + s_line
    
    for file in all_files:
        screen_events, normal_events = [], []
        if file in scr_files:
            blocks = parse_srt_file(os.path.join(scr_dir, file))
            for block in blocks:
                s_text = clean_ass_text(block['Text'])
                if not s_text: continue
                
                # === 应用正则到画面字 ===
                if replacements and target_mode in ['画面字', '全部']:
                    for pat, repl in replacements:
                        s_text = re.sub(pat, repl, s_text)
                if not s_text.strip(): continue # 如果替换后变成了空行，则跳过
                
                start_ass = srt_to_ass_time(block['Timeline'].split(' --> ')[0])
                end_ass = srt_to_ass_time(block['Timeline'].split(' --> ')[1])
                screen_events.append(f"Dialogue: 0,{start_ass},{end_ass},画面字,,0,0,0,,{s_text}")
                
        if file in norm_files:
            blocks = parse_srt_file(os.path.join(norm_dir, file))
            for block in blocks:
                n_text = clean_ass_text(block['Text'])
                if not n_text: continue
                
                # === 应用正则到对白字幕 ===
                if replacements and target_mode in ['对白字幕', '全部']:
                    for pat, repl in replacements:
                        n_text = re.sub(pat, repl, n_text)
                if not n_text.strip(): continue
                
                start_ass = srt_to_ass_time(block['Timeline'].split(' --> ')[0])
                end_ass = srt_to_ass_time(block['Timeline'].split(' --> ')[1])
                normal_events.append(f"Dialogue: 0,{start_ass},{end_ass},对白字幕,,0,0,0,,{n_text}")

        ass_events_str = "[Events]\nFormat: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text\n"
        ass_events_str += "\n".join(screen_events) + ("\n" if screen_events else "")
        ass_events_str += "\n".join(normal_events) + "\n"
        
        out_path = os.path.join(out_dir, file.rsplit('.', 1)[0] + '.ass')
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(f"{srt_script_info}\n{srt_styles_block}\n\n{ass_events_str}")
            
    return len(all_files)

def process_ass_split(in_dir, out_scr_dir, out_norm_dir, use_c1, bracket_str, use_c2, sel_effs, use_c3, sel_styles, to_srt=False):
    """根据组合条件将 ASS 拆分为画面字和普通字两个文件"""
    files = [f for f in os.listdir(in_dir) if f.lower().endswith('.ass')]
    if not files: raise ValueError("输入文件夹中没有找到 .ass 文件！")
    
    os.makedirs(out_scr_dir, exist_ok=True)
    os.makedirs(out_norm_dir, exist_ok=True)
    
    l_b, r_b = "", ""
    if use_c1 and len(bracket_str) >= 2:
        half = len(bracket_str) // 2
        l_b, r_b = bracket_str[:half], bracket_str[half:]
        
    def ass_to_srt_time(ass_time):
        h, m, s_ms = ass_time.strip().split(':')
        s, ms = s_ms.split('.')
        return f"{int(h):02d}:{int(m):02d}:{int(s):02d},{ms.ljust(3, '0')}"

    def convert_to_srt_blocks(ev_list):
        blocks = []
        idx = 1
        for ev in ev_list:
            if ev.startswith('Dialogue:'):
                parts = ev.split(',', 9)
                if len(parts) >= 10:
                    st = ass_to_srt_time(parts[1])
                    ed = ass_to_srt_time(parts[2])
                    txt = re.sub(r'\{.*?\}', '', parts[9]).replace('\\N', '\n').replace('\\n', '\n')
                    blocks.append(f"{idx}\n{st} --> {ed}\n{txt}\n")
                    idx += 1
        return "\n".join(blocks)
        
    processed_count = 0
    for file in files:
        filepath = os.path.join(in_dir, file)
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            lines = f.read().split('\n')
            
        h_lines, s_lines, ev_lines = [], [], []
        curr = "info"
        for line in lines:
            l = line.strip()
            if l.startswith('[V4+ Styles]'): curr = "styles"
            elif l.startswith('[Events]'): curr = "events"
            
            if curr == "info": h_lines.append(line)
            elif curr == "styles": s_lines.append(line)
            elif curr == "events": ev_lines.append(line)
            
        screen_ev, normal_ev = [], []
        for ev in ev_lines:
            if ev.startswith('Dialogue:'):
                parts = ev.split(',', 9)
                if len(parts) >= 10:
                    style = parts[3].strip()
                    effect = parts[8].strip()
                    txt = parts[9]
                    c_txt = re.sub(r'\{.*?\}', '', txt).strip()
                    
                    is_screen = True
                    if use_c1:
                        if not (l_b and r_b and c_txt.startswith(l_b) and c_txt.endswith(r_b)): is_screen = False
                    if use_c2:
                        if effect not in sel_effs: is_screen = False
                    if use_c3:
                        if style not in sel_styles: is_screen = False
                            
                    if is_screen: screen_ev.append(ev)
                    else: normal_ev.append(ev)
                else:
                    screen_ev.append(ev)
                    normal_ev.append(ev)
            else:
                screen_ev.append(ev)
                normal_ev.append(ev)
                
        base_name = os.path.splitext(file)[0]
        if to_srt:
            with open(os.path.join(out_scr_dir, base_name + '.srt'), 'w', encoding='utf-8') as f:
                f.write(convert_to_srt_blocks(screen_ev))
            with open(os.path.join(out_norm_dir, base_name + '.srt'), 'w', encoding='utf-8') as f:
                f.write(convert_to_srt_blocks(normal_ev))
        else:
            with open(os.path.join(out_scr_dir, file), 'w', encoding='utf-8') as f:
                f.write("\n".join(h_lines) + "\n" + "\n".join(s_lines) + "\n" + "\n".join(screen_ev) + "\n")
            with open(os.path.join(out_norm_dir, file), 'w', encoding='utf-8') as f:
                f.write("\n".join(h_lines) + "\n" + "\n".join(s_lines) + "\n" + "\n".join(normal_ev) + "\n")
            
        processed_count += 1
    return processed_count

# ================= UI 交互回调 =================

def run_ass_merge():
    dir1 = am_dir1_var.get().strip()
    dir2 = am_dir2_var.get().strip()
    out_dir = am_out_var.get().strip()
    
    if not dir1 or not dir2 or not out_dir:
        return messagebox.showwarning("警告", "请完整选择两个输入文件夹和输出文件夹！")
        
    try:
        count = process_ass_merge(dir1, dir2, out_dir)
        messagebox.showinfo("完成", f"ASS 合并成功！\n共合并了 {count} 个文件。")
    except Exception as e:
        messagebox.showerror("错误", f"合并失败:\n{str(e)}")

def run_time_split():
    in_dir = ts_in_var.get().strip()
    out_norm = ts_norm_var.get().strip()
    out_scr = ts_scr_var.get().strip()
    mode = ts_mode_var.get()
    
    if not in_dir or not out_norm or not out_scr:
        return messagebox.showwarning("警告", "请完整填写输入和输出文件夹！")
        
    try:
        count = process_time_split(in_dir, out_norm, out_scr, mode)
        messagebox.showinfo("完成", f"按时间轴拆分成功！\n共自动处理 {count} 个文件。")
    except Exception as e:
        messagebox.showerror("错误", f"拆分失败:\n{str(e)}")

def run_split():
    infile = split_in_var.get().strip()
    out_src = split_out_src_var.get().strip()
    out_tgt = split_out_tgt_var.get().strip()
    
    parts = [p.strip() for p in split_cols_var.get().strip().replace('，', ',').split(',') if p.strip()]
    if not infile or not out_src: return messagebox.showwarning("警告", "请至少填写输入文件和源语言输出目录！")
    
    if split_mode_var.get() == 1:
        if len(parts) < 4: return messagebox.showwarning("警告", "模式1 需要至少 4 个列名字母 (文件, ID, 时间, 内容)！")
        try:
            count = process_split(infile, out_src, out_tgt, parts)
            messagebox.showinfo("完成", f"成功拆分 {count} 个视频文件！")
        except Exception as e: messagebox.showerror("错误", f"拆分失败:\n{str(e)}")
    else:
        if len(parts) < 3: return messagebox.showwarning("警告", "模式2 需要至少 3 个列名字母 (文件, 时间, 内容)！")
        try:
            count = process_split_mode2(infile, out_src, out_tgt, parts, split_sheet_var.get().strip())
            messagebox.showinfo("完成", f"成功拆分 {count} 个视频文件！")
        except Exception as e: messagebox.showerror("错误", f"拆分失败:\n{str(e)}")

def run_merge():
    src_dir, tgt_dir, out_xls = merge_src_var.get().strip(), merge_tgt_var.get().strip(), merge_out_var.get().strip()
    src_name, tgt_name = merge_src_name_var.get().strip(), merge_tgt_name_var.get().strip()
    
    if not src_name: return messagebox.showerror("错误", "请输入源语言列名！")
    if not src_dir or not out_xls: return messagebox.showwarning("警告", "请完整填写源目录和输出文件路径！")
    
    if merge_mode_var.get() == 1:
        if not tgt_dir or not tgt_name: return messagebox.showwarning("警告", "模式1(双语对照)下，目标语言目录和列名必填！")
        try:
            count = process_merge(src_dir, tgt_dir, src_name, tgt_name, out_xls)
            messagebox.showinfo("完成", f"成功合并 {count} 个 SRT 文件到 Excel！")
        except Exception as e: messagebox.showerror("错误", f"合并失败:\n{str(e)}")
    else:
        try:
            count = process_merge_mode2(src_dir, tgt_dir, src_name, tgt_name, out_xls)
            messagebox.showinfo("完成", f"成功按模式2合并了 {count} 个 SRT 文件到 Excel！")
        except Exception as e: messagebox.showerror("错误", f"合并失败:\n{str(e)}")

def run_replace():
    report_file, srt_dir, out_summary = rep_report_var.get().strip(), rep_srt_var.get().strip(), rep_out_var.get().strip()
    parts = [p.strip() for p in rep_cols_var.get().strip().replace('，', ',').split(',')]
    if not report_file or not srt_dir or not out_summary: return messagebox.showwarning("警告", "请完整选择文件路径！")
    if len(parts) != 3 or not all(parts): return messagebox.showwarning("警告", "列名格式错误！")
    try:
        rep_count, file_count = process_replace(report_file, srt_dir, out_summary, parts[0], parts[1], parts[2])
        messagebox.showinfo("完成", f"替换完毕！\n共影响了 {file_count} 个 SRT 文件。\n合计成功替换 {rep_count} 条字幕，已生成展示表格。")
    except Exception as e: messagebox.showerror("错误", f"替换失败:\n{str(e)}")

def run_zip():
    target_dir, out_dir = zip_target_var.get().strip(), zip_out_var.get().strip()
    if not target_dir or not out_dir: return messagebox.showwarning("警告", "请选择需要打包的文件夹和输出文件夹！")
    try: max_f = int(zip_max_var.get().strip())
    except: return messagebox.showwarning("警告", "最大文件数必须是整数！")
    try:
        total_files, num_zips = process_zip(target_dir, out_dir, max_f)
        messagebox.showinfo("完成", f"打包完成！\n共处理了 {total_files} 个文件，\n生成了 {num_zips} 个分卷压缩包。")
    except Exception as e: messagebox.showerror("错误", f"打包失败:\n{str(e)}")

def run_ass_convert():
    srt_dir, out_dir = ass_srt_var.get().strip(), ass_out_var.get().strip()
    if not srt_dir or not out_dir: return messagebox.showwarning("警告", "请完整选择仅限 SRT 输入文件夹和 ASS 输出文件夹！")
    
    bracket = ass_bracket_var.get().strip()
    regex_text = ass_regex_text.get("1.0", tk.END)
    do_merge = ass_merge_var.get() == 1
    merge_report_path = ass_merge_report_var.get().strip()
    if do_merge and not merge_report_path: return messagebox.showwarning("警告", "勾选了合并功能，请指定合并报告的保存路径！")
    
    custom_style = {
        "play_res_x": ass_resx_var.get(), "play_res_y": ass_resy_var.get(),
        "n_font": ass_n_font_var.get(), "n_size": ass_n_size_var.get(), "n_color": ass_n_color_var.get(), "n_alpha": ass_n_alpha_var.get(), "n_out_color": ass_n_outcolor_var.get(), "n_out_alpha": ass_n_outalpha_var.get(),
        "n_margin_v": ass_n_marginv_var.get(), "n_margin_lr": ass_n_marginlr_var.get(), "n_outline": ass_n_outline_var.get(),
        "n_align": ass_n_align_var.get(), "n_shadow": ass_n_shadow_var.get(), "n_bold": ass_n_bold_var.get(), "n_italic": ass_n_italic_var.get(),
        "s_font": ass_s_font_var.get(), "s_size": ass_s_size_var.get(), "s_color": ass_s_color_var.get(), "s_alpha": ass_s_alpha_var.get(), "s_out_color": ass_s_outcolor_var.get(), "s_out_alpha": ass_s_outalpha_var.get(),
        "s_margin_v": ass_s_marginv_var.get(), "s_margin_lr": ass_s_marginlr_var.get(), "s_outline": ass_s_outline_var.get(),
        "s_align": ass_s_align_var.get(), "s_shadow": ass_s_shadow_var.get(), "s_bold": ass_s_bold_var.get(), "s_italic": ass_s_italic_var.get()
    }
    
    ref_cfg = None
    if ass_style_mode_5.get() == 1:
        ref_cfg = {
            'ref_path': ass_ref_path_5.get().strip(),
            'n_style': ass_ref_n_var_5.get().strip(),
            's_style': ass_ref_s_var_5.get().strip(),
            'font_mode': ass_ref_font_mode_5.get(),
            'override_font': ass_ref_override_font_5.get().strip()
        }
        if not ref_cfg['ref_path'] or not os.path.exists(ref_cfg['ref_path']):
            return messagebox.showwarning("警告", "请选择有效的外部参考 ASS 文件！")
        if not ref_cfg['n_style'] or not ref_cfg['s_style']:
            return messagebox.showwarning("警告", "请在应用前先扫描参考文件并选择样式！")
            
    try:
        count = process_srt_to_ass(srt_dir, out_dir, bracket, regex_text, custom_style, do_merge, merge_report_path, ass_style_mode_5.get(), ref_cfg)
        messagebox.showinfo("完成", f"转换完成！\n成功处理了 {count} 个 SRT 文件，已全部转为 ASS 并应用样式。")
    except Exception as e: messagebox.showerror("错误", f"转换失败:\n{str(e)}")

def run_column_copy():
    src_dir, tgt_dir = eff_src_var.get().strip(), eff_tgt_var.get().strip()
    out_dir, err_rep = eff_out_var.get().strip(), eff_err_var.get().strip()
    fmt, col_str = eff_fmt_var.get(), eff_col_var.get()
    
    if not src_dir or not tgt_dir or not out_dir: return messagebox.showwarning("警告", "请完整选择目录！")
    try:
        processed_count, err_count = process_column_copy_batch(src_dir, tgt_dir, out_dir, err_rep, fmt, col_str)
        if err_count > 0: messagebox.showwarning("部分完成", f"成功提取 {processed_count} 个文件！但检测到 {err_count} 处异常(行数缺失或时间轴错位)，已忽略时间轴强行复用并导出报告。")
        else: messagebox.showinfo("完成", f"完美处理 {processed_count} 个文件！所选列已严格按行数映射全部复制成功。")
    except Exception as e: messagebox.showerror("错误", f"处理失败:\n{str(e)}")

def run_srt_bilingual_split():
    in_d, out_d = bi_srt_var.get().strip(), bi_out_dir_var.get().strip()
    s1, s2 = bi_suf1_var.get().strip(), bi_suf2_var.get().strip()
    if not in_d or not out_d: return messagebox.showwarning("警告", "请选择输入和输出目录！")
    if not s1 or not s2: return messagebox.showwarning("警告", "请输入拆分后的语言后缀！")
    try:
        file_count, block_count = process_srt_bilingual_split_batch(in_d, out_d, s1, s2)
        messagebox.showinfo("完成", f"批量拆分成功！\n共处理 {file_count} 个文件（合计 {block_count} 条字幕），已导出单语 SRT。")
    except Exception as e: messagebox.showerror("错误", f"拆分失败:\n{str(e)}")

def run_merge_srt_to_ass():
    norm_d, scr_d, out_d = ms_norm_var.get().strip(), ms_scr_var.get().strip(), ms_out_var.get().strip()
    if not out_d or (not norm_d and not scr_d): return messagebox.showwarning("警告", "请至少选择一个输入和输出文件夹！")
    
    custom_style = {
        "play_res_x": m5_resx_var.get(), "play_res_y": m5_resy_var.get(),
        "n_font": m5_n_font_var.get(), "n_size": m5_n_size_var.get(), "n_color": m5_n_color_var.get(), "n_alpha": m5_n_alpha_var.get(), "n_out_color": m5_n_outcolor_var.get(), "n_out_alpha": m5_n_outalpha_var.get(),
        "n_margin_v": m5_n_marginv_var.get(), "n_margin_lr": m5_n_marginlr_var.get(), "n_outline": m5_n_outline_var.get(),
        "n_align": m5_n_align_var.get(), "n_shadow": m5_n_shadow_var.get(), "n_bold": m5_n_bold_var.get(), "n_italic": m5_n_italic_var.get(),
        "s_font": m5_s_font_var.get(), "s_size": m5_s_size_var.get(), "s_color": m5_s_color_var.get(), "s_alpha": m5_s_alpha_var.get(), "s_out_color": m5_s_outcolor_var.get(), "s_out_alpha": m5_s_outalpha_var.get(),
        "s_margin_v": m5_s_marginv_var.get(), "s_margin_lr": m5_s_marginlr_var.get(), "s_outline": m5_s_outline_var.get(),
        "s_align": m5_s_align_var.get(), "s_shadow": m5_s_shadow_var.get(), "s_bold": m5_s_bold_var.get(), "s_italic": m5_s_italic_var.get()
    }
    
    ref_cfg = None
    if ms_style_mode_9.get() == 1:
        ref_cfg = {
            'ref_path': ms_ref_path_9.get().strip(),
            'n_style': ms_ref_n_var_9.get().strip(),
            's_style': ms_ref_s_var_9.get().strip(),
            'font_mode': ms_ref_font_mode_9.get(),
            'override_font': ms_ref_override_font_9.get().strip()
        }
        if not ref_cfg['ref_path'] or not os.path.exists(ref_cfg['ref_path']):
            return messagebox.showwarning("警告", "请选择有效的外部参考 ASS 文件！")
        if not ref_cfg['n_style'] or not ref_cfg['s_style']:
            return messagebox.showwarning("警告", "请在应用前先扫描参考文件并选择样式！")
            
    regex_cfg = {
        'enable': ms_enable_regex_var.get() == 1,
        'target': ms_regex_target_var.get().strip(),
        'text': ms_regex_text.get("1.0", tk.END)
    }
            
    try:
        count = process_merge_srt_to_ass_batch(norm_d, scr_d, out_d, custom_style, ms_style_mode_9.get(), ref_cfg, regex_cfg)
        messagebox.showinfo("完成", f"合并转换完成！\n成功将 {count} 个双源 SRT 合并为了 ASS。")
    except Exception as e: messagebox.showerror("错误", f"处理失败:\n{str(e)}")

def scan_ass_split_features():
    d = split_ass_in_var.get().strip()
    if not d or not os.path.exists(d): return messagebox.showwarning("提示", "请先选择输入文件夹！")
    ass_files = [os.path.join(d, f) for f in os.listdir(d) if f.lower().endswith('.ass')]
    if not ass_files: return messagebox.showwarning("提示", "输入文件夹中未找到 .ass 文件！")
    
    effs, styles = set(), set()
    for filepath in ass_files:
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            for line in f:
                if line.startswith('Dialogue:'):
                    p = line.split(',', 9)
                    if len(p) >= 10:
                        styles.add(p[3].strip())
                        effs.add(p[8].strip())
                        
    lb_split_effs.delete(0, tk.END)
    for e in sorted(list(effs)): lb_split_effs.insert(tk.END, e)
    
    lb_split_styles.delete(0, tk.END)
    for s in sorted(list(styles)): lb_split_styles.insert(tk.END, s)
    
    messagebox.showinfo("成功", f"扫描完毕！\n共发现 {len(effs)} 种特效说明，{len(styles)} 种样式。")

def run_ass_split():
    i_d = split_ass_in_var.get().strip()
    o_s = split_ass_out_scr_var.get().strip()
    o_n = split_ass_out_norm_var.get().strip()
    
    if not i_d or not o_s or not o_n: return messagebox.showwarning("警告", "请完整选择输入和两个输出文件夹！")
        
    u1, u2, u3 = split_ass_c1_var.get(), split_ass_c2_var.get(), split_ass_c3_var.get()
    if not (u1 or u2 or u3): return messagebox.showwarning("警告", "请至少勾选一个拆分条件！")
        
    b_str = split_ass_bracket_var.get().strip()
    sel_effs = [lb_split_effs.get(i) for i in lb_split_effs.curselection()]
    sel_styles = [lb_split_styles.get(i) for i in lb_split_styles.curselection()]
    to_srt = split_ass_to_srt_var.get() == 1  # 接收是否转 SRT
    
    if u2 and not sel_effs: return messagebox.showwarning("警告", "勾选了特效条件，但未在列表中选中任何特效！")
    if u3 and not sel_styles: return messagebox.showwarning("警告", "勾选了样式条件，但未在列表中选中任何样式！")
    
    try:
        count = process_ass_split(i_d, o_s, o_n, u1, b_str, u2, sel_effs, u3, sel_styles, to_srt)
        messagebox.showinfo("完成", f"拆分成功！\n共处理了 {count} 个 ASS 文件，已分别输出。")
    except Exception as e: messagebox.showerror("错误", f"拆分失败:\n{str(e)}")

# ----- 预设通用操作 -----
def action_save_preset(var, current_list, combobox, filepath, min_len):
    val = var.get().strip().replace('，', ',')
    parts = [p.strip().upper() for p in val.split(',') if p.strip()]
    if len(parts) < min_len: return messagebox.showwarning("格式不规范", f"请输入标准的 {min_len} 个字母格式后再保存")
    val_clean = ", ".join(parts)
    if val_clean not in current_list:
        current_list.append(val_clean)
        combobox['values'] = current_list
        var.set(val_clean)
        save_presets_to_file(filepath, current_list)
        messagebox.showinfo("提示", "预设保存成功！")
    else: messagebox.showinfo("提示", "该预设已存在！")

def action_del_preset(var, current_list, combobox, filepath):
    val = var.get().strip()
    if val in current_list:
        current_list.remove(val)
        combobox['values'] = current_list
        var.set(current_list[0] if current_list else "")
        save_presets_to_file(filepath, current_list)
        messagebox.showinfo("提示", "预设已删除！")
    else: messagebox.showwarning("提示", "列表中没有此预设，无法删除。")

def choose_color(var, btn):
    # 动态获取当前触发按钮所在的顶层窗口
    parent_win = btn.winfo_toplevel()
    
    # 强制将颜色面板绑定到该子窗口上
    c = colorchooser.askcolor(title="选择颜色", initialcolor=var.get() or "#FFFFFF", parent=parent_win)
    if c[1]:
        var.set(c[1].upper())
        btn.config(bg=c[1].upper())
        
    # 选完颜色后，强制把这个子窗口拉回最顶层并恢复输入焦点
    parent_win.lift()
    parent_win.focus_set()

def update_color_btn(btn, color_hex):
    try: btn.config(bg=color_hex)
    except: pass

def ask_file(var, title, filetypes): var.set(filedialog.askopenfilename(title=title, filetypes=filetypes))
def ask_dir(var, title): var.set(filedialog.askdirectory(title=title))
def ask_save_file(var, title, filetypes, defaultextension): var.set(filedialog.asksaveasfilename(title=title, filetypes=filetypes, defaultextension=defaultextension))

# ================= GUI 界面构建 =================

root = tk.Tk()
root.title("字幕拆分与合并工具箱")
root.geometry("770x680")
root.minsize(500, 500)

# 跨平台主题与字体自适应
os_name = platform.system()
if os_name == 'Darwin':  # macOS 环境
    GLOBAL_FONT = 'PingFang SC'
    GLOBAL_FONT_TUPLE = (GLOBAL_FONT, 12) # Mac 渲染策略不同，字体需稍大
else:  # Windows / Linux 环境
    GLOBAL_FONT = 'Microsoft YaHei'
    GLOBAL_FONT_TUPLE = (GLOBAL_FONT, 10)

style = ttk.Style(root)
themes = style.theme_names()
if os_name == 'Darwin' and 'aqua' in themes:
    style.theme_use('aqua') # 激活原生 macOS 拟物化水晶主题
elif 'vista' in themes: 
    style.theme_use('vista')
elif 'winnative' in themes: 
    style.theme_use('winnative')

style.configure('TButton', font=GLOBAL_FONT_TUPLE)

# --- 新增：顶部导航分类栏 ---
nav_frame = tk.Frame(root)
nav_frame.pack(fill=tk.X, padx=10, pady=(10, 0))

category_var = tk.StringVar(value="SRT")

def switch_category():
    cat = category_var.get()
    nb_srt.pack_forget()
    nb_ass.pack_forget()
    nb_other.pack_forget()
    
    if cat == "SRT": nb_srt.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    elif cat == "ASS": nb_ass.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    else: nb_other.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

ttk.Radiobutton(nav_frame, text=" 📝 基本功能 ", variable=category_var, value="SRT", command=switch_category, style='Toolbutton').pack(side=tk.LEFT, padx=(0, 5), ipadx=10, ipady=3)
ttk.Radiobutton(nav_frame, text=" 🎬 高级功能 ", variable=category_var, value="ASS", command=switch_category, style='Toolbutton').pack(side=tk.LEFT, padx=5, ipadx=10, ipady=3)
ttk.Radiobutton(nav_frame, text=" 🛠️ 其他功能 ", variable=category_var, value="OTHER", command=switch_category, style='Toolbutton').pack(side=tk.LEFT, padx=5, ipadx=10, ipady=3)

nb_srt = ttk.Notebook(root)
nb_ass = ttk.Notebook(root)
nb_other = ttk.Notebook(root)

nb_srt.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
# ================= TAB 1: 拆分 =================
tab_split = ttk.Frame(nb_srt, padding=20)
nb_srt.add(tab_split, text=" XLSX 拆分为 SRT ")
tab_split.columnconfigure(1, weight=1)
current_presets_split = load_presets(PRESET_FILE_SPLIT, DEFAULT_PRESETS_SPLIT)
split_in_var = tk.StringVar()
split_out_src_var, split_out_tgt_var = tk.StringVar(), tk.StringVar()
split_cols_var = tk.StringVar(value=current_presets_split[0] if current_presets_split else "A, B, C, D")

split_mode_var = tk.IntVar(value=1)
split_sheet_var = tk.StringVar()

f_mode = ttk.Frame(tab_split)
f_mode.grid(row=0, column=0, columnspan=3, sticky="w", pady=5)
ttk.Radiobutton(f_mode, text="模式1: 番茄导出格式", variable=split_mode_var, value=1, command=lambda: update_split_mode()).pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(f_mode, text="模式2: CPP格式", variable=split_mode_var, value=2, command=lambda: update_split_mode()).pack(side=tk.LEFT, padx=5)

def scan_sheets():
    f = split_in_var.get().strip()
    if not f or not os.path.exists(f): return messagebox.showwarning("警告", "请先选择有效的Excel文件！")
    try:
        xl = pd.ExcelFile(f)
        cb_sheet['values'] = xl.sheet_names
        if xl.sheet_names: split_sheet_var.set(xl.sheet_names[0])
        messagebox.showinfo("成功", f"扫描到 {len(xl.sheet_names)} 个 Sheet")
    except Exception as e: messagebox.showerror("错误", str(e))

btn_scan_sheet = ttk.Button(f_mode, text="扫描Sheet", command=scan_sheets, state="disabled")
btn_scan_sheet.pack(side=tk.LEFT, padx=5)
cb_sheet = ttk.Combobox(f_mode, textvariable=split_sheet_var, state="disabled", width=15)
cb_sheet.pack(side=tk.LEFT, padx=5)

ttk.Label(tab_split, text="双语 XLSX/CSV 文件:").grid(row=1, column=0, sticky="e", padx=(0,10), pady=10)
ttk.Entry(tab_split, textvariable=split_in_var).grid(row=1, column=1, sticky="ew", padx=5, pady=10)
ttk.Button(tab_split, text="浏览...", command=lambda: ask_file(split_in_var, "选择文件", [("Excel", "*.xlsx"), ("CSV", "*.csv")])).grid(row=1, column=2, padx=(5,0), pady=10)

f_s = ttk.Frame(tab_split)
f_s.grid(row=2, column=0, columnspan=3, sticky="w", pady=5)
lbl_split_hint = ttk.Label(f_s, text="列名 (模式1: 文件, ID, 时间, 内容):")
lbl_split_hint.pack(side=tk.LEFT, padx=(0,5))
cb_s = ttk.Combobox(f_s, textvariable=split_cols_var, values=current_presets_split, width=15)
cb_s.pack(side=tk.LEFT, padx=(0, 10))
ttk.Button(f_s, text="保存预设", command=lambda: action_save_preset(split_cols_var, current_presets_split, cb_s, PRESET_FILE_SPLIT, 3)).pack(side=tk.LEFT, padx=5)
ttk.Button(f_s, text="删除预设", command=lambda: action_del_preset(split_cols_var, current_presets_split, cb_s, PRESET_FILE_SPLIT)).pack(side=tk.LEFT, padx=5)

def update_split_mode():
    if split_mode_var.get() == 1:
        cb_sheet.config(state="disabled")
        btn_scan_sheet.config(state="disabled")
        lbl_split_hint.config(text="列名 (模式1: 文件名, ID, 时间, 内容1, (内容2)):")
        split_cols_var.set("A, B, C, D")
    else:
        cb_sheet.config(state="readonly")
        btn_scan_sheet.config(state="normal")
        lbl_split_hint.config(text="列名 (模式2: 文件名, 时间, 内容1, (内容2)):")
        split_cols_var.set("A, B, C, D")

ttk.Label(tab_split, text="源语言 SRT 输出目录:").grid(row=3, column=0, sticky="e", padx=(0,10), pady=10)
ttk.Entry(tab_split, textvariable=split_out_src_var).grid(row=3, column=1, sticky="ew", padx=5, pady=10)
ttk.Button(tab_split, text="浏览...", command=lambda: ask_dir(split_out_src_var, "选择目录")).grid(row=3, column=2, padx=(5,0), pady=10)

ttk.Label(tab_split, text="目标语言 SRT 输出目录(可选):").grid(row=4, column=0, sticky="e", padx=(0,10), pady=10)
ttk.Entry(tab_split, textvariable=split_out_tgt_var).grid(row=4, column=1, sticky="ew", padx=5, pady=10)
ttk.Button(tab_split, text="浏览...", command=lambda: ask_dir(split_out_tgt_var, "选择目录")).grid(row=4, column=2, padx=(5,0), pady=10)

ttk.Button(tab_split, text="开始拆分", command=run_split, style='TButton').grid(row=5, column=0, columnspan=3, pady=20, ipadx=20, ipady=5)

# ================= TAB 2: 合并 =================
tab_merge = ttk.Frame(nb_srt, padding=20)
nb_srt.add(tab_merge, text=" SRT 合并为 XLSX ")
tab_merge.columnconfigure(1, weight=1)

merge_mode_var = tk.IntVar(value=1)
merge_src_var, merge_tgt_var, merge_out_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
merge_src_name_var, merge_tgt_name_var = tk.StringVar(), tk.StringVar()

f_merge_mode = ttk.Frame(tab_merge)
f_merge_mode.grid(row=0, column=0, columnspan=3, sticky="w", pady=5)
ttk.Radiobutton(f_merge_mode, text="模式1: 番茄导出格式", variable=merge_mode_var, value=1).pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(f_merge_mode, text="模式2: CPP格式", variable=merge_mode_var, value=2).pack(side=tk.LEFT, padx=5)

ttk.Label(tab_merge, text="源语言 SRT 目录:").grid(row=1, column=0, sticky="e", padx=(0,10), pady=10)
ttk.Entry(tab_merge, textvariable=merge_src_var).grid(row=1, column=1, sticky="ew", padx=5, pady=10)
ttk.Button(tab_merge, text="浏览...", command=lambda: ask_dir(merge_src_var, "选择目录")).grid(row=1, column=2, padx=(5,0), pady=10)

ttk.Label(tab_merge, text="目标语言 SRT 目录(可选):").grid(row=2, column=0, sticky="e", padx=(0,10), pady=10)
ttk.Entry(tab_merge, textvariable=merge_tgt_var).grid(row=2, column=1, sticky="ew", padx=5, pady=10)
ttk.Button(tab_merge, text="浏览...", command=lambda: ask_dir(merge_tgt_var, "选择目录")).grid(row=2, column=2, padx=(5,0), pady=10)

f_m = ttk.Frame(tab_merge)
f_m.grid(row=3, column=1, sticky="w", pady=5)
ttk.Label(f_m, text="源列名:").pack(side=tk.LEFT)
ttk.Entry(f_m, textvariable=merge_src_name_var, width=10).pack(side=tk.LEFT, padx=(5,15))
ttk.Label(f_m, text="目标列名:").pack(side=tk.LEFT)
ttk.Entry(f_m, textvariable=merge_tgt_name_var, width=10).pack(side=tk.LEFT, padx=(5,0))

ttk.Label(tab_merge, text="保存为 XLSX 文件:").grid(row=4, column=0, sticky="e", padx=(0,10), pady=10)
ttk.Entry(tab_merge, textvariable=merge_out_var).grid(row=4, column=1, sticky="ew", padx=5, pady=10)
ttk.Button(tab_merge, text="浏览...", command=lambda: ask_save_file(merge_out_var, "保存", [("Excel", "*.xlsx")], ".xlsx")).grid(row=4, column=2, padx=(5,0), pady=10)
ttk.Button(tab_merge, text="开始合并", command=run_merge, style='TButton').grid(row=5, column=0, columnspan=3, pady=20, ipadx=20, ipady=5)

# ================= TAB 3: 替换 =================
tab_rep = ttk.Frame(nb_srt, padding=20)
nb_srt.add(tab_rep, text=" 根据报告修改 SRT ")
tab_rep.columnconfigure(1, weight=1)
current_presets_rep = load_presets(PRESET_FILE_REP, DEFAULT_PRESETS_REP)
rep_report_var, rep_srt_var, rep_out_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
rep_cols_var = tk.StringVar(value=current_presets_rep[0] if current_presets_rep else "A, B, E")

ttk.Label(tab_rep, text="QA 报告 (Excel/CSV):").grid(row=0, column=0, sticky="e", padx=(0,10), pady=10)
ttk.Entry(tab_rep, textvariable=rep_report_var).grid(row=0, column=1, sticky="ew", padx=5, pady=10)
ttk.Button(tab_rep, text="浏览...", command=lambda: ask_file(rep_report_var, "选择文件", [("Excel", "*.xlsx"), ("CSV", "*.csv")])).grid(row=0, column=2, padx=(5,0), pady=10)

f_r = ttk.Frame(tab_rep)
f_r.grid(row=1, column=0, columnspan=3, sticky="w", pady=5, padx=20)
ttk.Label(f_r, text="列名 (文件, ID, 内容):").pack(side=tk.LEFT, padx=(0,5))
cb_r = ttk.Combobox(f_r, textvariable=rep_cols_var, values=current_presets_rep, width=15)
cb_r.pack(side=tk.LEFT, padx=(0, 10))
ttk.Button(f_r, text="保存预设", command=lambda: action_save_preset(rep_cols_var, current_presets_rep, cb_r, PRESET_FILE_REP, 3)).pack(side=tk.LEFT, padx=5)
ttk.Button(f_r, text="删除预设", command=lambda: action_del_preset(rep_cols_var, current_presets_rep, cb_r, PRESET_FILE_REP)).pack(side=tk.LEFT, padx=5)

ttk.Label(tab_rep, text="需修改的 SRT 文件夹:").grid(row=2, column=0, sticky="e", padx=(0,10), pady=10)
ttk.Entry(tab_rep, textvariable=rep_srt_var).grid(row=2, column=1, sticky="ew", padx=5, pady=10)
ttk.Button(tab_rep, text="浏览...", command=lambda: ask_dir(rep_srt_var, "选择目录")).grid(row=2, column=2, padx=(5,0), pady=10)

ttk.Label(tab_rep, text="保存替换展示表格:").grid(row=3, column=0, sticky="e", padx=(0,10), pady=10)
ttk.Entry(tab_rep, textvariable=rep_out_var).grid(row=3, column=1, sticky="ew", padx=5, pady=10)
ttk.Button(tab_rep, text="浏览...", command=lambda: ask_save_file(rep_out_var, "保存", [("Excel", "*.xlsx")], ".xlsx")).grid(row=3, column=2, padx=(5,0), pady=10)
ttk.Button(tab_rep, text="开始替换", command=run_replace, style='TButton').grid(row=5, column=0, columnspan=3, pady=15, ipadx=20, ipady=5)

# ================= TAB 8: 双语 SRT 批量拆分 =================
tab_bi = ttk.Frame(nb_srt, padding=20)
nb_srt.add(tab_bi, text=" 双语 SRT 批量拆分 ")
tab_bi.columnconfigure(1, weight=1)

bi_srt_var, bi_out_dir_var = tk.StringVar(), tk.StringVar()
bi_suf1_var, bi_suf2_var = tk.StringVar(value="语言1"), tk.StringVar(value="语言2")

ttk.Label(tab_bi, text="双语 SRT 输入文件夹:").grid(row=0, column=0, sticky="e", pady=15, padx=(0,10))
ttk.Entry(tab_bi, textvariable=bi_srt_var).grid(row=0, column=1, sticky="ew", padx=5)
ttk.Button(tab_bi, text="浏览...", command=lambda: ask_dir(bi_srt_var, "选择输入文件夹")).grid(row=0, column=2, padx=5)

ttk.Label(tab_bi, text="上方语言文件后缀:").grid(row=1, column=0, sticky="e", pady=10, padx=(0,10))
ttk.Entry(tab_bi, textvariable=bi_suf1_var, width=15).grid(row=1, column=1, sticky="w", padx=5)

ttk.Label(tab_bi, text="下方语言文件后缀:").grid(row=2, column=0, sticky="e", pady=10, padx=(0,10))
ttk.Entry(tab_bi, textvariable=bi_suf2_var, width=15).grid(row=2, column=1, sticky="w", padx=5)

ttk.Label(tab_bi, text="拆分后单语保存目录:").grid(row=3, column=0, sticky="e", pady=15, padx=(0,10))
ttk.Entry(tab_bi, textvariable=bi_out_dir_var).grid(row=3, column=1, sticky="ew", padx=5)
ttk.Button(tab_bi, text="浏览...", command=lambda: ask_dir(bi_out_dir_var, "选择输出目录")).grid(row=3, column=2, padx=5)

ttk.Label(tab_bi, text="* 注：支持 1-4 行的混合长段，4行(3换行符)将自动完美 2+2 切分。", foreground="gray").grid(row=4, column=0, columnspan=3, pady=(0,10))
ttk.Button(tab_bi, text="开始批量拆分双语", command=run_srt_bilingual_split, style='TButton').grid(row=5, column=0, columnspan=3, pady=10, ipadx=20, ipady=5)

# ================= TAB 10: 自动时间轴拆分 =================
tab_ts = ttk.Frame(nb_srt, padding=20)
nb_srt.add(tab_ts, text=" 自动时间轴拆分 ")
tab_ts.columnconfigure(1, weight=1)

ts_in_var = tk.StringVar()
ts_norm_var = tk.StringVar()
ts_scr_var = tk.StringVar()
ts_mode_var = tk.StringVar(value="SRT")

ttk.Label(tab_ts, text="输入文件夹:").grid(row=0, column=0, sticky="e", pady=10, padx=(0,10))
ttk.Entry(tab_ts, textvariable=ts_in_var).grid(row=0, column=1, sticky="ew", padx=5)
ttk.Button(tab_ts, text="浏览...", command=lambda: ask_dir(ts_in_var, "选择输入文件夹")).grid(row=0, column=2, padx=5)

f_ts_mode = ttk.Frame(tab_ts)
f_ts_mode.grid(row=1, column=0, columnspan=3, sticky="w", pady=5)
ttk.Radiobutton(f_ts_mode, text="处理 SRT", variable=ts_mode_var, value="SRT").pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(f_ts_mode, text="处理 ASS", variable=ts_mode_var, value="ASS").pack(side=tk.LEFT, padx=5)

ttk.Label(tab_ts, text="对白输出目录:").grid(row=2, column=0, sticky="e", pady=10, padx=(0,10))
ttk.Entry(tab_ts, textvariable=ts_norm_var).grid(row=2, column=1, sticky="ew", padx=5)
ttk.Button(tab_ts, text="浏览...", command=lambda: ask_dir(ts_norm_var, "选择输出目录")).grid(row=2, column=2, padx=5)

ttk.Label(tab_ts, text="画面字输出目录:").grid(row=3, column=0, sticky="e", pady=10, padx=(0,10))
ttk.Entry(tab_ts, textvariable=ts_scr_var).grid(row=3, column=1, sticky="ew", padx=5)
ttk.Button(tab_ts, text="浏览...", command=lambda: ask_dir(ts_scr_var, "选择输出目录")).grid(row=3, column=2, padx=5)

ttk.Label(tab_ts, text="* 智能侦测：工具会逐行扫描文件，一旦发现某行字幕的【开始时间】早于上一行的【结束时间】，\n即自动判定该行为分界点。该行及下方划为对白字幕，上方划为画面字。", foreground="gray").grid(row=4, column=0, columnspan=3, pady=(5,10), sticky="w")
ttk.Button(tab_ts, text="开始自动拆分", command=run_time_split, style='TButton').grid(row=5, column=0, columnspan=3, pady=10, ipadx=20, ipady=5)

# ================= TAB 4: 打包 =================
tab_zip = ttk.Frame(nb_other, padding=20)
nb_other.add(tab_zip, text=" 分部分打包 ")
tab_zip.columnconfigure(1, weight=1)
zip_target_var, zip_out_var, zip_max_var = tk.StringVar(), tk.StringVar(), tk.StringVar(value="50")

ttk.Label(tab_zip, text="需打包的文件夹:").grid(row=0, column=0, sticky="e", padx=(0,10), pady=20)
ttk.Entry(tab_zip, textvariable=zip_target_var).grid(row=0, column=1, sticky="ew", padx=5, pady=20)
ttk.Button(tab_zip, text="浏览...", command=lambda: ask_dir(zip_target_var, "选择")).grid(row=0, column=2, padx=(5,0), pady=20)
ttk.Label(tab_zip, text="单包最大文件数:").grid(row=1, column=0, sticky="e", padx=(0,10), pady=10)
ttk.Entry(tab_zip, textvariable=zip_max_var, width=15).grid(row=1, column=1, sticky="w", padx=5, pady=10)
ttk.Label(tab_zip, text="ZIP 输出文件夹:").grid(row=2, column=0, sticky="e", padx=(0,10), pady=20)
ttk.Entry(tab_zip, textvariable=zip_out_var).grid(row=2, column=1, sticky="ew", padx=5, pady=20)
ttk.Button(tab_zip, text="浏览...", command=lambda: ask_dir(zip_out_var, "选择")).grid(row=2, column=2, padx=(5,0), pady=20)
ttk.Button(tab_zip, text="开始打包", command=run_zip, style='TButton').grid(row=4, column=0, columnspan=3, pady=25, ipadx=20, ipady=5)

def build_style_tab(parent, font_v, size_v, col_v, ocol_v, mv_v, mlr_v, out_v, align_v, shad_v, bold_v, ita_v, alpha_v=None, oalpha_v=None):
    container = ttk.Frame(parent)
    # 将 rowspan 改为 2，完美适配全新的极限两行布局
    container.grid(row=0, column=0, rowspan=2, columnspan=8, sticky="ew")
    
    # ---- 第一行：字体排版与颜色 (极致紧凑整合) ----
    r1 = ttk.Frame(container)
    r1.pack(fill=tk.X, pady=(2, 2), anchor="w")
    
    ttk.Label(r1, text="字体:").pack(side=tk.LEFT)
    cb = ttk.Combobox(r1, textvariable=font_v, width=12)
    cb.pack(side=tk.LEFT, padx=(2, 5))
    
    ttk.Label(r1, text="字号:").pack(side=tk.LEFT)
    ttk.Entry(r1, textvariable=size_v, width=3).pack(side=tk.LEFT, padx=(2, 5))
    
    ttk.Checkbutton(r1, text="加粗", variable=bold_v).pack(side=tk.LEFT, padx=(0, 2))
    ttk.Checkbutton(r1, text="斜体", variable=ita_v).pack(side=tk.LEFT, padx=(0, 5))
    
    ttk.Label(r1, text="主色:").pack(side=tk.LEFT)
    c_btn = tk.Button(r1, width=2, bg="#FFFFFF", relief="ridge", command=lambda: choose_color(col_v, c_btn))
    c_btn.pack(side=tk.LEFT, padx=(2, 5))
    if alpha_v is not None:
        ttk.Label(r1, text="透:").pack(side=tk.LEFT)
        ttk.Entry(r1, textvariable=alpha_v, width=2).pack(side=tk.LEFT, padx=(2, 5))
        
    ttk.Label(r1, text="描边色:").pack(side=tk.LEFT)
    oc_btn = tk.Button(r1, width=2, bg="#000000", relief="ridge", command=lambda: choose_color(ocol_v, oc_btn))
    oc_btn.pack(side=tk.LEFT, padx=(2, 5))
    if oalpha_v is not None:
        ttk.Label(r1, text="透:").pack(side=tk.LEFT)
        ttk.Entry(r1, textvariable=oalpha_v, width=2).pack(side=tk.LEFT, padx=(2, 0))

    # ---- 第二行：特效参数与边距调参 (极致紧凑整合) ----
    r2 = ttk.Frame(container)
    r2.pack(fill=tk.X, pady=(2, 5), anchor="w")
    
    ttk.Label(r2, text="粗细:").pack(side=tk.LEFT)
    ttk.Entry(r2, textvariable=out_v, width=2).pack(side=tk.LEFT, padx=(2, 5))
    
    ttk.Label(r2, text="阴影:").pack(side=tk.LEFT)
    ttk.Entry(r2, textvariable=shad_v, width=2).pack(side=tk.LEFT, padx=(2, 5))

    ttk.Label(r2, text="对齐:").pack(side=tk.LEFT)
    ttk.Combobox(r2, textvariable=align_v, values=[str(i) for i in range(1,10)], width=2, state="readonly").pack(side=tk.LEFT, padx=(2, 5))
    
    ttk.Label(r2, text="上下边距:").pack(side=tk.LEFT)
    ttk.Entry(r2, textvariable=mv_v, width=3).pack(side=tk.LEFT, padx=(2, 5))
    
    ttk.Label(r2, text="左右边距:").pack(side=tk.LEFT)
    ttk.Entry(r2, textvariable=mlr_v, width=3).pack(side=tk.LEFT, padx=(2, 5))
    
    # 嵌套 open_visual_adjuster (内部逻辑和绑定完全一致)
    def open_visual_adjuster(parent_win, mv_var, mlr_var, align_var, font_var, size_var, col_var, ocol_var, out_var, bold_var, ita_var, alpha_var, oalpha_var):
        res_win = tk.Toplevel(parent_win)
        res_win.title("输入画板参数")
        res_win.resizable(False, False)
        
        f_res = ttk.Frame(res_win, padding=20)
        f_res.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(f_res, text="宽 (Width):").grid(row=0, column=0, pady=5, padx=5, sticky="e")
        ent_w = ttk.Entry(f_res, width=15); ent_w.insert(0, "1080"); ent_w.grid(row=0, column=1, pady=5, sticky="w")
        ttk.Label(f_res, text="高 (Height):").grid(row=1, column=0, pady=5, padx=5, sticky="e")
        ent_h = ttk.Entry(f_res, width=15); ent_h.insert(0, "1920"); ent_h.grid(row=1, column=1, pady=5, sticky="w")
        
        bg_color_var = tk.StringVar(value="#000000")
        ttk.Label(f_res, text="画板背景色:").grid(row=2, column=0, pady=5, padx=5, sticky="e")
        f_bg = ttk.Frame(f_res)
        f_bg.grid(row=2, column=1, sticky="w", pady=5)
        bg_btn = tk.Button(f_bg, width=4, bg="#000000", command=lambda: choose_color(bg_color_var, bg_btn))
        bg_btn.pack(side=tk.LEFT)
        
        def open_canvas():
            try: w, h = int(ent_w.get()), int(ent_h.get())
            except: return messagebox.showwarning("错误", "请输入有效的数字！")
            
            vid_bg_color = bg_color_var.get() or "#000000"
            res_win.destroy()
            
            cv_win = tk.Toplevel(parent_win)
            cv_win.title("拖拽调整边距 (完美支持透明度与色彩混合预览)")
            
            screen_w = parent_win.winfo_screenwidth()
            screen_h = parent_win.winfo_screenheight()
            max_h = screen_h * 0.75
            max_w = screen_w * 0.75
            scale = min(max_w / w, max_h / h)
            
            cw, ch = int(w * scale), int(h * scale)
            cv_win.resizable(False, False) 
            
            top_frame = tk.Frame(cv_win)
            top_frame.pack(fill=tk.X, pady=10)
            
            canvas = tk.Canvas(cv_win, width=cw, height=ch, bg="#333333", highlightthickness=2, highlightbackground="#555555")
            canvas.pack(padx=20, pady=(0, 20))
            
            rect_id = canvas.create_rectangle(0, 0, cw, ch, fill=vid_bg_color, outline="gray")
            
            def get_val(var, default, is_int=True):
                if not var: return default
                try: 
                    val = var.get()
                    return int(val) if is_int else (val if val else default)
                except: return default
                 
            def blend_hex(fg_hex, bg_hex, alpha_hex):
                try:
                    a = int(alpha_hex, 16)
                    if a < 0: a = 0
                    if a > 255: a = 255
                    opacity = (255 - a) / 255.0

                    fg = fg_hex.lstrip('#')
                    bg = bg_hex.lstrip('#')
                    if len(fg) != 6: fg = "FFFFFF"
                    if len(bg) != 6: bg = "000000"

                    r_fg, g_fg, b_fg = int(fg[0:2], 16), int(fg[2:4], 16), int(fg[4:6], 16)
                    r_bg, g_bg, b_bg = int(bg[0:2], 16), int(bg[2:4], 16), int(bg[4:6], 16)

                    r = int(r_fg * opacity + r_bg * (1 - opacity))
                    g = int(g_fg * opacity + g_bg * (1 - opacity))
                    b = int(b_fg * opacity + b_bg * (1 - opacity))

                    return f"#{r:02X}{g:02X}{b:02X}"
                except:
                    return fg_hex if fg_hex.startswith("#") else "#FFFFFF"

            drag_data = {"mv": get_val(mv_var, 20), "mlr": get_val(mlr_var, 20)}
            text_ids = []
            crosshair_ids = []
            mouse_pos = {"x": cw // 2, "y": ch // 2, "show": False}
            
            def draw_preview():
                for tid in text_ids + crosshair_ids: canvas.delete(tid)
                text_ids.clear()
                crosshair_ids.clear()
                
                align = get_val(align_var, 2)
                mv, mlr = drag_data["mv"], drag_data["mlr"]
                
                try:
                    f_name = get_val(font_var, GLOBAL_FONT, False)
                except NameError:
                    f_name = get_val(font_var, "Microsoft YaHei", False)
                f_size = max(8, int(get_val(size_var, 60) * scale * 0.8))
                b_str = "bold" if get_val(bold_var, 0) == 1 else "normal"
                i_str = "italic" if get_val(ita_var, 0) == 1 else "roman"
                tk_font = (f_name, f_size, b_str, i_str)
                
                fg_col_raw = get_val(col_var, "#FFFFFF", False)
                bg_col_raw = get_val(ocol_var, "#000000", False)
                fg_alpha = get_val(alpha_var, "00", False)
                bg_alpha = get_val(oalpha_var, "00", False)
                
                fg_col = blend_hex(fg_col_raw, vid_bg_color, fg_alpha)
                bg_col = blend_hex(bg_col_raw, vid_bg_color, bg_alpha)
                
                out_w = max(0, int(get_val(out_var, 2) * scale))
                
                if align in [1, 2, 3]: cy = ch - (mv * scale)
                elif align in [7, 8, 9]: cy = mv * scale
                else: cy = ch / 2
                
                if align in [1, 4, 7]: cx = mlr * scale
                elif align in [3, 6, 9]: cx = cw - (mlr * scale)
                else: cx = cw / 2
                
                anc_map = {1: "sw", 2: "s", 3: "se", 4: "w", 5: "center", 6: "e", 7: "nw", 8: "n", 9: "ne"}
                anc = anc_map.get(align, "s")
                
                if out_w > 0:
                    offsets = [(out_w,0), (-out_w,0), (0,out_w), (0,-out_w), (out_w,out_w), (-out_w,-out_w), (out_w,-out_w), (-out_w,out_w)]
                    for dx, dy in offsets:
                        tid = canvas.create_text(cx+dx, cy+dy, text="【示例字幕预览】", fill=bg_col, font=tk_font, anchor=anc)
                        text_ids.append(tid)
                        
                tid = canvas.create_text(cx, cy, text="【示例字幕预览】", fill=fg_col, font=tk_font, anchor=anc)
                text_ids.append(tid)

                if mouse_pos["show"]:
                    mx, my = mouse_pos["x"], mouse_pos["y"]
                    l1 = canvas.create_line(0, my, cw, my, fill="#00FF00", dash=(4, 4))
                    l2 = canvas.create_line(mx, 0, mx, ch, fill="#00FF00", dash=(4, 4))
                    crosshair_ids.extend([l1, l2])
                    
                    info_txt = f"左右边距(Margin L/R): {drag_data['mlr']} \n上下边距(Margin V): {drag_data['mv']}"
                    tx = mx + 10 if mx < cw - 180 else mx - 10
                    ty = my + 10 if my < ch - 40 else my - 30
                    anc_txt = "nw" if mx < cw - 180 else "ne"
                    
                    tid_shadow = canvas.create_text(tx+1, ty+1, text=info_txt, fill="#000000", font=("Microsoft YaHei", 9, "bold"), anchor=anc_txt)
                    tid_main = canvas.create_text(tx, ty, text=info_txt, fill="#00FF00", font=("Microsoft YaHei", 9, "bold"), anchor=anc_txt)
                    crosshair_ids.extend([tid_shadow, tid_main])
            
            draw_preview()
            
            def update_mouse(event, is_drag=False):
                nx, ny = max(0, min(event.x, cw)), max(0, min(event.y, ch))
                mouse_pos["x"] = nx
                mouse_pos["y"] = ny
                mouse_pos["show"] = True
                
                if is_drag:
                    align = get_val(align_var, 2)
                    if align in [1, 2, 3]: new_mv = (ch - ny) / scale
                    elif align in [7, 8, 9]: new_mv = ny / scale
                    else: new_mv = drag_data["mv"] 
                    
                    if align in [1, 4, 7]: new_mlr = nx / scale
                    elif align in [3, 6, 9]: new_mlr = (cw - nx) / scale
                    else: new_mlr = drag_data["mlr"]
                    
                    drag_data["mv"] = int(new_mv)
                    drag_data["mlr"] = int(new_mlr)
                    
                draw_preview()

            canvas.bind("<Motion>", lambda e: update_mouse(e, False))
            canvas.bind("<B1-Motion>", lambda e: update_mouse(e, True))
            canvas.bind("<Leave>", lambda e: [mouse_pos.update({"show": False}), draw_preview()])
            
            def confirm_and_sync():
                mv_var.set(str(drag_data["mv"]))
                mlr_var.set(str(drag_data["mlr"]))
                draw_preview() 
                messagebox.showinfo("同步成功", "边距参数已写入主界面！\n画板已根据主界面最新的颜色/对齐/透明度等参数刷新预览。", parent=cv_win)
                
            ttk.Button(top_frame, text="✅ 确认修改并同步预览", command=confirm_and_sync, style='TButton').pack(ipadx=10, ipady=2)
            ttk.Label(top_frame, text="拖拽文字调整边距。如果在主界面修改了透明度/颜色/字体等，点击上方按钮即可刷新预览！", foreground="gray").pack(pady=(5,0))
            
        ttk.Button(f_res, text="确定并打开画板", command=open_canvas, style='TButton').grid(row=3, column=0, columnspan=2, pady=(15, 0), ipadx=10, ipady=2)

    ttk.Button(r2, text="📐 图形调参", command=lambda: open_visual_adjuster(parent, mv_v, mlr_v, align_v, font_v, size_v, col_v, ocol_v, out_v, bold_v, ita_v, alpha_v, oalpha_v)).pack(side=tk.LEFT, padx=(5, 0))

    return cb, c_btn, oc_btn

def scan_ass_for_styles(path):
    s = {}
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8-sig') as f:
            in_s = False
            for line in f:
                l = line.strip()
                if l.startswith('[V4+ Styles]'): in_s = True; continue
                if l.startswith('[Events]'): break
                if in_s and l.startswith('Style:'): s[l.split('Style:')[1].split(',')[0].strip()] = l
    return s

# ================= TAB 5: 纯净 SRT 转 ASS =================
tab_ass = ttk.Frame(nb_ass, padding=10)
nb_ass.add(tab_ass, text=" SRT转ASS ")
tab_ass.columnconfigure(1, weight=1)

ass_srt_var, ass_out_var = tk.StringVar(), tk.StringVar()
ass_bracket_var = tk.StringVar(value="【】")

ttk.Label(tab_ass, text="仅限 SRT 输入文件夹:").grid(row=0, column=0, sticky="e", padx=(0,5), pady=5)
ttk.Entry(tab_ass, textvariable=ass_srt_var).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
ttk.Button(tab_ass, text="浏览...", command=lambda: ask_dir(ass_srt_var, "选择目录")).grid(row=0, column=2, padx=(5,0), pady=5)

ttk.Label(tab_ass, text="ASS 输出文件夹:").grid(row=1, column=0, sticky="e", padx=(0,5), pady=5)
ttk.Entry(tab_ass, textvariable=ass_out_var).grid(row=1, column=1, sticky="ew", padx=5, pady=5)
ttk.Button(tab_ass, text="浏览...", command=lambda: ask_dir(ass_out_var, "选择目录")).grid(row=1, column=2, padx=(5,0), pady=5)

f_ass_txt = ttk.LabelFrame(tab_ass, text="文本处理 (画面字提取、正则替换、合并)", padding=10)
f_ass_txt.grid(row=2, column=0, columnspan=3, sticky="ew", pady=10, padx=5)
f_ass_txt.columnconfigure(1, weight=1)

ttk.Label(f_ass_txt, text="画面字识别符号 (一对):").grid(row=0, column=0, sticky="e", padx=(0,5))
ttk.Entry(f_ass_txt, textvariable=ass_bracket_var).grid(row=0, column=1, sticky="ew", padx=5)
ttk.Label(f_ass_txt, text="例如: 【】 或 []", font=("Arial", 8)).grid(row=0, column=2, sticky="w", padx=5)

ttk.Label(f_ass_txt, text="正则批量替换:").grid(row=1, column=0, sticky="ne", padx=(0,5), pady=5)
ass_regex_text = tk.Text(f_ass_txt, height=3, width=50, font=('Arial', 9))
ass_regex_text.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
ass_regex_text.insert(tk.END, "示例_正则查找 >>> 示例_替换成什么\n例如将【】替换为[]，输入【(.*?)】 >>> [$1]")

ass_merge_var = tk.IntVar(value=0)
ass_merge_report_var = tk.StringVar()
ttk.Checkbutton(f_ass_txt, text="合并相邻且相同的字幕 (时间轴无缝衔接)", variable=ass_merge_var).grid(row=3, column=0, columnspan=2, sticky="w", padx=5, pady=5)
ttk.Label(f_ass_txt, text="合并报告导出至:").grid(row=4, column=0, sticky="e", padx=(0,5))
ttk.Entry(f_ass_txt, textvariable=ass_merge_report_var).grid(row=4, column=1, sticky="ew", padx=5)
ttk.Button(f_ass_txt, text="浏览...", command=lambda: ask_save_file(ass_merge_report_var, "保存合并报告", [("Excel", "*.xlsx")], ".xlsx")).grid(row=4, column=2, padx=5)

f_ass_style = ttk.LabelFrame(tab_ass, text="样式设置", padding=10)
f_ass_style.grid(row=3, column=0, columnspan=3, sticky="ew", pady=5, padx=5)

ass_style_mode_5 = tk.IntVar(value=0)

ass_frame_custom = ttk.Frame(f_ass_style)
f_ass_ref_5 = ttk.Frame(f_ass_style)

# --- 插入：分辨率 UI ---
ass_resx_var, ass_resy_var = tk.StringVar(value="1080"), tk.StringVar(value="1920")
f_ass_res_5 = ttk.Frame(ass_frame_custom)
f_ass_res_5.pack(fill=tk.X, pady=(0, 5))
ttk.Label(f_ass_res_5, text="视频分辨率 (宽/X):").pack(side=tk.LEFT)
ttk.Entry(f_ass_res_5, textvariable=ass_resx_var, width=8).pack(side=tk.LEFT, padx=5)
ttk.Label(f_ass_res_5, text="(高/Y):").pack(side=tk.LEFT)
ttk.Entry(f_ass_res_5, textvariable=ass_resy_var, width=8).pack(side=tk.LEFT, padx=5)
# ---------------------

def update_ass_style_mode_5():
    if ass_style_mode_5.get() == 0:
        ass_frame_custom.pack(fill=tk.BOTH, expand=True, pady=5)
        f_ass_ref_5.pack_forget()
    else:
        ass_frame_custom.pack_forget()
        f_ass_ref_5.pack(fill=tk.BOTH, expand=True, pady=5)

f_mode_btns_5 = ttk.Frame(f_ass_style)
f_mode_btns_5.pack(fill=tk.X, pady=2)
ttk.Radiobutton(f_mode_btns_5, text="使用下方自定义样式 (分 对白字幕/画面字)", variable=ass_style_mode_5, value=0, command=update_ass_style_mode_5).pack(side=tk.LEFT, padx=(0, 20))
ttk.Radiobutton(f_mode_btns_5, text="直接复用外部 ASS 样式 (完美保留排版/特效等)", variable=ass_style_mode_5, value=1, command=update_ass_style_mode_5).pack(side=tk.LEFT)

current_presets_ass = load_presets(PRESET_FILE_ASS, DEFAULT_PRESETS_ASS)
ass_preset_cb_var = tk.StringVar()

ass_n_font_var, ass_n_size_var, ass_n_color_var, ass_n_outcolor_var = tk.StringVar(value="SimHei"), tk.StringVar(value="60"), tk.StringVar(value="#FFFFFF"), tk.StringVar(value="#000000")
ass_n_marginv_var, ass_n_marginlr_var, ass_n_outline_var = tk.StringVar(value="20"), tk.StringVar(value="20"), tk.StringVar(value="2")
ass_n_align_var, ass_n_shadow_var, ass_n_bold_var, ass_n_italic_var = tk.StringVar(value="2"), tk.StringVar(value="0"), tk.IntVar(value=0), tk.IntVar(value=0)

ass_s_font_var, ass_s_size_var, ass_s_color_var, ass_s_outcolor_var = tk.StringVar(value="SimHei"), tk.StringVar(value="60"), tk.StringVar(value="#26E3FF"), tk.StringVar(value="#000000")
ass_s_marginv_var, ass_s_marginlr_var, ass_s_outline_var = tk.StringVar(value="850"), tk.StringVar(value="20"), tk.StringVar(value="2")
ass_s_align_var, ass_s_shadow_var, ass_s_bold_var, ass_s_italic_var = tk.StringVar(value="8"), tk.StringVar(value="0"), tk.IntVar(value=0), tk.IntVar(value=0)

style_notebook = ttk.Notebook(ass_frame_custom)
style_notebook.pack(fill=tk.X, pady=5)
tab_normal = ttk.Frame(style_notebook, padding=10); style_notebook.add(tab_normal, text=" 对白字幕 ")
tab_screen = ttk.Frame(style_notebook, padding=10); style_notebook.add(tab_screen, text=" 画面字 ")

ass_n_alpha_var = tk.StringVar(value="00"); ass_n_outalpha_var = tk.StringVar(value="00")
ass_s_alpha_var = tk.StringVar(value="00"); ass_s_outalpha_var = tk.StringVar(value="00")
n_cb, n_c_btn, n_oc_btn = build_style_tab(tab_normal, ass_n_font_var, ass_n_size_var, ass_n_color_var, ass_n_outcolor_var, ass_n_marginv_var, ass_n_marginlr_var, ass_n_outline_var, ass_n_align_var, ass_n_shadow_var, ass_n_bold_var, ass_n_italic_var, ass_n_alpha_var, ass_n_outalpha_var)
s_cb, s_c_btn, s_oc_btn = build_style_tab(tab_screen, ass_s_font_var, ass_s_size_var, ass_s_color_var, ass_s_outcolor_var, ass_s_marginv_var, ass_s_marginlr_var, ass_s_outline_var, ass_s_align_var, ass_s_shadow_var, ass_s_bold_var, ass_s_italic_var, ass_s_alpha_var, ass_s_outalpha_var)

if current_presets_ass: ass_preset_cb_var.set(list(current_presets_ass.keys())[0])
f_ps = create_ass_preset_bar(
    ass_frame_custom, 
    [ass_n_font_var, ass_n_size_var, ass_n_color_var, ass_n_outcolor_var, ass_n_marginv_var, ass_n_marginlr_var, ass_n_outline_var, ass_n_align_var, ass_n_shadow_var, ass_n_bold_var, ass_n_italic_var, ass_n_alpha_var, ass_n_outalpha_var],
    [ass_s_font_var, ass_s_size_var, ass_s_color_var, ass_s_outcolor_var, ass_s_marginv_var, ass_s_marginlr_var, ass_s_outline_var, ass_s_align_var, ass_s_shadow_var, ass_s_bold_var, ass_s_italic_var, ass_s_alpha_var, ass_s_outalpha_var],
    [n_c_btn, n_oc_btn, s_c_btn, s_oc_btn], ass_preset_cb_var, [ass_resx_var, ass_resy_var]
)

f_ps.pack(fill=tk.X, pady=(5, 5))
ttk.Button(f_ps, text="▶ 开始 SRT 转 ASS", command=run_ass_convert, style='TButton').pack(side=tk.LEFT, padx=(40, 0), ipadx=15, ipady=2)

# REF MODE 5
ass_ref_path_5 = tk.StringVar()
f_ref_top_5 = ttk.Frame(f_ass_ref_5)
f_ref_top_5.pack(anchor="w", pady=5)
ttk.Label(f_ref_top_5, text="外部 ASS 文件:").pack(side=tk.LEFT)
ttk.Entry(f_ref_top_5, textvariable=ass_ref_path_5, width=30).pack(side=tk.LEFT, padx=5)
ttk.Button(f_ref_top_5, text="浏览...", command=lambda: ask_file(ass_ref_path_5, "选择ASS", [("ASS","*.ass")])).pack(side=tk.LEFT)

f_ref_mid_5 = ttk.Frame(f_ass_ref_5)
f_ref_mid_5.pack(anchor="w", pady=5)

ass_ref_n_var_5, ass_ref_s_var_5 = tk.StringVar(), tk.StringVar()
# 修复：将下拉框的父容器改为 f_ref_mid_5
cb_ref_n_5 = ttk.Combobox(f_ref_mid_5, textvariable=ass_ref_n_var_5, width=15)
cb_ref_s_5 = ttk.Combobox(f_ref_mid_5, textvariable=ass_ref_s_var_5, width=15)

def scan_ref_5():
    s = scan_ass_for_styles(ass_ref_path_5.get().strip())
    k = list(s.keys())
    cb_ref_n_5['values'] = k; cb_ref_s_5['values'] = k
    if k: ass_ref_n_var_5.set(k[0]); ass_ref_s_var_5.set(k[0])

ttk.Button(f_ref_top_5, text="扫描样式 ->", command=scan_ref_5).pack(side=tk.LEFT, padx=5)

ttk.Label(f_ref_mid_5, text="赋予给普通字:").pack(side=tk.LEFT)
cb_ref_n_5.pack(side=tk.LEFT, padx=5)
ttk.Label(f_ref_mid_5, text="赋予给画面字:").pack(side=tk.LEFT, padx=10)
cb_ref_s_5.pack(side=tk.LEFT, padx=5)

f_ref_bot_5 = ttk.Frame(f_ass_ref_5)
f_ref_bot_5.pack(anchor="w", pady=5)
ass_ref_font_mode_5 = tk.IntVar(value=0)
ass_ref_override_font_5 = tk.StringVar(value="SimHei")
ttk.Radiobutton(f_ref_bot_5, text="保留参考样式的原始字体", variable=ass_ref_font_mode_5, value=0).pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(f_ref_bot_5, text="覆盖字体为:", variable=ass_ref_font_mode_5, value=1).pack(side=tk.LEFT, padx=5)
cb_ref_font_5 = ttk.Combobox(f_ref_bot_5, textvariable=ass_ref_override_font_5, width=25)
cb_ref_font_5.pack(side=tk.LEFT, padx=5)
ttk.Button(f_ref_bot_5, text="▶ 开始 SRT 转 ASS", command=run_ass_convert, style='TButton').pack(side=tk.LEFT, padx=(40, 0), ipadx=15, ipady=2)

update_ass_style_mode_5()

# ================= TAB 9: SRT 分类合并转 ASS =================
tab_ms = ttk.Frame(nb_ass, padding=10)
nb_ass.add(tab_ms, text=" SRT分类合并转ASS ")
tab_ms.columnconfigure(1, weight=1)

ms_norm_var, ms_scr_var, ms_out_var = tk.StringVar(), tk.StringVar(), tk.StringVar()

ttk.Label(tab_ms, text="对白字幕 SRT 文件夹:").grid(row=0, column=0, sticky="e", pady=5, padx=5)
ttk.Entry(tab_ms, textvariable=ms_norm_var).grid(row=0, column=1, sticky="ew", padx=5)
ttk.Button(tab_ms, text="浏览...", command=lambda: ask_dir(ms_norm_var, "选择对白字幕目录")).grid(row=0, column=2, padx=5)

ttk.Label(tab_ms, text="画面字 SRT 文件夹:").grid(row=1, column=0, sticky="e", pady=5, padx=5)
ttk.Entry(tab_ms, textvariable=ms_scr_var).grid(row=1, column=1, sticky="ew", padx=5)
ttk.Button(tab_ms, text="浏览...", command=lambda: ask_dir(ms_scr_var, "选择画面字目录")).grid(row=1, column=2, padx=5)

ttk.Label(tab_ms, text="合成 ASS 输出文件夹:").grid(row=2, column=0, sticky="e", pady=5, padx=5)
ttk.Entry(tab_ms, textvariable=ms_out_var).grid(row=2, column=1, sticky="ew", padx=5)
ttk.Button(tab_ms, text="浏览...", command=lambda: ask_dir(ms_out_var, "选择输出目录")).grid(row=2, column=2, padx=5)

# --- 新增：文本正则替换面板 ---
f_ms_txt = ttk.LabelFrame(tab_ms, text="文本处理 (合并前对源字幕文件进行正则替换，不影响源文件)", padding=10)
f_ms_txt.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 5), padx=5)

ms_enable_regex_var = tk.IntVar(value=0)
ttk.Checkbutton(f_ms_txt, text="开启正则替换", variable=ms_enable_regex_var).grid(row=0, column=0, sticky="w", padx=5)

f_ms_txt_sub = ttk.Frame(f_ms_txt)
f_ms_txt_sub.grid(row=0, column=1, sticky="w", padx=10)
ttk.Label(f_ms_txt_sub, text="应用到:").pack(side=tk.LEFT)
ms_regex_target_var = tk.StringVar(value="画面字")
ttk.Combobox(f_ms_txt_sub, textvariable=ms_regex_target_var, values=["画面字", "对白字幕", "全部"], width=10).pack(side=tk.LEFT, padx=5)

ms_regex_text = tk.Text(f_ms_txt, height=2, width=50, font=('Arial', 9))
ms_regex_text.grid(row=1, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
ms_regex_text.insert(tk.END, "示例_正则查找 >>> 示例_替换成什么\n例如将字幕首尾添加【】，输入(.*) >>> 【$1】\n")

f_ms_style = ttk.LabelFrame(tab_ms, text="样式设置 (画面字将自动排在普通字上层)", padding=10)
f_ms_style.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(5, 10), padx=5)

ms_style_mode_9 = tk.IntVar(value=0)

f_ms_frame_custom = ttk.Frame(f_ms_style)
f_ms_ref_9 = ttk.Frame(f_ms_style)

# --- 插入：分辨率 UI ---
m5_resx_var, m5_resy_var = tk.StringVar(value="1080"), tk.StringVar(value="1920")
f_ms_res = ttk.Frame(f_ms_frame_custom)
f_ms_res.pack(fill=tk.X, pady=(0, 5))
ttk.Label(f_ms_res, text="视频分辨率 (宽/X):").pack(side=tk.LEFT)
ttk.Entry(f_ms_res, textvariable=m5_resx_var, width=8).pack(side=tk.LEFT, padx=5)
ttk.Label(f_ms_res, text="(高/Y):").pack(side=tk.LEFT)
ttk.Entry(f_ms_res, textvariable=m5_resy_var, width=8).pack(side=tk.LEFT, padx=5)
# ---------------------

def update_ms_style_mode_9():
    if ms_style_mode_9.get() == 0:
        f_ms_frame_custom.pack(fill=tk.BOTH, expand=True, pady=5)
        f_ms_ref_9.pack_forget()
    else:
        f_ms_frame_custom.pack_forget()
        f_ms_ref_9.pack(fill=tk.BOTH, expand=True, pady=5)

ttk.Radiobutton(f_ms_style, text="使用下方自定义样式 (转换时分别分配给 对白字幕 和 画面字)", variable=ms_style_mode_9, value=0, command=update_ms_style_mode_9).pack(anchor="w", pady=2)
ttk.Radiobutton(f_ms_style, text="直接复用外部 ASS 样式 (完美保留所有排版与特效等私有参数)", variable=ms_style_mode_9, value=1, command=update_ms_style_mode_9).pack(anchor="w", pady=2)

m5_n_font_var, m5_n_size_var, m5_n_color_var, m5_n_outcolor_var = tk.StringVar(value="SimHei"), tk.StringVar(value="60"), tk.StringVar(value="#FFFFFF"), tk.StringVar(value="#000000")
m5_n_marginv_var, m5_n_marginlr_var, m5_n_outline_var = tk.StringVar(value="20"), tk.StringVar(value="20"), tk.StringVar(value="2")
m5_n_align_var, m5_n_shadow_var, m5_n_bold_var, m5_n_italic_var = tk.StringVar(value="2"), tk.StringVar(value="0"), tk.IntVar(value=0), tk.IntVar(value=0)

m5_s_font_var, m5_s_size_var, m5_s_color_var, m5_s_outcolor_var = tk.StringVar(value="SimHei"), tk.StringVar(value="60"), tk.StringVar(value="#26E3FF"), tk.StringVar(value="#000000")
m5_s_marginv_var, m5_s_marginlr_var, m5_s_outline_var = tk.StringVar(value="850"), tk.StringVar(value="20"), tk.StringVar(value="2")
m5_s_align_var, m5_s_shadow_var, m5_s_bold_var, m5_s_italic_var = tk.StringVar(value="8"), tk.StringVar(value="0"), tk.IntVar(value=0), tk.IntVar(value=0)

style_notebook_ms = ttk.Notebook(f_ms_frame_custom)
style_notebook_ms.pack(fill=tk.X, pady=5)
tab_norm_ms = ttk.Frame(style_notebook_ms, padding=10); style_notebook_ms.add(tab_norm_ms, text=" 对白字幕 ")
tab_scr_ms = ttk.Frame(style_notebook_ms, padding=10); style_notebook_ms.add(tab_scr_ms, text=" 画面字 ")

m5_n_alpha_var = tk.StringVar(value="00"); m5_n_outalpha_var = tk.StringVar(value="00")
m5_s_alpha_var = tk.StringVar(value="00"); m5_s_outalpha_var = tk.StringVar(value="00")
cb_msn, c_btn_msn, oc_btn_msn = build_style_tab(tab_norm_ms, m5_n_font_var, m5_n_size_var, m5_n_color_var, m5_n_outcolor_var, m5_n_marginv_var, m5_n_marginlr_var, m5_n_outline_var, m5_n_align_var, m5_n_shadow_var, m5_n_bold_var, m5_n_italic_var, m5_n_alpha_var, m5_n_outalpha_var)
cb_mss, c_btn_mss, oc_btn_mss = build_style_tab(tab_scr_ms, m5_s_font_var, m5_s_size_var, m5_s_color_var, m5_s_outcolor_var, m5_s_marginv_var, m5_s_marginlr_var, m5_s_outline_var, m5_s_align_var, m5_s_shadow_var, m5_s_bold_var, m5_s_italic_var, m5_s_alpha_var, m5_s_outalpha_var)

ms_preset_var = tk.StringVar()
if current_presets_ass: ms_preset_var.set(list(current_presets_ass.keys())[0])
f_ms_ps = create_ass_preset_bar(
    f_ms_frame_custom,
    [m5_n_font_var, m5_n_size_var, m5_n_color_var, m5_n_outcolor_var, m5_n_marginv_var, m5_n_marginlr_var, m5_n_outline_var, m5_n_align_var, m5_n_shadow_var, m5_n_bold_var, m5_n_italic_var, m5_n_alpha_var, m5_n_outalpha_var],
    [m5_s_font_var, m5_s_size_var, m5_s_color_var, m5_s_outcolor_var, m5_s_marginv_var, m5_s_marginlr_var, m5_s_outline_var, m5_s_align_var, m5_s_shadow_var, m5_s_bold_var, m5_s_italic_var, m5_s_alpha_var, m5_s_outalpha_var],
    [c_btn_msn, oc_btn_msn, c_btn_mss, oc_btn_mss], ms_preset_var, [m5_resx_var, m5_resy_var]
)

f_ms_ps.pack(fill=tk.X, pady=5)
ttk.Button(f_ms_ps, text="▶ 开始双源合并转 ASS", command=run_merge_srt_to_ass, style='TButton').pack(side=tk.LEFT, padx=(40, 0), ipadx=15, ipady=2)

# REF MODE 9
ms_ref_path_9 = tk.StringVar()
f_ref_top_9 = ttk.Frame(f_ms_ref_9)
f_ref_top_9.pack(anchor="w", pady=5)
ttk.Label(f_ref_top_9, text="外部 ASS 文件:").pack(side=tk.LEFT)
ttk.Entry(f_ref_top_9, textvariable=ms_ref_path_9, width=30).pack(side=tk.LEFT, padx=5)
ttk.Button(f_ref_top_9, text="浏览...", command=lambda: ask_file(ms_ref_path_9, "选择ASS", [("ASS","*.ass")])).pack(side=tk.LEFT)

f_ref_mid_9 = ttk.Frame(f_ms_ref_9)
f_ref_mid_9.pack(anchor="w", pady=5)

ms_ref_n_var_9, ms_ref_s_var_9 = tk.StringVar(), tk.StringVar()
# 修复：将下拉框的父容器改为 f_ref_mid_9
cb_ref_n_9 = ttk.Combobox(f_ref_mid_9, textvariable=ms_ref_n_var_9, width=15)
cb_ref_s_9 = ttk.Combobox(f_ref_mid_9, textvariable=ms_ref_s_var_9, width=15)

def scan_ref_9():
    s = scan_ass_for_styles(ms_ref_path_9.get().strip())
    k = list(s.keys())
    cb_ref_n_9['values'] = k; cb_ref_s_9['values'] = k
    if k: ms_ref_n_var_9.set(k[0]); ms_ref_s_var_9.set(k[0])

ttk.Button(f_ref_top_9, text="扫描样式 ->", command=scan_ref_9).pack(side=tk.LEFT, padx=5)

ttk.Label(f_ref_mid_9, text="赋予给普通字:").pack(side=tk.LEFT)
cb_ref_n_9.pack(side=tk.LEFT, padx=5)
ttk.Label(f_ref_mid_9, text="赋予给画面字:").pack(side=tk.LEFT, padx=10)
cb_ref_s_9.pack(side=tk.LEFT, padx=5)

f_ref_bot_9 = ttk.Frame(f_ms_ref_9)
f_ref_bot_9.pack(anchor="w", pady=5)
ms_ref_font_mode_9 = tk.IntVar(value=0)
ms_ref_override_font_9 = tk.StringVar(value="SimHei")
ttk.Radiobutton(f_ref_bot_9, text="保留参考样式的原始字体", variable=ms_ref_font_mode_9, value=0).pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(f_ref_bot_9, text="覆盖字体为:", variable=ms_ref_font_mode_9, value=1).pack(side=tk.LEFT, padx=5)
cb_ref_font_9 = ttk.Combobox(f_ref_bot_9, textvariable=ms_ref_override_font_9, width=25)
cb_ref_font_9.pack(side=tk.LEFT, padx=5)
ttk.Button(f_ref_bot_9, text="▶ 开始双源合并转 ASS", command=run_merge_srt_to_ass, style='TButton').pack(side=tk.LEFT, padx=(40, 0), ipadx=15, ipady=2)

update_ms_style_mode_9()

# ================= TAB 6: ASS 样式修改与复用 =================
tab_edit = ttk.Frame(nb_ass, padding=10)
nb_ass.add(tab_edit, text=" SRT/ASS编辑器 ")
tab_edit.columnconfigure(1, weight=1)

edit_in_var, edit_out_var = tk.StringVar(), tk.StringVar()

# ================= TAB 11: ASS 合并 =================
tab_am = ttk.Frame(nb_ass, padding=20)
nb_ass.add(tab_am, text=" ASS 合并 ")
tab_am.columnconfigure(1, weight=1)

am_dir1_var = tk.StringVar()
am_dir2_var = tk.StringVar()
am_out_var = tk.StringVar()

ttk.Label(tab_am, text="说明: 此功能会将 文件夹2 的字幕行 置于 文件夹1 的字幕行顶部，并自动合并去重样式头。", foreground="gray").grid(row=0, column=0, columnspan=3, pady=(0, 15), sticky="w")

ttk.Label(tab_am, text="文件夹1 (主文件/在下):").grid(row=1, column=0, sticky="e", pady=10, padx=(0,10))
ttk.Entry(tab_am, textvariable=am_dir1_var).grid(row=1, column=1, sticky="ew", padx=5)
ttk.Button(tab_am, text="浏览...", command=lambda: ask_dir(am_dir1_var, "选择文件夹1")).grid(row=1, column=2, padx=5)

ttk.Label(tab_am, text="文件夹2 (要拼接顶部的):").grid(row=2, column=0, sticky="e", pady=10, padx=(0,10))
ttk.Entry(tab_am, textvariable=am_dir2_var).grid(row=2, column=1, sticky="ew", padx=5)
ttk.Button(tab_am, text="浏览...", command=lambda: ask_dir(am_dir2_var, "选择文件夹2")).grid(row=2, column=2, padx=5)

ttk.Label(tab_am, text="合并后输出文件夹:").grid(row=3, column=0, sticky="e", pady=10, padx=(0,10))
ttk.Entry(tab_am, textvariable=am_out_var).grid(row=3, column=1, sticky="ew", padx=5)
ttk.Button(tab_am, text="浏览...", command=lambda: ask_dir(am_out_var, "选择输出文件夹")).grid(row=3, column=2, padx=5)

ttk.Button(tab_am, text="开始合并 ASS", command=run_ass_merge, style='TButton').grid(row=4, column=0, columnspan=3, pady=20, ipadx=20, ipady=5)

# ====== 核心修复：当输入文件夹发生改变时，自动清空旧的内存暂存 ======
def clear_memory_on_dir_change(*args):
    global_ass_memory_cache.clear()
edit_in_var.trace_add("write", clear_memory_on_dir_change)
# ==============================================================

ttk.Label(tab_edit, text="字幕输入文件夹:").grid(row=0, column=0, sticky="e", padx=(0,5), pady=5)
ttk.Entry(tab_edit, textvariable=edit_in_var).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
ttk.Button(tab_edit, text="浏览...", command=lambda: ask_dir(edit_in_var, "选择目录")).grid(row=0, column=2, padx=(5,0), pady=5)

ttk.Label(tab_edit, text="修改后字幕输出:").grid(row=1, column=0, sticky="e", padx=(0,5), pady=5)
ttk.Entry(tab_edit, textvariable=edit_out_var).grid(row=1, column=1, sticky="ew", padx=5, pady=5)
ttk.Button(tab_edit, text="浏览...", command=lambda: ask_dir(edit_out_var, "选择目录")).grid(row=1, column=2, padx=(5,0), pady=5)

# ================= 修复：赋予子面板垂直拉伸能力 =================
# 新增：告诉底层 grid 系统，把多余或缩小的垂直空间全都分配给 row=2（即子面板所在行）
tab_edit.rowconfigure(2, weight=1) 

edit_nb = ttk.Notebook(tab_edit)
# 修改：将 sticky="ew" 改为 sticky="nsew" (North-South-East-West)，允许上下左右全方位拉伸
edit_nb.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=10, padx=5)
# ==============================================================

ASS_COLS = ["0: Layer", "1: Start", "2: End", "3: Style", "4: Name", "5: MarginL", "6: MarginR", "7: MarginV", "8: Effect", "9: Text"]
SRT_COLS = ["0: ID", "1: Timeline", "2: Text"]

# ------ 功能0: 依指定列多选改样式 (替代原功能1和功能4) ------
etab_m0 = ttk.Frame(edit_nb, padding=10)
edit_nb.add(etab_m0, text="按指定列改样式")
etab_m0.columnconfigure(1, weight=1)

f_m0_top = ttk.Frame(etab_m0)
f_m0_top.grid(row=0, column=0, columnspan=3, sticky="ew", pady=5)
ttk.Label(f_m0_top, text="选择要扫描检索的列:").pack(side=tk.LEFT)
m0_col_var = tk.StringVar(value=ASS_COLS[3]) # 默认选择 Style 列
ttk.Combobox(f_m0_top, textvariable=m0_col_var, values=ASS_COLS, state="readonly", width=12).pack(side=tk.LEFT, padx=5)

def scan_m0_cols():
    d = edit_in_var.get().strip()
    if not d or not os.path.exists(d): return messagebox.showwarning("提示", "请先在上方选择 ASS 输入文件夹！")
    ass_files = [os.path.join(d, f) for f in os.listdir(d) if f.lower().endswith('.ass')]
    if not ass_files: return messagebox.showwarning("提示", "输入文件夹中未找到 .ass 文件！")
    
    col_idx = int(m0_col_var.get().split(':')[0])
    vals = set()
    for filepath in ass_files:
        file_name = os.path.basename(filepath)
        # --- 核心：优先从内存暂存中读取 ---
        if file_name in global_ass_memory_cache:
            lines = global_ass_memory_cache[file_name].split('\n')
        else:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f: lines = f.read().split('\n')
                global_ass_memory_cache[file_name] = "\n".join(lines)
            except: continue
            
        for line in lines:
            if line.startswith('Dialogue:'):
                p = line.strip().split(',', 9)
                if len(p) > col_idx and p[col_idx].strip():
                    vals.add(p[col_idx].strip())
                        
    lb_m0_vals.delete(0, tk.END)
    names = sorted(list(vals))
    for n in names: lb_m0_vals.insert(tk.END, n)
    if names: messagebox.showinfo("成功", f"扫描完毕，在当前暂存文件列中发现了 {len(names)} 个唯一数据！")
    else: messagebox.showwarning("提示", "未扫描到任何内容，该列可能全为空。")

ttk.Button(f_m0_top, text="🔍 一键扫描提取数据", command=scan_m0_cols).pack(side=tk.LEFT, padx=10)

f_m0_lb = ttk.Frame(etab_m0)
f_m0_lb.grid(row=1, column=0, columnspan=3, sticky="ew", pady=5)
ttk.Label(f_m0_lb, text="在下方选中要修改的数据项 (支持鼠标拖拽 / 按住Ctrl或Shift多选):").pack(anchor="w")
lb_m0_vals = tk.Listbox(f_m0_lb, selectmode=tk.EXTENDED, height=5, exportselection=False)
lb_m0_vals.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
sb_m0 = ttk.Scrollbar(f_m0_lb, command=lb_m0_vals.yview)
sb_m0.pack(side=tk.LEFT, fill=tk.Y)
lb_m0_vals.config(yscrollcommand=sb_m0.set)

f_m0_mid = ttk.Frame(etab_m0)
f_m0_mid.grid(row=2, column=0, columnspan=3, sticky="ew", pady=5)
ttk.Label(f_m0_mid, text="将选中的字幕行统一替换为新样式，命名为:").pack(side=tk.LEFT)
m0_target_style_var = tk.StringVar(value="Mod_Style")
ttk.Entry(f_m0_mid, textvariable=m0_target_style_var, width=15).pack(side=tk.LEFT, padx=5)

edit_m0_mode = tk.IntVar(value=0)
def update_m0_ui():
    if edit_m0_mode.get() == 0:
        f_m0_custom.grid(row=4, column=0, columnspan=3, sticky="ew")
        f_m0_ref.grid_remove()
    else:
        f_m0_custom.grid_remove()
        f_m0_ref.grid(row=4, column=0, columnspan=3, sticky="w")

ttk.Radiobutton(f_m0_mid, text="赋予下方自定义样式", variable=edit_m0_mode, value=0, command=update_m0_ui).pack(side=tk.LEFT, padx=(20, 5))
ttk.Radiobutton(f_m0_mid, text="从外部 ASS 偷取样式", variable=edit_m0_mode, value=1, command=update_m0_ui).pack(side=tk.LEFT)

# 自定义样式面板
f_m0_custom = ttk.Frame(etab_m0)
e_m0_font, e_m0_size, e_m0_col, e_m0_ocol = tk.StringVar(value="SimHei"), tk.StringVar(value="60"), tk.StringVar(value="#FFFFFF"), tk.StringVar(value="#000000")
e_m0_mv, e_m0_mlr, e_m0_outl = tk.StringVar(value="20"), tk.StringVar(value="20"), tk.StringVar(value="2")
e_m0_align, e_m0_shad, e_m0_bold, e_m0_ita = tk.StringVar(value="2"), tk.StringVar(value="0"), tk.IntVar(value=0), tk.IntVar(value=0)
e_m0_alpha, e_m0_outalpha = tk.StringVar(value="00"), tk.StringVar(value="00")
cb_m0, c_btn_m0, oc_btn_m0 = build_style_tab(f_m0_custom, e_m0_font, e_m0_size, e_m0_col, e_m0_ocol, e_m0_mv, e_m0_mlr, e_m0_outl, e_m0_align, e_m0_shad, e_m0_bold, e_m0_ita, e_m0_alpha, e_m0_outalpha)

e_m0_resx, e_m0_resy = tk.StringVar(value="1080"), tk.StringVar(value="1920")
f_m0_res = ttk.Frame(f_m0_custom)
f_m0_res.grid(row=3, column=0, columnspan=8, sticky="w", pady=(10, 0))
ttk.Label(f_m0_res, text="视频分辨率 (宽/X):").pack(side=tk.LEFT)
ttk.Entry(f_m0_res, textvariable=e_m0_resx, width=8).pack(side=tk.LEFT, padx=5)
ttk.Label(f_m0_res, text="(高/Y):").pack(side=tk.LEFT)
ttk.Entry(f_m0_res, textvariable=e_m0_resy, width=8).pack(side=tk.LEFT, padx=5)

m0_preset_var = tk.StringVar()
if current_presets_ass: m0_preset_var.set(list(current_presets_ass.keys())[0])
f_m0_ps = create_ass_preset_bar(f_m0_custom, [e_m0_font, e_m0_size, e_m0_col, e_m0_ocol, e_m0_mv, e_m0_mlr, e_m0_outl, e_m0_align, e_m0_shad, e_m0_bold, e_m0_ita, e_m0_alpha, e_m0_outalpha], None, [c_btn_m0, oc_btn_m0], m0_preset_var, [e_m0_resx, e_m0_resy])
f_m0_ps.grid(row=4, column=0, columnspan=8, sticky="w", pady=5)

# 外部参考面板
f_m0_ref = ttk.Frame(etab_m0)
m0_ref_path, m0_ref_style = tk.StringVar(), tk.StringVar()
f_m0_rtop = ttk.Frame(f_m0_ref)
f_m0_rtop.pack(anchor="w", pady=5)
ttk.Label(f_m0_rtop, text="外部 ASS 文件:").pack(side=tk.LEFT)
ttk.Entry(f_m0_rtop, textvariable=m0_ref_path, width=30).pack(side=tk.LEFT, padx=5)
ttk.Button(f_m0_rtop, text="浏览...", command=lambda: ask_file(m0_ref_path, "选择", [("ASS","*.ass")])).pack(side=tk.LEFT, padx=5)

m0_ref_cb = ttk.Combobox(f_m0_rtop, textvariable=m0_ref_style, width=15)
ttk.Button(f_m0_rtop, text="扫描样式 ->", command=lambda: scan_ref_for_cb(m0_ref_path.get(), m0_ref_cb, m0_ref_style)).pack(side=tk.LEFT, padx=5)
m0_ref_cb.pack(side=tk.LEFT, padx=5)

f_m0_rbot = ttk.Frame(f_m0_ref)
f_m0_rbot.pack(anchor="w", pady=5)
e_m0_font_mode, e_m0_override_font = tk.IntVar(value=0), tk.StringVar(value="SimHei")
ttk.Radiobutton(f_m0_rbot, text="保留参考样式的原始字体", variable=e_m0_font_mode, value=0).pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(f_m0_rbot, text="覆盖字体为:", variable=e_m0_font_mode, value=1).pack(side=tk.LEFT, padx=5)
cb_m0_ref_font = ttk.Combobox(f_m0_rbot, textvariable=e_m0_override_font, width=25)
cb_m0_ref_font.pack(side=tk.LEFT, padx=5)

update_m0_ui()

# === 注意：确保完全删除了 原功能4（etab_eff） 区块的所有代码 ===

# ------ 功能2: 根据字符重新分配 ------
etab_tag = ttk.Frame(edit_nb, padding=10); edit_nb.add(etab_tag, text="根据[]重划定对白/画面字")
etab_tag.columnconfigure(1, weight=1)
edit_m2_bracket = tk.StringVar(value="【】")
ttk.Label(etab_tag, text="画面字包裹符 (一对):").grid(row=0, column=0, sticky="e", pady=5)
ttk.Entry(etab_tag, textvariable=edit_m2_bracket, width=10).grid(row=0, column=1, sticky="w", padx=5)
ttk.Label(etab_tag, text="(如果整句话被该符号包裹，则判定为画面字，反之则为对白字幕)", font=("Arial", 8)).grid(row=0, column=1, columnspan=2, sticky="e")

edit_m2_mode = tk.IntVar(value=0)
f_m2_container = ttk.Frame(etab_tag)  

def update_m2_ui():
    if edit_m2_mode.get() == 0:
        f_m2_container.grid(row=2, column=0, columnspan=3, sticky="ew"); f_m2_ref.grid_remove()
    else:
        f_m2_container.grid_remove(); f_m2_ref.grid(row=2, column=0, columnspan=3, sticky="w")

ttk.Radiobutton(etab_tag, text="重新划分后，使用下方自定义样式赋予", variable=edit_m2_mode, value=0, command=update_m2_ui).grid(row=1, column=0, columnspan=2, sticky="w", pady=10)
ttk.Radiobutton(etab_tag, text="重新划分后，从外部 ASS 偷取样式赋予", variable=edit_m2_mode, value=1, command=update_m2_ui).grid(row=1, column=1, columnspan=2, sticky="w", pady=10)

f_m2_custom = ttk.Notebook(f_m2_container)
f_m2_custom.pack(fill=tk.BOTH, expand=True)
t2_n, t2_s = ttk.Frame(f_m2_custom, padding=10), ttk.Frame(f_m2_custom, padding=10)
f_m2_custom.add(t2_n, text=" 新的对白字幕样式 "); f_m2_custom.add(t2_s, text=" 新的画面字样式 ")

e_m2n_font, e_m2n_size, e_m2n_col, e_m2n_ocol = tk.StringVar(value="SimHei"), tk.StringVar(value="60"), tk.StringVar(value="#FFFFFF"), tk.StringVar(value="#000000")
e_m2n_mv, e_m2n_mlr, e_m2n_outl = tk.StringVar(value="20"), tk.StringVar(value="20"), tk.StringVar(value="2")
e_m2n_align, e_m2n_shad, e_m2n_bold, e_m2n_ita = tk.StringVar(value="2"), tk.StringVar(value="0"), tk.IntVar(value=0), tk.IntVar(value=0)
e_m2n_alpha, e_m2n_outalpha = tk.StringVar(value="00"), tk.StringVar(value="00")
cb_m2n, c_btn_m2n, oc_btn_m2n = build_style_tab(t2_n, e_m2n_font, e_m2n_size, e_m2n_col, e_m2n_ocol, e_m2n_mv, e_m2n_mlr, e_m2n_outl, e_m2n_align, e_m2n_shad, e_m2n_bold, e_m2n_ita, e_m2n_alpha, e_m2n_outalpha)

# --- 这里补回了被吞掉的画面字变量声明 ---
e_m2s_font, e_m2s_size, e_m2s_col, e_m2s_ocol = tk.StringVar(value="SimHei"), tk.StringVar(value="60"), tk.StringVar(value="#26E3FF"), tk.StringVar(value="#000000")
e_m2s_mv, e_m2s_mlr, e_m2s_outl = tk.StringVar(value="850"), tk.StringVar(value="20"), tk.StringVar(value="2")
e_m2s_align, e_m2s_shad, e_m2s_bold, e_m2s_ita = tk.StringVar(value="8"), tk.StringVar(value="0"), tk.IntVar(value=0), tk.IntVar(value=0)
# ----------------------------------------

e_m2s_alpha, e_m2s_outalpha = tk.StringVar(value="00"), tk.StringVar(value="00")
cb_m2s, c_btn_m2s, oc_btn_m2s = build_style_tab(t2_s, e_m2s_font, e_m2s_size, e_m2s_col, e_m2s_ocol, e_m2s_mv, e_m2s_mlr, e_m2s_outl, e_m2s_align, e_m2s_shad, e_m2s_bold, e_m2s_ita, e_m2s_alpha, e_m2s_outalpha)

# --- 新增：分辨率 UI ---
e_m2_resx, e_m2_resy = tk.StringVar(value="1080"), tk.StringVar(value="1920")
f_m2_res = ttk.Frame(f_m2_container)
f_m2_res.pack(fill=tk.X, pady=(5, 0))
ttk.Label(f_m2_res, text="视频分辨率 (宽/X):").pack(side=tk.LEFT)
ttk.Entry(f_m2_res, textvariable=e_m2_resx, width=8).pack(side=tk.LEFT, padx=5)
ttk.Label(f_m2_res, text="(高/Y):").pack(side=tk.LEFT)
ttk.Entry(f_m2_res, textvariable=e_m2_resy, width=8).pack(side=tk.LEFT, padx=5)

m2_preset_var = tk.StringVar()
if current_presets_ass: m2_preset_var.set(list(current_presets_ass.keys())[0])
f_m2_ps = create_ass_preset_bar(
    f_m2_container,
    [e_m2n_font, e_m2n_size, e_m2n_col, e_m2n_ocol, e_m2n_mv, e_m2n_mlr, e_m2n_outl, e_m2n_align, e_m2n_shad, e_m2n_bold, e_m2n_ita, e_m2n_alpha, e_m2n_outalpha],
    [e_m2s_font, e_m2s_size, e_m2s_col, e_m2s_ocol, e_m2s_mv, e_m2s_mlr, e_m2s_outl, e_m2s_align, e_m2s_shad, e_m2s_bold, e_m2s_ita, e_m2s_alpha, e_m2s_outalpha],
    [c_btn_m2n, oc_btn_m2n, c_btn_m2s, oc_btn_m2s], m2_preset_var, [e_m2_resx, e_m2_resy]
)
f_m2_ps.pack(fill=tk.X, pady=5)

f_m2_ref = ttk.Frame(etab_tag)
m2_ref_path, m2_ref_n, m2_ref_s = tk.StringVar(), tk.StringVar(), tk.StringVar()
f_m2_top = ttk.Frame(f_m2_ref)
f_m2_top.pack(anchor="w", pady=5)
ttk.Label(f_m2_top, text="提供样式的外部 ASS 文件:").pack(side=tk.LEFT)
ttk.Entry(f_m2_top, textvariable=m2_ref_path, width=30).pack(side=tk.LEFT, padx=5)
ttk.Button(f_m2_top, text="浏览...", command=lambda: ask_file(m2_ref_path, "选择", [("ASS","*.ass")])).pack(side=tk.LEFT, padx=5)

f_m2_mid = ttk.Frame(f_m2_ref)
f_m2_mid.pack(anchor="w", pady=5)

# 修复：将下拉框的父容器改为 f_m2_mid
m2_n_cb = ttk.Combobox(f_m2_mid, textvariable=m2_ref_n, width=15)
m2_s_cb = ttk.Combobox(f_m2_mid, textvariable=m2_ref_s, width=15)

def scan_ref_2():
    s = scan_ass_for_styles(m2_ref_path.get().strip())
    k = list(s.keys())
    m2_n_cb['values'] = k; m2_s_cb['values'] = k
    if k: m2_ref_n.set(k[0]); m2_ref_s.set(k[0])

ttk.Button(f_m2_top, text="扫描该参考文件中的样式 ->", command=scan_ref_2).pack(side=tk.LEFT, padx=5)

ttk.Label(f_m2_mid, text="将其赋予给对白字幕:").pack(side=tk.LEFT)
m2_n_cb.pack(side=tk.LEFT, padx=5)
ttk.Label(f_m2_mid, text="将其赋予给画面字:").pack(side=tk.LEFT, padx=10)
m2_s_cb.pack(side=tk.LEFT, padx=5)

f_m2_bot = ttk.Frame(f_m2_ref)
f_m2_bot.pack(anchor="w", pady=5)
e_m2_font_mode = tk.IntVar(value=0)
e_m2_override_font = tk.StringVar(value="SimHei")
ttk.Radiobutton(f_m2_bot, text="保留参考样式的原始字体", variable=e_m2_font_mode, value=0).pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(f_m2_bot, text="覆盖字体为:", variable=e_m2_font_mode, value=1).pack(side=tk.LEFT, padx=5)
cb_m2_ref_font = ttk.Combobox(f_m2_bot, textvariable=e_m2_override_font, width=25)
cb_m2_ref_font.pack(side=tk.LEFT, padx=5)


# ------ 功能3: 基于时间轴同步复制样式/特效 ------
etab_sync = ttk.Frame(edit_nb, padding=10)
edit_nb.add(etab_sync, text="依据时间轴复用样式")
etab_sync.columnconfigure(1, weight=1)

m3_ref_dir = tk.StringVar()
ttk.Label(etab_sync, text="提供样式/特效的参考文件夹:").grid(row=0, column=0, sticky="e", pady=5)
ttk.Entry(etab_sync, textvariable=m3_ref_dir).grid(row=0, column=1, sticky="ew", padx=5)
ttk.Button(etab_sync, text="浏览...", command=lambda: ask_dir(m3_ref_dir, "选择参考目录")).grid(row=0, column=2)

m3_sync_type = tk.IntVar(value=0)
ttk.Radiobutton(etab_sync, text="复用样式 (Style) 及排版", variable=m3_sync_type, value=0, command=lambda: m3_keep_font_cb.config(state="normal")).grid(row=1, column=0, sticky="w", pady=5)
ttk.Radiobutton(etab_sync, text="仅复用特效说明列 (Effect)", variable=m3_sync_type, value=1, command=lambda: m3_keep_font_cb.config(state="disabled")).grid(row=1, column=1, sticky="w", pady=5)

m3_keep_font = tk.IntVar(value=0)
m3_keep_font_cb = ttk.Checkbutton(etab_sync, text="替换样式时，保留原本的字体名称 (仅替换颜色、大小等)", variable=m3_keep_font)
m3_keep_font_cb.grid(row=2, column=0, columnspan=3, sticky="w", padx=20)

m3_err_rep = tk.StringVar()
ttk.Label(etab_sync, text="时间轴不匹配报错报告保存至:").grid(row=3, column=0, sticky="e", pady=10)
ttk.Entry(etab_sync, textvariable=m3_err_rep).grid(row=3, column=1, sticky="ew", padx=5)
ttk.Button(etab_sync, text="浏览...", command=lambda: ask_save_file(m3_err_rep, "保存", [("Excel", "*.xlsx")], ".xlsx")).grid(row=3, column=2)

# ------ 功能3: 批量/条件定位正则替换 (合并原功能5和6) ------
etab_f4 = ttk.Frame(edit_nb, padding=10)
edit_nb.add(etab_f4, text="批量/条件正则替换")

f4_format_var = tk.StringVar(value="ASS")
def update_f4_cols():
    if f4_format_var.get() == "ASS":
        cb_f4_tgt['values'] = ASS_COLS
        f4_target_col.set(ASS_COLS[9])
        cb_f4_c1['values'] = ASS_COLS
        f4_cond1_col.set(ASS_COLS[8])
        cb_f4_c2['values'] = ASS_COLS
        f4_cond2_col.set(ASS_COLS[3])
    else:
        cb_f4_tgt['values'] = SRT_COLS
        f4_target_col.set(SRT_COLS[2])
        cb_f4_c1['values'] = SRT_COLS
        f4_cond1_col.set(SRT_COLS[0])
        cb_f4_c2['values'] = SRT_COLS
        f4_cond2_col.set(SRT_COLS[1])

f4_top = ttk.Frame(etab_f4)
f4_top.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
ttk.Radiobutton(f4_top, text="处理 ASS 格式", variable=f4_format_var, value="ASS", command=update_f4_cols).pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(f4_top, text="处理 SRT 格式", variable=f4_format_var, value="SRT", command=update_f4_cols).pack(side=tk.LEFT, padx=5)

f4_target_col = tk.StringVar(value=ASS_COLS[9])
ttk.Label(etab_f4, text="需要应用正则替换的列:").grid(row=1, column=0, sticky="e", pady=5)
cb_f4_tgt = ttk.Combobox(etab_f4, textvariable=f4_target_col, values=ASS_COLS, width=15, state="readonly")
cb_f4_tgt.grid(row=1, column=1, sticky="w", padx=5)

ttk.Label(etab_f4, text="正则替换规则:").grid(row=2, column=0, sticky="ne", pady=5)
f4_regex_text = tk.Text(etab_f4, height=3, width=45, font=('Arial', 9))
f4_regex_text.grid(row=2, column=1, columnspan=3, sticky="w", padx=5, pady=5)
f4_regex_text.insert(tk.END, "示例_正则查找 >>> 示例_替换成什么\n【(.*?)】 >>> [$1]")

f4_cond_frame = ttk.LabelFrame(etab_f4, text="可选过滤条件 (都不勾选则为针对目标列的全量批量替换)", padding=10)
f4_cond_frame.grid(row=3, column=0, columnspan=4, sticky="ew", pady=10, padx=5)

f4_use_cond1 = tk.IntVar(value=0)
f4_cond1_col = tk.StringVar(value=ASS_COLS[8])
f4_cond1_val = tk.StringVar()
ttk.Checkbutton(f4_cond_frame, text="启用条件1 (列名):", variable=f4_use_cond1).grid(row=0, column=0, sticky="e", pady=5)
cb_f4_c1 = ttk.Combobox(f4_cond_frame, textvariable=f4_cond1_col, values=ASS_COLS, width=15, state="readonly")
cb_f4_c1.grid(row=0, column=1, sticky="w", padx=5)
ttk.Label(f4_cond_frame, text="包含/匹配(正则):").grid(row=0, column=2, sticky="e", padx=(10,5))
ttk.Entry(f4_cond_frame, textvariable=f4_cond1_val, width=20).grid(row=0, column=3, sticky="w", padx=5)

f4_use_cond2 = tk.IntVar(value=0)
f4_cond2_col = tk.StringVar(value=ASS_COLS[3])
f4_cond2_val = tk.StringVar()
ttk.Checkbutton(f4_cond_frame, text="启用条件2 (列名):", variable=f4_use_cond2).grid(row=1, column=0, sticky="e", pady=5)
cb_f4_c2 = ttk.Combobox(f4_cond_frame, textvariable=f4_cond2_col, values=ASS_COLS, width=15, state="readonly")
cb_f4_c2.grid(row=1, column=1, sticky="w", padx=5)
ttk.Label(f4_cond_frame, text="包含/匹配(正则):").grid(row=1, column=2, sticky="e", padx=(10,5))
ttk.Entry(f4_cond_frame, textvariable=f4_cond2_val, width=20).grid(row=1, column=3, sticky="w", padx=5)

# ------ 功能7: 手动选中修改样式 ------
etab_f7 = ttk.Frame(edit_nb, padding=5)
edit_nb.add(etab_f7, text="手动选中行改样式")

f7_top = ttk.Frame(etab_f7)
f7_top.pack(fill=tk.X, pady=5)
m7_file_var = tk.StringVar()

def load_m7_files():
    d = edit_in_var.get().strip()
    if not d or not os.path.exists(d): return messagebox.showwarning("提示", "请先在上方选择输入文件夹！")
    files = [f for f in os.listdir(d) if f.lower().endswith('.ass')]
    m7_file_cb['values'] = files
    if files: m7_file_var.set(files[0])
    
def load_m7_content():
    f = m7_file_var.get().strip()
    if not f: return
    p = os.path.join(edit_in_var.get().strip(), f)
    
    # --- 核心：优先从内存暂存中读取以保证多步操作连贯 ---
    if f in global_ass_memory_cache:
        lines = global_ass_memory_cache[f].split('\n')
    else:
        if not os.path.exists(p): return
        with open(p, 'r', encoding='utf-8-sig') as f_in: lines = f_in.read().split('\n')
        global_ass_memory_cache[f] = "\n".join(lines)
    
    m7_tree.delete(*m7_tree.get_children())
    curr, ev_lines_temp = "info", []
    for l in lines:
        if l.strip().startswith('[V4+ Styles]'): curr = "styles"
        elif l.strip().startswith('[Events]'): curr = "events"
        if curr == "events": ev_lines_temp.append(l)
        
    for i, ev in enumerate(ev_lines_temp):
        if ev.startswith('Dialogue:'):
            pts = ev.split(',', 9)
            if len(pts) >= 10: m7_tree.insert('', 'end', values=(i, f"{pts[1]}->{pts[2]}", pts[3], pts[8], pts[9]))

def open_video_preview():
    import threading
    import subprocess
    import tempfile
    
    if not HAS_CV2:
        return messagebox.showerror("缺少依赖", "请先在终端安装库：\npip install opencv-python pillow")
    
    f_name = m7_file_var.get().strip()
    v_dir = m7_video_dir.get().strip()
    if not f_name or not v_dir or not os.path.exists(v_dir): return messagebox.showwarning("警告", "请确保已选择有效的视频文件夹并加载了对应字幕！")
    
    base = os.path.splitext(f_name)[0]
    v_path = next((os.path.join(v_dir, base + ext) for ext in ['.mp4', '.mkv', '.avi', '.mov'] if os.path.exists(os.path.join(v_dir, base + ext))), None)
    if not v_path: return messagebox.showwarning("警告", f"未在视频文件夹找到与 {base} 同名的视频文件！")
    
    sel = m7_tree.selection()
    if not sel: return messagebox.showwarning("警告", "请先在列表中点击一句字幕，以作为预览起点！")

    def t2ms(t_str):
        h, m, s_ms = t_str.split(':')
        s, ms = s_ms.split('.')
        return (int(h)*3600 + int(m)*60 + int(s))*1000 + int(ms)*10

    start_ms = t2ms(m7_tree.item(sel[0], 'values')[1].split('->')[0])
    
    pv_win = tk.Toplevel(root)
    pv_win.title(f"实时视频调参 - {f_name}")
    
    try:
        cap = cv2.VideoCapture(v_path, cv2.CAP_ANY, [cv2.CAP_PROP_HW_ACCELERATION, cv2.VIDEO_ACCELERATION_ANY])
        if not cap.isOpened(): cap = cv2.VideoCapture(v_path)
    except:
        cap = cv2.VideoCapture(v_path)

    if not cap.isOpened():
        pv_win.destroy()
        return messagebox.showerror("错误", "无法读取视频流！")

    fps = cap.get(cv2.CAP_PROP_FPS) or 25.0
    vid_w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    vid_h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    
    if vid_w == 0 or vid_h == 0:
        pv_win.destroy()
        return messagebox.showerror("错误", "视频分辨率读取异常！")

    screen_w, screen_h = root.winfo_screenwidth() * 0.7, root.winfo_screenheight() * 0.7
    scale = min(screen_w / vid_w, screen_h / vid_h)
    cw, ch = int(vid_w * scale), int(vid_h * scale)
    
    pv_win.geometry(f"{cw + 40}x{ch + 80}")
    
    top_f = tk.Frame(pv_win); top_f.pack(fill=tk.X, pady=5)
    canvas = tk.Canvas(pv_win, width=cw, height=ch, bg="black", highlightthickness=0)
    canvas.pack(pady=5, expand=True)
    
    state = {
        "play": False, "ms": start_ms, "seek_req": start_ms, "auto_sync": False,
        "drag_mv": 20, "drag_mlr": 20, "drag_align": 2, "drag_target_id": sel[0] if sel else None,
        "mx": cw//2, "my": ch//2, "cross": False, "bg_img": None, "loop_id": None,
        "last_cache": "", "play_res_x": 1080.0, "play_res_y": 1920.0, 
        "styles": {}, "events": [], "hide_overlay": False, "is_ffmpeging": False
    }

    # ================= 核心1：实时同步解析内存大闸 =================
    def sync_ass_memory():
        content = global_ass_memory_cache.get(f_name, "")
        if content == state["last_cache"]: return False
        
        lines = content.split('\n')
        state["play_res_x"], state["play_res_y"] = 1080.0, 1920.0
        state["styles"].clear()
        state["events"].clear()
        
        tree_ids = m7_tree.get_children()
        idx = 0
        curr = ""
        
        for line in lines:
            l = line.strip()
            if l.startswith('[V4+ Styles]'): curr = "styles"
            elif l.startswith('[Events]'): curr = "events"
            elif l.startswith('PlayResX:'): 
                try: state["play_res_x"] = float(l.split(':')[1].strip())
                except: pass
            elif l.startswith('PlayResY:'): 
                try: state["play_res_y"] = float(l.split(':')[1].strip())
                except: pass
            
            if curr == "styles" and l.startswith('Style:'):
                pts = l.split('Style:')[1].split(',')
                if len(pts) >= 22: state["styles"][pts[0].strip()] = pts
            elif curr == "events" and l.startswith('Dialogue:'):
                pts = l.split(',', 9)
                if len(pts) >= 10:
                    tid = tree_ids[idx] if idx < len(tree_ids) else None
                    state["events"].append({"id": tid, "st": t2ms(pts[1].strip()), "ed": t2ms(pts[2].strip()), "style": pts[3].strip(), "text": pts[9].replace('\\N', '\n')})
                    idx += 1
                    
        state["last_cache"] = content
        return True

    # 初始化拉取一次数据
    sync_ass_memory()

    def parse_color(ass_c):
        try:
            s = ass_c.strip().upper().replace('&H', '')
            if len(s) >= 8: a, b, g, r = s[-8:-6], s[-6:-4], s[-4:-2], s[-2:]
            elif len(s) == 6: a, b, g, r = "00", s[-6:-4], s[-4:-2], s[-2:]
            else: s = s.zfill(6); a, b, g, r = "00", s[-6:-4], s[-4:-2], s[-2:]
            return f"#{r}{g}{b}"
        except: return "#FFFFFF"

    # ================= 核心2：FFmpeg 真实压制效果一键桥接 =================
    def render_ffmpeg_frame():
        if state["is_ffmpeging"]: return
        state["is_ffmpeging"] = True
        btn_ff.config(text="⏳ 正在调用 FFmpeg 渲染...", state=tk.DISABLED)
        
        def task():
            try:
                with tempfile.TemporaryDirectory() as td:
                    ass_path = os.path.join(td, "temp.ass")
                    with open(ass_path, "w", encoding="utf-8-sig") as f:
                        f.write(global_ass_memory_cache[f_name])
                    
                    cmd = ['ffmpeg', '-y', '-ss', str(state["ms"]/1000.0), '-i', v_path, '-vf', 'ass=temp.ass', '-vframes', '1', '-f', 'image2pipe', '-vcodec', 'png', '-']
                    
                    # 隐藏 Windows 弹出的黑色控制台窗口
                    startupinfo = None
                    if platform.system() == "Windows":
                        startupinfo = subprocess.STARTUPINFO()
                        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                        
                    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, cwd=td, startupinfo=startupinfo)
                    img_data, _ = proc.communicate()
                    
                    if img_data:
                        from io import BytesIO
                        img = Image.open(BytesIO(img_data)).convert("RGB")
                        img = img.resize((cw, ch), Image.Resampling.LANCZOS)
                        
                        def update_ui():
                            state["bg_img"] = ImageTk.PhotoImage(image=img)
                            canvas.delete("bg")
                            canvas.create_image(0, 0, anchor=tk.NW, image=state["bg_img"], tags="bg")
                            state["hide_overlay"] = True  # 隐藏 Tkinter 假字幕，展示纯纯的 FFmpeg 真字幕！
                            update_overlay()
                            btn_ff.config(text="📸 获取 FFmpeg 100% 真实压制帧", state=tk.NORMAL)
                            state["is_ffmpeging"] = False
                            
                        pv_win.after(0, update_ui)
                    else: raise Exception("FFmpeg 截取失败")
            except Exception as e:
                pv_win.after(0, lambda: [btn_ff.config(text="📸 获取 FFmpeg 100% 真实压制帧", state=tk.NORMAL), state.update({"is_ffmpeging": False}), messagebox.showerror("调用失败", f"请确保系统已安装 FFmpeg 并配置了环境变量！\n报错详情: {e}", parent=pv_win)])
                
        threading.Thread(target=task, daemon=True).start()

    # ================= 核心3：精准坐标轴 UI 渲染 =================
    # ================= 核心3：精准坐标轴 UI 渲染 (Subtitle Edit 高仿版) =================
    def update_overlay():
        canvas.delete("overlay")
        if state["hide_overlay"]: return
        
        active_subs = [sub for sub in state["events"] if sub["st"] <= state["ms"] <= sub["ed"]]
                
        if active_subs:
            state["auto_sync"] = True
            try:
                current_sel = m7_tree.selection()
                active_ids = [sub["id"] for sub in active_subs if sub["id"]]
                if set(current_sel) != set(active_ids) and active_ids:
                    m7_tree.selection_set(active_ids)
                    m7_tree.see(active_ids[0])
            except: pass
            state["auto_sync"] = False

            sys_fonts = tkfont.families()
            font_alias = {"SimHei": "黑体", "Microsoft YaHei": "微软雅黑", "SimSun": "宋体", "黑体": "SimHei", "微软雅黑": "Microsoft YaHei", "宋体": "SimSun"}
            
            ass_scale_x = cw / state["play_res_x"]
            ass_scale_y = ch / state["play_res_y"]
            
            # 建立真实包围盒缓存矩阵，用于实现和 Subtitle Edit 一致的真实物理碰撞检测
            occupied_bboxes = []
            
            for sub in active_subs:
                s_name = sub["style"]
                if s_name not in state["styles"]: continue

                pts = state["styles"][s_name]
                f_name_ass = pts[1].strip()
                if f_name_ass not in sys_fonts and f_name_ass in font_alias and font_alias[f_name_ass] in sys_fonts:
                    f_name_ass = font_alias[f_name_ass]

                f_size_ass = float(pts[2].strip())
                fg, bg = parse_color(pts[3].strip()), parse_color(pts[5].strip())
                bold = 1 if pts[7].strip() == '-1' else 0
                italic = 1 if pts[8].strip() == '-1' else 0
                out_w = max(0, int(float(pts[16].strip()) * ass_scale_y))
                align, mlr, mv = int(pts[18].strip()), int(pts[19].strip()), int(pts[21].strip())
                
                # 若当前字幕是被鼠标选中拖拽的焦点，覆盖为 UI 的临时坐标
                if sub["id"] == state.get("drag_target_id"):
                    mv, mlr, align = state["drag_mv"], state["drag_mlr"], state["drag_align"]

                # [核心修正 1]: Tkinter 的字体视觉上比 ASS 原生渲染略大，添加 0.85 黄金缩放系数
                px_size = max(8, int(f_size_ass * ass_scale_y * 0.85))
                tk_f = (f_name_ass, -px_size, "bold" if bold else "normal", "italic" if italic else "roman")
                
                cy = ch - (mv * ass_scale_y) if align in [1,2,3] else (mv * ass_scale_y if align in [7,8,9] else ch/2)
                cx = mlr * ass_scale_x if align in [1,4,7] else (cw - (mlr * ass_scale_x) if align in [3,6,9] else cw/2)
                anc = {1: "sw", 2: "s", 3: "se", 4: "w", 5: "center", 6: "e", 7: "nw", 8: "n", 9: "ne"}.get(align, "s")
                
                # [核心修正 2]: 多行字幕的断行对齐逻辑，修复排版乱跳
                just = {1: tk.LEFT, 4: tk.LEFT, 7: tk.LEFT, 3: tk.RIGHT, 6: tk.RIGHT, 9: tk.RIGHT}.get(align, tk.CENTER)
                
                # [核心修正 3]: 剥离 {xxx} 这种 Tkinter 无法识别渲染的 ASS 特效标签代码
                txt = re.sub(r'\{.*?\}', '', sub["text"])

                # [核心修正 4]: 获取该段文字真实的宽高包围盒，用于精准计算碰撞弹开
                dummy_id = canvas.create_text(cx, cy, text=txt, font=tk_f, anchor=anc, justify=just)
                bbox = canvas.bbox(dummy_id)
                canvas.delete(dummy_id)

                if bbox:
                    b_l, b_t, b_r, b_b = bbox
                    shift_y = 0
                    
                    # 模拟 Subtitle Edit 的堆叠碰撞逻辑 (Collision: Normal)
                    while True:
                        collision = False
                        for (u_l, u_t, u_r, u_b) in occupied_bboxes:
                            # 判断矩形是否相交 (留 2px 余量)
                            if not (b_r < u_l or b_l > u_r or b_b + shift_y + 2 < u_t or b_t + shift_y - 2 > u_b):
                                collision = True
                                break
                        if collision:
                            # 底部字幕向上堆叠，顶部字幕向下堆叠
                            if align in [1, 2, 3]: shift_y -= (b_b - b_t + 2)
                            elif align in [7, 8, 9]: shift_y += (b_b - b_t + 2)
                            else: shift_y -= (b_b - b_t + 2)
                        else:
                            break
                            
                    cy += shift_y
                    occupied_bboxes.append((b_l, b_t + shift_y, b_r, b_b + shift_y))

                # [核心修正 5]: 高精度自适应描边渲染 (消除描边十字间隙)
                if out_w > 0:
                    steps = max(8, int(out_w * 3)) # 自适应增加描边密度
                    for i in range(steps):
                        angle = i * (2 * math.pi / steps)
                        dx = out_w * math.cos(angle)
                        dy = out_w * math.sin(angle)
                        canvas.create_text(cx+dx, cy+dy, text=txt, fill=bg, font=tk_f, anchor=anc, justify=just, tags="overlay")
                
                # 绘制主字幕
                canvas.create_text(cx, cy, text=txt, fill=fg, font=tk_f, anchor=anc, justify=just, tags="overlay")
        
        # 拖拽辅助十字线 UI 绘制
        if not state["play"] and state["cross"]:
            mx, my = state["mx"], state["my"]
            canvas.create_line(0, my, cw, my, fill="#00FF00", dash=(4, 4), tags="overlay")
            canvas.create_line(mx, 0, mx, ch, fill="#00FF00", dash=(4, 4), tags="overlay")
            info = f"Margin L/R: {state['drag_mlr']}\nMargin V: {state['drag_mv']}"
            tx, ty = mx + 10 if mx < cw - 150 else mx - 10, my + 10 if my < ch - 40 else my - 30
            anc_txt = "nw" if mx < cw - 150 else "ne"
            canvas.create_text(tx+1, ty+1, text=info, fill="#000000", font=("Microsoft YaHei", 9, "bold"), anchor=anc_txt, tags="overlay")
            canvas.create_text(tx, ty, text=info, fill="#00FF00", font=("Microsoft YaHei", 9, "bold"), anchor=anc_txt, tags="overlay")

    def video_loop():
        if not pv_win.winfo_exists():
            cap.release(); return
        
        # 时刻监视内存变化，只要在主界面点了“暂存”，下一帧立刻更新排版！
        if sync_ass_memory(): update_overlay()
        
        frame_updated = False
        if state["seek_req"] is not None:
            cap.set(cv2.CAP_PROP_POS_MSEC, state["seek_req"])
            state["seek_req"] = None
            ret, frame = cap.read()
            if ret:
                state["ms"] = cap.get(cv2.CAP_PROP_POS_MSEC)
                frame = cv2.cvtColor(cv2.resize(frame, (cw, ch)), cv2.COLOR_BGR2RGB)
                state["bg_img"] = ImageTk.PhotoImage(image=Image.fromarray(frame))
                canvas.delete("bg")
                canvas.create_image(0, 0, anchor=tk.NW, image=state["bg_img"], tags="bg")
                frame_updated = True
        
        elif state["play"]:
            ret, frame = cap.read()
            if ret:
                state["ms"] = cap.get(cv2.CAP_PROP_POS_MSEC)
                frame = cv2.cvtColor(cv2.resize(frame, (cw, ch)), cv2.COLOR_BGR2RGB)
                state["bg_img"] = ImageTk.PhotoImage(image=Image.fromarray(frame))
                canvas.delete("bg")
                canvas.create_image(0, 0, anchor=tk.NW, image=state["bg_img"], tags="bg")
                frame_updated = True
            else: state["play"] = False 
                
        if frame_updated or state["play"]: update_overlay()
            
        state["loop_id"] = pv_win.after(int(1000 / fps) if state["play"] else 50, video_loop)

    def toggle_play(e=None):
        state["play"] = not state["play"]
        state["hide_overlay"] = False # 恢复实时 UI
        if not state["play"]: update_overlay()

    pv_win.bind("<space>", toggle_play)

    def on_mouse_move(e, is_drag=False):
        state["cross"], state["mx"], state["my"] = True, max(0, min(e.x, cw)), max(0, min(e.y, ch))
        if is_drag and not state["play"]:
            state["hide_overlay"] = False
            ass_scale_x, ass_scale_y = cw / state["play_res_x"], ch / state["play_res_y"]
            align = state["drag_align"]
            if align in [1, 2, 3]: state["drag_mv"] = int((ch - state["my"]) / ass_scale_y)
            elif align in [7, 8, 9]: state["drag_mv"] = int(state["my"] / ass_scale_y)
            if align in [1, 4, 7]: state["drag_mlr"] = int(state["mx"] / ass_scale_x)
            elif align in [3, 6, 9]: state["drag_mlr"] = int((cw - state["mx"]) / ass_scale_x)
        if not state["play"]: update_overlay()

    canvas.bind("<Motion>", lambda e: on_mouse_move(e, False))
    canvas.bind("<B1-Motion>", lambda e: on_mouse_move(e, True))
    canvas.bind("<Leave>", lambda e: state.update({"cross": False}) or (update_overlay() if not state["play"] else None))
    
    def on_tree_sync(e):
        if state["auto_sync"]: return
        sel = m7_tree.selection()
        if not sel: return
        
        focus_item = m7_tree.focus()
        target_item = focus_item if focus_item in sel else sel[0]
        
        # 【核心修复2】：将用户显式点击的这行字幕记录为“唯一拖拽目标”
        state["drag_target_id"] = target_item
        
        s_name = m7_tree.item(target_item, 'values')[2]
        if s_name in state["styles"]:
            state["drag_mv"] = int(state["styles"][s_name][21].strip())
            state["drag_mlr"] = int(state["styles"][s_name][19].strip())
            state["drag_align"] = int(state["styles"][s_name][18].strip())
        
        target_st = t2ms(m7_tree.item(target_item, 'values')[1].split('->')[0])
        target_ed = t2ms(m7_tree.item(target_item, 'values')[1].split('->')[1])
        if not (target_st <= state["ms"] <= target_ed): 
            state["seek_req"] = target_st
            state["hide_overlay"] = False # 恢复实时 UI
        else:
            # 如果仅仅是切换了同画面的拖拽焦点，也要刷新一下
            update_overlay()
            
    m7_tree.bind("<<TreeviewSelect>>", on_tree_sync)
    
    ttk.Button(top_f, text="▶/⏸ 播放/暂停视频 (Space)", command=toggle_play, style='TButton').pack(side=tk.LEFT, padx=10)
    ttk.Button(top_f, text="✅ 点击应用拖拽后的边距", command=lambda: e_m7_mv.set(str(state["drag_mv"])) or e_m7_mlr.set(str(state["drag_mlr"])), style='TButton').pack(side=tk.LEFT, padx=10)
    
    btn_ff = ttk.Button(top_f, text="📸 获取 FFmpeg 100% 真实压制帧", command=render_ffmpeg_frame, style='TButton')
    btn_ff.pack(side=tk.LEFT, padx=10)
    
    video_loop()

# === 将原本拥挤在一行的组件，分装到上下两个子容器中 ===
f7_top_1 = ttk.Frame(f7_top)
f7_top_1.pack(fill=tk.X, pady=(0, 2))

ttk.Button(f7_top_1, text="1. 刷新文件夹", command=load_m7_files).pack(side=tk.LEFT, padx=5)
m7_file_cb = ttk.Combobox(f7_top_1, textvariable=m7_file_var, state="readonly", width=25)
m7_file_cb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
ttk.Button(f7_top_1, text="2. 读取字幕内容", command=load_m7_content).pack(side=tk.LEFT, padx=2)

f7_top_2 = ttk.Frame(f7_top)
f7_top_2.pack(fill=tk.X, pady=(2, 0))

m7_video_dir = tk.StringVar()
ttk.Label(f7_top_2, text="参考视频夹(可选):").pack(side=tk.LEFT, padx=(5, 2))
ttk.Entry(f7_top_2, textvariable=m7_video_dir, width=15).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
ttk.Button(f7_top_2, text="浏览...", command=lambda: ask_dir(m7_video_dir, "选择视频")).pack(side=tk.LEFT, padx=2)
ttk.Button(f7_top_2, text="🎬 视频实时预览调参", command=open_video_preview).pack(side=tk.LEFT, padx=5)
# ==============================================================

f7_tree_f = ttk.Frame(etab_f7)
f7_tree_f.pack(fill=tk.X, pady=5)
m7_tree = ttk.Treeview(f7_tree_f, columns=("idx", "time", "style", "effect", "text"), show="headings", height=5)
m7_tree.heading("idx", text="序号"); m7_tree.column("idx", width=40, anchor="center")
m7_tree.heading("time", text="时间轴"); m7_tree.column("time", width=140)
m7_tree.heading("style", text="样式"); m7_tree.column("style", width=100)
m7_tree.heading("effect", text="特效"); m7_tree.column("effect", width=100)
m7_tree.heading("text", text="内容"); m7_tree.column("text", width=300)
m7_tree.pack(side=tk.LEFT, fill=tk.X, expand=True)
sb_m7 = ttk.Scrollbar(f7_tree_f, command=m7_tree.yview)
sb_m7.pack(side=tk.RIGHT, fill=tk.Y)
m7_tree.config(yscrollcommand=sb_m7.set)

f7_bot = ttk.Frame(etab_f7)
f7_bot.pack(fill=tk.BOTH, expand=True)

edit_m7_target_var = tk.StringVar(value="新选中样式")
f7_b_top = ttk.Frame(f7_bot)
f7_b_top.pack(fill=tk.X, pady=5)
ttk.Label(f7_b_top, text="赋予的新样式名称:").pack(side=tk.LEFT, padx=(0,5))
ttk.Entry(f7_b_top, textvariable=edit_m7_target_var, width=15).pack(side=tk.LEFT)

edit_m7_mode = tk.IntVar(value=0)
def update_m7_ui():
    if edit_m7_mode.get() == 0:
        f_m7_custom.pack(fill=tk.BOTH, expand=True); f_m7_ref.pack_forget()
    else:
        f_m7_custom.pack_forget(); f_m7_ref.pack(fill=tk.BOTH, expand=True)

ttk.Radiobutton(f7_b_top, text="使用下方自定义样式赋予", variable=edit_m7_mode, value=0, command=update_m7_ui).pack(side=tk.LEFT, padx=(15, 5))
ttk.Radiobutton(f7_b_top, text="从外部 ASS 偷取样式赋予", variable=edit_m7_mode, value=1, command=update_m7_ui).pack(side=tk.LEFT)

f_m7_custom = ttk.Frame(f7_bot)
e_m7_font, e_m7_size, e_m7_col, e_m7_ocol = tk.StringVar(value="SimHei"), tk.StringVar(value="60"), tk.StringVar(value="#FFFFFF"), tk.StringVar(value="#000000")
e_m7_mv, e_m7_mlr, e_m7_outl = tk.StringVar(value="20"), tk.StringVar(value="20"), tk.StringVar(value="2")
e_m7_align, e_m7_shad, e_m7_bold, e_m7_ita = tk.StringVar(value="2"), tk.StringVar(value="0"), tk.IntVar(value=0), tk.IntVar(value=0)
e_m7_alpha, e_m7_outalpha = tk.StringVar(value="00"), tk.StringVar(value="00")
cb_m7, c_btn_m7, oc_btn_m7 = build_style_tab(f_m7_custom, e_m7_font, e_m7_size, e_m7_col, e_m7_ocol, e_m7_mv, e_m7_mlr, e_m7_outl, e_m7_align, e_m7_shad, e_m7_bold, e_m7_ita, e_m7_alpha, e_m7_outalpha)

e_m7_resx, e_m7_resy = tk.StringVar(value="1080"), tk.StringVar(value="1920")
f_m7_res = ttk.Frame(f_m7_custom)
f_m7_res.grid(row=3, column=0, columnspan=8, sticky="w", pady=(5, 0))
ttk.Label(f_m7_res, text="视频分辨率 (宽/X):").pack(side=tk.LEFT)
ttk.Entry(f_m7_res, textvariable=e_m7_resx, width=8).pack(side=tk.LEFT, padx=5)
ttk.Label(f_m7_res, text="(高/Y):").pack(side=tk.LEFT)
ttk.Entry(f_m7_res, textvariable=e_m7_resy, width=8).pack(side=tk.LEFT, padx=5)

m7_preset_var = tk.StringVar()
if current_presets_ass: m7_preset_var.set(list(current_presets_ass.keys())[0])
f_m7_ps = create_ass_preset_bar(f_m7_custom, [e_m7_font, e_m7_size, e_m7_col, e_m7_ocol, e_m7_mv, e_m7_mlr, e_m7_outl, e_m7_align, e_m7_shad, e_m7_bold, e_m7_ita, e_m7_alpha, e_m7_outalpha], None, [c_btn_m7, oc_btn_m7], m7_preset_var, [e_m7_resx, e_m7_resy])
f_m7_ps.grid(row=4, column=0, columnspan=8, sticky="w", pady=5)

f_m7_ref = ttk.Frame(f7_bot)
m7_ref_path, m7_ref_style = tk.StringVar(), tk.StringVar()
f_m7_top = ttk.Frame(f_m7_ref); f_m7_top.pack(anchor="w", pady=5)
ttk.Label(f_m7_top, text="外部 ASS:").pack(side=tk.LEFT)
ttk.Entry(f_m7_top, textvariable=m7_ref_path, width=25).pack(side=tk.LEFT, padx=5)
ttk.Button(f_m7_top, text="浏览...", command=lambda: ask_file(m7_ref_path, "选择", [("ASS","*.ass")])).pack(side=tk.LEFT, padx=5)
m7_ref_cb = ttk.Combobox(f_m7_top, textvariable=m7_ref_style, width=15)
ttk.Button(f_m7_top, text="扫描样式 ->", command=lambda: scan_ref_for_cb(m7_ref_path.get(), m7_ref_cb, m7_ref_style)).pack(side=tk.LEFT, padx=5)
m7_ref_cb.pack(side=tk.LEFT, padx=5)

def scan_ref_for_cb(path, cb_widget, var_widget):
    s = scan_ass_for_styles(path.strip())
    k = list(s.keys()); cb_widget['values'] = k
    if k: var_widget.set(k[0])

e_m7_font_mode, e_m7_override_font = tk.IntVar(value=0), tk.StringVar(value="SimHei")
f_m7_bot2 = ttk.Frame(f_m7_ref); f_m7_bot2.pack(anchor="w", pady=5)
ttk.Radiobutton(f_m7_bot2, text="保留参考样式原字体", variable=e_m7_font_mode, value=0).pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(f_m7_bot2, text="覆盖字体为:", variable=e_m7_font_mode, value=1).pack(side=tk.LEFT, padx=5)
cb_m7_ref_font = ttk.Combobox(f_m7_bot2, textvariable=e_m7_override_font, width=25)
cb_m7_ref_font.pack(side=tk.LEFT, padx=5)
update_m7_ui()

# ------ 功能5 (对应底层的模式 5): 条件定位替换样式 ------
etab_f8 = ttk.Frame(edit_nb, padding=10)
edit_nb.add(etab_f8, text="条件定位替换样式")

f8_cond_frame = ttk.LabelFrame(etab_f8, text="定位条件 (必须至少启用一个条件，全为正则匹配)", padding=10)
f8_cond_frame.pack(fill=tk.X, pady=5)

f8_use_cond1 = tk.IntVar(value=0)
f8_cond1_col = tk.StringVar(value=ASS_COLS[9])
f8_cond1_val = tk.StringVar()
ttk.Checkbutton(f8_cond_frame, text="启用条件1 (列名):", variable=f8_use_cond1).grid(row=0, column=0, sticky="e", pady=5)
ttk.Combobox(f8_cond_frame, textvariable=f8_cond1_col, values=ASS_COLS, width=15, state="readonly").grid(row=0, column=1, sticky="w", padx=5)
ttk.Label(f8_cond_frame, text="包含/匹配(正则):").grid(row=0, column=2, sticky="e", padx=(10,5))
ttk.Entry(f8_cond_frame, textvariable=f8_cond1_val, width=20).grid(row=0, column=3, sticky="w", padx=5)

f8_use_cond2 = tk.IntVar(value=0)
f8_cond2_col = tk.StringVar(value=ASS_COLS[8])
f8_cond2_val = tk.StringVar()
ttk.Checkbutton(f8_cond_frame, text="启用条件2 (列名):", variable=f8_use_cond2).grid(row=1, column=0, sticky="e", pady=5)
ttk.Combobox(f8_cond_frame, textvariable=f8_cond2_col, values=ASS_COLS, width=15, state="readonly").grid(row=1, column=1, sticky="w", padx=5)
ttk.Label(f8_cond_frame, text="包含/匹配(正则):").grid(row=1, column=2, sticky="e", padx=(10,5))
ttk.Entry(f8_cond_frame, textvariable=f8_cond2_val, width=20).grid(row=1, column=3, sticky="w", padx=5)

f8_bot = ttk.Frame(etab_f8)
f8_bot.pack(fill=tk.BOTH, expand=True, pady=5)

edit_m8_target_var = tk.StringVar(value="新选中样式")
f8_b_top = ttk.Frame(f8_bot)
f8_b_top.pack(fill=tk.X, pady=5)
ttk.Label(f8_b_top, text="符合条件则赋予新样式名称:").pack(side=tk.LEFT, padx=(0,5))
ttk.Entry(f8_b_top, textvariable=edit_m8_target_var, width=15).pack(side=tk.LEFT)

edit_m8_mode = tk.IntVar(value=0)
def update_m8_ui():
    if edit_m8_mode.get() == 0:
        f_m8_custom.pack(fill=tk.BOTH, expand=True); f_m8_ref.pack_forget()
    else:
        f_m8_custom.pack_forget(); f_m8_ref.pack(fill=tk.BOTH, expand=True)

ttk.Radiobutton(f8_b_top, text="使用下方自定义样式赋予", variable=edit_m8_mode, value=0, command=update_m8_ui).pack(side=tk.LEFT, padx=(15, 5))
ttk.Radiobutton(f8_b_top, text="从外部 ASS 偷取样式赋予", variable=edit_m8_mode, value=1, command=update_m8_ui).pack(side=tk.LEFT)

f_m8_custom = ttk.Frame(f8_bot)
e_m8_font, e_m8_size, e_m8_col, e_m8_ocol = tk.StringVar(value="SimHei"), tk.StringVar(value="60"), tk.StringVar(value="#FFFFFF"), tk.StringVar(value="#000000")
e_m8_mv, e_m8_mlr, e_m8_outl = tk.StringVar(value="20"), tk.StringVar(value="20"), tk.StringVar(value="2")
e_m8_align, e_m8_shad, e_m8_bold, e_m8_ita = tk.StringVar(value="2"), tk.StringVar(value="0"), tk.IntVar(value=0), tk.IntVar(value=0)
e_m8_alpha, e_m8_outalpha = tk.StringVar(value="00"), tk.StringVar(value="00")
cb_m8, c_btn_m8, oc_btn_m8 = build_style_tab(f_m8_custom, e_m8_font, e_m8_size, e_m8_col, e_m8_ocol, e_m8_mv, e_m8_mlr, e_m8_outl, e_m8_align, e_m8_shad, e_m8_bold, e_m8_ita, e_m8_alpha, e_m8_outalpha)

e_m8_resx, e_m8_resy = tk.StringVar(value="1080"), tk.StringVar(value="1920")
f_m8_res = ttk.Frame(f_m8_custom)
f_m8_res.grid(row=3, column=0, columnspan=8, sticky="w", pady=(5, 0))
ttk.Label(f_m8_res, text="视频分辨率 (宽/X):").pack(side=tk.LEFT)
ttk.Entry(f_m8_res, textvariable=e_m8_resx, width=8).pack(side=tk.LEFT, padx=5)
ttk.Label(f_m8_res, text="(高/Y):").pack(side=tk.LEFT)
ttk.Entry(f_m8_res, textvariable=e_m8_resy, width=8).pack(side=tk.LEFT, padx=5)

m8_preset_var = tk.StringVar()
if current_presets_ass: m8_preset_var.set(list(current_presets_ass.keys())[0])
f_m8_ps = create_ass_preset_bar(f_m8_custom, [e_m8_font, e_m8_size, e_m8_col, e_m8_ocol, e_m8_mv, e_m8_mlr, e_m8_outl, e_m8_align, e_m8_shad, e_m8_bold, e_m8_ita, e_m8_alpha, e_m8_outalpha], None, [c_btn_m8, oc_btn_m8], m8_preset_var, [e_m8_resx, e_m8_resy])
f_m8_ps.grid(row=4, column=0, columnspan=8, sticky="w", pady=5)

f_m8_ref = ttk.Frame(f8_bot)
m8_ref_path, m8_ref_style = tk.StringVar(), tk.StringVar()
f_m8_top = ttk.Frame(f_m8_ref); f_m8_top.pack(anchor="w", pady=5)
ttk.Label(f_m8_top, text="外部 ASS:").pack(side=tk.LEFT)
ttk.Entry(f_m8_top, textvariable=m8_ref_path, width=30).pack(side=tk.LEFT, padx=5)
ttk.Button(f_m8_top, text="浏览...", command=lambda: ask_file(m8_ref_path, "选择", [("ASS","*.ass")])).pack(side=tk.LEFT, padx=5)
m8_ref_cb = ttk.Combobox(f_m8_top, textvariable=m8_ref_style, width=15)
ttk.Button(f_m8_top, text="扫描样式 ->", command=lambda: scan_ref_for_cb(m8_ref_path.get(), m8_ref_cb, m8_ref_style)).pack(side=tk.LEFT, padx=5)
m8_ref_cb.pack(side=tk.LEFT, padx=5)

e_m8_font_mode, e_m8_override_font = tk.IntVar(value=0), tk.StringVar(value="SimHei")
f_m8_bot2 = ttk.Frame(f_m8_ref); f_m8_bot2.pack(anchor="w", pady=5)
ttk.Radiobutton(f_m8_bot2, text="保留参考样式原字体", variable=e_m8_font_mode, value=0).pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(f_m8_bot2, text="覆盖字体为:", variable=e_m8_font_mode, value=1).pack(side=tk.LEFT, padx=5)
cb_m8_ref_font = ttk.Combobox(f_m8_bot2, textvariable=e_m8_override_font, width=25)
cb_m8_ref_font.pack(side=tk.LEFT, padx=5)
update_m8_ui()

def execute_ass_editor(stage_only=False):
    i_dir, o_dir = edit_in_var.get().strip(), edit_out_var.get().strip()
    if not i_dir or (not stage_only and not o_dir): return messagebox.showwarning("警告", "请填好输入和输出目录！")
    
    if not stage_only: os.makedirs(o_dir, exist_ok=True)
    mode = edit_nb.index(edit_nb.select())

    fmt = "ASS"
    if mode == 3: fmt = f4_format_var.get()
    
    ext = '.srt' if fmt == "SRT" else '.ass'
    files = [f for f in os.listdir(i_dir) if f.lower().endswith(ext)]
    if not files: return messagebox.showwarning("警告", f"输入文件夹中没有 {ext} 文件！")
    
    if mode == 2:
        ref_dir = m3_ref_dir.get().strip()
        if not ref_dir or not os.path.exists(ref_dir): return messagebox.showwarning("警告", "请选择有效的参考 ASS 文件夹！")
        sync_type = m3_sync_type.get()
        keep_font = m3_keep_font.get() == 1
        all_errors = []

    global_ref_resx, global_ref_resy = None, None
    rp = None
    if mode == 0:
        if edit_m0_mode.get() == 1: rp = m0_ref_path.get().strip()
        else:
            if e_m0_resx.get().strip(): global_ref_resx = f"PlayResX: {e_m0_resx.get().strip()}"
            if e_m0_resy.get().strip(): global_ref_resy = f"PlayResY: {e_m0_resy.get().strip()}"
    elif mode == 1:
        if edit_m2_mode.get() == 1: rp = m2_ref_path.get().strip()
        else:
            if e_m2_resx.get().strip(): global_ref_resx = f"PlayResX: {e_m2_resx.get().strip()}"
            if e_m2_resy.get().strip(): global_ref_resy = f"PlayResY: {e_m2_resy.get().strip()}"
    elif mode == 4:
        if edit_m7_mode.get() == 1: rp = m7_ref_path.get().strip()
        else:
            if e_m7_resx.get().strip(): global_ref_resx = f"PlayResX: {e_m7_resx.get().strip()}"
            if e_m7_resy.get().strip(): global_ref_resy = f"PlayResY: {e_m7_resy.get().strip()}"
    elif mode == 5:
        if edit_m8_mode.get() == 1: rp = m8_ref_path.get().strip()
        else:
            if e_m8_resx.get().strip(): global_ref_resx = f"PlayResX: {e_m8_resx.get().strip()}"
            if e_m8_resy.get().strip(): global_ref_resy = f"PlayResY: {e_m8_resy.get().strip()}"
    
    if rp and os.path.exists(rp):
        with open(rp, 'r', encoding='utf-8-sig') as f:
            for l in f:
                if l.startswith('PlayResX:'): global_ref_resx = l.strip()
                elif l.startswith('PlayResY:'): global_ref_resy = l.strip()

    regex_rules = []
    if mode == 3: raw_text = f4_regex_text.get("1.0", tk.END).split('\n')
    else: raw_text = []
    for line in raw_text:
        if '>>>' in line:
            pat, repl = line.split('>>>', 1)
            repl_python = re.sub(r'\$(\d+)', r'\\\1', repl.strip())
            regex_rules.append((pat.strip(), repl_python))

    for file in files:
        if mode == 4 and file != m7_file_var.get().strip(): continue
        
        in_path, out_path = os.path.join(i_dir, file), os.path.join(o_dir, file)
        
        # ====== 内存优先加载 ======
        if file in global_ass_memory_cache:
            content = global_ass_memory_cache[file]
        else:
            with open(in_path, 'r', encoding='utf-8-sig') as f: content = f.read()
            global_ass_memory_cache[file] = content

        # ====== SRT 格式文件独立处理逻辑 ======
        if fmt == "SRT" and mode == 3:
            blocks = re.split(r'\n\s*\n', content.strip())
            parsed_blocks = []
            for block in blocks:
                lines = block.strip().split('\n')
                if len(lines) >= 3:
                    parsed_blocks.append({'ID': lines[0].strip(), 'Timeline': lines[1].strip(), 'Text': "\n".join(lines[2:]).strip()})
                    
            for block in parsed_blocks:
                p = [block['ID'], block['Timeline'], block['Text']]
                c1_idx = int(f4_cond1_col.get().split(':')[0])
                c1_val = f4_cond1_val.get().strip()
                c2_idx = int(f4_cond2_col.get().split(':')[0])
                c2_val = f4_cond2_val.get().strip()
                tgt_idx = int(f4_target_col.get().split(':')[0])
                
                is_match = True
                if f4_use_cond1.get() == 1 and c1_val:
                    if not re.search(c1_val, p[c1_idx]): is_match = False
                if f4_use_cond2.get() == 1 and c2_val and is_match:
                    if not re.search(c2_val, p[c2_idx]): is_match = False
                    
                if is_match:
                    for pat, repl in regex_rules:
                        p[tgt_idx] = re.sub(pat, repl, p[tgt_idx])
                        
                block['ID'], block['Timeline'], block['Text'] = p[0], p[1], p[2]
                
            srt_content = [f"{b['ID']}\n{b['Timeline']}\n{b['Text']}\n" for b in parsed_blocks]
            final_content = "\n".join(srt_content)
            
            if stage_only: global_ass_memory_cache[file] = final_content
            else: 
                global_ass_memory_cache[file] = final_content
                with open(out_path, 'w', encoding='utf-8') as f: f.write(final_content)
            continue

        # ================= ASS 处理 =================
        lines = content.split('\n')
        h_lines, s_lines, ev_lines = [], [], []
        curr = "info"
        tgt_format_line = "Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding"

        for line in lines:
            l = line.strip()
            if l.startswith('[V4+ Styles]'): curr = "styles"
            elif l.startswith('[Events]'): curr = "events"
            
            if curr == "info": h_lines.append(line)
            elif curr == "styles": 
                s_lines.append(line)
                if l.startswith('Format:'): tgt_format_line = l
            elif curr == "events": ev_lines.append(line)
            
        if len(s_lines) <= 1: s_lines = ["[V4+ Styles]", tgt_format_line]

        ref_resx, ref_resy = global_ref_resx, global_ref_resy
        if mode == 2 and sync_type == 0:
            ref_path = os.path.join(ref_dir, file)
            if os.path.exists(ref_path):
                with open(ref_path, 'r', encoding='utf-8-sig') as f:
                    for l in f:
                        if l.startswith('PlayResX:'): ref_resx = l.strip()
                        elif l.startswith('PlayResY:'): ref_resy = l.strip()
        
        if ref_resx and ref_resy:
            has_x = has_y = False
            for i, l in enumerate(h_lines):
                if l.startswith('PlayResX:'): h_lines[i] = ref_resx; has_x = True
                elif l.startswith('PlayResY:'): h_lines[i] = ref_resy; has_y = True
            if not has_x: h_lines.append(ref_resx)
            if not has_y: h_lines.append(ref_resy)

        # ====== 功能0 ======
        if mode == 0:
            sel_idxs = lb_m0_vals.curselection()
            if not sel_idxs: return messagebox.showwarning("警告", "请在列表中选中至少一项要修改的数据！")
            sel_vals = set(lb_m0_vals.get(i) for i in sel_idxs)
            col_idx = int(m0_col_var.get().split(':')[0])
            
            new_style_name = m0_target_style_var.get().strip()
            if not new_style_name: return messagebox.showwarning("警告", "请输入赋予的新样式名称！")
            
            new_line = ""
            if edit_m0_mode.get() == 0:
                new_line = build_ass_style_line(new_style_name, e_m0_font.get(), e_m0_size.get(), e_m0_col.get(), e_m0_ocol.get(), e_m0_mv.get(), e_m0_mlr.get(), e_m0_outl.get(), e_m0_align.get(), e_m0_shad.get(), e_m0_bold.get(), e_m0_ita.get(), e_m0_alpha.get(), e_m0_outalpha.get())
            else:
                rp, rs = m0_ref_path.get(), m0_ref_style.get()
                if not os.path.exists(rp) or not rs: return messagebox.showwarning("警告", "请正确提供参考文件和样式！")
                ref_dict = scan_all_styles_from_ass(rp)
                if rs not in ref_dict: return messagebox.showwarning("错误", "参考中没找到该样式")
                new_line = rename_style_line(ref_dict[rs], new_style_name)
                if e_m0_font_mode.get() == 1:
                    new_line = replace_font_in_style(new_line, e_m0_override_font.get())

            matched_any = False
            new_ev = []
            for ev in ev_lines:
                if ev.startswith('Dialogue:'):
                    p = ev.split(',', 9)
                    if len(p) > col_idx and p[col_idx].strip() in sel_vals:
                        p[3] = new_style_name  # 统一将 Style 引用修改为新样式
                        matched_any = True
                        new_ev.append(",".join(p))
                    else: new_ev.append(ev)
                else: new_ev.append(ev)
            ev_lines = new_ev

            if matched_any:
                rep = False
                for i, sl in enumerate(s_lines):
                    if sl.startswith('Style:') and sl.split('Style:')[1].split(',')[0].strip() == new_style_name:
                        s_lines[i] = new_line; rep = True
                if not rep: s_lines.append(new_line)
                
        # ====== 功能1 ======
        elif mode == 1:
            b = edit_m2_bracket.get().strip()
            l_b, r_b = b[:len(b)//2] if len(b)>=2 else "", b[len(b)//2:] if len(b)>=2 else ""
            
            n_line, s_line = "", ""
            if edit_m2_mode.get() == 0:
                n_line = build_ass_style_line("对白字幕", e_m2n_font.get(), e_m2n_size.get(), e_m2n_col.get(), e_m2n_ocol.get(), e_m2n_mv.get(), e_m2n_mlr.get(), e_m2n_outl.get(), e_m2n_align.get(), e_m2n_shad.get(), e_m2n_bold.get(), e_m2n_ita.get(), e_m2n_alpha.get(), e_m2n_outalpha.get())
                s_line = build_ass_style_line("画面字", e_m2s_font.get(), e_m2s_size.get(), e_m2s_col.get(), e_m2s_ocol.get(), e_m2s_mv.get(), e_m2s_mlr.get(), e_m2s_outl.get(), e_m2s_align.get(), e_m2s_shad.get(), e_m2s_bold.get(), e_m2s_ita.get(), e_m2s_alpha.get(), e_m2s_outalpha.get())
                n_name, s_name = "对白字幕", "画面字"
            else:
                rp, rn, rs = m2_ref_path.get(), m2_ref_n.get(), m2_ref_s.get()
                if not os.path.exists(rp) or not rn or not rs: return messagebox.showwarning("警告", "请正确提供参考文件和两类样式！")
                ref_dict = scan_all_styles_from_ass(rp)
                n_line = rename_style_line(ref_dict[rn], rn)
                s_line = rename_style_line(ref_dict[rs], rs)
                if e_m2_font_mode.get() == 1:
                    n_line = replace_font_in_style(n_line, e_m2_override_font.get())
                    s_line = replace_font_in_style(s_line, e_m2_override_font.get())
                n_name, s_name = rn, rs

            for nl in [n_line, s_line]:
                nm = nl.split('Style:')[1].split(',')[0].strip()
                rep = False
                for i, l in enumerate(s_lines):
                    if l.strip().startswith('Style:') and l.split('Style:')[1].split(',')[0].strip() == nm:
                        s_lines[i] = nl; rep = True
                if not rep: s_lines.append(nl)

            new_ev = []
            for ev in ev_lines:
                parts = ev.split(',', 9)
                if len(parts) >= 10 and ev.strip().startswith('Dialogue:'):
                    txt = parts[9]
                    c_txt = re.sub(r'\{.*?\}', '', txt).strip()
                    if l_b and r_b and c_txt.startswith(l_b) and c_txt.endswith(r_b):
                        parts[3] = s_name
                        new_ev.append(",".join(parts))
                    elif l_b and r_b and l_b in txt and r_b in txt:
                        pat = f"{re.escape(l_b)}[\\s\\S]*?{re.escape(r_b)}"
                        s_t = clean_ass_text("\\N".join(re.findall(pat, txt)))
                        n_t = clean_ass_text(re.sub(pat, "", txt).strip())
                        if s_t:
                            s_p = list(parts); s_p[3] = s_name; s_p[9] = s_t
                            new_ev.append(",".join(s_p))
                        if n_t:
                            n_p = list(parts); n_p[3] = n_name; n_p[9] = n_t
                            new_ev.append(",".join(n_p))
                    else:
                        parts[3] = n_name
                        new_ev.append(",".join(parts))
                else: new_ev.append(ev)
            ev_lines = new_ev

        # ====== 功能2 ======
        elif mode == 2:
            ref_path = os.path.join(ref_dir, file)
            if not os.path.exists(ref_path):
                all_errors.append({'文件名': file, '时间轴': 'N/A', '文本': 'N/A', '错误': '参考文件夹中无同名文件'})
                continue
            
            ref_styles = scan_all_styles_from_ass(ref_path)
            ref_events = {}
            with open(ref_path, 'r', encoding='utf-8-sig') as f:
                for line in f:
                    if line.startswith('Dialogue:'):
                        p = line.strip().split(',', 9)
                        if len(p) >= 10: ref_events[(p[1].strip(), p[2].strip())] = p
            
            tgt_styles = scan_all_styles_from_ass(in_path)
            new_ev = []
            styles_to_add = {}

            for ev in ev_lines:
                if ev.startswith('Dialogue:'):
                    p = ev.split(',', 9)
                    if len(p) >= 10:
                        start, end = p[1].strip(), p[2].strip()
                        if (start, end) in ref_events:
                            rp = ref_events[(start, end)]
                            if sync_type == 1: p[8] = rp[8]
                            else:
                                new_style_name = rp[3]
                                orig_style_name = p[3]
                                p[3] = new_style_name
                                if new_style_name in ref_styles:
                                    st_line = ref_styles[new_style_name]
                                    if keep_font:
                                        orig_font = "Arial"
                                        if orig_style_name in tgt_styles:
                                            orig_font = tgt_styles[orig_style_name].split('Style:')[1].split(',')[1].strip()
                                        st_parts = st_line.split('Style:')[1].split(',')
                                        if len(st_parts) > 1: st_parts[1] = orig_font
                                        st_line = "Style:" + ",".join(st_parts)
                                    styles_to_add[new_style_name] = st_line
                        else:
                            all_errors.append({'文件名': file, '时间轴': f"{start} --> {end}", '文本': p[9], '错误': '在参考文件中未找到对应时间轴'})
                    new_ev.append(",".join(p))
                else: new_ev.append(ev)
            ev_lines = new_ev

            if sync_type == 0:
                for n_name, n_line in styles_to_add.items():
                    rep = False
                    for i, sl in enumerate(s_lines):
                        if sl.startswith('Style:') and sl.split('Style:')[1].split(',')[0].strip() == n_name:
                            s_lines[i] = n_line; rep = True
                    if not rep: s_lines.append(n_line)

        # ====== 功能3 (批量/条件正则替换 - ASS处理) ======
        elif mode == 3:
            tgt_idx = int(f4_target_col.get().split(':')[0])
            c1_idx = int(f4_cond1_col.get().split(':')[0])
            c1_val = f4_cond1_val.get().strip()
            c2_idx = int(f4_cond2_col.get().split(':')[0])
            c2_val = f4_cond2_val.get().strip()
            
            new_ev = []
            for ev in ev_lines:
                if ev.startswith('Dialogue:'):
                    p = ev.split(',', 9)
                    if len(p) >= 10:
                        is_match = True
                        if f4_use_cond1.get() == 1 and c1_val:
                            if not re.search(c1_val, p[c1_idx]): is_match = False
                        if f4_use_cond2.get() == 1 and c2_val and is_match:
                            if not re.search(c2_val, p[c2_idx]): is_match = False
                            
                        if is_match:
                            for pat, repl in regex_rules:
                                p[tgt_idx] = re.sub(pat, repl, p[tgt_idx])
                        new_ev.append(",".join(p))
                    else: new_ev.append(ev)
                else: new_ev.append(ev)
            ev_lines = new_ev

        # ====== 功能4 ======
        elif mode == 4:
            sel_items = m7_tree.selection()
            if not sel_items: return messagebox.showwarning("警告", "请在下方列表中点击选择你要修改的字幕行！")
            sel_indices = [int(m7_tree.item(item, 'values')[0]) for item in sel_items]
            
            new_style_name = edit_m7_target_var.get().strip()
            if not new_style_name: return messagebox.showwarning("警告", "请输入赋予的新样式名称！")
            
            new_line = ""
            if edit_m7_mode.get() == 0:
                new_line = build_ass_style_line(new_style_name, e_m7_font.get(), e_m7_size.get(), e_m7_col.get(), e_m7_ocol.get(), e_m7_mv.get(), e_m7_mlr.get(), e_m7_outl.get(), e_m7_align.get(), e_m7_shad.get(), e_m7_bold.get(), e_m7_ita.get(), e_m7_alpha.get(), e_m7_outalpha.get())
            else:
                rp, rs = m7_ref_path.get(), m7_ref_style.get()
                if not os.path.exists(rp) or not rs: return messagebox.showwarning("警告", "请正确提供参考文件和样式！")
                ref_dict = scan_all_styles_from_ass(rp)
                if rs not in ref_dict: return messagebox.showwarning("错误", "参考中没找到该样式")
                new_line = rename_style_line(ref_dict[rs], new_style_name)
                if e_m7_font_mode.get() == 1: new_line = replace_font_in_style(new_line, e_m7_override_font.get())
                
            rep = False
            for i, sl in enumerate(s_lines):
                if sl.startswith('Style:') and sl.split('Style:')[1].split(',')[0].strip() == new_style_name:
                    s_lines[i] = new_line; rep = True
            if not rep: s_lines.append(new_line)
            
            new_ev = []
            for i, ev in enumerate(ev_lines):
                if i in sel_indices and ev.startswith('Dialogue:'):
                    p = ev.split(',', 9)
                    if len(p) >= 10:
                        p[3] = new_style_name
                        new_ev.append(",".join(p))
                    else: new_ev.append(ev)
                else: new_ev.append(ev)
            ev_lines = new_ev

        # ====== 功能5: 条件定位替换样式 ======
        elif mode == 5:
            if f8_use_cond1.get() == 0 and f8_use_cond2.get() == 0:
                return messagebox.showwarning("警告", "请至少启用一个定位条件！")

            c1_idx = int(f8_cond1_col.get().split(':')[0])
            c1_val = f8_cond1_val.get().strip()
            c2_idx = int(f8_cond2_col.get().split(':')[0])
            c2_val = f8_cond2_val.get().strip()
            
            new_style_name = edit_m8_target_var.get().strip()
            if not new_style_name: return messagebox.showwarning("警告", "请输入赋予的新样式名称！")
            
            new_line = ""
            if edit_m8_mode.get() == 0:
                new_line = build_ass_style_line(new_style_name, e_m8_font.get(), e_m8_size.get(), e_m8_col.get(), e_m8_ocol.get(), e_m8_mv.get(), e_m8_mlr.get(), e_m8_outl.get(), e_m8_align.get(), e_m8_shad.get(), e_m8_bold.get(), e_m8_ita.get(), e_m8_alpha.get(), e_m8_outalpha.get())
            else:
                rp, rs = m8_ref_path.get(), m8_ref_style.get()
                if not os.path.exists(rp) or not rs: return messagebox.showwarning("警告", "请正确提供参考文件和样式！")
                ref_dict = scan_all_styles_from_ass(rp)
                if rs not in ref_dict: return messagebox.showwarning("错误", "参考中没找到该样式")
                new_line = rename_style_line(ref_dict[rs], new_style_name)
                if e_m8_font_mode.get() == 1: new_line = replace_font_in_style(new_line, e_m8_override_font.get())
                
            matched_any = False
            new_ev = []
            for ev in ev_lines:
                if ev.startswith('Dialogue:'):
                    p = ev.split(',', 9)
                    if len(p) >= 10:
                        is_match = True
                        if f8_use_cond1.get() == 1:
                            if not re.search(c1_val, p[c1_idx]): is_match = False
                        if f8_use_cond2.get() == 1 and is_match:
                            if not re.search(c2_val, p[c2_idx]): is_match = False
                            
                        if is_match:
                            p[3] = new_style_name
                            matched_any = True
                        new_ev.append(",".join(p))
                    else: new_ev.append(ev)
                else: new_ev.append(ev)
            ev_lines = new_ev

            if matched_any:
                rep = False
                for i, sl in enumerate(s_lines):
                    if sl.startswith('Style:') and sl.split('Style:')[1].split(',')[0].strip() == new_style_name:
                        s_lines[i] = new_line; rep = True
                if not rep: s_lines.append(new_line)

        final_content = "\n".join(h_lines) + "\n\n" + "\n".join(s_lines) + "\n\n" + "\n".join(ev_lines)
        
        # --- 根据按钮模式，分流输出目标 ---
        if stage_only:
            global_ass_memory_cache[file] = final_content
        else:
            global_ass_memory_cache[file] = final_content
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write(final_content)
            
    if stage_only:
        messagebox.showinfo("暂存成功", "【处理完毕】更改已无缝注入底层内存引擎！\n\n你可以随时切到其他标签页加载、刷新以进行二次、三次叠加操作。\n等所有流水线走完，只需点击底部【批量输出保存】按钮即可一次性落盘。")
    else:
        if mode == 2 and m3_err_rep.get().strip() and all_errors:
            pd.DataFrame(all_errors).to_excel(m3_err_rep.get().strip(), index=False)
            messagebox.showinfo("完成", f"落盘导出完成！但有 {len(all_errors)} 处时间轴不匹配，已导出报错报告。")
        else:
            messagebox.showinfo("完成", f"落盘完成！所有流式处理已同步到输出文件夹中的 {len(files)} 个文件！")

f_edit_bot_btns = ttk.Frame(tab_edit)
f_edit_bot_btns.grid(row=3, column=0, columnspan=3, pady=10)
ttk.Button(f_edit_bot_btns, text="💾 暂存当前更改到内存引擎", command=lambda: execute_ass_editor(stage_only=True), style='TButton').pack(side=tk.LEFT, padx=10, ipadx=10, ipady=5)
ttk.Button(f_edit_bot_btns, text="▶ 执行全部批处理并保存到输出文件夹", command=lambda: execute_ass_editor(stage_only=False), style='TButton').pack(side=tk.LEFT, padx=10, ipadx=10, ipady=5)

# ================= TAB 7: 批量复用指定列(SRT/ASS) =================
tab_eff = ttk.Frame(nb_ass, padding=20)
nb_ass.add(tab_eff, text=" 批量复用指定列(SRT/ASS) ")
tab_eff.columnconfigure(1, weight=1)

eff_fmt_var = tk.StringVar(value="ASS")
eff_col_var = tk.StringVar(value=ASS_COLS[8])

def update_eff_cols():
    if eff_fmt_var.get() == "ASS":
        cb_eff_col['values'] = ASS_COLS
        eff_col_var.set(ASS_COLS[8])
    else:
        cb_eff_col['values'] = SRT_COLS
        eff_col_var.set(SRT_COLS[2])

f_eff_top = ttk.Frame(tab_eff)
f_eff_top.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
ttk.Radiobutton(f_eff_top, text="处理 ASS 格式", variable=eff_fmt_var, value="ASS", command=update_eff_cols).pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(f_eff_top, text="处理 SRT 格式", variable=eff_fmt_var, value="SRT", command=update_eff_cols).pack(side=tk.LEFT, padx=5)

ttk.Label(f_eff_top, text="需要复用覆盖的列:").pack(side=tk.LEFT, padx=(20, 5))
cb_eff_col = ttk.Combobox(f_eff_top, textvariable=eff_col_var, values=ASS_COLS, width=15, state="readonly")
cb_eff_col.pack(side=tk.LEFT)

eff_src_var, eff_tgt_var = tk.StringVar(), tk.StringVar()
eff_out_var, eff_err_var = tk.StringVar(), tk.StringVar()

ttk.Label(tab_eff, text="提供数据的源文件夹:").grid(row=1, column=0, sticky="e", pady=10, padx=(0,10))
ttk.Entry(tab_eff, textvariable=eff_src_var).grid(row=1, column=1, sticky="ew", padx=5)
ttk.Button(tab_eff, text="浏览...", command=lambda: ask_dir(eff_src_var, "选择源文件夹")).grid(row=1, column=2, padx=5)

ttk.Label(tab_eff, text="待接收数据的目标文件夹:").grid(row=2, column=0, sticky="e", pady=10, padx=(0,10))
ttk.Entry(tab_eff, textvariable=eff_tgt_var).grid(row=2, column=1, sticky="ew", padx=5)
ttk.Button(tab_eff, text="浏览...", command=lambda: ask_dir(eff_tgt_var, "选择目标文件夹")).grid(row=2, column=2, padx=5)

ttk.Label(tab_eff, text="合成后的新文件输出至:").grid(row=3, column=0, sticky="e", pady=10, padx=(0,10))
ttk.Entry(tab_eff, textvariable=eff_out_var).grid(row=3, column=1, sticky="ew", padx=5)
ttk.Button(tab_eff, text="浏览...", command=lambda: ask_dir(eff_out_var, "选择输出文件夹")).grid(row=3, column=2, padx=5)

ttk.Label(tab_eff, text="行列不匹配报错报告保存至:").grid(row=4, column=0, sticky="e", pady=10, padx=(0,10))
ttk.Entry(tab_eff, textvariable=eff_err_var).grid(row=4, column=1, sticky="ew", padx=5)
ttk.Button(tab_eff, text="浏览...", command=lambda: ask_save_file(eff_err_var, "保存报错报告", [("Excel", "*.xlsx")], ".xlsx")).grid(row=4, column=2, padx=5)

ttk.Label(tab_eff, text="* 注：将基于【文件同名】和【行数顺序】进行精准的一对一覆盖提取。如果时间轴错位会记录在报告中", foreground="gray").grid(row=5, column=0, columnspan=3, pady=(0,10))
ttk.Button(tab_eff, text="执行指定列批量复制", command=run_column_copy, style='TButton').grid(row=6, column=0, columnspan=3, pady=10, ipadx=20, ipady=5)

# ================= TAB 10: ASS 拆分 (画面字/普通字) =================
tab_ass_split = ttk.Frame(nb_ass, padding=10)
nb_ass.add(tab_ass_split, text=" ASS拆分(画面/对白) ")
tab_ass_split.columnconfigure(1, weight=1)

split_ass_in_var = tk.StringVar()
split_ass_out_scr_var, split_ass_out_norm_var = tk.StringVar(), tk.StringVar()

ttk.Label(tab_ass_split, text="ASS 输入文件夹:").grid(row=0, column=0, sticky="e", pady=5, padx=5)
ttk.Entry(tab_ass_split, textvariable=split_ass_in_var).grid(row=0, column=1, sticky="ew", padx=5)
f_split_in_btns = ttk.Frame(tab_ass_split)
f_split_in_btns.grid(row=0, column=2, sticky="w", padx=5)
ttk.Button(f_split_in_btns, text="浏览...", command=lambda: ask_dir(split_ass_in_var, "选择目录")).pack(side=tk.LEFT)
ttk.Button(f_split_in_btns, text="🔍 扫描特效与样式", command=scan_ass_split_features).pack(side=tk.LEFT, padx=5)

ttk.Label(tab_ass_split, text="画面字 ASS 存至:").grid(row=1, column=0, sticky="e", pady=5, padx=5)
ttk.Entry(tab_ass_split, textvariable=split_ass_out_scr_var).grid(row=1, column=1, sticky="ew", padx=5)
ttk.Button(tab_ass_split, text="浏览...", command=lambda: ask_dir(split_ass_out_scr_var, "选择目录")).grid(row=1, column=2, sticky="w", padx=5)

ttk.Label(tab_ass_split, text="普通字 ASS 存至:").grid(row=2, column=0, sticky="e", pady=5, padx=5)
ttk.Entry(tab_ass_split, textvariable=split_ass_out_norm_var).grid(row=2, column=1, sticky="ew", padx=5)
ttk.Button(tab_ass_split, text="浏览...", command=lambda: ask_dir(split_ass_out_norm_var, "选择目录")).grid(row=2, column=2, sticky="w", padx=5)

f_split_cond = ttk.LabelFrame(tab_ass_split, text="拆分判定条件 (可多选组合，选中项需同时满足即为画面字，其余为普通字)", padding=10)
f_split_cond.grid(row=3, column=0, columnspan=3, sticky="ew", pady=10, padx=5)

# 条件 1
split_ass_c1_var = tk.IntVar(value=0)
split_ass_bracket_var = tk.StringVar(value="【】")
f_sc1 = ttk.Frame(f_split_cond)
f_sc1.pack(fill=tk.X, pady=2)
ttk.Checkbutton(f_sc1, text="条件1: 文本前后包含指定符号组合:", variable=split_ass_c1_var).pack(side=tk.LEFT)
ttk.Entry(f_sc1, textvariable=split_ass_bracket_var, width=10).pack(side=tk.LEFT, padx=5)
ttk.Label(f_sc1, text="(如【】、{}，匹配去除特效后的纯文本)", foreground="gray").pack(side=tk.LEFT)

# 条件 2
split_ass_c2_var = tk.IntVar(value=0)
f_sc2 = ttk.Frame(f_split_cond)
f_sc2.pack(fill=tk.X, pady=5)
ttk.Checkbutton(f_sc2, text="条件2: 包含在以下选中的【特效说明 Effect】内 (支持按住 Ctrl 多选):", variable=split_ass_c2_var).pack(anchor="w")
f_sc2_lb = ttk.Frame(f_sc2)
f_sc2_lb.pack(fill=tk.X, padx=20, pady=2)
lb_split_effs = tk.Listbox(f_sc2_lb, selectmode=tk.MULTIPLE, height=4, exportselection=False)
lb_split_effs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
sb_split_effs = ttk.Scrollbar(f_sc2_lb, command=lb_split_effs.yview)
sb_split_effs.pack(side=tk.LEFT, fill=tk.Y)
lb_split_effs.config(yscrollcommand=sb_split_effs.set)

# 条件 3
split_ass_c3_var = tk.IntVar(value=0)
f_sc3 = ttk.Frame(f_split_cond)
f_sc3.pack(fill=tk.X, pady=5)
ttk.Checkbutton(f_sc3, text="条件3: 包含在以下选中的【样式名称 Style】内 (支持按住 Ctrl 多选):", variable=split_ass_c3_var).pack(anchor="w")
f_sc3_lb = ttk.Frame(f_sc3)
f_sc3_lb.pack(fill=tk.X, padx=20, pady=2)
lb_split_styles = tk.Listbox(f_sc3_lb, selectmode=tk.MULTIPLE, height=4, exportselection=False)
lb_split_styles.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
sb_split_styles = ttk.Scrollbar(f_sc3_lb, command=lb_split_styles.yview)
sb_split_styles.pack(side=tk.LEFT, fill=tk.Y)
lb_split_styles.config(yscrollcommand=sb_split_styles.set)

# 新增：保存为 SRT 的勾选选项
split_ass_to_srt_var = tk.IntVar(value=0)
ttk.Checkbutton(tab_ass_split, text="拆分后自动剥离特效与样式，直接转为标准 SRT 格式保存", variable=split_ass_to_srt_var).grid(row=4, column=0, columnspan=3, sticky="w", padx=10, pady=(5, 0))

ttk.Button(tab_ass_split, text="▶ 开始拆分 ASS", command=run_ass_split, style='TButton').grid(row=5, column=0, columnspan=3, pady=10, ipadx=20, ipady=5)

# ================= TAB 11: ASS 样式预设提取 =================
tab_ext = ttk.Frame(nb_ass, padding=20)
nb_ass.add(tab_ext, text=" ASS 样式预设提取 ")

ext_file_var = tk.StringVar()
ext_style_var = tk.StringVar()
ext_preset_name = tk.StringVar()
ext_styles_cache = {}
ext_res_cache = {"x": "1080", "y": "1920"}

ttk.Label(tab_ext, text="选择用于提取的 ASS 文件:").grid(row=0, column=0, sticky="e", pady=10)
ttk.Entry(tab_ext, textvariable=ext_file_var, width=40).grid(row=0, column=1, sticky="w", padx=5)
ttk.Button(tab_ext, text="浏览...", command=lambda: ask_file(ext_file_var, "选择ASS", [("ASS","*.ass")])).grid(row=0, column=2, padx=5)

def run_scan_ext():
    f = ext_file_var.get().strip()
    if not os.path.exists(f): return messagebox.showwarning("错误", "文件不存在")
    s = scan_ass_for_styles(f)
    rx, ry = get_ass_resolution(f)
    ext_res_cache["x"] = rx
    ext_res_cache["y"] = ry
    
    ext_styles_cache.clear()
    ext_styles_cache.update(s)
    k = list(s.keys())
    cb_ext_style['values'] = k
    if k: ext_style_var.set(k[0])
    messagebox.showinfo("成功", f"扫描到 {len(k)} 个样式\n自动获取到视频分辨率: {rx} x {ry}")

ttk.Button(tab_ext, text="🔍 扫描样式", command=run_scan_ext).grid(row=0, column=3, padx=5)

ttk.Label(tab_ext, text="选择要提取的样式:").grid(row=1, column=0, sticky="e", pady=10)
cb_ext_style = ttk.Combobox(tab_ext, textvariable=ext_style_var, state="readonly", width=30)
cb_ext_style.grid(row=1, column=1, sticky="w", padx=5)

ttk.Label(tab_ext, text="保存为预设名称:").grid(row=2, column=0, sticky="e", pady=10)
ttk.Entry(tab_ext, textvariable=ext_preset_name, width=30).grid(row=2, column=1, sticky="w", padx=5)

def parse_ass_color(ass_str):
    """解析 ASS 颜色代码，返回 (HEX颜色, HEX透明度)"""
    try:
        s = ass_str.strip().upper().replace('&H', '')
        if len(s) >= 8:
            a, b, g, r = s[-8:-6], s[-6:-4], s[-4:-2], s[-2:]
        elif len(s) == 6:
            a, b, g, r = "00", s[-6:-4], s[-4:-2], s[-2:]
        else:
            s = s.zfill(6)
            a, b, g, r = "00", s[-6:-4], s[-4:-2], s[-2:]
        return f"#{r}{g}{b}", a
    except: 
        return "#FFFFFF", "00"

def parse_style_to_dict(style_line):
    parts = style_line.split('Style:')[1].split(',')
    c_hex, c_a = parse_ass_color(parts[3].strip())
    oc_hex, oc_a = parse_ass_color(parts[5].strip())
    return {
        "font": parts[1].strip(), "size": parts[2].strip(),
        "color": c_hex, "alpha": c_a, "out_color": oc_hex, "out_alpha": oc_a,
        "margin_v": parts[21].strip(), "margin_lr": parts[19].strip(),
        "outline": parts[16].strip(), "align": parts[18].strip(),
        "shadow": parts[17].strip(), "bold": 1 if parts[7].strip() == "-1" else 0,
        "italic": 1 if parts[8].strip() == "-1" else 0
    }

def save_ext_preset():
    name = ext_preset_name.get().strip()
    if not name: return messagebox.showwarning("警告", "请输入预设名称")
    target_style = ext_style_var.get()
    if target_style not in ext_styles_cache: return messagebox.showwarning("警告", "请先扫描并选择有效样式")
    
    d = DEFAULT_PRESETS_ASS["默认样式"].copy()
    pd_dict = parse_style_to_dict(ext_styles_cache[target_style])
    
    # 将提取到的样式同时应用给预设底层的双字段，确保在其他页面的面板加载时完美映射
    d.update({
        "play_res_x": ext_res_cache["x"], "play_res_y": ext_res_cache["y"],
        "n_font": pd_dict["font"], "n_size": pd_dict["size"], "n_color": pd_dict["color"], "n_alpha": pd_dict["alpha"], "n_out_color": pd_dict["out_color"], "n_out_alpha": pd_dict["out_alpha"], "n_margin_v": pd_dict["margin_v"], "n_margin_lr": pd_dict["margin_lr"], "n_outline": pd_dict["outline"], "n_align": pd_dict["align"], "n_shadow": pd_dict["shadow"], "n_bold": pd_dict["bold"], "n_italic": pd_dict["italic"],
        "s_font": pd_dict["font"], "s_size": pd_dict["size"], "s_color": pd_dict["color"], "s_alpha": pd_dict["alpha"], "s_out_color": pd_dict["out_color"], "s_out_alpha": pd_dict["out_alpha"], "s_margin_v": pd_dict["margin_v"], "s_margin_lr": pd_dict["margin_lr"], "s_outline": pd_dict["outline"], "s_align": pd_dict["align"], "s_shadow": pd_dict["shadow"], "s_bold": pd_dict["bold"], "s_italic": pd_dict["italic"]
    })

    current_presets_ass[name] = d

    save_presets_to_file(PRESET_FILE_ASS, current_presets_ass)
    update_all_ass_preset_cbs()
    messagebox.showinfo("成功", f"预设 [{name}] 已提取并保存！\n\n该预设已同步至所有编辑界面，且包含了完整的透明度、阴影等 13 项样式参数。")

ttk.Button(tab_ext, text="💾 提取并保存为全局 ASS 预设", command=save_ext_preset, style='TButton').grid(row=3, column=0, columnspan=3, pady=20, ipadx=20, ipady=5)

root.update_idletasks()
try:
    fonts = list(tkfont.families())
    n_cb['values'] = fonts
    s_cb['values'] = fonts
    cb_m0['values'] = fonts     # 新增
    cb_m2n['values'] = fonts
    cb_m2s['values'] = fonts
    cb_msn['values'] = fonts
    cb_mss['values'] = fonts
    
    cb_ref_font_5['values'] = fonts
    cb_ref_font_9['values'] = fonts
    cb_m0_ref_font['values'] = fonts  # 新增
    cb_m2_ref_font['values'] = fonts
    
    cb_m7['values'] = fonts
    cb_m7_ref_font['values'] = fonts
    
    cb_m8['values'] = fonts
    cb_m8_ref_font['values'] = fonts
except: pass

update_m0_ui()
update_m8_ui()
update_m2_ui()
update_ass_style_mode_5()
update_ms_style_mode_9()

root.mainloop()