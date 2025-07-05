import os
import json
import uuid
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import shutil
from datetime import datetime, timedelta


ITEMS_PER_PAGE = 10
CONFIG_PATH = "data/database_config.json"
LINKS_FOLDER = "links"
os.makedirs(LINKS_FOLDER, exist_ok=True)


class DataManager:
    def __init__(self, config):
        self.config = config
        self.data = {} # 各資料庫名稱對應的 DataFrame
        self.templates = {}
        self.groups = {}
        self.load_all()

    def load_all(self):
        for db_name, path in self.config.items():
            if os.path.exists(path):
                self.data[db_name] = pd.read_excel(path)
            else:
                self.data[db_name] = pd.DataFrame()

            # 模板與群組設定
            template_path = f"data/templates_{db_name}.json"
            group_path = f"data/groups_{db_name}.json"

            if os.path.exists(template_path):
                with open(template_path, "r", encoding="utf-8") as f:
                    self.templates[db_name] = json.load(f)
            else:
                self.templates[db_name] = list(self.data[db_name].columns)

            if os.path.exists(group_path):
                with open(group_path, "r", encoding="utf-8") as f:
                    self.groups[db_name] = json.load(f)
            else:
                self.groups[db_name] = {}

    def save_data(self, db_name):
        path = self.config[db_name]
        self.data[db_name].to_excel(path, index=False)

    def save_templates(self, db_name):
        with open(f"data/templates_{db_name}.json", "w", encoding="utf-8") as f:
            json.dump(self.templates[db_name], f, ensure_ascii=False, indent=2)

    def save_groups(self, db_name):
        with open(f"data/groups_{db_name}.json", "w", encoding="utf-8") as f:
            json.dump(self.groups[db_name], f, ensure_ascii=False, indent=2)

class App:
    def __init__(self, root):
        self.edit_mode = tk.BooleanVar(value=False)
        self.data_edit_mode = tk.BooleanVar(value=False)
        self.root = root
        self.root.title("資料管理系統主頁")
        self.database_config = self.load_config()
        self.data_manager = DataManager(self.database_config)
        self.summary_fields = {
            "車輛": ["車牌"],
            "廠商": ["名稱"]
        }

        self.build_home_page()

    def build_export_page(self):
        self.clear_window()
        tk.Label(self.root, text="請選擇要匯出的資料庫", font=("Arial", 14)).pack(pady=10)
        for db_name in self.database_config:
            tk.Button(self.root, text=db_name, width=20,
                      command=lambda db=db_name: self.build_export_field_selector(db)).pack(pady=5)
        tk.Button(self.root, text="🔙 返回主頁", command=self.build_home_page).pack(pady=20)

    def build_export_field_selector(self, db_name):
        self.clear_window()
        tk.Label(self.root, text=f"選擇要匯出的欄位：{db_name}", font=("Arial", 14)).pack(pady=10)
        self.current_database = db_name
        # ✅ 頂部按鈕列
        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10)

        tk.Button(top_frame, text="📤 匯出資料", command=lambda: self.export_selected_fields(
            db_name,
            [f for f, v in field_vars.items() if v.get() == 1]
        )).pack(side="left", padx=10)

        tk.Button(top_frame, text="🔙 返回", command=self.build_export_page).pack(side="left", padx=10)

        tk.Label(self.root, text=f"選擇要匯出的欄位：{db_name}", font=("Arial", 14)).pack(pady=5)

        # ✅ 中央置中的 canvas 區域
        canvas_frame = tk.Frame(self.root)
        canvas_frame.pack(fill="both", expand=True, padx=20, pady=10)

        canvas = tk.Canvas(canvas_frame)
        canvas.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))
        scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="n")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        group_defs = self.data_manager.groups.get(db_name, {})
        field_vars = {}

        for group_name, fields in group_defs.items():
            wrapper = tk.Frame(scrollable_frame)
            wrapper.pack(pady=5, fill="x")
            group_frame = tk.LabelFrame(wrapper, text=group_name, width=600)
            group_frame.pack(pady=5)
            group_var = tk.IntVar()

            def toggle_group(var=group_var, field_list=fields):
                for f in field_list:
                    if f in field_vars:
                        field_vars[f].set(var.get())

            tk.Checkbutton(group_frame, text="全選", variable=group_var, command=toggle_group).pack(anchor="w")

            for f in fields:
                field_vars[f] = tk.IntVar()
                tk.Checkbutton(group_frame, text=f, variable=field_vars[f]).pack(anchor="w")

    def export_selected_fields(self, db_name, selected_fields):
        df = self.data_manager.data[db_name]
        export_df = df[selected_fields].copy()
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
        filetypes=[("Excel 檔案", "*.xlsx")],
        initialfile=f"{db_name}_匯出.xlsx")
        if save_path:
            export_df.to_excel(save_path, index=False)
            messagebox.showinfo("匯出成功", f"已匯出至：\n{save_path}")

    def open_database(self, db_name):
        self.current_database = db_name
        self.current_page = 0
        self.build_data_page()

    def toggle_data_edit_mode(self):
        self.data_edit_mode.set(not self.data_edit_mode.get())
        self.build_data_page()

    def build_data_page(self):
        self.clear_window()
        control_frame = tk.Frame(self.root)
        control_frame.pack(pady=5)

        tk.Label(self.root, text=f"{self.current_database} 資料頁", font=("Arial", 14)).pack(pady=10)

        tk.Button(control_frame,
                text="✅ 完成編輯" if self.data_edit_mode.get() else "✏️ 編輯模式",
                command=self.toggle_data_edit_mode).pack(side="left", padx=5)

        if self.data_edit_mode.get():
            tk.Button(control_frame, text="➕ 新增資料", command=self.add_new_entry).pack(side="left", padx=5)

        tk.Button(control_frame, text="🔙 返回資料庫", command=self.open_db_select_page).pack(side="left", padx=5)

        self.grid_frame = tk.Frame(self.root)
        self.grid_frame.pack(padx=10, pady=10)

        self.refresh_grid()

    def load_config(self):
        if not os.path.exists(CONFIG_PATH):
            default_config = {
                "車輛": "data/vehicles.xlsx",
                "廠商": "data/vendors.xlsx"
            }
            os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(default_config, f, ensure_ascii=False, indent=2)
            return default_config
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)

    def save_config(self):
        os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(self.database_config, f, ensure_ascii=False, indent=2)

    def clear_window(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def build_home_page(self):
        self.clear_window()
        self.root.geometry("800x600")
        tk.Label(self.root, text="請選擇功能", font=("Arial", 16)).pack(pady=20)
        tk.Button(self.root, text="✏️ 編輯資料", width=20, height=2, command=self.open_db_select_page).pack(pady=10)
        tk.Button(self.root, text="📤 匯出資料", width=20, height=2, command=self.build_export_page).pack(pady=10)

    def delete_entry(self, index):
        df = self.data_manager.data[self.current_database]
        df.drop(index, inplace=True)
        df.reset_index(drop=True, inplace=True)
        self.data_manager.save_data(self.current_database)
        self.refresh_grid()

    def move_entry(self, index, direction):
        df = self.data_manager.data[self.current_database]
        new_index = index + direction
        if 0 <= new_index < len(df):
            df.iloc[[index, new_index]] = df.iloc[[new_index, index]].values
        self.data_manager.save_data(self.current_database)
        self.refresh_grid()

    def add_new_entry(self):
        df = self.data_manager.data[self.current_database]
        if df.empty and df.columns.empty:
            messagebox.showwarning("欄位未定義", f"「{self.current_database}」尚未設定任何欄位，請先編輯欄位模板或手動加入資料後再使用新增功能。")
            return
        new_row = {col: "" for col in df.columns}
        df.loc[len(df)] = new_row
        self.data_manager.save_data(self.current_database)
        self.refresh_grid()

    def refresh_grid(self):
        
        for widget in self.grid_frame.winfo_children():
            widget.destroy()

        df = self.data_manager.data[self.current_database]
        start = self.current_page * ITEMS_PER_PAGE
        end = start + ITEMS_PER_PAGE
        
        for display_index, idx in enumerate(df.index[start:end]):
            
            row = df.loc[idx]
            uuid_str = str(row.get("UUID", ""))
            highlight = self.should_highlight(uuid_str)
            frame = tk.Frame(self.grid_frame, bd=1, relief="solid", width=250, height=130,
                     bg="#ffffcc" if highlight else None)  # 黃色背景
            frame.pack_propagate(False)
            frame.grid(row=display_index // 2, column=display_index % 2, padx=5, pady=10)
            label_fields = self.summary_fields.get(self.current_database)
            if not label_fields:
                label_fields = [col for col in df.columns if col != "UUID"][:2]
            summary_lines = [f"{col}: {row.get(col, '')}" for col in label_fields]
            label_text = f"{self.current_database} #{idx + 1}\n" + "\n".join(summary_lines)
            tk.Label(frame, text=label_text, justify="left", bg="#ffffcc" if highlight else None).pack()
            if self.data_edit_mode.get():
                tk.Button(frame, text="🗑 刪除", command=lambda i=idx: self.delete_entry(i)).pack()
                if display_index > 0:
                    tk.Button(frame, text="↑", command=lambda i=idx: self.move_entry(i, -1)).pack()
                if display_index < len(df) - 1:
                    tk.Button(frame, text="↓", command=lambda i=idx: self.move_entry(i, 1)).pack()
            else:
                tk.Button(frame, text="查看詳情", command=lambda i=idx: self.open_detail(i)).pack()

    def open_db_select_page(self):
        self.clear_window()
        tk.Label(self.root, text="請選擇資料庫", font=("Arial", 14)).pack(pady=10)

        top_frame = tk.Frame(self.root)
        top_frame.pack()

        tk.Button(top_frame, text="✏️ 編輯" if not self.edit_mode.get() else "✅ 完成", command=self.toggle_edit_mode).pack(side="left", padx=5)
        if self.edit_mode.get():
            tk.Button(top_frame, text="➕ 新增資料庫", command=self.create_new_database).pack(side="left", padx=5)

        self.db_frame = tk.Frame(self.root)
        self.db_frame.pack(pady=10)

        self.render_db_list()

        tk.Button(self.root, text="🔙 返回首頁", command=self.build_home_page).pack(pady=5)

    def toggle_edit_mode(self):
        self.edit_mode.set(not self.edit_mode.get())
        self.open_db_select_page()

    def delete_database(self, name):
        if not messagebox.askyesno("確認刪除", f"確定要刪除「{name}」？此動作將永久移除對應資料、模板與分組設定。"):
            return

        # 1. 刪除 Excel 檔案
        excel_path = self.database_config.get(name)
        if excel_path and os.path.exists(excel_path):
            try:
                os.remove(excel_path)
            except Exception as e:
                messagebox.showwarning("刪除失敗", f"無法刪除 Excel 檔案：{e}")

        # 2. 刪除模板
        if name in self.data_manager.templates:
            del self.data_manager.templates[name]
            self.data_manager.save_templates()

        # 3. 刪除分組設定
        if name in self.data_manager.groups:
            del self.data_manager.groups[name]
            self.data_manager.save_groups()

        # 4. 刪除記憶體中的資料
        if name in self.data_manager.data:
            del self.data_manager.data[name]

        # 5. 刪除 config 記錄
        self.database_config.pop(name)
        self.save_config()

        # 6. 重新載入資料庫列表
        self.render_db_list()

    def move_database(self, index, direction):
        names = list(self.database_config.items())
        new_index = index + direction
        if 0 <= new_index < len(names):
            names[index], names[new_index] = names[new_index], names[index]
            self.database_config = dict(names)
            self.save_config()
            self.render_db_list()

    def render_db_list(self):
        for widget in self.db_frame.winfo_children():
            widget.destroy()

        db_names = list(self.database_config.keys())

        for idx, db_name in enumerate(db_names):
            row = tk.Frame(self.db_frame)
            row.pack(fill="x", pady=2)

            tk.Button(row, text=db_name, width=25,
                    command=lambda name=db_name: self.open_database(name)).pack(side="left")

            if self.edit_mode.get():
                tk.Button(row, text="🗑", command=lambda name=db_name: self.delete_database(name)).pack(side="left", padx=2)
                tk.Button(row, text="↑", command=lambda i=idx: self.move_database(i, -1)).pack(side="left", padx=2)
                tk.Button(row, text="↓", command=lambda i=idx: self.move_database(i, 1)).pack(side="left", padx=2)

    def create_new_database(self):
        new_win = tk.Toplevel(self.root)
        new_win.title("新增資料庫")

        tk.Label(new_win, text="資料庫名稱：").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        name_var = tk.StringVar()
        tk.Entry(new_win, textvariable=name_var, width=25).grid(row=0, column=1, pady=5)

        tk.Label(new_win, text="Excel 檔名：").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        file_var = tk.StringVar()
        tk.Entry(new_win, textvariable=file_var, width=25).grid(row=1, column=1, pady=5)

        def save():
            name = name_var.get().strip()
            if not name:
                messagebox.showerror("錯誤", "名稱不可為空")
                return
            if name in self.database_config:
                messagebox.showerror("錯誤", "名稱已存在")
                return

            # 設定檔案名稱與路徑
            filename = file_var.get().strip()
            if not filename:
                filename = name + ".xlsx"
            if not filename.endswith(".xlsx"):
                filename += ".xlsx"
            save_path = os.path.join("data", filename)

            # 建立空檔案（如果不存在）
            if not os.path.exists("data"):
                os.makedirs("data")

            default_columns = ["標題1"]
            df = pd.DataFrame(columns=default_columns)
            df.to_excel(save_path, index=False)

            # 註冊到 config
            self.database_config[name] = save_path
            self.save_config()

            # 寫入 templates 與 groups
            self.data_manager.templates[name] = default_columns
            self.data_manager.groups[name] = {"組別1": ["標題1"]}
            self.data_manager.save_templates(name)
            self.data_manager.save_groups(name)

            # 初始化記憶中的 data
            self.data_manager.data[name] = df

            new_win.destroy()
            self.open_db_select_page()

        tk.Button(new_win, text="建立資料庫", command=save).grid(row=2, column=0, columnspan=2, pady=10)

    def export_page_placeholder(self):
        messagebox.showinfo("尚未實作", "匯出頁面尚未完成，之後會加入欄位選擇與儲存功能。")

    def should_highlight(self, uuid_str):
        period_path = f"period/{uuid_str}_period_1.xlsx"
        if not os.path.exists(period_path):
            return False
        try:
            df = pd.read_excel(period_path)
            today = datetime.today()
            for _, row in df.iterrows():
                next_exec_str = str(row.get("下次執行日期", ""))
                remind_months = str(row.get("執行前__月提醒", ""))
                if not next_exec_str or not remind_months:
                    continue
                try:
                    next_date = datetime.strptime(next_exec_str, "%Y-%m-%d")
                    remind_delta = timedelta(days=int(remind_months) * 30)
                    if today >= next_date - remind_delta:
                        return True
                except:
                    continue
        except:
            return False
        return False

    def open_detail(self, index):
        
        if hasattr(self, 'current_detail_window') and self.current_detail_window.winfo_exists():
            self.current_detail_window.destroy()

        self.current_detail_index = index
        top = tk.Toplevel(self.root)
        self.current_detail_window = top
        top.title(f"{self.current_database} 詳細資料")
        top.geometry("850x650")

        is_editing = tk.BooleanVar(value=False)
        df = self.data_manager.data[self.current_database]
        row = df.loc[index]

        
        if 'UUID' not in df.columns:
            df['UUID'] = None
        if not pd.notnull(df.at[index, 'UUID']):
            df.at[index, 'UUID'] = str(uuid.uuid4())
        uuid_str = df.at[index, 'UUID']

        self.period_data = []
        self.period_path = os.path.join("period", f"{uuid_str}_period_1.xlsx")

        uuid_frame = tk.Frame(top)
        uuid_frame.pack(anchor="w", padx=10, pady=2)
        tk.Label(uuid_frame, text="UUID:").pack(side="left")
        uuid_entry = tk.Entry(uuid_frame, width=40)
        uuid_entry.insert(0, uuid_str)
        uuid_entry.config(state="readonly")
        uuid_entry.pack(side="left", padx=5)

        def copy_uuid():
            top.clipboard_clear()
            top.clipboard_append(uuid_str)
            top.update()
        tk.Button(uuid_frame, text="複製", command=copy_uuid).pack(side="left")

        button_frame = tk.Frame(top)
        button_frame.pack(pady=10)

        container = tk.Frame(top)
        canvas = tk.Canvas(container)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        container.pack(fill="both", expand=True)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))

        editable_groups = []

        def save_changes():
            new_fields = []
            new_groups = {}
            for group in editable_groups:
                group_name = group["title_var"].get().strip()
                if not group_name:
                    continue
                new_groups[group_name] = []
                for field_obj in group["fields"]:
                    key = field_obj["key_var"].get().strip()
                    val = field_obj["val_var"].get().strip()
                    if key:
                        new_fields.append(key)
                        new_groups[group_name].append(key)
                        if field_obj.get("type") == "external_link":
                            label = field_obj.get("label_var", tk.StringVar()).get().strip()
                            df.at[index, key] = json.dumps({"label": label, "path": val})
                        elif field_obj.get("type") == "internal_link":
                            label = field_obj.get("label_var", tk.StringVar()).get().strip()
                            df.at[index, key] = json.dumps({"label": label, "uuid": val})
                        else:
                            col_dtype = df[key].dtype if key in df.columns else object
                            try:
                                if pd.api.types.is_numeric_dtype(col_dtype):
                                    df.at[index, key] = str(val) if val else None
                                elif pd.api.types.is_bool_dtype(col_dtype):
                                    df.at[index, key] = val.lower() in ["true", "1", "yes"]
                                else:
                                    df.at[index, key] = str(val)
                            except ValueError:
                                df.at[index, key] = str(val)  # fallback
            # 儲存週期表格
            try:
                if hasattr(self, "period_data") and self.period_data:
                    
                    rows = []
                    for row_vars in self.period_data:
                        row = [v.get() for v in row_vars]
                        try:
                            if row[1] and row[3]:  # 有填寫「下次間隔」與「此次執行日期」
                                months = int(row[1])
                                exec_date = datetime.strptime(row[3], "%Y-%m-%d")
                                next_date = exec_date + timedelta(days=months * 30)
                                row[4] = next_date.strftime("%Y-%m-%d")
                            else:
                                row[4] = ""
                        except Exception:
                            row[4] = ""
                        rows.append(row)

                    df_period = pd.DataFrame(rows, columns=["標題", "下次間隔__月", "執行前__月提醒", "此次執行日期", "下次執行日期"])
                    os.makedirs(os.path.dirname(self.period_path), exist_ok=True)
                    df_period.to_excel(self.period_path, index=False)
            except Exception as e:
                print("儲存週期表格失敗:", e)

                df_period = pd.DataFrame(rows, columns=["標題", "下次間隔__月", "執行前__月提醒", "此次執行日期", "下次執行日期"])
                os.makedirs(os.path.dirname(self.period_path), exist_ok=True)
                df_period.to_excel(self.period_path, index=False)

            # 📝 儲存異動紀錄
            try:
                changes_path = os.path.join("data", f"changes_{self.current_database}.xlsx")
                os.makedirs("data", exist_ok=True)

                # 載入或初始化
                if os.path.exists(changes_path):
                    df_changes = pd.read_excel(changes_path)
                else:
                    df_changes = pd.DataFrame(columns=["標題", "異動日期", "異動前", "異動後", "uuid"])

                # 使用者輸入
                title = self.change_title_var.get().strip()
                after = self.change_after_var.get().strip()
                if title and after:
                    prev_rows = df_changes[df_changes["uuid"] == uuid_str]
                    prev_after = prev_rows["異動後"].iloc[-1] if not prev_rows.empty else "無"
                    now = datetime.today().strftime("%Y-%m-%d")

                    new_row = {
                        "標題": title,
                        "異動日期": now,
                        "異動前": prev_after,
                        "異動後": after,
                        "uuid": uuid_str
                    }
                    df_changes.loc[len(df_changes)] = new_row
                    df_changes.to_excel(changes_path, index=False)
            except Exception as e:
                print("異動紀錄儲存失敗：", e)
                        

            self.data_manager.templates[self.current_database] = list(dict.fromkeys(new_fields))
            self.data_manager.groups[self.current_database] = new_groups
            self.data_manager.save_templates(self.current_database)
            self.data_manager.save_groups(self.current_database)
            self.data_manager.save_data(self.current_database)

        def render_detail():
            df = self.data_manager.data[self.current_database]

            if is_editing.get():
                save_button.pack(side="left", padx=5)
                edit_button.pack_forget()
            else:
                save_button.pack_forget()
                edit_button.pack(side="left", padx=5)

            for widget in scrollable_frame.winfo_children():
                widget.destroy()

            row = df.loc[index]

            if not is_editing.get():
                groups = self.data_manager.groups.get(self.current_database, {})
                for group_name, fields in groups.items():
                    group_frame = tk.LabelFrame(scrollable_frame, text=group_name, padx=5, pady=5)
                    group_frame.pack(fill="x", padx=10, pady=5)

                    # 建立容器放置欄位內容，預設展開
                    content_frame = tk.Frame(group_frame)
                    content_frame.pack(fill="x")

                    def make_toggle_callback(cf=content_frame):
                        def toggle():
                            if cf.winfo_ismapped():
                                cf.pack_forget()
                            else:
                                cf.pack(fill="x")
                        return toggle

                    tk.Button(group_frame, text="展開 / 收起", command=make_toggle_callback()).pack(anchor="e")

                    # 欄位內容區域
                    for field in fields:
                        if not field:
                            continue
                        row_frame = tk.Frame(content_frame)
                        row_frame.pack(fill="x", pady=2)
                        tk.Label(row_frame, text=field, width=20, anchor="w").pack(side="left")
                        val = row.get(field, "")
                        try:
                            val_obj = json.loads(val)
                            if isinstance(val_obj, dict):
                                if "label" in val_obj and "path" in val_obj:
                                    def open_file(path=val_obj["path"]):
                                        import os, platform, subprocess
                                        if platform.system() == "Windows":
                                            os.startfile(path)
                                        elif platform.system() == "Darwin":
                                            subprocess.call(["open", path])
                                        else:
                                            subprocess.call(["xdg-open", path])
                                    tk.Button(row_frame, text=val_obj["label"], fg="blue", cursor="hand2", command=open_file).pack(side="left", padx=5)
                                elif "label" in val_obj and "uuid" in val_obj:
                                    def open_internal(uuid=val_obj["uuid"]):
                                        target_df = self.data_manager.data[self.current_database]
                                        if "UUID" not in target_df.columns:
                                            messagebox.showwarning("錯誤", "尚未建立 UUID 欄位")
                                            return
                                        target_rows = target_df[target_df["UUID"] == uuid]
                                        if not target_rows.empty:
                                            target_index = target_rows.index[0]
                                            self.open_detail(target_index)
                                        else:
                                            messagebox.showwarning("找不到資料", f"找不到 UUID 為 {uuid} 的資料")
                                    tk.Button(row_frame, text=val_obj["label"], fg="blue", cursor="hand2", command=open_internal).pack(side="left", padx=5)
                                else:
                                    raise ValueError
                            else:
                                raise ValueError
                        except Exception:
                            tk.Label(row_frame, text=str(val), anchor="w", width=40).pack(side="left", padx=5)
                        
                
                def create_new_table(callback=None):
                    table_folder = "tables"
                    os.makedirs(table_folder, exist_ok=True)
                    base = f"{uuid_str}_table_"
                    existing = [f for f in os.listdir(table_folder) if f.startswith(base)]
                    ids = [int(f.split("_")[-1].split(".")[0]) for f in existing if f.split("_")[-1].split(".")[0].isdigit()]
                    new_id = max(ids, default=0) + 1
                    new_path = os.path.join(table_folder, f"{base}{new_id}.xlsx")
                    df = pd.DataFrame([["欄位1", "欄位2"], ["內容1", "內容2"]])
                    title_df = pd.DataFrame({"title": [f"新表格{new_id}"]})

                    with pd.ExcelWriter(new_path, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, header=False, sheet_name="data")
                        title_df.to_excel(writer, index=False, sheet_name="metadata")

                    if callback:
                        callback()

                def open_table_editor(table_path, table_frame, refresh_callback=None):
                    def save_table():
                        try:
                            data = [[var.get() for var in row_vars] for row_vars in row_entries]
                            if not data or not data[0]:
                                messagebox.showwarning("警告", "表格內容為空，請至少保留一列一欄")
                                return

                            df = pd.DataFrame(data[1:], columns=data[0])
                            title_df = pd.DataFrame({"title": [title_var.get()]})

                            os.makedirs(os.path.dirname(table_path), exist_ok=True)
                            with pd.ExcelWriter(table_path, engine="openpyxl", mode="w") as writer:
                                df.to_excel(writer, index=False, sheet_name="data")
                                title_df.to_excel(writer, index=False, sheet_name="metadata")
                            messagebox.showinfo("成功", "表格已儲存")
                        except Exception as e:
                            messagebox.showerror("錯誤", f"儲存失敗：{e}")

                    edit_win = tk.Toplevel()
                    edit_win.title("編輯表格")
                    edit_win.geometry("1000x600")

                    try:
                        with pd.ExcelFile(table_path) as xls:
                            df = pd.read_excel(xls, sheet_name="data", header=None)
                            title_df = pd.read_excel(xls, sheet_name="metadata") if "metadata" in xls.sheet_names else pd.DataFrame()
                    except Exception as e:
                        messagebox.showerror("錯誤", f"讀取失敗：{e}")
                        return

                    title_var = tk.StringVar(value=title_df["title"].iloc[0] if not title_df.empty else "新表格")
                    tk.Entry(edit_win, textvariable=title_var, font=("Arial", 12)).pack(pady=5)

                    outer_frame = tk.Frame(edit_win)
                    outer_frame.pack(fill="both", expand=True)

                    xscroll = tk.Scrollbar(outer_frame, orient="horizontal")
                    yscroll = tk.Scrollbar(outer_frame, orient="vertical")
                    canvas = tk.Canvas(outer_frame, xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
                    xscroll.config(command=canvas.xview)
                    yscroll.config(command=canvas.yview)

                    xscroll.pack(side="bottom", fill="x")
                    yscroll.pack(side="right", fill="y")
                    canvas.pack(side="left", fill="both", expand=True)

                    inner_frame = tk.Frame(canvas)
                    canvas.create_window((0, 0), window=inner_frame, anchor="nw")
                    inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

                    row_entries = [[tk.StringVar(value=str(cell)) for cell in row] for _, row in df.iterrows()]
                    if not row_entries:
                        row_entries.append([tk.StringVar(value="欄位1"), tk.StringVar(value="欄位2")])
                    header = list(df.columns) if not df.empty else ["欄位1", "欄位2"]
                    row_entries.insert(0, [tk.StringVar(value=str(col)) for col in header])

                    def refresh_table_ui():
                        for widget in inner_frame.winfo_children():
                            widget.destroy()

                        for r_idx, row_vars in enumerate(row_entries):
                            row_frame = tk.Frame(inner_frame)
                            row_frame.pack(fill="x", pady=2)
                            for c_idx, var in enumerate(row_vars):
                                cell_frame = tk.Frame(row_frame)
                                cell_frame.pack(side="left")
                                tk.Entry(cell_frame, textvariable=var, width=15).pack()
                                if r_idx == 0:
                                    tk.Button(cell_frame, text="🗑 刪除列", command=lambda idx=c_idx: delete_column(idx)).pack()
                            if r_idx > 0:
                                tk.Button(row_frame, text="🗑 刪除行", command=lambda idx=r_idx: delete_row(idx)).pack(side="left", padx=2)

                    def add_column():
                        for row_vars in row_entries:
                            row_vars.append(tk.StringVar(value=""))
                        refresh_table_ui()

                    def delete_column(col_idx):
                        if row_entries and len(row_entries[0]) > 1:
                            for row_vars in row_entries:
                                if col_idx < len(row_vars):
                                    row_vars.pop(col_idx)
                            refresh_table_ui()

                    def add_row():
                        new_row = [tk.StringVar(value="") for _ in row_entries[0]]
                        row_entries.append(new_row)
                        refresh_table_ui()

                    def delete_row(idx):
                        if 0 < idx < len(row_entries):
                            row_entries.pop(idx)
                            refresh_table_ui()

                    def close_editor():
                        if refresh_callback:
                            refresh_callback()
                        edit_win.destroy()

                    refresh_table_ui()

                    action_frame = tk.Frame(edit_win)
                    action_frame.pack(pady=10)
                    tk.Button(action_frame, text="➕ 新增行", command=add_row).pack(side="left", padx=5)
                    tk.Button(action_frame, text="➕ 新增列", command=add_column).pack(side="left", padx=5)
                    tk.Button(action_frame, text="💾 儲存表格", command=save_table).pack(side="left", padx=5)
                    tk.Button(action_frame, text="❌ 關閉視窗", command=close_editor).pack(side="left", padx=5)

                def refresh_tables():
                    for widget in tables_container.winfo_children():
                        widget.destroy()
                    table_folder = "tables"
                    os.makedirs(table_folder, exist_ok=True)
                    for f in sorted(os.listdir("tables")):
                        if f.startswith(f"{uuid_str}_table_") and f.endswith(".xlsx"):
                            table_path = os.path.join("tables", f)
                            try:
                                meta = pd.read_excel(table_path, sheet_name="metadata")
                                title = meta.at[0, "title"] if "title" in meta.columns else f
                            except Exception:
                                title = f

                            frame = tk.LabelFrame(tables_container, text=title)
                            frame.pack(fill="x", padx=10, pady=5)

                            btn_row = tk.Frame(frame)
                            btn_row.pack(anchor="w", padx=5, pady=5)

                            is_expanded = tk.BooleanVar(value=False)
                            content_frame = tk.Frame(frame)
                            content_frame.pack(fill="x")

                            def toggle_expand(p=table_path, cf=content_frame, v=is_expanded):
                                if v.get():  # 如果已展開 ➜ 摺疊
                                    for widget in cf.winfo_children():
                                        widget.destroy()
                                    v.set(False)
                                else:  # 尚未展開 ➜ 展開
                                    try:
                                        df = pd.read_excel(p, sheet_name="data", header=None)
                                        for r_idx, row in df.iterrows():
                                            for c_idx, cell in enumerate(row):
                                                tk.Label(cf, text=str(cell), width=15, anchor="w", relief="groove").grid(row=r_idx, column=c_idx, sticky="nsew", padx=1, pady=1)
                                        v.set(True)
                                    except Exception as e:
                                        tk.Label(cf, text=f"讀取失敗: {e}", fg="red").pack()
                                        v.set(True)

                            def delete_table(path=table_path):
                                if messagebox.askyesno("刪除表格", "確定要刪除此表格？此操作不可復原。"):
                                    try:
                                        os.remove(path)
                                        refresh_tables()
                                    except Exception as e:
                                        messagebox.showerror("刪除失敗", f"無法刪除表格：{e}")

                            tk.Button(btn_row, text="展開/摺疊", command=toggle_expand).pack(side="left", padx=2)
                            tk.Button(btn_row, text="✏️ 編輯", command=lambda p=table_path, fr=frame: open_table_editor(p, fr, refresh_tables)).pack(side="left", padx=2)
                            tk.Button(btn_row, text="🗑 刪除", command=delete_table).pack(side="left", padx=2)

                # 📅 週期表格顯示（只讀模式）
                tk.Label(scrollable_frame, text="📅 週期表格", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)

                period_folder = "period"
                os.makedirs(period_folder, exist_ok=True)
                period_files = sorted(f for f in os.listdir(period_folder) if f.startswith(f"{uuid_str}_period_") and f.endswith(".xlsx"))

                for f in period_files:
                    period_path = os.path.join(period_folder, f)
                    try:
                        df = pd.read_excel(period_path)
                    except Exception as e:
                        df = pd.DataFrame([["讀取失敗", str(e)]])
                    
                    if df.empty:
                        continue

                    box = tk.LabelFrame(scrollable_frame, padx=5, pady=5)
                    box.pack(fill="x", padx=10, pady=5)

                    # 📌 顯示一次欄位標題
                    header_frame = tk.Frame(box)
                    header_frame.pack(fill="x", pady=2)
                    for col in df.columns:
                        tk.Label(header_frame, text=col, width=20, anchor="center", font=("Arial", 9, "bold")).pack(side="left", padx=2)

                    # 📌 顯示所有資料行
                    for _, row_ in df.iterrows():
                        row_frame = tk.Frame(box)
                        row_frame.pack(fill="x", pady=1)
                        for col in df.columns:
                            val = str(row_.get(col, ""))
                            tk.Label(row_frame, text=val, width=20, anchor="center").pack(side="left", padx=2)

                # 📝 異動紀錄顯示（只讀模式）
                tk.Label(scrollable_frame, text="📝 異動紀錄", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)

                changes_path = os.path.join("data", f"changes_{self.current_database}.xlsx")
                if os.path.exists(changes_path):
                    try:
                        df_changes = pd.read_excel(changes_path)
                        df_changes = df_changes[df_changes["uuid"] == uuid_str]
                    except Exception as e:
                        df_changes = pd.DataFrame([{"標題": "讀取失敗", "異動日期": str(e), "異動前": "", "異動後": ""}])
                else:
                    df_changes = pd.DataFrame()

                if not df_changes.empty:
                    frame = tk.Frame(scrollable_frame)
                    frame.pack(fill="x", padx=10, pady=5)

                    # 顯示欄位標題列
                    header = tk.Frame(frame)
                    header.pack(fill="x", pady=2)
                    for col in ["標題", "異動日期", "異動前", "異動後"]:
                        tk.Label(header, text=col, width=20, anchor="center", font=("Arial", 9, "bold")).pack(side="left", padx=2)

                    # 顯示每一筆紀錄
                    for _, row in df_changes.iterrows():
                        row_frame = tk.Frame(frame)
                        row_frame.pack(fill="x", pady=1)
                        for col in ["標題", "異動日期", "異動前", "異動後"]:
                            val = str(row.get(col, ""))
                            tk.Label(row_frame, text=val, width=20, anchor="center").pack(side="left", padx=2)
                else:
                    tk.Label(scrollable_frame, text="尚無異動紀錄", fg="gray").pack(anchor="w", padx=15, pady=5)
                
                # 📑 顯示模式下顯示自由表格
                label_frame = tk.Frame(scrollable_frame)
                label_frame.pack(anchor="w", padx=10, pady=5)
                tk.Label(label_frame, text="📑 自由表格", font=("Arial", 12, "bold")).pack(side="left")
                tk.Button(label_frame, text="➕ 新增表格", command=lambda: create_new_table(refresh_tables)).pack(side="left", padx=5)

                tables_container = tk.Frame(scrollable_frame)
                tables_container.pack(fill="x")

                refresh_tables()

                return

            def save_and_reload():
                save_changes()
                top.destroy()
                self.open_detail(index)

            def add_group():
                group_data = {"title_var": tk.StringVar(value="新組別"), "fields": []}
                editable_groups.append(group_data)
                render_detail()

            def move_group_up(idx):
                if idx > 0:
                    editable_groups[idx], editable_groups[idx - 1] = editable_groups[idx - 1], editable_groups[idx]
                    render_detail()

            def move_group_down(idx):
                if idx < len(editable_groups) - 1:
                    editable_groups[idx], editable_groups[idx + 1] = editable_groups[idx + 1], editable_groups[idx]
                    render_detail()

            def make_delete_callback(local_fields, f, rf):
                return lambda: (rf.destroy(), local_fields.remove(f))

            if not editable_groups:
                groups = self.data_manager.groups.get(self.current_database, {})
                for group_name, field_list in groups.items():
                    group_data = {"title_var": tk.StringVar(value=group_name), "fields": []}
                    for f in field_list:
                        key_var = tk.StringVar(value=f)
                        val_raw = row.get(f, "")
                        try:
                            val_obj = json.loads(val_raw)
                            if isinstance(val_obj, dict) and "label" in val_obj and "path" in val_obj:
                                label_var = tk.StringVar(value=val_obj["label"])
                                val_var = tk.StringVar(value=val_obj["path"])
                                field_data = {
                                    "key_var": key_var,
                                    "val_var": val_var,
                                    "label_var": label_var,
                                    "type": "external_link"
                                }
                            elif isinstance(val_obj, dict) and "label" in val_obj and "uuid" in val_obj:
                                label_var = tk.StringVar(value=val_obj["label"])
                                val_var = tk.StringVar(value=val_obj["uuid"])
                                field_data = {
                                    "key_var": key_var,
                                    "val_var": val_var,
                                    "label_var": label_var,
                                    "type": "internal_link"
                                }
                            else:
                                raise ValueError
                        except Exception:
                            val_var = tk.StringVar(value=str(val_raw))
                            field_data = {"key_var": key_var, "val_var": val_var}
                        
                        # ✅ 加入每個欄位到當前 group 中
                        group_data["fields"].append(field_data)
                    editable_groups.append(group_data)

            for i, group_data in enumerate(editable_groups):
                group_frame = tk.LabelFrame(scrollable_frame, text="", padx=5, pady=5)
                group_frame.pack(fill="x", padx=10, pady=5)

                header = tk.Frame(group_frame)
                header.pack(fill="x")

                tk.Entry(header, textvariable=group_data["title_var"], font=("Arial", 10, "bold"), width=30).pack(side="left")
                if i > 0:
                    tk.Button(header, text="↑", width=2, command=lambda idx=i: move_group_up(idx)).pack(side="right", padx=2)
                if i < len(editable_groups) - 1:
                    tk.Button(header, text="↓", width=2, command=lambda idx=i: move_group_down(idx)).pack(side="right", padx=2)
                tk.Button(header, text="刪除分組", width=6, command=lambda idx=i: (editable_groups.pop(idx), render_detail())).pack(side="right", padx=5)

                content_frame = tk.Frame(group_frame)
                content_frame.pack(fill="x")

                def add_field(local_fields=group_data["fields"], local_frame=content_frame):
                    key_var = tk.StringVar()
                    val_var = tk.StringVar()
                    row_frame = tk.Frame(local_frame)
                    row_frame.pack(fill="x", pady=2)
                    tk.Entry(row_frame, textvariable=key_var, width=20).pack(side="left")
                    tk.Entry(row_frame, textvariable=val_var, width=40).pack(side="left", padx=5)
                    field_obj = {"key_var": key_var, "val_var": val_var}
                    tk.Button(row_frame, text="刪除", command=make_delete_callback(local_fields, field_obj, row_frame)).pack(side="left")
                    local_fields.append(field_obj)

                for field_obj in group_data["fields"]:
                    row_frame = tk.Frame(content_frame)
                    row_frame.pack(fill="x", pady=2)
                    tk.Entry(row_frame, textvariable=field_obj["key_var"], width=20).pack(side="left")

                    if field_obj.get("type") == "external_link":
                        tk.Entry(row_frame, textvariable=field_obj["label_var"], width=20).pack(side="left", padx=5)
                        tk.Entry(row_frame, textvariable=field_obj["val_var"], width=30).pack(side="left", padx=5)

                        def browse_file(var=field_obj["val_var"]):
                            file_path = filedialog.askopenfilename()
                            if file_path:
                                var.set(file_path)

                        tk.Button(row_frame, text="選擇檔案", command=browse_file).pack(side="left")

                    elif field_obj.get("type") == "internal_link":
                        tk.Entry(row_frame, textvariable=field_obj["label_var"], width=20).pack(side="left", padx=5)
                        tk.Entry(row_frame, textvariable=field_obj["val_var"], width=30).pack(side="left", padx=5)

                        def paste_uuid(var=field_obj["val_var"]):
                            try:
                                clipboard_text = top.clipboard_get()
                                var.set(clipboard_text)
                            except tk.TclError:
                                pass

                        tk.Button(row_frame, text="貼上 UUID", command=paste_uuid).pack(side="left")

                    else:
                        # 嘗試從 dataframe 抓取所有該欄位曾用過的值
                        key = field_obj["key_var"].get().strip()
                        val_var = field_obj["val_var"]
                        df_all = self.data_manager.data[self.current_database]

                        # 取唯一非空值（字串格式）
                        if key in df_all.columns:
                            choices = sorted(set(str(v) for v in df_all[key].dropna().unique()))
                        else:
                            choices = []

                        combo = ttk.Combobox(row_frame, textvariable=val_var, values=choices, width=37)
                        combo.pack(side="left", padx=5)
                        combo.set(val_var.get())  # 預設值

                    tk.Button(row_frame, text="刪除", command=make_delete_callback(group_data["fields"], field_obj, row_frame)).pack(side="left")

                tk.Button(group_frame, text="新增欄位", command=add_field).pack(side="left", padx=5)

                def add_external_link(local_fields=group_data["fields"], local_frame=content_frame):
                    key_var = tk.StringVar()
                    label_var = tk.StringVar()
                    path_var = tk.StringVar()
                    row_frame = tk.Frame(local_frame)
                    row_frame.pack(fill="x", pady=2)
                    tk.Entry(row_frame, textvariable=key_var, width=20).pack(side="left")
                    tk.Entry(row_frame, textvariable=label_var, width=20).pack(side="left", padx=5)
                    tk.Entry(row_frame, textvariable=path_var, width=30).pack(side="left", padx=5)

                    def browse_file():
                        
                        file_path = filedialog.askopenfilename()
                        if file_path:
                            filename = os.path.basename(file_path)
                            unique_name = f"{uuid.uuid4().hex}_{filename}"
                            dest_path = os.path.join(LINKS_FOLDER, unique_name)
                            try:
                                shutil.copy2(file_path, dest_path)
                                path_var.set(dest_path)
                            except Exception as e:
                                messagebox.showerror("複製失敗", f"無法複製檔案：{e}")

                    tk.Button(row_frame, text="選擇檔案", command=browse_file).pack(side="left")
                    
                    field_obj = {
                        "key_var": key_var,
                        "val_var": path_var,
                        "label_var": label_var,
                        "type": "external_link"
                    }
                    tk.Button(row_frame, text="刪除", command=make_delete_callback(local_fields, field_obj, row_frame)).pack(side="left")
                    local_fields.append(field_obj)

                def add_internal_link(local_fields=group_data["fields"], local_frame=content_frame):
                    key_var = tk.StringVar()
                    label_var = tk.StringVar()
                    target_var = tk.StringVar()

                    row_frame = tk.Frame(local_frame)
                    row_frame.pack(fill="x", pady=2)
                    tk.Entry(row_frame, textvariable=key_var, width=20).pack(side="left")
                    tk.Entry(row_frame, textvariable=label_var, width=20).pack(side="left", padx=5)
                    tk.Entry(row_frame, textvariable=target_var, width=30).pack(side="left", padx=5)

                    def paste_uuid():
                        try:
                            clipboard_text = top.clipboard_get()
                            target_var.set(clipboard_text)
                        except tk.TclError:
                            pass

                    tk.Button(row_frame, text="貼上 UUID", command=paste_uuid).pack(side="left")

                    field_obj = {
                        "key_var": key_var,
                        "val_var": target_var,
                        "label_var": label_var,
                        "type": "internal_link"
                    }
                    tk.Button(row_frame, text="刪除", command=make_delete_callback(local_fields, field_obj, row_frame)).pack(side="left")
                    local_fields.append(field_obj)

                tk.Button(group_frame, text="新增內部連結",
                          command=lambda lf=group_data["fields"],
                          cf=content_frame: add_internal_link(lf, cf)).pack(side="left", padx=5)

                tk.Button(group_frame, text="新增外部連結",
                          command=lambda lf=group_data["fields"],
                          cf=content_frame: add_external_link(lf, cf)).pack(side="left", padx=5)

            tk.Button(scrollable_frame, text="新增分組", command=add_group).pack(pady=10)

            # 📝 異動紀錄編輯區
            tk.Label(scrollable_frame, text="📝 新增異動紀錄", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)

            change_title_var = tk.StringVar()
            change_after_var = tk.StringVar()

            form_frame = tk.Frame(scrollable_frame)
            form_frame.pack(fill="x", padx=15, pady=5)

            tk.Label(form_frame, text="異動標題：", width=10, anchor="e").pack(side="left", padx=2)
            tk.Entry(form_frame, textvariable=change_title_var, width=25).pack(side="left", padx=5)

            tk.Label(form_frame, text="異動後內容：", width=12, anchor="e").pack(side="left", padx=2)
            tk.Entry(form_frame, textvariable=change_after_var, width=40).pack(side="left", padx=5)

            # 傳遞到 save_changes 時使用
            self.change_title_var = change_title_var
            self.change_after_var = change_after_var

            
            
            ### 
            period_folder = "period"
            os.makedirs(period_folder, exist_ok=True)
            period_data = self.period_data
            period_path = self.period_path
            tk.Label(scrollable_frame, text="⏳ 週期表格", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)
            period_frame = tk.Frame(scrollable_frame)
            period_frame.pack(fill="x", padx=10, pady=5)

            tk.Label(period_frame, text="標題", width=20).grid(row=0, column=0)
            tk.Label(period_frame, text="下次間隔__月", width=15).grid(row=0, column=1)
            tk.Label(period_frame, text="執行前__月提醒", width=15).grid(row=0, column=2)
            tk.Label(period_frame, text="此次執行日期", width=20).grid(row=0, column=3)
            tk.Label(period_frame, text="下次執行日期", width=20).grid(row=0, column=4)

            period_data = []  # local reference
            self.period_data = period_data  # ✅ 讓 save_changes() 能存取

            # 讀取既有 period 表格內容
            if os.path.exists(period_path):
                try:
                    df_period = pd.read_excel(period_path)
                    for r_idx, row in df_period.iterrows():
                        row_vars = [tk.StringVar(value=str(row.get(col, ""))) for col in ["標題", "下次間隔__月", "執行前__月提醒", "此次執行日期", "下次執行日期"]]
                        period_data.append(row_vars)
                except Exception as e:
                    print("讀取週期表格失敗：", e)

            # 若無內容，自動加入一列空資料
            if not period_data:
                row_vars = [tk.StringVar() for _ in range(5)]
                period_data.append(row_vars)

            # 顯示每一列輸入欄位
            def render_period_rows():
                for widget in period_frame.winfo_children():
                    if int(widget.grid_info()["row"]) > 0:
                        widget.destroy()

                for r_idx, row_vars in enumerate(period_data):
                    for c_idx, var in enumerate(row_vars):
                        tk.Entry(period_frame, textvariable=var, width=20).grid(row=r_idx+1, column=c_idx)
                    def get_today(var=row_vars[3]):
                        var.set(datetime.today().strftime("%Y-%m-%d"))
                    tk.Button(period_frame, text="今天", command=get_today).grid(row=r_idx+1, column=5)

                    def delete_row(idx=r_idx):
                        period_data.pop(idx)
                        render_period_rows()
                    tk.Button(period_frame, text="刪除", command=delete_row).grid(row=r_idx+1, column=6)
                    # 自動計算欄位 - Label 形式
                    next_exec_label = tk.Label(period_frame, text="", width=20)
                    next_exec_label.grid(row=r_idx+1, column=4)

                    # 計算結果（嘗試將欄位自動更新）
                    def update_next_exec():
                        try:
                            months = int(row_vars[1].get())
                            exec_date = datetime.strptime(row_vars[3].get(), "%Y-%m-%d")
                            next_date = exec_date + timedelta(days=months * 30)
                            next_exec_label.config(text=next_date.strftime("%Y-%m-%d"))
                            row_vars[4].set(next_date.strftime("%Y-%m-%d"))  # 同步更新值
                        except:
                            next_exec_label.config(text="")

                    # 綁定內容變動時自動更新
                    row_vars[1].trace_add("write", lambda *args: update_next_exec())
                    row_vars[3].trace_add("write", lambda *args: update_next_exec())

                    # 初始更新
                    update_next_exec()

            render_period_rows()

            tk.Button(scrollable_frame, text="➕ 新增週期紀錄", command=lambda: (period_data.append([tk.StringVar() for _ in range(5)]), render_period_rows())).pack(padx=10, pady=5, anchor="w")

        def on_close():
            if is_editing.get():
                if messagebox.askyesno("尚未儲存", "尚未儲存變更，確定要關閉嗎？"):
                    top.destroy()
                    self.refresh_grid()
            else:
                top.destroy()
                self.refresh_grid()
        

        def toggle_edit():
            is_editing.set(True)
            save_button.pack(side="left", padx=5)
            edit_button.pack_forget()
            render_detail()

        def save_and_exit_edit():
            if is_editing.get():
                save_changes()
                top.destroy()
                self.open_detail(index)
                self.refresh_grid()

        

        top.protocol("WM_DELETE_WINDOW", on_close)
        save_button = tk.Button(button_frame, text="保存變更並退出編輯", command=save_and_exit_edit)
        edit_button = tk.Button(button_frame, text="編輯模式切換", command=toggle_edit)
        tk.Button(button_frame, text="關閉", command=on_close).pack(side="left", padx=5)

        render_detail()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
