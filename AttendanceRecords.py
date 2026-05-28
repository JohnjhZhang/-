import sys
import os
from datetime import datetime
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLineEdit, QComboBox, QDateEdit,
    QTextEdit, QLabel, QFileDialog, QMessageBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QGroupBox
)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QFont, QIntValidator

# ===================== 全局配置 =====================
VERSION = "v1.0.0"
COPYRIGHT = "©️JohnJHZhang"
EMAIL = "johnJHzhang@outlook.com"

LANG_KEYS = ["zh_CN", "zh_TW", "en", "ja", "ko"]
LANG_NAMES = ["简体中文", "繁體中文", "English", "日本語", "한국어"]
CURRENT_LANG = "zh_CN"

THEME_LIGHT = True
SAVE_PATH = ""
DATA = []
WEEK_LIST = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]

# 多语言翻译（学生人数改为总人数）
TRANSLATIONS = {
    "zh_CN": {"title": "课表记录器", "group_input": "信息录入", "class": "班级名称", "total": "总人数", "course": "课程名称", "date": "上课日期", "week": "星期", "remark": "备注", "status": "上课状态", "attend": "出勤", "absent": "缺勤", "leave": "请假", "add": "添加记录", "clear": "清空数据", "export": "导出Excel", "folder": "选择保存文件夹", "theme": "切换主题", "lang": "语言", "no_data": "暂无数据", "select_folder": "请先选择保存文件夹", "success": "导出成功！", "present": "实到人数", "absent_num": "缺勤人数", "ratio": "缺勤比例"},
    "zh_TW": {"title": "課表記錄器", "group_input": "資訊錄入", "class": "班級名稱", "total": "總人數", "course": "課程名稱", "date": "上課日期", "week": "星期", "remark": "備註", "status": "上課狀態", "attend": "出勤", "absent": "缺勤", "leave": "請假", "add": "新增記錄", "clear": "清空資料", "export": "匯出Excel", "folder": "選擇儲存資料夾", "theme": "切換主題", "lang": "語言", "no_data": "暫無資料", "select_folder": "請先選擇儲存資料夾", "success": "匯出成功！", "present": "實到人數", "absent_num": "缺勤人數", "ratio": "缺勤比例"},
    "en": {"title": "Schedule Recorder", "group_input": "Input Info", "class": "Class Name", "total": "Total", "course": "Course Name", "date": "Class Date", "week": "Weekday", "remark": "Remark", "status": "Status", "attend": "Attend", "absent": "Absent", "leave": "Leave", "add": "Add", "clear": "Clear", "export": "Export Excel", "folder": "Select Path", "theme": "Switch Theme", "lang": "Language", "no_data": "No Data", "select_folder": "Please select save path", "success": "Export Success!", "present": "Present", "absent_num": "Absent", "ratio": "Absent Ratio"},
    "ja": {"title": "時間割レコーダー", "group_input": "情報入力", "class": "クラス名", "total": "総人数", "course": "コース名", "date": "授業日", "week": "曜日", "remark": "備考", "status": "授業状況", "attend": "出席", "absent": "欠席", "leave": "休暇", "add": "追加", "clear": "クリア", "export": "Excel出力", "folder": "保存先選択", "theme": "テーマ切替", "lang": "言語", "no_data": "データなし", "select_folder": "保存先を選択してください", "success": "出力完了！", "present": "出席数", "absent_num": "欠席数", "ratio": "欠席率"},
    "ko": {"title": "시간표 기록기", "group_input": "정보 입력", "class": "반 이름", "total": "총 인원", "course": "강의 이름", "date": "수업 날짜", "week": "요일", "remark": "비고", "status": "수업 상태", "attend": "출석", "absent": "결석", "leave": "휴가", "add": "기록 추가", "clear": "전체 지우기", "export": "Excel 내보내기", "folder": "저장 경로 선택", "theme": "테마 전환", "lang": "언어", "no_data": "데이터 없음", "select_folder": "저장 경로를 선택하세요", "success": "내보내기 성공!", "present": "출석 인원", "absent_num": "결석 인원", "ratio": "결석 비율"}
}

# ===================== 主窗口 =====================
class ScheduleApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.t = TRANSLATIONS[CURRENT_LANG]
        self.setWindowTitle(self.t["title"])
        self.setMinimumSize(1200, 780)
        self.setFont(QFont("SF Pro Text", 12))

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.main_layout = QVBoxLayout(main_widget)
        self.main_layout.setContentsMargins(30, 30, 30, 30)
        self.main_layout.setSpacing(20)

        # 构建界面
        self.build_top_toolbar()
        self.build_input_group()
        self.build_btn_bar()
        self.build_path_label()
        self.build_excel_preview()
        self.build_footer()

        self.refresh_ui()
        self.apply_theme()
        self.bind_auto_calculate()  # 绑定自动计算

    # 顶部：语言+主题
    def build_top_toolbar(self):
        layout = QHBoxLayout()
        self.lang_lab = QLabel(self.t["lang"])
        self.lang_combo = QComboBox()
        self.lang_combo.addItems(LANG_NAMES)
        self.lang_combo.setFixedWidth(220)
        self.lang_combo.currentIndexChanged.connect(self.switch_lang)

        self.theme_btn = QPushButton(self.t["theme"])
        self.theme_btn.clicked.connect(self.switch_theme)
        self.theme_btn.setFixedWidth(130)

        layout.addStretch()
        layout.addWidget(self.lang_lab)
        layout.addWidget(self.lang_combo)
        layout.addSpacing(20)
        layout.addWidget(self.theme_btn)
        self.main_layout.addLayout(layout)

    # 输入区域（总人数+自动互算）
    def build_input_group(self):
        self.group_input = QGroupBox()
        layout = QVBoxLayout(self.group_input)
        layout.setSpacing(18)

        # 第一行：班级 + 总人数 + 课程名称
        row1 = QHBoxLayout()
        self.class_lab = QLabel()
        self.class_edit = QLineEdit()
        self.class_edit.setMinimumWidth(160)
        
        self.total_lab = QLabel()
        self.total_edit = QLineEdit()
        self.total_edit.setValidator(QIntValidator(0, 9999))  # 总人数数字校验
        self.total_edit.setMinimumWidth(120)
        
        self.course_lab = QLabel()
        self.course_edit = QLineEdit()
        self.course_edit.setMinimumWidth(160)
        
        row1.addWidget(self.class_lab)
        row1.addWidget(self.class_edit)
        row1.addSpacing(15)
        row1.addWidget(self.total_lab)
        row1.addWidget(self.total_edit)
        row1.addSpacing(15)
        row1.addWidget(self.course_lab)
        row1.addWidget(self.course_edit)

        # 第二行：日期 + 星期 + 上课状态
        row2 = QHBoxLayout()
        self.date_lab = QLabel()
        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        self.date_edit.setMinimumWidth(140)
        
        self.week_lab = QLabel()
        self.week_combo = QComboBox()
        self.week_combo.addItems(WEEK_LIST)
        self.week_combo.setMinimumWidth(100)
        
        self.status_lab = QLabel()
        self.status_combo = QComboBox()
        self.status_combo.setMinimumWidth(120)
        
        row2.addWidget(self.date_lab)
        row2.addWidget(self.date_edit)
        row2.addSpacing(15)
        row2.addWidget(self.week_lab)
        row2.addWidget(self.week_combo)
        row2.addSpacing(15)
        row2.addWidget(self.status_lab)
        row2.addWidget(self.status_combo)

        # 第三行：实到人数 + 缺勤人数 + 缺勤比例（自动计算）
        row3 = QHBoxLayout()
        self.present_lab = QLabel()
        self.present_edit = QLineEdit()
        self.present_edit.setValidator(QIntValidator(0, 9999))
        self.present_edit.setMinimumWidth(120)
        
        self.absent_num_lab = QLabel()
        self.absent_num_edit = QLineEdit()
        self.absent_num_edit.setValidator(QIntValidator(0, 9999))
        self.absent_num_edit.setMinimumWidth(120)
        
        self.ratio_lab = QLabel()
        self.ratio_show = QLabel("0.00%")
        self.ratio_show.setMinimumWidth(100)
        self.ratio_show.setStyleSheet("font-weight:bold; color:#d32f2f;")
        
        row3.addWidget(self.present_lab)
        row3.addWidget(self.present_edit)
        row3.addSpacing(15)
        row3.addWidget(self.absent_num_lab)
        row3.addWidget(self.absent_num_edit)
        row3.addSpacing(15)
        row3.addWidget(self.ratio_lab)
        row3.addWidget(self.ratio_show)

        # 第四行：备注
        row4 = QHBoxLayout()
        self.remark_lab = QLabel()
        self.remark_edit = QTextEdit()
        self.remark_edit.setMaximumHeight(100)
        row4.addWidget(self.remark_lab)
        row4.addWidget(self.remark_edit)

        layout.addLayout(row1)
        layout.addLayout(row2)
        layout.addLayout(row3)
        layout.addLayout(row4)
        self.main_layout.addWidget(self.group_input)

    # 按钮栏
    def build_btn_bar(self):
        layout = QHBoxLayout()
        layout.setSpacing(15)
        
        self.add_btn = QPushButton()
        self.clear_btn = QPushButton()
        self.export_btn = QPushButton()
        self.folder_btn = QPushButton()

        btn_style = """QPushButton{background:#007AFF;color:white;border-radius:8px;padding:10px 20px;min-height:38px;}QPushButton:hover{background:#0066CC;}"""
        for btn in [self.add_btn, self.clear_btn, self.export_btn, self.folder_btn, self.theme_btn]:
            btn.setStyleSheet(btn_style)

        self.add_btn.clicked.connect(self.add_record)
        self.clear_btn.clicked.connect(self.clear_all)
        self.export_btn.clicked.connect(self.export_excel)
        self.folder_btn.clicked.connect(self.select_folder)

        layout.addWidget(self.add_btn)
        layout.addWidget(self.clear_btn)
        layout.addWidget(self.export_btn)
        layout.addWidget(self.folder_btn)
        layout.addStretch()
        self.main_layout.addLayout(layout)

    def build_path_label(self):
        self.path_lab = QLabel()
        self.path_lab.setStyleSheet("color:#666;")
        self.main_layout.addWidget(self.path_lab)

    # 表格预览
    def build_excel_preview(self):
        self.table = QTableWidget()
        headers = ["班级名称", "总人数", "课程名称", "上课日期", "星期", "上课状态", "实到人数", "缺勤人数", "缺勤比例", "备注"]
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setStyleSheet("QTableWidget{border:1px solid #ddd;border-radius:8px;}")
        self.main_layout.addWidget(self.table)

    def build_footer(self):
        layout = QHBoxLayout()
        layout.addStretch()
        layout.addWidget(QLabel(VERSION))
        layout.addWidget(QLabel(f"{COPYRIGHT} | {EMAIL}"))
        self.main_layout.addLayout(layout)

    # ===================== 核心：自动计算逻辑 =====================
    def bind_auto_calculate(self):
        # 总人数/实到/缺勤 变化时自动计算
        self.total_edit.textChanged.connect(self.auto_calculate)
        self.present_edit.textChanged.connect(self.auto_calculate)
        self.absent_num_edit.textChanged.connect(self.auto_calculate)

    def auto_calculate(self):
        try:
            # 获取数值
            total = int(self.total_edit.text().strip() or 0)
            present_text = self.present_edit.text().strip()
            absent_text = self.absent_num_edit.text().strip()

            # 无总人数时重置
            if total <= 0:
                self.present_edit.clear()
                self.absent_num_edit.clear()
                self.ratio_show.setText("0.00%")
                return

            # 智能互算逻辑
            if present_text:  # 输入实到 → 算缺勤
                present = int(present_text)
                absent = total - present
                self.absent_num_edit.setText(str(max(absent, 0)))  # 不小于0
            elif absent_text:  # 输入缺勤 → 算实到
                absent = int(absent_text)
                present = total - absent
                self.present_edit.setText(str(max(present, 0)))  # 不小于0
            else:
                present = 0
                absent = 0

            # 计算缺勤比例
            ratio = (absent / total) * 100
            self.ratio_show.setText(f"{ratio:.2f}%")

        except Exception:
            self.ratio_show.setText("0.00%")

    # ===================== 语言/主题 =====================
    def refresh_ui(self):
        t = self.t
        self.setWindowTitle(t["title"])
        self.lang_lab.setText(t["lang"])
        self.theme_btn.setText(t["theme"])
        
        # 录入模块
        self.group_input.setTitle(t["group_input"])
        self.class_lab.setText(t["class"]+":")
        self.total_lab.setText(t["total"]+":")  # 总人数标签
        self.course_lab.setText(t["course"]+":")
        self.date_lab.setText(t["date"]+":")
        self.week_lab.setText(t["week"]+":")
        self.status_lab.setText(t["status"]+":")
        # 考勤标签
        self.present_lab.setText(t["present"]+":")
        self.absent_num_lab.setText(t["absent_num"]+":")
        self.ratio_lab.setText(t["ratio"]+":")
        self.remark_lab.setText(t["remark"]+":")
        
        # 提示文字
        self.class_edit.setPlaceholderText(t["class"])
        self.total_edit.setPlaceholderText(t["total"])
        self.course_edit.setPlaceholderText(t["course"])
        self.present_edit.setPlaceholderText(t["present"])
        self.absent_num_edit.setPlaceholderText(t["absent_num"])
        self.remark_edit.setPlaceholderText(t["remark"])
        
        # 按钮/状态
        self.add_btn.setText(t["add"])
        self.clear_btn.setText(t["clear"])
        self.export_btn.setText(t["export"])
        self.folder_btn.setText(t["folder"])
        self.status_combo.clear()
        self.status_combo.addItems([t["attend"], t["absent"], t["leave"]])
        self.path_lab.setText(f"{t['folder']}: {SAVE_PATH or t['no_data']}")

    def switch_lang(self, idx):
        global CURRENT_LANG
        CURRENT_LANG = LANG_KEYS[idx]
        self.t = TRANSLATIONS[CURRENT_LANG]
        self.refresh_ui()
        self.apply_theme()

    def switch_theme(self):
        global THEME_LIGHT
        THEME_LIGHT = not THEME_LIGHT
        self.apply_theme()

    def apply_theme(self):
        bg = "#fff" if THEME_LIGHT else "#121212"
        fg = "#000" if THEME_LIGHT else "#0f0"
        input_bg = "#ffffff" if THEME_LIGHT else "#2b2b2b"
        input_fg = "#000000" if THEME_LIGHT else "#ffffff"
        
        self.setStyleSheet(f"""
            QWidget{{background:{bg};color:{fg};}}
            QGroupBox{{border:1px solid #ccc;border-radius:8px;padding:15px;}}
            QLineEdit, QComboBox, QDateEdit, QTextEdit{{
                background:{input_bg};
                color:{input_fg};
                border:1px solid #ccc;
                border-radius:6px;
                padding:6px 8px;
            }}
        """)

    # ===================== 数据操作 =====================
    def select_folder(self):
        global SAVE_PATH
        path = QFileDialog.getExistingDirectory()
        if path:
            SAVE_PATH = path
            self.refresh_ui()

    def add_record(self):
        try:
            data = {
                "class": self.class_edit.text().strip(),
                "total": self.total_edit.text().strip() or "0",
                "course": self.course_edit.text().strip(),
                "date": self.date_edit.date().toString("yyyy-MM-dd"),
                "week": self.week_combo.currentText(),
                "status": self.status_combo.currentText(),
                "present": self.present_edit.text().strip() or "0",
                "absent": self.absent_num_edit.text().strip() or "0",
                "ratio": self.ratio_show.text(),
                "remark": self.remark_edit.toPlainText().strip()
            }
            if not data["class"]:
                return

            DATA.append(data)
            self.refresh_table()

            # 清空输入
            self.class_edit.clear()
            self.total_edit.clear()
            self.course_edit.clear()
            self.present_edit.clear()
            self.absent_num_edit.clear()
            self.remark_edit.clear()
        except Exception:
            pass

    def refresh_table(self):
        self.table.setRowCount(0)
        for row, item in enumerate(DATA):
            self.table.insertRow(row)
            vals = [
                item["class"], item["total"], item["course"], item["date"], 
                item["week"], item["status"], item["present"], item["absent"], 
                item["ratio"], item["remark"]
            ]
            for col, val in enumerate(vals):
                self.table.setItem(row, col, QTableWidgetItem(val))

    def clear_all(self):
        global DATA
        DATA.clear()
        self.table.setRowCount(0)
        self.class_edit.clear()
        self.total_edit.clear()
        self.course_edit.clear()
        self.present_edit.clear()
        self.absent_num_edit.clear()
        self.remark_edit.clear()

    def export_excel(self):
        if not SAVE_PATH:
            QMessageBox.information(self, self.t["title"], self.t["select_folder"])
            return
        if not DATA:
            QMessageBox.information(self, self.t["title"], self.t["no_data"])
            return
        file = os.path.join(SAVE_PATH, f"课表_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
        pd.DataFrame(DATA).to_excel(file, index=False)
        QMessageBox.information(self, self.t["title"], self.t["success"])

# ===================== 程序入口 =====================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ScheduleApp()
    window.show()
    sys.exit(app.exec())
