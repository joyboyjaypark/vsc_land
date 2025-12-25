import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit,
    QComboBox, QPushButton, QGridLayout, QMessageBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QGroupBox,
    QSizePolicy, QProgressBar, QInputDialog, QTabWidget,
    QVBoxLayout, QTextEdit,
    QListWidget, QListWidgetItem,
)
from PyQt5.QtWidgets import QHBoxLayout, QCheckBox
from PyQt5.QtGui import QColor, QBrush
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
import os

import requests
import xml.etree.ElementTree as ET
import csv
import datetime
import time
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.ticker import FuncFormatter
import functools
import traceback



# Numeric-aware table item: stores numeric value in UserRole and compares numerically
class NumericItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            a = self.data(Qt.UserRole)
            b = other.data(Qt.UserRole)
            if a is not None and b is not None:
                return int(a) < int(b)
        except Exception:
            pass
        return super().__lt__(other)


class VWorldAdmCodeGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("VWorld 행정구역 코드 조회")

        # 위젯들
        lbl_key = QLabel("API Key:")
        self.edit_key = QLineEdit("536CFED0-72BE-3E69-9C7F-1F44FED0E734")
        self.edit_key.setPlaceholderText("브이월드에서 발급받은 인증키")

        # Domain 입력 제거 (고정 또는 API에서 생략)

        # format selection removed; format is fixed to 'json'

        # numOfRows/pageNo 입력 제거 (고정값 사용)

        lbl_sido = QLabel("시/도:")
        self.combo_sido = QComboBox()
        self.combo_sido.addItem("선택")
        self.combo_sido.currentIndexChanged.connect(self.on_sido_changed)

        lbl_sigungu = QLabel("시군구:")
        self.combo_sigungu = QComboBox()
        self.combo_sigungu.addItem("선택")
        self.combo_sigungu.currentIndexChanged.connect(self.on_sigungu_changed)

        lbl_dong = QLabel("읍면동:")
        self.combo_dong = QComboBox()
        self.combo_dong.addItem("선택")
        self.combo_dong.currentIndexChanged.connect(self.on_dong_changed)

        lbl_selected = QLabel("행정코드:")
        self.edit_selected_admcode = QLineEdit()
        self.edit_selected_admcode.setReadOnly(True)

        # '요청 보내기' 버튼 removed; requests will run automatically on startup

        # region result table removed per request

        # 레이아웃
        layout = QGridLayout()
        layout.addWidget(lbl_key,      0, 0)
        layout.addWidget(self.edit_key, 0, 1)
        # 지역 그룹 박스에 시/도, 시군구, 읍면동 및 행정코드 배치
        group_region = QGroupBox("지역")
        group_layout = QGridLayout()
        group_layout.addWidget(lbl_sido, 0, 0)
        group_layout.addWidget(self.combo_sido, 0, 1)
        group_layout.addWidget(lbl_sigungu, 1, 0)
        group_layout.addWidget(self.combo_sigungu, 1, 1)
        group_layout.addWidget(lbl_dong, 2, 0)
        group_layout.addWidget(self.combo_dong, 2, 1)
        group_layout.addWidget(lbl_selected, 3, 0)
        group_layout.addWidget(self.edit_selected_admcode, 3, 1)
        group_region.setLayout(group_layout)
        layout.addWidget(group_region, 2, 0)

        # --- 아파트 실거래 입력 및 결과 ---
        lbl_apt_key = QLabel("Service Key:")
        # 기본값 설정
        self.edit_apt_key = QLineEdit("Nv0jBnCHJXCT20iu910K%2FIGnF556Vt2w06icWR2uj66dF73AiTNBXaM7bIS9Nu9C0cmB7sGVgpnbCiK01Qkgeg%3D%3D")
        self.edit_apt_key.setPlaceholderText("발급받은 인증키를 입력하세요 (URL 디코딩된 값 권장)")

        lbl_apt_lawd = QLabel("지역코드(LAWD_CD):")
        self.edit_apt_lawd = QLineEdit()
        self.edit_apt_lawd.setPlaceholderText("예: 11110")
        # 입력창이 그룹박스 너비 제약에 의해 너무 작아지지 않도록 확장 가능하도록 설정
        try:
            self.edit_apt_lawd.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        except Exception:
            pass

        lbl_apt_ymd = QLabel("조회기간(YYYY / MM):")
        # 년/월을 각각 선택하도록 콤보박스 2개씩 생성
        now = datetime.date.today()
        curr_year = now.year
        # 최근 10년을 선택지로 제공 (현재연도 ~ 현재-9)
        years = [str(y) for y in range(curr_year, curr_year - 10, -1)]
        months = [f"{i:02d}" for i in range(1, 13)]

        self.combo_apt_year_from = QComboBox()
        self.combo_apt_month_from = QComboBox()
        self.combo_apt_year_to = QComboBox()
        self.combo_apt_month_to = QComboBox()

        self.combo_apt_year_from.addItems(years)
        self.combo_apt_year_to.addItems(years)
        self.combo_apt_month_from.addItems(months)
        self.combo_apt_month_to.addItems(months)

        # 기본값: to = 현재, from = 12개월 전
        total = now.year * 12 + now.month - 1
        total_from = total - 12
        fy = total_from // 12
        fm = total_from % 12 + 1

        # year combo는 years가 내림차순이므로 인덱스를 찾아 설정
        try:
            self.combo_apt_year_from.setCurrentIndex(years.index(str(fy)))
        except ValueError:
            self.combo_apt_year_from.setCurrentIndex(len(years) - 1)
        self.combo_apt_month_from.setCurrentIndex(fm - 1)
        try:
            self.combo_apt_year_to.setCurrentIndex(years.index(str(now.year)))
        except ValueError:
            self.combo_apt_year_to.setCurrentIndex(0)
        self.combo_apt_month_to.setCurrentIndex(now.month - 1)

        self.btn_apt_fetch = QPushButton("아파트 조회")
        self.btn_apt_fetch.clicked.connect(self.on_apt_fetch)

        self.btn_apt_save = QPushButton("CSV로 저장")
        self.btn_apt_save.clicked.connect(self.on_apt_save_csv)
        self.btn_apt_save.setEnabled(False)

        self.btn_apt_cancel = QPushButton("취소")
        self.btn_apt_cancel.clicked.connect(self.on_apt_cancel)
        self.btn_apt_cancel.setEnabled(False)

        lbl_apt_url = QLabel("최종 URL:")
        self.edit_apt_url = QLineEdit()
        self.edit_apt_url.setReadOnly(True)

        self.apt_table = QTableWidget()
        apt_headers = [
            "아파트명", "아파트동", "전용면적", "계약일", "거래금액(만원)", "층", "건축년도", "법정동", "지번", "지역코드",
            "거래유형", "중개사소재지", "등기일자", "거래주체_매도자", "거래주체_매수자", "토지임대부여부"
        ]
        self.apt_table.setColumnCount(len(apt_headers))
        self.apt_table.setHorizontalHeaderLabels(apt_headers)
        hdr = self.apt_table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.Interactive)
        hdr.setSectionsMovable(True)
        hdr.setStretchLastSection(False)
        # 헤더 클릭으로 열을 오름/내림차순 정렬 가능하도록 설정
        self.apt_table.setSortingEnabled(True)
        # 헤더 우클릭으로 필터 입력/해제 가능하도록 설정
        hdr.setContextMenuPolicy(Qt.CustomContextMenu)
        hdr.customContextMenuRequested.connect(self.apt_header_context_menu)
        # 필터 저장용
        self.apt_filters = {}
        self.apt_rows_master = []
        self.apt_table.setEditTriggers(QTableWidget.NoEditTriggers)

        # 진행 상태 표시 (조회 진행률)
        self.status_label = QLabel("")
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setValue(0)

        # apt inputs layout: move Service Key under API Key
        layout.addWidget(lbl_apt_key, 1, 0)
        layout.addWidget(self.edit_apt_key, 1, 1)

        # 범위설정 그룹박스에 지역코드, 조회년월, 최종 URL 배치
        group_range = QGroupBox("범위설정")
        gr_layout = QGridLayout()
        # 내부 여백을 조금 주어 프레임 모서리와 간격 확보
        gr_layout.setContentsMargins(8, 8, 8, 8)
        # 컬럼 스트레치: 라벨(0) 고정, 입력창(1) 확장, 보조(2) 작게
        gr_layout.setColumnStretch(0, 0)
        gr_layout.setColumnStretch(1, 2)
        gr_layout.setColumnStretch(2, 1)
        gr_layout.addWidget(lbl_apt_lawd, 0, 0)
        # 지역코드 입력창을 두 칸(span)으로 넓혀 프레임 내부에서 최대한 확장
        gr_layout.addWidget(self.edit_apt_lawd, 0, 1, 1, 2)
        lbl_apt_ymd_from = QLabel("시작년월:")
        lbl_apt_ymd_to = QLabel("종료년월:")
        gr_layout.addWidget(lbl_apt_ymd_from, 1, 0)
        gr_layout.addWidget(self.combo_apt_year_from, 1, 1)
        gr_layout.addWidget(self.combo_apt_month_from, 1, 2)
        gr_layout.addWidget(lbl_apt_ymd_to, 2, 0)
        gr_layout.addWidget(self.combo_apt_year_to, 2, 1)
        gr_layout.addWidget(self.combo_apt_month_to, 2, 2)
        # 최종 URL은 종료년월 다음 줄에 배치 (URL 필드는 2열을 차지)
        gr_layout.addWidget(lbl_apt_url, 3, 0)
        gr_layout.addWidget(self.edit_apt_url, 3, 1, 1, 2)
        group_range.setLayout(gr_layout)
        # 범위설정 프레임을 지역 프레임 오른쪽에 배치
        layout.addWidget(group_range, 2, 1)

        # 버튼과 테이블을 한 줄 위로 이동하여 레이아웃을 압축
        layout.addWidget(self.btn_apt_fetch, 3, 0)
        layout.addWidget(self.btn_apt_save, 3, 1)
        layout.addWidget(self.btn_apt_cancel, 3, 2)
        layout.addWidget(self.apt_table, 4, 0, 1, 2)
        # 상태 표시줄: 상태 텍스트(왼쪽) + 진행바(오른쪽)
        layout.addWidget(self.status_label, 5, 0)
        layout.addWidget(self.progress_bar, 5, 1)

        # Put existing real-estate layout into first tab and add an economic-indicators tab
        self.tabs = QTabWidget()
        tab_real = QWidget()
        tab_real.setLayout(layout)

        tab_econ = QWidget()
        econ_layout = QGridLayout()
        # Group box to contain inputs (Data 1)
        group_data1 = QGroupBox("Data 1")
        group_layout = QGridLayout()
        # 경제지표 탭: 간소화된 UI — 인증키 입력 + 결과 콤보
        lbl_bok_key = QLabel("한국은행 인증키:")
        self.edit_bok_key = QLineEdit("TZ9P9GAR03LBXV2J3QGU")
        self.edit_bok_key.setPlaceholderText("한국은행 Open API 인증키")

        group_layout.addWidget(lbl_bok_key, 0, 0)
        group_layout.addWidget(self.edit_bok_key, 0, 1)

        # 콤보 제목
        lbl_bok_list = QLabel("서비스 통계 목록:")
        self.bok_combo = QComboBox()
        self.bok_combo.setEditable(False)
        # show up to 20 items in the dropdown popup instead of the default 10
        try:
            self.bok_combo.setMaxVisibleItems(20)
        except Exception:
            pass
        self.bok_combo.currentIndexChanged.connect(self.on_bok_select)

        group_layout.addWidget(lbl_bok_list, 1, 0)
        group_layout.addWidget(self.bok_combo, 1, 1, 1, 5)
        # 세부 목록 콤보박스 (서비스 통계 목록 선택 시 채워짐)
        lbl_bok_detail = QLabel("세부 목록:")
        self.bok_detail_combo = QComboBox()
        self.bok_detail_combo.setEditable(False)
        self.bok_detail_combo.currentIndexChanged.connect(self.on_bok_detail_select)
        group_layout.addWidget(lbl_bok_detail, 2, 0)
        group_layout.addWidget(self.bok_detail_combo, 2, 1, 1, 5)

        # 검색은 자동실행으로 처리하므로 버튼은 표시하지 않습니다.

        # mapping index -> STAT_CODE
        self.bok_index_to_code = {}

        # 기간 선택 콤보박스 (시작/종료)
        lbl_period_start = QLabel("기간 시작:")
        lbl_period_end = QLabel("기간 종료:")
        self.combo_period_start = QComboBox()
        self.combo_period_end = QComboBox()
        self.combo_period_start.setEditable(False)
        self.combo_period_end.setEditable(False)
        # make both period combos the same width for consistent layout
        try:
            fixed_w = 140
            self.combo_period_start.setFixedWidth(fixed_w)
            self.combo_period_end.setFixedWidth(fixed_w)
        except Exception:
            pass
        group_layout.addWidget(lbl_period_start, 3, 0)
        group_layout.addWidget(self.combo_period_start, 3, 1)
        group_layout.addWidget(lbl_period_end, 3, 2)
        group_layout.addWidget(self.combo_period_end, 3, 3)
        # saved list box under Data 1: shows appended selections
        lbl_saved_list = QLabel("저장된 목록:")
        self.bok_listbox = QListWidget()
        # storage for mapping listbox entries -> table row ranges
        self.bok_saved_ranges = []
        try:
            self.bok_listbox.setFixedHeight(100)
        except Exception:
            pass
        group_layout.addWidget(lbl_saved_list, 4, 0)
        group_layout.addWidget(self.bok_listbox, 4, 1, 1, 5)
        # 출력하기 버튼 및 결과 테이블
        self.btn_bok_print = QPushButton("출력하기")
        self.btn_bok_print.clicked.connect(self.on_bok_print)
        # 차트 생성 버튼
        self.btn_bok_chart = QPushButton("차트생성")
        self.btn_bok_chart.clicked.connect(self.on_bok_plot)
        # place group box into econ layout
        group_data1.setLayout(group_layout)
        econ_layout.addWidget(group_data1, 0, 0, 1, 7)

        econ_layout.addWidget(self.btn_bok_print, 1, 0)
        econ_layout.addWidget(self.btn_bok_chart, 1, 1)

        self.bok_result_table = QTableWidget()
        self.bok_result_table.setColumnCount(0)
        self.bok_result_table.setRowCount(0)
        self.bok_result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        econ_layout.addWidget(self.bok_result_table, 2, 0, 1, 7)

        tab_econ.setLayout(econ_layout)

        self.tabs.addTab(tab_real, "부동산 거래")
        self.tabs.addTab(tab_econ, "경제지표")

        # 지표누리 탭: 인증키 입력 (기본값 제공)
        tab_ind = QWidget()
        ind_layout = QGridLayout()
        lbl_ind_key = QLabel("지표누리 인증키:")
        self.edit_ind_key = QLineEdit("H4T022E22214155B")
        self.edit_ind_key.setPlaceholderText("지표누리 인증키")
        ind_layout.addWidget(lbl_ind_key, 0, 0)
        ind_layout.addWidget(self.edit_ind_key, 0, 1)
        # URL input below the key (default value)
        lbl_ind_url = QLabel("지표누리 URL:")
        self.edit_ind_url = QLineEdit("https://www.index.go.kr/unity/openApi/xml_idx.do?userId=youngbbo&idntfcId=H4T022E22214155B")
        self.edit_ind_url.setPlaceholderText("지표누리 API URL (기본)")
        ind_layout.addWidget(lbl_ind_url, 1, 0)
        ind_layout.addWidget(self.edit_ind_url, 1, 1, 1, 3)
        # 리스트받기 버튼
        self.btn_ind_list = QPushButton("리스트받기")
        self.btn_ind_list.clicked.connect(self.on_ind_list)
        ind_layout.addWidget(self.btn_ind_list, 2, 0)
        # 결과 테이블
        self.ind_table = QTableWidget()
        self.ind_table.setColumnCount(0)
        self.ind_table.setRowCount(0)
        self.ind_table.setEditTriggers(QTableWidget.NoEditTriggers)
        try:
            self.ind_table.setSortingEnabled(True)
            th = self.ind_table.horizontalHeader()
            th.setSectionsClickable(True)
            th.setSortIndicatorShown(True)
        except Exception:
            pass
        ind_layout.addWidget(self.ind_table, 3, 0, 1, 4)
        tab_ind.setLayout(ind_layout)
        self.tabs.addTab(tab_ind, "지표누리")

        # 출처 탭: 텍스트 박스 추가
        tab_source = QWidget()
        source_layout = QVBoxLayout()
        lbl_source = QLabel("출처 노트:")
        self.source_text = QTextEdit()
        self.source_text.setPlaceholderText("출처나 메모를 입력하세요.")
        # 초기 출처 내용 채우기
        self.source_text.setPlainText(
            "[부동산 거래]\n"
            "1. 주소검색\n"
            "  - 국토교통부 디지털트윈국토\n"
            "  - https://www.vworld.kr/v4po_openapi_s001.do\n"
            "  - https://www.vworld.kr/v4po_search.do?searchCaE=open&searchIdEW=%25EB%25B2%2595%25EC%25A0%2595%25EB%258F%2599%25EC%25A0%2595%25EB%25B3%25B4\n"
            "  - 개발키 : 536CFED0-72BE-3E69-9C7F-1F44FED0E734\n\n"
            "2. 매매내역\n"
            " - 공공데이터포털\n"
            " - 국토교통부_아파트 매매 실거래가 자료\n"
            " - Service Key : Nv0jBnCHJXCT20iu910K%2FIGnF556Vt2w06icWR2uj66dF73AiTNBXaM7bIS9Nu9C0cmB7sGVgpnbCiK01Qkgeg%3D%3D\n\n"
            "[경제지표]\n"
            "1. 경제통계\n"
            "  - 한국은행 경제통계시스템\n"
            "  - https://ecos.bok.or.kr/#/\n"
            "  - 한국은행 인증키 : TZ9P9GAR03LBXV2J3QGU\n\n"
            "[지표누리]\n"
            " - 지표누리공유서비스\n"
            " - https://www.index.go.kr/unity/openApi/openApiIntro.do\n"
            " - 인증키 : \tH4T022E22214155B"
        )
        source_layout.addWidget(lbl_source)
        source_layout.addWidget(self.source_text)
        tab_source.setLayout(source_layout)
        self.tabs.addTab(tab_source, "출처")

        main_layout = QGridLayout()
        main_layout.addWidget(self.tabs, 0, 0)
        self.setLayout(main_layout)
        self.resize(700, 500)

        # 그룹박스 너비를 윈도우 너비의 1/3으로 설정
        try:
            one_third = int(self.width() / 3)
        except Exception:
            one_third = 233
        try:
            group_region.setFixedWidth(one_third)
        except Exception:
            pass
        try:
            group_range.setFixedWidth(one_third)
        except Exception:
            pass

        # 자동으로 요청 실행 (이벤트 루프가 시작된 직후 호출)
        QTimer.singleShot(0, self.send_request)
        QTimer.singleShot(0, self.on_bok_search)

        # worker handle
        self._apt_worker = None


 

    def send_request(self):
        key = self.edit_key.text().strip()
        # domain and format inputs removed from UI; use fixed values
        fmt = "json"
        num_rows = "25"
        page_no = "1"

        if not key:
            QMessageBox.warning(self, "입력 오류", "API Key는 필수입니다.")
            return

        url = "http://api.vworld.kr/ned/data/admCodeList"

        params = {
            "key": key,
            "format": fmt,
            "numOfRows": num_rows,
            "pageNo": page_no
        }

        # `requests` 모듈이 설치되어 있지 않으면 사용자에게 안내
        try:
            import requests
        except ImportError:
            QMessageBox.critical(
                self,
                "의존성 오류",
                "`requests` 모듈이 설치되어 있지 않습니다.\nPowerShell에서 다음을 실행하세요:\npython -m pip install requests"
            )
            return

        try:
            resp = requests.get(url, params=params, timeout=10)
            resp.raise_for_status()

            pairs = []
            if fmt.lower() == "json":
                try:
                    data = resp.json()
                except ValueError:
                    QMessageBox.critical(self, "파싱 오류", "응답을 JSON으로 파싱할 수 없습니다.")
                    return

                def collect_from_json(obj):
                    if isinstance(obj, dict):
                        if "admCode" in obj and "admCodeNm" in obj:
                            pairs.append((obj.get("admCodeNm"), obj.get("admCode")))
                        for v in obj.values():
                            collect_from_json(v)
                    elif isinstance(obj, list):
                        for it in obj:
                            collect_from_json(it)

                collect_from_json(data)
            else:
                # xml
                try:
                    import xml.etree.ElementTree as ET
                    root = ET.fromstring(resp.text)
                except Exception:
                    QMessageBox.critical(self, "파싱 오류", "응답을 XML로 파싱할 수 없습니다.")
                    return

                for elem in root.iter():
                    children = {child.tag: (child.text or "") for child in list(elem)}
                    if "admCode" in children and "admCodeNm" in children:
                        pairs.append((children.get("admCodeNm"), children.get("admCode")))

            # 콤보박스(시/도) 항목 채우기 - admCodeNm의 유니크 값들
            # 또한 이름->admCode 리스트 매핑을 저장
            self.sido_map = {}
            for name, code in pairs:
                if not name:
                    continue
                self.sido_map.setdefault(name, []).append(code)

            adm_names = sorted(self.sido_map.keys())
            # block signals while we repopulate programmatically
            self.combo_sido.blockSignals(True)
            self.combo_sido.clear()
            if adm_names:
                self.combo_sido.addItem("선택")
                self.combo_sido.addItems(adm_names)
            else:
                self.combo_sido.addItem("없음")
            self.combo_sido.blockSignals(False)

            # 초기 시군구 콤보 초기화 (block signals)
            self.combo_sigungu.blockSignals(True)
            self.combo_sigungu.clear()
            self.combo_sigungu.addItem("선택")
            self.combo_sigungu.blockSignals(False)
            self.sigungu_map = {}
            self.last_si_pairs = []

            # 결과 테이블은 제거되었으므로 UI 업데이트는 생략합니다.
            pass
            if not pairs:
                QMessageBox.information(self, "결과 없음", "조회 결과에서 admCodeNm/admCode를 찾지 못했습니다.")
        except requests.exceptions.RequestException as e:
            QMessageBox.critical(self, "요청 실패", str(e))

    def on_sido_changed(self, index):
        name = self.combo_sido.currentText()
        if not name or name in ("선택", "없음"):
            if hasattr(self, 'edit_selected_admcode'):
                self.edit_selected_admcode.setText("")
            return

        # admCode 값 가져오기 (매핑에서 첫번째 코드 사용)
        codes = getattr(self, 'sido_map', {}).get(name)
        if not codes:
            QMessageBox.information(self, "정보", "선택한 시/도의 admCode를 찾을 수 없습니다.")
            return
        adm_code_full = codes[0]
        if hasattr(self, 'edit_selected_admcode'):
            self.edit_selected_admcode.setText(adm_code_full or "")
        # 지역코드(앞 5자리)를 아파트 입력 필드에 자동 반영
        if hasattr(self, 'edit_apt_lawd') and adm_code_full:
            try:
                self.edit_apt_lawd.setText(adm_code_full[:5])
            except Exception:
                pass
        adm_code_prefix = adm_code_full[:2]

        # 요청 파라미터: key는 UI의 값, format=json, numOfRows=200
        key = self.edit_key.text().strip()
        if not key:
            QMessageBox.warning(self, "입력 오류", "API Key는 필수입니다.")
            return

        try:
            import requests
        except ImportError:
            QMessageBox.critical(
                self,
                "의존성 오류",
                "`requests` 모듈이 설치되어 있지 않습니다.\nPowerShell에서 다음을 실행하세요:\npython -m pip install requests"
            )
            return

        url = "http://api.vworld.kr/ned/data/admSiList"
        params = {
            "key": key,
            "admCode": adm_code_prefix,
            "format": "json",
            "numOfRows": "200",
            "pageNo": "1",
        }

        try:
            resp = requests.get(url, params=params, timeout=10)
            resp.raise_for_status()
        except requests.exceptions.RequestException as e:
            QMessageBox.critical(self, "요청 실패", str(e))
            return

        # JSON 파싱 및 테이블에 표시
        try:
            data = resp.json()
        except ValueError:
            QMessageBox.critical(self, "파싱 오류", "응답을 JSON으로 파싱할 수 없습니다.")
            return

        pairs = []

        def collect_from_json(obj):
            if isinstance(obj, dict):
                if "admCode" in obj:
                    # 시군구명은 lowestAdmCodeNm이 우선, 없으면 admCodeNm 사용
                    name = obj.get("lowestAdmCodeNm") or obj.get("admCodeNm")
                    if name is not None:
                        pairs.append((name, obj.get("admCode")))
                for v in obj.values():
                    collect_from_json(v)
            elif isinstance(obj, list):
                for it in obj:
                    collect_from_json(it)

        collect_from_json(data)
        # 저장해두고 콤보박스(시군구) 채우기
        self.last_si_pairs = pairs
        self.sigungu_map = {}
        for n, c in pairs:
            if not n:
                continue
            self.sigungu_map.setdefault(n, []).append(c)

        sigungu_names = sorted(self.sigungu_map.keys())
        # block signals to avoid triggering on_sigungu_changed during programmatic update
        self.combo_sigungu.blockSignals(True)
        self.combo_sigungu.clear()
        if sigungu_names:
            self.combo_sigungu.addItem("선택")
            self.combo_sigungu.addItems(sigungu_names)
        else:
            self.combo_sigungu.addItem("없음")
        self.combo_sigungu.blockSignals(False)

        # 초기 읍면동 콤보 비우기
        if hasattr(self, 'combo_dong'):
            self.combo_dong.blockSignals(True)
            self.combo_dong.clear()
            self.combo_dong.addItem("선택")
            self.combo_dong.blockSignals(False)
            self.last_dong_pairs = []

        # 테이블 표시는 제거되어 있음
        if not pairs:
            QMessageBox.information(self, "결과 없음", "선택한 시/도에 대한 시군구 결과가 없습니다.")

    def on_sigungu_changed(self, index):
        name = self.combo_sigungu.currentText()
        if not name or name in ("선택", "없음"):
            # When user clears 시군구 selection, restore the 시/도-level admCode
            try:
                sel_sido = self.combo_sido.currentText()
                if sel_sido and sel_sido not in ("선택", "없음"):
                    codes = getattr(self, 'sido_map', {}).get(sel_sido)
                    if codes:
                        adm_code_full = codes[0]
                        if hasattr(self, 'edit_selected_admcode'):
                            self.edit_selected_admcode.setText(adm_code_full or "")
                        if hasattr(self, 'edit_apt_lawd') and adm_code_full:
                            try:
                                self.edit_apt_lawd.setText(adm_code_full[:5])
                            except Exception:
                                pass
                        # do not proceed with further requests
                        return
            except Exception:
                pass
            # fallback: clear selection display
            if hasattr(self, 'edit_selected_admcode'):
                self.edit_selected_admcode.setText("")
            return

        # 선택한 시군구의 admCode(첫번째)로 읍면동 목록 요청
        codes = getattr(self, 'sigungu_map', {}).get(name)
        if not codes:
            QMessageBox.information(self, "정보", "선택한 시군구의 admCode를 찾을 수 없습니다.")
            return
        adm_code_full = codes[0]
        if hasattr(self, 'edit_selected_admcode'):
            self.edit_selected_admcode.setText(adm_code_full or "")
        # 지역코드(앞 5자리)를 아파트 입력 필드에 자동 반영
        if hasattr(self, 'edit_apt_lawd') and adm_code_full:
            try:
                self.edit_apt_lawd.setText(adm_code_full[:5])
            except Exception:
                pass

        key = self.edit_key.text().strip()
        if not key:
            QMessageBox.warning(self, "입력 오류", "API Key는 필수입니다.")
            return

        try:
            import requests
        except ImportError:
            QMessageBox.critical(
                self,
                "의존성 오류",
                "`requests` 모듈이 설치되어 있지 않습니다.\nPowerShell에서 다음을 실행하세요:\npython -m pip install requests"
            )
            return

        url = "http://api.vworld.kr/ned/data/admDongList"
        params = {
            "key": key,
            "admCode": adm_code_full,
            "format": "json",
            "numOfRows": "200",
            "pageNo": "1",
        }

        try:
            resp = requests.get(url, params=params, timeout=10)
            resp.raise_for_status()
        except requests.exceptions.RequestException as e:
            QMessageBox.critical(self, "요청 실패", str(e))
            return

        try:
            data = resp.json()
        except ValueError:
            QMessageBox.critical(self, "파싱 오류", "응답을 JSON으로 파싱할 수 없습니다.")
            return

        pairs = []

        def collect_from_json(obj):
            if isinstance(obj, dict):
                if "admCode" in obj:
                    name_d = obj.get("lowestAdmCodeNm") or obj.get("admCodeNm")
                    if name_d is not None:
                        pairs.append((name_d, obj.get("admCode")))
                for v in obj.values():
                    collect_from_json(v)
            elif isinstance(obj, list):
                for it in obj:
                    collect_from_json(it)

        collect_from_json(data)

        # 저장 및 콤보(읍면동) 채우기
        self.last_dong_pairs = pairs
        self.dong_map = {}
        for n, c in pairs:
            if not n:
                continue
            self.dong_map.setdefault(n, []).append(c)

        dong_names = sorted(self.dong_map.keys())
        # block signals to avoid triggering on_dong_changed during programmatic update
        self.combo_dong.blockSignals(True)
        self.combo_dong.clear()
        if dong_names:
            self.combo_dong.addItem("선택")
            self.combo_dong.addItems(dong_names)
        else:
            self.combo_dong.addItem("없음")
        self.combo_dong.blockSignals(False)

        # 테이블 표시는 제거되어 있음
        pass
        if not pairs:
            QMessageBox.information(self, "결과 없음", "선택한 시군구에 대한 읍면동 결과가 없습니다.")

    def on_dong_changed(self, index):
        name = self.combo_dong.currentText()
        if not name or name in ("선택", "없음"):
            # show all if '선택' chosen
            if hasattr(self, 'edit_selected_admcode'):
                self.edit_selected_admcode.setText("")
            # table removed; nothing to update
            return

        pairs = getattr(self, 'last_dong_pairs', [])
        filtered = [(n, c) for (n, c) in pairs if n == name]

        # 표시할 admCode 설정 (첫번째 값)
        codes = getattr(self, 'dong_map', {}).get(name)
        if codes and hasattr(self, 'edit_selected_admcode'):
            self.edit_selected_admcode.setText(codes[0] or "")
        # 지역코드(앞 5자리)를 아파트 입력 필드에 자동 반영
        if codes and hasattr(self, 'edit_apt_lawd'):
            try:
                self.edit_apt_lawd.setText((codes[0] or "")[:5])
            except Exception:
                pass

        # table removed; nothing to update
        pass
        if not filtered:
            QMessageBox.information(self, "결과 없음", "선택한 읍면동에 대한 결과가 없습니다.")

    # =========== 한국은행(ECOS) 통계표 목록 조회 ===========
    def on_bok_search(self):
        # Use fixed parameters per requirements
        service = "StatisticTableList"
        key = self.edit_bok_key.text().strip()
        req_type = "xml"
        lang = "kr"
        start = "1"
        end = "1000"
        stat_code = ""

        if not key:
            QMessageBox.warning(self, "입력 오류", "한국은행 인증키를 입력하세요.")
            return

        base = "https://ecos.bok.or.kr/api"
        parts = [base, service, key, req_type, lang, start, end]
        if stat_code:
            parts.append(stat_code)
        url = "/".join(parts)

        try:
            resp = requests.get(url, timeout=15)
            resp.raise_for_status()
            data = resp.content
        except Exception as e:
            QMessageBox.critical(self, "요청 실패", f"API 요청 중 오류가 발생했습니다:\n{e}")
            return

        try:
            root = ET.fromstring(data)
        except Exception as e:
            QMessageBox.critical(self, "파싱 오류", f"응답 XML 파싱 실패:\n{e}")
            return

        nodes = root.findall('.//list') or root.findall('.//row') or root.findall('.//item')

        self.bok_combo.clear()
        self.bok_index_to_code.clear()

        for idx, node in enumerate(nodes):
            name = node.findtext('STAT_NAME') or node.findtext('STAT_NM') or ''
            srch = (node.findtext('SRCH_YN') or '').strip()
            code = node.findtext('STAT_CODE') or node.findtext('STAT_ID') or ''

            self.bok_combo.addItem(name)
            # 색상 처리: SRCH_YN == 'Y'이면 항목 글씨를 빨갛게 설정
            if srch.upper() == 'Y':
                try:
                    self.bok_combo.setItemData(idx, QBrush(QColor('red')), Qt.ForegroundRole)
                except Exception:
                    pass

            self.bok_index_to_code[idx] = code

        if not nodes:
            QMessageBox.information(self, "결과 없음", "조회된 결과가 없습니다.")

    def on_ind_list(self):
        # Fetch and display XML from the 지표누리 URL in the ind tab
        url = self.edit_ind_url.text().strip() if getattr(self, 'edit_ind_url', None) else ''
        if not url:
            QMessageBox.warning(self, "입력 오류", "지표누리 URL을 입력하세요.")
            return
        try:
            resp = requests.get(url, timeout=20)
            resp.raise_for_status()
            data = resp.content
        except Exception as e:
            QMessageBox.critical(self, "요청 실패", f"요청 중 오류가 발생했습니다:\n{e}")
            return

        try:
            root = ET.fromstring(data)
        except Exception as e:
            QMessageBox.critical(self, "파싱 오류", f"응답 XML 파싱 실패:\n{e}")
            return

        # Try common item containers
        items = root.findall('.//item') or root.findall('.//list') or root.findall('.//row') or []

        # If none found, search for a parent with repeated child tags
        if not items:
            for parent in root.iter():
                child_tags = [c.tag for c in list(parent) if c.tag]
                if not child_tags:
                    continue
                # find tag with highest occurrence
                from collections import Counter
                cnt = Counter(child_tags)
                most_common_tag, count = cnt.most_common(1)[0]
                if count > 1:
                    items = parent.findall(most_common_tag)
                    if items:
                        break

        # If still no repeated items, treat root's immediate children as a single-row table
        if not items:
            cols = [c.tag for c in list(root)]
            rows = [[(c.text or '').strip() if c is not None and c.text else '' for c in list(root)]] if cols else []
            if not rows:
                QMessageBox.information(self, "결과 없음", "표로 표시할 반복 항목을 찾을 수 없습니다.")
                return
            # populate table
            self.ind_table.setColumnCount(len(cols))
            self.ind_table.setHorizontalHeaderLabels(cols)
            self.ind_table.setRowCount(len(rows))
            for r, row in enumerate(rows):
                for c, val in enumerate(row):
                    self.ind_table.setItem(r, c, QTableWidgetItem(val))
            return

        # Build column set from all items' child tags
        col_set = []
        for it in items:
            for child in list(it):
                if child.tag not in col_set:
                    col_set.append(child.tag)

        # populate rows
        self.ind_table.setColumnCount(len(col_set))
        self.ind_table.setHorizontalHeaderLabels(col_set)
        self.ind_table.setRowCount(len(items))
        for r, it in enumerate(items):
            for c, tag in enumerate(col_set):
                try:
                    txt = it.findtext(tag) or ''
                except Exception:
                    txt = ''
                txt = (txt or '').strip()
                self.ind_table.setItem(r, c, QTableWidgetItem(txt))

        # 자동으로 첫 항목의 세부목록도 불러오도록 (있다면)
        try:
            if self.bok_combo.count() > 0:
                QTimer.singleShot(200, lambda: self.on_bok_select())
        except Exception:
            pass

    def _load_stat_item_list(self, stat_code):
        if not stat_code:
            # clear detail
            try:
                self.bok_detail_combo.clear()
            except Exception:
                pass
            return

        key = self.edit_bok_key.text().strip()
        if not key:
            return

        service = "StatisticItemList"
        req_type = "xml"
        lang = "kr"
        start = "1"
        end = "100"  # per request

        base = "https://ecos.bok.or.kr/api"
        parts = [base, service, key, req_type, lang, start, end, stat_code]
        url = "/".join(parts)

        try:
            resp = requests.get(url, timeout=15)
            resp.raise_for_status()
            data = resp.content
        except Exception as e:
            # show but don't block
            try:
                self.status_label.setText(f"세부목록 조회 실패: {e}")
            except Exception:
                pass
            return

        try:
            root = ET.fromstring(data)
        except Exception as e:
            try:
                self.status_label.setText(f"세부목목 파싱 실패: {e}")
            except Exception:
                pass
            return

        nodes = root.findall('.//list') or root.findall('.//row') or root.findall('.//item')

        self.bok_detail_combo.clear()
        self.bok_detail_index_to_pitem = {}
        self.bok_detail_index_to_info = {}
        # 맨 위에 빈 항목 추가
        try:
            self.bok_detail_combo.addItem("")
        except Exception:
            pass

        for idx, node in enumerate(nodes):
            item_name = (node.findtext('ITEM_NAME') or '').strip()
            cycle = (node.findtext('CYCLE') or '').strip()
            p_item_code = (node.findtext('P_ITEM_CODE') or '').strip()
            item_code = (node.findtext('ITEM_CODE') or '').strip()
            start_time = (node.findtext('START_TIME') or '').strip()
            end_time = (node.findtext('END_TIME') or '').strip()

            # display as ITEM_NAME_CYCLE_START_TIME_END_TIME (underscore-separated)
            parts = [item_name, cycle, start_time, end_time]
            # include only non-empty parts to avoid trailing underscores
            display = "_".join([p for p in parts if p])
            if not display:
                display = ''.join(node.itertext()).strip()[:100]

            # add item after the initial empty item -> mapping index = idx + 1
            add_index = idx + 1
            self.bok_detail_combo.addItem(display)
            if p_item_code:
                try:
                    self.bok_detail_combo.setItemData(add_index, QBrush(QColor('red')), Qt.ForegroundRole)
                except Exception:
                    pass

            self.bok_detail_index_to_pitem[add_index] = p_item_code
            self.bok_detail_index_to_info[add_index] = {
                'CYCLE': cycle,
                'ITEM_CODE': item_code,
                'START_TIME': start_time,
                'END_TIME': end_time,
                'P_ITEM_CODE': p_item_code,
            }

        if not nodes:
            try:
                self.status_label.setText("세부항목 없음")
            except Exception:
                pass

    def on_bok_detail_select(self):
        sel = self.bok_detail_combo.currentIndex()
        # index 0 is the intentionally empty top item
        if sel <= 0:
            # clear period combos when empty/top selected
            try:
                self.combo_period_start.clear()
                self.combo_period_end.clear()
            except Exception:
                pass
            return
        info = self.bok_detail_index_to_info.get(sel, {})
        # 기간 콤보 채우기: CYCLE이 'A'이면 연 단위(START_TIME~END_TIME)
        try:
            cycle = (info.get('CYCLE') or '').upper()
            start_t = (info.get('START_TIME') or '').strip()
            end_t = (info.get('END_TIME') or '').strip()
            self.combo_period_start.clear()
            self.combo_period_end.clear()

            if cycle == 'A' and start_t and end_t:
                try:
                    sy = int(start_t[:4])
                    ey = int(end_t[:4])
                except Exception:
                    sy = None; ey = None

                if sy is not None and ey is not None and sy <= ey:
                    years = [str(y) for y in range(sy, ey + 1)]
                    self.combo_period_start.addItems(years)
                    self.combo_period_end.addItems(years)
                    try:
                        self.combo_period_start.setCurrentIndex(0)
                        self.combo_period_end.setCurrentIndex(len(years) - 1)
                    except Exception:
                        pass

            elif cycle == 'Q' and start_t and end_t:
                # parse quarters from START_TIME/END_TIME
                def parse_quarter(t):
                    # accept formats: YYYYMM, YYYYQn, YYYY-Qn, YYYY.n (fallback)
                    if not t:
                        return None
                    t = t.strip()
                    # YYYYMM
                    if len(t) >= 6 and t[:6].isdigit():
                        y = int(t[:4])
                        m = int(t[4:6])
                        q = (m - 1) // 3 + 1
                        return (y, q)
                    # YYYYQn or YYYY-Qn
                    import re
                    m = re.match(r"^(\d{4})\D*Q?(\d)$", t, re.IGNORECASE)
                    if m:
                        y = int(m.group(1))
                        q = int(m.group(2))
                        return (y, q)
                    # try to parse leading year
                    try:
                        y = int(t[:4])
                        # default to Q1 if month not present
                        return (y, 1)
                    except Exception:
                        return None

                start_q = parse_quarter(start_t)
                end_q = parse_quarter(end_t)
                if start_q and end_q:
                    sy, sq = start_q
                    ey, eq = end_q
                    # convert to linear quarter index
                    start_idx = sy * 4 + (sq - 1)
                    end_idx = ey * 4 + (eq - 1)
                    if start_idx <= end_idx:
                        quarters = []
                        for idx in range(start_idx, end_idx + 1):
                            y = idx // 4
                            q = (idx % 4) + 1
                            quarters.append(f"{y}Q{q}")
                        self.combo_period_start.addItems(quarters)
                        self.combo_period_end.addItems(quarters)
                        try:
                            self.combo_period_start.setCurrentIndex(0)
                            self.combo_period_end.setCurrentIndex(len(quarters) - 1)
                        except Exception:
                            pass
            elif cycle == 'M' and start_t and end_t:
                # parse months from START_TIME/END_TIME; accept YYYYMM or YYYY-MM or YYYY.MM
                def parse_ym(t):
                    if not t:
                        return None
                    t = t.strip()
                    # raw digits YYYYMM
                    if len(t) >= 6 and t[:6].isdigit():
                        try:
                            y = int(t[:4]); m = int(t[4:6]);
                            return (y, m)
                        except Exception:
                            return None
                    import re
                    m = re.match(r"^(\d{4})\D?(\d{1,2})", t)
                    if m:
                        try:
                            return (int(m.group(1)), int(m.group(2)))
                        except Exception:
                            return None
                    try:
                        # fallback: year only
                        y = int(t[:4])
                        return (y, 1)
                    except Exception:
                        return None

                s = parse_ym(start_t)
                e = parse_ym(end_t)
                if s and e:
                    sy, sm = s
                    ey, em = e
                    # convert to month index
                    start_idx = sy * 12 + (sm - 1)
                    end_idx = ey * 12 + (em - 1)
                    if start_idx <= end_idx:
                        months = []
                        for idx in range(start_idx, end_idx + 1):
                            y = idx // 12
                            mth = (idx % 12) + 1
                            months.append(f"{y:04d}{mth:02d}")
                        self.combo_period_start.addItems(months)
                        self.combo_period_end.addItems(months)
                        try:
                            self.combo_period_start.setCurrentIndex(0)
                            self.combo_period_end.setCurrentIndex(len(months) - 1)
                        except Exception:
                            pass
            else:
                # 기타 사이클: 기본적으로 START_TIME/END_TIME을 단일 항목으로 넣음
                if start_t:
                    self.combo_period_start.addItem(start_t)
                if end_t:
                    self.combo_period_end.addItem(end_t)
        except Exception:
            pass

    def on_bok_print(self):
        # Build StatisticSearch URL and display results in table
        key = self.edit_bok_key.text().strip()
        if not key:
            QMessageBox.warning(self, "입력 오류", "한국은행 인증키를 입력하세요.")
            return

        stat_idx = self.bok_combo.currentIndex()
        if stat_idx < 0:
            QMessageBox.warning(self, "입력 오류", "서비스 통계 목록을 선택하세요.")
            return
        stat_code = self.bok_index_to_code.get(stat_idx, '').strip()
        if not stat_code:
            QMessageBox.warning(self, "입력 오류", "선택된 통계표의 STAT_CODE가 없습니다.")
            return

        detail_idx = self.bok_detail_combo.currentIndex()
        if detail_idx < 0:
            QMessageBox.warning(self, "입력 오류", "세부 목록을 선택하세요.")
            return
        detail_info = self.bok_detail_index_to_info.get(detail_idx, {})
        item_code1 = detail_info.get('ITEM_CODE', '')
        cycle = detail_info.get('CYCLE', '')

        # period values
        start_val = self.combo_period_start.currentText().strip() or ''
        end_val = self.combo_period_end.currentText().strip() or ''

        # prepare entry text for saved list (do not add until data fetched successfully)
        try:
            stat_name_text = self.bok_combo.currentText() or ''
            detail_text = self.bok_detail_combo.currentText() or ''
            entry = stat_name_text + (" :: " + detail_text if detail_text else "")
        except Exception:
            entry = ''

        # API params
        service = 'StatisticSearch'
        req_type = 'xml'
        lang = 'kr'
        start_idx = '1'
        end_cnt = '5000'  # per user request

        # Build path: base/service/key/type/lang/start/end/stat_code/주기/검색시작일자/검색종료일자/통계항목코드1/통계항목코드2/통계항목코드3/통계항목코드4
        base = 'https://ecos.bok.or.kr/api'
        parts = [base, service, key, req_type, lang, start_idx, end_cnt, stat_code, cycle, start_val, end_val]
        # 통계항목코드1..4: put item_code1 then placeholders
        parts.append(item_code1 or '?')
        parts.extend(['?', '?'])
        url = '/'.join(parts)

        try:
            resp = requests.get(url, timeout=30)
            resp.raise_for_status()
            data = resp.content
        except Exception as e:
            QMessageBox.critical(self, "요청 실패", f"StatisticSearch 요청 실패:\n{e}")
            return

        try:
            root = ET.fromstring(data)
        except Exception as e:
            QMessageBox.critical(self, "파싱 실패", f"응답 XML 파싱 실패:\n{e}")
            return

        nodes = root.findall('.//list') or root.findall('.//row') or root.findall('.//item')
        if not nodes:
            QMessageBox.information(self, "결과 없음", "조회된 데이터가 없습니다.")
            # do not clear existing table; simply return
            return

        # Determine column headers from first node children
        first = nodes[0]
        headers = [child.tag for child in list(first)]
        # unify headers order across nodes
        cols = headers[:]
        for n in nodes[1:]:
            for child in list(n):
                if child.tag not in cols:
                    cols.append(child.tag)

        # Populate table: append rows to existing table without deleting previous data
        try:
            # existing headers (if any)
            exist_col_count = self.bok_result_table.columnCount()
            existing_headers = [self.bok_result_table.horizontalHeaderItem(i).text() for i in range(exist_col_count)] if exist_col_count > 0 else []
            # union columns: preserve existing order then append new columns
            union_cols = existing_headers[:]
            for c in cols:
                if c not in union_cols:
                    union_cols.append(c)

            self.bok_result_table.setColumnCount(len(union_cols))
            self.bok_result_table.setHorizontalHeaderLabels(union_cols)

            existing_rows = self.bok_result_table.rowCount()
            self.bok_result_table.setRowCount(existing_rows + len(nodes))
            for r, node in enumerate(nodes):
                children = {c.tag: (c.text or '') for c in list(node)}
                for col in cols:
                    val = children.get(col, '')
                    cidx = union_cols.index(col)
                    # If this is the DATA_VALUE column, format with thousand separators
                    display_val = val
                    try:
                        if col == 'DATA_VALUE' and val is not None and val != '':
                            s = str(val).replace(',', '').strip()
                            sign = ''
                            if s.startswith(('+', '-')):
                                if s[0] == '-':
                                    sign = '-'
                                s = s[1:]
                            if '.' in s:
                                left, right = s.split('.', 1)
                                left_fmt = format(int(left), ',') if left.isdigit() else left
                                display_val = sign + left_fmt + '.' + right
                            else:
                                display_val = sign + format(int(s), ',') if s.isdigit() else val
                    except Exception:
                        display_val = val

                    item = QTableWidgetItem(display_val)
                    # numeric detection (store raw numeric value in UserRole)
                    try:
                        num = float(val.replace(',', ''))
                        item.setData(Qt.UserRole, num)
                    except Exception:
                        pass
                    self.bok_result_table.setItem(existing_rows + r, cidx, item)
            # after successfully appending rows, add the saved-list entry with a checked checkbox
            try:
                if entry:
                    item = QListWidgetItem()
                    widget = QWidget()
                    hl = QHBoxLayout()
                    chk = QCheckBox(entry)
                    chk.setChecked(True)
                    btn = QPushButton("삭제")
                    try:
                        btn.setFixedWidth(50)
                    except Exception:
                        pass
                    # connect delete with captured item
                    from functools import partial
                    btn.clicked.connect(partial(self._remove_saved_item, item))
                    hl.addWidget(chk)
                    hl.addWidget(btn)
                    hl.setContentsMargins(2, 2, 2, 2)
                    widget.setLayout(hl)
                    self.bok_listbox.addItem(item)
                    self.bok_listbox.setItemWidget(item, widget)
                    try:
                        item.setSizeHint(widget.sizeHint())
                    except Exception:
                        pass
                    if not hasattr(self, 'bok_saved_ranges'):
                        self.bok_saved_ranges = []
                    self.bok_saved_ranges.append({'label': entry, 'start': existing_rows, 'count': len(nodes)})
            except Exception:
                pass
            self.bok_result_table.resizeColumnsToContents()
        except Exception as e:
            QMessageBox.warning(self, "표시 실패", f"결과 표에 표시 중 오류:\n{e}")
            return

    def _remove_saved_item(self, item):
        try:
            idx = self.bok_listbox.row(item)
            if idx < 0:
                return
            # remove mapping
            mp = None
            try:
                mp = self.bok_saved_ranges.pop(idx)
            except Exception:
                mp = None
            # remove table rows corresponding to this saved range
            if mp:
                start = int(mp.get('start', 0))
                cnt = int(mp.get('count', 0))
                # clamp
                if cnt > 0:
                    for _ in range(cnt):
                        try:
                            self.bok_result_table.removeRow(start)
                        except Exception:
                            pass
                    # adjust subsequent saved_ranges' start indices
                    for r in self.bok_saved_ranges:
                        try:
                            if r.get('start', 0) > start:
                                r['start'] = max(0, int(r.get('start', 0)) - cnt)
                        except Exception:
                            pass
            # remove listbox item
            try:
                self.bok_listbox.takeItem(idx)
            except Exception:
                pass
        except Exception as e:
            # log traceback to debug_logs and show messagebox for easier debugging
            try:
                logs_dir = os.path.join(os.getcwd(), "debug_logs")
                os.makedirs(logs_dir, exist_ok=True)
                ts = int(time.time())
                fname = os.path.join(logs_dir, f"remove_saved_item_error_{ts}.log")
                with open(fname, 'w', encoding='utf-8') as fw:
                    fw.write("Error removing saved item:\n")
                    fw.write(str(e) + "\n\n")
                    traceback.print_exc(file=fw)
            except Exception:
                pass
            try:
                QMessageBox.critical(self, "삭제 오류", f"삭제 중 오류가 발생했습니다. 로그 파일을 확인하세요.\n{e}")
            except Exception:
                pass

    def on_bok_plot(self):
        try:
            col_count = self.bok_result_table.columnCount()
            headers = [
                (self.bok_result_table.horizontalHeaderItem(i).text()
                 if self.bok_result_table.horizontalHeaderItem(i) else '')
                for i in range(col_count)
            ]
            try:
                idx_time = headers.index('TIME')
            except ValueError:
                idx_time = None
            try:
                idx_val = headers.index('DATA_VALUE')
            except ValueError:
                idx_val = None
            try:
                idx_unit = headers.index('UNIT_NAME')
            except ValueError:
                idx_unit = None
            try:
                idx_stat = headers.index('STAT_NAME')
            except ValueError:
                idx_stat = None
            try:
                idx_item1 = headers.index('ITEM_NAME1')
            except ValueError:
                idx_item1 = None

            if idx_time is None or idx_val is None:
                QMessageBox.warning(self, "차트 생성 실패", "TIME 또는 DATA_VALUE 열을 찾을 수 없습니다.")
                return

            # Build per-saved-entry series (label, times[], values[]) so we can plot multiple series
            series_list = []
            unit = ''
            stat_name = ''
            item1 = ''
            rows = self.bok_result_table.rowCount()
            try:
                # If saved ranges exist and some are checked, build series from each checked saved range
                if getattr(self, 'bok_saved_ranges', None) and self.bok_listbox.count() > 0:
                    for i in range(self.bok_listbox.count()):
                        it = self.bok_listbox.item(i)
                        if it is None:
                            continue
                        w = self.bok_listbox.itemWidget(it)
                        if w is None:
                            continue
                        chk = w.findChild(QCheckBox)
                        if chk is None or not chk.isChecked():
                            continue
                        if i < len(self.bok_saved_ranges):
                            mp = self.bok_saved_ranges[i]
                            start = int(mp.get('start', 0))
                            cnt = int(mp.get('count', 0))
                            times_s = []
                            vals_s = []
                            for r in range(start, start + cnt):
                                t_item = self.bok_result_table.item(r, idx_time)
                                v_item = self.bok_result_table.item(r, idx_val)
                                t = t_item.text() if t_item else ''
                                v = v_item.text() if v_item else ''
                                try:
                                    num = float(str(v).replace(',', ''))
                                except Exception:
                                    num = None
                                # collect per-series
                                times_s.append(t)
                                vals_s.append(num)
                                if idx_unit is not None and not unit:
                                    u_item = self.bok_result_table.item(r, idx_unit)
                                    unit = u_item.text() if u_item else ''
                                if idx_stat is not None and not stat_name:
                                    s_item = self.bok_result_table.item(r, idx_stat)
                                    stat_name = s_item.text() if s_item else ''
                                if idx_item1 is not None and not item1:
                                    it_item = self.bok_result_table.item(r, idx_item1)
                                    item1 = it_item.text() if it_item else ''
                            label = mp.get('label', f'Series {i+1}')
                            series_list.append({'label': label, 'times': times_s, 'values': vals_s})
                else:
                    # fallback: single series containing all rows
                    times_s = []
                    vals_s = []
                    for r in range(rows):
                        t_item = self.bok_result_table.item(r, idx_time)
                        v_item = self.bok_result_table.item(r, idx_val)
                        t = t_item.text() if t_item else ''
                        v = v_item.text() if v_item else ''
                        try:
                            num = float(str(v).replace(',', ''))
                        except Exception:
                            num = None
                        times_s.append(t)
                        vals_s.append(num)
                        if idx_unit is not None and not unit:
                            u_item = self.bok_result_table.item(r, idx_unit)
                            unit = u_item.text() if u_item else ''
                        if idx_stat is not None and not stat_name:
                            s_item = self.bok_result_table.item(r, idx_stat)
                            stat_name = s_item.text() if s_item else ''
                        if idx_item1 is not None and not item1:
                            it_item = self.bok_result_table.item(r, idx_item1)
                            item1 = it_item.text() if it_item else ''
                    series_list.append({'label': 'Series 1', 'times': times_s, 'values': vals_s})
            except Exception:
                QMessageBox.information(self, "차트 없음", "플롯할 숫자 데이터가 없습니다.")
                return

            # Build union of time labels with granularity handling.
            # If any series contains monthly (or daily) granularity, produce a monthly
            # x-axis (YYYYMM). Annual-only values will be mapped to YYYY01 (Jan of year).
            import re as _re

            def to_yyyymm(t):
                try:
                    s = str(t).strip()
                    # YYYYMMDD -> YYYYMM
                    m = _re.match(r"^(\d{4})(\d{2})(\d{2})$", s)
                    if m:
                        return f"{m.group(1)}{m.group(2)}"
                    # YYYYMM
                    m = _re.match(r"^(\d{4})(\d{2})$", s)
                    if m:
                        return f"{m.group(1)}{m.group(2)}"
                    # YYYY-Qn or YYYYQn -> map to quarter start month
                    mq = _re.search(r"(\d{4})\D*Q(\d)", s, _re.IGNORECASE)
                    if mq:
                        y = int(mq.group(1)); q = int(mq.group(2)); mm = (q - 1) * 3 + 1
                        return f"{y}{mm:02d}"
                    # YYYY-MM or YYYY.MM or YYYY/MM or YYYY M formats
                    m2 = _re.match(r"^(\d{4})\D+(\d{1,2})", s)
                    if m2:
                        y = int(m2.group(1)); mm = int(m2.group(2))
                        return f"{y}{mm:02d}"
                    # pure year YYYY -> map to January of that year
                    if len(s) >= 4 and s[:4].isdigit():
                        return f"{s[:4]}01"
                except Exception:
                    pass
                return None

            has_monthly = False
            for ser in series_list:
                for t in ser.get('times', []):
                    if t is None:
                        continue
                    # if any time maps to a yyyymm (non-None) and original contains month/day/quarter, treat as monthly axis
                    s = str(t).strip()
                    if _re.match(r"^\d{6}$", _re.sub(r"\D", "", s)) and len(_re.sub(r"\D", "", s)) >= 6:
                        has_monthly = True
                        break
                    if _re.search(r"Q", s, _re.IGNORECASE):
                        has_monthly = True
                        break
                if has_monthly:
                    break

            if has_monthly:
                # build monthly union keys (YYYYMM)
                time_set = set()
                for ser in series_list:
                    for t in ser.get('times', []):
                        k = to_yyyymm(t)
                        if k:
                            time_set.add(k)
                union_times = sorted(time_set, key=lambda x: int(x))
            else:
                # fallback to year-only axis (as before)
                time_map = {}
                for ser in series_list:
                    for t in ser.get('times', []):
                        if t is None:
                            continue
                        # extract year
                        m = _re.search(r"(\d{4})", str(t))
                        if m:
                            kd = int(m.group(1))
                            if kd not in time_map:
                                time_map[kd] = str(kd)
                union_times = [lbl for kd, lbl in sorted(time_map.items(), key=lambda kv: kv[0])]
            if not union_times:
                QMessageBox.information(self, "차트 없음", "플롯할 숫자 데이터가 없습니다.")
                return

            # map each series to union_times indices
            x = list(range(len(union_times)))
            xticks = union_times
            plotted_series = []
            for ser in series_list:
                ymap = [None] * len(union_times)
                for t, v in zip(ser.get('times', []), ser.get('values', [])):
                    if t is None:
                        continue
                    if has_monthly:
                        k = to_yyyymm(t)
                    else:
                        # year-only mapping
                        m = _re.search(r"(\d{4})", str(t))
                        k = m.group(1) if m else None
                    if k is None:
                        continue
                    try:
                        idx = union_times.index(k)
                        ymap[idx] = v
                    except ValueError:
                        continue
                plotted_series.append({'label': ser.get('label', ''), 'y': ymap})

            # Attempt to set a font that supports Korean on Windows/Mac/Linux
            try:
                available = {f.name for f in fm.fontManager.ttflist}
                candidates = ['Malgun Gothic', '맑은 고딕', 'NanumGothic', 'AppleGothic', 'Noto Sans CJK KR', 'Noto Sans CJK JP', 'DejaVu Sans']
                use_font = next((c for c in candidates if c in available), None)
                if use_font:
                    plt.rcParams['font.family'] = use_font
                plt.rcParams['axes.unicode_minus'] = False
            except Exception:
                pass

            fig = plt.figure(figsize=(10, 5))
            base_ax = fig.add_subplot(111)
            # force default marker size to 3 for compact visuals; use doubled size when selected
            marker_size = 3
            try:
                plt.rcParams['lines.markersize'] = marker_size
            except Exception:
                pass

            # create one axis per series (shared x-axis). The first uses the base_ax,
            # subsequent axes are created via twinx and their spines shifted right.
            try:
                colors = plt.rcParams.get('axes.prop_cycle').by_key().get('color', ['C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7'])
            except Exception:
                colors = ['C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7']

            nseries = len(plotted_series)
            axes = []
            for si in range(nseries):
                if si == 0:
                    ax_i = base_ax
                else:
                    ax_i = base_ax.twinx()
                    # shift the spine to the right to avoid overlapping
                    try:
                        ax_i.spines['right'].set_position(('axes', 1.0 + 0.12 * (si - 1)))
                    except Exception:
                        pass
                axes.append(ax_i)

            # if many axes, expand right margin
            try:
                if nseries > 1:
                    right_margin = 0.85 + 0.12 * max(0, nseries - 2)
                    fig.subplots_adjust(right=min(0.98, right_margin))
            except Exception:
                pass

            plotted_series_updated = []
            for si, ser in enumerate(plotted_series):
                ys = ser['y']
                # convert None to NaN so matplotlib skips plotting those points
                ys_plot = [float(v) if v is not None else float('nan') for v in ys]
                ax_i = axes[si]
                color = colors[si % len(colors)]
                try:
                    import math
                    xi = [i for i, vv in enumerate(ys_plot) if not math.isnan(vv)]
                    yi = [ys_plot[i] for i in xi]
                    if xi:
                        line, = ax_i.plot(xi, yi, marker='o', linestyle='-', label=ser.get('label', ''), color=color)
                    else:
                        # no valid points
                        line = None
                    if line is not None:
                        line.set_markersize(marker_size)
                except Exception:
                    try:
                        line, = ax_i.plot(x, ys_plot, marker='o', linestyle='-', label=ser.get('label', ''))
                        line.set_markersize(marker_size)
                    except Exception:
                        line = None

                # label each y-axis with series label (and unit if available)
                ylbl = ser.get('label', '')
                if unit:
                    ylbl = f"{ylbl} ({unit})" if ylbl else unit
                try:
                    ax_i.set_ylabel(ylbl)
                    # set y-axis label and tick colors to match series color
                    try:
                        ax_i.yaxis.label.set_color(color)
                        ax_i.tick_params(axis='y', colors=color)
                    except Exception:
                        pass
                except Exception:
                    pass

                # per-axis y formatter
                try:
                    ax_i.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: format(int(v), ',') if abs(v - round(v)) < 1e-6 else format(v, ',.2f')))
                except Exception:
                    pass

                # create per-series annotation for tooltips
                try:
                    ann = ax_i.annotate("", xy=(0, 0), xytext=(15, 15), textcoords="offset points",
                                        bbox=dict(boxstyle="round", fc="w"), arrowprops=dict(arrowstyle="->"))
                    ann.set_visible(False)
                except Exception:
                    ann = None

                plotted_series_updated.append({'label': ser.get('label', ''), 'y': ys, 'y_plot': ys_plot, 'ax': ax_i, 'line': line, 'annot': ann})

            plotted_series = plotted_series_updated

            # Reduce number of x-tick labels if too many
            max_xticks = 20
            # Prefer year-aligned ticks: for monthly data use January indices; for quarterly use Q1.
            import re
            month_indices = []
            quarter_indices = []
            for i, t in enumerate(xticks):
                s = str(t)
                # YYYYMM e.g., 201901 or 2019-01 variants
                m = re.match(r"^(\d{4})(\d{2})$", s)
                if m:
                    mm = int(m.group(2))
                    if mm == 1:
                        month_indices.append(i)
                    continue
                # look for patterns like 2019Q1 or 2019-Q1
                mq = re.search(r"(\d{4})\D*Q(\d)", s, re.IGNORECASE)
                if mq:
                    q = int(mq.group(2))
                    if q == 1:
                        quarter_indices.append(i)

            if month_indices:
                cand = month_indices
            elif quarter_indices:
                cand = quarter_indices
            else:
                cand = list(range(len(x)))

            # If too many candidate ticks, thin them out to respect max_xticks
            if len(cand) > max_xticks:
                step = max(1, len(cand) // max_xticks)
                visible_x = [cand[i] for i in range(0, len(cand), step)]
            else:
                visible_x = cand

            # fallback: if visible_x empty, use automatic thinning over full range
            if not visible_x:
                if len(x) > max_xticks:
                    step = max(1, len(x) // max_xticks)
                    visible_x = x[::step]
                else:
                    visible_x = x

            visible_labels = [xticks[i] for i in visible_x]

            base_ax.set_xticks(visible_x)
            # shorten labels to avoid overlapping long texts
            def shorten(s, n=30):
                s = str(s)
                return s if len(s) <= n else s[:n-3] + '...'

            # convert TIME-like labels to year only (e.g., 202001->2020, 2010Q1->2010)
            import re
            def year_label(t):
                try:
                    if not t:
                        return ''
                    m = re.search(r'(\d{4})', str(t))
                    if m:
                        return m.group(1)
                    return str(t)
                except Exception:
                    return str(t)

            year_labels = [year_label(lbl) for lbl in visible_labels]
            base_ax.set_xticklabels([shorten(lbl, 20) for lbl in year_labels], rotation=45, ha='right')

            # bottom label with STAT_NAME and ITEM_NAME1 (truncate if very long)
            bottom_label = ' '.join(filter(None, [stat_name, item1]))
            if len(bottom_label) > 120:
                bottom_label = bottom_label[:117] + '...'

            # create legend per plotted series; place below plot
            try:
                if plotted_series:
                    handles = [p['line'] for p in plotted_series if p.get('line') is not None]
                    labels = [p.get('label', '') for p in plotted_series if p.get('line') is not None]
                    leg = base_ax.legend(handles, labels, loc='lower center', bbox_to_anchor=(0.5, -0.25), ncol=1, frameon=False)
                    for text in leg.get_texts():
                        text.set_fontsize(10)
                else:
                    if bottom_label:
                        fig.text(0.5, 0.01, bottom_label, ha='center', fontsize=10)
            except Exception:
                if bottom_label:
                    fig.text(0.5, 0.01, bottom_label, ha='center', fontsize=10)

            # draw faint major grid lines on the base axis
            try:
                base_ax.grid(True, which='major', color='gray', linestyle='-', linewidth=0.5, alpha=0.3)
            except Exception:
                pass

            fig.tight_layout(rect=[0, 0.08, 1, 1])

            def format_num(v):
                try:
                    if v is None:
                        return ''
                    if abs(v - round(v)) < 1e-6:
                        return format(int(round(v)), ',')
                    return format(v, ',.2f')
                except Exception:
                    return str(v)

            def update_annot(series_idx, xi):
                try:
                    ser = plotted_series[series_idx]
                    y_val = ser['y'][xi]
                    if y_val is None:
                        return
                    x_val = x[xi]
                    ann = ser.get('annot')
                    if ann is None:
                        return
                    ann.xy = (x_val, y_val)
                    label_x = xticks[xi]
                    lbl = ser.get('label', '')
                    txt = f"{lbl}\n{label_x}\n{format_num(y_val)} {unit if unit else ''}".strip()
                    ann.set_text(txt)
                    try:
                        ann.get_bbox_patch().set_alpha(0.9)
                    except Exception:
                        pass
                except Exception:
                    pass

            # selection state for ctrl-click comparisons
            selected_idxs = []  # list of (series_idx, x_idx) tuples
            sel_artists = []
            range_annot = fig.text(0.02, 0.95, '', transform=fig.transFigure, va='top', fontsize=9,
                                   bbox=dict(boxstyle='round', facecolor='w', alpha=0.9))
            range_annot.set_visible(False)

            def on_move(event):
                # show per-series annot when cursor near a point on any axis
                any_visible = False
                if event.inaxes in axes and event.xdata is not None and event.ydata is not None:
                    min_dist = float('inf')
                    nearest = (None, None)
                    for si, ser in enumerate(plotted_series):
                        for xi, yv in enumerate(ser['y']):
                            if yv is None:
                                continue
                            try:
                                xpix, ypix = ser['ax'].transData.transform((x[xi], yv))
                            except Exception:
                                continue
                            dx = xpix - event.x
                            dy = ypix - event.y
                            dist = (dx * dx + dy * dy) ** 0.5
                            if dist < min_dist:
                                min_dist = dist
                                nearest = (si, xi)
                    if nearest[0] is not None and min_dist < 10:
                        # hide all annots first
                        for ser in plotted_series:
                            a = ser.get('annot')
                            if a is not None:
                                a.set_visible(False)
                        update_annot(nearest[0], nearest[1])
                        a = plotted_series[nearest[0]].get('annot')
                        if a is not None:
                            a.set_visible(True)
                        fig.canvas.draw_idle()
                        any_visible = True
                if not any_visible:
                    for ser in plotted_series:
                        a = ser.get('annot')
                        if a is not None and a.get_visible():
                            a.set_visible(False)
                            fig.canvas.draw_idle()

            def on_click(event):
                try:
                    if event.inaxes not in axes or event.button != 1:
                        return
                    min_dist = float('inf')
                    nearest = (None, None)
                    for si, ser in enumerate(plotted_series):
                        for xi, yv in enumerate(ser['y']):
                            if yv is None:
                                continue
                            try:
                                xpix, ypix = ser['ax'].transData.transform((x[xi], yv))
                            except Exception:
                                continue
                            dx = xpix - event.x
                            dy = ypix - event.y
                            dist = (dx * dx + dy * dy) ** 0.5
                            if dist < min_dist:
                                min_dist = dist
                                nearest = (si, xi)
                    if nearest[0] is None or min_dist >= 10:
                        return

                    si, xi = nearest
                    sel_key = (si, xi)
                    found = None
                    for idx, p in enumerate(selected_idxs):
                        if p == sel_key:
                            found = idx
                            break
                    if found is not None:
                        selected_idxs.pop(found)
                        art = sel_artists.pop(found)
                        try:
                            art.remove()
                        except Exception:
                            pass
                    else:
                        if len(selected_idxs) >= 2:
                            selected_idxs.pop(0)
                            art = sel_artists.pop(0)
                            try:
                                art.remove()
                            except Exception:
                                pass
                        selected_idxs.append(sel_key)
                        yv = plotted_series[si]['y'][xi]
                        art_ax = plotted_series[si]['ax']
                        art, = art_ax.plot(x[xi], yv, marker='o', markersize=marker_size*2, markerfacecolor='red', markeredgecolor='red', markeredgewidth=1.0)
                        sel_artists.append(art)
                        try:
                            update_annot(si, xi)
                            a = plotted_series[si].get('annot')
                            if a is not None:
                                a.set_visible(True)
                        except Exception:
                            pass

                    # when two points selected, compute stats based on their x indices and series
                    if len(selected_idxs) == 2:
                        (s1, i1), (s2, i2) = selected_idxs[0], selected_idxs[1]
                        if i1 > i2:
                            s1, s2, i1, i2 = s2, s1, i2, i1
                        v1 = plotted_series[s1]['y'][i1]
                        v2 = plotted_series[s2]['y'][i2]
                        lbl1 = xticks[i1]
                        lbl2 = xticks[i2]

                        def parse_year_month(t):
                            try:
                                s = str(t).strip()
                                if len(s) >= 6 and s[:6].isdigit():
                                    return (int(s[:4]), int(s[4:6]))
                                import re
                                m_qu = re.match(r"^(\d{4})\D*Q(\d)", s, re.IGNORECASE)
                                if m_qu:
                                    y = int(m_qu.group(1)); q = int(m_qu.group(2)); month = (q - 1) * 3 + 1; return (y, month)
                                if len(s) >= 4 and s[:4].isdigit():
                                    return (int(s[:4]), 1)
                                return (None, None)
                            except Exception:
                                return (None, None)

                        y1, m1 = parse_year_month(lbl1)
                        y2, m2 = parse_year_month(lbl2)
                        years = None
                        if y1 is not None and y2 is not None:
                            months_diff = (y2 - y1) * 12 + (m2 - m1)
                            years = months_diff / 12.0 if months_diff != 0 else 0

                        pct = None
                        cagr = None
                        if v1 is not None and v2 is not None and v1 != 0:
                            try:
                                pct = (v2 - v1) / v1 * 100.0
                            except Exception:
                                pct = None
                        if pct is not None and years and years > 0 and v1 > 0:
                            try:
                                cagr = ((v2 / v1) ** (1.0 / years) - 1.0) * 100.0
                            except Exception:
                                cagr = None

                        pct_txt = f"{pct:.2f}%" if pct is not None else 'N/A'
                        cagr_txt = f"{cagr:.2f}%" if cagr is not None else 'N/A'
                        yrs_txt = f"{years:.2f}년" if years is not None else 'N/A'

                        range_text = f"기간: {lbl1} → {lbl2}\n기간 차이: {yrs_txt}\n증가율: {pct_txt}\n연평균: {cagr_txt}"
                        range_annot.set_text(range_text)
                        range_annot.set_visible(True)
                        fig.canvas.draw_idle()
                    else:
                        range_annot.set_visible(False)
                        fig.canvas.draw_idle()
                except Exception:
                    pass

            fig.canvas.mpl_connect("motion_notify_event", on_move)
            fig.canvas.mpl_connect("button_press_event", on_click)
            plt.show()
        except Exception as e:
            QMessageBox.warning(self, "차트 생성 오류", f"오류:\n{e}")

    def on_bok_select(self):
        sel = self.bok_combo.currentIndex()
        if sel < 0:
            return
        code = self.bok_index_to_code.get(sel, '')
        try:
            # show selected code in status label and copy to clipboard
            self.status_label.setText(f"선택 통계표코드: {code}")
            QApplication.clipboard().setText(code or "")
        except Exception:
            pass

        # 세부 목록 조회: 서비스는 StatisticItemList, 요청종료건수는 100
        try:
            self._load_stat_item_list(code)
        except Exception:
            pass

    def on_apt_cancel(self):
        if getattr(self, '_apt_worker', None) is None:
            return
        try:
            self.status_label.setText('취소 요청 중...')
            self._apt_worker.stop()
            self.btn_apt_cancel.setEnabled(False)
        except Exception:
            pass

    # =========== 아파트 실거래 관련 메서드 ===========
    def on_apt_fetch(self):
        lawd = self.edit_apt_lawd.text().strip()
        key = self.edit_apt_key.text().strip()
        # 년/월 콤보에서 YYYYMM 문자열을 조합
        from_year = self.combo_apt_year_from.currentText().strip()
        from_month = self.combo_apt_month_from.currentText().strip()
        to_year = self.combo_apt_year_to.currentText().strip()
        to_month = self.combo_apt_month_to.currentText().strip()
        from_ym = f"{from_year}{from_month}"
        to_ym = f"{to_year}{to_month}"

        # Allow searching by selected `시/도` when `지역코드(LAWD_CD)` is not manually provided.
        # If `lawd` is empty but a `시/도` is selected, construct a list of 5-digit LAWD codes
        # from `self.sido_map` and iterate over them. Otherwise require `lawd` as before.
        if not key or not from_ym or not to_ym:
            QMessageBox.warning(self, "입력 오류", "Service Key와 조회기간(시작/종료)은 필수입니다.")
            return

        # build lawd_list: prefer explicit 5-digit `lawd` if provided.
        # If user provided a short code (e.g. '28' from 시/도), treat it as non-specific
        # and derive 5-digit LAWD codes from `sigungu_map` (or `sido_map` as fallback).
        lawd_list = []
        try:
            sel_sido = self.combo_sido.currentText()
        except Exception:
            sel_sido = None

        if lawd and len(lawd) >= 5:
            lawd_list = [lawd[:5]]
        else:
            # derive from sigungu_map if possible
            try:
                codes_map = getattr(self, 'sigungu_map', {}) or {}
                seen = set()
                if sel_sido and sel_sido not in ("선택", "없음") and codes_map:
                    # sigungu_map maps names->list(codes)
                    for lst in codes_map.values():
                        for c in (lst or []):
                            if not c:
                                continue
                            prefix = (c[:5] or "").strip()
                            if prefix and prefix not in seen:
                                seen.add(prefix)
                                lawd_list.append(prefix)
                else:
                    # fallback: try sido_map entries for the selected sido
                    codes = getattr(self, 'sido_map', {}).get(sel_sido) or []
                    for c in codes:
                        if not c:
                            continue
                        prefix = (c[:5] or "").strip()
                        if prefix and prefix not in seen:
                            seen.add(prefix)
                            lawd_list.append(prefix)
            except Exception:
                lawd_list = []

        if not lawd_list:
            QMessageBox.warning(self, "입력 오류", "지역코드가 없거나 선택된 시/도가 없습니다. 지역코드를 입력하거나 시/도를 선택하세요.")
            return

        # prevent duplicate start if worker already running
        if getattr(self, '_apt_worker', None) is not None and getattr(self._apt_worker, 'isRunning', lambda: False)():
            QMessageBox.information(self, "진행 중", "이미 조회가 진행 중입니다.")
            return

        # reset filters and master rows for a fresh fetch
        try:
            self.apt_filters = {}
            self.apt_rows_master = []
        except Exception:
            pass

        # clear previous table data immediately so UI shows fresh state
        try:
            self.apt_table.setRowCount(0)
            try:
                self.apt_table.clearContents()
            except Exception:
                pass
            self.btn_apt_save.setEnabled(False)
            if hasattr(self, 'edit_apt_url'):
                self.edit_apt_url.setText("")
            try:
                self.status_label.setText("대기")
                self.progress_bar.setValue(0)
            except Exception:
                pass
        except Exception:
            pass

        # 구성될 최종 URL을 생성하여 표시 (pageNo 먼저, numOfRows 고정)
        api_url = "https://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
        # serviceKey: decode percent-encoding before passing to `params`
        # to avoid double-encoding by `requests` (matches browser behavior).
        # Show an example final URL using the first LAWD code
        params = {
            "serviceKey": requests.utils.unquote(key),
            "LAWD_CD": (lawd_list[0] if lawd_list else lawd),
            "DEAL_YMD": from_ym,
            "pageNo": "1",
            "numOfRows": "1000",
        }
        try:
            req = requests.Request('GET', api_url, params=params)
            prep = requests.Session().prepare_request(req)
            final_url = prep.url
        except Exception:
            final_url = api_url
        if hasattr(self, 'edit_apt_url'):
            self.edit_apt_url.setText(final_url)

        # 기간(start->end) 내 모든 월을 순회하여 데이터 수집
        months = self._months_between(from_ym, to_ym)
        all_rows = []
        total_months = len(months)
        # 준비: 진행바 설정 및 버튼 비활성화
        try:
            total_steps = total_months * max(1, len(lawd_list))
            self.progress_bar.setMaximum(total_steps)
            self.progress_bar.setValue(0)
            self.status_label.setText(f"진행: 0/{total_steps}")
        except Exception:
            pass
        self.btn_apt_fetch.setEnabled(False)
        self.btn_apt_save.setEnabled(False)

        # Run data fetching in a worker thread to avoid blocking the UI
        try:
            total_steps = total_months * max(1, len(lawd_list))
            self.progress_bar.setMaximum(total_steps)
            self.progress_bar.setValue(0)
            self.status_label.setText(f"진행: 0/{total_steps}")
        except Exception:
            pass

        self.btn_apt_fetch.setEnabled(False)
        self.btn_apt_save.setEnabled(False)

        # create worker (pass list of LAWD codes)
        self._apt_worker = AptFetchWorker(lawd_list, months, key)

        def _on_progress(cur, total):
            try:
                self.progress_bar.setValue(cur)
                self.status_label.setText(f"진행: {cur}/{total}")
                QApplication.processEvents()
            except Exception:
                pass

        def _on_finished(rows):
            try:
                # sort by 계약일
                def _parse_date(s):
                    try:
                        parts = s.split('-')
                        y = int(parts[0]); mo = int(parts[1]); d = int(parts[2])
                        return datetime.date(y, mo, d)
                    except Exception:
                        return datetime.date.min
                # If 읍면동 콤보에 값이 선택되어 있으면, 법정동(컬럼 인덱스 7)과 같은 행만 남깁니다.
                try:
                    sel_dong = self.combo_dong.currentText()
                    if sel_dong and sel_dong not in ("선택", "없음"):
                        rows = [r for r in rows if (r[7] or "").strip() == sel_dong]
                except Exception:
                    pass

                # trade_date is at column index 3 in worker row structure
                rows.sort(key=lambda r: _parse_date(r[3] if len(r) > 3 else ""))
                self.apt_rows_master = rows
                self.apt_filters = {}
                self.populate_apt_table(rows)
                self.btn_apt_save.setEnabled(bool(rows))
                self.status_label.setText(f"완료: {len(rows)}건")
                try:
                    # set progress to full (maximum may be months * LAWD count)
                    self.progress_bar.setValue(self.progress_bar.maximum())
                except Exception:
                    pass
            finally:
                self.btn_apt_fetch.setEnabled(True)
                self.btn_apt_cancel.setEnabled(False)
                self._apt_worker = None

        def _on_error(msg):
            QMessageBox.critical(self, "요청 실패", msg)
            try:
                self.btn_apt_fetch.setEnabled(True)
                self.btn_apt_save.setEnabled(False)
                self.btn_apt_cancel.setEnabled(False)
            except Exception:
                pass

        self._apt_worker.progress.connect(_on_progress)
        self._apt_worker.finished.connect(_on_finished)
        self._apt_worker.error.connect(_on_error)
        self._apt_worker.start()
        # enable cancel button while running
        try:
            self.btn_apt_cancel.setEnabled(True)
        except Exception:
            pass

        # Note: table will be populated by worker's finished handler;
        # avoid immediate clearing/populating here to prevent race conditions.

    def get_apt_trade_data(self, lawd_cd, deal_ymd, service_key):
        url = "https://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
        # serviceKey must be URL-decoded to avoid double-encoding
        params = {
            "serviceKey": requests.utils.unquote(service_key),
            "LAWD_CD": lawd_cd,
            "DEAL_YMD": deal_ymd,
            "pageNo": "1",
            "numOfRows": "1000",
        }

        try:
            # mimic browser headers to reduce server-side differences
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
                "Accept": "application/xml, text/xml, */*;q=0.01",
            }
            resp = requests.get(url, params=params, timeout=20, headers=headers)
            resp.raise_for_status()
            # save debug response into debug_logs
            try:
                logs_dir = os.path.join(os.getcwd(), "debug_logs")
                os.makedirs(logs_dir, exist_ok=True)
                ts = int(time.time())
                fname = os.path.join(logs_dir, f"debug_response_{lawd_cd}_{deal_ymd}_{ts}.xml")
                with open(fname, "wb") as fw:
                    fw.write(resp.content)
                meta = os.path.join(logs_dir, f"debug_response_{lawd_cd}_{deal_ymd}_{ts}.meta.txt")
                with open(meta, "w", encoding='utf-8') as fm:
                    fm.write(f"url: {resp.url}\nstatus: {resp.status_code}\nheaders: {dict(resp.headers)}\n")
            except Exception:
                pass
        except Exception as e:
            QMessageBox.critical(self, "요청 실패", str(e))
            return None

        try:
            root = ET.fromstring(resp.content)
        except Exception as e:
            QMessageBox.critical(self, "파싱 오류", f"응답 XML 파싱 실패: {e}")
            return None
        header = root.find("header")
        code = None
        msg = None
        if header is not None:
            code = header.findtext("resultCode")
            msg = header.findtext("resultMsg")

        items = root.findall("body/items/item")
        # 일부 API는 header의 resultCode가 '00'이 아니더라도 items를 반환할 수 있습니다.
        # items가 존재하면 데이터를 우선 사용하고, 비어있을 때만 에러로 처리합니다.
        if not items:
            if code is not None and code != "00":
                QMessageBox.warning(self, "API 오류", f"{code}: {msg}")
            else:
                QMessageBox.information(self, "결과 없음", "조회 결과가 없습니다.")
            return None
        def _find_text(item, candidates):
            # Try explicit tag names first, then fallback to substring match on tag names
            for cand in candidates:
                try:
                    v = item.findtext(cand)
                except Exception:
                    v = None
                if v:
                    return v
            # fallback: search child tags for keyword substring
            low_cands = [c.lower() for c in candidates]
            for child in list(item):
                t = (child.tag or "").lower()
                for lc in low_cands:
                    if lc in t:
                        return child.text or ""
            return ""

        def _norm_amount(s):
            if not s:
                return ""
            try:
                import re
                digits = re.sub(r"[^0-9]", "", s)
                return digits
            except Exception:
                return s

        def _norm_rgst(s):
            if not s:
                return ""
            try:
                import re
                s2 = s.strip()
                # two-digit year like 25.12.04 -> 2025-12-04
                m = re.match(r"^(\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})$", s2)
                if m:
                    yy = int(m.group(1))
                    yyyy = 2000 + yy if yy < 100 else yy
                    mm = int(m.group(2)); dd = int(m.group(3))
                    return f"{yyyy:04d}-{mm:02d}-{dd:02d}"
                m2 = re.match(r"^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})$", s2)
                if m2:
                    yyyy = int(m2.group(1)); mm = int(m2.group(2)); dd = int(m2.group(3))
                    return f"{yyyy:04d}-{mm:02d}-{dd:02d}"
                return s2
            except Exception:
                return s

        rows = []
        for it in items:
            trade_date = f"{it.findtext('dealYear') or ''}-{it.findtext('dealMonth') or ''}-{it.findtext('dealDay') or ''}"
            # normalize amount and registration date
            raw_amount = it.findtext("dealAmount") or ""
            amount_norm = _norm_amount(raw_amount)
            rgst_raw = it.findtext("rgstDate") or _find_text(it, ["registDay", "등기일자", "registrationDate", "rgstDate"])
            rgst_norm = _norm_rgst(rgst_raw)

            row = [
                it.findtext("aptNm") or "",
                # 아파트동
                it.findtext("aptDong") or _find_text(it, ["aptDong", "단지동", "동", "apt_dong"]),
                it.findtext("excluUseAr") or "",
                trade_date,
                amount_norm,
                it.findtext("floor") or "",
                it.findtext("buildYear") or "",
                it.findtext("umdNm") or "",
                it.findtext("jibun") or "",
                it.findtext("sggCd") or "",
                # 추가 필드들: 우선 명시적 태그 사용, 없으면 후보 탐색
                (it.findtext("dealingGbn") or _find_text(it, ["tradeType", "거래유형", "dealType", "dealingGbn"])),
                (it.findtext("estateAgentSggNm") or _find_text(it, ["bcnstAddr", "brokerAddr", "중개사소재지", "bcnstc", "estateAgentSggNm"])),
                rgst_norm,
                (it.findtext("slerGbn") or _find_text(it, ["seller", "거래주체정보_매도자", "매도자", "tradePartSeller", "slerGbn"])),
                _find_text(it, ["buyer", "거래주체정보_매수자", "매수자", "tradePartBuyer"]),
                _find_text(it, ["rentYn", "토지임대부", "landLease", "isLandLeaseApt"]),
            ]
            rows.append(row)

        return rows

    # -- Header filter handlers for apt_table --
    def apt_header_context_menu(self, pos):
        hdr = self.apt_table.horizontalHeader()
        col = hdr.logicalIndexAt(pos)
        if col < 0:
            return
        # toggle: if filter exists for this column, remove it; else ask for input
        if col in self.apt_filters and self.apt_filters[col]:
            # remove filter
            try:
                del self.apt_filters[col]
            except Exception:
                self.apt_filters.pop(col, None)
            self.apply_apt_filters()
            self.status_label.setText(f"필터 해제: 열 {col}")
            return

        text, ok = QInputDialog.getText(self, "열 필터", f"열 {col} 필터 텍스트 입력:")
        if not ok:
            return
        text = text.strip()
        if not text:
            return
        self.apt_filters[col] = text
        self.apply_apt_filters()
        self.status_label.setText(f"필터 적용: 열 {col} -> '{text}'")

    def apply_apt_filters(self):
        if not getattr(self, 'apt_rows_master', None):
            return
        if not self.apt_filters:
            rows = self.apt_rows_master
        else:
            def matches_all(row):
                for col, f in self.apt_filters.items():
                    try:
                        if f.lower() not in (row[col] or "").lower():
                            return False
                    except Exception:
                        return False
                return True
            rows = [r for r in self.apt_rows_master if matches_all(r)]
        self.populate_apt_table(rows)

    def _months_between(self, from_ym, to_ym):
        """Return list of YYYYMM strings from from_ym to to_ym inclusive, ascending."""
        try:
            fy = int(from_ym[:4]); fm = int(from_ym[4:6])
            ty = int(to_ym[:4]); tm = int(to_ym[4:6])
        except Exception:
            return []
        # normalize order
        start = fy * 100 + fm
        end = ty * 100 + tm
        if start > end:
            start, end = end, start
        months = []
        y = start // 100
        m = start % 100
        while True:
            months.append(f"{y}{m:02d}")
            if y * 100 + m >= end:
                break
            m += 1
            if m > 12:
                m = 1
                y += 1
        return months

    def populate_apt_table(self, rows):
        # disable sorting while populating to avoid race/ordering issues
        try:
            was_sorting = self.apt_table.isSortingEnabled()
        except Exception:
            was_sorting = False
        try:
            self.apt_table.setSortingEnabled(False)
        except Exception:
            pass
        try:
            self.apt_table.clearContents()
        except Exception:
            pass
        self.apt_table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                try:
                    txt = val if isinstance(val, str) else str(val)
                except Exception:
                    txt = ""
                # format 거래금액(만원) column with thousand separators (column index 4)
                if c == 4:
                    num_val = None
                    try:
                        import re
                        digits = re.sub(r"[^0-9]", "", txt)
                        if digits:
                            num_val = int(digits)
                            txt = f"{num_val:,}"
                    except Exception:
                        num_val = None
                    try:
                        if num_val is not None:
                            it = NumericItem(txt)
                            it.setData(Qt.UserRole, num_val)
                        else:
                            it = QTableWidgetItem(txt)
                    except Exception:
                        it = QTableWidgetItem(txt)
                    self.apt_table.setItem(r, c, it)
                else:
                    self.apt_table.setItem(r, c, QTableWidgetItem(txt))
        try:
            self.apt_table.setSortingEnabled(was_sorting)
        except Exception:
            pass

    def on_apt_save_csv(self):
        lawd = self.edit_apt_lawd.text().strip()
        from_year = self.combo_apt_year_from.currentText().strip()
        from_month = self.combo_apt_month_from.currentText().strip()
        to_year = self.combo_apt_year_to.currentText().strip()
        to_month = self.combo_apt_month_to.currentText().strip()
        from_ym = f"{from_year}{from_month}"
        to_ym = f"{to_year}{to_month}"
        filename = f"apt_trade_{lawd}_{from_ym}_{to_ym}.csv"
        try:
            with open(filename, "w", newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                headers = [self.apt_table.horizontalHeaderItem(i).text() for i in range(self.apt_table.columnCount())]
                writer.writerow(headers)
                for r in range(self.apt_table.rowCount()):
                    row = [self.apt_table.item(r, c).text() if self.apt_table.item(r, c) else "" for c in range(self.apt_table.columnCount())]
                    writer.writerow(row)
            QMessageBox.information(self, "저장 성공", f"{filename} 으로 저장되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "저장 실패", str(e))


class AptTradeGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("아파트 실거래 조회")

        lbl_key = QLabel("Service Key:")
        self.edit_key = QLineEdit()
        self.edit_key.setPlaceholderText("발급받은 인증키를 입력하세요 (URL 디코딩된 값 권장)")

        lbl_lawd = QLabel("지역코드(LAWD_CD):")
        self.edit_lawd = QLineEdit()
        self.edit_lawd.setPlaceholderText("예: 11110")

        lbl_ymd = QLabel("조회년월(DEAL_YMD, YYYYMM):")
        self.edit_ymd = QLineEdit()
        self.edit_ymd.setPlaceholderText("예: 202401")

        self.btn_fetch = QPushButton("조회")
        self.btn_fetch.clicked.connect(self.on_fetch)

        self.btn_save = QPushButton("CSV로 저장")
        self.btn_save.clicked.connect(self.on_save_csv)
        self.btn_save.setEnabled(False)

        self.table = QTableWidget()
        headers = ["아파트명", "전용면적", "계약일", "거래금액(만원)", "층", "건축년도", "법정동", "지번", "지역코드"]
        # 기존 AptTradeGUI 테이블 헤더도 아파트동과 추가 필드를 포함하도록 확장
        headers = [
            "아파트명", "아파트동", "전용면적", "계약일", "거래금액(만원)", "층", "건축년도", "법정동", "지번", "지역코드",
            "거래유형", "중개사소재지", "등기일자", "거래주체_매도자", "거래주체_매수자", "토지임대부여부"
        ]
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        th = self.table.horizontalHeader()
        th.setSectionResizeMode(QHeaderView.Interactive)
        th.setSectionsMovable(True)
        th.setStretchLastSection(False)
        # 헤더 클릭으로 열 정렬을 활성화
        self.table.setSortingEnabled(True)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        # header right-click filter
        th.setContextMenuPolicy(Qt.CustomContextMenu)
        th.customContextMenuRequested.connect(self.aptgui_header_context_menu)
        self.table_rows_master = []
        self.table_filters = {}
        # AptTradeGUI table header: context menu for filtering set in its init_ui

        layout = QGridLayout()
        layout.addWidget(lbl_key, 0, 0)
        layout.addWidget(self.edit_key, 0, 1)
        layout.addWidget(lbl_lawd, 1, 0)
        layout.addWidget(self.edit_lawd, 1, 1)
        layout.addWidget(lbl_ymd, 2, 0)
        layout.addWidget(self.edit_ymd, 2, 1)
        layout.addWidget(self.btn_fetch, 3, 0)
        layout.addWidget(self.btn_save, 3, 1)
        layout.addWidget(self.table, 4, 0, 1, 2)

        self.setLayout(layout)
        self.resize(900, 600)

    def on_fetch(self):
        lawd = self.edit_lawd.text().strip()
        ymd = self.edit_ymd.text().strip()
        key = self.edit_key.text().strip()

        if not key or not lawd or not ymd:
            QMessageBox.warning(self, "입력 오류", "Service Key, 지역코드, 조회년월을 모두 입력하세요.")
            return

        data = self.get_apt_trade_data(lawd, ymd, key)
        if data is None:
            QMessageBox.information(self, "결과", "데이터를 가져오지 못했습니다.")
            return

        # 저장: 원본 데이터 보관 및 필터 초기화
        try:
            self.table_rows_master = data
            self.table_filters = {}
        except Exception:
            pass

        self.populate_table(data)
        self.btn_save.setEnabled(bool(data))

    def get_apt_trade_data(self, lawd_cd, deal_ymd, service_key):
        url = "https://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
        params = {
            "serviceKey": requests.utils.unquote(service_key),
            "LAWD_CD": lawd_cd,
            "DEAL_YMD": deal_ymd,
            "numOfRows": "1000",
            "pageNo": "1",
        }

        try:
            resp = requests.get(url, params=params, timeout=20)
            resp.raise_for_status()
        except Exception as e:
            QMessageBox.critical(self, "요청 실패", str(e))
            return None

        try:
            root = ET.fromstring(resp.content)
        except Exception as e:
            QMessageBox.critical(self, "파싱 오류", f"응답 XML 파싱 실패: {e}")
            return None

        header = root.find("header")
        code = None
        msg = None
        if header is not None:
            code = header.findtext("resultCode")
            msg = header.findtext("resultMsg")

        items = root.findall("body/items/item")
        if not items:
            if code is not None and code != "00":
                QMessageBox.warning(self, "API 오류", f"{code}: {msg}")
            else:
                QMessageBox.information(self, "결과 없음", "조회 결과가 없습니다.")
            return None

        def _find_text(item, candidates):
            for cand in candidates:
                try:
                    v = item.findtext(cand)
                except Exception:
                    v = None
                if v:
                    return v
            low_cands = [c.lower() for c in candidates]
            for child in list(item):
                t = (child.tag or "").lower()
                for lc in low_cands:
                    if lc in t:
                        return child.text or ""
            return ""

        rows = []
        for it in items:
            trade_date = f"{it.findtext('dealYear') or ''}-{it.findtext('dealMonth') or ''}-{it.findtext('dealDay') or ''}"
            row = [
                it.findtext("aptNm") or "",
                # 아파트동
                it.findtext("aptDong") or _find_text(it, ["aptDong", "단지동", "동", "apt_dong"]),
                it.findtext("excluUseAr") or "",
                trade_date,
                (it.findtext("dealAmount") or "").strip(),
                it.findtext("floor") or "",
                it.findtext("buildYear") or "",
                it.findtext("umdNm") or "",
                it.findtext("jibun") or "",
                it.findtext("sggCd") or "",
                # 추가 필드들: 우선 명시적 태그 사용, 없으면 후보 탐색
                (it.findtext("dealingGbn") or _find_text(it, ["tradeType", "거래유형", "dealType", "dealingGbn"])),
                (it.findtext("estateAgentSggNm") or _find_text(it, ["bcnstAddr", "brokerAddr", "중개사소재지", "bcnstc", "estateAgentSggNm"])),
                (it.findtext("rgstDate") or _find_text(it, ["registDay", "등기일자", "registrationDate", "rgstDate"])),
                (it.findtext("slerGbn") or _find_text(it, ["seller", "거래주체정보_매도자", "매도자", "tradePartSeller", "slerGbn"])),
                _find_text(it, ["buyer", "거래주체정보_매수자", "매수자", "tradePartBuyer"]),
                _find_text(it, ["rentYn", "토지임대부", "landLease", "isLandLeaseApt"]),
            ]
            rows.append(row)

        return rows

    def populate_table(self, rows):
        try:
            was_sorting = self.table.isSortingEnabled()
        except Exception:
            was_sorting = False
        try:
            self.table.setSortingEnabled(False)
        except Exception:
            pass
        try:
            self.table.clearContents()
        except Exception:
            pass
        self.table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                try:
                    txt = val if isinstance(val, str) else str(val)
                except Exception:
                    txt = ""
                # format 거래금액(만원) column with thousand separators (column index 4)
                if c == 4:
                    num_val = None
                    try:
                        import re
                        digits = re.sub(r"[^0-9]", "", txt)
                        if digits:
                            num_val = int(digits)
                            txt = f"{num_val:,}"
                    except Exception:
                        num_val = None
                    try:
                        if num_val is not None:
                            it = NumericItem(txt)
                            it.setData(Qt.UserRole, num_val)
                        else:
                            it = QTableWidgetItem(txt)
                    except Exception:
                        it = QTableWidgetItem(txt)
                    self.table.setItem(r, c, it)
                else:
                    self.table.setItem(r, c, QTableWidgetItem(txt))
        try:
            self.table.setSortingEnabled(was_sorting)
        except Exception:
            pass

    # -- Header filter handlers for AptTradeGUI.table --
    def aptgui_header_context_menu(self, pos):
        th = self.table.horizontalHeader()
        col = th.logicalIndexAt(pos)
        if col < 0:
            return
        if col in self.table_filters and self.table_filters[col]:
            try:
                del self.table_filters[col]
            except Exception:
                self.table_filters.pop(col, None)
            self.apply_aptgui_filters()
            return
        text, ok = QInputDialog.getText(self, "열 필터", f"열 {col} 필터 텍스트 입력:")
        if not ok:
            return
        text = text.strip()
        if not text:
            return
        self.table_filters[col] = text
        self.apply_aptgui_filters()

    def apply_aptgui_filters(self):
        if not getattr(self, 'table_rows_master', None):
            return
        if not self.table_filters:
            rows = self.table_rows_master
        else:
            def matches_all(row):
                for col, f in self.table_filters.items():
                    try:
                        if f.lower() not in (row[col] or "").lower():
                            return False
                    except Exception:
                        return False
                return True
            rows = [r for r in self.table_rows_master if matches_all(r)]
        self.populate_table(rows)

    def on_save_csv(self):
        # 간단히 현재 테이블 내용을 CSV로 저장
        lawd = self.edit_lawd.text().strip()
        ymd = self.edit_ymd.text().strip()
        filename = f"apt_trade_{lawd}_{ymd}.csv"
        try:
            with open(filename, "w", newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
                writer.writerow(headers)
                for r in range(self.table.rowCount()):
                    row = [self.table.item(r, c).text() if self.table.item(r, c) else "" for c in range(self.table.columnCount())]
                    writer.writerow(row)
            QMessageBox.information(self, "저장 성공", f"{filename} 으로 저장되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "저장 실패", str(e))
 
# Worker thread to fetch apartment trade data to avoid blocking UI
class AptFetchWorker(QThread):
    progress = pyqtSignal(int, int)  # current, total
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, lawd, months, service_key, parent=None):
        super().__init__(parent)
        # `lawd` may be a single LAWD string or a list of LAWD strings.
        if isinstance(lawd, (list, tuple)):
            self.lawd_list = list(lawd)
        else:
            self.lawd_list = [lawd]
        self.months = months
        self.service_key = service_key
        self._stop = False

    def stop(self):
        # signal to stop after current request finishes
        self._stop = True

    def run(self):
        import requests, xml.etree.ElementTree as ET, time, re
        rows = []
        # total progress is (#lawd * #months); guard against zero
        total = max(1, len(self.months) * max(1, len([l for l in self.lawd_list if l])))
        logs_dir = os.path.join(os.getcwd(), "debug_logs")
        try:
            os.makedirs(logs_dir, exist_ok=True)
        except Exception:
            pass

        def _find_text(item, candidates):
            for cand in candidates:
                try:
                    v = item.findtext(cand)
                except Exception:
                    v = None
                if v:
                    return v
            low_cands = [c.lower() for c in candidates]
            for child in list(item):
                t = (child.tag or "").lower()
                for lc in low_cands:
                    if lc in t:
                        return child.text or ""
            return ""

        def _norm_amount(s):
            if not s:
                return ""
            try:
                return re.sub(r"[^0-9]", "", s)
            except Exception:
                return s

        def _norm_rgst(s):
            if not s:
                return ""
            try:
                s2 = s.strip()
                m = re.match(r"^(\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})$", s2)
                if m:
                    yy = int(m.group(1)); yyyy = 2000 + yy if yy < 100 else yy
                    mm = int(m.group(2)); dd = int(m.group(3))
                    return f"{yyyy:04d}-{mm:02d}-{dd:02d}"
                m2 = re.match(r"^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})$", s2)
                if m2:
                    yyyy = int(m2.group(1)); mm = int(m2.group(2)); dd = int(m2.group(3))
                    return f"{yyyy:04d}-{mm:02d}-{dd:02d}"
                return s2
            except Exception:
                return s

        step = 0
        for lawd in self.lawd_list:
            if not lawd:
                continue
            for ym in self.months:
                if getattr(self, '_stop', False):
                    self.error.emit('취소됨')
                    return
                url = "https://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
                params = {
                    "serviceKey": requests.utils.unquote(self.service_key),
                    "LAWD_CD": lawd,
                    "DEAL_YMD": ym,
                    "pageNo": "1",
                    "numOfRows": "1000",
                }
                headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
                    "Accept": "application/xml, text/xml, */*;q=0.01",
                }
                try:
                    resp = requests.get(url, params=params, timeout=30, headers=headers)
                    resp.raise_for_status()
                except Exception as e:
                    self.error.emit(str(e))
                    return
                if getattr(self, '_stop', False):
                    self.error.emit('취소됨')
                    return
                # save raw response
                try:
                    ts = int(time.time())
                    fname = os.path.join(logs_dir, f"debug_response_{lawd}_{ym}_{ts}.xml")
                    with open(fname, "wb") as fw:
                        fw.write(resp.content)
                    meta = os.path.join(logs_dir, f"debug_response_{lawd}_{ym}_{ts}.meta.txt")
                    with open(meta, "w", encoding='utf-8') as fm:
                        fm.write(f"url: {resp.url}\nstatus: {resp.status_code}\nheaders: {dict(resp.headers)}\n")
                except Exception:
                    pass

                try:
                    root = ET.fromstring(resp.content)
                except Exception as e:
                    self.error.emit(f"XML parse error ({ym}): {e}")
                    return

                items = root.findall("body/items/item")
                for it in items:
                    trade_date = f"{it.findtext('dealYear') or ''}-{it.findtext('dealMonth') or ''}-{it.findtext('dealDay') or ''}"
                    raw_amount = it.findtext("dealAmount") or ""
                    amount_norm = _norm_amount(raw_amount)
                    rgst_raw = it.findtext("rgstDate") or _find_text(it, ["registDay", "등기일자", "registrationDate", "rgstDate"])
                    rgst_norm = _norm_rgst(rgst_raw)
                    row = [
                        it.findtext("aptNm") or "",
                        it.findtext("aptDong") or _find_text(it, ["aptDong", "단지동", "동", "apt_dong"]),
                        it.findtext("excluUseAr") or "",
                        trade_date,
                        amount_norm,
                        it.findtext("floor") or "",
                        it.findtext("buildYear") or "",
                        it.findtext("umdNm") or "",
                        it.findtext("jibun") or "",
                        it.findtext("sggCd") or "",
                        (it.findtext("dealingGbn") or _find_text(it, ["tradeType", "거래유형", "dealType", "dealingGbn"])),
                        (it.findtext("estateAgentSggNm") or _find_text(it, ["bcnstAddr", "brokerAddr", "중개사소재지", "bcnstc", "estateAgentSggNm"])),
                        rgst_norm,
                        (it.findtext("slerGbn") or _find_text(it, ["seller", "거래주체정보_매도자", "매도자", "tradePartSeller", "slerGbn"])),
                        _find_text(it, ["buyer", "거래주체정보_매수자", "매수자", "tradePartBuyer"]),
                        _find_text(it, ["rentYn", "토지임대부", "landLease", "isLandLeaseApt"]),
                    ]
                    rows.append(row)
                step += 1
                try:
                    self.progress.emit(min(step, total), total)
                except Exception:
                    pass
        self.finished.emit(rows)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = VWorldAdmCodeGUI()
    win.show()
    sys.exit(app.exec_())
