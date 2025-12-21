import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit,
    QComboBox, QPushButton, QGridLayout, QMessageBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QGroupBox,
    QSizePolicy, QProgressBar, QInputDialog
)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
import os

import requests
import xml.etree.ElementTree as ET
import csv
import datetime
import time


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

        self.setLayout(layout)
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
            # show all if '선택' chosen
            if hasattr(self, 'edit_selected_admcode'):
                self.edit_selected_admcode.setText("")
            # table removed; nothing to update
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

        if not key or not lawd or not from_ym or not to_ym:
            QMessageBox.warning(self, "입력 오류", "Service Key, 지역코드, 조회기간(시작/종료)을 모두 입력하세요.")
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
        params = {
            "serviceKey": requests.utils.unquote(key),
            "LAWD_CD": lawd,
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
            self.progress_bar.setMaximum(total_months)
            self.progress_bar.setValue(0)
            self.status_label.setText(f"진행: 0/{total_months}")
        except Exception:
            pass
        self.btn_apt_fetch.setEnabled(False)
        self.btn_apt_save.setEnabled(False)

        # Run data fetching in a worker thread to avoid blocking the UI
        try:
            self.progress_bar.setMaximum(total_months)
            self.progress_bar.setValue(0)
            self.status_label.setText(f"진행: 0/{total_months}")
        except Exception:
            pass

        self.btn_apt_fetch.setEnabled(False)
        self.btn_apt_save.setEnabled(False)

        # create worker
        self._apt_worker = AptFetchWorker(lawd, months, key)

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
                # trade_date is at column index 3 in worker row structure
                rows.sort(key=lambda r: _parse_date(r[3] if len(r) > 3 else ""))
                self.apt_rows_master = rows
                self.apt_filters = {}
                self.populate_apt_table(rows)
                self.btn_apt_save.setEnabled(bool(rows))
                self.status_label.setText(f"완료: {len(rows)}건")
                self.progress_bar.setValue(total_months)
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
        self.lawd = lawd
        self.months = months
        self.service_key = service_key
        self._stop = False

    def stop(self):
        # signal to stop after current month finishes
        self._stop = True

    def run(self):
        import requests, xml.etree.ElementTree as ET, time, re
        rows = []
        total = len(self.months)
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

        for i, ym in enumerate(self.months, start=1):
            if getattr(self, '_stop', False):
                # emit error to indicate cancellation
                self.error.emit('취소됨')
                return
            url = "https://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
            params = {
                "serviceKey": requests.utils.unquote(self.service_key),
                "LAWD_CD": self.lawd,
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
                fname = os.path.join(logs_dir, f"debug_response_{self.lawd}_{ym}_{ts}.xml")
                with open(fname, "wb") as fw:
                    fw.write(resp.content)
                meta = os.path.join(logs_dir, f"debug_response_{self.lawd}_{ym}_{ts}.meta.txt")
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
            self.progress.emit(i, total)
        self.finished.emit(rows)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = VWorldAdmCodeGUI()
    win.show()
    sys.exit(app.exec_())
