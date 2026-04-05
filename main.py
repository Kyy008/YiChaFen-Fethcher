import json
import re
import sys
import time
import uuid
from typing import Any, Dict, Optional

import requests
from pypinyin import lazy_pinyin
try:
    from PyQt6.QtCore import QObject, QThread, pyqtSignal
    from PyQt6.QtWidgets import (
        QApplication,
        QCheckBox,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QPlainTextEdit,
        QScrollArea,
        QVBoxLayout,
        QWidget,
    )
except ImportError as exc:
    raise SystemExit("请先安装 PyQt6：pip install PyQt6") from exc


BASE_URL = "https://www.yichafen.com"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/132.0.0.0 Safari/537.36 "
        "MicroMessenger/7.0.20.1781(0x6700143B) NetType/WIFI "
        "MiniProgramEnv/Mac MacWechat/WMPF "
        "MacWechat/3.8.7(0x13080712) "
        "UnifiedPCMacWechat(0xf2641702) XWEB/18788"
    ),
    "xweb_xhr": "1",
    "Accept": "*/*",
    "Content-Type": "application/x-www-form-urlencoded",
    "Referer": "https://servicewechat.com/wx9300e397f1f28ec7/65/page-frame.html",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "zh-CN,zh;q=0.9",
}

LABEL_KEYS = ("name", "label", "title", "columnName", "desc")
CHINESE_CHAR_PATTERN = re.compile(r"[\u4e00-\u9fff]")


def now_ms() -> str:
    return str(int(time.time() * 1000))


def make_nonce() -> str:
    return uuid.uuid4().hex


def to_submit_field_name(display_name: str) -> str:
    normalized_parts: list[str] = []

    for char in display_name.strip():
        if CHINESE_CHAR_PATTERN.fullmatch(char):
            normalized_parts.extend(lazy_pinyin(char))
            continue
        if char.isascii() and char.isalnum():
            normalized_parts.append(char.lower())

    return f"s_{''.join(normalized_parts)}"


def ensure_ok_json(resp: requests.Response, step: str) -> Dict[str, Any]:
    resp.raise_for_status()
    try:
        data = resp.json()
    except Exception as exc:
        raise RuntimeError(f"{step} 返回不是 JSON: {resp.text[:500]}") from exc

    if not isinstance(data, dict):
        raise RuntimeError(f"{step} 返回格式异常: {data!r}")

    if data.get("status") != 1:
        raise RuntimeError(f"{step} 失败: {json.dumps(data, ensure_ascii=False)}")

    return data


def get_qz_config(
    session: requests.Session,
    base_url: str,
    uq_code: str,
    openid_code: str,
    headers: Dict[str, str],
) -> Dict[str, Any]:
    params = {
        "uqCode": uq_code,
        "doorToken": "",
        "openidCode": openid_code,
        "timestamp": now_ms(),
        "nonce": make_nonce(),
    }

    resp = session.get(
        f"{base_url}/Miniprogram/Api/qz",
        params=params,
        headers=headers,
        timeout=20,
    )
    return ensure_ok_json(resp, "qz")


def verify_params(
    session: requests.Session,
    base_url: str,
    uq_code: str,
    openid_code: str,
    headers: Dict[str, str],
    kkey: str,
    form_fields: Dict[str, str],
) -> str:
    params = {
        "kkey": kkey,
        "uqCode": uq_code,
    }
    form = dict(form_fields)
    form.update(
        {
            "openidCode": openid_code,
            "timestamp": now_ms(),
            "nonce": make_nonce(),
        }
    )

    resp = session.post(
        f"{base_url}/Miniprogram/Api/verifyParams",
        params=params,
        data=form,
        headers=headers,
        timeout=20,
    )
    data = ensure_ok_json(resp, "verifyParams")

    token = data.get("data", {}).get("token")
    if not token:
        raise RuntimeError(f"verifyParams 没返回 token: {json.dumps(data, ensure_ascii=False)}")
    return str(token)


def get_result(
    session: requests.Session,
    base_url: str,
    uq_code: str,
    openid_code: str,
    headers: Dict[str, str],
    token: str,
) -> Dict[str, Any]:
    params = {
        "uqCode": uq_code,
        "token": token,
        "openidCode": openid_code,
        "timestamp": now_ms(),
        "nonce": make_nonce(),
    }

    resp = session.get(
        f"{base_url}/Miniprogram/Api/subjectResultV4",
        params=params,
        headers=headers,
        timeout=20,
    )
    return ensure_ok_json(resp, "subjectResultV4")


def choose_column_key(column: Dict[str, Any], fallback_index: int) -> str:
    for label_key in LABEL_KEYS:
        label_value = str(column.get(label_key) or "").strip()
        if label_value:
            return label_value

    pinyin = str(column.get("pinyin") or "").strip()
    if pinyin:
        return pinyin

    return f"field_{fallback_index}"


def normalize_value(value: Any) -> Optional[str]:
    if value is None:
        return None
    return str(value)


def extract_record(record: Dict[str, Any], include_empty: bool) -> Dict[str, Optional[str]]:
    column_list = record.get("columnList") or []
    extracted: Dict[str, Optional[str]] = {}

    for index, column in enumerate(column_list, start=1):
        if not isinstance(column, dict):
            continue
        key = choose_column_key(column, index)
        value = normalize_value(column.get("value"))
        if value is None and not include_empty:
            continue
        extracted[key] = value

    return extracted


def extract_result_records(
    result_json: Dict[str, Any],
    include_empty: bool,
) -> list[Dict[str, Optional[str]]]:
    data = result_json.get("data", {})
    record_list = data.get("recordList") or []
    if not isinstance(record_list, list):
        return []

    records: list[Dict[str, Optional[str]]] = []
    for record in record_list:
        if not isinstance(record, dict):
            continue
        records.append(extract_record(record, include_empty=include_empty))
    return records


def query_result(
    session: requests.Session,
    base_url: str,
    uq_code: str,
    openid_code: str,
    headers: Dict[str, str],
    form_fields: Dict[str, str],
) -> Dict[str, Any]:
    qz_json = get_qz_config(
        session=session,
        base_url=base_url,
        uq_code=uq_code,
        openid_code=openid_code,
        headers=headers,
    )
    kkey = qz_json.get("data", {}).get("kkey")
    if not kkey:
        raise RuntimeError(f"qz 没返回 kkey: {json.dumps(qz_json, ensure_ascii=False)[:1000]}")

    token = verify_params(
        session=session,
        base_url=base_url,
        uq_code=uq_code,
        openid_code=openid_code,
        headers=headers,
        kkey=str(kkey),
        form_fields=form_fields,
    )
    return get_result(
        session=session,
        base_url=base_url,
        uq_code=uq_code,
        openid_code=openid_code,
        headers=headers,
        token=token,
    )


def build_output(
    result_json: Dict[str, Any],
    include_empty: bool,
) -> Dict[str, Any]:
    records = extract_result_records(result_json, include_empty=include_empty)
    return {
        "record_count": len(records),
        "records": records,
    }


class QueryWorker(QObject):
    finished = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(
        self,
        uq_code: str,
        openid_code: str,
        form_fields: Dict[str, str],
        include_empty: bool,
        dump_raw: bool,
    ) -> None:
        super().__init__()
        self.uq_code = uq_code
        self.openid_code = openid_code
        self.form_fields = form_fields
        self.include_empty = include_empty
        self.dump_raw = dump_raw

    def run(self) -> None:
        try:
            with requests.Session() as session:
                result_json = query_result(
                    session=session,
                    base_url=BASE_URL,
                    uq_code=self.uq_code,
                    openid_code=self.openid_code,
                    headers=HEADERS,
                    form_fields=self.form_fields,
                )

            if self.dump_raw:
                payload = json.dumps(result_json, ensure_ascii=False, indent=2)
            else:
                payload = json.dumps(
                    build_output(result_json, include_empty=self.include_empty),
                    ensure_ascii=False,
                    indent=2,
                )
            self.finished.emit(payload)
        except Exception as exc:
            self.failed.emit(str(exc))


class FieldRow(QWidget):
    def __init__(self, remove_callback, display_name: str = "", value_text: str = "") -> None:
        super().__init__()
        self.remove_callback = remove_callback

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(6)

        top_layout = QHBoxLayout()
        top_layout.setContentsMargins(0, 0, 0, 0)

        self.display_name_edit = QLineEdit()
        self.display_name_edit.setPlaceholderText("小程序中看到的字段名")
        self.display_name_edit.setText(display_name)
        self.display_name_edit.textChanged.connect(self.update_submit_name)

        self.value_edit = QLineEdit()
        self.value_edit.setPlaceholderText("字段值")
        self.value_edit.setText(value_text)

        self.submit_name_edit = QLineEdit()
        self.submit_name_edit.setReadOnly(True)

        self.remove_button = QPushButton("删除")
        self.remove_button.clicked.connect(self.handle_remove)

        top_layout.addWidget(QLabel("字段"))
        top_layout.addWidget(self.display_name_edit, 2)
        top_layout.addWidget(QLabel("值"))
        top_layout.addWidget(self.value_edit, 3)
        top_layout.addWidget(self.remove_button)

        submit_layout = QHBoxLayout()
        submit_layout.setContentsMargins(0, 0, 0, 0)
        submit_layout.addWidget(QLabel("提交字段名"))
        submit_layout.addWidget(self.submit_name_edit)

        outer_layout.addLayout(top_layout)
        outer_layout.addLayout(submit_layout)

        self.update_submit_name()

    def handle_remove(self) -> None:
        self.remove_callback(self)

    def update_submit_name(self) -> None:
        self.submit_name_edit.setText(to_submit_field_name(self.display_name_edit.text()))


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.thread: Optional[QThread] = None
        self.worker: Optional[QueryWorker] = None
        self.field_rows: list[FieldRow] = []
        self.setWindowTitle("易查分查询工具")
        self.resize(960, 720)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        uq_layout = QHBoxLayout()
        self.uq_code_edit = QLineEdit()
        self.uq_code_edit.setPlaceholderText("必填")
        uq_layout.addWidget(QLabel("UQ Code"))
        uq_layout.addWidget(self.uq_code_edit)
        main_layout.addLayout(uq_layout)

        openid_layout = QHBoxLayout()
        self.openid_code_edit = QLineEdit()
        self.openid_code_edit.setPlaceholderText("必填")
        openid_layout.addWidget(QLabel("OpenID Code"))
        openid_layout.addWidget(self.openid_code_edit)
        main_layout.addLayout(openid_layout)

        field_header_layout = QHBoxLayout()
        field_header_layout.addWidget(QLabel("查询字段"))
        self.add_field_button = QPushButton("添加字段")
        self.add_field_button.clicked.connect(lambda: self.add_field_row())
        field_header_layout.addWidget(self.add_field_button)
        field_header_layout.addStretch()
        main_layout.addLayout(field_header_layout)

        self.field_container = QWidget()
        self.field_layout = QVBoxLayout(self.field_container)
        self.field_layout.setContentsMargins(0, 0, 0, 0)
        self.field_layout.setSpacing(8)
        self.field_layout.addStretch()

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(self.field_container)
        main_layout.addWidget(scroll_area, 3)

        option_layout = QHBoxLayout()
        self.include_empty_checkbox = QCheckBox("包含空值字段")
        self.dump_raw_checkbox = QCheckBox("显示原始返回")
        option_layout.addWidget(self.include_empty_checkbox)
        option_layout.addWidget(self.dump_raw_checkbox)
        option_layout.addStretch()
        main_layout.addLayout(option_layout)

        action_layout = QHBoxLayout()
        self.query_button = QPushButton("开始查询")
        self.query_button.clicked.connect(self.start_query)
        action_layout.addStretch()
        action_layout.addWidget(self.query_button)
        main_layout.addLayout(action_layout)

        self.result_edit = QPlainTextEdit()
        self.result_edit.setReadOnly(True)
        main_layout.addWidget(self.result_edit, 4)

        self.add_field_row("姓名")
        self.add_field_row("学号")

    def add_field_row(self, display_name: str = "", value_text: str = "") -> None:
        row = FieldRow(self.remove_field_row, display_name=display_name, value_text=value_text)
        self.field_rows.append(row)
        self.field_layout.insertWidget(self.field_layout.count() - 1, row)

    def remove_field_row(self, row: FieldRow) -> None:
        if len(self.field_rows) == 1:
            QMessageBox.warning(self, "无法删除", "至少保留一个查询字段。")
            return
        self.field_rows.remove(row)
        row.setParent(None)
        row.deleteLater()

    def collect_form_fields(self) -> Dict[str, str]:
        fields: Dict[str, str] = {}
        for row in self.field_rows:
            display_name = row.display_name_edit.text().strip()
            key = row.submit_name_edit.text().strip()
            value = row.value_edit.text()
            if not display_name and not value:
                continue
            if not display_name:
                raise ValueError("存在未填写展示字段名的行。")
            if key == "s_":
                raise ValueError(f"字段 {display_name} 无法转换成有效的提交字段名。")
            if not value:
                raise ValueError(f"字段 {display_name} 的值不能为空。")
            if key in fields:
                raise ValueError(f"字段 {display_name} 转换后与其他字段重复：{key}")
            fields[key] = value

        if not fields:
            raise ValueError("至少填写一个查询字段。")
        return fields

    def set_querying(self, querying: bool) -> None:
        self.query_button.setEnabled(not querying)
        self.add_field_button.setEnabled(not querying)
        self.uq_code_edit.setEnabled(not querying)
        self.openid_code_edit.setEnabled(not querying)
        self.include_empty_checkbox.setEnabled(not querying)
        self.dump_raw_checkbox.setEnabled(not querying)
        for row in self.field_rows:
            row.display_name_edit.setEnabled(not querying)
            row.value_edit.setEnabled(not querying)
            row.remove_button.setEnabled(not querying)
        if querying:
            self.result_edit.setPlainText("查询中，请稍候...")

    def start_query(self) -> None:
        uq_code = self.uq_code_edit.text().strip()
        openid_code = self.openid_code_edit.text().strip()

        if not uq_code:
            QMessageBox.warning(self, "缺少参数", "UQ Code 为必填项。")
            return
        if not openid_code:
            QMessageBox.warning(self, "缺少参数", "OpenID Code 为必填项。")
            return

        try:
            form_fields = self.collect_form_fields()
        except ValueError as exc:
            QMessageBox.warning(self, "字段错误", str(exc))
            return

        self.set_querying(True)

        self.thread = QThread(self)
        self.worker = QueryWorker(
            uq_code=uq_code,
            openid_code=openid_code,
            form_fields=form_fields,
            include_empty=self.include_empty_checkbox.isChecked(),
            dump_raw=self.dump_raw_checkbox.isChecked(),
        )
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.handle_query_success)
        self.worker.failed.connect(self.handle_query_failure)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(self.cleanup_worker)
        self.thread.start()

    def handle_query_success(self, result_text: str) -> None:
        self.result_edit.setPlainText(result_text)
        self.set_querying(False)

    def handle_query_failure(self, error_message: str) -> None:
        self.result_edit.setPlainText(error_message)
        self.set_querying(False)
        QMessageBox.critical(self, "查询失败", error_message)

    def cleanup_worker(self) -> None:
        if self.worker is not None:
            self.worker.deleteLater()
        self.worker = None
        self.thread = None


def main() -> None:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
