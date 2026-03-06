
"""
USB-util: USB/BTデバイス情報のスキャン・表示・保存・COMポート逆引きを行うPython GUI/CLIツール

- USBデバイスの詳細情報取得（PyUSB）
- usb.idsによるベンダー名・製品名補完
- JSON保存・読み込み
- CustomTkinter GUIで情報表示
- COMポート情報の取得・逆引き（pyserial/win32com）
- コマンドラインからCOMポート取得も可能
- 指定したVID:PIDのデバイスへシリアルコマンド送受信（pyserial）
"""

# 主な機能
# - PyUSBでUSBデバイスの詳細情報取得
# - usb.idsによるベンダー名・製品名補完
# - JSON保存・読み込み
# - CustomTkinter GUIで情報表示
# - 権限不足/バックエンド未導入時のエラー表示

import argparse
import os
import sys
from pathlib import Path

from typing import Any, Dict, List, Optional, Tuple

import customtkinter as ctk

import threading
import time

from core.com_ports import ComPortManager
from core.bootstrap import setup_services as bootstrap_setup_services
from core.device_models import UsbDeviceSnapshot, UsbSnapshotRepository, UsbSnapshotService
from core.scanners import DeviceScanner
from core.usb_ids import UsbIdsDatabase
from ui.view_model import UsbDevicesViewModel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USB_JSON_PATH = os.path.join(BASE_DIR, "usb_devices.json")
_SERVICE_SINGLETON: Optional[UsbSnapshotService] = None

AILAB_ROOT = Path(__file__).resolve().parents[2]
AK_GUIPARTS_ROOT = os.path.abspath(
    str(AILAB_ROOT / "lab_automation_libs" / "gui_parts" / "aist-guiparts")
)
if AK_GUIPARTS_ROOT not in sys.path and os.path.isdir(AK_GUIPARTS_ROOT):
    sys.path.insert(0, AK_GUIPARTS_ROOT)

from aist_guiparts.ui_base import BaseApp  # noqa: E402


class UsbDevicesApp:
    """Render USB device snapshots via CustomTkinter."""

    def __init__(self, view_model: UsbDevicesViewModel) -> None:
        self.view_model = view_model
        print(f"[DEBUG] UsbDevicesApp initial snapshots count: {self.view_model.device_count()}")

        self.app = BaseApp(theme="dark")
        self.app.title("AUTO Kobo USB Utility")
        self.app.geometry("900x600")
        self.app.build_menu()

        self.combo: Optional[ctk.CTkComboBox] = None
        self.device_count_label: Optional[ctk.CTkLabel] = None
        self.device_listbox: Optional[ctk.CTkScrollableFrame] = None
        self.device_list_items: List[ctk.CTkFrame] = []
        self.info_labels: Dict[str, ctk.CTkLabel] = {}
        self.error_label: Optional[ctk.CTkLabel] = None
        self.detail_box: Optional[ctk.CTkTextbox] = None

        self._ui_spacing = self._spacing_config()
        self._ui_fonts = self._font_config()

        self._setup_layout()
        self._show_scanning_indicator()
        threading.Thread(target=self._background_initial_scan, daemon=True).start()

    def _background_initial_scan(self):
        self.view_model.refresh()
        self.app.after(0, self._finish_scanning_indicator)

    def _show_scanning_indicator(self):
        if self.device_listbox:
            for item in self.device_list_items:
                item.destroy()
            self.device_list_items.clear()
            self._scanning_blink_state = True
            self._scanning_label = ctk.CTkLabel(
                self.device_listbox,
                text="🔄 デバイス検索中... 🔄",
                font=("Meiryo", 16, "bold"),
                text_color="#FF8800"
            )
            self._scanning_label.pack(pady=30)
            self.device_list_items.append(self._scanning_label)
            self._blink_scanning_label()

    def _blink_scanning_label(self):
        # ブリンク（点滅）処理
        if not hasattr(self, "_scanning_label") or self._scanning_label is None:
            return
        self._scanning_blink_state = not getattr(self, "_scanning_blink_state", False)
        if self._scanning_blink_state:
            self._scanning_label.configure(text="🔄 デバイス検索中... 🔄")
        else:
            self._scanning_label.configure(text="   デバイス検索中...   ")
        self.app.after(600, self._blink_scanning_label)

    def _finish_scanning_indicator(self):
        # スキャン完了時にインジケータを消してリストを更新
        if hasattr(self, "_scanning_label") and self._scanning_label:
            self._scanning_label.destroy()
            self._scanning_label = None
        self._apply_view_model(update_combo=True, rebuild_list=True)

    def run(self) -> None:
        self.app.mainloop()

    def _spacing_config(self) -> dict:
        return {
            "layout": {
                "root": (20, 20),             # ウィンドウ全体の外側マージン
                "content_gap": (10, 5),      # 上部コントロールバーと下部エリアの間隔
            },
            "header": {
                "pady": (0, 10),
            },
            "top_controls": {
                "frame_padx": 10,             # 上部コントロールバーの左右余白
                "frame_pady": (10, 5),        # 上部コントロールバーの上下余白
                "title_padx": 10,             # タイトルラベルと左端の距離
                "combo_padx": 20,             # コンボボックス左右余白
                "button_padx": 10,            # 再読み込みボタン左右余白
            },
            "device_summary": {
                "header_pad": (10, 0),        # 左ペイン見出しの余白
                "frame_padx": 10,             # サマリーリスト左右余白
                "frame_pady": 10,             # サマリーリスト上下余白
                "item_padx": 6,               # サマリー項目枠の左右余白
                "item_pady": 4,               # サマリー項目枠の上下余白
                "label_padx": 8,              # サマリー項目テキスト左右余白
                "label_pady": 6,              # サマリー項目テキスト上下余白
            },
            "device_info": {
                "frame_padx": 10,             # デバイス情報/JSONペインの左右余白
                "frame_pady": 10,             # デバイス情報/JSONペインの上下余白
                "label_padx": (10, 6),        # 詳細ラベルと左端との距離
                "value_padx": (0, 10),        # 詳細値と右端との距離
                "row_pady": (0, 0),           # 詳細項目の行間
                "error_pady": (8, 0),         # エラー表示の上下余白
                "json_padx": 10,              # JSONテキストボックス左右余白
                "json_pady": 10,              # JSONテキストボックス上下余白
            },
        }

    def _font_config(self) -> dict:
        return {
            "top_controls": {"title": ("Meiryo", 20, "bold")},
            "section_heading": ("Meiryo", 14, "bold"),
            "device_summary": {"counter": ("Meiryo", 12)},
            "device_info": {"label": ("Meiryo", 12, "bold"), "value": ("Meiryo", 12), "error": ("Meiryo", 11)},
        }

    # ------------------------------------------------------------------ layout
    def _setup_layout(self) -> None:
        spacing = self._ui_spacing

        layout_spacing = spacing["layout"]
        summary_spacing = spacing["device_summary"]
        info_spacing = spacing["device_info"]

        header = self.app.build_default_titlebar("AK USB Utility", logo_height=36)
        header.pack(fill="x", padx=layout_spacing["root"][0], pady=spacing["header"]["pady"])

        root_padx, root_pady = layout_spacing["root"]
        main_frame = ctk.CTkFrame(self.app)
        main_frame.pack(fill="both", expand=True, padx=root_padx, pady=(0, root_pady))

        self._setup_top_controls(main_frame)

        bottom_frame = ctk.CTkFrame(main_frame)
        bottom_frame.pack(fill="both", expand=True, padx=info_spacing["frame_padx"], pady=layout_spacing["content_gap"])
        bottom_frame.grid_columnconfigure(0, weight=0)
        bottom_frame.grid_columnconfigure(1, weight=1)
        bottom_frame.grid_columnconfigure(2, weight=1)
        bottom_frame.grid_rowconfigure(0, weight=1)

        summary_frame = ctk.CTkFrame(bottom_frame)
        summary_frame.grid(row=0, column=0, sticky="nsew", padx=(0, summary_spacing["frame_padx"]))
        self._setup_summary_section(summary_frame)

        detail_parent = ctk.CTkFrame(bottom_frame)
        detail_parent.grid(row=0, column=1, sticky="nsew", padx=(0, info_spacing["frame_padx"]))
        self._setup_detail_section(detail_parent)

        json_parent = ctk.CTkFrame(bottom_frame)
        json_parent.grid(row=0, column=2, sticky="nsew", padx=(info_spacing["frame_padx"], 0))
        self._setup_json_section(json_parent)

    def _setup_top_controls(self, parent: ctk.CTkFrame) -> None:
        controls_spacing = self._ui_spacing["top_controls"]
        controls_fonts = self._ui_fonts["top_controls"]

        frame = ctk.CTkFrame(parent)
        frame.pack(fill="x", padx=controls_spacing["frame_padx"], pady=controls_spacing["frame_pady"])

        title = ctk.CTkLabel(frame, text="USB/BTデバイス選択", font=controls_fonts["title"])
        title.pack(side="left", padx=controls_spacing["title_padx"])

        options = self.view_model.get_options()
        self.combo = ctk.CTkComboBox(frame, values=options, width=300)
        self.combo.pack(side="left", padx=controls_spacing["combo_padx"])
        self.combo.configure(command=self._on_selection_change)

        reload_btn = ctk.CTkButton(frame, text="再読み込み", command=self._reload_snapshots, width=120)
        reload_btn.pack(side="right", padx=controls_spacing["button_padx"])

    def _setup_summary_section(self, parent: ctk.CTkFrame) -> None:
        summary_spacing = self._ui_spacing["device_summary"]
        section_heading_font = self._ui_fonts["section_heading"]
        summary_fonts = self._ui_fonts["device_summary"]

        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", padx=summary_spacing["frame_padx"], pady=summary_spacing["header_pad"])

        device_list_label = ctk.CTkLabel(header, text="USB/BTデバイス一覧", font=section_heading_font)
        device_list_label.pack(side="left")

        self.device_count_label = ctk.CTkLabel(header, text="(0)", font=summary_fonts["counter"])
        self.device_count_label.pack(side="right")

        self.device_listbox = ctk.CTkScrollableFrame(parent, width=260, fg_color=("gray12", "gray18"))
        self.device_listbox.pack(
            fill="both",
            expand=True,
            padx=summary_spacing["frame_padx"],
            pady=summary_spacing["frame_pady"],
        )

    def _setup_detail_section(self, parent: ctk.CTkFrame) -> None:
        info_spacing = self._ui_spacing["device_info"]
        section_heading_font = self._ui_fonts["section_heading"]
        info_fonts = self._ui_fonts["device_info"]

        heading = ctk.CTkLabel(parent, text="デバイス情報", font=section_heading_font)
        heading.pack(anchor="w", padx=info_spacing["frame_padx"], pady=(10, 0))

        info_frame = ctk.CTkScrollableFrame(parent)
        info_frame.pack(fill="both", expand=True, padx=info_spacing["frame_padx"], pady=info_spacing["frame_pady"])
        info_frame.grid_columnconfigure(0, weight=1)

        fields = [
            "VID",
            "usb.ids Vendor",
            "PID",
            "usb.ids Product",
            "Manufacturer",
            "Product",
            "Serial",
            "Identity",
            "Bus",
            "Address",
            "Port Path",
            "Class Guess",
            "COMポート",
            "接続経路",
            "LocationInformation",
            "BLE Address",
            "BLE Name",
            "BLE RSSI",
            "BLE UUIDs",
        ]

        for row, label in enumerate(fields):
            label_kwargs = {"text": f"{label}: ―", "font": info_fonts["label"], "justify": "left", "anchor": "w"}
            if label == "COMポート":
                label_kwargs["text_color"] = "red"
            row_label = ctk.CTkLabel(info_frame, **label_kwargs)
            row_label.grid(
                row=row,
                column=0,
                sticky="w",
                padx=info_spacing["label_padx"],
                pady=info_spacing["row_pady"],
            )
            self.info_labels[label] = row_label

        self.error_label = ctk.CTkLabel(info_frame, text="", font=info_fonts["error"], text_color="red")
        self.error_label.grid(row=len(fields), column=0, columnspan=2, sticky="w", padx=info_spacing["frame_padx"], pady=info_spacing["error_pady"])

    def _setup_json_section(self, parent: ctk.CTkFrame) -> None:
        info_spacing = self._ui_spacing["device_info"]
        section_heading_font = self._ui_fonts["section_heading"]
        info_fonts = self._ui_fonts["device_info"]

        heading = ctk.CTkLabel(parent, text="デバイス詳細 (JSON)", font=section_heading_font)
        heading.pack(anchor="w", padx=info_spacing["frame_padx"], pady=(10, 0))

        self.detail_box = ctk.CTkTextbox(parent, font=info_fonts["value"])
        self.detail_box.pack(fill="both", expand=True, padx=info_spacing["json_padx"], pady=info_spacing["json_pady"])

    # -------------------------------------------------------------- view syncs
    def _apply_view_model(self, update_combo: bool = True, rebuild_list: bool = True) -> None:
        if update_combo and self.combo:
            options = self.view_model.get_options()
            self.combo.configure(values=options)
            selected = self.view_model.selected_option()
            if selected:
                self.combo.set(selected)

        if rebuild_list:
            self._rebuild_device_list()

        self._update_detail_section()

    def _rebuild_device_list(self) -> None:
        if not self.device_listbox:
            return
        for item in self.device_list_items:
            item.destroy()
        self.device_list_items.clear()

        snapshots = self.view_model.snapshots
        if self.device_count_label:
            self.device_count_label.configure(text=f"({len(snapshots)})")

        for index, snapshot in enumerate(snapshots):
            self._add_device_list_item(index, snapshot)

    def _add_device_list_item(self, index: int, snapshot: UsbDeviceSnapshot) -> None:
        if not self.device_listbox:
            return

        item_frame = ctk.CTkFrame(self.device_listbox, corner_radius=8)
        entries = self.view_model.list_entries()
        if index < len(entries):
            text, dimmed = entries[index]
        else:
            text, dimmed = (snapshot.key(), False)
        if dimmed:
            item_frame.configure(fg_color=("#9fd3ff", "#75b5f2"))
        else:
            item_frame.configure(fg_color=("#50a7ff", "#2b82d9"))
        item_frame.pack(fill="x", padx=self._ui_spacing["device_summary"]["item_padx"], pady=self._ui_spacing["device_summary"]["item_pady"])

        text_color = ("#1f3c66", "#0f2845") if dimmed else ("black", "black")
        item_label = ctk.CTkLabel(item_frame, text=text, justify="left", anchor="w", text_color=text_color)
        item_label.pack(fill="x", padx=self._ui_spacing["device_summary"]["label_padx"], pady=self._ui_spacing["device_summary"]["label_pady"])

        item_label.bind("<Button-1>", lambda _event, idx=index: self._on_list_item_clicked(idx))
        item_frame.bind("<Button-1>", lambda _event, idx=index: self._on_list_item_clicked(idx))

        self.device_list_items.append(item_frame)

    def _update_detail_section(self) -> None:
        selected = self.view_model.current_snapshot()
        if selected is None:
            for label in self.info_labels.values():
                label.configure(text=label.cget("text").split(":")[0] + ": ―")
            if self.detail_box:
                self.detail_box.delete("1.0", "end")
            if self.error_label:
                self.error_label.configure(text="")
            return

        details = self.view_model.info_values()
        for key, value in details.items():
            if key in self.info_labels:
                self.info_labels[key].configure(text=f"{key}: {value}")

        if self.detail_box:
            self.detail_box.delete("1.0", "end")
            self.detail_box.insert("1.0", self.view_model.detail_json())

        if self.error_label:
            self.error_label.configure(text=self.view_model.error_message())

    def _on_list_item_clicked(self, index: int) -> None:
        self.view_model.select_by_index(index)
        self._apply_view_model(update_combo=True, rebuild_list=True)

    def _on_selection_change(self, selected: str) -> None:
        self.view_model.select_by_key(selected)
        self._apply_view_model(update_combo=False, rebuild_list=True)

    def _reload_snapshots(self) -> None:
        self._show_scanning_indicator()
        threading.Thread(target=self._background_reload_scan, daemon=True).start()

    def _background_reload_scan(self):
        self.view_model.refresh()
        self.app.after(0, self._finish_scanning_indicator)


def _get_service_singleton() -> UsbSnapshotService:
    """Return a lazily-instantiated UsbSnapshotService for convenience helpers."""

    global _SERVICE_SINGLETON
    if _SERVICE_SINGLETON is None:
        scanner = DeviceScanner(ble_timeout=5.0)
        repository = UsbSnapshotRepository(USB_JSON_PATH)
        _SERVICE_SINGLETON = UsbSnapshotService(scanner, repository)
    return _SERVICE_SINGLETO


def get_com_port_for_device(
    vid: str,
    pid: str,
    serial: Optional[str] = None,
    *,
    refresh: bool = False,
) -> Optional[str]:
    """
    VID/PID(/Serial)から一致するデバイスのCOMポート名を返す便利関数。

    外部モジュールから `from usb_util_gui import get_com_port_for_device` として呼び出す想定。

    Args:
        vid: ベンダーID（例: "0x1234"）
        pid: プロダクトID（例: "0x5678"）
        serial: シリアル番号（任意）
        refresh: Trueにすると照合前にUSBデバイスを再スキャン

    Returns:
        一致するCOMポート名。検出できなければNone。
    """

    service = _get_service_singleton()
    port = service.get_com_port_for_device(vid, pid, serial, refresh=refresh)
    if port or refresh:
        return port

    # 一度もスナップショットが保存されていないケースへのフォールバックとしてスキャン
    service.refresh()
    return service.get_com_port_for_device(vid, pid, serial, refresh=False)


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    """CLI用の引数を定義し、与えられたargvからNamespaceを生成する。"""
    parser = argparse.ArgumentParser(description="USB device viewer and identifier")
    parser.add_argument("vid", nargs="?", help="Vendor ID (e.g. 0x1234)")
    parser.add_argument("pid", nargs="?", help="Product ID (e.g. 0x5678)")
    parser.add_argument("serial", nargs="?", help="Serial number (optional)")
    parser.add_argument(
        "--refresh",
        action="store_true",
        help="Rescan devices before returning results",
    )
    parser.add_argument(
        "--send",
        help="MatchingデバイスのCOMポートへ送信するコマンド文字列（pyserial required）",
    )
    parser.add_argument(
        "--baudrate",
        type=int,
        default=9600,
        help="送受信に利用するボーレート（default: 9600）",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=2.0,
        help="送受信のタイムアウト秒数（default: 2.0）",
    )
    parser.add_argument(
        "--read-bytes",
        type=int,
        default=None,
        metavar="N",
        help="受信時に読み取るバイト数。未指定の場合は読み取りを行わない",
    )
    parser.add_argument(
        "--read-until",
        help="この文字列（デコード後）が現れるまで読み取る。--read-bytesより優先",
    )
    parser.add_argument(
        "--encoding",
        default="utf-8",
        help="コマンド送信およびレスポンスデコード時に用いる文字コード（default: utf-8）",
    )
    parser.add_argument(
        "--append-newline",
        action="store_true",
        help="送信前にコマンド末尾へ改行(\\n)を付与する",
    )
    parser.add_argument(
        "--self-test",
        action="store_true",
        help="環境診断を行い、検出したUSBスナップショットとCOMポート情報を表示する",
    )
    return parser.parse_args(argv)


def run_cli(service: UsbSnapshotService, ids_db: UsbIdsDatabase, args: argparse.Namespace) -> int:
    """CLI要求を処理し、指定VID/PIDのUSBデバイス情報を標準出力へ表示する。"""
    # 例) `python usb_util_gui.py 0x25a4 0x9311` の処理本体。
    #     `--send` オプションを付与するとシリアルコマンド送信も行う。
    if not args.vid or not args.pid:
        print("VID と PID を指定してください", file=sys.stderr)
        return 2

    results = service.find_device_connections(args.vid, args.pid, args.serial, refresh=args.refresh)
    if not results:
        print("該当するUSBデバイスが見つかりません", file=sys.stderr)
        return 1

    for entry in results:
        snapshot: UsbDeviceSnapshot = entry["snapshot"]
        identity = entry["identity"]
        port_path = "-".join(map(str, entry["port_path"])) if entry["port_path"] else "-"
        vendor_label, product_label = snapshot.resolve_names(ids_db)
        vendor_raw = snapshot.manufacturer.strip() if snapshot.manufacturer else "―"
        product_raw = snapshot.product.strip() if snapshot.product else "―"
        vendor_display = vendor_label or "不明"
        product_display = product_label or "不明"
        print(f"Identity: {identity}")
        print(f"  VID:PID : {snapshot.vid}:{snapshot.pid}")
        print(f"  Serial  : {entry['serial'] or '―'}")
        print(f"  Vendor  : {vendor_display} (raw: {vendor_raw})")
        print(f"  Product : {product_display} (raw: {product_raw})")
        print(f"  PortPath: {port_path}")
        print(f"  Bus/Addr: {snapshot.bus}/{snapshot.address}")
        print(f"  COM Port: {entry['com_port'] or '情報なし'}")
        print("")

    if args.send:
        read_bytes = args.read_bytes if args.read_bytes and args.read_bytes > 0 else None
        try:
            command_result = service.send_serial_command(
                args.vid,
                args.pid,
                args.send,
                serial=args.serial,
                refresh=False,
                baudrate=args.baudrate,
                timeout=args.timeout,
                read_bytes=read_bytes,
                read_until=args.read_until,
                encoding=args.encoding,
                append_newline=args.append_newline,
            )
        except RuntimeError as exc:
            print(str(exc), file=sys.stderr)
            return 3

        print("=== コマンド送受信結果 ===")
        print(f"Port     : {command_result['port']}")
        print(f"Written  : {command_result['bytes_written']} bytes")
        response_raw = command_result.get("response_raw", b"")
        response_text = command_result.get("response_text", "")
        if response_raw:
            print(f"Response(raw): {response_raw}")
        else:
            print("Response(raw): (empty)")
        if response_text:
            print(f"Response(txt): {response_text}")
        elif response_raw:
            print("Response(txt): (decode失敗)")
        else:
            print("Response(txt): (no data)")

    return 0


def run_self_test(
    service: UsbSnapshotService,
    snapshots: List[UsbDeviceSnapshot],
    scan_error: Optional[str],
) -> int:
    """
    USBスキャンとCOMポート列挙の結果をダンプし、環境診断に役立てる。
    """
    # 例) `python usb_util_gui.py --self-test` で実行される診断ルーチン。

    print("=== USB-util Self Test ===")
    print(f"USB JSON path: {USB_JSON_PATH}")
    print(f"Snapshot count: {len(snapshots)}")
    if scan_error:
        print(f"Scan error: {scan_error}", file=sys.stderr)

    if snapshots:
        print("\n--- Snapshots (up to 5 entries) ---")
        shown = 0
        for snapshot in snapshots:
            if snapshot.error:
                print(f"* Error snapshot: {snapshot.error}")
                continue
            print(f"* {snapshot.identity()}")
            print(f"    VID:PID : {snapshot.vid}:{snapshot.pid}")
            print(f"    Serial  : {snapshot.serial or '―'}")
            print(f"    Class   : {snapshot.class_guess}")
            shown += 1
            if shown >= 5:
                break
        if len(snapshots) > shown:
            print(f"... ({len(snapshots) - shown} more snapshots)")
    else:
        print("No snapshots available. USBデバイスが接続されているか確認してください。")

    ports = ComPortManager.get_com_ports()
    print("\n--- COM Ports ---")
    if not ports:
        print("検出されたCOMポートはありません。pyserialがインストール済みか確認してください。")
    else:
        for port in ports:
            device = port.get("device") or "-"
            desc = port.get("description") or port.get("hwid") or "-"
            vid = port.get("vid") or "-"
            pid = port.get("pid") or "-"
            serial = port.get("serial_number") or "-"
            print(f"* {device}: {desc}")
            print(f"    VID:PID={vid}:{pid} Serial={serial}")

    if not scan_error and snapshots:
        print("\nSelf test completed successfully.")
        return 0
    print("\nSelf test completed with warnings.")
    return 1 if scan_error else 0


def run_gui(
    ids_db: UsbIdsDatabase,
    service: UsbSnapshotService,
    snapshots: List[UsbDeviceSnapshot],
) -> None:
    """GUIアプリケーションを組み立て、初期データを読み込んで起動する。"""
    view_model = UsbDevicesViewModel(service, ids_db)
    view_model.load_initial(snapshots)
    app = UsbDevicesApp(view_model)
    app.run()


def main(argv: Optional[List[str]] = None) -> None:
    """CLI経路とGUI経路を切り替えつつ、USB情報ツールのエントリーポイントを提供する。"""
    args = parse_args(argv)
    service, snapshots, scan_error = bootstrap_setup_services(USB_JSON_PATH)
    ids_db = UsbIdsDatabase()

    if args.self_test:
        exit_code = run_self_test(service, snapshots, scan_error)
        sys.exit(exit_code)

    if args.vid and args.pid:
        exit_code = run_cli(service, ids_db, args)
        sys.exit(exit_code)

    run_gui(ids_db, service, snapshots)


if __name__ == "__main__":
    main()
