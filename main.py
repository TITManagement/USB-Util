"""
USB-util: USBデバイス情報のスキャン・表示・保存・COMポート逆引きを行うPython GUI/CLIツール

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
#
# クラス構成
# - UsbScanner: USBデバイスのスキャン
# - UsbIdsDatabase: usb.idsパースと名称解決
# - UsbDevicesApp: GUI表示
import argparse
import os
import sys
from typing import Any, Dict, List, Optional, Tuple

from core.com_ports import ComPortManager
from core.models import UsbDeviceSnapshot
from core.repository import UsbSnapshotRepository
from core.service import UsbSnapshotService
from core.scanner import UsbScanner
from ui.app import UsbDevicesApp
from ui.view_model import UsbDevicesViewModel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USB_JSON_PATH = os.path.join(BASE_DIR, "usb_devices.json")
_SERVICE_SINGLETON: Optional[UsbSnapshotService] = None


def _get_service_singleton() -> UsbSnapshotService:
    """Return a lazily-instantiated UsbSnapshotService for convenience helpers."""

    global _SERVICE_SINGLETON
    if _SERVICE_SINGLETON is None:
        scanner = UsbScanner()
        repository = UsbSnapshotRepository(USB_JSON_PATH)
        _SERVICE_SINGLETON = UsbSnapshotService(scanner, repository)
    return _SERVICE_SINGLETON


def get_com_port_for_device(
    vid: str,
    pid: str,
    serial: Optional[str] = None,
    *,
    refresh: bool = False,
) -> Optional[str]:
    """
    VID/PID(/Serial)から一致するデバイスのCOMポート名を返す便利関数。

    外部モジュールから `from main import get_com_port_for_device` として呼び出す想定。

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


def find_usb_ids_path() -> str:
    """Return the first existing usb.ids path across supported platforms."""
    candidates: List[Optional[str]] = []
    env_path = os.environ.get("USB_IDS_PATH")
    if env_path:
        candidates.append(env_path)
    candidates.append(os.path.join(BASE_DIR, "usb.ids"))
    candidates.append(os.path.join(os.getcwd(), "usb.ids"))
    candidates.extend([
        "/usr/share/hwdata/usb.ids",
        "/usr/share/misc/usb.ids",
        "/var/lib/usbutils/usb.ids",
        "/opt/homebrew/share/hwdata/usb.ids",
        "/opt/local/share/hwdata/usb.ids",
    ])
    if sys.platform.startswith("win"):
        program_data = os.environ.get("ProgramData")
        if program_data:
            candidates.append(os.path.join(program_data, "usb.ids"))
    for path in candidates:
        if path and os.path.exists(path):
            return os.path.abspath(path)
    return os.path.join(BASE_DIR, "usb.ids")


def _normalize_usb_id(value: Any) -> Optional[str]:
    """値をUSB ID表記(4桁の16進文字列)に正規化する。"""
    if value is None:
        return None
    if isinstance(value, int):
        return format(value, "04x")
    text = str(value).strip().lower()
    if text.startswith("0x"):
        text = text[2:]
    return text.zfill(4)


class UsbIdsDatabase:
    """
    usb.idsファイルをパースし、ベンダーID・プロダクトIDから名称を解決するクラス。
    - reload(): キャッシュ再構築
    - lookup(vid, pid): ベンダー名・製品名取得
    """

    # usb.idsのパスを受け取り、ベンダー/プロダクト名の辞書を構築
    def __init__(self, ids_path: Optional[str] = None) -> None:
        """usb.idsのパスを保持し、後続処理のためのキャッシュ領域を初期化する。"""
        # 初期化。ids_pathはusb.idsファイルのパス
        self.ids_path = ids_path or find_usb_ids_path()
        self._cache: Optional[Dict[str, Dict[str, Any]]] = None

    def reload(self) -> None:
        """内部キャッシュを無効化し、次回lookup時に再パースさせる。"""
        # キャッシュをクリアして再パース可能にする
        self._cache = None

    def lookup(self, vid: Any, pid: Any) -> Tuple[Optional[str], Optional[str]]:
        """
        ベンダーID・プロダクトIDを受け取り、ベンダー名・製品名を返却する。
        一致する情報がなければタプルの要素はいずれもNoneとなる。
        """
        # ベンダーID・プロダクトIDから名称を取得
        vendors = self._ensure_cache()
        norm_vid = _normalize_usb_id(vid)
        norm_pid = _normalize_usb_id(pid)
        if norm_vid is None or norm_pid is None:
            return None, None
        vendor_entry = vendors.get(norm_vid)
        if not vendor_entry:
            return None, None
        vendor_name = vendor_entry.get("name")
        product_entry = vendor_entry["products"].get(norm_pid)
        product_name = product_entry.get("name") if product_entry else None
        return vendor_name, product_name

    def _ensure_cache(self) -> Dict[str, Dict[str, Any]]:
        """キャッシュが未構築であればusb.idsを読み込み辞書へ展開する。"""
        # キャッシュがなければパースして構築
        if self._cache is None:
            self._cache = self._parse_usb_ids(self.ids_path)
        return self._cache

    @staticmethod
    def _parse_usb_ids(ids_path: str) -> Dict[str, Dict[str, Any]]:
        """usb.idsファイルを逐次読み込み、ベンダー・プロダクト階層の辞書を生成する。"""
        # usb.idsファイルをパースして辞書化
        vendors: Dict[str, Dict[str, Any]] = {}
        current_vendor: Optional[str] = None
        current_product: Optional[str] = None
        try:
            with open(ids_path, "r", encoding="utf-8") as infile:
                for raw_line in infile:
                    line = raw_line.rstrip("\n")
                    if not line or line.startswith("#"):
                        continue
                    if line.startswith("\t\t"):
                        # インターフェース情報
                        if current_vendor is None or current_product is None:
                            continue
                        parts = line.strip().split(None, 1)
                        if not parts:
                            continue
                        interface_code = parts[0].lower()
                        interface_name = parts[1].strip() if len(parts) > 1 else ""
                        product_entry = vendors[current_vendor]["products"].setdefault(
                            current_product, {"name": "", "interfaces": []}
                        )
                        product_entry["interfaces"].append(
                            {"code": interface_code, "name": interface_name}
                        )
                        continue
                    if line.startswith("\t"):
                        # プロダクト情報
                        if current_vendor is None:
                            continue
                        parts = line.strip().split(None, 1)
                        if not parts:
                            continue
                        product_id = _normalize_usb_id(parts[0])
                        product_name = parts[1].strip() if len(parts) > 1 else ""
                        vendors[current_vendor]["products"][product_id] = {
                            "name": product_name,
                            "interfaces": [],
                        }
                        current_product = product_id
                        continue
                    # ベンダー情報
                    parts = line.split(None, 1)
                    if not parts:
                        continue
                    vendor_id = _normalize_usb_id(parts[0])
                    vendor_name = parts[1].strip() if len(parts) > 1 else ""
                    vendors.setdefault(vendor_id, {"name": vendor_name, "products": {}})
                    vendors[vendor_id]["name"] = vendor_name
                    current_vendor = vendor_id
                    current_product = None
        except OSError:
            return {}
        return vendors




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


def setup_services() -> Tuple[UsbSnapshotService, List[UsbDeviceSnapshot], Optional[str]]:
    """スキャナ・リポジトリ・サービスを組み立て、最新スナップショットを取得する。"""
    scanner = UsbScanner()
    repository = UsbSnapshotRepository(USB_JSON_PATH)
    service = UsbSnapshotService(scanner, repository)
    snapshots, scan_error = service.refresh()
    print(f"[DEBUG] scan() snapshots count: {len(snapshots)}")
    if not snapshots:
        print("[DEBUG] USBデバイスが1つも取得できませんでした")
    if scan_error:
        print(scan_error, file=sys.stderr)
    return service, snapshots, scan_error


def run_cli(service: UsbSnapshotService, ids_db: UsbIdsDatabase, args: argparse.Namespace) -> int:
    """CLI要求を処理し、指定VID/PIDのUSBデバイス情報を標準出力へ表示する。"""
    # 例) `python main.py 0x25a4 0x9311` の処理本体。
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
    # 例) `python main.py --self-test` で実行される診断ルーチン。

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
    service, snapshots, scan_error = setup_services()
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
