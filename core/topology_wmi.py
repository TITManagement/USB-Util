"""Windows-specific USB topology resolver using WMI."""

from __future__ import annotations

import platform
import re
from typing import Any, Dict, Iterable, List, Optional, Tuple

from core.models import UsbDeviceSnapshot

PORT_TOKEN_RE = re.compile(r"(Port_#\d+|Hub_#\d+)", re.IGNORECASE)


def annotate_windows_topology(snapshots: Iterable[UsbDeviceSnapshot]) -> None:
    """Augment snapshots with hub/port topology information on Windows."""

    if not snapshots:
        return
    if platform.system().lower() != "windows":
        return
    try:
        import wmi  # type: ignore
    except ImportError:
        return

    resolver = _TopologyResolver(wmi.WMI())
    mapping = resolver.build_mapping()
    if not mapping:
        return

    for snapshot in snapshots:
        key_with_serial = _snapshot_key(snapshot, include_serial=True)
        key_without_serial = _snapshot_key(snapshot, include_serial=False)
        entry = mapping.get(key_with_serial) or mapping.get(key_without_serial)
        if not entry:
            continue
        snapshot.topology_chain = entry.get("port_hub_chain", [])
        snapshot.location_information = entry.get("location_information", "")
        snapshot.location_fallback = entry.get("location_fallback", "")
        snapshot.usb_controllers = entry.get("usb_controllers", [])


class _TopologyResolver:
    def __init__(self, wmi_client: "wmi.WMI") -> None:  # type: ignore
        self._wmi = wmi_client

    def build_mapping(self) -> Dict[Tuple[str, str, str], Dict[str, List[str]]]:
        dep_to_ctrl = self._map_entity_to_controller()
        ctrl_names = self._controller_names()
        mapping: Dict[Tuple[str, str, str], Dict[str, List[str]]] = {}

        for dev in self._wmi.Win32_PnPEntity():
            device_id = _norm(getattr(dev, "DeviceID", ""))
            if not device_id.startswith("USB\\"):
                continue
            vid, pid = _parse_vid_pid(device_id)
            if not vid or not pid:
                continue
            serial = _parse_serial_from_pnpid(device_id)
            location_info = _norm(getattr(dev, "LocationInformation", ""))
            chain = _parse_location_chain(location_info)
            controllers = [
                ctrl_names.get(ctrl_id, ctrl_id)
                for ctrl_id in dep_to_ctrl.get(device_id, [])
            ]

            entry = {
                "pnp_device_id": device_id,
                "location_information": location_info,
                "location_fallback": "",
                "port_hub_chain": chain,
                "usb_controllers": controllers,
            }

            key_with_serial = (vid, pid, serial)
            key_without_serial = (vid, pid, "")
            mapping[key_with_serial] = entry
            mapping.setdefault(key_without_serial, entry)

        return mapping

    def _map_entity_to_controller(self) -> Dict[str, List[str]]:
        dep_to_ctrl: Dict[str, List[str]] = {}

        def extract_deviceid(relpath: Optional[str]) -> Optional[str]:
            if not relpath:
                return None
            match = re.search(r'DeviceID="([^\"]+)"', relpath)
            return match.group(1) if match else None

        for rel in self._wmi.Win32_USBControllerDevice():
            dep_path = getattr(rel, "Dependent", None)
            ant_path = getattr(rel, "Antecedent", None)
            dep_id = extract_deviceid(dep_path)
            ant_id = extract_deviceid(ant_path)
            if dep_id and ant_id:
                dep_to_ctrl.setdefault(dep_id, []).append(ant_id)

        return dep_to_ctrl

    def _controller_names(self) -> Dict[str, str]:
        names: Dict[str, str] = {}
        for ctrl in self._wmi.Win32_USBController():
            device_id = _norm(getattr(ctrl, "DeviceID", ""))
            if device_id:
                names[device_id] = _norm(getattr(ctrl, "Name", "")) or device_id
        return names


def _snapshot_key(snapshot: UsbDeviceSnapshot, *, include_serial: bool) -> Tuple[str, str, str]:
    vid = _normalize_vid_pid(snapshot.vid)
    pid = _normalize_vid_pid(snapshot.pid)
    serial = _normalize_serial(snapshot.serial) if include_serial else ""
    return vid, pid, serial


def _parse_location_chain(location_info: str) -> List[str]:
    if not location_info:
        return []
    return [token for token in PORT_TOKEN_RE.findall(location_info)]


def _parse_vid_pid(pnp_device_id: str) -> Tuple[str, str]:
    match = re.search(r"VID_([0-9A-Fa-f]{4}).*PID_([0-9A-Fa-f]{4})", pnp_device_id or "")
    if match:
        return match.group(1).upper(), match.group(2).upper()
    return "", ""


def _parse_serial_from_pnpid(pnp_device_id: str) -> str:
    if not pnp_device_id:
        return ""
    parts = re.split(r"[\\#]", pnp_device_id)
    return parts[-1].upper() if parts else ""


def _normalize_vid_pid(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    if text.startswith("0X"):
        text = text[2:]
    return text.zfill(4)


def _normalize_serial(value: Any) -> str:
    text = (str(value).strip() if value is not None else "").upper()
    if not text or text in {"取得不可", "-", ""}:
        return ""
    return text


def _norm(value: Optional[str]) -> str:
    return (value or "").strip()
