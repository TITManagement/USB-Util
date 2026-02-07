# -*- coding: utf-8 -*-
"""
COM Port Inspector (FULL) - PowerShell不要版 (Windows)
- pyserial: COM列挙
- WMI: PnP情報 / USBコントローラ関連 / LocationInformation（Port_# / Hub_#）
- 分類: USB / Bluetooth / RS-232(内蔵/PCI) / Unknown
- 出力: 人間向け表示 + JSON

依存:
  pip install pyserial wmi
実行:
  python com_port_inspector_full.py
"""




# -------------------- utils --------------------

def norm(s: Optional[str]) -> str:
    return (s or "").strip()

def upper(s: Optional[str]) -> str:
    return norm(s).upper()

VIDPID_RE = re.compile(r"VID_([0-9A-Fa-f]{4}).*PID_([0-9A-Fa-f]{4})")
PORT_TOKEN_RE = re.compile(r"(Port_#\d+|Hub_#\d+)", re.IGNORECASE)

def parse_vid_pid(s: str) -> Tuple[Optional[str], Optional[str]]:
    m = VIDPID_RE.search(s or "")
    if m:
        return m.group(1).upper(), m.group(2).upper()
    return None, None

def parse_serial_tail(pnp_device_id: str) -> str:
    parts = re.split(r"[\\#]", pnp_device_id or "")
    return parts[-1] if parts else ""

def parse_location_chain(location_info: str) -> List[str]:
    if not location_info:
        return []
    return PORT_TOKEN_RE.findall(location_info)

def safe_int_hex(n: Optional[int]) -> Optional[str]:
    return f"{n:04X}" if n is not None else None


# -------------------- data models --------------------

@dataclass
class PortInfo:
    com: str
    description: str
    hwid: str
    vid_hex: Optional[str]
    pid_hex: Optional[str]
    serial: Optional[str]
    location: Optional[str]
    manufacturer: Optional[str]
    product: Optional[str]


@dataclass
class WmiPnP:
    device_id: str
    name: str
    pnp_class: str
    manufacturer: str
    location_info: str
    status: str


# -------------------- gather from pyserial --------------------

def gather_ports() -> List[PortInfo]:
    out: List[PortInfo] = []
    for p in list_ports.comports():
        hwid = norm(getattr(p, "hwid", ""))
        vid_hex = safe_int_hex(getattr(p, "vid", None))
        pid_hex = safe_int_hex(getattr(p, "pid", None))

        serial = norm(getattr(p, "serial_number", "")) or ""
        if not serial and hwid:
            serial = parse_serial_tail(hwid)

        if (not vid_hex or not pid_hex) and hwid:
            v2, p2 = parse_vid_pid(hwid)
            vid_hex = vid_hex or v2
            pid_hex = pid_hex or p2

        out.append(
            PortInfo(
                com=p.device,
                description=norm(getattr(p, "description", "")),
                hwid=hwid,
                vid_hex=vid_hex,
                pid_hex=pid_hex,
                serial=serial or None,
                location=norm(getattr(p, "location", "")) or None,
                manufacturer=norm(getattr(p, "manufacturer", "")) or None,
                product=norm(getattr(p, "product", "")) or None,
            )
        )
    return out


# -------------------- WMI indices --------------------

def build_pnp_index(w: wmi.WMI) -> Dict[str, WmiPnP]:
    idx: Dict[str, WmiPnP] = {}
    for dev in w.Win32_PnPEntity():
        did = norm(getattr(dev, "DeviceID", ""))  # = PNPDeviceID相当
        if not did:
            continue
        idx[did] = WmiPnP(
            device_id=did,
            name=norm(getattr(dev, "Name", "")),
            pnp_class=norm(getattr(dev, "PNPClass", "")),
            manufacturer=norm(getattr(dev, "Manufacturer", "")),
            location_info=norm(getattr(dev, "LocationInformation", "")),
            status=norm(getattr(dev, "Status", "")),
        )
    return idx

def build_controller_names(w: wmi.WMI) -> Dict[str, str]:
    names: Dict[str, str] = {}
    for c in w.Win32_USBController():
        did = norm(getattr(c, "DeviceID", ""))
        if did:
            names[did] = norm(getattr(c, "Name", "")) or did
    return names

def map_dependent_to_controllers(w: wmi.WMI) -> Dict[str, List[str]]:
    """
    Win32_USBControllerDevice:
      Antecedent: Win32_USBController
      Dependent : Win32_PnPEntity
    依存デバイス(DeviceID) -> コントローラ(DeviceID) の逆引きを作る
    """
    dep_to_ctrl = defaultdict(list)

    def extract_deviceid(relpath: str) -> Optional[str]:
        m = re.search(r'DeviceID="([^"]+)"', relpath or "")
        return m.group(1) if m else None

    for rel in w.Win32_USBControllerDevice():
        dep_path = getattr(rel, "Dependent", None)
        ant_path = getattr(rel, "Antecedent", None)
        dep_id = extract_deviceid(dep_path)
        ant_id = extract_deviceid(ant_path)
        if dep_id and ant_id:
            dep_to_ctrl[dep_id].append(ant_id)

    return dep_to_ctrl


# -------------------- classification --------------------

def classify_port(port: PortInfo, pnp: Optional[WmiPnP]) -> Tuple[str, float, List[str]]:
    """
    kind: USB / Bluetooth / RS-232/PCI/ACPI / Unknown
    """
    hw = upper(port.hwid)
    desc = upper(port.description)
    reasons: List[str] = []

    # Bluetooth
    if "BTHENUM" in hw or "BLUETOOTH" in hw or "RFCOMM" in desc or "BLUETOOTH" in desc:
        reasons.append("hwid/desc indicates Bluetooth (BTHENUM/BLUETOOTH/RFCOMM)")
        return "Bluetooth", 0.95, reasons

    # USB
    if port.vid_hex and port.pid_hex:
        reasons.append("VID/PID available (likely USB-serial)")
        return "USB", 0.92, reasons
    if hw.startswith("USB\\") or " USB" in (" " + desc):
        reasons.append("hwid/desc suggests USB")
        return "USB", 0.75, reasons

    # ACPI/PCI serial
    if hw.startswith("ACPI\\") or hw.startswith("PCI\\") or "PNP0501" in hw or "PNP0500" in hw:
        reasons.append("hwid indicates ACPI/PCI standard serial")
        return "RS-232/PCI/ACPI", 0.85, reasons

    if pnp and upper(pnp.pnp_class) == "PORTS":
        if "USB\\" not in hw and "BTHENUM" not in hw:
            reasons.append("PnPClass=Ports but not USB/Bluetooth")
            return "RS-232/PCI/ACPI", 0.60, reasons

    return "Unknown", 0.30, ["no strong indicators found"]


# -------------------- correlation / reporting --------------------

def correlate_full() -> List[dict]:
    if platform.system().lower() != "windows":
        raise RuntimeError("Windows-only: this script uses WMI.")

    w = wmi.WMI()
    ports = gather_ports()

    pnp_idx = build_pnp_index(w)
    dep_to_ctrl = map_dependent_to_controllers(w)
    ctrl_names = build_controller_names(w)

    results = []
    for p in ports:
        # pyserial.hwid が WMI Win32_PnPEntity.DeviceID と一致することが多い
        pnp = pnp_idx.get(p.hwid)

        if not pnp:
            # 大文字小文字差などのゆるい一致
            ph = upper(p.hwid)
            for k, v in pnp_idx.items():
                if upper(k) == ph:
                    pnp = v
                    break

        loc_info = pnp.location_info if pnp else ""
        chain = parse_location_chain(loc_info)

        controllers = []
        for ctrl_id in dep_to_ctrl.get(p.hwid, []):
            controllers.append(ctrl_names.get(ctrl_id, ctrl_id))

        kind, conf, reasons = classify_port(p, pnp)

        results.append({
            "com": p.com,
            "kind": kind,
            "confidence": conf,
            "reasons": reasons,

            "description": p.description,
            "hwid": p.hwid,

            "vid_pid": f"{p.vid_hex}:{p.pid_hex}" if p.vid_hex and p.pid_hex else None,
            "serial_guess": p.serial,

            "pyserial_location": p.location,              # 例: "1-4.3"
            "wmi_location_information": loc_info,         # 例: "Port_#0003.Hub_#0006"
            "wmi_chain": chain,                           # 例: ["Port_#0003","Hub_#0006"]

            "usb_host_controllers": controllers,
            "wmi_device_name": pnp.name if pnp else None,
            "wmi_manufacturer": pnp.manufacturer if pnp else None,
            "wmi_status": pnp.status if pnp else None,
            "wmi_pnp_class": pnp.pnp_class if pnp else None,
        })

    return results


def pretty_print(results: List[dict]) -> None:
    print("=== COM Port Inspector (FULL / no PowerShell) ===\n")
    for r in results:
        print(f"{r['com']:8} [{r['kind']:<14}] conf={r['confidence']:.2f}  {r['description']}")
        print(f"  HWID: {r['hwid']}")
        if r["vid_pid"]:
            print(f"  VID:PID: {r['vid_pid']}  SerialGuess: {r['serial_guess']}")
        else:
            print(f"  SerialGuess: {r['serial_guess']}")

        if r["wmi_device_name"] or r["wmi_manufacturer"]:
            print(f"  WMI: name='{r['wmi_device_name']}'  manu='{r['wmi_manufacturer']}'  class='{r['wmi_pnp_class']}'  status='{r['wmi_status']}'")

        if r["wmi_chain"]:
            print(f"  Topology (WMI): {' -> '.join(r['wmi_chain'])}   (from LocationInformation)")
        elif r["wmi_location_information"]:
            print(f"  Topology (WMI): (unparsed) {r['wmi_location_information']}")
        else:
            print("  Topology (WMI): (none)")

        if r["pyserial_location"]:
            print(f"  Topology (pyserial fallback): {r['pyserial_location']}   (e.g. root-port.subport)")

        if r["usb_host_controllers"]:
            for i, name in enumerate(r["usb_host_controllers"], 1):
                print(f"  HostController[{i}]: {name}")
        else:
            print("  HostController: (unknown)")

        print(f"  Reasons: {'; '.join(r['reasons'])}")
        print("")


def main():
    results = correlate_full()
    pretty_print(results)
    out = "com_port_inspector_full.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()