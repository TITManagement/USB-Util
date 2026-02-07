# -*- coding: utf-8 -*-
"""
PowerShell 不使用。WMI + pyserial だけで
  (1) COMポート列挙
  (2) PNPDeviceID / LocationInformation の解析
  (3) 所属 USB ホストコントローラ名の特定（Win32_USBControllerDevice 経由）
  (4) Port_#xxxx.Hub_#yyyy 連鎖の可視化（ハブ階層の番号列）
を行い、「どのハブの何番ポートか」をできるだけ具体的に出力する。

依存: pip install pyserial wmi
実行 (Windowsのみ): python tools/diagnostics/com_usb_topology_wmi.py
"""



# ---- utils ------------------------------------------------------------------

def norm(s): return (s or "").strip()
def upper(s): return norm(s).upper()

PORT_TOKEN_RE = re.compile(r"(Port_#\d+|Hub_#\d+)", re.IGNORECASE)

def parse_location_chain(location_info: str):
    """
    "LocationInformation" 例:
      "Port_#0003.Hub_#0006"
      "Port_#0002.Hub_#0004"
    を ['Port_#0003','Hub_#0006'] のような配列に。
    """
    if not location_info:
        return []
    return PORT_TOKEN_RE.findall(location_info)

def parse_vid_pid(pnp_device_id: str):
    m = re.search(r'VID_([0-9A-Fa-f]{4}).*PID_([0-9A-Fa-f]{4})', pnp_device_id or "")
    if m:
        return m.group(1).upper(), m.group(2).upper()
    return None, None

def parse_serial_from_pnpid(pnp_device_id: str):
    if not pnp_device_id:
        return ""
    parts = re.split(r'[\\#]', pnp_device_id)
    return parts[-1] if parts else ""

# ---- inventory via WMI ------------------------------------------------------

def list_com_ports_pyserial():
    rows = []
    for p in list_ports.comports():
        rows.append({
            "com": p.device,                         # "COM7"
            "description": norm(getattr(p, "description", "")),
            "vid_hex": f"{p.vid:04X}" if p.vid is not None else None,
            "pid_hex": f"{p.pid:04X}" if p.pid is not None else None,
            "serial": norm(getattr(p, "serial_number", "")),
            "hwid": norm(getattr(p, "hwid", "")),
            "location": norm(getattr(p, "location", "")),
            "manufacturer": norm(getattr(p, "manufacturer", "")),
            "product": norm(getattr(p, "product", "")),
        })
    return rows

def build_pnp_index(w):
    """
    Win32_PnPEntity を索引化。
    - key: DeviceID（=PNPDeviceID）
    - value: dict(name, manufacturer, class_guid, pnp_class, location_info, status)
    """
    idx = {}
    for dev in w.Win32_PnPEntity():
        did = norm(dev.DeviceID)
        if not did:
            continue
        idx[did] = {
            "device_id": did,
            "name": norm(dev.Name),
            "manufacturer": norm(dev.Manufacturer),
            "class_guid": norm(dev.ClassGuid),
            "pnp_class": norm(dev.PNPClass),
            "location_info": norm(dev.LocationInformation),
            "status": norm(dev.Status),
        }
    return idx

def map_entity_to_controller(w):
    """
    Win32_USBControllerDevice は
      Antecedent: Win32_USBController
      Dependent : Win32_PnPEntity  (USB配下デバイス)
    の関連。Dependent(=PnP) → Antecedent(=Controller) を逆引きテーブル化。
    """
    dep_to_ctrl = defaultdict(list)
    # 参照先オブジェクトは "__RELPATH" に WMI パスが入る。そこから DeviceID を抜く。
    def extract_deviceid(relpath):
        # 例: Win32_PnPEntity.DeviceID="USB\\VID_0403&PID_6001\\A50285BI"
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

def build_controller_names(w):
    """
    Win32_USBController の DeviceID => Name を収集
    例: "PCI\\VEN_8086&DEV_A12F&CC_0C03" => "Intel(R) USB 3.0 eXtensible Host Controller ..."
    """
    names = {}
    for c in w.Win32_USBController():
        did = norm(c.DeviceID)
        if did:
            names[did] = norm(c.Name)
    return names

# ---- correlate & report -----------------------------------------------------

def correlate_with_topology():
    if platform.system().lower() != "windows":
        print("Windows専用（WMI使用）。このスクリプトはWindows環境でのみ実行してください。")
        sys.exit(1)

    if wmi is None:
        print("Windows専用の wmi ライブラリが見つかりません。`pip install wmi` を実行してください。")
        sys.exit(1)

    w = wmi.WMI()

    coms = list_com_ports_pyserial()
    pnp_index = build_pnp_index(w)
    dep_to_ctrl = map_entity_to_controller(w)
    ctrl_names = build_controller_names(w)

    results = []
    for c in coms:
        # COM側 PNPDeviceID は pyserial.hwid に含まれることが多い
        pnpid = c.get("hwid", "")
        # hwidが "USB\\VID_xxxx&PID_yyyy\\SER" 形式ならそれを优先キーに
        if not pnpid.startswith("USB") and "USB" in pnpid:
            # hwid例: "USB VID:PID=0403:6001 SER=A50285BI LOCATION=1-3"
            # このケースは "USB\\VID_XXXX&PID_YYYY\\SER" を探す
            m = re.search(r'(USB\\VID_[0-9A-Fa-f]{4}&PID_[0-9A-Fa-f]{4}\\[^ ]+)', pnpid)
            if m:
                pnpid = m.group(1)

        vid_hex, pid_hex = c["vid_hex"], c["pid_hex"]
        if (not vid_hex or not pid_hex) and pnpid.startswith("USB\\VID_"):
            v2, p2 = parse_vid_pid(pnpid)
            vid_hex = vid_hex or v2
            pid_hex = pid_hex or p2

        serial_guess = c.get("serial") or parse_serial_from_pnpid(pnpid)

        # PnP情報から LocationInformation を拾う
        loc_info = ""
        if pnpid in pnp_index:
            loc_info = pnp_index[pnpid]["location_info"]
        # それが無ければ pyserial.location を使う（"1-4.3" 等）
        loc_fallback = c.get("location")

        chain = parse_location_chain(loc_info)
        # 所属コントローラ名（複数ある場合があるので全部拾う）
        controllers = []
        for ctrl_id in dep_to_ctrl.get(pnpid, []):
            name = ctrl_names.get(ctrl_id) or ctrl_id
            controllers.append(name)

        results.append({
            "com": c["com"],
            "description": c["description"],
            "vid_hex": vid_hex, "pid_hex": pid_hex,
            "serial": serial_guess,
            "pnp_device_id": pnpid,
            "location_information": loc_info,
            "location_fallback": loc_fallback,
            "port_hub_chain": chain,           # ['Port_#0003','Hub_#0006', ...]
            "usb_controllers": controllers,    # ["Intel(R) USB 3.2 ...", ...]
        })

    # 表示
    print("=== COM ↔ USB トポロジ（ハブ番号/ポート番号 & コントローラ名）===\n")
    for r in results:
        print(f"{r['com']}  {r['description']}")
        print(f"  VID:PID={r['vid_hex']}:{r['pid_hex']}  Serial={r['serial']}")
        print(f"  PNPDeviceID: {r['pnp_device_id']}")
        if r["port_hub_chain"]:
            print(f"  Chain: {' -> '.join(r['port_hub_chain'])}")
        else:
            print("  Chain: (LocationInformation なし / 解析不可)")
        if r["location_fallback"]:
            print(f"  Location(fallback): {r['location_fallback']}  (例: ルート=1, 下位=4.3)")
        if r["usb_controllers"]:
            for i, name in enumerate(r["usb_controllers"], 1):
                print(f"  HostController[{i}]: {name}")
        else:
            print("  HostController: (未特定)")
        print("")

    # JSON保存
    with open("com_usb_topology_wmi.json", "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print("JSONを書き出しました: com_usb_topology_wmi.json")

if __name__ == "__main__":
    correlate_with_topology()
