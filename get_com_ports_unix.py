import sys
import platform

try:
    import serial.tools.list_ports as list_ports
except ImportError:
    print("pyserialが必要です。pip install pyserial を実行してください。", file=sys.stderr)
    sys.exit(1)

def get_com_ports():
    com_ports = []
    for p in list_ports.comports():
        # USB-Serialデバイスのみ抽出（Mac/Linux共通）
        if p.vid is not None and p.pid is not None:
            com_ports.append({
                "device": p.device,
                "description": p.description,
                "hwid": p.hwid,
                "vid": hex(p.vid),
                "pid": hex(p.pid),
                "serial_number": p.serial_number,
                "manufacturer": p.manufacturer,
                "product": p.product,
            })
    return com_ports

if __name__ == "__main__":
    os_name = platform.system()
    print(f"OS: {os_name}")
    for port in get_com_ports():
        print(port)
