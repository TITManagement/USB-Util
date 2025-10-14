import sys
import win32com.client

def get_com_ports():
    com_ports = []
    try:
        wmi = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        for device in wmi.ConnectServer(".", "root\\cimv2").ExecQuery("SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'"):
            # 例: USB\VID_XXXX&PID_YYYY\SERIAL
            com_ports.append({
                "name": device.Name,
                "pnp_id": device.PNPDeviceID
            })
    except Exception as e:
        print("COMポート情報取得エラー:", e, file=sys.stderr)
    return com_ports

if __name__ == "__main__":
    for port in get_com_ports():
        print(port)
