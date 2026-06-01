import os
import sys
import time
from pathlib import Path
from winrt.windows.devices.bluetooth.advertisement import (
    BluetoothLEAdvertisementWatcher,
    BluetoothLEScanningMode
)
from winrt.windows.storage.streams import DataReader
import win32com.client as win32

from tiltclass import TiltClass
from excelexportclass import ExcelExportClass

# Tilt UUIDs (normalized)
TILTS = {
    'a495bb10c5b14b44b5121370f02d74de': 'Red',
    'a495bb20c5b14b44b5121370f02d74de': 'Green',
    'a495bb30c5b14b44b5121370f02d74de': 'Black',
    'a495bb40c5b14b44b5121370f02d74de': 'Purple',
    'a495bb50c5b14b44b5121370f02d74de': 'Orange',
    'a495bb60c5b14b44b5121370f02d74de': 'Blue',
    'a495bb70c5b14b44b5121370f02d74de': 'Yellow',
    'a495bb80c5b14b44b5121370f02d74de': 'Pink'
}

def parse_ibeacon_payload(data_bytes):
    b = bytes(data_bytes)
    if len(b) < 23:
        return None

    for i in range(0, len(b) - 22):
        if b[i] == 0x02 and b[i+1] == 0x15:
            start = i + 2
            uuid = ''.join(f"{x:02x}" for x in b[start:start+16])
            major = int.from_bytes(b[start+16:start+18], 'big')
            minor = int.from_bytes(b[start+18:start+20], 'big')
            return uuid, major, minor
    return None

class OneShotTiltCapture:
    def __init__(self):
        self.captured = False

        # Bluetooth watcher
        self.watcher = BluetoothLEAdvertisementWatcher()
        self.watcher.scanning_mode = BluetoothLEScanningMode.ACTIVE
        self.watcher.add_received(self._on_received)

        # Tilt objects
        self.tilts = {uuid: TiltClass(uuid, name) for uuid, name in TILTS.items()}

    # -----------------------------
    # Excel Export
    # -----------------------------
    def get_base_path(self):
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def _write_excel(self, tilt, temp_f, gravity, rssi):
        exporter = ExcelExportClass()

        base_path = self.get_base_path()
        file_path = os.path.join(base_path, "Tilt_Reading.xlsx")
        
        wb = exporter.open_file(file_path, tilt.tiltName)
        ws = exporter.open_sheet(wb, tilt.tiltName)

        exporter.export_to_file(ws, tilt.tiltName, temp_f, gravity, rssi)

        wb.Save()
        wb.Close(SaveChanges=True)

        exporter.excel.Quit()

        for name in ("ws", "wb", "exporter"):
            try:
                del name
            except Exception:
                pass

    # -----------------------------
    # Bluetooth callback
    # -----------------------------
    def _on_received(self, sender, args):
        if self.captured:
            return
        
        adv = args.advertisement
        rssi = args.raw_signal_strength_in_dbm 
        
        if len(adv.manufacturer_data) == 0:
            return

        for md in adv.manufacturer_data:
            reader = DataReader.from_buffer(md.data)
            buf = reader.read_buffer(md.data.length)
            data = bytes(memoryview(buf))

            parsed = parse_ibeacon_payload(data)
            if not parsed:
                continue

            uuid, major, minor = parsed
            uuid = uuid.lower()

            if uuid in self.tilts:
                tilt = self.tilts[uuid]

                # Update using TiltClass logic
                tilt.tiltUpdate(major, minor)

                # Get converted values using class methods
                temp_c = tilt.tempCelsius()
                temp_f = tilt.tempFahrenheit()
                gravity = tilt.specificGravity()

                print(f"Tilt Found: {tilt.tiltName}")
                print(f"Temp (°F): {temp_f}")
                print(f"Gravity: {gravity:.4f}")
                print(f"RSSI: {rssi} dBm")

                self._write_excel(tilt, temp_f, gravity, rssi)

                self.captured = True
                self.watcher.stop()
                return
        
    def run(self):
        print("Waiting for Tilt broadcast...")
        self.watcher.start()

        while not self.captured:
            time.sleep(0.1)


if __name__ == "__main__":
    OneShotTiltCapture().run()

