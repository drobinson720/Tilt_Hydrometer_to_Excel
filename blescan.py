from winrt.windows.devices.bluetooth.advertisement import (
    BluetoothLEAdvertisementWatcher,
    BluetoothLEScanningMode
)
from winrt.windows.storage.streams import DataReader
import time
from typing import List, Dict, Optional


def _bytes_to_hex_string(b: bytes) -> str:
    return ''.join(f'{x:02x}' for x in b)

def _bytes_to_int_be(b: bytes) -> int:
    return int.from_bytes(b, byteorder='big')

def _parse_ibeacon_from_bytes(b: bytes) -> Optional[Dict]:
    if not b or len(b) < 25:
        return None
    for i in range(0, len(b) - 24):
        if b[i] == 0x02 and b[i+1] == 0x15:
            start = i + 2
            uuid = _bytes_to_hex_string(b[start:start+16])
            major = _bytes_to_int_be(b[start+16:start+18])
            minor = _bytes_to_int_be(b[start+18:start+20])
            return {'uuid': uuid, 'major': major, 'minor': minor}
    return None


class WinRTScanner:
    def __init__(self):
        self.watcher = BluetoothLEAdvertisementWatcher()
        self.watcher.scanning_mode = BluetoothLEScanningMode.ACTIVE

        self._results: List[Dict] = []
        self.watcher.add_received(self._on_received)

    def _on_received(self, sender, args):
        adv = args.advertisement

        for md in adv.manufacturer_data:
            reader = DataReader.from_buffer(md.data)
            data = bytes(reader.read_bytes(md.data.length))

            parsed = _parse_ibeacon_from_bytes(data)
            if parsed:
                self._results.append(parsed)

    def scan(self, timeout: float = 5.0) -> List[Dict]:
        self._results.clear()
        self.watcher.start()

        t0 = time.time()
        while time.time() - t0 < timeout:
            time.sleep(0.05)

        self.watcher.stop()
        return self._results.copy()


def parseEvents_async(timeout: float = 5.0) -> List[Dict]:
    scanner = WinRTScanner()
    return scanner.scan(timeout)


def parseEvents(sock=None, loop_count: int = 100) -> List[Dict]:
    per_iter = 0.05
    total_timeout = max(1.0, loop_count * per_iter)
    scanner = WinRTScanner()
    return scanner.scan(total_timeout)
