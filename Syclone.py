import os
import re
import time
import threading
import pythoncom
import serial
import subprocess
import winsound
import tkinter as tk
import win32com.client as win32
import matplotlib.pyplot as plt

from datetime import datetime
from pathlib import Path
from tkinter import filedialog
from serial.tools import list_ports
from queue import Queue
from collections import deque

excel_app = None
excel_sink = None
excel_filepath = None

write_in_progress = False
completion_sound_played = False
measurement_queue = Queue() # global
first_sample_received = False
excel_alive = True
excel_closed = False

active_fill = {
    "sheet": None,
    "row": None,
    "col": None,
    "count": 0,
    "next_index": 0
}

HERE = os.path.dirname(os.path.abspath(__file__))
WAV_PATH = os.path.join(HERE, "ready.wav")

# For live plotting
plot_x = deque(maxlen=500)   # time (s)
plot_y = deque(maxlen=500)   # dose (nSv/h)
plot_lock = threading.Lock()
plot_start_time = time.time()

# ---------------------------------------------------------------------------
# Excel MsgBox helper using raw COM Invoke (always works)
# ---------------------------------------------------------------------------
def excel_msgbox(excel_app, text, buttons=4+32+256, title="Prompt"):
    """
    Calls Excel.Application.MsgBox directly using its COM dispatch ID,
    bypassing missing attributes in win32com wrappers.
    """
    DISPID_MSGBOX = 0x3EB  # = 1003 decimal

    return excel_app._oleobj_.InvokeTypes(
        DISPID_MSGBOX,
        0,
        1,                      # DISPATCH_METHOD
        (3, 0),                # return type: int (VbMsgBoxResult)
        ((8, 1), (3, 1), (8, 1)),  # (text, buttons, title)
        text,
        buttons,
        title
    )

# ---------------------------------------------------------------------------
def excel_is_running():
    try:
        proc = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE"],
            capture_output=True,
            text=True,
            creationflags=0x08000000
        )
        return "EXCEL.EXE" in proc.stdout
    except Exception:
        return True

# ---------------------------------------------------------------------------
def try_fill_next_cell(value):
    """Write next Syclone sample into Excel measurement cell(s)."""
    global active_fill, completion_sound_played

    sheet = active_fill.get("sheet")
    positions = active_fill.get("positions")

    if sheet is None or not positions:
        return

    idx = active_fill.get("next_index", 0)
    if idx >= len(positions):
        return  # all samples filled

    row, col = positions[idx]
    cell = sheet.Cells(row, col)
    cell.Value = float(value)

    active_fill["next_index"] = idx + 1
    
    # -------------------------------------------------
    # COMPLETION DETECTED
    # -------------------------------------------------
    if active_fill["next_index"] == len(positions):
        if not completion_sound_played:
            completion_sound_played = True
            print("[Syclone] Last sample collected.")
            completion_whistle()


# ---------------------------------------------------------------------------
# Excel event handler
# ---------------------------------------------------------------------------
class ExcelEvents:
    def OnSheetChange(self, sheet, target):
        global pending_ops

        try:
            text = str(target.Value).strip()
        except Exception:
            return

        sheet_name = sheet.Name
        addr = target.Address
        key = (sheet_name, addr)

        # # -----------------------------------------------------------
        # # 1) Is this a reply (Y/N) to a pending overwrite question?
        # # -----------------------------------------------------------
        # if key in pending_ops and text:
        #     entry = pending_ops[key]
        #     row   = entry["row"]
        #     col   = entry["col"]
        #     rows  = entry["rows"]
        #     cols  = entry["cols"]
        #     old_width = entry["old_width"]

        #     # Restore original width
        #     col_obj = sheet.Columns(col)
        #     col_obj.ColumnWidth = old_width

        #     # Remove pending state
        #     del pending_ops[key]

        #     reply = text.upper()

        #     if reply == "Y":
        #         # Clear prompt cell formatting
        #         target.Value = ""
        #         target.Interior.ColorIndex = 0
        #         target.Font.Bold = False

        #         # Overwrite and build grid
        #         self.generate_grid(sheet, row, col, rows, cols, timestamp_mode)

        #     elif reply == "N":
        #         # User cancelled: just clear prompt cell
        #         target.Value = ""
        #         target.Interior.ColorIndex = 0
        #         target.Font.Bold = False

        #     else:
        #         # Invalid answer, restore pending state
        #         pending_ops[key] = {
        #             "row": row,
        #             "col": col,
        #             "rows": rows,
        #             "cols": cols,
        #             "old_width": old_width,
        #         }

        #     return

        # -----------------------------------------------------------
        # 2) Fresh command: parse "Syclone N" or "Syclone R x C"
        # -----------------------------------------------------------
        # Must start with "Syclone"
        if not text.lower().startswith("syclone"):
            return

        # "Syclone R x C"  (or "Syclone R xC", "Syclone R x  C", etc.)
        m_grid = re.match(r"(?i)^syclone\s+(\d+)\s*[x×]\s*(\d+)\s*$", text)
        m_line = re.match(r"(?i)^syclone\s+(\d+)\s*$", text)

        if m_grid:
            rows = int(m_grid.group(1))
            cols = int(m_grid.group(2))
        elif m_line:
            rows = 1
            cols = int(m_line.group(1))
            timestamp_mode = True
        else:
            # Not a recognized command
            return

        if rows <= 0 or cols <= 0:
            return

        row = target.Row
        col = target.Column

        target.Value = ""
        self.generate_grid(sheet, row, col, rows, cols, timestamp_mode)

    def OnSheetSelectionChange(self, sheet, target):
        print(f"[Excel] Selection changed to {sheet.Name}!{target.Address}")

    # Application-level event (fires only if Excel fully quits)
    def OnQuit(self):
        global excel_closed
        print("[Excel] Excel quitting.")
        excel_closed = True

    # -------------------------------------------------------------------
    # Grid builder: rows x cols bands of (title, measurement, spacer)
    # -------------------------------------------------------------------
    def generate_grid(self, sheet, start_row, start_col, rows, cols, timestamp_mode=False):
        """
        Builds a dense grid where:
          - Each Syclone command consumes ONE row
          - Timestamp + samples are on the same row
          - No spacer rows
          - No titles
          - No borders
        """
        global active_fill, completion_sound_played

        completion_sound_played = False
        positions = []

        for r in range(rows):
            row = start_row + r   # <-- DENSE rows, no gaps

            # Timestamp cell
            if timestamp_mode:
                ts_cell = sheet.Cells(row, start_col)
                ts_cell.Value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ts_cell.Font.Bold = True
                ts_cell.HorizontalAlignment = -4108
                col_offset = 1
            else:
                col_offset = 0

            # Sample cells
            for c in range(cols):
                col = start_col + col_offset + c
                dcell = sheet.Cells(row, col)
                dcell.Value = ""
                dcell.HorizontalAlignment = -4108
                positions.append((row, col))

        active_fill = {
            "sheet": sheet,
            "positions": positions,
            "next_index": 0
        }

        # Keep cursor on timestamp cell
        try:
            sheet.Application.Goto(sheet.Cells(start_row, start_col))
        except Exception:
            pass

class WorkbookEvents:
    def OnBeforeClose(self, Cancel):
        global excel_closed
        print("[WorkbookEvents] Workbook BeforeClose fired.")
        excel_closed = True

# ---------------------------------------------------------------------------
def excel_thread():
    global excel_app, excel_workbook, excel_app_events, excel_book_events, excel_filepath, excel_closed

    print("[Excel thread] starting...")
    pythoncom.CoInitialize()

    try:
        excel_app = win32.Dispatch("Excel.Application")
        excel_app.Visible = True
        print("[Excel thread] Excel.Application created.")
    except Exception as e:
        print("Failed to start Excel:", e)
        return

    try:
        excel_workbook = excel_app.Workbooks.Open(str(Path(excel_filepath).resolve()), ReadOnly=False)
        print("[Excel thread] Workbook loaded:", excel_workbook.Name)
    except Exception as e:
        print("Failed to open workbook:", e)
        return

    try:
        # Application-level events (SheetChange, Quit, etc.)
        excel_app_events = win32.WithEvents(excel_app, ExcelEvents)
        # Workbook-level events (BeforeClose)
        excel_book_events = win32.WithEvents(excel_workbook, WorkbookEvents)
        print("[Excel thread] Excel event listeners attached.\n")
    except Exception as e:
        print("Failed to attach Excel events:", e)
        return

    try:
        while True:
            pythoncom.PumpWaitingMessages()

            # If the workbook is closing, bail out
            if excel_closed:
                print("[Excel thread] excel_closed flag set. Exiting Excel thread.")
                break

            # Fill samples
            while not measurement_queue.empty():
                try_fill_next_cell(measurement_queue.get())

            time.sleep(0.05)

    finally:
        pythoncom.CoUninitialize()
        print("[Excel thread] Excel thread ended.")


# ---------------------------------------------------------------------------
# Serial devices
# ---------------------------------------------------------------------------
def find_bluetooth_serial_ports():
    from serial.tools import list_ports
    candidates = []

    for p in list_ports.comports():
        desc = p.description.lower()
        hwid = p.hwid.lower()

        if ("bluetooth" in desc or "spp" in desc or "serial" in desc) and "usb" not in desc:
            candidates.append(p.device)

    return candidates

# ---------------------------------------------------------------------------
def detect_syclone_data_port(calibration_seconds=2.0, baudrate=115200):
    """
    Detects which Bluetooth serial port receives Syclone data.
    Returns an *open* serial.Serial object ready to use.
    """

    ports = find_bluetooth_serial_ports()
    if not ports:
        print("No Syclone-like Bluetooth ports found.")
        return None

    print("Candidate ports:", ports)
    print(f"Trigger one Syclone measurement in the next {calibration_seconds} seconds...")

    ser_map = {}
    byte_counts = {p: 0 for p in ports}

    # --- Phase 1: open all candidate ports
    for p in ports:
        try:
            ser = serial.Serial(p, baudrate=baudrate, timeout=0.1)
            ser_map[p] = ser
            print(f"  Opened {p}")
        except Exception as e:
            print(f"  Could not open {p}: {e}")

    if not ser_map:
        print("Could not open any candidate ports.")
        return None

    # --- Phase 2: read incoming data
    t0 = time.time()
    while time.time() - t0 < calibration_seconds:
        for p, ser in ser_map.items():
            try:
                waiting = ser.in_waiting
                if waiting:
                    data = ser.read(waiting)
                    byte_counts[p] += len(data)
            except Exception:
                pass
        time.sleep(0.05)

    # --- Phase 3: pick the port with the most data
    best_port = max(byte_counts, key=byte_counts.get)
    if byte_counts[best_port] == 0:
        print("No data observed on any candidate port.")
        # Close everything
        for ser in ser_map.values():
            try: ser.close()
            except: pass
        return None

    print(f"Detected Syclone port: {best_port} ({byte_counts[best_port]} bytes)")

    # --- Phase 4: keep BEST port open, close the others
    winning_ser = ser_map[best_port]

    for p, ser in ser_map.items():
        if p != best_port:
            try:
                ser.close()
                print(f"Closed unused port {p}")
            except:
                pass

    # We return the OPEN serial.Serial object
    return winning_ser

# ---------------------------------------------------------------------------
def get_syclone_port():
    port = detect_syclone_data_port()
    if port:
        print(f"Syclone data port = {port}")
        return port

    print("Syclone not detected.")
    return None

# ---------------------------------------------------------------------------
def pick_file():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx;*.xlsm;*.xls")]
    )

# ---------------------------------------------------------------------------
def bcd_to_int(b):
    return (b >> 4) * 10 + (b & 0x0F)

# ---------------------------------------------------------------------------
def parse_packet(pkt):
    if len(pkt) != 50 or pkt[0] != 0x43 or pkt[1] != 0x59:
        return None

    # Raw dose rate in nR/s from packet
    dose_nrs = int.from_bytes(pkt[38:42], "little") * 0.1  # 0.1 nR/s units

    # Convert nR/s -> nSv/h  (nSv/h = 36 * nR/s)
    dose_nsvh = dose_nrs * 36.0

    year   = 2000 + bcd_to_int(pkt[22])
    month  = bcd_to_int(pkt[23])
    day    = bcd_to_int(pkt[24])
    hour   = bcd_to_int(pkt[25])
    minute = bcd_to_int(pkt[26])
    second = bcd_to_int(pkt[27])

    timestamp = f"{year:04d}-{month:02d}-{day:02d} {hour:02d}:{minute:02d}:{second:02d}"

    # Now returns dose in nSv/h
    return dose_nsvh, timestamp

# ---------------------------------------------------------------------------
def syclone_listener_thread(ser):
    """
    Continuously reads Syclone packets from an *open* serial port,
    parses them, and feeds:
      - the console (print)
      - the live matplotlib plot (via plot_x / plot_y)
    """
    global plot_start_time

    print("[Syclone] Listener thread started.")
    buffer = bytearray()

    while True:
        try:
            chunk = ser.read(50)
            if not chunk:
                continue

            buffer.extend(chunk)

            # Process complete 50-byte packets
            while len(buffer) >= 50:
                pkt = buffer[:50]
                buffer = buffer[50:]

                parsed = parse_packet(pkt)
                if parsed:
                    dose_nsvh, timestamp = parsed
                    now = time.time()

                    global first_sample_received
                    if not first_sample_received:
                        first_sample_received = True

                    print(f"[Syclone] {dose_nsvh:6.1f} nSv/h  @ {timestamp}")

                    with plot_lock:
                        if not plot_x:
                            plot_start_time = now
                        t_rel = now - plot_start_time
                        plot_x.append(t_rel)
                        plot_y.append(dose_nsvh)

                    # Queue for Excel
                    measurement_queue.put(dose_nsvh)

        except Exception as e:
            print("[Syclone] Error in listener:", e)
            break


# ---------------------------------------------------------------------------
def is_excel_running():
    """
    Returns True if EXCEL.EXE is running, False otherwise.
    Uses 'tasklist', so this is Windows-only (which is fine here).
    """
    try:
        proc = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE"],
            capture_output=True,
            text=True,
            creationflags=0x08000000  # CREATE_NO_WINDOW to avoid flashing a console
        )
        return "EXCEL.EXE" in proc.stdout
    except Exception:
        # If something goes wrong, assume Excel is still running to avoid
        # killing the app spuriously.
        return True

# ---------------------------------------------------------------------------        
def completion_whistle():
    winsound.PlaySound(WAV_PATH, winsound.SND_FILENAME)
        
# ---------------------------------------------------------------------------        
def main():
    global excel_filepath, first_sample_received

    # 1) Connect to Syclone first
    syclone_ser = detect_syclone_data_port()
    if syclone_ser:
        print("Syclone serial stream is live.")

        t2 = threading.Thread(target=syclone_listener_thread, args=(syclone_ser,))
        t2.daemon = True
        t2.start()
    else:
        print("Syclone not detected. Exiting.")
        return

    # 2) Wait for first sample BEFORE calling pick_file()
    print("[Waiting for first Syclone sample to start the plot...]")
    while not first_sample_received:
        time.sleep(0.05)

    print("[First sample received. Launching live plot...]")

    # 3) Create plot now (before Excel dialogs)
    plt.ion()
    fig, ax = plt.subplots()
    line, = ax.plot([], [], "b-")
    ax.set_title("Syclone Dose Rate")
    ax.set_xlabel("Time (s)")
    ax.set_ylabel("Dose (nSv/h)")
    plt.show(block=False)

    # 4) NOW load Excel (file dialog no longer blocks the plot)
    filepath = pick_file()
    if not filepath:
        print("[No file selected]")
        return

    excel_filepath = filepath

    print("Starting Excel event thread...")

    t = threading.Thread(target=excel_thread)
    t.daemon = True
    t.start()

    print("[Listening for Excel events. Plot running. CTRL+C to exit]\n")

    # 5) Plot update loop + Excel window watchdog
    try:
        while True:

            # ---- Plot updates ----
            with plot_lock:
                xs = list(plot_x)
                ys = list(plot_y)

            if xs:
                line.set_data(xs, ys)

                xmax = xs[-1]
                xmin = max(0, xmax - 300)
                ax.set_xlim(xmin, xmax + 1)

                ymin = min(ys)
                ymax = max(ys)
                if ymin == ymax:
                    ymin -= 1
                    ymax += 1
                ax.set_ylim(ymin * 0.9, ymax * 1.1)

            fig.canvas.draw_idle()
            fig.canvas.flush_events()
            time.sleep(0.1)

            # ---- Excel watchdog ----
            if excel_closed:
                print("[Excel] Excel was closed. Exiting...")
                os._exit(0)

    except KeyboardInterrupt:
        print("\nStopped.")


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
    completion_whistle()
    