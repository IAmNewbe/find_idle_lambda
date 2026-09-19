"""
Copyright by Ahmad Ahlul Hikam
Idle Lambda Finder
==================
Input file dari "Manage WDM Trail" (.xlsx), ambil kolom
"Source Channel", pakai C-band 50GHz (80 channel) (disuruh pak Rupas ehe)
untuk menemukan channel yang masih kosong (idle).

Jalankan:
  python src/find_idle_lambda.py --config config/config.yaml
"""
import argparse
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

sys.path.append(str(Path(__file__).parent))
from utils import (
    load_config,
    find_input_files,
    open_workbook_safely,
    find_header_row,
    parse_channel_number,
)


def build_channel_plan(config: dict) -> dict:
    """
    cihuy
    """
    total = config["total_channels"]
    freq1 = config["freq_channel1_thz"]
    spacing_thz = config["spacing_ghz"] / 1000.0
    plan = {}
    for n in range(1, total + 1):
        freq = round(freq1 - (n - 1) * spacing_thz, 3)
        wavelength = round(C_NM_THZ / freq, 2)
        plan[n] = {"frequency_thz": freq, "wavelength_nm": wavelength}
    return plan


C_NM_THZ = 299792.458  # konstanta kecepatan cahaya (nm . THz), ngapain ini yak wkwk


def find_flex_grid_100g(channel_plan: dict, idle_list: list) -> list[dict]:
    """
    Cari slot Flex Grid 100G yang tersedia: dua channel 50G idle yang
    berurutan (n, n+1). Titik tengah frekuensi keduanya jadi center
    frequency slot 100G (lebar slot = 2 x spacing 50G = 100GHz,
    yaitu center +- 50GHz).
    """
    idle_set = set(idle_list)
    slots = []
    for n in sorted(idle_list):
        if (n + 1) not in idle_set:
            continue
        freq_a = channel_plan[n]["frequency_thz"]
        freq_b = channel_plan[n + 1]["frequency_thz"]
        center_freq = round((freq_a + freq_b) / 2, 4)
        center_wavelength = round(C_NM_THZ / center_freq, 2)
        half_slot_thz = 0.05  # 50 GHz = 0.05 THz
        slots.append(
            {
                "pair": f"{n} - {n + 1}",
                "channel_a": n,
                "channel_b": n + 1,
                "center_freq_thz": center_freq,
                "center_wavelength_nm": center_wavelength,
                "slot_low_thz": round(center_freq - half_slot_thz, 3),
                "slot_high_thz": round(center_freq + half_slot_thz, 3),
            }
        )
    return slots


def extract_usage_from_file(path: Path, config: dict) -> list[dict]:
    """Baca satu file, kembalikan list dict info pemakaian channel."""
    wb = open_workbook_safely(path)
    channel_col = config["channel_column"]
    info_cols = config.get("info_columns", [])

    usages = []
    for ws in wb.worksheets:
        header_row, headers = find_header_row(ws, channel_col)
        if header_row is None:
            continue  # sheet ini tidak punya kolom yang dicari, skip

        col_index = {name: idx for idx, name in enumerate(headers)}
        channel_idx = col_index[channel_col]

        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            if channel_idx >= len(row):
                continue
            channel_no = parse_channel_number(row[channel_idx])
            if channel_no is None:
                continue

            info = {}
            for col_name in info_cols:
                idx = col_index.get(col_name)
                info[col_name] = row[idx] if idx is not None and idx < len(row) else None

            usages.append(
                {
                    "channel": channel_no,
                    "source_file": path.name,
                    **info,
                }
            )
    return usages


def collect_all_usage(config: dict) -> dict:
    """Kumpulkan pemakaian channel dari semua file input. Return dict channel -> list of usage."""
    files = find_input_files(config["input_folder"], config["input_pattern"])
    print(f"[INFO] Ditemukan {len(files)} file input:")
    for f in files:
        print(f"       - {f.name}")

    usage_by_channel: dict[int, list] = {}
    for f in files:
        usages = extract_usage_from_file(f, config)
        print(f"[INFO] {f.name}: {len(usages)} baris channel terbaca")
        for u in usages:
            usage_by_channel.setdefault(u["channel"], []).append(u)

    return usage_by_channel


def style_header(ws, n_cols: int, color="305496"):
    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    font = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")


def write_report(channel_plan: dict, usage_by_channel: dict, config: dict) -> Path:
    output_folder = Path(config["output_folder"])
    output_folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = output_folder / f"Lambda_Mapping_{timestamp}.xlsx"

    wb = Workbook()

    # --- Sheet 1: Mapping semua channel ---
    ws = wb.active
    ws.title = "Lambda Mapping"
    info_cols = config.get("info_columns", [])
    headers = ["Channel", "Wavelength (nm)", "Frequency (THz)", "Status"] + [
        f"Used By - {c}" for c in info_cols
    ] + ["Source File"]
    ws.append(headers)

    used_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    idle_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    for n in sorted(channel_plan.keys()):
        plan = channel_plan[n]
        usages = usage_by_channel.get(n, [])
        status = "USED" if usages else "IDLE"

        if usages:
            first = usages[0]
            info_values = [first.get(c) for c in info_cols]
            source_files = ", ".join(sorted(set(u["source_file"] for u in usages)))
        else:
            info_values = [None for _ in info_cols]
            source_files = ""

        row = [n, plan["wavelength_nm"], plan["frequency_thz"], status] + info_values + [source_files]
        ws.append(row)

        row_idx = ws.max_row
        fill = used_fill if status == "USED" else idle_fill
        for col_idx in range(1, len(headers) + 1):
            ws.cell(row=row_idx, column=col_idx).fill = fill

    style_header(ws, len(headers))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(h) + 4)

    # --- Sheet 2: Ringkasan ---
    ws2 = wb.create_sheet("Ringkasan")
    total = len(channel_plan)
    used_count = sum(1 for n in channel_plan if n in usage_by_channel)
    idle_count = total - used_count
    idle_list = sorted(n for n in channel_plan if n not in usage_by_channel)

    ws2.append(["Keterangan", "Nilai"])
    ws2.append(["Total Channel", total])
    ws2.append(["Terpakai (USED)", used_count])
    ws2.append(["Kosong (IDLE)", idle_count])
    ws2.append(["Persentase Idle", f"{idle_count/total*100:.1f}%"])
    ws2.append([])
    ws2.append(["Daftar Channel Idle:"])
    for i in range(0, len(idle_list), 15):
        ws2.append([", ".join(str(c) for c in idle_list[i:i + 15])])

    style_header(ws2, 2)
    ws2.column_dimensions["A"].width = 45
    ws2.column_dimensions["B"].width = 20

    # --- Sheet 3: Flex Grid 100G Available ---
    ws3 = wb.create_sheet("Flex Grid 100G")
    flex_slots = find_flex_grid_100g(channel_plan, idle_list)

    flex_headers = [
        "Pasangan Channel 50G",
        "Channel A",
        "Channel B",
        "Center Frequency (THz)",
        "Center Wavelength (nm)",
        "Slot Range (THz)",
    ]
    ws3.append(flex_headers)

    flex_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    for slot in flex_slots:
        row = [
            slot["pair"],
            slot["channel_a"],
            slot["channel_b"],
            slot["center_freq_thz"],
            slot["center_wavelength_nm"],
            f"{slot['slot_low_thz']} - {slot['slot_high_thz']}",
        ]
        ws3.append(row)
        row_idx = ws3.max_row
        for col_idx in range(1, len(flex_headers) + 1):
            ws3.cell(row=row_idx, column=col_idx).fill = flex_fill

    style_header(ws3, len(flex_headers))
    ws3.freeze_panes = "A2"
    if flex_slots:
        ws3.auto_filter.ref = ws3.dimensions
    for i, h in enumerate(flex_headers, start=1):
        ws3.column_dimensions[get_column_letter(i)].width = max(16, len(h) + 4)

    if not flex_slots:
        ws3.append(["Tidak ada pasangan channel idle yang berurutan saat ini."])

    wb.save(out_path)
    return out_path, used_count, idle_count, idle_list, flex_slots

#anjay dibaca dong
def main():
    parser = argparse.ArgumentParser(description="Cari channel/lambda idle dari export WDM Trail")
    parser.add_argument("--config", default="config/config.yaml")
    args = parser.parse_args()

    print("=" * 55)
    print("  IDLE LAMBDA FINDER - MULAI")
    print("=" * 55)

    config = load_config(args.config)
    channel_plan = build_channel_plan(config)
    usage_by_channel = collect_all_usage(config)
    out_path, used_count, idle_count, idle_list, flex_slots = write_report(
        channel_plan, usage_by_channel, config
    )

    print("-" * 55)
    print(f"[INFO] Total channel      : {len(channel_plan)}")
    print(f"[INFO] Terpakai (USED)    : {used_count}")
    print(f"[INFO] Idle               : {idle_count}")
    print(f"[INFO] Channel idle       : {idle_list}")
    print(f"[INFO] Slot Flex Grid 100G tersedia: {len(flex_slots)}")
    for slot in flex_slots:
        print(
            f"         - Channel {slot['pair']} -> center {slot['center_freq_thz']} THz "
            f"/ {slot['center_wavelength_nm']} nm"
        )
    print(f"[SUKSES] Laporan tersimpan di: {out_path}")
    print("=" * 55)


if __name__ == "__main__":
    main()
