from flask import Flask, render_template, request, jsonify, send_file
import json
import io
import os
import threading
import sys
import webview  # WebView kütüphanesi eklendi
from datetime import datetime

app = Flask(__name__)
# ─── Optimizasyon: Guillotine kesim algoritması ───────────────────────────────
def optimize_cuts(pieces, sheet_w, sheet_h, blade=4):
    """
    Basit guillotine bin-packing.
    pieces: [(w, h, label, qty), ...]
    Döner: sheets_used (kaç levha), placed (hangi parça nerede), waste_pct
    """
    # Parçaları büyükten küçüğe sırala
    expanded = []
    for w, h, label, qty, bant in pieces:
        for i in range(qty):
            expanded.append((w, h, label, bant))
    expanded.sort(key=lambda p: p[0]*p[1], reverse=True)

    sheets = []  # her sheet: list of free rectangles
    placements = []  # (sheet_idx, x, y, w, h, label)

    def fits(rect, pw, ph):
        return rect[2] >= pw and rect[3] >= ph

    def split(rect, pw, ph):
        rx, ry, rw, rh = rect
        # Sağ kalan ve alt kalan
        right = (rx + pw + blade, ry, rw - pw - blade, ph)
        bottom = (rx, ry + ph + blade, rw, rh - ph - blade)
        result = []
        if right[2] > 0 and right[3] > 0:
            result.append(right)
        if bottom[2] > 0 and bottom[3] > 0:
            result.append(bottom)
        return result

    for pw, ph, label, bant in expanded:
        placed = False
        for si, free_rects in enumerate(sheets):
            for ri, rect in enumerate(free_rects):
                # Normal yön
                if fits(rect, pw, ph):
                    placements.append((si, rect[0], rect[1], pw, ph, label, bant))
                    new_rects = split(rect, pw, ph)
                    free_rects.pop(ri)
                    free_rects.extend(new_rects)
                    placed = True
                    break
                # 90° döndür
                elif fits(rect, ph, pw):
                    placements.append((si, rect[0], rect[1], ph, pw, label + " (D)", bant))
                    new_rects = split(rect, ph, pw)
                    free_rects.pop(ri)
                    free_rects.extend(new_rects)
                    placed = True
                    break
            if placed:
                break
        if not placed:
            si = len(sheets)
            sheets.append([(0, 0, sheet_w, sheet_h)])
            rect = sheets[si][0]
            if fits(rect, pw, ph):
                placements.append((si, rect[0], rect[1], pw, ph, label, bant))
                new_rects = split(rect, pw, ph)
                sheets[si].pop(0)
                sheets[si].extend(new_rects)
            else:
                placements.append((si, rect[0], rect[1], ph, pw, label + " (D)", bant))
                new_rects = split(rect, ph, pw)
                sheets[si].pop(0)
                sheets[si].extend(new_rects)

    total_area = sheet_w * sheet_h * len(sheets)
    used_area = sum(p[3]*p[4] for p in placements)
    waste_pct = round((1 - used_area/total_area)*100, 1) if total_area else 0

    return len(sheets), placements, waste_pct


# ─── Maliyet Hesaplama ────────────────────────────────────────────────────────
def calculate_cost(data):
    pieces = data.get("pieces", [])
    pricing = data.get("pricing", {})
    sheet_w = float(data.get("sheet_w", 2440))
    sheet_h = float(data.get("sheet_h", 1220))
    material = data.get("material", "MDF")

    # Levha fiyatı (adet)
    sheet_price = float(pricing.get("sheet_price", 0))
    # Kesim ücreti (kesim başına)
    cut_price = float(pricing.get("cut_price", 0))
    # PVC bant fiyatı (metre başına)
    band_price = float(pricing.get("band_price", 0))
    # Nakliye
    shipping = float(pricing.get("shipping", 0))

    # Parça listesini hazırla — giriş CM, içeride MM'e çevir
    piece_list = []
    total_band_m = 0.0

    for p in pieces:
        w_cm = float(p.get("w", 0))
        h_cm = float(p.get("h", 0))
        w = w_cm * 10   # cm → mm
        h = h_cm * 10   # cm → mm
        qty = int(p.get("qty", 1))
        bant_sides = p.get("bant", [])

        piece_list.append((w, h, p.get("name", "Parça"), qty, bant_sides))

        # Bant hesabı — mm → metre (/1000)
        for side in bant_sides:
            if side in ["sol", "sag"]:
                total_band_m += (h / 1000) * qty
            elif side in ["ust", "alt"]:
                total_band_m += (w / 1000) * qty

    # Optimizasyon (testere kalınlığı kullanıcıdan)
    blade = float(data.get("blade", 3))
    num_sheets, placements, waste_pct = optimize_cuts(piece_list, sheet_w, sheet_h, blade=blade)

    # Kesim sayısı = levha başına kesim (guillotine bölme sayısı)
    # Her levhadaki parça sayısına göre hesapla
    cuts_per_sheet = {}
    for p in placements:
        si = p[0]
        cuts_per_sheet[si] = cuts_per_sheet.get(si, 0) + 1
    # Her levhada n parça varsa yaklaşık n+1 kesim (satır+sütun)
    total_cuts = sum(max(1, v) for v in cuts_per_sheet.values())

    # Maliyet
    cost_sheets = num_sheets * sheet_price
    cost_cuts = num_sheets * cut_price   # Kesim ücreti plaka başına
    cost_band = total_band_m * band_price
    cost_total = cost_sheets + cost_cuts + cost_band + shipping

    return {
        "num_sheets": num_sheets,
        "waste_pct": waste_pct,
        "total_cuts": total_cuts,
        "total_band_m": round(total_band_m, 2),
        "cost_sheets": round(cost_sheets, 2),
        "cost_cuts": round(cost_cuts, 2),
        "cost_band": round(cost_band, 2),
        "cost_shipping": round(shipping, 2),
        "cost_total": round(cost_total, 2),
        "placements": [
            {
                "sheet": p[0]+1,
                "x": p[1], "y": p[2],
                "w": p[3], "h": p[4],
                "label": p[5],
                "bant": p[6],
                # cm değerler etiket için
                "w_cm": round(p[3]/10, 1),
                "h_cm": round(p[4]/10, 1),
            }
            for p in placements
        ],
        "sheet_w": sheet_w,
        "sheet_h": sheet_h,
    }


# ─── Routes ──────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/calculate", methods=["POST"])
def api_calculate():
    data = request.json
    result = calculate_cost(data)
    return jsonify(result)

@app.route("/api/export_excel", methods=["POST"])
def api_export_excel():
    try:
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        return jsonify({"error": "openpyxl yuklu degil. pip install openpyxl"}), 500

    try:
        data = request.json
        result = calculate_cost(data)
        pieces = data.get("pieces", [])
        pricing = data.get("pricing", {})

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Kesim Listesi"

        purp   = "4A1A8A"
        lpurp  = "D4C4F0"
        header_fill = PatternFill("solid", fgColor=purp)
        sub_fill    = PatternFill("solid", fgColor="7B5EA7")
        alt_fill    = PatternFill("solid", fgColor=lpurp)
        white_fill  = PatternFill("solid", fgColor="FFFFFF")

        bold_white = Font(bold=True, color="FFFFFF", size=11)
        bold_light = Font(bold=True, color="EDE0FF", size=10)
        normal     = Font(color="1A0533", size=10)

        thin   = Side(style="thin", color="A66EE8")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        center     = Alignment(horizontal="center", vertical="center")
        left_align = Alignment(horizontal="left",   vertical="center")

        # Baslik
        ws.merge_cells("A1:H1")
        ws["A1"] = f"Atölyemhanem — KESIM LISTESI — {datetime.now().strftime('%d.%m.%Y')}"
        ws["A1"].font      = Font(bold=True, color="FFFFFF", size=14)
        ws["A1"].fill      = header_fill
        ws["A1"].alignment = center

        # Levha bilgisi
        ws.merge_cells("A2:H2")
        ws["A2"] = f"Levha: {data.get('material','MDF')}  |  {data.get('sheet_w')}x{data.get('sheet_h')} mm  |  Levha Adedi: {result['num_sheets']}  |  Fire: %{result['waste_pct']}"
        ws["A2"].font      = bold_light
        ws["A2"].fill      = sub_fill
        ws["A2"].alignment = center

        # Kesim listesi baslik
        headers = ["#", "Parca Adi", "Boy (cm)", "En (cm)", "Adet", "Bant Kenarlari", "Bant (m)", "Levha #"]
        ws.append([])
        row = ws.max_row
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=i, value=h)
            cell.font      = bold_white
            cell.fill      = PatternFill("solid", fgColor="2D0D5C")
            cell.alignment = center
            cell.border    = border

        # Parca satirlari
        placements_by_label = {}
        for p in result["placements"]:
            lbl = p["label"].replace(" (D)","")
            if lbl not in placements_by_label:
                placements_by_label[lbl] = []
            placements_by_label[lbl].append(str(p["sheet"]))

        bant_map = {"ust":"UK-1","alt":"UK-2","sol":"KK-1","sag":"KK-2"}
        for idx, p in enumerate(pieces):
            w    = float(p.get("w", 0))
            h    = float(p.get("h", 0))
            qty  = int(p.get("qty", 1))
            bant_sides = p.get("bant", [])
            bant_str = ", ".join(bant_map.get(s, s) for s in bant_sides) or "—"
            name = p.get("name", f"Parca {idx+1}")
            band_m = 0
            for side in bant_sides:
                if side in ["sol", "sag"]:
                    band_m += (h * 10 / 1000) * qty
                else:
                    band_m += (w * 10 / 1000) * qty
            sheets_used = ", ".join(sorted(set(placements_by_label.get(name, []))))

            row_data = [idx+1, name, h, w, qty, bant_str, round(band_m,2), sheets_used]
            ws.append(row_data)
            rn   = ws.max_row
            fill = alt_fill if idx % 2 == 0 else white_fill
            for col in range(1, 9):
                c = ws.cell(row=rn, column=col)
                c.fill      = fill
                c.font      = normal
                c.alignment = center if col != 2 else left_align
                c.border    = border

        # Maliyet tablosu
        ws.append([])
        ws.append([])
        cost_title_row = ws.max_row
        ws.merge_cells(f"A{cost_title_row}:H{cost_title_row}")
        t = ws.cell(row=cost_title_row, column=1, value="MALIYET OZETI")
        t.font      = bold_white
        t.fill      = header_fill
        t.alignment = center

        cost_rows = [
            ("Levha Maliyeti", result["cost_sheets"],  f"{result['num_sheets']} levha x {pricing.get('sheet_price',0)} TL"),
            ("Kesim Ucreti",   result["cost_cuts"],    f"{result['num_sheets']} plaka x {pricing.get('cut_price',0)} TL"),
            ("PVC Bant",       result["cost_band"],    f"{result['total_band_m']} m x {pricing.get('band_price',0)} TL"),
            ("Nakliye",        result["cost_shipping"], ""),
            ("TOPLAM",         result["cost_total"],   ""),
        ]

        for i, (label, amount, note) in enumerate(cost_rows):
            ws.append(["", label, "", "", f"{amount} TL", "", note, ""])
            rn = ws.max_row
            is_total = label == "TOPLAM"
            ws.merge_cells(f"B{rn}:D{rn}")
            ws.merge_cells(f"F{rn}:H{rn}")
            fill = PatternFill("solid", fgColor="1A0533") if is_total else (alt_fill if i%2==0 else white_fill)
            font = Font(bold=True, color="FFFFFF", size=12) if is_total else normal
            for col in [2,5,7]:
                c = ws.cell(row=rn, column=col)
                c.fill      = fill
                c.font      = font
                c.alignment = center
                c.border    = border

        col_widths = [5, 28, 10, 10, 7, 22, 10, 12]
        for i, w2 in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w2

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        fname = f"kesim_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        return send_file(output,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name=fname)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/import_excel", methods=["POST"])
def api_import_excel():
    try:
        import openpyxl
    except ImportError:
        return jsonify({"error": "openpyxl yuklu degil"}), 500
    try:
        if 'file' not in request.files:
            return jsonify({"error": "Dosya bulunamadı"}), 400
        f = request.files['file']
        wb = openpyxl.load_workbook(f, data_only=True)
        ws = wb.active
        rows = list(ws.values)

        # Başlık satırını bul
        header_row = -1
        col_map = {}
        for i, row in enumerate(rows[:20]):
            row_upper = [str(c).strip().upper() if c is not None else '' for c in row]
            keys = {'TANIM', 'BOY', 'EN', 'ADET'}
            if any(k in row_upper for k in keys):
                header_row = i
                for j, c in enumerate(row_upper):
                    if c == 'TANIM':          col_map['tanim'] = j
                    elif c == 'BOY':          col_map['boy'] = j
                    elif c == 'EN':           col_map['en'] = j
                    elif c in ('ADET','AD.'): col_map['adet'] = j
                    elif c == 'UK1':          col_map['uk1'] = j
                    elif c == 'UK2':          col_map['uk2'] = j
                    elif c == 'KK1':          col_map['kk1'] = j
                    elif c == 'KK2':          col_map['kk2'] = j
                # TANIM bulunamadıysa, skip sütunlardan ilk metin sütununu al
                if 'tanim' not in col_map:
                    skip = {'NO','BOY','EN','ADET','AD.','UK1','UK2','KK1','KK2','ÖLÇÜ','BANT','MM',''}
                    for j, c in enumerate(row_upper):
                        if j == 0: continue
                        if c not in skip and c:
                            col_map['tanim'] = j
                            break
                break

        if header_row == -1:
            return jsonify({"error": "Başlık satırı bulunamadı (TANIM, BOY, EN, ADET)"}), 400

        pieces = []
        for row in rows[header_row + 1:]:
            def get(idx):
                if idx is None or idx >= len(row): return ''
                v = row[idx]
                return str(v).strip() if v is not None else ''

            tanim = get(col_map.get('tanim'))
            boy_s = get(col_map.get('boy')).replace(',', '.')
            en_s  = get(col_map.get('en')).replace(',', '.')
            try: boy = float(boy_s)
            except: boy = 0
            try: en = float(en_s)
            except: en = 0
            try: adet = int(float(get(col_map.get('adet')) or 1))
            except: adet = 1

            if not tanim and boy == 0 and en == 0: continue
            if boy == 0 or en == 0: continue

            def chk(key):
                v = get(col_map.get(key))
                return bool(v and v != '0')

            pieces.append({
                'name': tanim or f'Parça {len(pieces)+1}',
                'h': boy, 'w': en, 'qty': adet,
                'uk1': chk('uk1'), 'uk2': chk('uk2'),
                'kk1': chk('kk1'), 'kk2': chk('kk2')
            })

        if not pieces:
            return jsonify({"error": "Geçerli parça bulunamadı"}), 400

        return jsonify({"pieces": pieces, "count": len(pieces)})

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# ─── WebView Başlatıcı ───────────────────────────────────────────────────────
def run_flask():
    # Flask sunucusunu sessizce arka planda çalıştırır
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)

if __name__ == "__main__":
    # 1. Flask'ı ayrı bir thread üzerinde başlat
    flask_thread = threading.Thread(target=run_flask)
    flask_thread.daemon = True
    flask_thread.start()

    # 2. WebView Penceresini oluştur
    webview.create_window(
        'Atölyemhanem v2 — Kesim Programı', 
        'http://127.0.0.1:5000',
        width=1280,
        height=900,
        resizable=True,
        min_size=(300, 300)
    )
    
    # 3. Uygulamayı başlat
    webview.start()
    sys.exit()