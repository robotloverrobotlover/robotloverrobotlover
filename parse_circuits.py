import openpyxl

def get_cell_color(cell):
    fill = cell.fill
    if fill and fill.fill_type not in (None, "none"):
        try:
            rgb = fill.fgColor.rgb
        except Exception:
            return None
        if rgb and rgb not in ("00000000", "FF000000", "FFFFFFFF"):
            color_map = {
                "FF2E7D32": "green",  "FF00695C": "teal",   "FF1565C0": "blue",
                "FF00AA00": "green",  "FF007070": "teal",   "FF005500": "dark green",
                "FF008000": "green",  "FF006400": "dark green",
                "FF4E7C2F": "mid green", "FF1F5C1F": "dark green",
                "FF6B8E23": "olive",  "FF90C060": "light green",
            }
            r = rgb.upper()
            return color_map.get(r, f"#{r[-6:]}")
    return None

def parse_excel(filepath):
    wb = openpyxl.load_workbook(filepath)
    sheet = None
    for name in wb.sheetnames:
        if "mpls" in name.lower() or "pre" in name.lower():
            sheet = wb[name]; break
    if sheet is None:
        sheet = wb.active

    merged_map = {}
    for rng in sheet.merged_cells.ranges:
        tl = (rng.min_row, rng.min_col)
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                merged_map[(r, c)] = tl

    def real_value(row_1, col_1):
        key = (row_1, col_1)
        if key in merged_map:
            tr, tc = merged_map[key]
            return sheet.cell(tr, tc).value
        return sheet.cell(row_1, col_1).value

    header_row = col_fp = col_comm = col_circ = col_new = col_name = col_email = None
    site_headers = []

    for row in sheet.iter_rows():
        for cell in row:
            if "full/partial" in str(cell.value or "").lower():
                header_row = cell.row
                break
        if header_row:
            break

    if not header_row:
        raise ValueError("Header row not found")

    for cell in sheet[header_row]:
        v = str(cell.value or "").strip().lower()
        c = cell.column
        if "full/partial" in v:    col_fp    = c
        if "comments"     in v:    col_comm  = c
        if "circuit #"    in v and "new" not in v: col_circ  = c
        if "new circuit"  in v:    col_new   = c
        if "circuit name" in v:    col_name  = c
        if "email" in v or "cau" in v: col_email = c
        if str(cell.value or "").strip().upper().startswith("SITE"):
            site_headers.append((c, str(cell.value).strip().upper()))

    first_data = header_row + 1
    for r in range(header_row + 1, sheet.max_row + 1):
        key = (r, col_circ)
        if key in merged_map and merged_map[key][0] <= header_row:
            first_data = r + 1
        else:
            first_data = r
            break

    circuits = []
    visited = set()
    r = first_data

    while r <= sheet.max_row:
        if r in visited:
            r += 1; continue

        circ_val = str(real_value(r, col_circ) or "").strip()
        name_val = str(real_value(r, col_name) or "").strip()

        if not circ_val and not name_val:
            r += 1; continue

        key = (r, col_circ)
        if key in merged_map and merged_map[key][0] < first_data:
            r += 1; continue

        span_end = r
        for rr in range(r + 1, sheet.max_row + 1):
            k = (rr, col_circ)
            if k in merged_map and merged_map[k][0] == r:
                span_end = rr
            else:
                break

        def cv(col):
            v = real_value(r, col)
            return str(v or "").strip()

        fp    = cv(col_fp)
        comm  = cv(col_comm)  if col_comm  else ""
        circ  = cv(col_circ)
        new_c = cv(col_new)   if col_new   else ""
        name  = cv(col_name)
        email = cv(col_email) if col_email else ""

        connections = []
        for rr in range(r, span_end + 1):
            row_sites = []
            for col_1, site_name in site_headers:
                cell = sheet.cell(rr, col_1)
                v = str(cell.value or "").strip()
                if v:
                    color = get_cell_color(cell)
                    row_sites.append((site_name, v, color))
            if row_sites:
                connections.append(row_sites)

        visited.update(range(r, span_end + 1))

        circuits.append({
            "full_partial": fp, "comments": comm,
            "circuit": circ,    "new_circuit": new_c,
            "name": name,       "email": email,
            "connections": connections,
        })

        r = span_end + 1

    return circuits

def format_output(circuits):
    lines = []
    for i, c in enumerate(circuits, 1):
        lines.append(f"Block {i}")
        lines.append(f"1. Circuit #:     {c['circuit'] or '(blank)'}")
        lines.append(f"2. New Circuit #: {c['new_circuit'] or '(blank)'}")
        lines.append(f"3. Circuit Name:  {c['name'] or '(blank)'}")

        connections = c["connections"]
        if not connections:
            lines.append("   (No site connections found)")
        else:
            for conn_idx, row_sites in enumerate(connections, 1):
                lines.append(f"   Connection {conn_idx}:")
                for site_name, port_val, color in row_sites:
                    color_str = f", color: {color}" if color else ""
                    lines.append(f"      {site_name} (port: {port_val}{color_str})")
                if conn_idx < len(connections):
                    lines.append("")

        extras = []
        if c["full_partial"]: extras.append(f"Full/Partial: {c['full_partial']}")
        if c["comments"]:     extras.append(f"Comments: {c['comments']}")
        if c["email"]:        extras.append(f"Email/CAU ID: {c['email']}")
        if extras:
            lines.append("   [" + " | ".join(extras) + "]")
        lines.append("")
    return "\n".join(lines)

circuits = parse_excel("/home/claude/Altemp2_template.xlsx")
text = format_output(circuits)
print(text)
with open("/home/claude/circuits_output.txt", "w") as f:
    f.write(text)
