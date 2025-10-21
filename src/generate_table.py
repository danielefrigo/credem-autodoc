def generate_table(
  doc,
  header = [],
  body = [[]]
):
    num_cols = len(header)
    num_rows = len(body)
    table = doc.add_table(rows=num_rows+1, cols=num_cols  )
    table.style = "Credem"

    hdr_cells = table.rows[0].cells
    for h in range(num_cols):
        hdr_cells[h].text = header[h]

    for r in range(num_rows):
        body_row = table.rows[r+1].cells
        for c in range(num_cols):
            body_row[c].text = body[r][c]

