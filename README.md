# Koha Patron Image Mapper (Excel → Cardnumber Image Extractor)

This tool extracts embedded images from an Excel file (`.xlsx`) and renames them using a **cardnumber / registration number**. It is designed for **Koha ILS bulk patron image upload preparation**.

---

## 🎯 Purpose

Koha requires patron images to be named using:

```

cardnumber.jpg

````

This tool automates:

- Extracting embedded images from Excel
- Reading cardnumber from a specified column
- Mapping images to correct row
- Renaming images for Koha bulk upload

---

## 📁 Excel Format Assumption

![Screenshot](/excel.png)

- Each row represents one patron
- Images must be embedded inside Excel cells

---

## ⚙️ Full Setup (Linux / Ubuntu / WSL)

### 1. Create Project Folder

```bash
mkdir -p ~/excel_image_extract
cd ~/excel_image_extract
````

---

### 2. Install Dependencies

```bash
sudo apt update
sudo apt install -y python3 python3-venv python3-pip zip unzip
```

---

### 3. Create Virtual Environment

```bash
python3 -m venv venv
source venv/bin/activate
```

---

### 4. Install Python Packages

```bash
pip install openpyxl lxml
```

---

## 📥 Add Excel File

Copy your Excel file into the project folder:

```bash
cp /home/username/newinput.xlsx ~/excel_image_extract/
```

---

## 🧠 How It Works

1. Reads Excel file using `openpyxl`
2. Extracts cardnumber from Column C
3. Opens `.xlsx` as ZIP archive
4. Reads images from `xl/media/`
5. Reads drawing XML mapping (`xl/drawings/`)
6. Matches image → row → cardnumber
7. Saves image as:

```
final_photos/<cardnumber>.jpg
```

---

## 🧾 Python Script (final_map.py)

Create script:

```bash
nano final_map.py
```

Paste the following code:

```python
from openpyxl import load_workbook
import zipfile
from lxml import etree
import os

excel_file = "newinput.xlsx"
output_folder = "final_photos"

os.makedirs(output_folder, exist_ok=True)

# -------------------------
# STEP 1: Map Row → Cardnumber
# -------------------------
wb = load_workbook(excel_file)
ws = wb.active

reg_map = {}
for row in range(2, ws.max_row + 1):
    reg = ws.cell(row=row, column=3).value  # Column C Registration/Admision/Cardnumber
    if reg:
        reg_map[row] = str(reg)

# -------------------------
# STEP 2: Read Excel as ZIP
# -------------------------
with zipfile.ZipFile(excel_file, 'r') as z:

    # Extract images
    media = {
        os.path.basename(f): z.read(f)
        for f in z.namelist()
        if "xl/media/" in f
    }

    # Find drawing files
    drawings = [f for f in z.namelist() if "xl/drawings/drawing" in f]

    for d in drawings:
        xml = z.read(d)
        tree = etree.fromstring(xml)

        ns = {
            "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        }

        rel_file = d.replace("drawings/", "drawings/_rels/") + ".rels"

        if rel_file not in z.namelist():
            continue

        rel_xml = z.read(rel_file)
        rel_tree = etree.fromstring(rel_xml)

        rel_map = {}
        for rel in rel_tree:
            rel_map[rel.attrib["Id"]] = os.path.basename(rel.attrib["Target"])

        for anchor in tree.findall(".//xdr:twoCellAnchor", ns):

            from_node = anchor.find("xdr:from", ns)
            pic = anchor.find(".//xdr:pic//a:blip", ns)

            if from_node is None or pic is None:
                continue

            row = int(from_node.find("xdr:row", ns).text) + 1
            rId = pic.attrib.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
            )

            if row in reg_map and rId in rel_map:

                img_name = rel_map[rId]
                reg_no = reg_map[row]

                if img_name in media:

                    out_file = os.path.join(output_folder, f"{reg_no}.jpg")

                    with open(out_file, "wb") as f:
                        f.write(media[img_name])

                    print("Saved:", out_file)

print("DONE - ALL IMAGES MAPPED")
```

---

## 🚀 Run the Script

```bash
python3 final_map.py
```

---

## 📂 Output

All images will be saved here:

```
final_photos/
```

Example output:

```
12345.jpg
67890.jpg
112233.jpg
```

---

## 📦 Create Zip Package

```bash
zip -r koha_patron_images.zip final_photos
```
---

## ⚠️ Notes

* Column C must contain unique cardnumbers
* Only embedded Excel images supported
* Each row must correspond to correct image placement
* If multiple images exist per row, first valid mapping is used

---

## 🛠️ Requirements

* Python 3.8+
* openpyxl
* lxml
* Linux / WSL recommended

