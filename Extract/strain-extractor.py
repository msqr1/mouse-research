import sys, glob, xlsxwriter, xmltodict

weekDir = sys.argv[1]
output = sys.argv[2]
animalDirs = glob.glob(f"{weekDir}/*")
animalIDs = [i[len(weekDir) + 1:] for i in animalDirs]
animalCnt = len(animalDirs)

def getStrainIterables(globExpr, cellCnt):
  XMLs = []
  IDs = []
  for i in range(animalCnt):
    a = glob.glob(f"{animalDirs[i]}/{globExpr}.xml")
    if len(a) > 0:
      XMLs.append(a[0])
      IDs.append(animalIDs[i])
  return zip(range(2, len(IDs) * cellCnt, cellCnt), XMLs, IDs, strict=True)

book = xlsxwriter.Workbook(output)
apex = book.add_worksheet("Apex")
pp = book.add_worksheet("PP")
smv = book.add_worksheet("SMV")

for i in [apex, pp, smv]:
  i.write_row("A1", ["ID", "Cell type"])
  i.merge_range("C1:E1", "Endo")
  i.merge_range("F1:H1", "Myo")
  i.merge_range("I1:K1", "Epi")
  i.write_row("C2", ["TTP", "Peak", "ES"] * 3)

cellTp = ["13-apical anterior", "16-apical lateral", "15-apical inferior", "14-apical septal"]
cellCnt = len(cellTp)
for i, xml, id  in getStrainIterables("Apex*/*TTP(apex)", cellCnt):
  worksheets = None
  with open(xml) as f:
    worksheets = xmltodict.parse(f.read())["Workbook"]["Worksheet"]
  data = []
  for j in range(cellCnt):
    data.append([])
  for ws in worksheets:
    name = ws["@ss:Name"]
    if name == "Strain-Endo TTP" or name == "Strain-Myo TTP" or name == "Strain-Epi TTP":
      rows = ws["Table"]["Row"]
      start = 14 if len(rows) == 20 else 7
      for j in range(cellCnt):
        cells = rows[start + j]["Cell"]
        data[j].append(cells[1]["Data"]["#text"])
        data[j].append(cells[2]["Data"]["#text"])
        data[j].append(cells[3]["Data"]["#text"])
  apex.merge_range(i, 0, i + cellCnt - 1, 0, id)
  apex.write_column(i, 1, cellTp)
  for j in range(cellCnt):
    apex.write_row(i + j, 2, data[j])

cellTp = ["07-mid anterior", "12-mid anterolateral", "11-mid inferolateral", "10-mid inferior", "09-mid inferoseptal", "08-mid anteroseptal"]
cellCnt = len(cellTp)
for i, xml, id  in getStrainIterables("PP*/*TTP(pm)", cellCnt):
  worksheets = None
  with open(xml) as f:
    worksheets = xmltodict.parse(f.read())["Workbook"]["Worksheet"]
  data = []
  for j in range(cellCnt):
    data.append([])
  for ws in worksheets:
    name = ws["@ss:Name"]
    if name == "Strain-Endo TTP" or name == "Strain-Myo TTP" or name == "Strain-Epi TTP":
      rows = ws["Table"]["Row"]
      start = 16 if len(rows) == 24 else 7
      for j in range(cellCnt):
        cells = rows[start + j]["Cell"]
        data[j].append(cells[1]["Data"]["#text"])
        data[j].append(cells[2]["Data"]["#text"])
        data[j].append(cells[3]["Data"]["#text"])
  pp.merge_range(i, 0, i + cellCnt - 1, 0, id)
  pp.write_column(i, 1, cellTp)
  for j in range(cellCnt):
    pp.write_row(i + j, 2, data[j])

cellTp = ["01-basal anterior", "06-basal anterolateral", "05-basal inferolateral", "04-basal inferior", "03-basal inferoseptal", "02-basal anteroseptal"]
cellCnt = len(cellTp)
for i, xml, id  in getStrainIterables("sMV*/*TTP(mv)", cellCnt):
  worksheets = None
  with open(xml) as f:
    worksheets = xmltodict.parse(f.read())["Workbook"]["Worksheet"]
  data = []
  for j in range(cellCnt):
    data.append([])
  for ws in worksheets:
    name = ws["@ss:Name"]
    if name == "Strain-Endo TTP" or name == "Strain-Myo TTP" or name == "Strain-Epi TTP":
      rows = ws["Table"]["Row"]
      start = 16 if len(rows) == 24 else 7
      for j in range(cellCnt):
        cells = rows[start + j]["Cell"]
        data[j].append(cells[1]["Data"]["#text"])
        data[j].append(cells[2]["Data"]["#text"])
        data[j].append(cells[3]["Data"]["#text"])
  smv.merge_range(i, 0, i + cellCnt - 1, 0, id)
  smv.write_column(i, 1, cellTp)
  for j in range(cellCnt):
    smv.write_row(i + j, 2, data[j])

for i in [apex, pp, smv]:
  i.autofit()
  
book.close()