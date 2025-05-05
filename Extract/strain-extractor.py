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
    IDs.append(animalIDs[i])
    XMLs.append(a[0] if len(a) > 0 else '')
  return zip(range(2, len(IDs) * cellCnt, cellCnt), XMLs, IDs, strict=True)

book = xlsxwriter.Workbook(output)

class Input():
  def __init__(self, sheet, cellTp, globExpr):
    self.sheet = sheet
    self.cellTp = cellTp
    self.cellCnt = len(cellTp)
    self.globExpr = globExpr

apex = book.add_worksheet("Apex")
pp = book.add_worksheet("PP")
smv = book.add_worksheet("SMV")
longAxis = book.add_worksheet("Long Axis")

for i in [apex, pp, smv, longAxis]:
  i.write_row("A2", ["ID", "Cell type"])
  i.merge_range("C1:E1", "Endo")
  i.merge_range("F1:H1", "Myo")
  i.merge_range("I1:K1", "Epi")
  i.write_row("C2", ["TTP", "Peak", "ES"] * 3)

for input in [
  Input(apex, ["13-apical anterior", "16-apical lateral", "15-apical inferior", "14-apical septal"], "Apex*/*TTP(apex)"),
  Input(pp, ["07-mid anterior", "12-mid anterolateral", "11-mid inferolateral", "10-mid inferior", "09-mid inferoseptal", "08-mid anteroseptal"], "PP*/*TTP(pm)"),
  Input(smv, ["01-basal anterior", "06-basal anterolateral", "05-basal inferolateral", "04-basal inferior", "03-basal inferoseptal", "02-basal anteroseptal"], "sMV*/*TTP(mv)"),
  Input(longAxis, ["05-basal inferolateral", "11-mid inferolateral", "16-apical lateral", "17-apex", "14-apical septal", "08-mid anteroseptal", "02-basal anteroseptal"], "Long Axis*/*TTP(pslax)")
]:
  maxRowLen = 12 + input.cellCnt * 2
  skipRadial = 10 + input.cellCnt
  for i, xml, id  in getStrainIterables(input.globExpr, input.cellCnt):
    input.sheet.merge_range(i, 0, i + input.cellCnt - 1, 0, id)
    input.sheet.write_column(i, 1, input.cellTp)
    if xml == '':
      continue
    worksheets = None
    with open(xml) as f:
      worksheets = xmltodict.parse(f.read())["Workbook"]["Worksheet"]
    data = []
    for j in range(input.cellCnt):
      data.append([])
    for ws in worksheets:
      name = ws["@ss:Name"]
      if name == "Strain-Endo TTP" or name == "Strain-Myo TTP" or name == "Strain-Epi TTP":
        rows = ws["Table"]["Row"]
        start = skipRadial if len(rows) == maxRowLen else 7
        for j in range(input.cellCnt):
          cells = rows[start + j]["Cell"]
          data[j].append(cells[1]["Data"]["#text"])
          data[j].append(cells[2]["Data"]["#text"])
          data[j].append(cells[3]["Data"]["#text"])
    for j in range(input.cellCnt):
      input.sheet.write_row(i + j, 2, data[j])
for i in [apex, pp, smv, longAxis]:
  i.autofit()

book.close()