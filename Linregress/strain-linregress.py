import scipy, openpyxl, xlsxwriter, sys, glob

def getLinregressData(sheets, row, col, cnt):
  x = [] # Week, 0 = baseline
  y = [] # Average cell values
  for i, sheet in enumerate(sheets):
    top = sheet.cell(row, col).value
    if top != None:
      sum = float(top)
      for k in range(row + 1, row + cnt):
        sum += float(sheet.cell(k, col).value)
      x.append(i)
      y.append(sum / cnt)
  return x,y

# Sort files from weeks, and assume that is the correct time order
dataFiles = glob.glob(f"{sys.argv[1]}/*.xlsx")
dataFiles.sort()
wbs = [openpyxl.load_workbook(i, read_only=True) for i in dataFiles]
book = xlsxwriter.Workbook(sys.argv[2])
apex = book.add_worksheet("Apex")
apexs = [i["Apex"] for i in wbs]
pp = book.add_worksheet("PP")
pps = [i["PP"] for i in wbs]
smv = book.add_worksheet("SMV")
smvs = [i["SMV"] for i in wbs]
longAxis = book.add_worksheet("Long Axis")
longAxes = [i["Long Axis"] for i in wbs]

for i in zip([apex, pp, smv, longAxis], [apexs, pps, smvs, longAxes], [4, 6, 6, 7]):
  i[0].write_row("A3", ["ID", *["Slope", "P-value", "R^2-value"] * 9])
  for j,k in zip(range(1, 20, 9), ["Endo", "Myo", "Epi"], strict=True):
    i[0].merge_range(0, j, 0, j + 8, k)
    for l,m in zip(range(j, j + 9, 3), ["TTP", "Peak", "Es"], strict=True):
      i[0].merge_range(1, l, 1, l + 2, m)

  # j is for output sheet
  # k is for strain
  for j,k in zip(range(3, 8), range(3, 4 + 4 * i[2], i[2]), strict=True):
    data = []
    for l in range(3, 12):
      res = scipy.stats.linregress(*getLinregressData(i[1], k, l, i[2]))
      data.append(res.slope)
      data.append(res.pvalue)
      data.append(res.rvalue ** 2)
    i[0].write_row(j, 0, [i[1][0].cell(k, 1).value, *data])
  i[0].autofit()
book.close()
for i in wbs:
  i.close()