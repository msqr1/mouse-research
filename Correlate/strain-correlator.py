import csv
import scipy, openpyxl, xlsxwriter, sys, glob

def getStrainData(sheets, row, col, cnt):
  y = []   
  for sheet in sheets:
    top = sheet.cell(row, col).value
    if top != None:
      sum = float(top)
      for k in range(row + 1, row + cnt):
        sum += float(sheet.cell(k, col).value)
      y.append(sum / cnt)
  return y

def getMVData(input, top, left):
  data = []
  for i in range(left, left + 5):
    if input[top][i] != '':
      data.append(float(input[top][i]))
  return data

# Sort files from weeks, and assume that is the correct time order
dataFiles = glob.glob(f"{sys.argv[1]}/*.xlsx")
dataFiles.sort()
wbs = [openpyxl.load_workbook(i, read_only=True) for i in dataFiles]
book = xlsxwriter.Workbook(sys.argv[4])
apex = book.add_worksheet("Apex")
apexs = [i["Apex"] for i in wbs]
pp = book.add_worksheet("PP")
pps = [i["PP"] for i in wbs]
smv = book.add_worksheet("SMV")
smvs = [i["SMV"] for i in wbs]
longAxis = book.add_worksheet("Long Axis")
longAxes = [i["Long Axis"] for i in wbs]

mvData = []
with open(sys.argv[2]) as f:
  reader = csv.reader(f)
  for i in reader:
    mvData.append(i)
for i in zip([apex, pp], [apexs, pps], [4, 6]):
  i[0].write_row("A3", ["ID", "", *["Coefficient", "P-value"] * 9])
  for j,k in zip(range(2, 15, 6), ["Endo", "Myo", "Epi"], strict=True):
    i[0].merge_range(0, j, 0, j + 5, k)
    for l,m in zip(range(j, j + 6, 2), ["TTP", "Peak", "Es"], strict=True):
      i[0].merge_range(1, l, 1, l + 1, m)

  # j is for mv and output sheet
  # k is for strain
  for j,k in zip(range(1, 22, 4), range(3, 4 + 5 * i[2], i[2]), strict=True):
    funcData = [getMVData(mvData, l, 4) for l in range(j, j + 4)]
    strainData = [getStrainData(i[1], k, l, i[2]) for l in range(3, 12)]
    for l in enumerate(funcData):
      data = []
      for m in strainData:
        res = scipy.stats.pearsonr(l[1], m)
        data.append(res.correlation)
        data.append(res.pvalue)
      i[0].write_row(j + 2 + l[0], 2, data)
    i[0].merge_range(j + 2, 0, j + 5, 0, mvData[j][0])
    i[0].write_column(j + 2, 1, ["AET", "IVCT", "IVRT", "MV E"])
  i[0].autofit()

mvData = []
with open(sys.argv[3]) as f:
  reader = csv.reader(f)
  for i in reader:
    mvData.append(i)
for i in zip([smv, longAxis], [smvs, longAxes], [6, 7]):
  i[0].write_row("A3", ["ID", "", *["Coefficient", "P-value"] * 9])
  for j,k in zip(range(2, 15, 6), ["Endo", "Myo", "Epi"], strict=True):
    i[0].merge_range(0, j, 0, j + 5, k)
    for l,m in zip(range(j, j + 6, 2), ["TTP", "Peak", "Es"], strict=True):
      i[0].merge_range(1, l, 1, l + 1, m)

  # j is for mv and output sheet
  # k is for strain
  for j,k in zip(range(1, 22, 4), range(3, 4 + 5 * i[2], i[2]), strict=True):
    funcData = [getMVData(mvData, l, 4) for l in range(j, j + 4)]
    strainData = [getStrainData(i[1], k, l, i[2]) for l in range(3, 12)]
    for l in enumerate(funcData):
      data = []
      for m in strainData:
        res = scipy.stats.pearsonr(l[1], m)
        data.append(res.correlation)
        data.append(res.pvalue)
      i[0].write_row(j + 2 + l[0], 2, data)
    i[0].merge_range(j + 2, 0, j + 5, 0, mvData[j][0])
    i[0].write_column(j + 2, 1, ["AET", "IVCT", "IVRT", "MV E"])
  i[0].autofit()

book.close()
for i in wbs:
  i.close()