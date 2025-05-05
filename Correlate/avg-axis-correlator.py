import scipy.stats, csv, sys, xlsxwriter

def getMVData(input, top, left):
  data = []
  for i in range(left, left + 5):
    if input[top][i] != '':
      data.append(float(input[top][i]))
  return data

def getAxisData(input, top, left):
  data = []
  for i in range(left, left + 17, 4):
    if input[top][i] != '':
      data.append(float(input[top][i]))
  return data

with open(sys.argv[1]) as mv, open(sys.argv[2]) as axis:
  book = xlsxwriter.Workbook(sys.argv[3])
  sheet = book.add_worksheet()
  sheet.merge_range("C1:D1", "AET")
  sheet.merge_range("E1:F1", "IVCT")
  sheet.merge_range("G1:H1", "IVRT")
  sheet.merge_range("I1:J1", "MV E")
  sheet.write_row("A2", ["ID", "G type", *["Coefficient", "P-value"] * 4])
  mvData = []
  axisData = []
  reader = csv.reader(mv)
  for i in reader:
    mvData.append(i)
  reader = csv.reader(axis)
  for i in reader:
    axisData.append(i)
  gType = axisData[1][1:5]

  # i is for mv and output sheet
  # j is for axis
  for i,j in zip(range(1, 22, 4), range(2, 8), strict=True):
    sheet.merge_range(i + 1, 0, i + 4, 0, mvData[i][0])
    sheet.write_column(i + 1, 1, gType)
    aet = getMVData(mvData, i, 4)
    ivct = getMVData(mvData, i + 1, 4)
    ivrt = getMVData(mvData, i + 2, 4)
    mve = getMVData(mvData, i + 3, 4)
    gt1 = getAxisData(axisData, j, 1)
    gt2 = getAxisData(axisData, j, 2)
    gt3 = getAxisData(axisData, j, 3)
    gt4 = getAxisData(axisData, j, 4)
    for k in enumerate([gt1, gt2, gt3, gt4]):
      data = []
      for l in [aet, ivct, ivrt, mve]:
        res = scipy.stats.pearsonr(k[1], l)
        data.append(res.correlation)
        data.append(res.pvalue)
      sheet.write_row(i + k[0] + 1, 2, data)
  sheet.autofit()
  book.close()