import scipy.stats, csv, sys, xlsxwriter

def getMVData(input, top, left):
  data = []
  for i in range(left, left + 5):
    if input[top][i] != '':
      data.append(float(input[top][i]))
  return data

def getStrainData(input, top, left, cnt):
  data = []  
  for i in range(left, left + 15, 3):
    if input[top][i] != '':
      data.append(sum([float(input[j][i]) for j in range(top, top + cnt)]) / cnt)
  return data

with open(sys.argv[1]) as mv, open(sys.argv[2]) as strain:
  book = xlsxwriter.Workbook(sys.argv[3])
  sheet = book.add_worksheet()
  sheet.merge_range("C1:D1", "AET")
  sheet.merge_range("E1:F1", "IVCT")
  sheet.merge_range("G1:H1", "IVRT")
  sheet.merge_range("I1:J1", "MV E")
  sheet.write_row("A2", ["ID", "TTP/Peak/ES", *["Coefficient", "P-value"] * 4])
  mvData = []
  strainData = []
  reader = csv.reader(mv)
  for row in reader:
    mvData.append(row)
  reader = csv.reader(strain)
  for row in reader:
    strainData.append(row)
  tpCnt = 1
  while strainData[tpCnt + 2][0] == '':
    tpCnt += 1

  # i is for mv
  # j is for strain
  # k is for the output sheet
  for i,j,k in zip(range(1, 22, 4), range(2, tpCnt * 6, tpCnt), range(2, 20, 3), strict=True):
    aet = getMVData(mvData, i, 4)
    ivct = getMVData(mvData, i + 1, 4)
    ivrt = getMVData(mvData, i + 2, 4)
    mve = getMVData(mvData, i + 3, 4)
    ttp = getStrainData(strainData, j, 2, tpCnt)
    peak = getStrainData(strainData, j, 3, tpCnt)
    es = getStrainData(strainData, j, 4, tpCnt)
    sheet.merge_range(k, 0, k + 2, 0, mvData[i][0])
    sheet.write_column(k, 1, ["TTP", "Peak", "ES"])
    for l in enumerate([ttp, peak, es]):
      data = []
      for m in [aet, ivct, ivrt, mve]:
        res = scipy.stats.pearsonr(l[1], m)
        data.append(res.correlation)
        data.append(res.pvalue)
      sheet.write_row(k + l[0], 2, data)
  sheet.autofit()
  book.close()
 