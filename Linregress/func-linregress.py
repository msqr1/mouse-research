import scipy, csv, sys, xlsxwriter

def getLinregressData(data, top, left):
  x = [] # Week, 0 = baseline
  y = [] # Average cell values
  for i in range(5):
    if data[top][left + i] != '':
      x.append(i)
      y.append(float(data[top][left + i]))
  return x,y

# Output book
book = xlsxwriter.Workbook(sys.argv[2])
sheet = book.add_worksheet()
sheet.write_row("A1", ["ID", "Measurement", "Slope", "P-value"])
with open(sys.argv[1]) as f:
  data = []
  reader = csv.reader(f)
  for row in reader:
    data.append(row)
  tpCnt = 4
  for i in range(1, tpCnt * 6, tpCnt):
    sheet.merge_range(i, 0, i + tpCnt - 1, 0, data[i][0])
    for j in range(i, i + tpCnt):
      res = scipy.stats.linregress(*getLinregressData(data, i, 4))
      sheet.write_row(j, 1, [data[j][1], res.slope, res.pvalue])
sheet.autofit()
book.close()