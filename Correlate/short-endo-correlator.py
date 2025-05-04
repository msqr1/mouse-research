import scipy.stats, csv, sys, xlsxwriter

def getMVData(input, top, left):
  data = []
  for i in range(left, left + 5):
    if input[top][i] != '':
      data.append(float(input[top][i]))
  return data

def getEndoData(input, top, left):
  data = []
  for i in range(left, left + 10, 2):
    if input[top][i] != '':
      data.append(float(input[top][i]))
  return data

with open(sys.argv[1]) as mv, open(sys.argv[2]) as endo:
  book = xlsxwriter.Workbook(sys.argv[3])
  sheet = book.add_worksheet()
  sheet.merge_range("C1:D1", "AET")
  sheet.merge_range("E1:F1", "IVCT")
  sheet.merge_range("G1:H1", "IVRT")
  sheet.merge_range("I1:J1", "MV E")
  sheet.write_row("A2", ["ID", "Endo type", *["Coefficient", "P-value"] * 4])
  mvData = []
  endoData = []