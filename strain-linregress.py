import scipy, csv, sys

def getLinregressData(data, top, left, cnt):
  x = [] # Week, 0 = baseline
  y = [] # Average cell values
  for i in range(5):
    if data[top][left + i * 3] != '':
      x.append(i)
      y.append(sum([float(data[j][left + i * 3]) for j in range(top, top + cnt)]) / cnt)
  return x,y

with open(sys.argv[1]) as f:
  data = []
  reader = csv.reader(f)
  for row in reader:
    data.append(row)
  tpCnt = 1
  while data[tpCnt + 2][0] == '':
    tpCnt += 1
  for i in range(2, tpCnt * 6, tpCnt):
    print(f"ID: {data[i][0]}")
    res = scipy.stats.linregress(*getLinregressData(data, i, 2, tpCnt))
    print("  TTP:")
    print(f"    Slope: {res.slope}")
    print(f"    p-value: {res.pvalue}")
    res = scipy.stats.linregress(*getLinregressData(data, i, 3, tpCnt))
    print(f"  Peak:")
    print(f"    Slope: {res.slope}")
    print(f"    p-value: {res.pvalue}")
    res = scipy.stats.linregress(*getLinregressData(data, i, 3, tpCnt))
    print(f"  ES:")
    print(f"    Slope: {res.slope}")
    print(f"    p-value: {res.pvalue}\n")