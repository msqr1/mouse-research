import sys, glob, xlsxwriter

weekDir = sys.argv[1]
output = sys.argv[2]
animalDirs = glob.glob(f"{weekDir}/*")
animalIDs = [i[len(weekDir) + 1:] for i in animalDirs]
book = xlsxwriter.Workbook(output)

apex = book.add_worksheet("Apex")
pp = book.add_worksheet("PP")
smv = book.add_worksheet("SMV")
longAxis = book.add_worksheet("Long Axis")
for i in [apex, pp, smv]:
  i.write_row("A1", ["ID", "EndoGCS", "EndoROT"])

longAxis.write_row("A1", ["ID", "EndoGCS*", "EndoGLS"])

# For apex, pp and smv the EndoGCS is always at line 102, and EndoROT is at line 106
# For long axis, EndoGCS is at line 106, and EndoGLS is at line 107
arr = [
  [apex, "Apex*/*Analysis(ref. view apex).txt"],
  [pp, "PP*/*Analysis(ref. view pm).txt"],
  [smv, "sMV*/*Analysis(ref. view mv).txt"]
]

for i in range(0, len(animalDirs)):
  for j in arr:
    j[0].write(f"A{2 + i}", animalIDs[i])
    a = glob.glob(f"{animalDirs[i]}/{j[1]}")
    if(len(a)) > 0:
      with open(a[0]) as f:
        for k in range(0, 101):
          next(f)
        endoGCS = f.readline()[8:-3].strip()
        for k in range(0, 3):
          next(f)
        j[0].write_row(f"B{2 + i}", [endoGCS, f.readline()[8:-5].strip()])
  longAxis.write(f"A{2 + i}", animalIDs[i])
  a = glob.glob(f"{animalDirs[i]}/Long Axis*/*Analysis(ref. view pslax).txt")
  if(len(a)) > 0:
    with open(a[0]) as f:
      for k in range(0, 105):
        next(f)
      longAxis.write_row(f"B{2 + i}", [f.readline()[9:-3].strip(), f.readline()[8:-3].strip()])
book.close()