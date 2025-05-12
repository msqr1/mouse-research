import sys, glob, xlsxwriter

# Week directory (raw data)
weekDir = sys.argv[1]

# Output sheet
out = sys.argv[2]

animalDirs = glob.glob(f"{weekDir}/*")
animalIDs = [i[len(weekDir) + 1:] for i in animalDirs]
book = xlsxwriter.Workbook(out)

apex = book.add_worksheet("Apex")
pp = book.add_worksheet("PP")
smv = book.add_worksheet("SMV")
longAxis = book.add_worksheet("Long Axis")
for i in [apex, pp, smv]:
  i.write_row("A1", ["ID", "MyoGCS", "EndoGCS", "EndoROT", "MyoROT"])

longAxis.write_row("A1", ["ID", "MyoGCS*", "MyoGLS", "EndoGCS*", "EndoGLS"])

# Information line number:

# apex, pp and smv:
#   101: MyoGCS
#   102: EndoGCS
#   106: EndoRot
#   107: MyoROT

# long axis:
#   105: MyoGCS*
#   106: MyoGLS
#   107: EndoGCS*
#   108: EndoGLS

arr = [
  [apex, "Apex*/*Analysis(ref. view apex).txt"],
  [pp, "PP*/*Analysis(ref. view pm).txt"],
  [smv, "sMV*/*Analysis(ref. view mv).txt"]
]

for i in range(len(animalDirs)):
  for j in arr:
    j[0].write(1 + i, 0, animalIDs[i])
    a = glob.glob(f"{animalDirs[i]}/{j[1]}")
    if(len(a)) > 0:
      with open(a[0]) as f:
        for k in range(100):
          next(f)
        myoGCS = float(f.readline()[7:-3])
        endoGCS = float(f.readline()[8:-3])
        for k in range(3):
          next(f)
        endoROT = float(f.readline()[8:-5])
        myoROT = float(f.readline()[7:-5])
        j[0].write_row(1 + i, 1, [myoGCS, endoGCS, endoROT, myoROT])
  longAxis.write(1 + i, 0, animalIDs[i])
  a = glob.glob(f"{animalDirs[i]}/Long Axis*/*Analysis(ref. view pslax).txt")
  if(len(a)) > 0:
    with open(a[0]) as f:
      for k in range(0, 104):
        next(f)
      myoGCS = float(f.readline()[8:-3])
      myoGLS = float(f.readline()[7:-3])
      endoGCS = float(f.readline()[9:-3])
      endoGLS = float(f.readline()[8:-3])
      longAxis.write_row(1 + i, 1, [myoGCS, myoGLS, endoGCS, endoGLS])
book.close()