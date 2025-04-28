import sys, glob, xlsxwriter, xmltodict

weekDir = sys.argv[1]
output = sys.argv[2]
animalDirs = glob.glob(f"{weekDir}/*")
animalIDs = [i[len(weekDir) + 1:] for i in animalDirs]
book = xlsxwriter.Workbook(output)

apex = book.add_worksheet("Apex")
pp = book.add_worksheet("PP")
smv = book.add_worksheet("SMV")
longAxis = book.add_worksheet("Long Axis")

# For normal Ws
for i in [apex, pp, smv, longAxis]:
  i.write_row("A1", ["ID", "Cell type"])
  i.set_column(1, 1, 19)
  i.merge_range("C1:H1", "Peak")
  i.merge_range("I1:N1", "TTP")
  i.merge_range("O1:T1", "ES")
  i.merge_range("F2:H2", "Circumferential")
  i.merge_range("L2:N2", "Circumferential")
  i.merge_range("R2:T2", "Circumferential")
  for j in ["C3", "F3", "I3", "L3", "O3", "R3"]:
    i.write_row(j, ["Endo", "Myo", "Epi"])

for i in [apex, pp, smv]:
  i.merge_range("C2:E2", "Radial")
  i.merge_range("I2:K2", "Radial")
  i.merge_range("O2:Q2", "Radial")

longAxis.merge_range("C2:E2", "Longitudinal")
longAxis.merge_range("I2:K2", "Longitudinal")
longAxis.merge_range("O2:Q2", "Longitudinal")

for i in range(0, len(animalDirs)):
  apex.merge_range(f"A{4 + 4 * i}:A{7 + 4 * i}", animalIDs[i])
  apex.write_column(f"B{4 + 4 * i}:B{7 + 4 * i}", ["13-apical anterior", "16-apical lateral", "15-apical inferior", "14-apical septal"])
  with open(glob.glob(f"{animalDirs[i]}/Apex*/*TTP(apex).xml")[0]) as f: 
    for ws in xmltodict.parse(f.read())["Workbook"]["Worksheet"]:
      name = ws["@ss:Name"]
      rows = ws["Table"]["Row"]
      if name == "Strain-Endo TTP":
        if(len(rows) == 20):
          for j in range(0, 4):
            cells = rows[7 + j]["Cell"]
            apex.write(f"C{4 + 4 * i + j}", cells[1]["Data"]["#text"])
            apex.write(f"I{4 + 4 * i + j}", cells[2]["Data"]["#text"])
            apex.write(f"O{4 + 4 * i + j}", cells[3]["Data"]["#text"])
          
          for j in range(0, 4):
            cells = rows[14 + j]["Cell"]
            apex.write(f"F{4 + 4 * i + j}", cells[1]["Data"]["#text"])
            apex.write(f"L{4 + 4 * i + j}", cells[2]["Data"]["#text"])
            apex.write(f"R{4 + 4 * i + j}", cells[3]["Data"]["#text"])
        else:
          for j in range(0, 4):
            cells = rows[7 + j]["Cell"]
            apex.write(f"F{4 + 4 * i + j}", cells[1]["Data"]["#text"])
            apex.write(f"L{4 + 4 * i + j}", cells[2]["Data"]["#text"])
            apex.write(f"R{4 + 4 * i + j}", cells[3]["Data"]["#text"])
      elif name == "Strain-Myo TTP":
        if(len(rows) == 20):
          for j in range(0, 4):
            cells = rows[7 + j]["Cell"]
            apex.write(f"D{4 + 4 * i + j}", cells[1]["Data"]["#text"])
            apex.write(f"J{4 + 4 * i + j}", cells[2]["Data"]["#text"])
            apex.write(f"P{4 + 4 * i + j}", cells[3]["Data"]["#text"])
          for j in range(0, 4):
            cells = rows[14 + j]["Cell"]
            apex.write(f"G{4 + 4 * i + j}", cells[1]["Data"]["#text"])
            apex.write(f"M{4 + 4 * i + j}", cells[2]["Data"]["#text"])
            apex.write(f"S{4 + 4 * i + j}", cells[3]["Data"]["#text"])
        else:
          for j in range(0, 4):
            cells = rows[7 + j]["Cell"]
            apex.write(f"G{4 + 4 * i + j}", cells[1]["Data"]["#text"])
            apex.write(f"M{4 + 4 * i + j}", cells[2]["Data"]["#text"])
            apex.write(f"S{4 + 4 * i + j}", cells[3]["Data"]["#text"])
      elif name == "Strain-Epi TTP":
        if(len(rows) == 20):
          for j in range(0, 4):
            cells = rows[7 + j]["Cell"]
            apex.write(f"E{4 + 4 * i + j}", cells[1]["Data"]["#text"])
            apex.write(f"K{4 + 4 * i + j}", cells[2]["Data"]["#text"])
            apex.write(f"Q{4 + 4 * i + j}", cells[3]["Data"]["#text"])
          for j in range(0, 4):
            cells = rows[14 + j]["Cell"]
            apex.write(f"H{4 + 4 * i + j}", cells[1]["Data"]["#text"])
            apex.write(f"N{4 + 4 * i + j}", cells[2]["Data"]["#text"])
            apex.write(f"T{4 + 4 * i + j}", cells[3]["Data"]["#text"])
        else:
          cells = rows[7 + j]["Cell"]
          apex.write(f"H{4 + 4 * i + j}", cells[1]["Data"]["#text"])
          apex.write(f"N{4 + 4 * i + j}", cells[2]["Data"]["#text"])
          apex.write(f"T{4 + 4 * i + j}", cells[3]["Data"]["#text"])

  pp.merge_range(f"A{4 + 6 * i}:A{9 + 6 * i}", animalIDs[i])
  pp.write_column(f"B{4 + 6 * i}:B{7 + 6 * i}", ["07-mid anterior", "12-mid anterolateral", "11-mid inferolateral", "10-mid inferior", "09-mid inferoseptal", "08-mid anteroseptal"])
  with open(glob.glob(f"{animalDirs[i]}/PP*/*TTP(pm).xml")[0]) as f:
    for ws in xmltodict.parse(f.read())["Workbook"]["Worksheet"]:
      name = ws["@ss:Name"]
      rows = ws["Table"]["Row"]
      if name == "Strain-Endo TTP":
        if(len(rows) == 24):
          for j in range(0, 6):
            cells = rows[7 + j]["Cell"]
            pp.write(f"C{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            pp.write(f"I{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            pp.write(f"O{4 + 6 * i + j}", cells[3]["Data"]["#text"])

          for j in range(0, 6):
            cells = rows[16 + j]["Cell"]
            pp.write(f"F{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            pp.write(f"L{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            pp.write(f"R{4 + 6 * i + j}", cells[3]["Data"]["#text"])
        else:
          for j in range(0, 6):
            cells = rows[7 + j]["Cell"]
            pp.write(f"F{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            pp.write(f"L{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            pp.write(f"R{4 + 6 * i + j}", cells[3]["Data"]["#text"])

      elif name == "Strain-Myo TTP":
        if(len(rows) == 24):
          for j in range(0, 6):
            cells = rows[7 + j]["Cell"]
            pp.write(f"D{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            pp.write(f"J{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            pp.write(f"P{4 + 6 * i + j}", cells[3]["Data"]["#text"])
          for j in range(0, 6):
            cells = rows[16 + j]["Cell"]
            pp.write(f"G{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            pp.write(f"M{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            pp.write(f"S{4 + 6 * i + j}", cells[3]["Data"]["#text"])
        else:
          for j in range(0, 6):
            cells = rows[7 + j]["Cell"]
            pp.write(f"G{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            pp.write(f"M{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            pp.write(f"S{4 + 6 * i + j}", cells[3]["Data"]["#text"])
      elif name == "Strain-Epi TTP":
        if(len(rows) == 24):
          for j in range(0, 6):
            cells = rows[7 + j]["Cell"]
            pp.write(f"E{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            pp.write(f"K{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            pp.write(f"Q{4 + 6 * i + j}", cells[3]["Data"]["#text"])
    
          for j in range(0, 6):
            cells = rows[16 + j]["Cell"]
            pp.write(f"H{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            pp.write(f"N{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            pp.write(f"T{4 + 6 * i + j}", cells[3]["Data"]["#text"])
        else:
          for j in range(0, 6):
            cells = rows[7 + j]["Cell"]
            pp.write(f"H{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            pp.write(f"N{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            pp.write(f"T{4 + 6 * i + j}", cells[3]["Data"]["#text"])

  smv.merge_range(f"A{4 + 6 * i}:A{9 + 6 * i}", animalIDs[i])
  smv.write_column(f"B{4 + 6 * i}:B{7 + 6 * i}", ["01-basal anterior", "06-basal anterolateral", "05-basal inferolateral", "04-basal inferior", "03-basal inferoseptal", "02-basal anteroseptal"])
  with open(glob.glob(f"{animalDirs[i]}/sMV*/*TTP(mv).xml")[0]) as f:
    for ws in xmltodict.parse(f.read())["Workbook"]["Worksheet"]:
      name = ws["@ss:Name"]
      rows = ws["Table"]["Row"]
      if name == "Strain-Endo TTP":
        if(len(rows) == 24):
          for j in range(0, 6):
            cells = rows[7 + j]["Cell"]
            smv.write(f"C{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            smv.write(f"I{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            smv.write(f"O{4 + 6 * i + j}", cells[3]["Data"]["#text"])

          for j in range(0, 6):
            cells = rows[16 + j]["Cell"]
            smv.write(f"F{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            smv.write(f"L{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            smv.write(f"R{4 + 6 * i + j}", cells[3]["Data"]["#text"])
        else:
          for j in range(0, 6):
            cells = rows[7 + j]["Cell"]
            smv.write(f"F{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            smv.write(f"L{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            smv.write(f"R{4 + 6 * i + j}", cells[3]["Data"]["#text"])

      elif name == "Strain-Myo TTP":
        if(len(rows) == 24):
          for j in range(0, 6):
            cells = rows[7 + j]["Cell"]
            smv.write(f"D{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            smv.write(f"J{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            smv.write(f"P{4 + 6 * i + j}", cells[3]["Data"]["#text"])
          for j in range(0, 6):
            cells = rows[16 + j]["Cell"]
            smv.write(f"G{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            smv.write(f"M{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            smv.write(f"S{4 + 6 * i + j}", cells[3]["Data"]["#text"])
        else:
          for j in range(0, 6):
            cells = rows[7 + j]["Cell"]
            smv.write(f"G{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            smv.write(f"M{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            smv.write(f"S{4 + 6 * i + j}", cells[3]["Data"]["#text"])
      elif name == "Strain-Epi TTP":
        if(len(rows) == 24):
          for j in range(0, 6):
            cells = rows[7 + j]["Cell"]
            smv.write(f"E{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            smv.write(f"K{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            smv.write(f"Q{4 + 6 * i + j}", cells[3]["Data"]["#text"])
    
          for j in range(0, 6):
            cells = rows[16 + j]["Cell"]
            smv.write(f"H{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            smv.write(f"N{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            smv.write(f"T{4 + 6 * i + j}", cells[3]["Data"]["#text"])
        else:
          for j in range(0, 6):
            cells = rows[7 + j]["Cell"]
            smv.write(f"H{4 + 6 * i + j}", cells[1]["Data"]["#text"])
            smv.write(f"N{4 + 6 * i + j}", cells[2]["Data"]["#text"])
            smv.write(f"T{4 + 6 * i + j}", cells[3]["Data"]["#text"])
  
  longAxis.merge_range(f"A{4 + 7 * i}:A{10 + 7 * i}", animalIDs[i])
  longAxis.write_column(f"B{4 + 7 * i}:B{7 + 7 * i}", ["05-basal inferolateral", "11-mid inferolateral", "16-apical lateral", "17-apex", "14-apical septal", "08-mid anteroseptal", "02-basal anteroseptal"])
  with open(glob.glob(f"{animalDirs[i]}/Long Axis*/*TTP(pslax).xml")[0]) as f:
    for ws in xmltodict.parse(f.read())["Workbook"]["Worksheet"]:
      name = ws["@ss:Name"]
      rows = ws["Table"]["Row"]
      if name == "Strain-Endo TTP":
        if(len(rows) == 26):
          for j in range(0, 7):
            cells = rows[7 + j]["Cell"]
            longAxis.write(f"C{4 + 7 * i + j}", cells[1]["Data"]["#text"])
            longAxis.write(f"I{4 + 7 * i + j}", cells[2]["Data"]["#text"])
            longAxis.write(f"O{4 + 7 * i + j}", cells[3]["Data"]["#text"])

          for j in range(0, 7):
            cells = rows[17 + j]["Cell"]
            longAxis.write(f"F{4 + 7 * i + j}", cells[1]["Data"]["#text"])
            longAxis.write(f"L{4 + 7 * i + j}", cells[2]["Data"]["#text"])
            longAxis.write(f"R{4 + 7 * i + j}", cells[3]["Data"]["#text"])
        else:
          for j in range(0, 7):
            cells = rows[7 + j]["Cell"]
            longAxis.write(f"F{4 + 7 * i + j}", cells[1]["Data"]["#text"])
            longAxis.write(f"L{4 + 7 * i + j}", cells[2]["Data"]["#text"])
            longAxis.write(f"R{4 + 7 * i + j}", cells[3]["Data"]["#text"])

      elif name == "Strain-Myo TTP":
        if(len(rows) == 26):
          for j in range(0, 7):
            cells = rows[7 + j]["Cell"]
            longAxis.write(f"D{4 + 7 * i + j}", cells[1]["Data"]["#text"])
            longAxis.write(f"J{4 + 7 * i + j}", cells[2]["Data"]["#text"])
            longAxis.write(f"P{4 + 7 * i + j}", cells[3]["Data"]["#text"])
          for j in range(0, 7):
            cells = rows[17 + j]["Cell"]
            longAxis.write(f"G{4 + 7 * i + j}", cells[1]["Data"]["#text"])
            longAxis.write(f"M{4 + 7 * i + j}", cells[2]["Data"]["#text"])
            longAxis.write(f"S{4 + 7 * i + j}", cells[3]["Data"]["#text"])
        else:
          for j in range(0, 7):
            cells = rows[7 + j]["Cell"]
            longAxis.write(f"G{4 + 7 * i + j}", cells[1]["Data"]["#text"])
            longAxis.write(f"M{4 + 7 * i + j}", cells[2]["Data"]["#text"])
            longAxis.write(f"S{4 + 7 * i + j}", cells[3]["Data"]["#text"])

      elif name == "Strain-Epi TTP":
        if(len(rows) == 26):
          for j in range(0, 7):
            cells = rows[7 + j]["Cell"]
            longAxis.write(f"E{4 + 7 * i + j}", cells[1]["Data"]["#text"])
            longAxis.write(f"K{4 + 7 * i + j}", cells[2]["Data"]["#text"])
            longAxis.write(f"Q{4 + 7 * i + j}", cells[3]["Data"]["#text"])
    
          for j in range(0, 7):
            cells = rows[17 + j]["Cell"]
            longAxis.write(f"H{4 + 7 * i + j}", cells[1]["Data"]["#text"])
            longAxis.write(f"N{4 + 7 * i + j}", cells[2]["Data"]["#text"])
            longAxis.write(f"T{4 + 7 * i + j}", cells[3]["Data"]["#text"])
        else:
          for j in range(0, 7):
            cells = rows[7 + j]["Cell"]
            longAxis.write(f"H{4 + 7 * i + j}", cells[1]["Data"]["#text"])
            longAxis.write(f"N{4 + 7 * i + j}", cells[2]["Data"]["#text"])
            longAxis.write(f"T{4 + 7 * i + j}", cells[3]["Data"]["#text"])

book.close()