import csv, sys, matplotlib as mpl, matplotlib.pyplot as plt

# ONE sheet from avg-col-extractor, converted to CSV
colData = sys.argv[1]

# Graph title
title = sys.argv[2]

# y-axis label
yLabel = sys.argv[3]

# Output image
out = sys.argv[4]

with open(colData) as f:
  data = []
  reader = csv.reader(f)
  for row in reader:
    data.append(row)
  toPlot = []
  tickLabels = []
  for i in range(1, len(data[0])):
    toPlot.append([])
    tickLabels.append(f"Week {i - 1}")
    for j in range(1, len(data)):
      a = data[j][i]
      if(a != ''):
        toPlot[-1].append(abs(float(a)))
  fig, ax = plt.subplots(figsize=(16, 9))
  thickness = 12
  #thickness -= 2
  prop = {"linewidth": thickness}
  bplot = ax.boxplot(toPlot, tick_labels=tickLabels, patch_artist=True, showfliers=True, boxprops=prop, medianprops=prop, whiskerprops=prop, capprops=prop, flierprops={'markerfacecolor': 'black', "markersize": thickness, **prop}, widths=.4)
  for axis in ['top', 'bottom', 'left', 'right']:
    ax.spines[axis].set_linewidth(thickness)
  ax.yaxis.set_tick_params(width=thickness, length=20, direction="inout", labelsize=27, pad=10)
  ax.xaxis.set_tick_params(width=thickness, length=20, direction="inout", labelsize=35, pad=10)
  plt.ylabel(yLabel, fontdict={"fontsize": 40}, labelpad=25)
  plt.title(title, fontdict={"fontsize": 45}, pad=25)
  plt.savefig(out, dpi=150)
  #plt.show()