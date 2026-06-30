import pandas as pd
import datetime as dt
import numpy as np
import matplotlib.pyplot as plt
import io
import base64

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_PATH = "./DS Sorting\\"
DATA_PATH = BASE_PATH + "Data\\"
HTML_PATH = BASE_PATH + "Html\\"
WORK_PATH = BASE_PATH + "work3\\"
BI_PATH = "./bi_data/DS_Sorting_Weekly\\BI Data\\"


def save_bi_csv(df, name):
    df.to_csv(WORK_PATH + name)
    df.to_csv(BI_PATH + name)

FILE_KPI_PUTAWAY = DATA_PATH + 'PUTAWAY_UNIT.csv'
FILE_KPI_SORT = DATA_PATH + 'SORTING_MODULE.csv'
FILE_102 = DATA_PATH + '102.csv'

# ── Auto-calculate dates: previous Monday → previous Saturday ──────────────────
# This script is intended to be run on a Monday.
# "Previous Monday" = the Monday of the prior week (7 days ago).
today = dt.date.today()
days_since_monday = today.weekday()  # Monday=0
if days_since_monday == 0:
    prev_monday = today - dt.timedelta(days=7)
else:
    prev_monday = today - dt.timedelta(days=days_since_monday + 7)

prev_saturday = prev_monday + dt.timedelta(days=5)

# Build date list Mon-Sat (6 days)
TargetDate1 = [(prev_monday + dt.timedelta(days=i)).strftime('%Y%m%d') for i in range(6)] + ['0000']
TargetDate2 = [(prev_monday + dt.timedelta(days=i)).strftime('%Y%m%d') for i in range(6)]

sHtmlDate = prev_monday.strftime('%m/%d/%Y')
sHtmlFile = prev_monday.strftime('%Y%m%d') + '.html'

# ── Configuration ──────────────────────────────────────────────────────────────
iSortKPI_Available = 1  # KPI files available: 0=No, 1=Yes

FIG_TITLE = f'DS Sorting KPI [Operation Date: week {sHtmlDate}]'
HTML_NAME = 'DS_Sorting_KPI_' + sHtmlFile
SORT_INTERVAL = 20

# ── Colors ─────────────────────────────────────────────────────────────────────
BLUE = 'dodgerblue'
GREEN = 'darkseagreen'
RED = 'red'
PINK = 'lightpink'
ORANGE = 'darkorange'



def calc_time(timeFrom, timeTo, pick_count, iSec):

  if pick_count == 1:
    tspend = iSec
    ispeed = round((tspend / pick_count) * 60, 1)

  else:
    if timeTo == timeFrom:
      tspend0 = timeTo + dt.timedelta(minutes=1)
      tspend0 = tspend0 - timeFrom
    else:
      tspend0 = timeTo - timeFrom
    
    tspend1 = tspend0.total_seconds()

    ispeed = 0
    if tspend1 > 0:
      tspend = tspend1 / 60
      ispeed = round((tspend / pick_count) * 60, 1)
    else:
      tspend0 = 0.7
      ispeed = round((tspend0 / pick_count) * 60, 1)

  return tspend, ispeed

def img2html(fig):
    sio = io.BytesIO()
    fig.savefig(sio, format='png')
    image_bin = base64.b64encode(sio.getvalue())
    return HTML_TMP.format(image_bin=str(image_bin)[2:-1])


### Graph Plotting #####################################################################################################################


def set_graph_label2(ax1, ax2, stitle, x_label1_flg, x_label, y_label1_flg, y_label1, y_label2_flg, y_label2, xtickrotation, fontSize):
    ax1.set_title(stitle)
    ax1.grid(axis="y")
    if x_label1_flg == 1:
        ax1.set_ylabel(x_label)

    if y_label1_flg == 1:
        ax1.set_ylabel(y_label1)

    if y_label2_flg == 1:
        ax2.set_ylabel(y_label2)

    ax1.set_xticks(ax1.get_xticks())
    xticklabels = ax1.get_xticklabels()
    ax1.set_xticklabels(xticklabels, fontsize=fontSize, rotation=xtickrotation)


def plot_3bar_3line(ax, bar_width, df_AB, x_key, y_key1, y_key2, stitle, x_label, y_label1, y_label2):
    #bar_width = 0.3
    ax2 = ax.twinx()
    ax.grid(axis="y")

    shifts = ['shift1', 'shift2']
    index = np.arange(len(shifts))

    def _draw(category, offset, bar_label, bar_color, line_label, line_color, always_line):
        raw = df_AB[df_AB['Mix_Solid'] == category]
        n = raw.shape[0]
        if n == 0:
            return
        sub = raw.set_index(x_key).reindex(shifts).fillna(0)
        bars = ax.bar(index + offset, sub[y_key1], width=bar_width,
                      label=bar_label, alpha=0.6, color=bar_color)
        ax.bar_label(bars, label_type='edge')
        if always_line or n == 2:
            ax2.plot(index + bar_width, sub[y_key2],
                     label=line_label, color=line_color, marker='o')

    print(df_AB)
    _draw('Mx_Sort', 0,             'Mixsort count',    PINK,  'Mix sort speed',   'hotpink', True)
    _draw('Mx_Repl', bar_width,     'Mix repl count',   BLUE,  'Mix repl speed',   'blue',    False)
    _draw('So_Repl', bar_width * 2, 'Solid repl count', GREEN, 'Solid repl speed', 'green',   False)

    plt.xticks(index + bar_width, ('Shift1', 'Shift2'))
    ax.grid(axis="y")

    list_Shift = df_AB[x_key].values.tolist()
    list_ispeed = df_AB['ispeed'].values.tolist()

    ax_pos = ax2.get_position()

    for iLoop, value in enumerate(list_ispeed):
      if list_Shift[iLoop] == 'shift1':
        ax2.text(ax_pos.x0 + 0.22, list_ispeed[iLoop]-5, value)
      else:
        ax2.text(ax_pos.x0 + 1.15, list_ispeed[iLoop]-5, value)

    ymin, ymax = ax.get_ylim()
    ax.set_ylim(ymin, ymax + 800)
    ax2.set_ylim(0,90)

    set_graph_label2(ax, ax2, stitle, 1, x_label, 1, y_label1, 1, y_label2, 0, 12)
    #ax2.axhline(BASE_TIME, 0, 1, c="purple", ls='--')

    hans, labs = ax.get_legend_handles_labels()
    ax.legend(handles=hans, labels=labs, bbox_to_anchor=(1.1, 0.9), loc='upper left', borderaxespad=0)
    ax2.legend(loc="upper left", bbox_to_anchor=(1.1, 0.6))


def plot_4bar_4line(ax, df_A, df_B, x_key, y_key1, y_key2, y_key3,  y_key4,
                     stitle, x_label, y_label1, y_label2, base_line, bar_color1, plot_color1, bar_color2, plot_color2, iSepaleter):
    width = 0.32
    ax2 = ax.twinx()
    ax.grid(axis="y")

    pA1 = ax.bar(df_A[x_key], df_A[y_key1], align="edge", width=-width, label='Shift1 Sort_Count', alpha=0.6, color=PINK) 
    pA2 = ax.bar(df_A[x_key], df_A[y_key2], align="edge", width=width, label='Shift1 Repl_Count', alpha=0.6, color=BLUE ) 
    pA3 = ax2.plot(df_A[x_key],df_A[y_key3], label='Shift1 Sort_Speed', color="hotpink", marker='o') 
    pA4 = ax2.plot(df_A[x_key],df_A[y_key4], label='Shift1 Repl_Speed', color="blue", marker='o') 

    pB1 = ax.bar(df_B[x_key], df_B[y_key1], align="edge", width=-width, label='Shift2 Sort_Count', alpha=0.6, color=PINK) 
    pB2 = ax.bar(df_B[x_key], df_B[y_key2], align="edge", width=width, label='Shift2 Repl_Count', alpha=0.6, color=BLUE) 
    pB3 = ax2.plot(df_B[x_key],df_B[y_key3], label='Shift2 Sort_Speed', linestyle='dashed', color="hotpink", marker='*') 
    pA4 = ax2.plot(df_B[x_key],df_B[y_key4], label='Shift2 Repl_Speed', linestyle='dashed', color="blue", marker='*') 

    ax.bar_label(pA1,label_type = 'edge')
    ax.bar_label(pA2,label_type = 'edge')
    ax.bar_label(pB1,label_type = 'edge')
    ax.bar_label(pB2,label_type = 'edge')

    set_graph_label2(ax, ax2, stitle, 1, x_label, 1, y_label1, 1, y_label2, 270, 8)

    ymin, ymax = ax.get_ylim()
    ax.set_ylim(ymin, ymax + 100)

    ymin2, ymax2 = ax2.get_ylim()
    ax2.set_ylim(0, ymax2 + 5)

    ax.grid(axis="y")

    ax.legend([pA1, pA2],
              {'Sort_Count','Repl_Count'}, 
              bbox_to_anchor=(1.07, 1), loc='upper left')
    ax2.legend(loc="upper left", bbox_to_anchor=(1.07, 0.8))
    
    ax.vlines(iSepaleter, ymin, ymax - 10, linestyles='dashed', colors='purple', linewidth = 2)


def plot_2bar_2line(ax, df_A, df_B, x_key, y_key1, y_key2,
                     stitle, x_label, y_label1, y_label2, bar_color1, iSepaleter):
    bar_width = 0.4
    ax2 = ax.twinx()
    ax.grid(axis="y")

    pA1 = ax.bar(df_A[x_key], df_A[y_key1], width=bar_width, alpha=0.6, color=bar_color1) 
    pA3 = ax2.plot(df_A[x_key],df_A[y_key2], label='Shift1 Repl_Speed', linestyle='solid', color="green", marker='o') 

    pB1 = ax.bar(df_B[x_key], df_B[y_key1], width=bar_width, alpha=0.6, color=bar_color1) 
    pB3 = ax2.plot(df_B[x_key],df_B[y_key2], label='Shift2 Repl_Speed', linestyle='dashed', color="green", marker='*') 

    ax.bar_label(pA1,label_type = 'edge')
    ax.bar_label(pB1,label_type = 'edge')

    set_graph_label2(ax, ax2, stitle, 1, x_label, 1, y_label1, 1, y_label2, 270, 8)

    ymin, ymax = ax.get_ylim()
    ax.set_ylim(ymin, ymax + 100)

    ymin2, ymax2 = ax2.get_ylim()
    ax2.set_ylim(0, ymax2 + 2)


    ax.grid(axis="y")
    ax.legend([pA1],
              {'Repl_Count'}, 
              bbox_to_anchor=(1.1, 1), loc='upper left')
    ax2.legend(loc="upper left", bbox_to_anchor=(1.1, 0.8))

    ax.vlines(iSepaleter, ymin, ymax, linestyles='dashed', colors='purple', linewidth = 2)

### Glaph Plot End ####################################################################################################################

### Functions Start ####################################################################################################################
### Common    Start ####################################################################################################################

def calc_speed_2(df, item_list1, group_list1):

  df_sum = df[item_list1].groupby(group_list1)\
    .agg(tspend=('tspend', 'sum'),sortCount=('sortCount', 'sum'))

  df_sum = df_sum.reset_index()

  df_sum = df_sum.assign(ispeed=0)
  if df_sum.shape[0] > 0:
    df_sum['ispeed'] = df_sum.apply(lambda x: round((x['tspend'] / x['sortCount']) * 60, 1), axis=1)

  return df_sum


def set_per_shift(flg_RS, df):
  if flg_RS == 'R':
    df_rmx= df[df['Mix_Solid'] == 'MO']
    df_rmx = calc_speed_2(df_rmx, ['sShift','Mix_Solid','sortCount', 'tspend'], ['sShift','Mix_Solid'])
    print(df_rmx)

    for iLoop1 in range(df_rmx.shape[0]):
      df_rmx.iat[iLoop1,1] = 'Mx_Repl'
      #df_rmx.iat[0,1] = 'Mx_Repl'
      #df_rmx.iat[1,1] = 'Mx_Repl'
    print(df_rmx)

    df_rmx2= df[df['Mix_Solid'] == 'SO']
    df_rmx2 = calc_speed_2(df_rmx2, ['sShift','Mix_Solid','sortCount', 'tspend'], ['sShift','Mix_Solid'])
    print(df_rmx2)
    for iLoop1 in range(df_rmx2.shape[0]):
      df_rmx2.iat[iLoop1,1] = 'So_Repl'
    #df_rmx2.iat[0,1] = 'So_Repl'
    #df_rmx2.iat[1,1] = 'So_Repl'
    print(df_rmx2)

    df_repl = pd.concat([df_rmx, df_rmx2], ignore_index=True)
    print(df_repl)

    df_sum = df_repl
  else:
    df_s = calc_speed_2(df, ['sShift','Mix_Solid','sortCount', 'tspend'], ['sShift','Mix_Solid'])

    print(df_s)
    if df_s.shape[0] <= 0:
      wList = [['shift1', 'Mx_Sort',0,0,0],['shift2', 'Mx_Sort',0,0,0]]
      w_df= pd.DataFrame(data = wList, columns = ["sShift", "Mix_Solid",  "tspend",  "sortCount",  "ispeed"])
      print(w_df)

    elif df_s.shape[0] == 1:
      df_s.iat[0,1] = 'Mx_Sort'
      if df_s.iat[0,0] == 'shift1':
        wList = [['shift2', 'Mx_Sort',0,0,0]]
      else:
        wList = [['shift1', 'Mx_Sort',0,0,0]]

      w_df= pd.DataFrame(data = wList, columns = ["sShift", "Mix_Solid",  "tspend",  "sortCount",  "ispeed"])
      df_s = pd.concat([df_s, w_df], ignore_index= True)

    else:
      for iLoop1 in range(df_s.shape[0]):
        df_s.iat[iLoop1,1] = 'Mx_Sort'

    df_sum = df_s

  print(df_sum)
  return df_sum


def set_per_date(flg_RS, df):
  if flg_RS == 'R':
    df_rmx= df[df['Mix_Solid'] == 'MO']
    df_rmx = calc_speed_2(df_rmx, ['sOperationDate', 'sShift', 'Mix_Solid', 'sortCount','tspend'], ['sOperationDate','sShift','Mix_Solid'])
    for iLoop1 in range(df_rmx.shape[0]):
      df_rmx.iat[iLoop1,2] = 'Mx_Repl'

    df_rmx2= df[df['Mix_Solid'] == 'SO']
    df_rmx2 = calc_speed_2(df_rmx2, ['sOperationDate', 'sShift', 'Mix_Solid', 'sortCount','tspend'], ['sOperationDate','sShift','Mix_Solid'])

    for iLoop1 in range(df_rmx2.shape[0]):
      df_rmx2.iat[iLoop1,2] = 'So_Repl'

    df_w = pd.concat([df_rmx, df_rmx2], ignore_index=True)

  else:
    df_s = calc_speed_2(df, ['sOperationDate', 'sShift', 'Mix_Solid', 'sortCount','tspend'], ['sOperationDate','sShift','Mix_Solid'])

    for iLoop1 in range(df_s.shape[0]):
      df_s.iat[iLoop1,2] = 'Mx_Sort'

    df_w = df_s

  print (df_w.head())

  df_w['sDate2'] = df_w['sShift'].astype(str) + '_' + df_w['sOperationDate'].astype(str)

  df_w = df_w[['sDate2','sShift','Mix_Solid','sortCount','ispeed']]
  df_w = df_w.sort_values(['sDate2','Mix_Solid'])

  print(df_w.head())

  return df_w


def set_Mix_Solid(iCount):
  if iCount > 1:
    sRtn = 'MO'
  else:
    sRtn = 'SO'
  return sRtn


def read_csv_file(file_name, read_columns, set_column_name, duplicate_column):
  
  df = pd.read_csv(file_name, usecols=read_columns)
  df.columns = set_column_name
  #print(df_sort0.head())

  print(df.duplicated().sum())
  df = df.drop_duplicates(subset=duplicate_column, keep='first')
  #print(df.shape[0])

  return df

### Common  End ####################################################################################################################

### 102 Data ###################################################################################################################
def set_102_0():

  df_0 = pd.read_csv(FILE_102, usecols=[1,3,8,21])
  df_0.columns = ['Container','Unit','Module','Mix_Solid']
  print(df_0.head())
  print(df_0.shape[0])

  return df_0

### Sort & Replenishment Common ################################################################################################

def set_com1(df, iOpeCode_Colomn):

  df = df.assign(sOperationDate = '', tOperationDate = '', tOperationDateTime = '', sShift = '', OPCD2 = '')
  print(df.head())

  for iLoop1 in range(df.shape[0]):
    sShift = ''
    sOpeCD2 = ''
    tOperationDateTime = ''
#    print(df.iat[iLoop1,0])
    sOperationDate0 = df.iat[iLoop1,0].astype(str)

    sOpeCD = df.iat[iLoop1, iOpeCode_Colomn]
    if sOpeCD is np.nan:
       sOpeCD = ""

    tOperationDateTime = dt.date(int(sOperationDate0[0:4]),int(sOperationDate0[4:6]),int(sOperationDate0[6:8]))

    tOperationDate1_w1 = dt.date(int(sOperationDate0[0:4]),int(sOperationDate0[4:6]),int(sOperationDate0[6:8]))
    tOperationTime = dt.time(int(sOperationDate0[8:10]),int(sOperationDate0[10:12]))
    if tOperationTime > dt.time(5,0):
      #process time between 0am to 5am change to previous day's operation 
      sOperationDate1 = sOperationDate0[0:8] 
    else:      
      tOperationDate1_w2 = tOperationDate1_w1 + dt.timedelta(days=-1)
      sOperationDate1 = tOperationDate1_w2.strftime('%Y%m%d')

    if tOperationTime > dt.time(5,00) and tOperationTime < dt.time(16,30):
      #1st Shift
      sShift = 'shift1'
    else:      
      #2nd Shift
      sShift = 'shift2' 

    tDateTime_W = sOperationDate0[0:4] + '/' + sOperationDate0[4:6] + '/' + sOperationDate0[6:8] + " " + sOperationDate0[8:10] + ':' + sOperationDate0[10:12]
    tOperationDateTime = dt.datetime.strptime(tDateTime_W, '%Y/%m/%d %H:%M')

    sOpeCD2 = sShift + '_' + sOpeCD

    iRow = 6
    df.iat[iLoop1,iRow] = sOperationDate1
    df.iat[iLoop1,iRow + 1] = dt.date(int(sOperationDate1[0:4]),int(sOperationDate1[4:6]),int(sOperationDate1[6:8]))
    df.iat[iLoop1,iRow + 2] = tOperationDateTime
    df.iat[iLoop1,iRow + 3] = sShift
    df.iat[iLoop1,iRow + 4] = sOpeCD2

  return df


def set_sum1(df, sSortDel_flg):
  iG = 0
  iSortCount = 0

  if sSortDel_flg == 'S':
    df1 = df.sort_values(['OPCD2', 'tOperationDateTime'])
  else:
    df1 = df[df['Spend'] >= 0]
    df1 = df1.sort_values(['Mix_Solid','OPCD2', 'tOperationDateTime'])

  df1 = df1.assign(tdiff=0, iG_1='', sortCount=0)
  print(df1.head())

#  ope_Flg = df1.iat[0,10]
  for iLoop2 in range(df1.shape[0]):
    if iLoop2 == 0:
      tdiff = 0
      iG = 1
    else:
      tdiff = df1.iat[iLoop2,8] - df1.iat[iLoop2 -1, 8]
      tdiff = tdiff.total_seconds() / 60

    if abs(tdiff) >= SORT_INTERVAL:
      iSortCount = 0
      iG += 1

    iSortCount += 1
    sOpe = df1.iloc[iLoop2,10]
    iColumn = 12
    df1.iat[iLoop2,iColumn + 0] = tdiff
    df1.iat[iLoop2,iColumn + 1] = sOpe + '_' + str(iG)
    df1.iat[iLoop2,iColumn + 2] = iSortCount

#  print(df1.head())  
  return df1


def set_sum2(df, sSD):

  df_sum = df[['sOperationDate','tOperationDateTime', 'sShift', 'Mix_Solid','OPCD', 'OPCD2', 'iG_1']]\
    .groupby(['sOperationDate', 'sShift', 'Mix_Solid','OPCD', 'OPCD2', 'iG_1'])\
    .agg(
      tDate1=('tOperationDateTime', 'min'),
      tDate2=('tOperationDateTime', 'max'),
      sortCount=('tOperationDateTime', 'count')
    )

  df_sum = df_sum.reset_index()

  if sSD == 'Sort':
    iSec = 0.7
  else:
    iSec = 0.2

  df1 = df_sum.assign(tspend=1,ispeed=0)
  if df1.shape[0] > 0:
    df1[['tspend', 'ispeed']] = df1.apply(lambda x: calc_time(x['tDate1'], x['tDate2'], x['sortCount'], iSec), axis=1, result_type="expand")

  #print(df1.head())  
  return df1


def set_sum3(flg_sr, df):
  df_sum = calc_speed_2(df, ['sShift', 'Mix_Solid','OPCD', 'OPCD2', 'sortCount','tspend'],  ['sShift', 'Mix_Solid','OPCD', 'OPCD2'])

  if flg_sr == 'S':
    df_sum = df_sum.rename(columns={'sortCount': 'SortCount'})
  else:
    df_sum = df_sum.rename(columns={'sortCount': 'ReplCount'})

#  print(df_sum.head())
  return df_sum


def set_mix_ope1(df_sort, df_repl):
  df_merge = pd.merge(df_sort, df_repl, on=['OPCD2'], how='outer')
  df_merge.columns = ['OPCD','Sort_Count','Sort_spend','Sort_Speed','Repl_Count','Repl_spend','Repl_Speed']
  df_merge = df_merge.sort_values(['OPCD'])
  df_merge = df_merge.fillna(0)

  df_merge = df_merge.assign(total_Count=0)
  for iLoop1 in range(df_merge.shape[0]):
    df_merge.iat[iLoop1,7] = df_merge.iat[iLoop1,1].astype(int) + df_merge.iat[iLoop1,3].astype(int)

  #print(df_merge.head())
  return df_merge


def set_top5(df, key1, key2, sortBy):
  df_1 = df[df[key1].str.contains(key2)]
  df_1 = df_1.sort_values(by = sortBy, ascending=False).head(5)
  print(df_1)
  return df_1

def set_average(df, sMS, sShift):
  print(df)

  if sMS == 'Mx':
    sort_count = df['Sort_Count'].mean()
    sort_spend = df['Sort_spend'].mean()
    sort_speed = round((sort_spend/sort_count) * 60, 1)

    repl_count = df['Repl_Count'].mean()
    repl_spend = df['Repl_spend'].mean()
    repl_speed = round((repl_spend/repl_count) * 60, 1)

    df_add = pd.DataFrame({'OPCD': [sShift + '_Average'], 
                          'Sort_Count':[sort_count],
                          'Sort_spend':[sort_spend],
                          'Sort_Speed':[sort_speed],
                          'Repl_Count':[repl_count],
                          'Repl_spend':[repl_spend],
                          'Repl_Speed':[repl_speed],
                          'total_Count':[0]})

  else:
    repl_count = df['ReplCount'].mean()
    repl_spend = df['tspend'].mean()
    repl_speed = round((repl_spend/repl_count) * 60, 1)

    df_add = pd.DataFrame({'OPCD2': [sShift + '_Average'], 
                          'ReplCount':[repl_count],
                          'tspend':[repl_spend],
                          'ispeed':[repl_speed]})

  print(df_add)
  
  df_1 = pd.concat([df, df_add])
  print(df_1)

  return df_1

### Sort&Replenishment Common End   #################################################################################################


### Functions End ####################################################################################################################

fig = plt.figure(figsize=(18, 11))
fig.suptitle(FIG_TITLE, fontsize=20)
fig.subplots_adjust(hspace=0.7, wspace=0.5, top=0.92, bottom=0.08, left=0.1)
ax_w = 18
ax_h = 3

############# Read 102 #####################################################################################################
df_102 = set_102_0()

############# <delivering> Start #####################################################################################################
df_deli0 = read_csv_file(FILE_KPI_PUTAWAY, [0,1,2,3,6,10], ['Stat_Time','End_Time','Spend','OPCD','Container','Module'], ['Module'])
print(df_deli0.shape[0])
if df_deli0.shape[0] <= 0:
  iFlgDeli = 0
else:
  iFlgDeli = 1

df_deli1 = set_com1(df_deli0, 3)
df_deli1 = df_deli1[df_deli1['sOperationDate'].isin(TargetDate2)]

df_102_w = df_102[['Module','Mix_Solid']]
df_deli2 = pd.merge(df_deli1, df_102_w, on=['Module'], how='left')

df_deli3 = df_deli2.copy(deep=True)
df_deli3 = set_sum1(df_deli3, 'R')

df_deli4 = set_sum2(df_deli3, 'Repl')
df_deli5 = set_sum3('R', df_deli4)

df_deli_shift1 = set_per_shift('R', df_deli4)
save_bi_csv(df_deli_shift1, "df_deli_shift1.csv")

############# <delivering> End #####################################################################################################


############# <Sorting> Start #####################################################################################################
def build_df_sort_day1_from_dates(dates_yyyymmdd: list[str]) -> pd.DataFrame:
    rows = []
    for d in dates_yyyymmdd:
        rows.append([f"shift1_{d}", "shift1", "Mx_Sort", 0, 0])
        rows.append([f"shift2_{d}", "shift2", "Mx_Sort", 0, 0])

    df = pd.DataFrame(
        rows,
        columns=["sDate2", "sShift", "Mix_Solid", "sortCount", "ispeed"]
    )
    return df


if iSortKPI_Available == 1:
  #SortKPI is Aveilable

  df_sort0 = read_csv_file(FILE_KPI_SORT, [0,1,2,5,6,7], ['Stat_Time','End_Time','OPCD','Step','Unit','Module'], ['Module'])
  print('Sorting_count', df_sort0.shape[0])
  if df_sort0.shape[0] <= 0:
    iFlgSort = 0

    df_sort_shift1 = pd.DataFrame([['shift1', 'Mx_Sort',0,0,0], ['shift2', 'Mx_Sort',0,0,0]], columns=['sShift','Mix_Solid','tspend','sortCount','ispeed'])
    print(df_sort_shift1)
    save_bi_csv(df_sort_shift1, 'df_sort_shift1.csv')

    df_sort_day1 = build_df_sort_day1_from_dates(TargetDate2)

    df_sort5 = pd.DataFrame([['shift1', 'MO', 'XXX','XXX', 0,0,0], ['shift2', 'MO', 'XXX','XXX', 0,0,0]],
                             columns=['sShift', 'Mix_Solid', 'OPCD', 'OPCD2', 'tspend', 'SortCount', 'ispeed'])

  else:
    iFlgSort = 1

    df_sort1 = set_com1(df_sort0, 2)
    df_sort1 = df_sort1[df_sort1['sOperationDate'].isin(TargetDate2)]

    df_102_w = df_102[['Module','Mix_Solid']]
    df_sort2 = pd.merge(df_sort1, df_102_w, on=['Module'], how='left')

    # Exclude Solid modules
    df_sort2 = df_sort2[df_sort2['Mix_Solid'] == 'MO']

    df_sort3 = df_sort2.copy(deep=True)
    df_sort3 = set_sum1(df_sort3, 'S')

    df_sort4 = set_sum2(df_sort3, 'Sort')

    df_sort5 = set_sum3('S', df_sort4)

    df_sort_shift1 = set_per_shift('S', df_sort4)
    save_bi_csv(df_sort_shift1, 'df_sort_shift1.csv')

    df_sort_day1 = set_per_date('s', df_sort4)
    save_bi_csv(df_sort_day1, 'df_sort_day1.csv')

else:
    iFlgSort = 0

    df_sort_shift1 = pd.DataFrame([['shift1', 'Mx_Sort',0,0,0], ['shift2', 'Mx_Sort',0,0,0]], columns=['sShift','Mix_Solid','tspend','sortCount','ispeed'])
    print(df_sort_shift1)
    save_bi_csv(df_sort_shift1, 'df_sort_shift1.csv')

    df_sort_day1 = build_df_sort_day1_from_dates(TargetDate2)

    print(df_sort_day1)

    df_sort5 = pd.DataFrame([['shift1', 'MO', 'XXX','XXX', 0,0,0], ['shift2', 'MO', 'XXX','XXX', 0,0,0]], columns=['sShift', 'Mix_Solid',	'OPCD',	'OPCD2',	'tspend',	'SortCount',	'ispeed'])
    print(df_sort5)
    save_bi_csv(df_sort5, 'df_sort5.csv')


############# <Sorting> End #####################################################################################################

df_g1 = pd.concat([df_deli_shift1, df_sort_shift1],ignore_index=True)

df_g1_pv = pd.pivot_table(df_g1, index='sShift', columns='Mix_Solid', values=['sortCount', 'ispeed'], aggfunc=np.max)
print(df_g1_pv)

_col_rename = {
    ('ispeed',     'Mx_Repl'): 'Mix_Repl_Speed',
    ('ispeed',     'Mx_Sort'): 'Mix_Sort_Speed',
    ('ispeed',     'So_Repl'): 'Solid_Repl_Speed',
    ('sortCount',  'Mx_Repl'): 'Mix_Repl_Count',
    ('sortCount',  'Mx_Sort'): 'Mix_Sort_Count',
    ('sortCount',  'So_Repl'): 'Solid_Repl_Count',
}
df_g1_pv.columns = [_col_rename.get(col, '_'.join(col)) for col in df_g1_pv.columns]
df_g1_pv = df_g1_pv.reset_index()
print(df_g1_pv)

df_g1_pv = df_g1_pv.reindex(columns=['sShift','Mix_Sort_Count','Mix_Sort_Speed','Mix_Repl_Count','Mix_Repl_Speed','Solid_Repl_Count','Solid_Repl_Speed'])
df_g1_pv = df_g1_pv.fillna(0)
print(df_g1_pv)
ax1 = fig.add_subplot(ax_h, ax_w, (1,5))
plot_3bar_3line(ax1, 0.28, df_g1, 
                'sShift', 'sortCount','ispeed', \
                "Sorting & Put away per Shift", \
                'Shift', 'Module count', 'process speed per module (seconds)')


############# Per day Start #####################################################################################################

df_deli_day1 = set_per_date('R', df_deli4)
save_bi_csv(df_deli_day1, 'df_deli_day1.csv')
print(df_deli_day1)


df_per_day = pd.concat([df_deli_day1, df_sort_day1],ignore_index=True)
save_bi_csv(df_per_day, 'df_per_day.csv')
print(df_per_day)


####################
df_day_pv1 = df_deli_day1[df_deli_day1['Mix_Solid']=='So_Repl']
df_day_pv1 =df_day_pv1[['sDate2','sShift','sortCount','ispeed']]
df_day_pv1.columns=['sDate2','shift1','Solid_Repl_Count','Solid_Repl_Speed']
print(df_day_pv1)

df_day_pv2 = df_deli_day1[df_deli_day1['Mix_Solid']=='Mx_Repl']
df_day_pv2 =df_day_pv2[['sDate2','sShift','sortCount','ispeed']]
df_day_pv2.columns=['sDate2','shift1','Mix_Repl_Count','Mix_Repl_Speed']
print(df_day_pv2)

df_day_pv3 = df_sort_day1[df_sort_day1['Mix_Solid']=='Mx_Sort']
df_day_pv3 =df_day_pv3[['sDate2','sShift','sortCount','ispeed']]
df_day_pv3.columns=['sDate2','shift1','Mix_Sort_Count','Mix_Sort_Speed']
print(df_day_pv3)


df_day_pv4 = pd.merge(df_day_pv1,df_day_pv2,on=['sDate2','shift1'],how='outer')
df_day_pv4 = pd.merge(df_day_pv4,df_day_pv3,on=['sDate2','shift1'],how='outer')
df_day_pv4 = df_day_pv4.fillna(0)
df_day_pv = pd.pivot_table(df_per_day, index='sDate2', columns='Mix_Solid', values=['sortCount', 'ispeed'], aggfunc=np.max)

_day_col_rename = {
    ('sortCount', 'Mx_Sort'): 'Mix_Sort_Count',
    ('sortCount', 'Mx_Repl'): 'Mix_Repl_Count',
    ('sortCount', 'So_Repl'): 'Solid_Repl_Count',
    ('ispeed',    'Mx_Sort'): 'Mix_Sort_Speed',
    ('ispeed',    'Mx_Repl'): 'Mix_Repl_Speed',
    ('ispeed',    'So_Repl'): 'Solid_Repl_Speed',
}
df_day_pv.columns = [_day_col_rename.get(col, '_'.join(col)) for col in df_day_pv.columns]
df_day_pv = df_day_pv.reset_index().rename(columns={'sDate2': 'Date2'})
print(df_day_pv)

df_day_pv =df_day_pv.reindex(columns=['Date2','Mix_Sort_Count','Mix_Sort_Speed','Mix_Repl_Count','Mix_Repl_Speed','Solid_Repl_Count','Solid_Repl_Speed'])
df_day_pv = df_day_pv.fillna(0)
print(df_day_pv)
save_bi_csv(df_day_pv, 'df_day_pv.csv')


ax1 = fig.add_subplot(ax_h, ax_w, (19,25))

plot_4bar_4line(ax1, 
                df_day_pv[df_day_pv['Date2'].str.contains('shift1_')],
                df_day_pv[df_day_pv['Date2'].str.contains('shift2_')],
                'Date2', 'Mix_Sort_Count','Mix_Repl_Count', 'Mix_Sort_Speed', 'Mix_Repl_Speed',\
                "Mix sort and Put away per day", \
                'Shift', 'Module count', 'Speed per module (seconds)', 0, PINK, GREEN, BLUE, RED, 4.5)


ax1 = fig.add_subplot(ax_h, ax_w, (30,34))

plot_2bar_2line(ax1,
                df_day_pv[df_day_pv['Date2'].str.contains('shift1_')],
                df_day_pv[df_day_pv['Date2'].str.contains('shift2_')],
                'Date2', 'Solid_Repl_Count','Solid_Repl_Speed',
                "Solid Put away per day", \
                'Shift', 'Module count', 'Repl_Speed per Module (seconds)',
                GREEN, 4.5)

############# Per day End #####################################################################################################

############# Per Operater Start #####################################################################################################
### Mix sort & put away ###

df_deli_w = df_deli5[df_deli5['Mix_Solid'] == 'MO']
df_deli_w = df_deli_w[['OPCD2','ReplCount','tspend','ispeed']]

df_mix_Ope = set_mix_ope1(df_sort5[['OPCD2','SortCount','tspend','ispeed']],df_deli_w)
df_shift1_top = set_top5(df_mix_Ope, 'OPCD', 'shift1_', 'total_Count')
df_shift1_ope_average = set_average(df_shift1_top,'Mx','shift1')
save_bi_csv(df_shift1_ope_average, 'df_shift1_ope_Mix_average.csv')

df_shift2_top = set_top5(df_mix_Ope, 'OPCD', 'shift2_', 'total_Count')
df_shift2_ope_average = set_average(df_shift2_top,'Mx','shift2')
save_bi_csv(df_shift2_ope_average, 'df_shift2_ope_Mix_average.csv')


ax1 = fig.add_subplot(ax_h, ax_w, (37,43))

plot_4bar_4line(ax1, 
                df_shift1_ope_average,
                df_shift2_ope_average,
                'OPCD', 'Sort_Count','Repl_Count', 'Sort_Speed', 'Repl_Speed',\
                "Mix sort and Put away per Operater", \
                'Operater', 'Module count', 'Speed per module (seconds)', 0, PINK, GREEN, BLUE, RED, 5.5)


### Solid put away ###
df_solid_Repl_Ope = df_deli5[df_deli5['Mix_Solid'] == 'SO']
df_solid_Repl_Ope = df_solid_Repl_Ope[['OPCD2','ReplCount','tspend','ispeed']]


ax1 = fig.add_subplot(ax_h, ax_w, (48,52))

df_shift1_top = set_top5(df_solid_Repl_Ope, 'OPCD2', 'shift1_', 'ReplCount')
df_shift1_ope_average = set_average(df_shift1_top,'SO','shift1')
save_bi_csv(df_shift1_ope_average, 'df_shift1_ope_Solid_average.csv')

df_shift2_top = set_top5(df_solid_Repl_Ope, 'OPCD2', 'shift2_', 'ReplCount')
df_shift2_ope_average = set_average(df_shift2_top,'SO','shift2')
save_bi_csv(df_shift2_ope_average, 'df_shift2_ope_Solid_average.csv')

plot_2bar_2line(ax1,
                df_shift1_ope_average,
                df_shift2_ope_average,
                'OPCD2', 'ReplCount','ispeed',
                "Solid Put away per Operater", \
                'Operater', 'Module count', 'Repl_Speed per Module (seconds)',
                GREEN, 5.5)

############# Per Operater End #####################################################################################################

HTML_TMP = """
<!doctype html>
<html lang="ja">
    <head>
        
        </head>
  <body>
    <img src="data:image/png;base64,{image_bin}">
  </body>
</html>
"""

html = img2html(fig)
plt.close()

with open(HTML_PATH + HTML_NAME, "w") as w:
    w.write(html)
