import pandas as pd
import datetime as dt
import numpy as np
import matplotlib.pyplot as plt
import io
import base64
import sys
import json

kpifilePath = './IP Picking/Data/'
kpiHTMLPath = './IP Picking/Html/'
workPath = './IP Picking/work9/'

kpiCSV = 'Picking_MODULE.csv'

# Config injection: when called from IP_Picking_KPI_auto.py, config is passed as JSON via sys.argv[1]
if len(sys.argv) > 1:
    _cfg = json.loads(sys.argv[1])
    TargeDate = _cfg['TargeDate']
    TargeDate1 = _cfg['TargeDate1']
    FIG_TITLE = _cfg['FIG_TITLE']
    OrderGroups = _cfg['OrderGroups']
else:
    # Fallback: original hardcoded values for manual use
    TargeDate = ['20260330','20260331','20260401','20260402','20260403']
    TargeDate1 = ['20260330','20260331','20260401','20260402','20260403','20260404']
    FIG_TITLE = 'IP Picking KPI, Pallet code-S3 [Order week 03/23/2026]'
PICK_INTERVAL = 20
BASE_TIME = 75
#G_H = 11
#G_W = 5

blue = 'dodgerblue'
#green = 'lightseagreen'
green = 'darkseagreen'
red = 'red'
pink = 'lightpink'
#palevioletred
#orange = 'orange'
orange = 'darkorange'

#PCodes_list = ['A1', 'A2', 'A3', 'B1', 'BX', 'C', 'D1', 'D2', 'E', 'F', 'G','H','J','K','U']
#PCodes_list = ['A1', 'A2', 'A3', 'B1', 'BX', 'C', 'C2', 'D1', 'D2', 'E', 'F', 'G','HJK','U']
PCodes_list = ['A1', 'A2', 'A3', 'B1', 'C', 'D1', 'D2', 'E1', 'E2', 'F', 'G', 'HJK','U']

PCodes = [['A1'],
          ['A2'],
          ['A3','AX'],
          ['B1','B2','BX'],
          ['C1','C2','C3','CX'],
          ['D1'],
          ['D2','D3','DX'],
          ['E1'],
          ['E2','E3','EX'],
          ['F1','F2','FX'],
          ['G1'],
          ['H1','HX','JX','K1','K2','KX'],
          ['U1','UX']]

PCodes2 = ['A1',
          'A2',
          'A3,AX',
          'B1,B2,BX',
          'C1,C2,C3,CX',
          'D1',
          'D2,D3,DX',
          'E1',
          'E2,E3,EX',
          'F1,F2,FX',
          'G1',
          'H1,HX,JX,K1,K2,KX',
          'U1,UX']


PCodes_title = ['A1[A1]',
          'A2[A2]',
          'A3[A3,AX]',
          'B1[B1,B2,BX]',
          'C[C1,C2,C3,CX]',
          'D1[D1]',
          'D2[D2,D3,DX]',
          'E1[E1]',
          'E2[E2,E3,EX]',
          'F[F1,F2,FX]',
          'G[G1]',
          'HJK[H1,HX,JX,K1,K2,KX]',
          'U[U1,UX]']
#          'C2[C2]',


#OrderNos = [['01','02','03'],['04','05','06'],['07','08','09'],['10','11','12'],
#           ['13','14','15'],['16','17','18'],['19','20','21'],['22','23','24'],
#           ['25','26','27'],['28','29','30']]

#OrderNos2 = ['010203','040506','070809','101112', '131415','161718','192021','222324','252627','282930','supplemnt/BO']

WeekDays = ['<Monday>','<Tuesday>','<Wednesday>','<Thursday>','<Friday>','<Saturday>']

# OrderGroups: only define fallback if not already set from config
if len(sys.argv) <= 1:
    OrderGroups = [
    ['20260330-01','2026033001','2026033002','2026033003'],
    ['20260330-02','2026033004','2026033005','2026033006'],
    ['20260330-03','2026033007','2026033008','2026033009'],
    ['20260330-04','2026033010','2026033011','2026033012'],
    ['20260330-05','2026033013','2026033014','2026033015'],
    ['20260330-06','2026033016','2026033017','2026033018'],
    ['20260330-07','2026033019','2026033020','2026033021'],
    ['20260330-08','2026033022','2026033023','2026033024'],
    ['20260330-09','2026033025','2026033026','2026033027'],
    ['20260330-10','2026033028','2026033029','2026033030'],
    ['20260330-11','2026033031','2026033032','2026033033'],
    ['20260330-12','2026033034','2026033035','2026033036'],
    ['20260331-01','2026033101','2026033102','2026033103'],
    ['20260331-02','2026033104','2026033105','2026033106'],
    ['20260331-03','2026033107','2026033108','2026033109'],
    ['20260331-04','2026033110','2026033111','2026033112'],
    ['20260331-05','2026033113','2026033114','2026033115'],
    ['20260331-06','2026033116','2026033117','2026033118'],
    ['20260331-07','2026033119','2026033120','2026033121'],
    ['20260331-08','2026033122','2026033123','2026033124'],
    ['20260331-09','2026033125','2026033126','2026033127'],
    ['20260331-10','2026033128','2026033129','2026033130'],
    ['20260331-11','2026033131','2026033132','2026033133'],
    ['20260331-12','2026033134','2026033135','2026033136'],
    ['20260401-01','2026040101','2026040102','2026040103'],
    ['20260401-02','2026040104','2026040105','2026040106'],
    ['20260401-03','2026040107','2026040108','2026040109'],
    ['20260401-04','2026040110','2026040111','2026040112'],
    ['20260401-05','2026040113','2026040114','2026040115'],
    ['20260401-06','2026040116','2026040117','2026040118'],
    ['20260401-07','2026040119','2026040120','2026040121'],
    ['20260401-08','2026040122','2026040123','2026040124'],
    ['20260401-09','2026040125','2026040126','2026040127'],
    ['20260401-10','2026040128','2026040129','2026040130'],
    ['20260401-11','2026040131','2026040132','2026040133'],
    ['20260401-12','2026040134','2026040135','2026040136'],
    ['20260402-01','2026040201','2026040202','2026040203'],
    ['20260402-02','2026040204','2026040205','2026040206'],
    ['20260402-03','2026040207','2026040208','2026040209'],
    ['20260402-04','2026040210','2026040211','2026040212'],
    ['20260402-05','2026040213','2026040214','2026040215'],
    ['20260402-06','2026040216','2026040217','2026040218'],
    ['20260402-07','2026040219','2026040220','2026040221'],
    ['20260402-08','2026040222','2026040223','2026040224'],
    ['20260402-09','2026040225','2026040226','2026040227'],
    ['20260402-10','2026040228','2026040229','2026040230'],
    ['20260402-11','2026040231','2026040232','2026040233'],
    ['20260402-12','2026040234','2026040235','2026040236'],
    ['20260403-01','2026040301','2026040302','2026040303'],
    ['20260403-02','2026040304','2026040305','2026040306'],
    ['20260403-03','2026040307','2026040308','2026040309'],
    ['20260403-04','2026040310','2026040311','2026040312'],
    ['20260403-05','2026040313','2026040314','2026040315'],
    ['20260403-06','2026040316','2026040317','2026040318'],
    ['20260403-07','2026040319','2026040320','2026040321'],
    ['20260403-08','2026040322','2026040323','2026040324'],
    ['20260403-09','2026040325','2026040326','2026040327'],
    ['20260403-10','2026040328','2026040329','2026040330'],
    ['20260403-11','2026040331','2026040332','2026040333'],
    ['20260403-12','2026040334','2026040335','2026040336'],
    ['20260404-01','2026040401','2026040402','2026040403'],
    ['20260404-02','2026040404','2026040405','2026040406'],
    ['20260404-03','2026040407','2026040408','2026040409'],
    ['20260404-04','2026040410','2026040411','2026040412'],
    ['20260404-05','2026040413','2026040414','2026040415'],
    ['20260404-06','2026040416','2026040417','2026040418'],
    ['20260404-07','2026040419','2026040420','2026040421'],
    ['20260404-08','2026040422','2026040423','2026040424'],
    ['20260404-09','2026040425','2026040426','2026040427'],
    ['20260404-10','2026040428','2026040429','2026040430'],
    ['20260404-11','2026040431','2026040432','2026040433'],
    ['20260404-12','2026040434','2026040435','2026040436']
    ]

def formatDate(sDate):
    #print(sDate)
    #print(type(sDate))
    tDateW1 = dt.date(int(sDate[0:4]),int(sDate[4:6]),int(sDate[6:8]))

    sTime = sDate[8:10] + ':' + sDate[10:12]
    tTime = dt.time(int(sDate[8:10]),int(sDate[10:12]))
    if tTime > dt.time(5,0):
      #process time between 0am to 5am change to previous day's operation 
      #1st Shift
      sDate1 = sDate[0:4] + '/' + sDate[4:6] + '/' + sDate[6:8] 
    else:      
      #2nd Shift
      tDateW2 = tDateW1 + dt.timedelta(days=-1)
      sDate1 = tDateW2.strftime('%Y/%m/%d')

    sDate2 = sDate[0:4] + '/' + sDate[4:6] + '/' + sDate[6:8] + " " + sDate[8:10] + ':' + sDate[10:12] 
    tDate = dt.datetime.strptime(sDate2, '%Y/%m/%d %H:%M')

    sShift = ''
    if tTime > dt.time(6,00) and tTime < dt.time(16,30):
      #1st Shift
      sShift1 = 'shift1' 
      sShift2 = '1' 
    else:      
      #2nd Shift
      sShift1 = 'shift2' 
      sShift2 = '2' 

    #print(type(tDate), tDate)
    return sDate1, sDate2, tDate, sShift1, sShift2 

def calc_time(timeFrom, timeTo, pick_count):

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
    tspend0 = 1
    ispeed = round((tspend0 / pick_count) * 60, 1)

  return tspend, ispeed


def set_orderNo(sOrderNo, sBackOrder):
#  OrderNo0 = sOrderNo + sBackOrder
  OrderNo0 = sOrderNo
  orderNo1 = sOrderNo[:8]
  orderNo2 = sOrderNo[-2:]
  orderNo3 = 'supplemnt/BO'

  iLoop = 0
  flg_sp = 1
  flg_bo = 0
  for order in OrderGroups:
#    if orderNo2 in OrderGroups[iLoop]:
    if sOrderNo in OrderGroups[iLoop]:
      if '*' in sBackOrder:
        flg_bo = 1
        OrderNo0 = sOrderNo + sBackOrder
      else:
        flg_bo = 0

#      orderNo3 = OrderNos2[iLoop]
      orderNo3 = OrderGroups[iLoop][0]
      flg_sp = 0
      break
    iLoop += 1

  if flg_sp == 1 or flg_bo ==1:
    orderNo2 = 'SP/BO'

  return OrderNo0, orderNo1, orderNo2, orderNo3


def set_pCode(spalletCode, orderNo2):

  if orderNo2 == 'SP/BO':
       sPCode1 = 'SP/BO'
       sPCode2 = 'supplemnt/BO'
  else:
    iLoop = 0
    sPCode1 = ""
    sPCode2 = ""
    for pp in PCodes:
      if spalletCode in PCodes[iLoop]:
        #if spalletCode == 'JX' or spalletCode == 'H1' or spalletCode == 'HX':
        if spalletCode == 'JX':
          i=1
        
        sPCode1 = PCodes_list[iLoop]
        sPCode2 = PCodes2[iLoop]
      iLoop += 1

  return sPCode1, sPCode2


def img2html(fig):
    sio = io.BytesIO()
    fig.savefig(sio, format='png')
    image_bin = base64.b64encode(sio.getvalue())
    return HTML_TMP.format(image_bin=str(image_bin)[2:-1])


### Glaph Plot Start ####################################################################################################################

def set_graph_label1(ax1, stitle, x_label1_flg, x_label, y_label1_flg, y_label1, xtickrotation):
    ax1.set_title(stitle)
    ax1.grid(axis="y")
    if x_label1_flg == 1:
        ax1.set_ylabel(x_label)

    if y_label1_flg == 1:
        ax1.set_ylabel(y_label1)

    ax1.set_xticks(ax1.get_xticks())
    xticklabels = ax1.get_xticklabels()
    ax1.set_xticklabels(xticklabels, fontsize=12, rotation=xtickrotation)


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


def plot_title(fig, ax,sRoute):
    ax_pos = ax.get_position()
    fig.text(0.02, ax_pos.y0 +0.03, sRoute, size = 28)


def plot_1bar(ax, df_A, x_key, y_key1, stitle, x_label, y_label1):
    #width = 0.35
    ax.grid(axis="y")

    pA1 = ax.bar(df_A[x_key], df_A[y_key1], label=y_key1, alpha=0.5, color=blue) 
  
    ax.bar_label(pA1,label_type = 'edge')
    set_graph_label1(ax, stitle, 0, "", 1, y_label1, 270)

    ax.grid(axis="y")
    plt.legend(bbox_to_anchor=(1.02, 1), loc='upper left')


def plot_plot(ax, df, x_key, stitle, x_label, y_label):
  df_T = df.T
  print(df_T)
  ax.grid(axis="y")

  for iLoop in range(df_T.shape[1]):
    pA1 = ax.plot(['Shift1', 'Shift2'], df_T.iloc[:, iLoop], marker='o')

  ymin, ymax = plt.ylim()
  plt.ylim(ymin - 5, ymax + 5)

  ax.legend(df.index, bbox_to_anchor=(1.02, 1), loc='upper left')

  set_graph_label1(ax, stitle, 1, x_label, 1, y_label, 0)
  ax.set_ylim(30, 80)
  ax.grid(axis="y")


def plot_2bar_1line(ax, df_AB, df_A, df_B, x_key, y_key1, y_key2, stitle, x_label, y_label1, y_label2):
    width = 0.3
    ax2 = ax.twinx()
    ax.grid(axis="y")

    pA1 = ax.bar(df_A[x_key], df_A[y_key1], label='Shift1_pick_box',  alpha=0.5, color=green) 
    ax.bar_label(pA1,label_type = 'edge')
#    pA2 = ax.bar(df_A[x_key], df_A[y_key2], label='Shift1_pick_time', hatch='/', alpha=0.5, align="edge", width=width, color=blue) 
#    ax.bar_label(pA2,label_type = 'edge')

    pB1 = ax.bar(df_B[x_key], df_B[y_key1], label='Shift2_pick_box',  alpha=0.5, color=orange)
    ax.bar_label(pB1,label_type = 'edge')
#    pB2 = ax.bar(df_B[x_key], df_B[y_key2], label='Shift2_pick_time', hatch='/', alpha=0.5, align="edge", width=width, color=orange) 
#    ax.bar_label(pB2,label_type = 'edge')

    pA3 = ax2.plot(df_AB[x_key], df_AB['ispeed'], color="blue", marker='o')

    list_DayClass = df_AB[x_key].values.tolist()
    list_ispeed = df_AB['ispeed'].values.tolist()

    for iLoop, value in enumerate(list_ispeed):
        ax2.text(list_DayClass[iLoop], list_ispeed[iLoop]-5, value)

    ymin, ymax = plt.ylim()
    plt.ylim(ymin - 5, ymax + 5)

    ax2.set_ylim(30, 80)

    #ax.legend([pA1, pA2, pB1, pB2],{'Shift1_pick_box','Shift1_pick_time','Shift2_pick_box','Shift2_pick_time'}, bbox_to_anchor=(1.02, 1), loc='upper left')

    set_graph_label2(ax, ax2, stitle, 1, x_label, 1, y_label1, 1, y_label2, 0, 12)

    ax.grid(axis="y")


def plot_1bar_1line(ax, df_A, x_key, y_key1, y_key2, stitle, x_label, y_label1, y_label2):
#  width = 0.3
  ax2 = ax.twinx()
  ax.grid(axis="y")

  pA1 = ax.bar(df_A[x_key], df_A[y_key1], alpha=0.5,  label='pick Module Count', color=blue) 
  ax.bar_label(pA1,label_type = 'edge')

  pPlotA = ax2.plot(df_A[x_key],df_A[y_key2], label='pick speed', color=red, marker='o') 

  set_graph_label2(ax, ax2, stitle, 1, x_label, 1, y_label1, 1, y_label2, 270, 8)

  # text shows in plot
  list_DayClass = df_A[x_key].values.tolist()
  list_ispeed = df_A['ispeed'].values.tolist()
  for iLoop, value in enumerate(list_ispeed):
    ax2.text(list_DayClass[iLoop], list_ispeed[iLoop]-5, value)

  ymin, ymax = plt.ylim()
  plt.ylim(ymin - 5, ymax + 30)

  hans, labs = ax.get_legend_handles_labels()
  ax.legend(handles=hans, labels=labs, bbox_to_anchor=(1.07, 0.9), loc='upper left', borderaxespad=0)

  ax2.legend(loc="center left", bbox_to_anchor=(1.07, 0.8))
  ax.grid(axis="y")


def plot_1bar_1line_1(ax, df_A, x_key, y_key1, y_key2, stitle, x_label, y_label1, y_label2):
  width = 0.3
  ax2 = ax.twinx()
  ax.grid(axis="y")
  print(df_A.head())

  pA1 = ax.bar(df_A[x_key], df_A[y_key1], alpha=0.5, color=blue) 
  ax.bar_label(pA1,label_type = 'edge')

  pPlotA = ax2.plot(df_A[x_key],df_A[y_key2], color="blue", label='shift1_speed', marker='o') 

  set_graph_label2(ax, ax2, stitle, 1, x_label, 1, y_label1, 1, y_label2, 270, 8)

  # text shows in plot
  list_DayClass = df_A[x_key].values.tolist()
  list_ispeed = df_A['ispeed'].values.tolist()
  for iLoop, value in enumerate(list_ispeed):
    ax2.text(list_DayClass[iLoop], list_ispeed[iLoop]-5, value)

  ymin, ymax = plt.ylim()
  plt.ylim(ymin - 5, ymax + 30)
  
  ax.axhline(BASE_TIME, 0, 1, c="purple", ls='--')
  
  ax.grid(axis="y")

### Glaph Plot End ####################################################################################################################

### Functions Start ####################################################################################################################

def df_stuck1(df1, key, column):
  df1_w = df1[df1['Shift1'] == 'shift1']
  df1_w = df1_w[[key, column]]
  df1_w.columns = [key, 'Shift1']

  df2_w = df1[df1['Shift1'] == 'shift2']
  df2_w = df2_w[[key, column]]
  df2_w.columns = [key, 'Shift2']

  df_all_w = pd.DataFrame()
  df_all_w = pd.merge(df1_w, df2_w)
  df_all_w.set_index(key, inplace=True)
  print(df_all_w)
  return df_all_w


def calc_speed_minmax(df, item_list1, group_list1, col_list1, item_list2, group_list2, col_list2,  item_list3, group_list3, col_list3):

  df_min = df[item_list1].groupby(group_list1).agg({'min'})
  df_min = df_min.reset_index()
  df_min = df_min.droplevel(1, axis=1)
  df_min.columns = col_list1

  df_max = df[item_list2].groupby(group_list2).agg({'max'})
  df_max = df_max.reset_index()
  df_max = df_max.droplevel(1, axis=1)
  df_max.columns = col_list2

  df_cnt = df[item_list3].groupby(group_list3).agg({'count'})
  df_cnt = df_cnt.reset_index()
  df_cnt = df_cnt.droplevel(1, axis=1)
  df_cnt.columns = col_list3

  df_sum = pd.DataFrame()
  df_sum = pd.merge(df_min, df_max)
  df_sum = pd.merge(df_sum, df_cnt)

  df_sum = df_sum.assign(tspend=1,ispeed=0)
  if df_sum.shape[0] > 0:
    df_sum[['tspend', 'ispeed']] = df_sum.apply(lambda x: calc_time(x['tDate1'], x['tDate2'], x['pickCount']), axis=1, result_type="expand")

  return df_sum


def calc_speed(df, item_list1, group_list1,  item_list2, group_list2):

  df_1 = df[item_list1].groupby(group_list1).agg({'sum'})
  df_1 = df_1.reset_index()
  df_1 = df_1.droplevel(1, axis=1)
  df_1.columns = item_list1

  df_2 = df[item_list2].groupby(group_list2).agg({'sum'})
  df_2 = df_2.reset_index()
  df_2 = df_2.droplevel(1, axis=1)
  df_2.columns = item_list2

  df_sum = pd.DataFrame()
  df_sum = pd.merge(df_1, df_2)

  df_sum = df_sum.assign(ispeed=0)
  if df_sum.shape[0] > 0:
    df_sum['ispeed'] = df_sum.apply(lambda x: round((x['tspend'] / x['pickCount']) * 60, 1), axis=1)

  #df_sum.to_csv('./work_internal/df_sum2.csv')
  
  return df_sum


#def set_sum1(df):
#  df_sum = calc_speed_minmax(df,['OrderNo3', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate'],
#                                ['OrderNo3', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2'], 
#                                ['OrderNo3', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate1'],
#
#                                ['OrderNo3', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate'], 
#                                ['OrderNo3', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2'],
#                                ['OrderNo3', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate2'],

#                                ['OrderNo3', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate'],
#                                ['OrderNo3', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2'],
#                                ['OrderNo3', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'pickCount'])
#  df_sum.to_csv('./work9/df_sum1.csv')
#  return df_sum


def set_sum1_PB(df):

  df_sum = calc_speed_minmax(df,['OrderNo3', 'OrderNo', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate'],
                                ['OrderNo3', 'OrderNo', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2'], 
                                ['OrderNo3', 'OrderNo', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate1'],

                                ['OrderNo3', 'OrderNo', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate'], 
                                ['OrderNo3', 'OrderNo', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2'],
                                ['OrderNo3', 'OrderNo', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate2'],

                                ['OrderNo3', 'OrderNo', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate'],
                                ['OrderNo3', 'OrderNo', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2'],
                                ['OrderNo3', 'OrderNo', 'iG_order1', 'iG_order2', 'iG_order3', 'OPCD', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'pickCount'])

  #df_sum.to_csv('./work9/df_sum1_PB.csv')
  return df_sum

def set_sum2(df):

  df_sum = calc_speed(df,['OrderNo3', 'iG_order3', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'pickCount'],
                          ['OrderNo3', 'iG_order3', 'pCode1', 'pCode2', 'Shift1', 'Shift2'], 
                          ['OrderNo3', 'iG_order3', 'pCode1', 'pCode2', 'Shift1', 'Shift2', 'tspend'], 
                          ['OrderNo3', 'iG_order3', 'pCode1', 'pCode2', 'Shift1', 'Shift2'])

  df_sum.to_csv(workPath + "df_sum2.csv")  
  return df_sum


def sum_per_shift(df):

  df_sum = calc_speed(df,['Shift1','pickCount'],['Shift1'], 
                     ['Shift1', 'tspend'], ['Shift1'])
  print(df_sum)

  df_sum.to_csv(workPath + "df_per_shift" + ".csv")
  return df_sum


def sum_per_shift_speed(df):

  df_0 = df[~df['pCode1'].str.contains('SP/BO')]

  df_sum = calc_speed(df_0,['Shift1', 'pCode1', 'pCode2', 'pickCount'],['Shift1','pCode1', 'pCode2'], 
                          ['Shift1', 'pCode1', 'pCode2','tspend'], ['Shift1', 'pCode1', 'pCode2'])

  df_sum = df_stuck1(df_sum, 'pCode1', 'ispeed')
  df_sum.to_csv(workPath + "df_per_shift_speed" + ".csv")
  return df_sum


def set_day_pcode_total(df):

#  df_1 = df[df['OrderNo2'] != 'ZZ']
  df_1 = df[['pCode2', 'Module']].groupby(['pCode2']).agg({'count'})
  df_1 = df_1.reset_index()
  df_1 = df_1.droplevel(1, axis=1)
  df_1.columns = ['pCode2', 'pickCount']
  print(df_1.head())

  df_1.to_csv(workPath + "set_day_pcode_total" + ".csv")
  return df_1


def set_ope0(df):
  df_1 = df[~df['pCode1'].str.contains('SP/BO')]
  print(df_1.head())
  df_1 = df_1[['tDate', 'OPCD', 'pCode1', 'OrderNo3','Module', 'Shift1']]

  df_1 = df_1.sort_values(['OPCD', 'tDate'])
  df_1 = df_1.assign(tdiff=0, iG_1='', iG_2='', PickCount=0)
#  df.to_csv('./work9/df_zzz0_' + str(pCode1) + '_' + str(shift) + '.csv')

  Ope_Flg = df_1.iat[0,1]
  pCode_Flg = df_1.iat[0,2]
  OrderNo1_Flg = df_1.iat[0,3]

#  iSpend = 0
  iG = 0
  iPickCount = 0
  tdiff = 0

  for iLoop2 in range(df_1.shape[0]):
    if iLoop2 == 0:
      tdiff = 0
      iG_order = 1
    else:
      tdiff = df_1.iat[iLoop2,0] - df_1.iat[iLoop2 -1, 0]
      tdiff = tdiff.total_seconds() / 60

    if abs(tdiff) > 1000:
      iii = 1

    if abs(tdiff) >= PICK_INTERVAL \
        or pCode_Flg != df_1.iat[iLoop2,2] \
        or Ope_Flg != df_1.iat[iLoop2,1] \
        or OrderNo1_Flg != df_1.iat[iLoop2,3]:
      
      if OrderNo1_Flg != df.iat[iLoop2,3]:
        iG_order = 1
      else:
        iG_order += 1

      
      iPickCount = 0
      #iG_order += 1
      pCode_Flg = df_1.iat[iLoop2,2]
      Ope_Flg = df_1.iat[iLoop2,1]
      OrderNo1_Flg = df_1.iat[iLoop2,3]

    iPickCount += 1
    sShift = df_1.iloc[iLoop2,5]
    #OrderNo = df.iloc[iLoop2,6]
    #OrderDegit = df.iloc[iLoop2,2]
    df_1.iat[iLoop2,6] = tdiff
    df_1.iat[iLoop2,7] = Ope_Flg + '_' + str(iG_order)  
    df_1.iat[iLoop2,8] = sShift + '_' + Ope_Flg  
    df_1.iat[iLoop2,9] = iPickCount

  print(df_1.head())

  df_1.to_csv(workPath + "df_ope0.csv")
  return df_1


def set_ope1(df):

  df_sum = calc_speed_minmax(df,['OPCD', 'tDate', 'iG_1', 'iG_2'], ['OPCD', 'iG_1', 'iG_2'], ['OPCD', 'iG_1', 'iG_2', 'tDate1'],
                                ['OPCD', 'tDate', 'iG_1', 'iG_2'], ['OPCD', 'iG_1', 'iG_2'], ['OPCD', 'iG_1', 'iG_2', 'tDate2'],
                                ['OPCD', 'tDate', 'iG_1', 'iG_2'], ['OPCD', 'iG_1', 'iG_2'], ['OPCD', 'iG_1', 'iG_2', 'pickCount'])

  df_sum.to_csv(workPath + "df_ope1.csv")
  return df_sum


def set_ope2(df):

  df_sum = calc_speed(df,['OPCD', 'iG_2', 'pickCount'],['OPCD','iG_2'], 
                         ['OPCD', 'iG_2', 'tspend'], ['OPCD','iG_2'])

  df_sum.to_csv(workPath + "df_ope2.csv")
  return df_sum


def set_week1(df):

  print(df.shape[0])
  df_1 = df[~df['pCode1'].str.contains('SP/BO')]
#  print(df_1.head())
  print(df_1.shape[0])

#  df_sum = calc_speed(df_1,['OrderNo3', 'pCode1', 'pCode2', 'pickCount'],['OrderNo3', 'pCode1', 'pCode2'], 
#                          ['OrderNo3',  'pCode1', 'pCode2', 'tspend'], ['OrderNo3', 'pCode1', 'pCode2'])
  df_sum = calc_speed(df_1,['OrderNo3', 'iG_order3', 'pCode1', 'pCode2', 'pickCount'],['OrderNo3', 'iG_order3','pCode1', 'pCode2'], 
                          ['OrderNo3', 'iG_order3', 'pCode1', 'pCode2', 'tspend'], ['OrderNo3', 'iG_order3', 'pCode1', 'pCode2'])

  df_sum.to_csv(workPath + "df_week1.csv")
  return df_sum


def set_pcode0(df):
  iG_order = 0
  iPickCount = 0

  #df = df.sort_values(['pCode1','OrderNo3','tDate'])
#  df = df.sort_values(['pCode1','OrderNo1','OPCD', 'tDate'])
  df = df.sort_values(['pCode1','OrderNo3','OPCD','tDate'])
  df = df.assign(tdiff=0, iG_order1='', iG_order2='', iG_order3='', PickCount=0)
#  df.to_csv('./work9/set_sum0_0.csv')

  print(df.head())
  pCode_Flg = df.iat[0,11]
  Ope_Flg = df.iat[0,2]
#  OrderNo1_Flg = df.iat[0,6]
  OrderNo1_Flg = df.iat[0,8]
  for iLoop2 in range(df.shape[0]):
    if iLoop2 == 0:
      tdiff = 0
      iG_order = 1
    else:
      tdiff = df.iat[iLoop2,1] - df.iat[iLoop2 -1, 1]
      tdiff = tdiff.total_seconds() / 60
      #print(df_order_w_shift1.iat[iLoop2,1],df_order_w_shift1.iat[iLoop2 -1,1],tdiff)

    if abs(tdiff) >= PICK_INTERVAL \
        or pCode_Flg != df.iat[iLoop2,11] \
        or Ope_Flg != df.iat[iLoop2,2] \
        or OrderNo1_Flg != df.iat[iLoop2,8]:
      
      if OrderNo1_Flg != df.iat[iLoop2,8]:
        iG_order = 1
      else:
        iG_order += 1
      
      iPickCount = 0
      #iG_order += 1
      pCode_Flg = df.iat[iLoop2,11]
      Ope_Flg = df.iat[iLoop2,2]
      OrderNo1_Flg = df.iat[iLoop2,8]

    iPickCount += 1
    #sShift = df.iloc[iLoop2,17]
    #OrderNo = df.iloc[iLoop2,6]
    OrderDegit = df.iloc[iLoop2,8]
    sShift = str(df.iloc[iLoop2,17])
    df.iat[iLoop2,18] = tdiff
    df.iat[iLoop2,19] = str(OrderDegit) + '_' + pCode_Flg + '_' + sShift + '_' + Ope_Flg
    df.iat[iLoop2,20] = str(OrderDegit) + '_' + pCode_Flg + '_' + str(iG_order)
    df.iat[iLoop2,21] = str(OrderDegit) + '_' + pCode_Flg
    df.iat[iLoop2,22] = iPickCount
    
  print(df.head())
  df.to_csv(workPath + "df_pcode0.csv")
  return df

def set_pcode1(df):

  df_sum = calc_speed_minmax(df,['OrderNo3','iG_order1','iG_order2','iG_order3','OPCD','pCode1', 'pCode2', 'Shift1', 'tDate'],
                                ['OrderNo3','iG_order1','iG_order2','iG_order3','OPCD','pCode1', 'pCode2', 'Shift1'], 
                                ['OrderNo3','iG_order1','iG_order2','iG_order3','OPCD','pCode1', 'pCode2', 'Shift1', 'tDate1'],

                                ['OrderNo3','iG_order1','iG_order2','iG_order3','OPCD','pCode1', 'pCode2', 'Shift1', 'tDate'], 
                                ['OrderNo3','iG_order1','iG_order2','iG_order3','OPCD','pCode1', 'pCode2', 'Shift1'],
                                ['OrderNo3','iG_order1','iG_order2','iG_order3','OPCD','pCode1', 'pCode2', 'Shift1', 'tDate2'],

                                ['OrderNo3','iG_order1','iG_order2','iG_order3','OPCD','pCode1', 'pCode2', 'Shift1','tDate'],
                                ['OrderNo3','iG_order1','iG_order2','iG_order3','OPCD','pCode1', 'pCode2', 'Shift1'],
                                ['OrderNo3','iG_order1','iG_order2','iG_order3','OPCD','pCode1', 'pCode2', 'Shift1', 'pickCount'])

#  df_sum = calc_speed_minmax(df,['OrderNo3','iG_order2','iG_order3','pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate'],
#                                ['OrderNo3','iG_order2','iG_order3','pCode1', 'pCode2', 'Shift1', 'Shift2'], 
#                                ['OrderNo3','iG_order2','iG_order3','pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate1'],

#                                ['OrderNo3','iG_order2','iG_order3','pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate'], 
#                                ['OrderNo3','iG_order2','iG_order3','pCode1', 'pCode2', 'Shift1', 'Shift2'],
#                                ['OrderNo3','iG_order2','iG_order3','pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate2'],

#                                ['OrderNo3','iG_order2','iG_order3','pCode1', 'pCode2', 'Shift1', 'Shift2', 'tDate'],
#                                ['OrderNo3','iG_order2','iG_order3','pCode1', 'pCode2', 'Shift1', 'Shift2'],
#                                ['OrderNo3','iG_order2','iG_order3','pCode1', 'pCode2', 'Shift1', 'Shift2', 'pickCount'])

  print(df_sum.head())
  df_sum.to_csv(workPath + "df_pcode1.csv")
  return df_sum


def set_pcode2(df):

  df_sum = calc_speed(df,['pCode1', 'pCode2', 'Shift1', 'pickCount'],
                          ['pCode1', 'pCode2', 'Shift1'], 
                          ['pCode1', 'pCode2', 'Shift1', 'tspend'], 
                          ['pCode1', 'pCode2', 'Shift1'])

  df_sum.to_csv(workPath + "df_pcode2.csv")  
  return df_sum


def formatDate2(sDate, sTime):
    tDate = dt.datetime.strptime(sDate, '%Y/%m/%d')
    tDateTime = dt.datetime.strptime(sDate + ' ' + sTime, '%Y/%m/%d %H:%M:%S')

    #print(type(tDate), tDate)
    return tDate, tDateTime


def set_IP202(df):
  print(df.head())
  df_1 = df[df['DOCK CODE'].str.contains('S3')]
  df_1['PLAN SHIP DATE'] = df_1['PLAN SHIP DATE'].str.replace('-','/')
  df_1['PLAN SHIP TIME'] = df_1['PLAN SHIP TIME'].str.replace('.',':')
  df_1['SHIPMENT LOAD DATE'] = df_1['SHIPMENT LOAD DATE'].str.replace('-','/')
  df_1['SHIPMENT LOAD TIME'] = df_1['SHIPMENT LOAD TIME'].str.replace('.',':')
  print(df_1.head())
#  df_1 = df_1[['tDate', 'OPCD', 'pCode1', 'OrderNo3','Module', 'Shift1']]

  df_x2 = pd.DataFrame()
  #df_1 = df_1[df_1['ORDER NO'].str.contains()]
  for ww in TargeDate1:
    df_x1 = df_1.copy(deep=True) 
    df_x1 = df_x1[df_x1['ORDER NO'].str.contains(ww)]

    df_x2 = pd.concat([df_x2, df_x1], axis=0)
    print(df_x2.head())

  df_2 = df_x2.assign(tPlanDate='', tPlanDateTime='',tLoadDate='', tLoadDateTime='')
#  print(df_2.head())

  df_2[['tPlanDate', 'tPlanDateTime']] = df_2.apply(lambda x: formatDate2(x['PLAN SHIP DATE'], x['PLAN SHIP TIME']), axis=1, result_type="expand")
  df_2[['tLoadDate', 'tLoadDateTime']] = df_2.apply(lambda x: formatDate2(x['SHIPMENT LOAD DATE'], x['SHIPMENT LOAD TIME']), axis=1, result_type="expand")

  df_3=df_2[['TEMP.TRAILER', 'TRAILER NO','ORDER NO','PALLETIZASION','MODULE NO','tPlanDate','tPlanDateTime','tLoadDate','tLoadDateTime']]
  print(df_3.head())

  return df_3


### Functions End ####################################################################################################################
##### Power BI #######

#IP202file = kpifilePath + IP202

#df_IP202 = pd.read_csv(IP202file, usecols=[0,1,2,3,5,7,8,9,11,12,13,14,17,18])
#df_IP202.to_csv('./work9/df_IP202.csv')

#df_IP202_1 = set_IP202(df_IP202)
#df_IP202_1.to_csv('./work9/df_IP202_1.csv')


##### PEnd ower BI #######

######################################################################
kpifile = kpifilePath + kpiCSV

df_KPI = pd.read_csv(kpifile, usecols=[0,1,4,5,6,7,8,9,10,11])

df_KPI[['sDate', 'sDateTime','tDate', 'sShift1', 'sShift2']] = df_KPI.apply(lambda x: formatDate(str(x['PROCESS TIME'])), axis=1, result_type="expand")
df_KPI.to_csv(workPath + "df_KPI0.csv")

df_KPI[['ORDER NO', 'OrderNo1', 'OrderNo2', 'OrderNo3']] = df_KPI.apply(lambda x: set_orderNo(str(x['ORDER NO']), str(x['BACK ORDER'])), axis=1, result_type="expand")

df_KPI[['pCode1', 'pCode2']] = df_KPI.apply(lambda x: set_pCode(str(x['PALLET CODE']), str(x['OrderNo2'])), axis=1, result_type="expand")

df_KPI = df_KPI.reindex(columns=['sDate', 'tDate', 'OPCD', 'DOCK', 'ORDER NO', 'BACK ORDER', 'OrderNo1', 'OrderNo2', 'OrderNo3', 
                                 'ROUTE', 'PALLET CODE', 'pCode1', 'pCode2', 'PRODUCT CODE', 'MODULE NO','PICK LOCATION', 'sShift1', 'sShift2'])
df_KPI.columns = ['sDate', 'tDate','OPCD','Dock', 'OrderNo', 'BackOrder', 'OrderNo1', 'OrderNo2', 'OrderNo3', 'Route', 
                  'PalletCode', 'pCode1', 'pCode2', 'Parts', 'Module', 'Location', 'Shift1', 'Shift2']
print(df_KPI.shape[0])
df_KPI.to_csv(workPath + "df_KPI.csv")



#fig = plt.figure(figsize=(35, 50)) #Mon to Fri
fig = plt.figure(figsize=(40, 80))  #Mon to Sat
fig.suptitle(FIG_TITLE, fontsize=50)
fig.subplots_adjust(hspace=0.6, wspace=0.5, top = 0.95, bottom = 0.05, left = 0.14, right=0.8)

#ax_w = 7  #Mon to Sat
ax_w = 6  #Mon to Fri
ax_h = 17


df_KPI_0 = df_KPI.copy(deep=True)     
df_KPI_0 = df_KPI_0[df_KPI_0['Dock'] == 'S3']
print(df_KPI_0.shape[0])
df_KPI_0 = df_KPI_0[df_KPI_0['OrderNo1'].isin(TargeDate1)]
print(df_KPI_0.shape[0])
print(df_KPI_0.head())


df_KPI_1 = df_KPI_0[~df_KPI_0['PalletCode'].isnull()]
#df_KPI_1 = df_KPI_1[df_KPI_1['OrderNo3'].isin(OrderNos2)]

df_KPI_1.to_csv(workPath + "df_KPI_1.csv")

if df_KPI_1.shape[0] == 0:
    available = sorted(df_KPI['OrderNo1'].dropna().unique().tolist())
    print("ERROR: No rows match the requested report week after filtering.", file=sys.stderr)
    print(f"  Requested TargeDate1: {TargeDate1}", file=sys.stderr)
    print(f"  OrderNo1 values present in {kpiCSV}: {available}", file=sys.stderr)
    print("  Check that Data/Picking_MODULE.csv contains the correct week's export,", file=sys.stderr)
    print("  or re-run and enter a Monday date that matches the data.", file=sys.stderr)
    sys.exit(2)

df_KPI_2 = df_KPI_1.copy(deep=True)     

df_day_per_pcode_all = set_day_pcode_total(df_KPI_2)
df_day_per_pcode_all.to_csv(workPath + "df_day_per_pcode_all.csv")

ax1 = fig.add_subplot(ax_h, ax_w, 3)
plot_1bar(ax1, df_day_per_pcode_all, 'pCode2', 'pickCount', 'Total pick volume' ,'pCode', 'Module count')

#####################################################################################

df_w = df_KPI_1.copy(deep=True) 
df_pcode0 = set_pcode0(df_w)
df_pcode1 = set_pcode1(df_pcode0)
df_pcode2 = set_pcode2(df_pcode1)

df_sum1_PB = set_sum1_PB(df_pcode0)


df_pcode_per_shift = sum_per_shift(df_pcode1)
print(df_pcode_per_shift.head())
df_pcode_per_shift.to_csv(workPath + "df_pcode_per_shift_main.csv")
ax1 = fig.add_subplot(ax_h, ax_w, 1)
plot_2bar_1line(ax1, df_pcode_per_shift, df_pcode_per_shift[df_pcode_per_shift['Shift1'] == 'shift1'],
                df_pcode_per_shift[df_pcode_per_shift['Shift1'] == 'shift2'],
                'Shift1', 'pickCount','tspend', \
                "Pick_count and Pick_speed - All Pallet Code", \
                'Shift', 'Module count', 'Pick_Speed per Module (seconds)')


df_pcode_per_shift_speed = sum_per_shift_speed(df_pcode1)
print(df_pcode_per_shift_speed.head())
ax1 = fig.add_subplot(ax_h, ax_w, 2)
plot_plot(ax1, df_pcode_per_shift_speed, 'Shift1', "Pick_speed - All Pallet Code", 'pCode1', 'Pick_Speed per Module (seconds)')

#####################################################################################
#df_w = df_KPI_1.copy(deep=True) 
#df_sum0 = set_sum0(df_w)
#df_sum1 = set_sum1(df_sum0)

df_ope = df_KPI_1.copy(deep=True) 
df_ope0 = set_ope0(df_ope)
df_ope1 = set_ope1(df_ope0)
df_ope2 = set_ope2(df_ope1)

df_ope2 = df_ope2.sort_values(['iG_2'])
print(df_ope2)

ax1 = fig.add_subplot(ax_h, ax_w, (4, 5))
plot_1bar_1line(ax1, df_ope2,
                'iG_2', 'pickCount','ispeed', \
                "Pick_count and Pick_speed - All Pallet Code", \
                'Operator Code', 'Module count', 'Pick_Speed per Module (seconds)')


#df_week = df_sum1.copy(deep=True) 
df_week = df_pcode1.copy(deep=True) 
df_week1 = set_week1(df_week)
print(df_week1.head())

ax_count = 2
iLoop = 0
for pp in PCodes_list:
  #param_pcode = PCodes_list[iLoop]
  print(PCodes_list[iLoop])

#  df_w1 = df_sum2.copy(deep=True) 
  df_w1 = df_pcode2.copy(deep=True) 

  df_w1 = df_w1[df_w1['pCode1'] == pp]

  ax_count += 1
  df_ww = sum_per_shift(df_w1)
  #print(df_ww.head())
  ax1 = fig.add_subplot(ax_h, ax_w, (iLoop + 1) * ax_w + 1)
  plot_2bar_1line(ax1, df_ww, df_ww[df_ww['Shift1'] == 'shift1'], df_ww[df_ww['Shift1'] == 'shift2'],
                  'Shift1', 'pickCount','tspend', "Pick_count and Pick_speed per shift\n" + PCodes_title[iLoop],
                  'Shift', 'Module count', 'Pick_Speed per Module (seconds)')

  ax_week = (iLoop + 1) * ax_w + 2
  iweek = 0

  
  for ww in TargeDate:
    df_x1 = df_week1.copy(deep=True) 

    df_x1 = df_x1[df_x1['pCode1'] == pp]
    print(df_x1.head())

#    sww1 = ww.strftime('%Y%m%d')
 #   print(sww1)

    df_x1 = df_x1[df_x1['OrderNo3'].str.contains(ww)]
#    df_x1 = df_x1[df_x1['iG_order3'].str.contains(ww)]
    print(df_x1.head())

#    df_ww = sum_per_shift(df_w1)
    ax1 = fig.add_subplot(ax_h, ax_w, ax_week)
    plot_1bar_1line_1(ax1, df_x1,
                'iG_order3', 'tspend','ispeed', \
                "Pick_time and Pick_speed\n" + WeekDays[iweek], \
                'order Number', 'spend time (minutes)', 'Pick_Speed per Module (seconds)')
    ax_week += 1
    iweek += 1

  plot_title(fig, ax1, PCodes_title[iLoop])

  iLoop += 1


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

with open(kpiHTMLPath + 'IP_Picking_KPI.html', "w") as w:
    w.write(html)
