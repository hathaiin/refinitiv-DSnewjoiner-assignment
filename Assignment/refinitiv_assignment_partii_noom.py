import pandas as pd
import datetime
import calendar
from collections import defaultdict

def import_data(path,sheet_name):
  data_excel = pd.ExcelFile(path)
  data_sheet = pd.read_excel(data_excel, sheet_name, index_col='Date')
  data = data_sheet.sort_index()
  return data

'''
*** Check Full Week Function ***
Return : Dictionary of 4 months data
Keys   -> month name
Values -> list of 4 week dicts that contain 5 values of data 

e.g. {'Dec': [{Timestamp(Mon) : bid close, 
               Timestamp(Tue) : bid close,
               Timestamp(Wed) : bid close,
               Timestamp(Thu) : bid close,
               Timestamp(Fri) : bid close },

               {Timestamp(Mon) : bid close,
               Timestamp(Tue) : bid close, ... }, ...],
      'Mar': [{Timestamp(Mon) : bid close, ... }, ...], ... }
                                
'''
def check_full_week(data):
  check_week = 0
  data_dict = {}
  my_dict = defaultdict(list)

  for index, value in data.iterrows():
    month = index.strftime("%b")                # get month name (e.g. 2019-03-01 = Mar) 
    week_num = index.strftime("%w")             # get date number of week (e.g. Mon = 1, Tue = 2, Fri = 5)

    if month in ('Mar','Jun','Sep','Dec'):
      if int(week_num) == int(check_week+1):    # check week(num) day (start at Mon(1))
        if len(my_dict[month]) != 4:            # check last week of month (count number of week ,must 4)
          data_dict.update({index : value[0]})  # keep date:bid into dict
          check_week+=1                         # check next date

        if check_week == 5 :                    # meet Fri day then append to that month dict and clear dict
          check_week = 0
          my_dict[month].append(data_dict)
          data_dict = {}
  return my_dict

def max_profit(bidClose_list):
  max_diff = 0
  max_idx = 0

  bid_min = min(bidClose_list[:4])                              # find min value of week (except value on Friday)
  buy_idx = bidClose_list.index(bid_min)                        # get index of min value

  for i,bid_val in enumerate(bidClose_list[buy_idx+1:]):        # calculate max difference from next of min value [buy_idx+1:] to end list
    diff = bid_val - bid_min
    if diff > max_diff:
      max_diff = diff                                           # update max diff value
      max_idx = i                                               # max diff index

  return buy_idx, buy_idx+max_idx+1, max_diff
  
def main():

  path = '/Users/hathaiin/Desktop/Refinitiv/Assignment/PreciousMetalSpot.xlsx'
  sheet_name = 'Gold Spot'

  df_gold = import_data(path,sheet_name)
  my_dict = check_full_week(df_gold)

  find_max_profit = []
  for month,value in my_dict.items():
    for week in range(len(my_dict[month])) :
      for date_key,bid_value in my_dict[month][week].items():
        print(date_key, bid_value)
        find_max_profit.append(bid_value)

      ### Show Calendar 2019 ###
      print()
      num = int(date_key.strftime("%m"))
      print(calendar.month(2019,num))

      # print(find_max_profit)
      # print()
      
      ### MaxProfit Function ###
      buy_idx, sell_idx, max_diff = max_profit(find_max_profit)
      # print(buy_idx, sell_idx, max_diff)
      # print('Buy        : ' ,list(my_dict[month][week].values())[buy_idx])
      # print('Sell       : ' ,list(my_dict[month][week].values())[sell_idx])
      # print('Max Profit : ' ,max_diff)

      buyValue = list(my_dict[month][week].values())[buy_idx]
      sellValue = list(my_dict[month][week].values())[sell_idx]

      buyDate = str(list(my_dict[month][week].keys())[buy_idx]).replace(" 00:00:00", "")
      sellDate = str(list(my_dict[month][week].keys())[sell_idx]).replace(" 00:00:00", "")

      firstWeekDate = str(list(my_dict[month][week].keys())[0]).replace(" 00:00:00", "")
      endWeekDate = str(list(my_dict[month][week].keys())[-1]).replace(" 00:00:00", "")

      print('Max Profit Week {} to {} is {:.4f}, buy at {:.4f} on {}, sell at {:.4f} on {}'\
            .format(firstWeekDate, endWeekDate, max_diff, buyValue, buyDate, sellValue, sellDate))
      print('----------------------------------------------------------------------------------------------------------------------')
      find_max_profit = []

if __name__ == "__main__":
  main()

# test_max_pro = [1498.81, 1485.27, 1498.6, 1510.4167, 1511.2979]   
# test_max_pro1 = [1485.27, 1498.81,  1498.6, 1510.4167, 1511.2979] 
# test_max_pro2 = [1485.27, 1498.81,  1498.6, 1511.2979, 1510.4167]     
# test_max_pro3 = [1498.81, 1485.27,  1498.6, 1511.2979, 1510.4167]
