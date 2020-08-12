import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

def import_data(path_file,sheet_name):
  data_excel = pd.ExcelFile(path_file)
  data = pd.read_excel(data_excel, sheet_name, index_col='Date')
  return data

"""
Bid Close  -->  Today(T) Close Price     ex. 1 Jan 2019
Bid Close1 -->  Tomorrow(T+1) Close Price  ex. 2 Jan 2019

Daily Return = [[Close Price(T+1) - Close Price(T)]/Close Price(T)]*100

Note: Date sort by Descending (use shift1)
"""
def daily_return(data):
  data['Bid Close1'] = data['Bid Close'].shift(1)                                         # Next day Bid Close
  data['Daily Return'] = ((data['Bid Close1']-data['Bid Close'])/data['Bid Close'])*100   # Percentage of Daily Return
  return data

def visualize(data,data_name):
  dr_data = data['Daily Return']
  
  plt.figure(figsize=(8,5))
  plt.hist(dr_data, bins=16, rwidth=0.5)
  plt.grid(axis='y', alpha=0.75)

  # 2 digits(float) of min, max daily return
  min = round(dr_data.min(),2)
  max = round(dr_data.max(),2)

  # plt.xticks(np.arange(min, max, 0.5), rotation=90)
  plt.xticks(np.arange(-8.0, 8.0, 0.5), rotation=90)
  plt.yticks(np.arange(0.0, 110.0, 10.0))

  plt.title(data_name, fontsize=15, fontweight='bold')
  plt.xlabel('% Daily Return')
  plt.show()

def main():

  path_file = '/Users/hathaiin/Desktop/Refinitiv/Assignment/PreciousMetalSpot.xlsx'

  gold = import_data(path_file,'Gold Spot')
  silver = import_data(path_file,'Silver Spot')
  platinum = import_data(path_file,'Platinum Spot')
  palladium = import_data(path_file,'Palladium Spot')

  gold_result = daily_return(gold)
  silver_result = daily_return(silver)
  platinum_result = daily_return(platinum)
  palladium_result = daily_return(palladium)

  visualize(gold_result, 'Gold')
  visualize(silver_result, 'Silver')
  visualize(platinum_result, 'Platinum')
  visualize(palladium_result, 'Palladium')

if __name__ == "__main__":
    main()
