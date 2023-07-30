import pandas as pd

class Excel_Utility:

  def __init__(self,filepath):
      self.filepath = filepath
      self.read_excel_file()

  def read_excel_file(self):
    self.data = pd.read_excel(self.filepath,sheet_name=0)

  def display_data(self):
    print(self.data)

  def get_columns(self):
    file = self.data
    columns = file.columns
    print('Columns: ', end=" ")
    for column in columns:
      print(column,end=" ")
    return columns

  def sort_by_column_name(self,column_name,kind='quicksort',ascending=True):
      data=self.data
      self.data = data.sort_values(by=column_name,axis=0,ascending=ascending,kind=kind)
      print(f'After sorting: \n{self.data}')

  def sort_by_column_number(self,index,kind='quicksort',ascending=True):
      data=self.data
      index = list(map(lambda ele:ele-1,index))
      self.data = data.sort_values(by=list(self.data.columns[index]),axis=0,ascending=ascending,kind=kind)
      print(f'After sorting: \n{self.data}')

  def filter_by_field(self,field,condition):
    data = self.data
    print(data[(data[field]==condition)])
  
  def write_to_excel(self):
    self.data.to_excel(self.filepath)

def display_options():
  print('Options')
  print('1. View data')
  print('2. View column names')
  print('3. Sort by column numbers')
  print('4. Sort by column name')
  print('5. Write the changes to the file')
  print('6. Exit')

def sort_by_index(utlis):
  print('Note: If you want to sort with multiple columns, please enter space seperated values')
  index = input('Enter index number: ')
  index = list(map(lambda ele:int(ele),index.split(' ')))
  columns= utils.get_columns()
  while max(index)>len(columns):
    print('\nNote: If you want to sort with multiple columns, please enter space seperated values')
    index = input('Enter column number: ')
    try:
      index = list(map(lambda ele:int(ele),index.split(' ')))
    except Exception as e:
      index = list(map(lambda ele:int(ele),index))
  print('\n1. Quicksort')
  print('2. Mergesort')
  print('3. Heapsort')
  sort_style = input('Enter sorting style: ')
  while sort_style not in ['1','2','3']:
    sort_style = input('Enter sorting style: ')
  sort_style = int(sort_style)
  if sort_style ==1:
    sort_style = 'quicksort'
  elif sort_style ==2:
    sort_style = 'mergesort'
  elif sort_style ==3:
    sort_style = 'heapsort'
  print('1. Ascending Order')
  print('2. Descending Order')
  order = input('Select order: ')
  while order not in ['1','2']:
    order = input('Select order: ')
  order=int(order)
  if order ==1:
    order = True
  elif order ==2:
    order = False
  utils.sort_by_column_number(index=index,kind=sort_style,ascending=order)

def sort_by_column(utils):
  columns= utils.get_columns()
  print('\nNote: If you want to sort with multiple columns, please enter space seperated values')
  column_names = input('Enter column name: ')
  column_names = column_names.split(' ')
  flag= False
  for column in column_names:
    if column not in columns:
      flag = True
  while flag:
      print('\nNote: If you want to sort with multiple columns, please enter space seperated values')
      column_names = input('Enter column name: ')
      column_names = column_names.split(' ')
      for column in column_names:
        if column in columns:
          flag = False
  print('1. Quicksort')
  print('2. Mergesort')
  print('3. Heapsort')
  sort_style = input('Enter sorting style: ')
  while sort_style not in ['1','2','3']:
    sort_style = input('Enter sorting style: ')
  sort_style = int(sort_style)
  if sort_style ==1:
    sort_style = 'quicksort'
  elif sort_style ==2:
    sort_style = 'mergesort'
  elif sort_style ==3:
    sort_style = 'heapsort'
  print('1. Ascending Order')
  print('2. Descending Order')
  order = input('Select order: ')
  while order not in ['1','2']:
    order = input('Select order: ')
  order=int(order)
  if order ==1:
    order = True
  elif order ==2:
    order = False
  utils.sort_by_column_name(column_names,kind=sort_style,ascending=order)
def take_user_input(utils):
  while True:
    display_options()
    try:
      selection = int(input('Please Enter Your Option: '))
      if selection==1:
        utils.display_data()
      elif selection==2:
        utils.get_columns()
      elif selection==3:
        sort_by_index(utils)
      elif selection==4:
        sort_by_column(utils)
      elif selection==5:
        utils.write_to_excel()
      elif selection==6:
        break

    except Exception:
      print('\nPlease select one of the given options properly.')
      take_user_input(utils)
if __name__=='__main__':
  filepath = input('Enter the excel file path: ')
  utils = Excel_Utility(filepath)
  take_user_input(utils)
