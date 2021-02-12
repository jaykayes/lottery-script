#!/usr/bin/env python
# coding: utf-8

# In[1]:


# fix long autocomplete in jupyter notebook
get_ipython().run_line_magic('config', 'Completer.use_jedi = False')


# In[2]:


import os
import xlwt
import pandas as pd
from datetime import datetime
from random import sample
from pathlib import Path


# In[3]:


forms_dir = Path('../2020-02-11_handout/')
result_dir = forms_dir

TC_filename           = r'SE T&C Form.csv'
applications_filename = r'SE Application Form.csv'
inventory_filename    = r'SE Inventory - Inventory.csv'


# In[4]:


# build all the paths to the input and output files
result_filename = '{}_handout.xls'.format(datetime.strftime(datetime.today(), '%Y-%m-%d'))

TC_path           = Path(forms_dir, TC_filename)
inventory_path    = Path(forms_dir, inventory_filename)
applications_path = Path(forms_dir, applications_filename)

result_path       = Path(result_dir, result_filename)

for path in [TC_path, inventory_path, applications_path]:
    if not os.path.isfile(path):
        raise ValueError('{} does not exist. Check input files.'.format(path))

# if the directory for the results does not exist, make it
if not os.path.isdir(result_path.parent):
    os.mkdir(result_path.parent)


# In[5]:


# read the files
applications = pd.read_csv(applications_path,
                           usecols=['Timestamp', 'Username', 'Name', 'Equipment Sjoeskrenten', 'Equipment Ski/Snowscooter'],
                           parse_dates=['Timestamp'],
                           dtype={'Equipment Sjoeskrenten': str, 'Equipment Ski/Snowscooter': str})
tc_form = pd.read_csv(TC_path, usecols=['Name', 'E-Mail'])
inventory = pd.read_csv(inventory_path)


# In[6]:


# get indices, where inventory of the containers start and ent
start_sk_inventory = inventory.index[inventory['Name'] == 'Sjoeskrenten Container:'].to_list()[0]
end_sk_inventory   = inventory.index[inventory['Name'] == 'Snowscooter Container:'].to_list()[0]

# split for sjoerskrenten and snowscooter containers
sk_inventory = inventory.iloc[start_sk_inventory + 1:end_sk_inventory, :].copy()
ss_inventory = inventory.iloc[end_sk_inventory + 1: ,:].copy()

# shift indices to match line numbers
sk_inventory = sk_inventory.shift(2)
ss_inventory = ss_inventory.shift(2)

# drop empty lines, only keep those with items in 'Name'
sk_inventory.dropna(subset=['Name'], inplace=True)
ss_inventory.dropna(subset=['Name'], inplace=True)


# In[7]:


# drop duplicates in applications
# search for duplicates in Name and Username sepereately, with both in a list it will only find duplicates with both
applications.drop_duplicates('Name', keep='last', inplace=True)
applications.drop_duplicates('Username', keep='last', inplace=True)

# change column names for quicker typing
applications.rename(columns={'Equipment Sjoeskrenten':'SK', 'Equipment Ski/Snowscooter':'SS'}, inplace=True)


# In[8]:


#ToDO: kick out people who did not sign T&C and who are behind deadline


# In[9]:


# loop through applications and put the items on 'want list' and put in, who wants them
want_dict_sk = {}
want_dict_ss = {}

for _, person in applications.iterrows():
    itemlist_SS = []
    itemlist_SK = []
    # clean up input and make items into list of integers
    # if input does not convert to integers, skip it
    try:
        clean = person['SK'].rstrip(', ')
        itemlist_SK = [int(s.lstrip(' ')) for s in clean.split(',')]
    except:
        pass
        if str(person['SK']) != 'nan':
            print('wrong format for', person['Name'], 'in SK list')
            print(person['SK'])
    try:
        clean = person['SS'].rstrip(', ')
        itemlist_SS = [int(s.lstrip(' ')) for s in clean.split(',')]
    except:
        pass
        if str(person['SS']) != 'nan':
            print('wrong format for', person['Name'], 'in SS list')
            print(person['SS'])

    # add items to want dict, put people onto items
    for item in itemlist_SK:
        if item not in sk_inventory.index.to_list(): # check if item is in the inventory list
            pass
        else:
            if item in want_dict_sk.keys(): # if item is already in the list, append the new name
                want_dict_sk[item] = want_dict_sk[item] + [person['Name']]
            else:
                want_dict_sk[item] = [person['Name']]
    
    for item in itemlist_SS:
        if item not in ss_inventory.index.to_list(): # check if item is in the inventory list
            pass
        else:
            if item in want_dict_ss.keys(): # if item is already in the list, append the new name
                want_dict_ss[item] = want_dict_ss[item] + [person['Name']]
            else:
                want_dict_ss[item] = [person['Name']]


# In[10]:


# check demand of every item, and if neccessary, do the lottery

# SK container
won_dict_sk = {}
for item, applicants in want_dict_sk.items():
    demand = len(applicants)
    stock = sk_inventory['Number'][item]
    
    if demand > stock:
        won = sample(applicants, int(stock))
        won_dict_sk[item] = won
    else:
        won_dict_sk[item] = applicants


# In[11]:


# pay special attention to the skis and boots and poles
# do the lottery for skis only. Everybody who gets skis, will get boots
ski_names = ('Fjell skis /w Telemark 3-pin binding', 'Fjell skis /w BC binding', 'Cross country skis', 'Randonee skis', 'Freeride skis', 'Snowboard')
won_dict_ski_readable = {ski:[] for ski in ski_names}
ski_indices = {}

for ski in ski_names:
    # get the item numbers for every ski type
    items_skis = ss_inventory.index[[ski == name for name in ss_inventory['Name']]]
    ski_indices[ski] = items_skis
    
# for every ski type, check who wants it
for ski in ski_names:
    # iterate through all ski types
    for item in ski_indices[ski]:
        # check if skis are requested
        if item in want_dict_ss.keys():
            applicants = want_dict_ss[item]
            demand = len(applicants)
            stock = ss_inventory['Number'][item]
            
            if demand > stock:
                won = sample(applicants, int(stock))
                won_dict_ski_readable[ski] += won
            else:
                won_dict_ski_readable[ski] += applicants

won_dict_ski = {}
# make a dict with the written name, better to check
for ski, people in won_dict_ski_readable.items():
    item = ski_indices[ski][0]
    won_dict_ski[item] = people


# In[12]:


# Lottery on boots is kind of useless. When people got skis, they just have to find some boots that fit.
# If you do the lottery on boots too, it is possible that someone gets skis, but no boots
# rather do the lottery on skins


# In[13]:


# boots only for people who got skis
boot_names = ('Fjellski shoes Telemark', 'Fjellski shoes BC', 'Cross Country shoes', 'Randonne ski boots', 'Freeride Boots', 'Snow board boots')
boot_indices = {}

won_dict_boots = {}

for boots in boot_names:
    # get the item numbers for every boot type
    items_boots = ss_inventory.index[[boots in name for name in ss_inventory['Name']]]
    boot_indices[boots] = items_boots # save inventory numbers for every kind of boot

# go through list of ski winners
for ski, people in won_dict_ski_readable.items():        
    # find boots for skis
    index = ski_names.index(ski) # order of boot_names and order of ski_names has to match up
    boots = boot_names[index]
    
    # find inventory numbers of those boots
    items = boot_indices[boots]
    # check demand for every inventory number
    for item in items:
        if item in want_dict_ss.keys(): # check if anybody wants this size of boots
            applicants = want_dict_ss[item]
            demand = len(applicants)
            stock = ss_inventory['Number'][item]

            if demand > stock:
                won = sample(applicants, int(stock))
                won_dict_boots[item] = won
            else:
                won_dict_boots[item] = applicants
# In[14]:


# delete skis and boots from want list snow scooter to not do the lottery on them again

indices_to_delete = []
for index in ski_indices.values():
    indices_to_delete += [i for i in index]

for index in boot_indices.values():
    indices_to_delete += [i for i in index]

for index in indices_to_delete:
    if index in want_dict_ss.keys(): # only delte stuff from the list, if it is really in there
        del want_dict_ss[index]


# In[15]:


# Lottery for the rest of the container
# check demand of every item, and if neccessary, do the lottery

# SS container
won_dict_ss = {}
for item, applicants in want_dict_ss.items():
    demand = len(applicants)
    stock = ss_inventory['Number'][item]
    
    if demand > stock:
        won = sample(applicants, int(stock))
        won_dict_ss[item] = won
    else:
        won_dict_ss[item] = applicants


# In[16]:


# now go through all the winner lists and gather the items one person has won
winner_sk = {}
winner_ss = {}

for item, winners in won_dict_sk.items():
    item_name = sk_inventory['Name'][item]
    for person in winners:
        if person in winner_sk.keys():
            winner_sk[person] += [item]
        else:
            winner_sk[person] = [item]

for _dict in [won_dict_ss, won_dict_boots, won_dict_ski]:
    for item, winners in _dict.items():
        item_name = ss_inventory['Name'][item]
        for person in winners:
            if person in winner_ss.keys():
                winner_ss[person] += [item]
            else:
                winner_ss[person] = [item]


# In[17]:


# use item names instead of numbers
winner_sk_readable = {}
for winner, item in winner_sk.items():
    names = [sk_inventory['Name'][i] for i in item]
    winner_sk_readable[winner] = names
    
winner_ss_readable = {}
for winner, item in winner_ss.items():
    names = [ss_inventory['Name'][i] for i in item]
    winner_ss_readable[winner] = names


# In[18]:


# sort winner alphabeticaly
sorted_sk = {}
for name, items in sorted(winner_sk_readable.items()):
    sorted_sk[name] = items

sorted_ss = {}
for name, items in sorted(winner_ss_readable.items()):
    sorted_ss[name] = items


# In[19]:


# write everything to an excel sheet

wb = xlwt.Workbook() 
line_width = 20

style_header_container = xlwt.easyxf("alignment: wrap True; font: bold on, height 280")
style_header           = xlwt.easyxf("alignment: wrap True; borders: left thin, right thin, top thin, bottom thin; font: bold on")
style                  = xlwt.easyxf("alignment: wrap True, vert centre; borders: left thin, right thin, top thin, bottom thin")

  
# create the first sheet for the Sjoeskrenten results
sheet_sk = wb.add_sheet('Sjoeskrenten')
# create the second sheet for the Snowscooter results
sheet_ss = wb.add_sheet('Snowscooter')

for sheet, result, header in zip([sheet_sk, sheet_ss], [sorted_sk, sorted_ss], ['Sjoeskrenten', 'Snowscooter']):
    # set size for columns
    sheet.col(0).width = 256 * line_width + 1000
    sheet.col(1).width = 256 * line_width + 2000

    sheet.col(2).width = 4000
    sheet.col(3).width = 5000

    sheet.write_merge(0, 0, 0, 1, '{h} {a}'.format(h=header, a=datetime.strftime(datetime.today(), '%d.%m.%Y')), style_header_container)
    sheet.row(0).height_mismatch = True       # for the adjustment of the row height
    sheet.row(0).height = 400

    # write header
    sheet.write(2, 0, 'Name', style_header)
    sheet.write(2, 1, 'Equipment', style_header)
    sheet.write(2, 2, 'Comments', style_header)
    sheet.write(2, 3, 'Signature', style_header)

    row = 3 #start row
    for name, items in result.items():
        # separate items by linebreak
        formatted_items = ''
        sheet.write(row, 0, name, style)
        for item in items:
            formatted_items = formatted_items + '\n' + item

        sheet.write(row, 1, formatted_items, style)
        sheet.row(row).height_mismatch = True
        sheet.row(row).height = (len(items) + 2) * 256
        sheet.write(row, 2, '', style)
        sheet.write(row, 3, '', style)

        row = row + 1
    
wb.save(result_path)


# In[20]:


print('Written results to {}'.format(result_path))
print('Done.')

