#!/usr/bin/env python
# coding: utf-8

import os
import re
import xlwt
import pickle
import pandas as pd
from datetime import datetime
from random import sample, shuffle
from pathlib import Path

def main():
    forms_dir = Path('./sample_data')
    result_dir = Path(forms_dir, 'results')

    last_lottery = '2021-04-20 16:00'
    deadline = '2021-05-04 16:00'

    applications_filename = r'sample_from_sharepoint.xlsx'
    inventory_filename    = r'SE Inventory.xlsx'

    winner_file_ss = 'winner_file_ss.pickle'
    winner_file_sk = 'winner_file_sk.pickle'

    end_sk_inventory = 999 # all ss container stuff is bigger than 1000

    # hardcode indices of skis:
    # ['Fjell skis /w Telemark 3-pin binding', 'Fjell skis /w BC binding', 'Cross country skis', 'Randonee skis', 'Freeride skis', 'Snowboard']
    ski_ind_list = [1119, 1120, 1121, 1122, 1123, 1124]

    # build all the paths to the input and output files
    today_string = datetime.strftime(datetime.today(), '%Y-%m-%d')
    result_filename = f'{today_string}_handout.xls'

    inventory_path    = Path(forms_dir, inventory_filename)
    applications_path = Path(forms_dir, applications_filename)

    result_path       = Path(result_dir, result_filename)

    for path in [inventory_path, applications_path]:
        if not os.path.isfile(path):
            raise ValueError(f'{path} does not exist. Check input files.')

    # if the directory for the results does not exist, make it
    if not os.path.isdir(result_path.parent):
        os.mkdir(result_path.parent)


    # read the files
    applications = pd.read_excel(applications_path,
                               usecols=['Completion time', 'Terms and Conditions', 'Name', 'Item Numbers'],
                               parse_dates=['Completion time'],
                               dtype={'Item Numbers': str})
    inventory = pd.read_excel(inventory_path, index_col=0)

    # clean up inventory
    # drop everything that is not an inventory item, eg headers
    inventory = inventory[inventory.index.notna()]

    # drop duplicates in applications
    # search for duplicates in Name and Username sepereately, with both in a list it will only find duplicates with both
    applications.drop_duplicates('Name', keep='last', inplace=True)

    ### clean up applications ###
    # check if applications were submitted in the time frame    
    deadline = datetime.strptime(deadline, '%Y-%m-%d %H:%M')
    lasttime = datetime.strptime(last_lottery, '%Y-%m-%d %H:%M')

    # remove the timezone data from the applications, so they can be compared to the deadline
    application_times = [t.replace(tzinfo=None) for t in applications['Completion time']]
    before_deadline = [t < deadline for t in application_times]
    after_last = [t > lasttime for t in application_times]

    keep = [a and b for a,b in zip(before_deadline, after_last)]
    applications = applications[keep]

    # drop terms and conditions deniers (should only occur with too old data)
    applications = applications[applications['Terms and Conditions'].notna()]


    ### write want dicts ###
    # loop through applications and put the items on 'want list' and put in, who wants them
    want_dict_ss = {}
    want_dict_sk = {}


    for _, person in applications.iterrows():
        itemlist = []
        # clean up input and make items into list of integers
        # if input does not convert to integers, skip it
        if isinstance(person['Item Numbers'], str):
            for split in re.findall(r'\d+', person['Item Numbers']): # \d+ finds all numbers in a row '023a4. 5' = [023, 4, 5]
                try:
                    itemlist.append(int(split))
                except:
                    print(person['Name'], person['SK'], split)

        # add items to want dict, put people onto items
        for item in itemlist:
            if item not in inventory.index.to_list(): # check if item is in the inventory list
                pass
            else:
                if item <= end_sk_inventory:
                    if item in want_dict_sk.keys(): # if item is already in the list, append the new name
                        want_dict_sk[item] = want_dict_sk[item] + [person['Name']]
                    else:
                        want_dict_sk[item] = [person['Name']]
                else:
                    if item in want_dict_ss.keys(): # if item is already in the list, append the new name
                        want_dict_ss[item] = want_dict_ss[item] + [person['Name']]
                    else:
                        want_dict_ss[item] = [person['Name']]


    # Lottery
    won_dict_ski = do_ski_lottery(ski_ind_list, want_dict_ss, inventory)
    won_dict_ss = do_lottery(want_dict_ss, inventory)
    won_dict_sk = do_lottery(want_dict_sk, inventory)


    # now go through all the winner lists and gather the items one person has won
    winner_sk = gather_wins(won_dict_sk, inventory)
    winner_ss = gather_wins(won_dict_ss, inventory)
    winner_ski = gather_wins(won_dict_ski, inventory)

    # add ski winners to the winner dict
    for name, ski in winner_ski.items():
        if name in winner_ss.keys():
            winner_ss[name] += ski
        else:
            winner_ss[name] = ski


    # make lists with readable items
    winner_sk_readable = make_readable(winner_sk, inventory)
    winner_ss_readable = make_readable(winner_ss, inventory)

    # sort winner alphabeticaly
    sorted_sk = sort_by_name(winner_sk_readable)
    sorted_ss = sort_by_name(winner_ss_readable)

    ### write everything to an excel sheet
    write_to_excel(['Sjoerskrenten', 'Snowscooter'], [sorted_sk, sorted_ss], result_path)

    # save the winners to pickle
    with open(Path(result_dir, winner_file_ss), 'wb') as fp:
        pickle.dump(winner_ss, fp)

    with open(Path(result_dir, winner_file_sk), 'wb') as fp:
        pickle.dump(winner_sk, fp)


    print(f'Written results to {result_path}')
    print(f'Written at {datetime.now()}')
    print('Done.')


def do_lottery(want_dict, inventory):
    won_dict = {}
    
    for item, applicants in want_dict.items():
        demand = len(applicants)
        stock = inventory['Number'][item]
        
        # check demand
        if demand > stock:
            # draw random sample
            won = sample(applicants, int(stock))
            won_dict[item] = won
            # print('item', item, 'stock', stock, 'applicants', len(applicants), 'winners', len(won), won)
        else:
            # enough items -> everybody gets one
            won_dict[item] = applicants
        
    return won_dict


def do_ski_lottery(ski_ind_list, want_dict, inventory):
    # pay special attention to the skis and boots and poles
    # do the lottery for skis only. Everybody who gets skis, will get boots


    # shuffle, so that there is no bias for handing out skis because of the order in ski_names
    shuffle(ski_ind_list)
    won_dict_ski = {i:[] for i in ski_ind_list}

    for item in ski_ind_list:
        if item in want_dict.keys():
            applicants_all = set(want_dict[item]) # don't let people apply twice for skis to increase chances

            # delete applicants who already have an other type of ski
            already_won = []
            # make list of people who already won skis
            for winners in won_dict_ski.values():
                already_won += winners

            # delete winners form list of applicants, so they don't get two pairs of skis
            applicants = [person for person in applicants_all if person not in already_won]

            demand = len(applicants)
            stock = inventory['Number'][item]

            if demand > stock:
                won = sample(applicants, int(stock))
                won_dict_ski[item] += won
            else:
                won_dict_ski[item] += applicants


    # Lottery on boots is kind of useless. When people got skis, they just have to find some boots that fit.
    # If you do the lottery on boots too, it is possible that someone gets skis, but no boots
    # rather do the lottery on skins


    # get indices of boots
    boot_names = ('Fjellski shoes Telemark', 'Fjellski shoes BC', 'Cross Country shoes', 'Randonne ski boots', 'Freeride Boots', 'Snow board boots','Fjell ski skins short', 'Poles')
    boot_indices = {}

    for boots in boot_names:
        # get the item numbers for every boot type
        items_boots = inventory.index[[boots in name for name in inventory['Name']]]
        boot_indices[boots] = items_boots # save inventory numbers for every kind of boot

    # delete skis and boots from want list snow scooter to not do the lottery on them again

    indices_to_delete = ski_ind_list

    for index in boot_indices.values():
        indices_to_delete += [i for i in index]

    for index in indices_to_delete:
        if index in want_dict.keys(): # only delte stuff from the list, if it is really in there
            del want_dict[index]

    return won_dict_ski


def gather_wins(won_dict, inventory):
    '''make a dict with NAME:[items] out of the lists with items:[NAMES]'''
    winners_dict = {}
    
    for item, winners in won_dict.items():
        for person in winners:
            if person in winners_dict.keys():
                winners_dict[person] += [item]
            else:
                winners_dict[person] = [item]
    
    return winners_dict


def make_readable(winners, inventory):
    '''use item names instead of numbers'''

    winner_readable = {}
    
    for winner, item in winners.items():
        names = [inventory['Name'][i] for i in item]
        winner_readable[winner] = names
    
    return winner_readable


def sort_by_name(winner_readable):
    '''sort winner lists alphabetically'''

    sorted_dict = {}
    for name, items in sorted(winner_readable.items()):
        sorted_dict[name] = items
    
    return sorted_dict


def write_to_excel(sheet_names, winnerdicts, result_path):
    assert len(sheet_names) == len(winnerdicts), 'you need one sheet name for every winner dict'

    wb = xlwt.Workbook() 
    line_width = 20

    style_header_container = xlwt.easyxf("alignment: wrap True; font: bold on, height 280")
    style_header           = xlwt.easyxf("alignment: wrap True; borders: left thin, right thin, top thin, bottom thin; font: bold on")
    style                  = xlwt.easyxf("alignment: wrap True, vert centre; borders: left thin, right thin, top thin, bottom thin")

    # create the sheets
    sheetlist = [wb.add_sheet(name) for name in sheet_names]

    today_string = datetime.strftime(datetime.today(), '%Y-%m-%d')

    for sheet, result, header in zip(sheetlist, winnerdicts, sheet_names):
        # set size for columns
        sheet.col(0).width = 256 * line_width + 1000
        sheet.col(1).width = 256 * line_width + 2000

        sheet.col(2).width = 4000
        sheet.col(3).width = 5000

        sheet.write_merge(0, 0, 0, 1, f'{header} {today_string}', style_header_container)
        sheet.row(0).height_mismatch = True       # for the adjustment of the row height
        sheet.row(0).height = 400

        # write header
        sheet.write(2, 0, 'Name', style_header)
        sheet.write(2, 1, 'Equipment', style_header)
        sheet.write(2, 2, 'Comments', style_header)
        sheet.write(2, 3, 'Signature', style_header)

        row = 3 # start row
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



if __name__ == "__main__":
    main()