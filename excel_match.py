# To match two different dataset store in Excel
# Source File "customer_source.xlsx"
# Target File "sites.xlsx" with sheet "WA 326", "ACT 464", "QLD 330", "SA 227", "NSW 420", "VIC 315"
# Match the column 'Site Name' with Target 'Name'. If can't match, match with 'Address'.
# Then get the 'Customer ID' as 'Biller Code'
# Return the result in 'result.xlsx'
# Target File Columns :
# Site Name : Biller Code : Name : Address : Customer ID
# Data Source Problem
# 1) Duplicate record found in Source's name. How to tackle? e.g. Albany Regional Hospital
# Matching Algorithm :
# - strip first
# - convert all the string into lower case for comparison
# - replace '-',',','/' by whitespace, delete '(', ')'
# - split the string by ' '
# - then match value by value
# - assume first value MUST match; (Street address) if the 1st value is number, 2nd value must match
# - define the short form :
#   {'uni':'university','rd':'road','tce':'terrace','ave':'avenue','ltd':'limited'}
# - ST : at the end of string -> street, otherwise ->saint
# - handle "City of XXX"
# - Site name match with Alliance name and Alliance address
# - Site address match with Alliance address ONLY

def matchWord (target, source) :
    """Match the string from target to source\n
       Parameters : source_str, target_str\n
       Return match count in integer. 0 - Not found \n
              target length (after special split) \n
              source length (after special split)"""

    target = target.strip().lower().replace('-',' ').replace(',', ' ').replace('/','').replace('&','').replace('(',' ').\
                replace(')',' ').replace('.',' ').replace("'",'')
    source = source.strip().lower().replace('-',' ').replace(',', ' ').replace('/','').replace('&','').replace('(',' ').\
                replace(')',' ').replace('.',' ').replace("'",'')
    l_target = target.split()
    l_source = source.split()
    target_len = len(l_target)
    source_len = len(l_source)

    # Data Preparation
    full_name = {'uni':'university','rd':'road','tce':'terrace','ave':'avenue','ltd':'limited',\
                 'st':'saint','hwy':'highway'}   # convert the short form to full name

    count = 0
    st_count = 0
    st_count_2nd = 0

    if target_len == 0 or source_len == 0 : return count, target_len, source_len
    # Return if the first word doesn't match, except Street number
    len_1st_target = len(l_target[0])
    for s in range(len_1st_target):
        num_chk_1st = l_target[0][s].isnumeric()        #   Handle L1, U1&2, etc.
        if num_chk_1st is True: break


    if target_len > 1 : num_chk_2nd = l_target[1].isnumeric()   # handle street name 100-101 xxx street
    else : num_chk_2nd = False

    if l_target[0] != l_source[0] and num_chk_1st == False: return count, target_len, source_len

    for i in range(target_len) :
        if i == target_len - 1 and l_target[i] == 'st' : l_target[i] = 'street'     # Set st = street at the end of string
        if l_target[i] in full_name : l_target[i] = full_name[l_target[i]]          # convert short form

        for r in range(source_len) :
            if r == source_len - 1 and l_source[r] == 'st': l_source[r] = 'street'  # Set st = street at the end of string
            if l_source[r] in full_name: l_source[r] = full_name[l_source[r]]       # convert short form

            # print(l_target[i], l_source[r])
            if l_target[i] == l_source[r] :
                count += 1
                if num_chk_1st and num_chk_2nd and i == 2 : st_count_2nd += 1
                if num_chk_1st and i == 1: st_count += 1    # For street address mapping
                break       # prevent duplicated count e.g. 'tower' counts twice for 'xxx tower xxx tower'
            # print(count, st_count, st_count_2nd, i , r)

        if num_chk_1st and num_chk_2nd and i == 2 and st_count_2nd == 0: return 0, target_len, source_len  # For 2nd number street address only
        if num_chk_1st and i == 1 and num_chk_2nd == False and st_count == 0 : return 0, target_len, source_len  # For 1st number street address only
    if 'city of' in target and 'city of' in source and count <=2 : return 0, target_len, source_len  # Handle 'City of XXX' matching

    return count, target_len, source_len


import pandas as pd

# Read Excel File
with pd.ExcelFile(r'customer_source.xlsx') as source_xlsx:
    df_source = pd.read_excel(source_xlsx, sheet_name='Sheet1', keep_default_na=False, usecols=[0,1,2], index_col=None,
                                names = ['name', 'id', 'addr'], header=0)

with pd.ExcelFile(r'sites.xlsx') as target_xlsx:
    df_target_1 = pd.read_excel(target_xlsx,sheet_name='WA 326', keep_default_na=False, usecols=[0,1,2,10], index_col=None,
                                names = ['Seq', 'Site_Name', 'Biller_Code', 'Site_Address'], header=0)
    df_target_2 = pd.read_excel(target_xlsx, sheet_name='ACT 464', keep_default_na=False, usecols=[0, 1, 2, 3],
                                index_col=None, names=['Seq', 'Site_Name', 'Biller_Code', 'Site_Address'], header=0)
    df_target_3 = pd.read_excel(target_xlsx, sheet_name='QLD 330', keep_default_na=False, usecols=[0, 1, 2, 3],
                                index_col=None, names=['Seq', 'Site_Name', 'Biller_Code', 'Site_Address'], header=0)
    df_target_4 = pd.read_excel(target_xlsx, sheet_name='SA 227', keep_default_na=False, usecols=[0, 1, 2, 3],
                                index_col=None, names=['Seq', 'Site_Name', 'Biller_Code', 'Site_Address'], header=0)
    df_target_5 = pd.read_excel(target_xlsx, sheet_name='NSW 420', keep_default_na=False, usecols=[0, 1, 2, 3],
                                index_col=None, names=['Seq', 'Site_Name', 'Biller_Code', 'Site_Address'], header=0)
    df_target_6 = pd.read_excel(target_xlsx, sheet_name='VIC 315', keep_default_na=False, usecols=[0, 1, 2, 3],
                                index_col=None, names=['Seq', 'Site_Name', 'Biller_Code', 'Site_Address'], header=0)

len_source = len(df_source.index)


#-------------------------Target 1 'WA 326----------------------------------
len_target1 = len(df_target_1.index)

df_target_1 = df_target_1[['Seq','Site_Name', 'Site_Address', 'Biller_Code']]
df_target_1['Site_Match'] = ''      # Percentage of Target Match
df_target_1['Alliance_Match'] = ''  # Match Ratio
df_target_1['Alliance_Name'] = ''
df_target_1['Alliance_Address'] = ''
df_target_1['Alliance_ID'] = ''

for r in range(len_target1) :                     # Search the Target File
    site_name = df_target_1['Site_Name'][r]
    site_addr = df_target_1['Site_Address'][r]

    match_count =[[0,0],0]                  #[Match Count, Row index)
    # match_ratio = 0
    tmp_name_count = [[0,0],0]
    tmp_addr_count = [[0,0],0]
    tmp_compare = [[0,0],0]

    for i in range(len_source) :        # Search in Source File
        tmp_source_name = df_source['name'][i]
        tmp_source_addr = df_source['addr'][i]

        if len(tmp_source_name) > 0 :
            match_result = matchWord(site_name, tmp_source_name)
            tmp_name_count = [match_result, i]    # Site Name vs Alliance Name

        if len(tmp_source_addr) > 0 :
            match_result = matchWord(site_name, tmp_source_addr)
            tmp_addr_count = [match_result, i]    # Site Name vs Alliance Address

            if len(site_addr) > 0 and 'no address' not in site_addr.strip().lower() and tmp_name_count[0][0] > 0:
                match_addr_result = matchWord(site_addr, tmp_source_addr)        # # Site Address vs Alliance Address
                if match_addr_result[0] > match_result[0] : tmp_addr_count = [match_addr_result, i]

        if tmp_name_count[0][0] >= tmp_addr_count[0][0] : # Compare Name vs Address
            tmp_compare = tmp_name_count
        else :
            tmp_compare = tmp_addr_count

        if tmp_compare[0][0] > match_count[0][0] :    # Filter maximum match count in source
            match_count = tmp_compare
    # Update Excel File
    if match_count[0][0] > 0 :
        df_target_1['Site_Match'].iloc[r] = int(match_count[0][0] / match_count[0][1] * 100)
        df_target_1['Alliance_Match'].iloc[r] = int(match_count[0][0]/match_count[0][2]*100)
        df_target_1['Alliance_Name'].iloc[r] = df_source['name'][match_count[1]]
        df_target_1['Alliance_Address'].iloc[r] = df_source['addr'][match_count[1]]
        df_target_1['Alliance_ID'].iloc[r] = df_source['id'][match_count[1]]
        df_target_1['Biller_Code'].iloc[r] = df_source['id'][match_count[1]]
    else :
        df_target_1['Site_Match'].iloc[r] = match_count[0][0]
        df_target_1['Alliance_Match'].iloc[r] = 0
        df_target_1['Alliance_Name'].iloc[r] = 'NOT FOUND'
        df_target_1['Alliance_Address'].iloc[r] = 'NOT FOUND'
        df_target_1['Alliance_ID'].iloc[r] = 'NOT FOUND'
        df_target_1['Biller_Code'].iloc[r] = 'NOT FOUND'

#-------------------------Target 1 'WA 326----------------------------------

#-------------------------Target 2 ACT 464----------------------------------
len_target2 = len(df_target_2.index)

df_target_2 = df_target_2[['Seq','Site_Name', 'Site_Address', 'Biller_Code']]
df_target_2['Site_Match'] = ''      # Percentage of Target Match
df_target_2['Alliance_Match'] = ''  # Match Ratio
df_target_2['Alliance_Name'] = ''
df_target_2['Alliance_Address'] = ''
df_target_2['Alliance_ID'] = ''

for r in range(len_target2) :                     # Search the Target File
    site_name = df_target_2['Site_Name'][r]
    site_addr = df_target_2['Site_Address'][r]

    match_count =[[0,0],0]                  #[Match Count, Row index)
    # match_ratio = 0
    tmp_name_count = [[0,0],0]
    tmp_addr_count = [[0,0],0]
    tmp_compare = [[0,0],0]

    for i in range(len_source) :        # Search in Source File
        tmp_source_name = df_source['name'][i]
        tmp_source_addr = df_source['addr'][i]

        if len(tmp_source_name) > 0 :
            match_result = matchWord(site_name, tmp_source_name)
            tmp_name_count = [match_result, i]    # Site Name vs Alliance Name

        if len(tmp_source_addr) > 0 :
            match_result = matchWord(site_name, tmp_source_addr)
            tmp_addr_count = [match_result, i]    # Site Name vs Alliance Address

            if len(site_addr) > 0 and 'no address' not in site_addr.strip().lower() and tmp_name_count[0][0] > 0:
                match_addr_result = matchWord(site_addr, tmp_source_addr)        # # Site Address vs Alliance Address
                if match_addr_result[0] > match_result[0] : tmp_addr_count = [match_addr_result, i]

        if tmp_name_count[0][0] >= tmp_addr_count[0][0] : # Compare Name vs Address
            tmp_compare = tmp_name_count
        else :
            tmp_compare = tmp_addr_count

        if tmp_compare[0][0] > match_count[0][0] :    # Filter maximum match count in source
            match_count = tmp_compare
    # Update Excel File
    if match_count[0][0] > 0 :
        df_target_2['Site_Match'].iloc[r] = int(match_count[0][0] / match_count[0][1] * 100)
        df_target_2['Alliance_Match'].iloc[r] = int(match_count[0][0]/match_count[0][2]*100)
        df_target_2['Alliance_Name'].iloc[r] = df_source['name'][match_count[1]]
        df_target_2['Alliance_Address'].iloc[r] = df_source['addr'][match_count[1]]
        df_target_2['Alliance_ID'].iloc[r] = df_source['id'][match_count[1]]
        df_target_2['Biller_Code'].iloc[r] = df_source['id'][match_count[1]]
    else :
        df_target_2['Site_Match'].iloc[r] = match_count[0][0]
        df_target_2['Alliance_Match'].iloc[r] = 0
        df_target_2['Alliance_Name'].iloc[r] = 'NOT FOUND'
        df_target_2['Alliance_Address'].iloc[r] = 'NOT FOUND'
        df_target_2['Alliance_ID'].iloc[r] = 'NOT FOUND'
        df_target_2['Biller_Code'].iloc[r] = 'NOT FOUND'

#-------------------------Target 2 'ACT 464----------------------------------

#-------------------------Target 3 QLD 330----------------------------------
len_target3 = len(df_target_3.index)

df_target_3 = df_target_3[['Seq','Site_Name', 'Site_Address', 'Biller_Code']]
df_target_3['Site_Match'] = ''      # Percentage of Target Match
df_target_3['Alliance_Match'] = ''  # Match Ratio
df_target_3['Alliance_Name'] = ''
df_target_3['Alliance_Address'] = ''
df_target_3['Alliance_ID'] = ''

for r in range(len_target3) :                     # Search the Target File
    site_name = df_target_3['Site_Name'][r]
    site_addr = df_target_3['Site_Address'][r]

    match_count =[[0,0],0]                  #[Match Count, Row index)
    # match_ratio = 0
    tmp_name_count = [[0,0],0]
    tmp_addr_count = [[0,0],0]
    tmp_compare = [[0,0],0]

    for i in range(len_source) :        # Search in Source File
        tmp_source_name = df_source['name'][i]
        tmp_source_addr = df_source['addr'][i]

        if len(tmp_source_name) > 0 :
            match_result = matchWord(site_name, tmp_source_name)
            tmp_name_count = [match_result, i]    # Site Name vs Alliance Name

        if len(tmp_source_addr) > 0 :
            match_result = matchWord(site_name, tmp_source_addr)
            tmp_addr_count = [match_result, i]    # Site Name vs Alliance Address

            if len(site_addr) > 0 and 'no address' not in site_addr.strip().lower() and tmp_name_count[0][0] > 0:
                match_addr_result = matchWord(site_addr, tmp_source_addr)        # # Site Address vs Alliance Address
                if match_addr_result[0] > match_result[0] : tmp_addr_count = [match_addr_result, i]

        if tmp_name_count[0][0] >= tmp_addr_count[0][0] : # Compare Name vs Address
            tmp_compare = tmp_name_count
        else :
            tmp_compare = tmp_addr_count

        if tmp_compare[0][0] > match_count[0][0] :    # Filter maximum match count in source
            match_count = tmp_compare
    # Update Excel File
    if match_count[0][0] > 0 :
        df_target_3['Site_Match'].iloc[r] = int(match_count[0][0] / match_count[0][1] * 100)
        df_target_3['Alliance_Match'].iloc[r] = int(match_count[0][0]/match_count[0][2]*100)
        df_target_3['Alliance_Name'].iloc[r] = df_source['name'][match_count[1]]
        df_target_3['Alliance_Address'].iloc[r] = df_source['addr'][match_count[1]]
        df_target_3['Alliance_ID'].iloc[r] = df_source['id'][match_count[1]]
        df_target_3['Biller_Code'].iloc[r] = df_source['id'][match_count[1]]
    else :
        df_target_3['Site_Match'].iloc[r] = match_count[0][0]
        df_target_3['Alliance_Match'].iloc[r] = 0
        df_target_3['Alliance_Name'].iloc[r] = 'NOT FOUND'
        df_target_3['Alliance_Address'].iloc[r] = 'NOT FOUND'
        df_target_3['Alliance_ID'].iloc[r] = 'NOT FOUND'
        df_target_3['Biller_Code'].iloc[r] = 'NOT FOUND'

#-------------------------Target 3 QLD 330----------------------------------

#-------------------------Target 4 SA 227----------------------------------
len_target4 = len(df_target_4.index)

df_target_4 = df_target_4[['Seq','Site_Name', 'Site_Address', 'Biller_Code']]
df_target_4['Site_Match'] = ''      # Percentage of Target Match
df_target_4['Alliance_Match'] = ''  # Match Ratio
df_target_4['Alliance_Name'] = ''
df_target_4['Alliance_Address'] = ''
df_target_4['Alliance_ID'] = ''

for r in range(len_target4) :                     # Search the Target File
    site_name = df_target_4['Site_Name'][r]
    site_addr = df_target_4['Site_Address'][r]

    match_count =[[0,0],0]                  #[Match Count, Row index)
    # match_ratio = 0
    tmp_name_count = [[0,0],0]
    tmp_addr_count = [[0,0],0]
    tmp_compare = [[0,0],0]

    for i in range(len_source) :        # Search in Source File
        tmp_source_name = df_source['name'][i]
        tmp_source_addr = df_source['addr'][i]

        if len(tmp_source_name) > 0 :
            match_result = matchWord(site_name, tmp_source_name)
            tmp_name_count = [match_result, i]    # Site Name vs Alliance Name

        if len(tmp_source_addr) > 0 :
            match_result = matchWord(site_name, tmp_source_addr)
            tmp_addr_count = [match_result, i]    # Site Name vs Alliance Address

            if len(site_addr) > 0 and 'no address' not in site_addr.strip().lower() and tmp_name_count[0][0] > 0:
                match_addr_result = matchWord(site_addr, tmp_source_addr)        # # Site Address vs Alliance Address
                if match_addr_result[0] > match_result[0] : tmp_addr_count = [match_addr_result, i]

        if tmp_name_count[0][0] >= tmp_addr_count[0][0] : # Compare Name vs Address
            tmp_compare = tmp_name_count
        else :
            tmp_compare = tmp_addr_count

        if tmp_compare[0][0] > match_count[0][0] :    # Filter maximum match count in source
            match_count = tmp_compare
    # Update Excel File
    if match_count[0][0] > 0 :
        df_target_4['Site_Match'].iloc[r] = int(match_count[0][0] / match_count[0][1] * 100)
        df_target_4['Alliance_Match'].iloc[r] = int(match_count[0][0]/match_count[0][2]*100)
        df_target_4['Alliance_Name'].iloc[r] = df_source['name'][match_count[1]]
        df_target_4['Alliance_Address'].iloc[r] = df_source['addr'][match_count[1]]
        df_target_4['Alliance_ID'].iloc[r] = df_source['id'][match_count[1]]
        df_target_4['Biller_Code'].iloc[r] = df_source['id'][match_count[1]]
    else :
        df_target_4['Site_Match'].iloc[r] = match_count[0][0]
        df_target_4['Alliance_Match'].iloc[r] = 0
        df_target_4['Alliance_Name'].iloc[r] = 'NOT FOUND'
        df_target_4['Alliance_Address'].iloc[r] = 'NOT FOUND'
        df_target_4['Alliance_ID'].iloc[r] = 'NOT FOUND'
        df_target_4['Biller_Code'].iloc[r] = 'NOT FOUND'

#-------------------------Target 4 SA 227----------------------------------

#-------------------------Target 5 NSW 420----------------------------------
len_target5 = len(df_target_5.index)

df_target_5 = df_target_5[['Seq','Site_Name', 'Site_Address', 'Biller_Code']]
df_target_5['Site_Match'] = ''      # Percentage of Target Match
df_target_5['Alliance_Match'] = ''  # Match Ratio
df_target_5['Alliance_Name'] = ''
df_target_5['Alliance_Address'] = ''
df_target_5['Alliance_ID'] = ''

for r in range(len_target5) :                     # Search the Target File
    site_name = df_target_5['Site_Name'][r]
    site_addr = df_target_5['Site_Address'][r]

    match_count =[[0,0],0]                  #[Match Count, Row index)
    # match_ratio = 0
    tmp_name_count = [[0,0],0]
    tmp_addr_count = [[0,0],0]
    tmp_compare = [[0,0],0]

    for i in range(len_source) :        # Search in Source File
        tmp_source_name = df_source['name'][i]
        tmp_source_addr = df_source['addr'][i]

        if len(tmp_source_name) > 0 :
            match_result = matchWord(site_name, tmp_source_name)
            tmp_name_count = [match_result, i]    # Site Name vs Alliance Name

        if len(tmp_source_addr) > 0 :
            match_result = matchWord(site_name, tmp_source_addr)
            tmp_addr_count = [match_result, i]    # Site Name vs Alliance Address

            if len(site_addr) > 0 and 'no address' not in site_addr.strip().lower() and tmp_name_count[0][0] > 0:
                match_addr_result = matchWord(site_addr, tmp_source_addr)        # # Site Address vs Alliance Address
                if match_addr_result[0] > match_result[0] : tmp_addr_count = [match_addr_result, i]

        if tmp_name_count[0][0] >= tmp_addr_count[0][0] : # Compare Name vs Address
            tmp_compare = tmp_name_count
        else :
            tmp_compare = tmp_addr_count

        if tmp_compare[0][0] > match_count[0][0] :    # Filter maximum match count in source
            match_count = tmp_compare
    # Update Excel File
    if match_count[0][0] > 0 :
        df_target_5['Site_Match'].iloc[r] = int(match_count[0][0] / match_count[0][1] * 100)
        df_target_5['Alliance_Match'].iloc[r] = int(match_count[0][0]/match_count[0][2]*100)
        df_target_5['Alliance_Name'].iloc[r] = df_source['name'][match_count[1]]
        df_target_5['Alliance_Address'].iloc[r] = df_source['addr'][match_count[1]]
        df_target_5['Alliance_ID'].iloc[r] = df_source['id'][match_count[1]]
        df_target_5['Biller_Code'].iloc[r] = df_source['id'][match_count[1]]
    else :
        df_target_5['Site_Match'].iloc[r] = match_count[0][0]
        df_target_5['Alliance_Match'].iloc[r] = 0
        df_target_5['Alliance_Name'].iloc[r] = 'NOT FOUND'
        df_target_5['Alliance_Address'].iloc[r] = 'NOT FOUND'
        df_target_5['Alliance_ID'].iloc[r] = 'NOT FOUND'
        df_target_5['Biller_Code'].iloc[r] = 'NOT FOUND'

#-------------------------Target 5 NSW 420----------------------------------
#-------------------------Target 6 VIC 315----------------------------------
len_target6 = len(df_target_6.index)

df_target_6 = df_target_6[['Seq','Site_Name', 'Site_Address', 'Biller_Code']]
df_target_6['Site_Match'] = ''      # Percentage of Target Match
df_target_6['Alliance_Match'] = ''  # Match Ratio
df_target_6['Alliance_Name'] = ''
df_target_6['Alliance_Address'] = ''
df_target_6['Alliance_ID'] = ''

for r in range(len_target6) :                     # Search the Target File
    site_name = df_target_6['Site_Name'][r]
    site_addr = df_target_6['Site_Address'][r]

    match_count =[[0,0],0]                  #[Match Count, Row index)
    # match_ratio = 0
    tmp_name_count = [[0,0],0]
    tmp_addr_count = [[0,0],0]
    tmp_compare = [[0,0],0]

    for i in range(len_source) :        # Search in Source File
        tmp_source_name = df_source['name'][i]
        tmp_source_addr = df_source['addr'][i]

        if len(tmp_source_name) > 0 :
            match_result = matchWord(site_name, tmp_source_name)
            tmp_name_count = [match_result, i]    # Site Name vs Alliance Name

        if len(tmp_source_addr) > 0 :
            match_result = matchWord(site_name, tmp_source_addr)
            tmp_addr_count = [match_result, i]    # Site Name vs Alliance Address

            if len(site_addr) > 0 and 'no address' not in site_addr.strip().lower() and tmp_name_count[0][0] > 0:
                match_addr_result = matchWord(site_addr, tmp_source_addr)        # # Site Address vs Alliance Address
                if match_addr_result[0] > match_result[0] : tmp_addr_count = [match_addr_result, i]

        if tmp_name_count[0][0] >= tmp_addr_count[0][0] : # Compare Name vs Address
            tmp_compare = tmp_name_count
        else :
            tmp_compare = tmp_addr_count

        if tmp_compare[0][0] > match_count[0][0] :    # Filter maximum match count in source
            match_count = tmp_compare
    # Update Excel File
    if match_count[0][0] > 0 :
        df_target_6['Site_Match'].iloc[r] = int(match_count[0][0] / match_count[0][1] * 100)
        df_target_6['Alliance_Match'].iloc[r] = int(match_count[0][0]/match_count[0][2]*100)
        df_target_6['Alliance_Name'].iloc[r] = df_source['name'][match_count[1]]
        df_target_6['Alliance_Address'].iloc[r] = df_source['addr'][match_count[1]]
        df_target_6['Alliance_ID'].iloc[r] = df_source['id'][match_count[1]]
        df_target_6['Biller_Code'].iloc[r] = df_source['id'][match_count[1]]
    else :
        df_target_6['Site_Match'].iloc[r] = match_count[0][0]
        df_target_6['Alliance_Match'].iloc[r] = 0
        df_target_6['Alliance_Name'].iloc[r] = 'NOT FOUND'
        df_target_6['Alliance_Address'].iloc[r] = 'NOT FOUND'
        df_target_6['Alliance_ID'].iloc[r] = 'NOT FOUND'
        df_target_6['Biller_Code'].iloc[r] = 'NOT FOUND'

#-------------------------Target 6 VIC 315----------------------------------


with pd.ExcelWriter(r'result.xlsx', index=False, na_rep='') as writer :
    df_target_1.to_excel(writer, sheet_name='WA 326', index=False)
    df_target_2.to_excel(writer, sheet_name='ACT 464', index=False)
    df_target_3.to_excel(writer, sheet_name='QLD 330', index=False)
    df_target_4.to_excel(writer, sheet_name='SA 227', index=False)
    df_target_5.to_excel(writer, sheet_name='NSW 420', index=False)
    df_target_6.to_excel(writer, sheet_name='VIC 315', index=False)
