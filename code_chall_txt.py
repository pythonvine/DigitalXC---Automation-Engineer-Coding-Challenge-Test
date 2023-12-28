#Install the required packages
#pip install pandas openpyxl

import pandas as pd
import json

def extract_group(file, colomn_name, group_name, remove_string):
    data = pd.read_excel(file)
    
    comments = data[colomn_name]
    # 
    group_count={}
    group_count.update({"Group_name": "Number of occurrences"})
    output = 'output.txt'
    # with open('output.txt', 'w') as file:
    #      file.write()
    #      file.close()

    for comment in comments:
        # print(comment)
        if group_name in comment:
            group = comment.split(group_name)[1].strip().replace(remove_string, '')
            group_list = [group.strip() for group in group.split(",")]
        
        for group in group_list:
                group_count[group] = group_count.get(group, 0) + 1

    # df = pd.DataFrame(list(group_count.items()), columns=["Group_name", "Number of occurrences"])
    # df.to_excel('output.xlsx', index=False)
    for group, count in group_count.items():

# Save to a text file
        with open('output.txt', 'a') as file:
            file.write(f"{group}\t\t\t{count}\n")
            file.close()
   
file = 'coding_challenge_test.xlsx'
target_string = "Groups : [code]<I>"
colomn = 'Additional comments'
remove_string = '</I>[/code]'

extract_group(file, colomn, target_string, remove_string)
