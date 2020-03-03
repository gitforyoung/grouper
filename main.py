import pandas as pd
import random, math
import os, sys

# Initial Screen Printing
initial_text = """
======== Group Maker v1.0 ========
자동 그룹 생성 프로그램입니다.
프로그램을 명단 파일(excel)과 같은
경로에 저장해 주세요.
==================================
                        - Young
"""
print(initial_text)
os.system('pause')
os.system('cls')

guide_text = """
============ ! 중요 ! ============
excel은 다음과 같은 양식으로 작성 후
파일명은 "list.xlsl"로 저장해 주세요.
각 팀의 멤버 수 만큼 리더를 선정해 주세요.

   ID   |   이름   |   리더
 S-112  |  김민준  |
 S-013  |  이서준  |    o
 A-102  |  최예준  |
==================================
"""
print(guide_text)
os.system('pause')
os.system('cls')

# Exception
try:
    data = pd.read_excel("./list.xlsx")
except FileNotFoundError:
    print('그룹을 생성할 데이터(list.xlsl)가 없습니다.\n데이터를 저장한 후 다시 실행해 주세요.\n')
    os.system('pause')
    sys.exit()
if len(data)==0:
    print('데이터의 내용이 없습니다.\n내용을 추가한 후 다시 실행해 주세요.\n')
    os.system('pause')
    sys.exit()

# Get input data
while True:
    try:
        team_size = int(input("각 팀의 멤버 수를 입력하세요 : "))
        if team_size==0:
            raise ZeroDivisionError
        break;
    except (ValueError, ZeroDivisionError):
        print("\n0이 아닌 숫자를 입력하세요.\n")
        os.system('pause')
        os.system('cls')
 
group_size = math.ceil( (data.index[-1] + 1) / team_size )
print("\n입력한 각 팀의 멤버 수는",team_size, ". 생성할 팀 개수는",group_size,".\n")
os.system('pause')

# Classify members that are selected / not selected.
picked_members = []
other_members =[]
column_name = list(data)
ranked_data = data.sort_values(by=column_name[2])

for i in range(group_size):
    picked_members.append(ranked_data.iloc[i,0])
random.shuffle(picked_members)

for i in data.index:
    if data.iloc[i,0] not in picked_members:
        other_members.append(data.iloc[i,0])
random.shuffle(other_members)

# Calculate number of teams
groupNumber = []
count = 0
for i in range(len(other_members)):
    groupNumber.append((count//(team_size-1))+1)
    count = count + 1

# Make result
df = pd.DataFrame()
df['groups'] = groupNumber
df['ids'] = other_members

final_group = []
for i in range(group_size):
    final_group.append([])

for i in range(len(final_group)):
    final_group[i].append(picked_members[i])
    for j in df.index:
        if df.iloc[j,0]==(i+1):
            final_group[i].append(df.iloc[j][1])
    random.shuffle(final_group[i])
    final_group[i].insert(0,i+1)

# Append seat numbers to original data frame
seat_number = []
for i in data.index:
    for team in range(len(final_group)):
        for member in range(1, len(final_group[team])):
            if data.iloc[i,0] == final_group[team][member]:
                input_data = str(team+1) + "-" + str(member)
                seat_number.append(input_data)
data["Seat"]=seat_number

# Make "team" table
for team in range(len(final_group)):
    for member in range(1, len(final_group[team])):
        for i in data.index:
            if data.iloc[i,0] == final_group[team][member]:
                input_data = "("+str(final_group[team][member])+")"+data.iloc[i,1]
                final_group[team][member] = input_data

labels=['Team', 'membes']
final_df = pd.DataFrame.from_records(final_group)
final_df.rename(columns={0:"Team"},inplace=True)
final_df.set_index('Team',inplace=True)

# Print results
print (final_df)
data.set_index(column_name[0],inplace=True)
print(data)

# Save results
final_df.to_excel("./team.xlsx")
data.to_excel("./seat.xlsx")