# use Map.txt
# print and rewrite abbr. for map name


with open('Map.txt','r') as file:
    lines = file.readlines()

new_lines = []
for line in lines:
    lineList = line.split()
    map_lvl = lineList[0] # map level in 1 column
    map_1name = lineList[1] # map first name in 2 column
    map_1name_character = map_1name[0] # first character of map first name
    map_2name = lineList[2] # map second name
    map_2name_character = map_2name[0]
    map_abbr = map_1name_character + map_2name_character
    new_line = map_lvl +'\t'+ map_1name +' '+ map_2name +'\t'+ map_abbr + '\n'
    new_lines.append(new_line)

with open('Map.txt','w') as file:
    file.writelines(new_lines)