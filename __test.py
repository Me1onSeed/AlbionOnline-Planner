with open('Map.txt', 'r') as f:  
    for line in f:  
        columns = line.split()  # 按空格分割行内容  
        column1 = columns[1]  # 获取第3列数据  
        column2 = columns[2]  # 获取第3列数据  
        column3 = columns[3]  # 获取第4列数据  
        print("\""+column1+" "+column2+" - "+column3+"\",")