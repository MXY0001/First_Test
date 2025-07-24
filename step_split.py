

def split_1(list_value):
    num = 0
    index = []
    for i in range(len(list_value)):
        if i == len(list_value) - 1:
            index.append(list_value[num:len(list_value)])
            continue
        if list_value[i + 1] - list_value[i] > 1:
            index.append(list_value[num:i + 1])
            num = i + 1
    return index

