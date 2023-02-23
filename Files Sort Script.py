import os

path = ""

files = os.listdir(path)
list_of_names_lst = []

for file in files:
    only_name = file.split(".")[0]
    split_name_lst = only_name.split("-")
    list_of_names_lst.append(split_name_lst)

sorted(list_of_names_lst, key=lambda x: (x[1], x[0]))
i = 0
for name in list_of_names_lst:
    i += 1
    if len(name[0]) == 1:
        name[0] = "0" + name[0]
    if len(name[1]) == 1:
        name[1] = "0" + name[1]
    print(str(i) + ") " + name[0] + "." + name[1] + "." + name[2])
