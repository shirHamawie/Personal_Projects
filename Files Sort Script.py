import os


def print_green(output):
    print("\033[92m {}\033[00m".format(output))


def sort_n_print(path):
    files = os.listdir(path)
    list_of_names_lst = []

    for file in files:
        only_name = file.split(".")[0]
        split_name_lst = only_name.split("-")
        list_of_names_lst.append(split_name_lst)

    list_of_names_lst = sorted(list_of_names_lst, key=lambda x: (int(x[1]), int(x[0])))

    i = 0
    print_green(path.split("\\")[-1])
    for name in list_of_names_lst:
        i += 1
        if len(name[0]) == 1:
            name[0] = "0" + name[0]
        if len(name[1]) == 1:
            name[1] = "0" + name[1]
        print(str(i) + ") " + name[0] + "." + name[1] + "." + name[2])


main_path = "C:\\Users\\t-shamawie\\Videos\\Recordings"

folders = os.listdir(main_path)
for folder in folders:
    temp = os.path.join(main_path, folder)
    try:
        sort_n_print(temp)
    except:
        continue
