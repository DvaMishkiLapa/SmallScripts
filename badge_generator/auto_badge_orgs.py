def replace_names(fio, line):
    new_line = line
    for x, y in zip(['orgsurname', 'orgname', 'orgpatronymic'], [0, 1, 2]):
        new_line = new_line.replace(x, fio[y])
    return new_line


def names_gen(file):
    names_list = []
    with open(file, 'r', encoding='utf-8') as f:
        for line in f:
            fio_list = line[:-1].split(' ')
            if len(fio_list) > 4:
                fio_list[3] += ' ' + fio_list.pop(4)
            names_list.append(fio_list)
        return names_list


names_list = names_gen('list_orgs.txt')
iterator = 0
f_write = open('auto_orgs.fodg', 'w', encoding='utf-8')
with open('orgs.fodg', 'r', encoding='utf-8') as f_read:
    for line in f_read:
        if ('orgname' in line) and len(names_list) > iterator:
            f_write.write(replace_names(names_list[iterator], line))
            iterator += 1
        else:
            f_write.write(line)
f_write.close()
