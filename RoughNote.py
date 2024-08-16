def printTable(tab):
    for j in range(len(tab[0])):
        for i in range(len(tab)):
            m = max([len(s) for s in tab[i]])
            try:
                print(tab[i][j].rjust(m), end=' ')
            except IndexError:
                print(' '.rjust(m),end=' ')
        print('')

tableData = [['apples', 'oranges', 'cherries', 'banana'],
             ['Alice', 'Bob', 'Carol', 'David'],
             ['dogs', 'cats', 'moose', 'goose'],
             ['dogs', 'cats', 'moose'],
             ['apples', 'oranges', 'cherries', 'banana']]


printTable(tableData)
