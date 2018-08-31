import string


def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) +1
    return num

run=1
while run ==1:
    inputchars=input('put in excel column letters: ')
    if not inputchars[0] in string.ascii_letters:
        run = 0
        break

    number=col2num(inputchars)-1

    print(('that is column _________________#'+ str(number)+ '_________________, assuming you start from column 0'))
