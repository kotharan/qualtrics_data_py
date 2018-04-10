chars=[]
inputchars=input('put in excel column letters: ')
#give me a bit more time and ill have this program able to actually turn excel column letters into usable column numbers.
for x in inputchars:
    chars.append(x)

nums=[]
for x in chars:
    nums.append((ord(x)-96))

print(('that is column #'+ str(sum(nums))+ ', assuming you start from column 0'))
