
def addthemup(n1,n2,n3,n4,n5,n6,n7):

    try:
        n1=float(n1)
    except (ValueError, IndexError):
        n1=0
    try:
        n2=float(n2)
    except (ValueError, IndexError):
        n2=0
    try:
        n3=float(n3)
    except (ValueError, IndexError):
        n3=0
    try:
        n4=float(n4)
    except (ValueError, IndexError):
        n4=0
    try:
        n5=float(n5)
    except (ValueError, IndexError):
        n5=0
    try:
        n6=float(n6)
    except (ValueError, IndexError):
        n6=0
    try:
        n7=float(n7)
    except (ValueError, IndexError):
        n7=0
    out= (n3+n4+n5+n6+n7)
    return str(out)

def multiplythethree(n1,n2,n3,n4,schoolname):
        try:
            n1=float(n1)
        except (ValueError, IndexError):
            n1=0
        try:
            n2=float(n2)
        except (ValueError, IndexError):
            n2=0
        try:
            n3=float(n3)
        except (ValueError, IndexError):
            n3=0
        try:
            n4=float(n4)
        except (ValueError, IndexError):
            n4=0

        if (n1+n2)==0:
            print(schoolname,': no student numbers found')
            return 'no student numbers found'
        else:
            students =n1+n2
        if (n3==0)&(n4 !=0):
            price = n4
            #print('using n4=',str(n4), ' total= ',str(students*price))
            return str(students*price)
        elif n3!=0:
            price= n3
            #print('using n3=',str(n3), ' total= ',str(students*price))
            return str(students*price)
        elif (n3==0)&(n4 ==0):
            print(schoolname,': no price per student found')
            return 'no price per student found'
        else:
            print(schoolname,': unknown error')
            return 'unknown error'
