# -*- coding: UTF-8 -*-



def titleToNumber(s):
        a={chr(i):i-64 for i in range(65,91)}
        num=0
        b=s[::-1]
        for i,j in enumerate(b):
            num+=a[j]*(26**i)
        return num

    
print(titleToNumber("X")-titleToNumber("C"))
