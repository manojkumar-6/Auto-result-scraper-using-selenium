import collections
from collections import Counter
s=[-2,1,-3,4,-1,2,1,-5,4]
v =[]
p=[]
f=0
min="a"
for i in range(0,len(s)):
    for j in range(i+1,len(s)):
        sa=s[i:j+1]
        print(sa)
        c=sum(sa)
        #print(c)
        if min=="a":
            min=c
        elif c > min:
            min=c
            p.clear()
            p.append(sa)
        
        #print(min)
print(p,f)
print(min)
z=0
for i in s:
    z+=i
    if z<0:
        z=0
    else:
        z+=i
print(z)



