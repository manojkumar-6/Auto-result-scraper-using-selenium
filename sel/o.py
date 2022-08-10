from autoscraper import AutoScraper
url="https://www.amazon.in/s?k=mobiles&ref=nb_sb_noss_2"
s=AutoScraper()
r=s.build(url)
l=[]
p=[]
w=r[0]
o=w[0]
for i in r:
            if o in i:
                p.append(i)
            else:
                l.append(i)
g=zip(l,p)
for i in g:
    print(i)