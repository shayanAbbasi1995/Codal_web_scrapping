x = 'from codal_main_function import *\ncodal_search_for_links(1,10)\n'
for i in range(1, 426):
    if i < 10:
        start='00'+str(i)
    elif i < 100:
        start = '0'+str(i)
    else:
        start=str(i)
    y=x.replace('(1,10)','('+str(i)+','+str(i)+')')
    f=open('execute_coda_'+start+'.py', mode='w')
    f.write(y)
    f.close()