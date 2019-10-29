import xlrd
f = open('12334.txt','w')
workbook = xlrd.open_workbook('1.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')
num_rows = worksheet.nrows - 1
def wri(v,nb):
            
            k=nb
            j=nb
            num=nb-1
            count=0
            lar=nb+n
            
            f.write('\n')
            f.write('{"%s":['%(c))
            for k in range(nb,v):
                k=nb
                m=worksheet.cell(k,3)
                
                p=m.value
                
        
                
                """print(k)"""
                j=k
                if(k==lar):
                    break
        
        
                for j in range(k,v+1):
                    
                    
                    
                    no=worksheet.cell(j,3)
                    l=no.value
                    zero=worksheet.cell(j,0).value
                    four=worksheet.cell(j,4).value
                    five=worksheet.cell(j,5).value
                    seven=worksheet.cell(j,7).value
                    eight=worksheet.cell(j,8).value
                    if(j==k):
                        f.write('\n')
                        f.write('{"%s":['%(l))
                        f.write('\n')
                        
                        
                        f.write('{"State":"%s","City":"%s","Loction":"%s","AscName":"%s ","AscAddress":"%s","ContactNumber":"%s","Lat":"%s","Long":"%s"}'%(c,l,l,zero,four,five,seven,eight))
                        
                        num=num+1
                        continue
                
                    elif l==p and j>k:
                        f.write(',')
                        
                        num=num+1
                        count=count+1
                        
                        
                        
                        
                        f.write('\n')
                        f.write('{"State":"%s","City":"%s","Loction":"%s","AscName":"%s ","AscAddress":"%s","ContactNumber":"%s","Lat":"%s","Long":"%s"}'%(c,l,l,zero,four,five,seven,eight))
                        continue
                    else:
                        if(j>=v):
                            f.write(']}')
                        else:
                            f.write(']},')
                            
                            break
                       
                    
                
                
                
            
                
                
                nb=num+1
                """print(nb)"""

countt=0    
while(1):
   
    print("Do you want to continue (yes/no)")
    check=input()
    

    if(check=='yes'):
        if(countt>0):
            f.write(',')
        countt=countt+1
        print("enter the starting serial number of required state in excel")
        var=int(input())
        i=var-1
        if(i==1):
            f.write('{"state":[')
        c=""
        n=0

        v=i+80
        if(v>num_rows):
            
            v=num_rows
            

        nb=i
        w=worksheet.cell(i,2)

        for i in range(nb,v+1):
            if(i==nb):
                c=w.value
        
            d=worksheet.cell(i,2)
            if(d.value==c):
                n=n+1
                
                
                if(i==v):
                    
                    wri(v,nb)
                    f.write(']}')
                    break
            
        
            else:
                v=nb+n
                wri(v,nb)
                f.write(']}')
                print("state done :",c)
                print("enter next input as :",v+1)
                break
    
        
        
    
    else:
        
        f.write(']}')
        f.close()
        break

