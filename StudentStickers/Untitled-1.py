s = "8 9 10 "
x=0
while True: 
    try:
        print(s.split(" ")[x])
        x+=1
    except:
        break
print("done")
    
