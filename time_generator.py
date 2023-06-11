
def gentime(start_hour,stop_hour,step_hour,start_minute,stop_minute,step_minute):
    r = 0
    for hour in range(start_hour,stop_hour,step_hour):
        for minute in range(start_minute,stop_minute,step_minute):
            # Format the hour and minute with leading zeros
            formatted_hour = f"{hour:02d}"
            formatted_minute = f"{minute:02d}"
            
            # Print or use the formatted hour and minute
            print(f"{formatted_hour}:{formatted_minute}")
            r+=1
    print(r)

# start_hour 0,stop_hour 23(+1),step_hour 1,start_minute 0,stop_minute 30(+1),step_minute 30

def main(start_hour,end_hour,start_min,end_min):
    if end_hour == 0:
        if start_hour == 0 or start_hour > 0:
            gentime(start_hour,24,1,start_min,end_min+1,30)
            gentime(0,1,1,start_min,end_min,30)
    elif end_hour - start_hour > 0:
        gentime(start_hour,end_hour+1,1,start_min,end_min+1,30)
    else:
        gentime(start_hour,24,1,start_min,end_min+1,30)
        gentime(0,end_hour+1,1,start_min,end_min+1,30)
        
main(1,23,0,30)