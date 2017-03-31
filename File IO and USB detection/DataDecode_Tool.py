import os,struct,string,time,xlwt,shutil 
from ctypes import windll

HEAD_STRUCT='=LQ'
TRACK_NUM = 9999+1
FILE_NUM = 8+1

def get_driveStatus():
    devices = []
    record_deviceBit = windll.kernel32.GetLogicalDrives()#The GetLogicalDrives function retrieves a bitmask
                                                         #representing the currently available disk drives.
    for label in string.uppercase : #The uppercase letters 'A-Z'
        if record_deviceBit & 1:
            devices.append(label)
        record_deviceBit >>= 1
    return devices


def read_DataDecode(r_path,w_path,outfileName):
    os.chdir(w_path)
    fw = open(outfileName+".txt",'w')
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Statistics",cell_overwrite_ok=True)

    os.chdir(r_path)
    for n in range(0,FILE_NUM):
        fileName = str("%02d" % n )+'.dat'
        if os.path.exists(fileName):
            print fileName +' is exist\n'
            fr = open(fileName,'rb')
            os.chdir(w_path)
            fw.write(fileName+':\n')
            fw.write('------------------------\n')
            os.chdir(r_path)
            a=0
            for m in range(0,TRACK_NUM):
                (count,time)=struct.unpack(HEAD_STRUCT,fr.read(4+8))
                sheet1.write(0, 0+n*4, fileName)
                sheet1.write(2, 0+n*4, "Track NO")
                sheet1.write(2, 1+n*4, "Count")
                sheet1.write(2, 2+n*4, "Time(s)")
                sheet1.write(2, 3+n*4, " ")
                if(count and m>0):
                    print "track:%d ,count:%d, time:%d"%(m,count,time)
                    fw.write("track:%d ,count:%d, time:%d\n"%(m,count,time))
                    sheet1.write(a+3,0+n*4,m)
                    sheet1.write(a+3,1+n*4,count)
                    sheet1.write(a+3,2+n*4,time)
                    sheet1.write(a+3,3+n*4, " ")
                    a=a+1
                elif(m==0):
                    print "Total language count:%d"%(count)
                    print'------------------------\n'
                    fw.write("Total language count:%d\n" % count)
                    fw.write('------------------------\n')
                    sheet1.write(1, 0+n*4, "Total:")
                    sheet1.write(1,1+n*4,count)
            fr.close()
            fw.write('\n')
        else:
            print fileName +' is not exist\n'
        print'\n'
    fw.close()
    os.chdir(w_path)
    book.save("DataDecode_%s"%(outfileName)+".xls")

def copy_DataDecode(r_path,w_path):
    os.chdir(r_path)
    print "Collecting....."
    for n in range(0,FILE_NUM):
        fileName = str("%02d" % n )+'.dat'
        if os.path.exists(fileName):
            os.chdir(w_path)
            if os.path.exists("DataDecode/"+fileName):
                fw = open("DataDecode/"+fileName,'rb+')
                init_flag=1
            else:
                fw = open("DataDecode/"+fileName,'wb+')
                init_flag=0
                
            os.chdir(r_path)
            fr = open(fileName,'rb')
            for m in range(0,TRACK_NUM):
                (count_r,time_r)=struct.unpack(HEAD_STRUCT,fr.read(4+8))
                if count_r ==0:
                    continue
                os.chdir(w_path)
                if init_flag==1:
                    fw.seek(struct.calcsize(HEAD_STRUCT)*m)
                    (count_w,time_w)=struct.unpack(HEAD_STRUCT,fw.read(4+8))
                    fw.seek(0,0)
                   # print "\n\n-------------------------------\n"
                    count = count_r+count_w
                    time = time_r+time_w
                    fw.seek(struct.calcsize(HEAD_STRUCT)*m)
                    data = struct.pack(HEAD_STRUCT,count,time)
                   # print "file:%d,track:%d ,count:%d, time:%f"%(n,m,count,time)
                   # print "-------------------------------\n\n"
                else:
                    data = struct.pack(HEAD_STRUCT,count_r,time_r)
                fw.write(data)
                os.chdir(r_path) 
            fr.close()
    fw.close()   

def creat_emptyFile(w_path):
    os.chdir(w_path)
    for n in range(1,FILE_NUM):
        fileName = str("%02d" % n )+'.dat'
        fw = open("DataDecode/"+fileName,'wb')
        fw.truncate(struct.calcsize(HEAD_STRUCT)*TRACK_NUM)
        fw.close()
        
if __name__ == '__main__':
    now_path = os.getcwd()
    original = set(get_driveStatus())
    cmd =0
    while(cmd!='q'):
        print "\n"
        print "Refresh usb detection (r)"
        print "Start to analysis the every file of data (a)"
        print "Sum all file of data (s)"
        print "Export final data (e)"
        print "Delete final file of data(D)"
        print "Quit (q)\n"
        cmd = raw_input('Please input command:')
        print " " 
        if(cmd == 'q'):    
            print('exit program')
        elif(cmd =='r'):
            add_device =  set(get_driveStatus())- original
            subt_device = original - set(get_driveStatus())

            if (len(add_device)):
                print "There were %d"% (len(add_device))
                for drive in add_device:
                    print "The USB added: %s." % (drive)
                 
            elif(len(subt_device)):
                print "There were %d"% (len(subt_device))
                original = set(get_driveStatus())
                for drive in subt_device:
                    print "The USB remove: %s." % (drive)
            else:
                print("No anything detected!\n")
        elif(cmd =='a'):
            usb_device =  set(get_driveStatus())- original
            if (not len(usb_device)):
                print ("No any device!\n")
            else:    
                for drive in usb_device:
                    r_path = drive +":\binary"
                    if os.system("cd " + r_path) == 0:
                        read_DataDecode(r_path,now_path,"DataDecode_%s"%(drive))
        elif(cmd =='s'):
            if not os.path.exists("DataDecode"):
                os.makedirs("DataDecode")
                creat_emptyFile(now_path)
            usb_device =  set(get_driveStatus())- original
            if (not len(usb_device)):
                print ("No any device!\n")
            else:
                for drive in usb_device:
                    r_path = drive +":\binary"
                    if os.system("cd " + r_path) == 0:
                        copy_DataDecode(r_path,now_path)
        elif(cmd =='e'):
            os.chdir(now_path)
            if not os.path.exists("DataDecode"):
                print("You are not doing any collect!\n")
                continue
            if os.system("cd " + now_path+"/DataDecode") == 0:
                read_DataDecode(now_path+"/DataDecode",now_path,"Data_Decode")
        elif(cmd =='D'):
            if os.path.exists("DataDecode"):
                os.chdir(now_path)
                shutil.rmtree("DataDecode")
                print("All file of data collect is deleted")
            else:
                print("No any file of data collect!")
        else:
            print ("Unknown command!")
    
