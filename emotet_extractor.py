import base64
import string
import sys
import argparse
import olefile
import oledump




def get_printable_strings(string_a, min=0):
    all_strings = []
    for i in string_a:
        #result = ""
        for c in i:
            if c in string.printable:
                all_strings.append(c)


    return all_strings

        

def clean_up(script_in):
    script_out = []

    for i in script_in:
        i=i.replace("'","")
        i=i.replace("+","")
        i=i.replace('"',"")
        i=i.replace("(","")
        i=i.replace(')',"")
        script_out.append(i)

    return script_out


#TODO: Different delimiters for parameters
def grab_para(script_in):
    start = 0
    counter=0
    position = 0
    parameter = []
    
    while start < len(script_in) and position > -1:
        position=script_in.find ("(",position)
        counter = 0
        if position > -1:
            start=position
            
            while (script_in[position] != ")" or counter != 1) and position<len(script_in):
                if script_in[position] == "(":
                    counter=counter+1
                if script_in[position] == ")":
                    counter=counter-1
                position = position+1
            
            if position > -1:
                parameter.append(script_in[start+1:position])
                start = counter+1
        
  
    return parameter


def C2SIP3(string):
    if sys.version_info[0] > 2:
        if type(string) == bytes:
            a= ''.join([chr(x) for x in string])

            return a
        else:
            return string
    else:
        return string


          


if __name__ == "__main__":


    parser = argparse.ArgumentParser(description='Extract urls from emotet .doc-files')
    parser.add_argument('file', action="store", help='the emotet .doc-file')
    parser.add_argument('-s', action="store_true", default=False, dest='print_script', help='print the emotet vba-script after base64-decode')
    parser.add_argument('-d', action="store", dest="delimiter", help="manually set the split delimiter" )
    parser.add_argument('-i', action="store_true", dest="print_input", help="print the obfuscatet script as axtrcted from the .doc")

    args = parser.parse_args()

    filename = args.file
    input_strings = []
    printable_strings = ""
    ole = olefile.OleFileIO(open(filename, 'rb').read())
    decoders=""


   
    for _, _, _, _, stream in oledump.OLEGetStreams(ole,False):
        try:       
            input_strings.append(C2SIP3(stream))
        except:
            next
 



    printable_strings = printable_strings.join(get_printable_strings(input_strings))
    printable_strings = printable_strings.strip()


    if args.print_input:
        print(printable_strings)
        print ("--------------")


    start=0
    end=0
    found=False 
    lookup="powershell"
    l_counter=1
    possible_script = []


    

    while (l_counter < 10):
        while (start != -1):
            if args.delimiter:
                s=printable_strings.split(args.delimiter)            
            else:
                print("Von "+lookup[0:l_counter])
                print("Bis "+lookup[l_counter:2*l_counter])
                start = printable_strings.find(lookup[0:l_counter],start)
                end = printable_strings.find(lookup[l_counter:2*l_counter],start+end)
                print(l_counter)
                print("Start: "+str(start))
                print("End: "  +str(end))
                
                #if start==36363:
                #    print (printable_strings[start+l_counter:end])

                if ( start+l_counter<end):
                    #print("T= "+printable_strings[start+l_counter:end])
                    s = printable_strings.split(printable_strings[start+l_counter:end])
                    
                else:
                    s= []
                    


            o = ""
            o = o.join(s)

            start2=o.lower().find("powershell")
            o=o[start2:]



            if o[0:13].lower().startswith("powershell "):
                found=True
                               
                possible_script.append(o)
                print("Script added")

            if end != -1:
                print("end+1")
                end = end +1

            elif start != -1:
                print("start+1")
                start=start+1
                end=0

            print(start)
            print(end)

        l_counter = l_counter+1  
        start=0
        end=0

    print(len(possible_script))




    if found:
        for i in possible_script:
            
            end=i.find("=",start)
            end2 = i.find(",",start)
            print ("End"+str(end))
            print (end2)
            if end2 < end:
                end=end2
            else:
                end=end+1

            print(end)
            s=i[14:end]
            print(s)

            #remove trailing chars at the end, so we get a base64 string
            cut= len(s)%4
            print(cut)
            if cut==3:
                s+="="            

    
            try:
                script_bytes = base64.b64decode(s)
                script = script_bytes.decode('UTF-16LE')
                print(script)
            except:
                print("Decode error")
                continue

            if args.print_script:
                print(script)
                print ("--------------")

            

            urls=[]
            
            arg=clean_up(grab_para(script))
            
            for i in arg:
                i.lower()
                if i.find("http") > -1:
                    s1 = i.find("http")
                    s2 = i.find("http",s1+1)
                    sep = i[s2-1] #Separator for the split

                    while s2 != -1:        
                        s2=i.find(sep,s1+1)
                        if s2 != -1:
                            urls.append(i[s1:s2])
                            s1=s2+1
                        else: #The last URL is not finished by the separator, so print the rest
                            urls.append(i[s1:]) 

            if len(urls) == 0:
                
                s1 = script.find("http")
                s2 = script.find("http",s1+1)
                sep = script[s2-1] #Separator for the split

                while s2 != -1:        
                    s2=script.find(sep,s1+1)
                    if s2 != -1:
                        urls.append(script[s1:s2])
                        s1=s2+1
                    else: #The last URL is not finished by the separator, so lets find the ' or "" at the end of the string
                        t1= script.find("'", s1)
                        t2= script.find("\"", s1)
                        if t1 < t2:
                            urls.append(script[s1:t1])
                        else:
                            urls.append(script[s1:t2])                    
        
            if len(urls) > 0:
                for i in urls:
                    print(i)            
            else:
                print ("No URLs found")
    else:
        print("No VBA-script found")            
