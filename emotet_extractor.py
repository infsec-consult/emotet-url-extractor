import base64
import string
import sys
import argparse

def strings(filename, min=400):
    with open(filename, errors="ignore") as f:  
        result = ""
        for c in f.read():
            if c in string.printable:
                result += c
                continue
            if len(result) >= min and result[0]=="p":
                yield result
            result = ""
        if len(result) >= min:  
            yield result


def clean_up(script_in):
    script_out = []

    for i in script_in:
        i=i.replace("'","")
        i=i.replace("+","")
        i=i.replace('"',"")
        script_out.append(i)

    return script_out

def grab_para(script_in):
    start = 0
    counter=0
    parameter = []
    
    while start < len(script_in) and counter > -1:
        counter=script_in.find ("(",start)
        if counter > -1:
            start=counter
            counter = script_in.find(")",start)
            if counter > -1:
                parameter.append(script_in[start+1:counter])
                start = counter+1

         
    return parameter

            


if __name__ == "__main__":


    parser = argparse.ArgumentParser(description='Extract urls from emotet .doc-files')
    parser.add_argument('file', action="store", help='the emotet .doc-file')
    parser.add_argument('-s', action="store_true", default=False, dest='print_script', help='print the emotet vba-script after base64-decode')
    parser.add_argument('-d', action="store", dest="delimiter", help="manually set the split delimiter" )
    parser.add_argument('-i', action="store_true", dest="print_input", help="print the obfuscatet script as axtrcted from the .doc")

    args = parser.parse_args()






    filename = args.file
    input = ""
    input = input.join(strings(filename))
    input = input.strip()

    if args.print_input:
        print(input)
        print ("--------------")

    start = 0
    end=0
    counter=0
    found=0

    while (found==0 and end+counter<len(input)):
        if args.delimiter:
            s=input.split(args.delimiter)            
        else:
            while (end+counter<len(input) and input[end+counter]!="o"):
                counter=counter+1
        
            end=end+counter
            s = input.split(input[start+1:end])

        o = ""
        o = o.join(s)

        if o[0:13].lower().startswith("powershell "):
            found=1
        else:
            end=0
            counter=counter+1

    if found==1:
        s=o[14:]


        #remove trailing chars at the end, so we get a base64 string
        cut= len(s)%4
        if cut!=0:
            s=s[:-cut]            

 
        script_bytes = base64.b64decode(s)
        script = script_bytes.decode('UTF-16LE')

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
            
