import base64
import string
import sys

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


if __name__ == "__main__":

    if len(sys.argv) != 2:
        print ("exactly 1 filename needed")
        exit(0)

    filename = str(sys.argv[1])
    input = ""
    input = input.join(strings(filename))
    input = input.strip()


    start = 0
    end=0
    counter=0
    found=0

    while (found==0 and end+counter<len(input)):
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


        s1 = script.find("http")
        s2 = script.find("http",s1+1)
        sep = script[s2-1] #Separator for the split

        while s2 != -1:        
            s2=script.find(sep,s1+1)
            if s2 != -1:
                print(script[s1:s2])
                s1=s2+1
            else: #The last URL is not finished by the separator, so lets find the ' or "" at the end of the string
                t1= script.find("'", s1)
                t2= script.find("\"", s1)
                if t1 < t2:
                    print (script[s1:t1])
                else:
                    print(script[s1:t2])

  

