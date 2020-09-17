import base64
import string
import sys
import argparse
import olefile
import oledump




def get_printable_strings(string_input, min=0):
    printable_strings = []
    for i in string_input:        
        for c in i:
            #if c in string.digits or c in string.ascii_letters or c in string.punctuation:
            if c in string.printable:
                printable_strings.append(c)


    return printable_strings

        

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
def get_parameters(script_in):
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


def b2s(string):
    if type(string) == bytes:
        a= ''.join([chr(x) for x in string])
        return a
    else:
        return string


          


if __name__ == "__main__":


    parser = argparse.ArgumentParser(description='Extract urls from emotet .doc-files')
    parser.add_argument('file', action="store", help='the emotet .doc-file')
    parser.add_argument('-s', action="store_true", default=False, dest='print_script', help='print the emotet vba-script after base64-decode')
    parser.add_argument('-d', action="store", dest="delimiter", help="manually set the split delimiter" )
    parser.add_argument('-g', action="store_true", default=False, dest="greedy", help="More detailed check but longer runtime" )

    args = parser.parse_args()

    filename = args.file
    input_strings = []
    printable_string = ""
    ole = olefile.OleFileIO(open(filename, 'rb').read())
    decoders=""

    if args.greedy:
        print("Greedy search startet. This may take some time")

   #Grab the content of all streams
    for _, _, _, _, stream in oledump.OLEGetStreams(ole,False):
        try:       
            input_strings.append(b2s(stream))
        except:
            next
 
    #We are looking for scriptcode, so filter for printable strings
    printable_string = printable_string.join(get_printable_strings(input_strings))
    printable_string = printable_string.strip()

    




    #Find any occurrence of "powershell" with any delimiter inside the word
    #e.g. "p*o*w*e*r*s*h*e*l*l" or "po|we|rs|he|ll" or "powerE3$$d_shell"
    #and split the rest of the string with the delimiter
    
    #possible start of the delimiter powershell in the string
    start=0
    #possible end of the delimiter powershell in the string
    end=0

    #found any script
    found=False 

    lookup="powershell"

    #Length of the substrings of "powershell" we are looking for 
    #first we look for 1 character substrings "p" and "o"
    #any characters between "p" and "o" are used as an delimiter to split the string. 
    #If the join of the result starts with the word "powershell" we found the script. 
    #If not continue with the next occurance of "p" and "o".
    #If all "p" and "o" are checked move on to 2 characters "po" and "we" and 3 characters ...
    lookup_length=1
    possible_script = []

    while (lookup_length < 10):
        while (start != -1):
            #Just for development: Submit a delimiter
            if args.delimiter:
                s=printable_string.split(args.delimiter)            
            else:
                
                if args.greedy:
                    start = printable_string.find(lookup[0:lookup_length],start)
                    end = printable_string.find(lookup[lookup_length:2*lookup_length],end)
                else:
                    start = printable_string.find(lookup[0:lookup_length],start)
                    end = printable_string.find(lookup[lookup_length:2*lookup_length],start+end)

                if ( start+lookup_length<end):

                    s = printable_string.split(printable_string[start+lookup_length:end])
                    
                else:
                    s= []
                    


            possible_powershell = ""
            possible_powershell = possible_powershell.join(s)



            #get the startposition of a possible powershell script and delete all previous characters
            start2=possible_powershell.lower().find("powershell")
            possible_powershell=possible_powershell[start2:]

           
            if possible_powershell[0:13].lower().startswith("powershell "):
                found=True
                possible_script.append(possible_powershell)


            if end != -1 and args.greedy:
                end = end +1

            elif start != -1:
                start=start+1
                end=0


        lookup_length = lookup_length+1  
        start=0
        end=0



    #If we found at least one possible script try to base64-decode and look for URLs
    if found:
        urls=[]
        for i in possible_script:
            #the base64 code always starts with "jab"
            s_end = i.lower().find("j")
            #first character of the base64 code
            c = i[s_end]

            #the possible script starts after "powershell -e ". now lets find the end
            #The end is the first non alphanumeric sign or an "="
            while (c.isalnum() or c == "=") and s_end>-1:
                s_end+=1
                if s_end <len(i):
                    c = i[s_end]
                else:
                    s_end = -1

            end = s_end

            #grab the base64-code
            s=i[i.lower().find("j"):end]
            

            #add "=" at the end, so we get a base64 string
            cut= len(s)%4
            if cut==3:
                s+="="            

            #Hopefully we got are real base64-encoded script
            try:
                script_bytes = base64.b64decode(s)
                script = script_bytes.decode('UTF-16LE',errors="ignore")
            except:
                continue

            if args.print_script:
                print(script)
                print ("--------------")

            

            
            #URLs are inside the parameters, so we clean up the script by e.g. removeing ",+" 
            #and grab any strings inside an pair of "(" and ")" as parameters
            arg=clean_up(get_parameters(script))
                        
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
                        else: #The last URL is not finished by the separator, so print the rest of the parameter
                            urls.append(i[s1:]) 

            #if we found no URL in the parameter, lets try something else
            #simply look for any substrings starting with "http"
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
        
        #If we found at least one URL, sort and unique the list
        if len(urls) > 0:            
            urls.sort()
            output = []
            for x in urls:
                if x not in output:
                    output.append(x)
                
            
            for i in output:
                print(i)            
        else:
            #Probably some new obfuscation in the script
            print ("No URLs found")
    else:
        #Perhaps a new way to hide the script code
        #send a link for the sample at @infsec_consult on twitter 
        print("No VBA-script found")        