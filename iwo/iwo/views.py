from django.http import HttpResponse,request,HttpRequest,FileResponse, StreamingHttpResponse
from django.shortcuts import render,redirect
from django.contrib.sessions.models import Session
import docxtpl
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxtpl import DocxTemplate
from datetime import datetime,date
import io
import zipfile
import sendfile
import tempfile
from docx import Document




HTML_STRING = """
<h1>Hello Word</h1>
"""

def home(request):

    request.session.flush()
    request.session.modified = True

    return redirect("/input2/")

def input(request):

    this_day = date.today().__str__()

    return render(request,"input2.html",{'date1':this_day})

def readword(request):
    if request.method == "POST":

        print(request.FILES['word_file'])
        doc = Document(request.FILES['word_file'])
        allText = []
        for p in doc.paragraphs:

            allText.append(p.text)

        for item in allText:
            if "dummy" in item:
                idx = allText.index(item)

        data_list = allText[idx].split("|")

        print(data_list)

        request.session["countycode"] = data_list[1]
        request.session["caseyear"] = data_list[2]
        request.session["fileno"] = data_list[3]
        request.session["suffix"] = data_list[4]
        request.session["casenumber"]=data_list[5]
        request.session["division"]=data_list[6]
        request.session["obligorname"]=data_list[7]
        request.session["obligoraddress1"]=data_list[8]
        request.session["obligoraddress2"]=data_list[9]
        request.session["obligoraddress3"]=data_list[10]
        request.session["obligeename2"]=data_list[11]
        request.session["obligeeaddress1"]=data_list[12]
        request.session["obligeeaddress2"]=data_list[13]
        request.session["obligeeaddress3"]=data_list[14]
        request.session["petitioner"]=data_list[15]
        request.session["respondent"]=data_list[16]
        request.session["date1"]=data_list[17]
        request.session["iwo"]=data_list[18]
        request.session["amendiwo"]=data_list[19]
        request.session["lump"]=data_list[20]
        request.session["termiwo"]=data_list[21]
        request.session["cse"]=data_list[22]
        request.session["court"]=data_list[23]
        request.session["attorney"]=data_list[24]
        request.session["private"]=data_list[25]
        request.session["statetribeterritory"]=data_list[26]
        request.session["remittanceid"]=data_list[27]
        request.session["citycountydisttribe"]=data_list[28]
        request.session["orderid"]=data_list[29]
        request.session["privateindividualentity"]=data_list[30]
        request.session["caseid"]=data_list[31]
        request.session["employername"]=data_list[32]
        request.session["employeraddress1"]=data_list[33]
        request.session["employeraddress2"]=data_list[34]
        request.session["employeraddress3"]=data_list[35]
        request.session["fein"]=data_list[36]
        request.session["employeename"]=data_list[37]
        request.session["employeessn"]=data_list[38]
        request.session["employeedob"]=data_list[39]
        request.session["obligeename"]=data_list[40]



        request.session["child1"]=data_list[41]
        request.session["child1dob"]=data_list[42]
        request.session["child2"]=data_list[43]
        request.session["child2dob"]=data_list[44]
        request.session["child3"]=data_list[45]
        request.session["child3dob"]=data_list[46]
        request.session["child4"]=data_list[47]
        request.session["child4dob"]=data_list[48]
        request.session["child5"]=data_list[49]
        request.session["child5dob"]=data_list[50]
        request.session["child6"]=data_list[51]
        request.session["child6dob"]=data_list[52]
        request.session["orderfromstate"]=data_list[53]
        request.session["dollar1"]=data_list[54]
        request.session["per1"]=data_list[55]
        request.session["dollar2"]=data_list[56]
        request.session["per2"]=data_list[57]
        request.session["yes12"]=data_list[58]
        request.session["no12"]=data_list[59]
        request.session["dollar3"]=data_list[60]
        request.session["per3"]=data_list[61]
        request.session["dollar4"]=data_list[62]
        request.session["per4"]=data_list[63]
        request.session["dollar5"]=data_list[64]
        request.session["per5"]=data_list[65]
        request.session["dollar6"]=data_list[66]
        request.session["per6"]=data_list[67]
        request.session["dollar7"]=data_list[68]
        request.session["per7"]=data_list[69]
        request.session["other"]=data_list[70]
        request.session["totalwithhold"]=data_list[71]
        request.session["per8"]=data_list[72]
        request.session["permonth"]=data_list[73]
        request.session["pertwoweeks"]=data_list[74]
        request.session["persemimonth"]=data_list[75]
        request.session["perweek"]=data_list[76]
        request.session["lumpsum"]=data_list[77]
        request.session["doumentid"]=data_list[78]
        request.session["principal"]=data_list[79]
        request.session["days1"]=data_list[80]
        request.session["daysof"]=data_list[81]
        request.session["days2"]=data_list[82]
        request.session["withholdpercent"]=data_list[83]
        request.session["statetribeterritory2"]=data_list[84]
        request.session["remitto"]=data_list[85]
        request.session["remitaddress"]=data_list[86]
        request.session["remitid"]=data_list[87]
        request.session["liability"]=data_list[88]
        request.session["antidiscrimination"]=data_list[89]
        request.session["supplemental"]=data_list[90]
        request.session["sender"]=data_list[91]
        request.session["sendertel"]=data_list[92]
        request.session["senderfax"]=data_list[93]
        request.session["senderwebsite"]=data_list[94]
        request.session["noticeto"]=data_list[95]
        request.session["sender2"]=data_list[96]
        request.session["sender2tel"]=data_list[97]
        request.session["senderfax2"]=data_list[98]
        request.session["senderwebsite2"]=data_list[99]
        request.session["arrears"]=data_list[100]
        request.session["arrearsrate"]=data_list[101]
        request.session["arrearsper"]=data_list[102]
        request.session["arrearspayment"]=data_list[103]
        request.session["arrearspaymentper"]=data_list[104]
        request.session["deductfull"]=data_list[105]
        request.session["deductpercent"]=data_list[106]
        request.session["deductpercentamount"]=data_list[107]
        request.session["deductnone"]=data_list[108]
        request.session["childinit1"]=data_list[109]
        request.session["childdobtwo1"]=data_list[110]
        request.session["dob181"]=data_list[111]
        request.session["allremaining1"]=data_list[112]
        request.session["childinit2"]=data_list[113]
        request.session["childdobtwo2"]=data_list[114]
        request.session["dob182"]=data_list[115]
        request.session["allremaining2"]=data_list[116]
        request.session["childinit3"]=data_list[117]
        request.session["childdobtwo3"]=data_list[118]
        request.session["dob183"]=data_list[119]
        request.session["allremaining3"]=data_list[120]
        request.session["childinit4"]=data_list[121]
        request.session["childdobtwo4"]=data_list[122]
        request.session["dob184"]=data_list[123]
        request.session["allremaining4"]=data_list[124]
        request.session["childinit5"]=data_list[125]
        request.session["childdobtwo5"]=data_list[126]
        request.session["dob185"]=data_list[127]
        request.session["allremaining5"]=data_list[128]
        request.session["childinit6"]=data_list[129]
        request.session["childdobtwo6"]=data_list[130]
        request.session["dob186"]=data_list[131]
        request.session["allremaining6"]=data_list[132]


        return redirect("/input2/")


def download(request):

    if request.method == "POST":
        message = 'message'
        
        if request.POST.get('termiwo') != "on":
            doc = DocxTemplate("DjangoIWOTemplate.docx")
        else:
            doc = DocxTemplate("DjangoIWOTemplateShort.docx")
       
        data = [
            request.POST.get('countycode'),
            request.POST.get('caseyear'),
            request.POST.get('fileno'),
            request.POST.get('suffix'),
            request.POST.get('casenumber'),
            request.POST.get('division'),
            request.POST.get('obligorname'),
            request.POST.get('obligoraddress1'),
            request.POST.get('obligoraddress2'),
            request.POST.get('obligoraddress3'),            
            request.POST.get('obligeename2'),
            request.POST.get('obligeeaddress1'),
            request.POST.get('obligeeaddress2'),
            request.POST.get('obligeeaddress3'),
            request.POST.get('petitioner'),
            request.POST.get('respondent'),
            request.POST.get('date1'),
            request.POST.get('iwo'),
            request.POST.get('amendiwo'),
            request.POST.get('lump'),
            request.POST.get('termiwo'),
            request.POST.get('cse'),
            request.POST.get('court'),
            request.POST.get('attorney'),
            request.POST.get('private'),
            request.POST.get('statetribeterritory'),
            request.POST.get('remittanceid'),
            request.POST.get('citycountydisttribe'),
            request.POST.get('orderid'),
            request.POST.get('privateindividualentity'),
            request.POST.get('caseid'),
            request.POST.get('employername'),
            request.POST.get('employeraddress1'),
            request.POST.get('employeraddress2'),
            request.POST.get('employeraddress3'),
            request.POST.get('fein'),
            request.POST.get('employeename'),
            request.POST.get('employeessn'),
            request.POST.get('employeedob'),   
            request.POST.get('obligeename'), 
            request.POST.get('child1'),
            request.POST.get('child1dob'),
            request.POST.get('child2'),
            request.POST.get('child2dob'),
            request.POST.get('child3'),
            request.POST.get('child3dob'),
            request.POST.get('child4'),
            request.POST.get('child4dob'),
            request.POST.get('child5'),
            request.POST.get('child5dob'),
            request.POST.get('child6'),
            request.POST.get('child6dob'),
            request.POST.get('orderfromstate'),
            request.POST.get('dollar1'),
            request.POST.get('per1'),
            request.POST.get('dollar2'),
            request.POST.get('per2'),
            request.POST.get('yes12'),
            request.POST.get('no12'),
            request.POST.get('dollar3'),
            request.POST.get('per3'),
            request.POST.get('dollar4'),
            request.POST.get('per4'),
            request.POST.get('dollar5'),
            request.POST.get('per5'),
            request.POST.get('dollar6'),
            request.POST.get('per6'),
            request.POST.get('dollar7'),
            request.POST.get('per7'),
            request.POST.get('other'),
            request.POST.get('totalwithhold'),
            request.POST.get('per8'),
            request.POST.get('permonth'),
            request.POST.get('pertwoweeks'),
            request.POST.get('persemimonth'),
            request.POST.get('perweek'),
            request.POST.get('lumpsum'),
            request.POST.get('documentid'),
            request.POST.get('principal'),
            request.POST.get('days1'),
            request.POST.get('daysof'),
            request.POST.get('days2'),
            request.POST.get('withholdpercent'),
            request.POST.get('statetribeterritory2'),
            request.POST.get('remitto'),
            request.POST.get('remitaddress'),
            request.POST.get('remitid'),
            request.POST.get('liability'),
            request.POST.get('antidiscrimination'),
            request.POST.get('supplemental'),
            request.POST.get('sender'),
            request.POST.get('sendertel'),
            request.POST.get('senderfax'),
            request.POST.get('senderwebsite'),
            request.POST.get('noticeto'),
            request.POST.get('sender2'),
            request.POST.get('sender2tel'),
            request.POST.get('senderfax2'),
            request.POST.get('senderwebsite2'),
            request.POST.get('arrears'),
            request.POST.get('arrearsrate'),
            request.POST.get('arrearsper'),
            request.POST.get('arrearspayment'),
            request.POST.get('arrearspaymentper'),
            request.POST.get('deductfull'),
            request.POST.get('deductpercent'),
            request.POST.get('deductpercentamount'),
            request.POST.get('deductnone'),
            request.POST.get('childinit1'),
            request.POST.get('childdobtwo1'),
            request.POST.get('dob181'),
            request.POST.get('allremaining1'),
            request.POST.get('childinit2'),
            request.POST.get('childdobtwo2'),
            request.POST.get('dob182'),
            request.POST.get('allremaining2'),
            request.POST.get('childinit3'),
            request.POST.get('childdobtwo3'),
            request.POST.get('dob183'),
            request.POST.get('allremaining3'),
            request.POST.get('childinit4'),
            request.POST.get('childdobtwo4'),
            request.POST.get('dob184'),
            request.POST.get('allremaining4'),
            request.POST.get('childinit5'),
            request.POST.get('childdobtwo5'),
            request.POST.get('dob185'),
            request.POST.get('allremaining5'),
            request.POST.get('childinit6'),
            request.POST.get('childdobtwo6'),
            request.POST.get('dob186'),
            request.POST.get('allremaining6'),
        ]

        data_string = "dummy"

        for d in data:
            data_string = data_string + "|" + str(d)

        temp_obligorname = request.POST.get('obligorname').split(',')
        obligornamebig = temp_obligorname[1].upper() + " " + temp_obligorname[0].upper()
        obligornamesmall = temp_obligorname[1] + " " + temp_obligorname[0]

        temp_obligeename = request.POST.get('obligeename2').split(',')
        obligeenamebig = temp_obligeename[1].upper() + " " + temp_obligeename[0].upper()
        obligeenamesmall = temp_obligeename[1] + " " + temp_obligeename[0]


        if request.POST.get('obligoraddress1') != "" and request.POST.get('obligoraddress2') == "" and request.POST.get('obligoraddress3') == "":
            obligoraddress = request.POST.get('obligoraddress1')
        elif request.POST.get('obligoraddress1') != "" and request.POST.get('obligoraddress2') != "" and request.POST.get('obligoraddress3') == "":
            obligoraddress = request.POST.get('obligoraddress1') + ", " + request.POST.get('obligoraddress2')
        else:
            obligoraddress = request.POST.get('obligoraddress1') + ", " + request.POST.get('obligoraddress2') + ", " + request.POST.get('obligoraddress3')

        if request.POST.get('obligeeaddress1') != "" and request.POST.get('obligeeaddress2') == "" and request.POST.get('obligeeaddress3') == "":
            obligeeaddress = request.POST.get('obligeeaddress1')
        elif request.POST.get('obligeeaddress1') != "" and request.POST.get('obligeeaddress2') != "" and request.POST.get('obligeeaddress3') == "":
            obligeeaddress = request.POST.get('obligeeaddress1') + ", " + request.POST.get('obligeeaddress2')
        else:
            obligeeaddress = request.POST.get('obligeeaddress1') + ", " + request.POST.get('obligeeaddress2') + ", " + request.POST.get('obligeeaddress3')

        if len(request.POST.get('employername')) < 23:
            employername_short = request.POST.get('employername') + " "*(23-len(request.POST.get('employername')))
        else:
            employername_short = "".join(list(request.POST.get('employername'))[:18]) + ". . ."
        
        county_parts_list = request.POST.get('countycode').split(',')
        circuit = county_parts_list[2].upper()
        county = county_parts_list[1].upper()

        if request.POST.get('petitioner') == 'on':
            petitioner_big = obligornamebig
            respondent_big = obligeenamebig
        else:
            petitioner_big = obligeenamebig
            respondent_big = obligornamebig

        if request.POST.get('arrears') == "":
            arrears = "NA"
        else:
            arrears = request.POST.get('arrears')

        if request.POST.get('arrearsrate') == "":
            arrearsrate = "NA"
        else:
            arrearsrate = request.POST.get('arrearsrate')

        if request.POST.get('arrearsper') == "":
            arrearsper = "NA"
        else:
            arrearsper = request.POST.get('arrearsper')


        liability_list = request.POST.get('liability').split()
        if len(liability_list) > 0:
            i = 1
            temp_liability_list = [liability_list[0]]
            while i != 0 and i <= len(liability_list)-1:
                temp_liability_list.append(liability_list[i])
                if len(" ".join(temp_liability_list)) <=93:
                    liability1 = " ".join(temp_liability_list)
                    temp_liability_list[0] = liability1
                    temp_liability_list.pop(1)
                    print(liability1)
                    i += 1
                else:
                    i = 0
        else: 
            
            liability1 = ""

        if len(liability1) > 0 and len(liability_list) > len(liability1.split()):
            i = 1
            liability_list = liability_list[len(liability1.split()):]
            temp_liability_list = [liability_list[0]]
            while i != 0 and i <= len(liability_list) - 1:
                temp_liability_list.append(liability_list[i])
                if len(" ".join(temp_liability_list)) <=93:
                    liability2 = " ".join(temp_liability_list)
                    temp_liability_list[0] = liability2
                    temp_liability_list.pop(1)
                    print(liability2)
                    i += 1
                else:
                    i = 0
        else: 

            liability2 = "" 

        if len(liability2) and len(liability_list) > len(liability2.split()):
            i = 1
            liability_list = liability_list[len(liability2.split()):]
            temp_liability_list = [liability_list[0]]
            while i != 0 and i <= len(liability_list) -1:
                temp_liability_list.append(liability_list[i])
                if len(" ".join(temp_liability_list)) <=93:
                    liability3 = " ".join(temp_liability_list)
                    temp_liability_list[0] = liability3
                    temp_liability_list.pop(1)
                    print(liability3)
                    i += 1
                else:
                    i = 0
        else: 

            liability3 = ""    
        
        if len(liability3) and len(liability_list) > len(liability3.split()):
            i = 1
            liability_list = liability_list[len(liability2.split()):]
            temp_liability_list = [liability_list[0]]
            while i != 0 and i <= len(liability_list) -1:
                temp_liability_list.append(liability_list[i])
                if len(" ".join(temp_liability_list)) <=93:
                    liability4 = " ".join(temp_liability_list)
                    temp_liability_list[0] = liability4
                    temp_liability_list.pop(1)
                    print(liability4)
                    i += 1
                else:
                    i = 0
        else: 

            liability4 = "" 
            
               

        liability = liability1 + "_"*(93-len(liability1)) + "\n" + liability2 + "_"*(93-len(liability2)) + "\n" + liability3 + "_"*(93-len(liability3)) + "\n" + liability4 + "_"*(93-len(liability4))


        antidisc_list = request.POST.get('antidiscrimination').split()
        if len(antidisc_list) > 0:
            i = 1
            temp_antidisc_list = [antidisc_list[0]]
            while i != 0 and i <= len(antidisc_list)-1:
                temp_antidisc_list.append(antidisc_list[i])
                if len(" ".join(temp_antidisc_list)) <=93:
                    antidisc1 = " ".join(temp_antidisc_list)
                    temp_antidisc_list[0] = antidisc1
                    temp_antidisc_list.pop(1)
                    print(antidisc1)
                    i += 1
                else:
                    i = 0
        else: 
            
            antidisc1 = ""

        if len(antidisc1) > 0 and len(antidisc_list) > len(antidisc1.split()):
            i = 1
            antidisc_list = antidisc_list[len(antidisc1.split()):]
            temp_antidisc_list = [antidisc_list[0]]
            while i != 0 and i <= len(antidisc_list) - 1:
                temp_liability_list.append(antidisc_list[i])
                if len(" ".join(temp_antidisc_list)) <=93:
                    antidisc2 = " ".join(temp_antidisc_list)
                    temp_antidisc_list[0] = antidisc2
                    temp_antidisc_list.pop(1)
                    print(antidisc2)
                    i += 1
                else:
                    i = 0
        else: 

            antidisc2 = "" 

        if len(antidisc2) and len(antidisc_list) > len(antidisc2.split()):
            i = 1
            antidisc_list = antidisc_list[len(antidisc2.split()):]
            temp_antidisc_list = [antidisc_list[0]]
            while i != 0 and i <= len(antidisc_list) -1:
                temp_antidisc_list.append(antidisc_list[i])
                if len(" ".join(temp_antidisc_list)) <=93:
                    antidisc3 = " ".join(temp_antidisc_list)
                    temp_antidisc_list[0] = antidisc3
                    temp_antidisc_list.pop(1)
                    print(antidisc3)
                    i += 1
                else:
                    i = 0
        else: 

            antidisc3 = ""    
        
        if len(antidisc3) and len(antidisc_list) > len(antidisc3.split()):
            i = 1
            antidisc_list = antidisc_list[len(antidisc3.split()):]
            temp_antidisc_list = [antidisc_list[0]]
            while i != 0 and i <= len(antidisc_list) -1:
                temp_antidisc_list.append(antidisc_list[i])
                if len(" ".join(temp_antidisc_list)) <=93:
                    antidisc4 = " ".join(temp_antidisc_list)
                    temp_antidisc_list[0] = antidisc4
                    temp_antidisc_list.pop(1)
                    print(antidisc4)
                    i += 1
                else:
                    i = 0
        else: 

            antidisc4 = "" 
        
        antidisc = antidisc1 + "_"*(93-len(antidisc1)) + "\n" + antidisc2 + "_"*(93-len(antidisc2)) + "\n" + antidisc3 + "_"*(93-len(antidisc3)) + "\n" + antidisc4 + "_"*(93-len(antidisc4))



        supp_list = request.POST.get('supplemental').split()
        if len(supp_list) > 0:
            i = 1
            temp_supp_list = [supp_list[0]]
            while i != 0 and i <= len(supp_list)-1:
                temp_supp_list.append(supp_list[i])
                if len(" ".join(temp_supp_list)) <=93:
                    supp1 = " ".join(temp_supp_list)
                    temp_supp_list[0] = supp1
                    temp_supp_list.pop(1)
                    print(supp1)
                    i += 1
                else:
                    i = 0
        else: 
            
            supp1 = ""

        if len(supp1) > 0 and len(supp_list) > len(supp1.split()):
            i = supp_list
            supp_list = supp_list[len(supp1.split()):]
            temp_supp_list = [supp_list[0]]
            while i != 0 and i <= len(supp_list) - 1:
                temp_supp_list.append(supp_list[i])
                if len(" ".join(temp_supp_list)) <=93:
                    supp2 = " ".join(temp_supp_list)
                    temp_supp_list[0] = supp2
                    temp_supp_list.pop(1)
                    print(supp2)
                    i += 1
                else:
                    i = 0
        else: 

            supp2 = "" 
        if len(supp2) > 0 and len(supp_list) > len(supp2.split()):
            i = supp_list
            supp_list = supp_list[len(supp2.split()):]
            temp_supp_list = [supp_list[0]]
            while i != 0 and i <= len(supp_list) - 1:
                temp_supp_list.append(supp_list[i])
                if len(" ".join(temp_supp_list)) <=93:
                    supp3 = " ".join(temp_supp_list)
                    temp_supp_list[0] = supp3
                    temp_supp_list.pop(1)
                    print(supp3)
                    i += 1
                else:
                    i = 0
        else: 

            supp3 = "" 

        if len(supp3) > 0 and len(supp_list) > len(supp3.split()):
            i = supp_list
            supp_list = supp_list[len(supp3.split()):]
            temp_supp_list = [supp_list[0]]
            while i != 0 and i <= len(supp_list) - 1:
                temp_supp_list.append(supp_list[i])
                if len(" ".join(temp_supp_list)) <=93:
                    supp4 = " ".join(temp_supp_list)
                    temp_supp_list[0] = supp4
                    temp_supp_list.pop(1)
                    print(supp4)
                    i += 1
                else:
                    i = 0
        else: 

            supp4 = "" 

        if len(supp4) > 0 and len(supp_list) > len(supp4.split()):
            i = supp_list
            supp_list = supp_list[len(supp4.split()):]
            temp_supp_list = [supp_list[0]]
            while i != 0 and i <= len(supp_list) - 1:
                temp_supp_list.append(supp_list[i])
                if len(" ".join(temp_supp_list)) <=93:
                    supp5 = " ".join(temp_supp_list)
                    temp_supp_list[0] = supp5
                    temp_supp_list.pop(1)
                    print(supp5)
                    i += 1
                else:
                    i = 0
        else: 

            supp5 = "" 

        if len(supp5) > 0 and len(supp_list) > len(supp5.split()):
            i = supp_list
            supp_list = supp_list[len(supp5.split()):]
            temp_supp_list = [supp_list[0]]
            while i != 0 and i <= len(supp_list) - 1:
                temp_supp_list.append(supp_list[i])
                if len(" ".join(temp_supp_list)) <=93:
                    supp6 = " ".join(temp_supp_list)
                    temp_supp_list[0] = supp6
                    temp_supp_list.pop(1)
                    print(supp6)
                    i += 1
                else:
                    i = 0
        else: 

            supp6 = "" 

        supp = supp1 + "_"*(93-len(supp1)) + "\n" + supp2 + "_"*(93-len(supp2)) + "\n" + supp3 + "_"*(93-len(supp3)) + "\n" + supp4 + "_"*(93-len(supp4)) + "\n" + supp5 + "_"*(93-len(supp5)) + "\n" + supp6 + "_"*(93-len(supp6))



        context = {
            "countycode":request.POST.get('countycode'),
            "caseyear":request.POST.get('caseyear'),
            "fileno":request.POST.get('fileno'),
            "suffix":request.POST.get('suffix'),
            "casenumber":request.POST.get('casenumber'),
            "division":request.POST.get('division'),
            "obligorname":request.POST.get('obligorname'),
            "obligoraddress":obligoraddress,
            "obligeename2":request.POST.get('obligeename2'),
            "obligeeaddress":obligeeaddress,
            "petitioner":request.POST.get('petitioner'),
            "respondent":request.POST.get('respondent'),
            "date1": datetime.strftime(datetime.strptime(request.POST.get('date1'),"%Y-%m-%d"),"%m/%d/%Y") + "_"*(20-len(datetime.strftime(datetime.strptime(request.POST.get('date1'),"%Y-%m-%d"),"%m/%d/%Y"))),
            "iwo": request.POST.get('iwo'),
            "amendiwo": request.POST.get('amendiwo'),
            "lump": request.POST.get('lump'),
            "termiwo": request.POST.get('termiwo'),
            "cse": request.POST.get('cse'),
            "court": request.POST.get('court'),
            "attorney": request.POST.get('attorney'),
            "private": request.POST.get('private'),
            "statetribeterritory": request.POST.get('statetribeterritory') + "_"*(24-len(request.POST.get('statetribeterritory'))),
            "statetribeterritory2": request.POST.get('statetribeterritory') + "_"*(17-len(request.POST.get('statetribeterritory'))),
            "statetribeterritory3": request.POST.get('statetribeterritory') + "_"*(17-len(request.POST.get('statetribeterritory'))),
            "statetribeterritory4": request.POST.get('statetribeterritory') + "_"*(17-len(request.POST.get('statetribeterritory'))),
            "statetribeterritory10": request.POST.get('statetribeterritory') + "_"*(43-len(request.POST.get('statetribeterritory'))),
            "remittanceid": request.POST.get('remittanceid') + "_"*(21-len(request.POST.get('remittanceid'))),
            "citycountydisttribe": request.POST.get('citycountydisttribe') + "_"*(22-len(request.POST.get('citycountydisttribe'))),
            "orderid": request.POST.get('orderid') + "_"*(40-len(request.POST.get('orderid'))),
            "orderid2": request.POST.get('orderid') + "_"*(44-len(request.POST.get('orderid'))),
            "orderid3": request.POST.get('orderid') + "_"*(44-len(request.POST.get('orderid'))),
            "orderid4": request.POST.get('orderid') + "_"*(44-len(request.POST.get('orderid'))),
            "privateindividualentity": request.POST.get('privateindividualentity') + "_"*(22-len(request.POST.get('privateindividualentity'))),
            "caseid": request.POST.get('caseid') + "_"*(40-len(request.POST.get('caseid'))),
            "caseid2": request.POST.get('caseid') + "_"*(33-len(request.POST.get('caseid'))),
            "caseid3": request.POST.get('caseid') + "_"*(33-len(request.POST.get('caseid'))),
            "caseid4": request.POST.get('caseid') + "_"*(33-len(request.POST.get('caseid'))),
            "employername": request.POST.get('employername') + "_"*(40-len(request.POST.get('employername'))),
            "employername2":employername_short,
            "employername3":employername_short,
            "employername4":employername_short,
            "employeraddress1": request.POST.get('employeraddress1') + "_"*(40-len(request.POST.get('employeraddress1'))),
            "employeraddress2": request.POST.get('employeraddress2') + "_"*(40-len(request.POST.get('employeraddress2'))),
            "employeraddress3": request.POST.get('employeraddress3') + "_"*(40-len(request.POST.get('employeraddress3'))),
            "fein":request.POST.get('fein') + "_"*(19-len(request.POST.get('fein'))),
            "fein2":request.POST.get('fein') + "_"*(17-len(request.POST.get('fein'))),
            "fein3":request.POST.get('fein') + "_"*(17-len(request.POST.get('fein'))),
            "fein4":request.POST.get('fein') + "_"*(17-len(request.POST.get('fein'))),
            "employeename": request.POST.get('employeename') + "_"*(42-len(request.POST.get('employeename'))),
            "employeename2": request.POST.get('employeename') + "_"*(47-len(request.POST.get('employeename'))),
            "employeename2": request.POST.get('employeename') + "_"*(47-len(request.POST.get('employeename'))),
            "employeename2": request.POST.get('employeename') + "_"*(47-len(request.POST.get('employeename'))),
            "employeessn": request.POST.get('employeessn') + "_"*(42-len(request.POST.get('employeessn'))),
            "employeessn2": request.POST.get('employeessn') + "_"*(22-len(request.POST.get('employeessn'))),
            "employeessn3": request.POST.get('employeessn') + "_"*(22-len(request.POST.get('employeessn'))),
            "employeessn4": request.POST.get('employeessn') + "_"*(22-len(request.POST.get('employeessn'))),
            "employeedob": request.POST.get('employeedob') + "_"*(42-len(request.POST.get('employeedob'))),
            "obligeename": request.POST.get('obligeename') + "_"*(42-len(request.POST.get('obligeename'))),
            "child1":request.POST.get('child1') + "_"*(28-len(request.POST.get('child1'))),
            "child1dob":request.POST.get('child1dob') + "_"*(17-len(request.POST.get('child1dob'))),
            "child2":request.POST.get('child2') + "_"*(28-len(request.POST.get('child2'))),
            "child2dob":request.POST.get('child2dob') + "_"*(17-len(request.POST.get('child2dob'))),
            "child3":request.POST.get('child3') + "_"*(28-len(request.POST.get('child3'))),
            "child3dob":request.POST.get('child3dob') + "_"*(17-len(request.POST.get('child3dob'))),
            "child4":request.POST.get('child4') + "_"*(28-len(request.POST.get('child4'))),
            "child4dob":request.POST.get('child4dob') + "_"*(17-len(request.POST.get('child4dob'))),
            "child5":request.POST.get('child5') + "_"*(28-len(request.POST.get('child5'))),
            "child5dob":request.POST.get('child5dob') + "_"*(17-len(request.POST.get('child5dob'))),
            "child6":request.POST.get('child6') + "_"*(28-len(request.POST.get('child6'))),
            "child6dob":request.POST.get('child6dob') + "_"*(17-len(request.POST.get('child6dob'))),
            "orderfromstate":request.POST.get('orderfromstate') + "_"*(43-len(request.POST.get('orderfromstate'))),
            "dollar1":request.POST.get('dollar1') + "_"*(12-len(request.POST.get('dollar1'))),
            "per1":request.POST.get('per1') + "_"*(13-len(request.POST.get('per1'))),
            "dollar2":request.POST.get('dollar2') + "_"*(12-len(request.POST.get('dollar2'))),
            "per2":request.POST.get('per2') + "_"*(13-len(request.POST.get('per2'))),
            "yes12":request.POST.get('yes12'),
            "no12":request.POST.get('no12'),
            "dollar3":request.POST.get('dollar3') + "_"*(12-len(request.POST.get('dollar3'))),
            "per3":request.POST.get('per3') + "_"*(13-len(request.POST.get('per3'))),
            "dollar4":request.POST.get('dollar4') + "_"*(12-len(request.POST.get('dollar4'))),
            "per4":request.POST.get('per4') + "_"*(13-len(request.POST.get('per4'))),
            "dollar5":request.POST.get('dollar5') + "_"*(12-len(request.POST.get('dollar5'))),
            "per5":request.POST.get('per5') + "_"*(13-len(request.POST.get('per5'))),
            "dollar6":request.POST.get('dollar6') + "_"*(12-len(request.POST.get('dollar6'))),
            "per6":request.POST.get('per6') + "_"*(13-len(request.POST.get('per6'))),
            "dollar7":request.POST.get('dollar7') + "_"*(12-len(request.POST.get('dollar7'))),
            "per7":request.POST.get('per7') + "_"*(13-len(request.POST.get('per7'))),
            "other":request.POST.get('other') + "_"*(40-len(request.POST.get('other'))),
            "totalwithhold":request.POST.get('totalwithhold') + "_"*(10-len(request.POST.get('totalwithhold'))),
            "per8":request.POST.get('per8') + "_"*(13-len(request.POST.get('per8'))),
            "permonth": request.POST.get('permonth') + "_"*(11-len(request.POST.get('permonth'))),
            "pertwoweeks": request.POST.get('pertwoweeks') + "_"*(10-len(request.POST.get('pertwoweeks'))),
            "persemimonth": request.POST.get('persemimonth') + "_"*(11-len(request.POST.get('persemimonth'))),
            "perweek": request.POST.get('perweek') + "_"*(10-len(request.POST.get('perweek'))),
            "lumpsum":request.POST.get('lumpsum') + "_"*(10-len(request.POST.get('lumpsum'))),
            "documentid":request.POST.get('documentid') + "_"*(27-len(request.POST.get('documentid'))),
            "principal":request.POST.get('principal') + "_"*(17-len(request.POST.get('per1'))),
            "days1":request.POST.get('days1') + "_"*(5-len(request.POST.get('days1'))),
            "daysof":request.POST.get('daysof') + "_"*(11-len(request.POST.get('daysof'))),
            "days2":request.POST.get('days2') + "_"*(5-len(request.POST.get('days2'))),
            "withholdpercent":request.POST.get('withholdpercent') + "_"*(4-len(request.POST.get('withholdpercent'))),
            "statetribeterritory2":request.POST.get('statetribeterritory2') + "_"*(17-len(request.POST.get('statetribeterritory2'))),
            "remitto":request.POST.get('remitto') + "_"*(56-len(request.POST.get('remitto'))),
            "remitaddress":request.POST.get('remitaddress') + "_"*(66-len(request.POST.get('remitaddress'))),
            "remitid":request.POST.get('remitid') + "_"*(11-len(request.POST.get('remitid'))),
            "liability":liability,
            "antidiscrimination":antidisc,
            "supplemental":request.POST.get('supplemental') + "_"*(558-len(request.POST.get('supplemental'))),
            "sender":request.POST.get('sender') + "_"*(28-len(request.POST.get('sender'))),
            "sendertel":request.POST.get('sendertel') + "_"*(14-len(request.POST.get('sendertel'))),
            "senderfax":request.POST.get('senderfax') + "_"*(13-len(request.POST.get('senderfax'))),
            "senderwebsite":request.POST.get('senderwebsite') + "_"*(29-len(request.POST.get('senderwebsite'))),
            "noticeto":request.POST.get('noticeto') + "_"*(89-len(request.POST.get('noticeto'))),
            "sender2":request.POST.get('sender2') + "_"*(24-len(request.POST.get('sender2'))),
            "sender2tel":request.POST.get('sender2tel') + "_"*(14-len(request.POST.get('sender2tel'))),
            "senderfax2":request.POST.get('senderfax2') + "_"*(13-len(request.POST.get('senderfax2'))),
            "senderwebsite2":request.POST.get('senderwebsite2') + "_"*(29-len(request.POST.get('senderwebsite2'))),
            "circuit":circuit,
            "county":county,
            "petitionerbig":petitioner_big,
            "respondentbig":respondent_big,
            "employerupper":request.POST.get('employername').upper(),
            "obligornamesmall":obligornamesmall,
            "obligeenamesmall":obligeenamesmall,
            "arrears":arrears,
            "arrearsrate":arrearsrate,
            "arrearsper":arrearsper,
            "arrearspayment":request.POST.get('arrearspayment'),
            "arrearspaymentper":request.POST.get('arrearspaymentper'),
            "deductfull":request.POST.get('deductfull'),
            "deductpercent":request.POST.get('deductpercent'),
            "deductpercentamount":request.POST.get('deductpercentamount'),
            "deductnone":request.POST.get('deductnone'),
            "childinit1":request.POST.get('childinit1'),
            "childdobtwo1":request.POST.get('childdobtwo1'),
            "dob181":request.POST.get('dob181'),
            "allremaining1":request.POST.get('allremaining1'),
            "childinit2":request.POST.get('childinit2'),
            "childdobtwo2":request.POST.get('childdobtwo2'),
            "dob182":request.POST.get('dob182'),
            "allremaining2":request.POST.get('allremaining2'),
            "childinit3":request.POST.get('childinit3'),
            "childdobtwo3":request.POST.get('childdobtwo3'),
            "dob183":request.POST.get('dob183'),
            "allremaining3":request.POST.get('allremaining3'),
            "childinit4":request.POST.get('childinit4'),
            "childdobtwo4":request.POST.get('childdobtwo4'),
            "dob184":request.POST.get('dob184'),
            "allremaining4":request.POST.get('allremaining4'),
            "childinit5":request.POST.get('childinit5'),
            "childdobtwo5":request.POST.get('childdobtwo5'),
            "dob185":request.POST.get('dob185'),
            "allremaining5":request.POST.get('allremaining5'),
            "childinit6":request.POST.get('childinit6'),
            "childdobtwo6":request.POST.get('childdobtwo6'),
            "dob186":request.POST.get('dob186'),
            "allremaining6":request.POST.get('allremaining6'),
            "data_string":data_string
        }



        doc.render(context)
        doc2 = doc.render(context)
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        response = StreamingHttpResponse(streaming_content=buffer,content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        response["Content-Disposition"] = 'attachment;filename="test.docx"'
        response['Content-Encoding'] = 'UTF-8'
  
    return response

