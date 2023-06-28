from django.http import StreamingHttpResponse
from django.shortcuts import render,redirect
from django.contrib.sessions.models import Session
from docxtpl import DocxTemplate
from datetime import datetime,date
import io
from docx import Document
import os

def format_dollars(amount):
    amount = amount.replace("$","").replace(",","")
    if len(amount) > 9:
        amount = amount[0:len(amount)-9] + "," + amount[len(amount)-9:len(amount)-6] + "," + amount[-6:]
    elif len(amount) > 6:
        amount = amount[0:len(amount)-6] + "," + amount[-6:]
    else:
        amount = amount
    return amount

def home(request):

    request.session.flush()
    request.session.modified = True

    return redirect("/bar-services/iwo/input/")

def input(request):

    this_day = date.today().__str__()
    return render(request,"input.html",{'date1':this_day})

def help(request):
    return render(request,'help.html')

# Function handling import of data from previously generated IWO/Addendum .docx file.  Data is stored in a comma-separated string in each .docx template which remains hidden
# under white font color.  If new fields are added to the IWO/Addendum templates, the new data must be referenced by the same index as that in the "data" list contained
# in the "download" function below.

def readword(request):
    if request.method == "POST":

        try:
            doc = Document(request.FILES['word_file'])
            allText = []
            for p in doc.paragraphs:

                allText.append(p.text)
            
            for item in allText:
                if "dummy" in item:
                    idx = allText.index(item)

            data_list = allText[idx].split("|")

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
            request.session["global_format_off"]=data_list[133]
            request.session["obligorname_noformat1"]=data_list[134]
            request.session["obligeename_noformat1"]=data_list[135]
            request.session["child1_noformat1"]=data_list[136]
            request.session["child2_noformat1"]=data_list[137]
            request.session["child3_noformat1"]=data_list[138]
            request.session["child4_noformat1"]=data_list[139]
            request.session["child5_noformat1"]=data_list[140]
            request.session["child6_noformat1"]=data_list[141]
            request.session["employername2"]=data_list[142]
            request.session["fein2"]=data_list[143]
            request.session["employeename2"]=data_list[144]
            request.session["employeessn2"]=data_list[145]
            request.session["caseid2"]=data_list[146]
            request.session["orderid2"]=data_list[147]

            return redirect("/bar-services/iwo/input/")
    
        except:

            return redirect("/bar-services/iwo/import-error/")
        
def import_error(request):
    return render(request,'importerr.html')

# Function handling population of data from form-fields to .docx template and download to user's machine.

def download(request):

    if request.method == "POST":

        if request.POST.get('termiwo') != "on":
            doc = DocxTemplate("DjangoIWOTemplate.docx")
        else:
            doc = DocxTemplate("DjangoIWOTemplateShort.docx")
       
        # Data string written to .docx template in comma-separated form under white font color.  Index of any new values added must correspond with indexes 
        # referenced in the "readword" function above.

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
            request.POST.get('global_format_off'),
            request.POST.get('obligorname_noformat1'),
            request.POST.get('obligeename_noformat1'),
            request.POST.get('child1_noformat1'),
            request.POST.get('child2_noformat1'),
            request.POST.get('child3_noformat1'),
            request.POST.get('child4_noformat1'),
            request.POST.get('child5_noformat1'),
            request.POST.get('child6_noformat1'),    
            request.POST.get('employername2'),
            request.POST.get('fein'),
            request.POST.get('employeename2'),
            request.POST.get('employeessn2'),
            request.POST.get('caseid2'),
            request.POST.get('orderid2'),
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
            employername_short = request.POST.get('employername') + "_"*(23-len(request.POST.get('employername')))
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
            arrears = "$" + format_dollars(request.POST.get('arrears'))

        if request.POST.get('arrearsrate') == "":
            arrearsrate = "NA"
        else:
            arrearsrate = "$" + format_dollars(request.POST.get('arrearsrate'))

        if request.POST.get('arrearsper') == "":
            arrearsper = "NA"
        else:
            arrearsper = request.POST.get('arrearsper')


        liability = request.POST.get('liability')

        if len(liability.split()) == 1:
            liability = liability + "_"*(93-len(liability)) + "\n" + "_"*93 + "\n" + "_"*93 + "\n" + "_"*93
        else:
            liability_list = liability.split()

            if len(liability_list) > 0:
                i = 1
                temp_liability_list = [liability_list[0]]
                while  i != 0 and i <= len(liability_list) - 1:
                    temp_liability_list.append(liability_list[i])
                    if len("_".join(temp_liability_list)) <= 93:
                        liability1 = "_".join(temp_liability_list)
                        temp_liability_list[0] = liability1
                        temp_liability_list.pop(1)
                        add_back = liability_list[i]
                        i += 1
                    elif len(liability_list) == 1:
                        liability1 = liability_list[0]
                    else:
                        i = 0
                for item in liability1.split("_"):
                    if item in liability_list:
                        liability_list.remove(item)
                if len(liability_list) > 0:
                    liability_list.insert(0,add_back)
            else:
                liability1 = ""
            
            if len(liability_list) > 0:
                i = 1
                liability_list = liability_list[len(liability1.split()):]
                temp_liability_list = [liability_list[0]]
                if len(liability_list) - 1 > 0:
                    while i != 0 and i <= len(liability_list) - 1:
                        temp_liability_list.append(liability_list[i])
                        if len("_".join(temp_liability_list)) <= 93:
                            liability2 = "_".join(temp_liability_list)
                            temp_liability_list[0] = liability2
                            temp_liability_list.pop(1)
                            add_back = liability_list[i]
                            i += 1
                        else:
                            i = 0
                elif len(liability_list) == 1:
                    liability2 = liability_list[0]
                else:
                    liability2 = ""
                
                for item in liability2.split("_"):
                    if item in liability_list:
                        liability_list.remove(item)
                if len(liability_list) > 0:
                    liability_list.insert(0,add_back)
            else:
                liability2 = ""
            
            if len(liability_list) > 0:
                i = 1
                liability_list = liability_list[len(liability2.split()):]
                temp_liability_list = [liability_list[0]]
                if len(liability_list) - 1 > 0:
                    while i != 0 and i <= len(liability_list) - 1:
                        temp_liability_list.append(liability_list[i])
                        if len("_".join(temp_liability_list)) <= 93:
                            liability3 = "_".join(temp_liability_list)
                            temp_liability_list[0] = liability3
                            temp_liability_list.pop(1)
                            add_back = liability_list[i]
                            i += 1
                        else:
                            i = 0
                elif len(liability_list) == 1:
                    liability3 = liability_list[0]
                else:
                    liability3 = ""
                
                for item in liability3.split("_"):
                    if item in liability_list:
                        liability_list.remove(item)
                if len(liability_list) > 0:
                    liability_list.insert(0,add_back)
            else:
                liability3 = ""

            if len(liability_list) > 0:
                i = 1
                liability_list = liability_list[len(liability3.split()):]
                temp_liability_list = [liability_list[0]]
                if len(liability_list) - 1 > 0:
                    while i != 0 and i <= len(liability_list) -1:
                        temp_liability_list.append(liability_list[i])
                        if len("_".join(temp_liability_list)) <=93:
                            liability4 = "_".join(temp_liability_list)
                            temp_liability_list[0] = liability4
                            temp_liability_list.pop(1)
                            add_back = liability_list[i]
                            i += 1
                        else:
                            i = 0
                elif len(liability_list) == 1:
                    liability4 = liability_list[0]
                else:
                    liability4 = ""
            else:
                liability4 = ""

            liability = liability1 + "_"*(93-len(liability1)) + "\n" + liability2 + "_"*(93-len(liability2)) + "\n" + liability3 + "_"*(93-len(liability3)) + "\n" + liability4 + "_"*(93-len(liability4))


        antidisc = request.POST.get('antidiscrimination')

        if len(antidisc.split()) == 1:
            antidisc = antidisc + "_"*(93-len(antidisc)) + "\n" + "_"*93 + "\n" + "_"*93 + "\n" + "_"*93
        else:
            antidisc_list = antidisc.split()

            if len(antidisc_list) > 0:
                i = 1
                temp_antidisc_list = [antidisc_list[0]]
                while  i != 0 and i <= len(antidisc_list) - 1:
                    temp_antidisc_list.append(antidisc_list[i])
                    if len("_".join(temp_antidisc_list)) <= 93:
                        antidisc1 = "_".join(temp_antidisc_list)
                        temp_antidisc_list[0] = antidisc1
                        temp_antidisc_list.pop(1)
                        add_back = antidisc_list[i]
                        i += 1
                    elif len(antidisc_list) == 1:
                        antidisc1 = antidisc_list[0]
                    else:
                        i = 0
                for item in antidisc1.split("_"):
                    if item in antidisc_list:
                        antidisc_list.remove(item)
                if len(antidisc_list) > 0:
                    antidisc_list.insert(0,add_back)
            else:
                antidisc1 = "" 

            if len(antidisc_list) > 0:
                i = 1
                antidisc_list = antidisc_list[len(antidisc1.split()):]
                temp_antidisc_list = [antidisc_list[0]]
                if len(antidisc_list) - 1 > 0:
                    while i != 0 and i <= len(antidisc_list) - 1:
                        temp_antidisc_list.append(antidisc_list[i])
                        if len("_".join(temp_antidisc_list)) <= 93:
                            antidisc2 = "_".join(temp_antidisc_list)
                            temp_antidisc_list[0] = antidisc2
                            temp_antidisc_list.pop(1)
                            add_back = antidisc_list[i]
                            i += 1
                        else:
                            i = 0
                elif len(antidisc_list) == 1:
                    antidisc2 = antidisc_list[0]
                else:
                    antidisc2 = ""
                
                for item in antidisc2.split("_"):
                    if item in antidisc_list:
                        antidisc_list.remove(item)
                if len(antidisc_list) > 0:
                    antidisc_list.insert(0,add_back)
            else:
                antidisc2 = ""
            
            if len(antidisc_list) > 0:
                i = 1
                antidisc_list = antidisc_list[len(antidisc2.split()):]
                temp_antidisc_list = [antidisc_list[0]]
                if len(antidisc_list) - 1 > 0:
                    while i != 0 and i <= len(antidisc_list) - 1:
                        temp_antidisc_list.append(antidisc_list[i])
                        if len("_".join(temp_antidisc_list)) <= 93:
                            antidisc3 = "_".join(temp_antidisc_list)
                            temp_antidisc_list[0] = antidisc3
                            temp_antidisc_list.pop(1)
                            add_back = antidisc_list[i]
                            i += 1
                        else:
                            i = 0
                elif len(antidisc_list) == 1:
                    antidisc3 = antidisc_list[0]
                else:
                    antidisc3 = ""
                
                for item in antidisc3.split("_"):
                    if item in antidisc_list:
                        antidisc_list.remove(item)
                if len(antidisc_list) > 0:
                    antidisc_list.insert(0,add_back)
            else:
                antidisc3 = ""

            if len(antidisc_list) > 0:
                i = 1
                antidisc_list = antidisc_list[len(antidisc3.split()):]
                temp_antidisc_list = [antidisc_list[0]]
                if len(antidisc_list) - 1 > 0:
                    while i != 0 and i <= len(antidisc_list) -1:
                        temp_antidisc_list.append(antidisc_list[i])
                        if len("_".join(temp_antidisc_list)) <=93:
                            antidisc4 = "_".join(temp_antidisc_list)
                            temp_antidisc_list[0] = antidisc4
                            temp_antidisc_list.pop(1)
                            add_back = antidisc_list[i]
                            i += 1
                        else:
                            i = 0
                elif len(antidisc_list) == 1:
                    antidisc4 = antidisc_list[0]
                else:
                    antidisc4 = ""
            else:
                antidisc4 = ""

            antidisc = antidisc1 + "_"*(93-len(antidisc1)) + "\n" + antidisc2 + "_"*(93-len(antidisc2)) + "\n" + antidisc3 + "_"*(93-len(antidisc3)) + "\n" + antidisc4 + "_"*(93-len(antidisc4))


        supp = request.POST.get('supplemental')

        if len(supp.split()) == 1:
            supp = supp + "_"*(93-len(supp)) + "\n" + "_"*93 + "\n" + "_"*93 + "\n" + "_"*93 + "\n" + "_"*93 + "\n" + "_"*93
        else:
            supp_list = supp.split()

            if len(supp_list) > 0:
                i = 1
                temp_supp_list = [supp_list[0]]
                while  i != 0 and i <= len(supp_list) - 1:
                    temp_supp_list.append(supp_list[i])
                    if len("_".join(temp_supp_list)) <= 93:
                        supp1 = "_".join(temp_supp_list)
                        temp_supp_list[0] = supp1
                        temp_supp_list.pop(1)
                        add_back = supp_list[i]
                        i += 1
                    elif len(supp_list) == 1:
                        supp1 = supp_list[0]
                    else:
                        i = 0
                for item in supp1.split("_"):
                    if item in supp_list:
                        supp_list.remove(item)
                if len(supp_list) > 0:
                    supp_list.insert(0,add_back)
            else:
                supp1 = ""

            if len(supp_list) > 0:
                i = 1
                supp_list = supp_list[len(supp1.split()):]
                temp_supp_list = [supp_list[0]]
                if len(supp_list) - 1 > 0:
                    while i != 0 and i <= len(supp_list) - 1:
                        temp_supp_list.append(supp_list[i])
                        if len("_".join(temp_supp_list)) <= 93:
                            supp2 = "_".join(temp_supp_list)
                            temp_supp_list[0] = supp2
                            temp_supp_list.pop(1)
                            add_back = supp_list[i]
                            i += 1
                        else:
                            i = 0
                elif len(supp_list) == 1:
                    supp2 = supp_list[0]
                else:
                    supp2 = ""
                
                for item in supp2.split("_"):
                    if item in supp_list:
                        supp_list.remove(item)
                if len(supp_list) > 0:
                    supp_list.insert(0,add_back)
            else:
                supp2 = ""
            
            if len(supp_list) > 0:
                i = 1
                supp_list = supp_list[len(supp2.split()):]
                temp_supp_list = [supp_list[0]]
                if len(supp_list) - 1 > 0:
                    while i != 0 and i <= len(supp_list) - 1:
                        temp_supp_list.append(supp_list[i])
                        if len("_".join(temp_supp_list)) <= 93:
                            supp3 = "_".join(temp_supp_list)
                            temp_supp_list[0] = supp3
                            temp_supp_list.pop(1)
                            add_back = supp_list[i]
                            i += 1
                        else:
                            i = 0
                elif len(supp_list) == 1:
                    supp3 = supp_list[0]
                else:
                    supp3 = ""
                
                for item in supp3.split("_"):
                    if item in supp_list:
                        supp_list.remove(item)
                if len(supp_list) > 0:
                    supp_list.insert(0,add_back)
            else:
                supp3 = ""

            if len(supp_list) > 0:
                i = 1
                supp_list = supp_list[len(supp3.split()):]
                temp_supp_list = [supp_list[0]]
                if len(supp_list) - 1 > 0:
                    while i != 0 and i <= len(supp_list) - 1:
                        temp_supp_list.append(supp_list[i])
                        if len("_".join(temp_supp_list)) <= 93:
                            supp4 = "_".join(temp_supp_list)
                            temp_supp_list[0] = supp4
                            temp_supp_list.pop(1)
                            add_back = supp_list[i]
                            i += 1
                        else:
                            i = 0
                elif len(supp_list) == 1:
                    supp4 = supp_list[0]
                else:
                    supp4 = ""
                
                for item in supp4.split("_"):
                    if item in supp_list:
                        supp_list.remove(item)
                if len(supp_list) > 0:
                    supp_list.insert(0,add_back)
            else:
                supp4 = ""
            
            if len(supp_list) > 0:
                i = 1
                supp_list = supp_list[len(supp4.split()):]
                temp_supp_list = [supp_list[0]]
                if len(supp_list) - 1 > 0:
                    while i != 0 and i <= len(supp_list) - 1:
                        temp_supp_list.append(supp_list[i])
                        if len("_".join(temp_supp_list)) <= 93:
                            supp5 = "_".join(temp_supp_list)
                            temp_supp_list[0] = supp5
                            temp_supp_list.pop(1)
                            add_back = supp_list[i]
                            i += 1
                        else:
                            i = 0
                elif len(supp_list) == 1:
                    supp5 = supp_list[0]
                else:
                    supp5 = ""
                
                for item in supp5.split("_"):
                    if item in supp_list:
                        supp_list.remove(item)
                if len(supp_list) > 0:
                    supp_list.insert(0,add_back)
            else:
                supp5 = ""

            if len(supp_list) > 0:
                i = 1
                supp_list = supp_list[len(supp5.split()):]
                temp_supp_list = [supp_list[0]]
                if len(supp_list) - 1 > 0:
                    while i != 0 and i <= len(supp_list) -1:
                        temp_supp_list.append(supp_list[i])
                        if len("_".join(temp_supp_list)) <=93:
                            supp6 = "_".join(temp_supp_list)
                            temp_supp_list[0] = supp6
                            temp_supp_list.pop(1)
                            add_back = supp_list[i]
                            i += 1
                        else:
                            i = 0
                elif len(supp_list) == 1:
                    supp6 = supp_list[0]
                else:
                    supp6 = ""
            else:
                supp6 = ""

            supp = supp1 + "_"*(93-len(supp1)) + "\n" + supp2 + "_"*(93-len(supp2)) + "\n" + supp3 + "_"*(93-len(supp3)) + "\n" + supp4 + "_"*(93-len(supp4)) + "\n" + supp5 + "_"*(93-len(supp5)) + "\n" + supp6 + "_"*(93-len(supp6))


        all_remaining_list = []

        for i in range(1,7):
            if len(request.POST.get('allremaining' + str(i))) > 0:
                all_remaining_list.append("$" + format_dollars(request.POST.get('allremaining' + str(i))))
            else:
                all_remaining_list.append(request.POST.get('allremaining' + str(i)))

        # Internal function to create three (3) values for certain form fields to account for strings which, based on length, will cause spillover and throw off
        # spacing and formatting in .docx template.  Three values are generated corresponding to 1) a string size that will fit in the alloted space in the template, 
        # 2) a string that would spillover the alloted space but whose font-size is reduced to make fit in the template, and 3) a string that would spillover the alloted
        # space but receives an ellipses ("...") at the end of the string.  All three values are populated to the template, two of which will always be empty strings.

        def format_lengthy(element,len_dict):
            
            if len(request.POST.get(element)) < len_dict['short']:
                field1 = request.POST.get(element)[:len_dict['short']] + "_"*(len_dict['short'] - len(request.POST.get(element)))
                field2 = ""
                field3 = ""
            elif len(request.POST.get(element)) < len_dict['long']:
                field1 = ""
                field2 = request.POST.get(element)[:len_dict['long']] + "_"*(len_dict['long'] - len(request.POST.get(element)))
                field3 = ""
            else:
                field1 = ""
                field2 = ""
                field3 = request.POST.get(element)[:len_dict['ellipse']] + "..." + "_"*(len_dict['ellipse'] - len(request.POST.get(element)))
            return [field1,field2,field3]

        # Dictionary of values to populate .docx template.  Various values append underscore characters to the end of the string value in order to preserve
        # the visual appearance of the federal IWO form.  These values must be calculated manually by counting the number of character spaces required in the .docx 
        # template.  The template must use fixed-width fonts of size 10 for strings that fit the alloted space and 7.5 for those that would spill over.  The
        # templates utilized in this system use fixed-width font Courier New.

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
            "date1": datetime.strftime(datetime.strptime(request.POST.get('date1'),"%Y-%m-%d"),"%m/%d/%Y") + "_"*(12-len(datetime.strftime(datetime.strptime(request.POST.get('date1'),"%Y-%m-%d"),"%m/%d/%Y"))),
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
            "remittanceidshort": format_lengthy("remittanceid",{"short":21,"long":28,"ellipse":25})[0],
            "remittanceidlong": format_lengthy("remittanceid",{"short":21,"long":28,"ellipse":25})[1],
            "remittanceidellipse": format_lengthy("remittanceid",{"short":21,"long":28,"ellipse":25})[2],
            "citycountydisttribe": request.POST.get('citycountydisttribe') + "_"*(22-len(request.POST.get('citycountydisttribe'))),
            "orderid": request.POST.get('orderid') + "_"*(40-len(request.POST.get('orderid'))),
            "orderid2": request.POST.get('orderid') + "_"*(44-len(request.POST.get('orderid'))),
            "privateindividualentity": request.POST.get('privateindividualentity') + "_"*(22-len(request.POST.get('privateindividualentity'))),
            "caseid": request.POST.get('caseid') + "_"*(40-len(request.POST.get('caseid'))),
            "caseid2": request.POST.get('caseid') + "_"*(33-len(request.POST.get('caseid'))),
            "employernameshort": format_lengthy("employername",{"short":40,"long":53,"ellipse":50})[0],
            "employernamelong": format_lengthy("employername",{"short":40,"long":53,"ellipse":50})[1],
            "employernameellipse": format_lengthy("employername",{"short":40,"long":53,"ellipse":50})[2],
            "employername2":employername_short,
            "employername3":employername_short,
            "employername4":employername_short,
            "employeraddress1": request.POST.get('employeraddress1') + "_"*(40-len(request.POST.get('employeraddress1'))),
            "employeraddress2": request.POST.get('employeraddress2') + "_"*(40-len(request.POST.get('employeraddress2'))),
            "employeraddress3": request.POST.get('employeraddress3') + "_"*(40-len(request.POST.get('employeraddress3'))),
            "fein":request.POST.get('fein') + "_"*(19-len(request.POST.get('fein'))),
            "fein2":request.POST.get('fein') + "_"*(17-len(request.POST.get('fein'))),
            "employeename": request.POST.get('employeename') + "_"*(42-len(request.POST.get('employeename'))),
            "employeename2": request.POST.get('employeename') + "_"*(47-len(request.POST.get('employeename'))),
            "employeessn": request.POST.get('employeessn') + "_"*(42-len(request.POST.get('employeessn'))),
            "employeessn2": request.POST.get('employeessn') + "_"*(22-len(request.POST.get('employeessn'))),
            "employeedob": request.POST.get('employeedob') + "_"*(42-len(request.POST.get('employeedob'))),
            "obligeename": request.POST.get('obligeename') + "_"*(42-len(request.POST.get('obligeename'))),
            "child1short":format_lengthy("child1",{"short":28,"long":37,"ellipse":34})[0],
            "child1long":format_lengthy("child1",{"short":28,"long":37,"ellipse":34})[1],
            "child1ellipse":format_lengthy("child1",{"short":28,"long":37,"ellipse":34})[2],
            "child1dob":request.POST.get('child1dob') + "_"*(17-len(request.POST.get('child1dob'))),
            "child2short":format_lengthy("child2",{"short":28,"long":37,"ellipse":34})[0],
            "child2long":format_lengthy("child2",{"short":28,"long":37,"ellipse":34})[1],
            "child2ellipse":format_lengthy("child2",{"short":28,"long":37,"ellipse":34})[2],            
            "child2dob":request.POST.get('child2dob') + "_"*(17-len(request.POST.get('child2dob'))),
            "child3short":format_lengthy("child3",{"short":28,"long":37,"ellipse":34})[0],
            "child3long":format_lengthy("child3",{"short":28,"long":37,"ellipse":34})[1],
            "child3ellipse":format_lengthy("child3",{"short":28,"long":37,"ellipse":34})[2],            
            "child3dob":request.POST.get('child3dob') + "_"*(17-len(request.POST.get('child3dob'))),
            "child4short":format_lengthy("child4",{"short":28,"long":37,"ellipse":34})[0],
            "child4long":format_lengthy("child4",{"short":28,"long":37,"ellipse":34})[1],
            "child4ellipse":format_lengthy("child4",{"short":28,"long":37,"ellipse":34})[2],            
            "child4dob":request.POST.get('child4dob') + "_"*(17-len(request.POST.get('child4dob'))),
            "child5short":format_lengthy("child5",{"short":28,"long":37,"ellipse":34})[0],
            "child5long":format_lengthy("child5",{"short":28,"long":37,"ellipse":34})[1],
            "child5ellipse":format_lengthy("child5",{"short":28,"long":37,"ellipse":34})[2],            
            "child5dob":request.POST.get('child5dob') + "_"*(17-len(request.POST.get('child5dob'))),
            "child6short":format_lengthy("child6",{"short":28,"long":37,"ellipse":34})[0],
            "child6long":format_lengthy("child6",{"short":28,"long":37,"ellipse":34})[1],
            "child6ellipse":format_lengthy("child6",{"short":28,"long":37,"ellipse":34})[2],            
            "child6dob":request.POST.get('child6dob') + "_"*(17-len(request.POST.get('child6dob'))),
            "orderfromstate":request.POST.get('orderfromstate') + "_"*(43-len(request.POST.get('orderfromstate'))),
            "dollar1":"_"*(12-len(format_dollars(request.POST.get('dollar1')))) + format_dollars(request.POST.get('dollar1')),
            "per1":request.POST.get('per1') + "_"*(13-len(request.POST.get('per1'))),
            "dollar2":"_"*(12-len(format_dollars(request.POST.get('dollar2')))) + format_dollars(request.POST.get('dollar2')),
            "per2":request.POST.get('per2') + "_"*(13-len(request.POST.get('per2'))),
            "yes12":request.POST.get('yes12'),
            "no12":request.POST.get('no12'),
            "dollar3":"_"*(12-len(format_dollars(request.POST.get('dollar3')))) + format_dollars(request.POST.get('dollar3')),
            "per3":request.POST.get('per3') + "_"*(13-len(request.POST.get('per3'))),
            "dollar4":"_"*(12-len(format_dollars(request.POST.get('dollar4')))) + format_dollars(request.POST.get('dollar4')),
            "per4":request.POST.get('per4') + "_"*(13-len(request.POST.get('per4'))),
            "dollar5":"_"*(12-len(format_dollars(request.POST.get('dollar5')))) + format_dollars(request.POST.get('dollar5')),
            "per5":request.POST.get('per5') + "_"*(13-len(request.POST.get('per5'))),
            "dollar6":"_"*(12-len(format_dollars(request.POST.get('dollar6')))) + format_dollars(request.POST.get('dollar6')),
            "per6":request.POST.get('per6') + "_"*(13-len(request.POST.get('per6'))),
            "dollar7":"_"*(12-len(format_dollars(request.POST.get('dollar7')))) + format_dollars(request.POST.get('dollar7')),
            "per7":request.POST.get('per7') + "_"*(13-len(request.POST.get('per7'))),
            "other":request.POST.get('other') + "_"*(40-len(request.POST.get('other'))),
            "totalwithhold":"_"*(10-len(format_dollars(request.POST.get('totalwithhold')))) + format_dollars(request.POST.get('totalwithhold')),
            "per8":request.POST.get('per8') + "_"*(13-len(request.POST.get('per8'))),
            "permonth":"_"*(11-len(format_dollars(request.POST.get('permonth')))) + format_dollars(request.POST.get('permonth')),
            "pertwoweeks":"_"*(10-len(format_dollars(request.POST.get('pertwoweeks')))) + format_dollars(request.POST.get('pertwoweeks')),
            "persemimonth":"_"*(11-len(format_dollars(request.POST.get('persemimonth')))) + format_dollars(request.POST.get('persemimonth')),
            "perweek":"_"*(10-len(format_dollars(request.POST.get('perweek')))) + format_dollars(request.POST.get('perweek')),
            "lumpsum":"_"*(10-len(format_dollars(request.POST.get('lumpsum')))) +format_dollars(request.POST.get('lumpsum')),
            "documentid":request.POST.get('documentid') + "_"*(27-len(request.POST.get('documentid'))),
            "principal":request.POST.get('principal') + "_"*(17-len(request.POST.get('statetribeterritory'))),
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
            "supplemental":supp,
            "sendershort":format_lengthy("sender",{"short":28,"long":37,"ellipse":34})[0],
            "senderlong":format_lengthy("sender",{"short":28,"long":37,"ellipse":34})[1],
            "senderellipse":format_lengthy("sender",{"short":28,"long":37,"ellipse":34})[2],
            "sendertel":request.POST.get('sendertel') + "_"*(14-len(request.POST.get('sendertel'))),
            "senderfax":request.POST.get('senderfax') + "_"*(13-len(request.POST.get('senderfax'))),
            "senderwebsite":request.POST.get('senderwebsite') + "_"*(29-len(request.POST.get('senderwebsite'))),
            "noticetoshort":format_lengthy("noticeto",{"short":89,"long":120,"ellipse":117})[0],
            "noticetolong":format_lengthy("noticeto",{"short":89,"long":120,"ellipse":117})[1],
            "noticetoellipse":format_lengthy("noticeto",{"short":89,"long":120,"ellipse":117})[2],
            "sender2short":format_lengthy("sender2",{"short":25,"long":33,"ellipse":30})[0],
            "sender2long":format_lengthy("sender2",{"short":25,"long":33,"ellipse":30})[1],
            "sender2ellipse":format_lengthy("sender2",{"short":25,"long":33,"ellipse":30})[2],
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
            "arrearspayment":format_dollars(request.POST.get('arrearspayment')),
            "arrearspaymentper":request.POST.get('arrearspaymentper'),
            "deductfull":request.POST.get('deductfull'),
            "deductpercent":request.POST.get('deductpercent'),
            "deductpercentamount":request.POST.get('deductpercentamount') + "%",
            "deductnone":request.POST.get('deductnone'),
            "childinit1":request.POST.get('childinit1'),
            "childdobtwo1":request.POST.get('childdobtwo1'),
            "dob181":request.POST.get('dob181'),
            "allremaining1":all_remaining_list[0],
            "childinit2":request.POST.get('childinit2'),
            "childdobtwo2":request.POST.get('childdobtwo2'),
            "dob182":request.POST.get('dob182'),
            "allremaining2":all_remaining_list[1],
            "childinit3":request.POST.get('childinit3'),
            "childdobtwo3":request.POST.get('childdobtwo3'),
            "dob183":request.POST.get('dob183'),
            "allremaining3":all_remaining_list[2],
            "childinit4":request.POST.get('childinit4'),
            "childdobtwo4":request.POST.get('childdobtwo4'),
            "dob184":request.POST.get('dob184'),
            "allremaining4":all_remaining_list[3],
            "childinit5":request.POST.get('childinit5'),
            "childdobtwo5":request.POST.get('childdobtwo5'),
            "dob185":request.POST.get('dob185'),
            "allremaining5":all_remaining_list[4],
            "childinit6":request.POST.get('childinit6'),
            "childdobtwo6":request.POST.get('childdobtwo6'),
            "dob186":request.POST.get('dob186'),
            "allremaining6":all_remaining_list[5],
            "data_string":data_string,
            "global_format_off":request.POST.get('global_format_off'),
            "obligorname_noformat":request.POST.get('obligorname_noformat1'),
            "obligeenname_noformat":request.POST.get('obligeename_noformat1'),
            "child1_noformat":request.POST.get('child1_noformat1'),
            "child2_noformat":request.POST.get('child2_noformat1'),
            "child3_noformat":request.POST.get('child3_noformat1'),
            "child4_noformat":request.POST.get('child4_noformat1'),
            "child5_noformat":request.POST.get('child5_noformat1'),
            "child6_noformat":request.POST.get('child6_noformat1'),
        }

        for item in request.POST.dict().items():
            print(item[0])
            request.session[item[0]]=item[1]
        request.session.modified = True

        filename = request.POST.get('casenumber').replace("-","_") + "_" + datetime.now().strftime("%m.%d.%Y_%I%M%S%p") + ".docx"

        doc.render(context)
        doc2 = doc.render(context)
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        response = StreamingHttpResponse(streaming_content=buffer,content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        response["Content-Disposition"] = 'attachment;filename=' + filename
        response['Content-Encoding'] = 'UTF-8'

        log = open("log.txt","a")
        if os.stat("log.txt").st_size == 0:
            log.write("DATE" + " "*16 + "TIME" + " "*16 + "CIRCUIT" + " "*13 + "COUNTY" + " "*14 + "CASENUMBER\n")
            log.write("-"*105 + "\n")
        log.write(datetime.now().strftime("%m/%d/%Y          %I:%M:%S %p") + " "*9 + request.POST.get('countycode').split(",")[2] + " "*(20-len(request.POST.get('countycode').split(",")[2])) + request.POST.get('countycode').split(",")[1] + " "*(20 - len(request.POST.get('countycode').split(",")[1])) + request.POST.get('casenumber') + "\n")
        log.close()


    return response

