import xmlrpc.client as xmlrpclib
import base64
import os
from myVars import *
import datetime as DT
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time
from slugify import slugify
import csv

def main():
    print("Lets get started!")
    common_proxy = xmlrpclib.ServerProxy(url+'common')
    uid = common_proxy.login(DB,USER,PASS)
    models = xmlrpclib.ServerProxy('{}2/object'.format(url))
    allData = gatherData(uid, models)
    partsSpecialists = filterSales(allData)
    decryptPhotos(partsSpecialists)
    allSlugs = genSlugs(partsSpecialists)
    genVcards(partsSpecialists, allSlugs)
    allHtml = genHtml(partsSpecialists, allSlugs)
    writeCsv(partsSpecialists, allHtml, allSlugs)


def getFields(uid, models):
    return models.execute_kw(DB, uid, PASS, 'hr.employee', 'fields_get', [], {'attributes': ['string', 'help', 'type']})


def gatherData(uid, models):
    allData = models.execute_kw(DB, uid, PASS, 'hr.employee', 'search_read', [[['name', '!=', ['Benjamin Krentz', 'Albert Alexander', 'Ryan Loos']]]], {'fields': ['name', 'work_phone', 'department_id', 'job_id', 'job_title', 'work_phone', 'mobile_phone', 'work_email', 'image_256', 'x_studio_exclude_from_website', "work_location_id", "first_contract_date", "x_studio_employee_bio"]})
    
    for idx, person in enumerate(allData):
        if 'x_studio_exclude_from_website' in person:
            if person['x_studio_exclude_from_website'] == True:
                del allData[idx]
    return allData


def filterSales(allData):
    newData = []
    for person in allData:
        if not(person['department_id'][0] == 2 and person['image_256']):
            continue
        else:
            newData.append(person)
    return newData

def decryptPhotos(allData):
    for photo in allData:
            try:
                with open((f"./photos/{photo['name']}.jpg").replace(" ", "-").lower(), "wb") as fh:
                    fh.write(base64.b64decode((photo['image_256'])))
            except:
                os.remove((f"./photos/{photo['name']}.jpg").replace(" ", "-").lower())


def genVcards(partsSpecialists, allSlugs):
    vcardData = ""
    for person, slug in zip(partsSpecialists, allSlugs):
        vcardData = f'BEGIN:VCARD\n'
        vcardData += f'VERSION:3.0\n'
        vcardData += f'FN;CHARSET=UTF-8:{person["name"]}\n'
        vcardData += f'N;CHARSET=UTF-8:;{person["name"].replace(" ", ";")};;\n'
        vcardData += f'EMAIL;CHARSET=UTF-8;type=WORK,INTERNET:{person["work_email"]}\n'
        # vcardData += f'TEL;TYPE=CELL:716-219-8106\n'
        vcardData += f'TEL;TYPE=WORK,VOICE:{person["work_phone"]}\n'
        if(str(person["work_location_id"][0]) == "1"):
            vcardData += "ADR;CHARSET=UTF-8;TYPE=WORK:;;2712 West Ave;Newfane;NY;14108;USA\n"
        elif(str(person["work_location_id"][0]) == "2"):
            vcardData += "ADR;CHARSET=UTF-8;TYPE=WORK:;;611 Jamison Road;Elma;NY;14059;USA\n"
        vcardData += f'TITLE;CHARSET=UTF-8:{person["job_title"]}\n'
        vcardData += f'ORG;CHARSET=UTF-8:ConEquip Parts\n'
        vcardData += f'URL;type=WORK;CHARSET=UTF-8:https://www.conequip.com\n'
        vcardData += f'Conequip sells construction equipment parts \n'
        vcardData += f'REV:2022-04-06T21:06:34.631Z\n'
        vcardData += f'END:VCARD\n'
        
        with open((f"./vcards/{slug}.vcf"), "w") as fh:
            fh.write(vcardData)

def genSlugs(partsSpecialists):
    allSlugs = []
    for person in partsSpecialists:
        allSlugs.append(slugify(person["name"]))
    return allSlugs

def genHtml(partsSpecialists, allSlugs):
    allHtml = []
    tempHtml = ""
    tempBio = ""
    for person, slug in zip(partsSpecialists, allSlugs):
        tempHtml = ""
        tempHtml += f'<div class="container"><div class="main-body"> <!-- Breadcrumb --><nav aria-label="breadcrumb" class="main-breadcrumb"><ol class="breadcrumb"><li class="breadcrumb-item"><a href="/about-conequip-team">About Our Team</a></li><li class="breadcrumb-item active" aria-current="page">'
        tempHtml += f'{person["name"]}</li></ol></nav><!-- /Breadcrumb --><div class="row gutters-sm"><div class="col-lg-12 mb-5"><div class="card bg-light"><div class="card-body"><div class="d-flex flex-column align-items-center text-center"> <img src="/pub/media/wysiwyg/conequip/employees/2022/'
        tempHtml += f'{slug}.jpg" alt="Admin" class="rounded-circle border" width="150"><div class="mt-3"><h4>'
        tempHtml += f'{person["name"]}</h4><p class="text-secondary mb-1">'
        tempHtml += f'{person["job_title"]}</p><a href="/vcf/'
        tempHtml += f'{slug}.vcf"><button class="btn btn-primary m-2"><i class="fa fa-address-card-o"></i> Download Contact To Your Phone</button></a><br><a href="mailto:'
        tempHtml += f'{person["work_email"]}"><button class="btn btn-outline-primary m-2"><i class="fa fa-envelope-o"></i> Email</button></a> <a href="sms:'
        tempHtml += f'{person["work_phone"]}"><button class="btn btn-outline-primary m-2"><i class="fa fa-comment-o"></i> Text</button></a> <a href="tel:'
        tempHtml += f'{person["work_phone"]}"><button class="btn btn-outline-primary m-2"><i class="fa fa-phone"></i> Call</button></a> </div></div></div></div></div><div class="col-lg-12"><div class="card mb-5"><div class="card-body bg-light"><div class="row"><div class="col-sm-3"><h6 class="mb-0"><strong>Full Name:</strong></h6></div><div class="col-sm-9 text-secondary"> '
        tempHtml += f'{person["name"]}</div></div><hr><div class="row"><div class="col-sm-3"><h6 class="mb-0"><strong>Joined ConEquip:</strong></h6></div><div class="col-sm-9 text-secondary">'
        tempHtml += f'{month[person["first_contract_date"][5:7]]} {person["first_contract_date"][0:4]}</div></div><hr><div class="row"><div class="col-sm-3"><h6 class="mb-0"><strong>Email:</strong></h6></div><div class="col-sm-9 text-secondary">'
        tempHtml += f'{person["work_email"]}</div></div><hr><div class="row"><div class="col-sm-3"><h6 class="mb-0"><strong>Direct Phone:</strong></h6></div><div class="col-sm-9 text-secondary"> <a href="tel:'
        tempHtml += f'{person["work_phone"]}">'
        tempHtml += f'{person["work_phone"]}</a> </div></div><hr><div class="row"><div class="col-sm-3"><h6 class="mb-0"><strong>Location:</strong></h6></div><div class="col-sm-9 text-secondary"> '
        if(str(person["work_location_id"][0]) == "1"):
            tempHtml += "Newfane, NY"
        elif(str(person["work_location_id"][0]) == "2"):
            tempHtml += "Elma, NY"
        tempHtml += f'</div></div><hr><div class="row"><div class="col-sm-3"><h6 class="mb-0"><strong>Bio:</strong></h6><br></div><div class="col-sm-12">'
        if person["x_studio_employee_bio"] != False:
            tempBio = person["x_studio_employee_bio"].split('\n')
            for line in tempBio:
                if line != "":
                    tempHtml += f"<p>{line}</p>"
        else:
            tempHtml += f'<p>Coming Soon...</p>'
        tempHtml += f'</div></div></div></div></div></div></div></div>'
        allHtml.append(tempHtml)
    return allHtml


def writeCsv(partsSpecialists, allHtml, allSlugs):
    header = ['name', 'slug', 'html', 'metaTitle', 'metaDescription']
    data = []
    for person, html, slug in zip(partsSpecialists, allHtml, allSlugs):
        data.append([person["name"], slug, html, f'{person["name"]} - {person["job_title"]} - ConEquip Parts and Equipment', f'{person["name"]} - {person["job_title"]} - ConEquip Parts and Equipment since {month[person["first_contract_date"][5:7]]} {person["first_contract_date"][0:4]}'])

    with open('empBios.csv', 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(data)


if __name__ == "__main__":
    main()
