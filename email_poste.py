import csv
import xlrd
import urllib
import urllib2
import sys
import cookielib
import pycurl
import json
import io
import requests
import os
from datetime import datetime

db_from_poste	= 'test.txt'
file_addr_poste	= 'prova2.xlsx'
url		= 'https://myposteimpresa.poste.it/online/login'
username	= 'xxx'
password	= 'xxx'


list_city = []
with open(db_from_poste, 'rb') as csvfile:
  spamreader = csv.reader(csvfile, delimiter=',', quotechar='\n')
  for row in spamreader:
    list_city.append(row)
    
    
    
# Open the workbook
xl_workbook = xlrd.open_workbook(file_addr_poste)

# List sheet names, and pull a sheet by name
#
sheet_names = xl_workbook.sheet_names()
#print('Sheet Names', sheet_names)

xl_sheet = xl_workbook.sheet_by_index(0)
#Pull the first row by index
#  (rows/columns are also zero-indexed)
#
row = xl_sheet.row(0)  # 1st row

# Print 1st row values and types
#
from xlrd.sheet import ctype_text   

#print('(Column #) type:value')
for idx, cell_obj in enumerate(row):
  cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
  #print('(%s) %s %s' % (idx, cell_type_str, cell_obj.value))

#Authentication Start
form = { "username" : username,"password" : password}
encodedForm = urllib.urlencode(form)
USER_AGENT = 'Mozilla/5.0 (X11; U; Linux i686; en-GB; rv:1.9.0.5) Gecko/2008121622 Ubuntu/8.10 (intrepid) Firefox/3.0.5'
output  = io.BytesIO()
c = pycurl.Curl()
c.setopt(c.URL, url)
c.setopt(pycurl.FOLLOWLOCATION, 1)
c.setopt(c.POSTFIELDS, encodedForm)
c.setopt(pycurl.POST, 1)
c.setopt(c.VERBOSE, False)
c.setopt(pycurl.WRITEFUNCTION, output.write)
c.setopt(pycurl.COOKIEFILE,'cookie.txt')
c.setopt(pycurl.ENCODING, 'gzip, deflate')
c.setopt(pycurl.USERAGENT, USER_AGENT)
c.setopt(pycurl.CONNECTTIMEOUT, 30)
c.setopt(pycurl.TIMEOUT, 30)
c.setopt(pycurl.SSL_VERIFYPEER, 0);
c.perform()
#Authentication End
print "Process started("+str(datetime.now())+"): Authentication performed!"


error    = []
imported = []

num_cols = xl_sheet.ncols   # Number of columns
for row_idx in range(0, xl_sheet.nrows):    # Iterate through rows
  if row_idx>0:
    print ('-'*40)
    print ('Parsing row: %s...' % row_idx)   # Print row number
    #for col_idx in range(0, num_cols):  # Iterate through columns
    #  cell_obj = xl_sheet.cell(row_idx, col_idx) # Get cell object by row, col
    #  print ('Column: [%s] cell_obj: [%s]' % (col_idx, int(cell_obj.value) if col_idx==2 else cell_obj.value))
    
    azienda		= xl_sheet.cell(row_idx, 0).value
    indirizzo		= xl_sheet.cell(row_idx, 1).value
    cap			= int(xl_sheet.cell(row_idx, 2).value)
    comune		= xl_sheet.cell(row_idx, 3).value
    provincia		= xl_sheet.cell(row_idx, 4).value
    sito		= xl_sheet.cell(row_idx, 5).value
    indirizzo_sito	= xl_sheet.cell(row_idx, 6).value
    documento		= xl_sheet.cell(row_idx, 7).value
    address_present	= False
    for city in list_city:
      if (int(city[3])==int(cap) or int(cap)==50020 or int(cap) in( 95100, 80100) ):
	address_present	= True
	break;
    if address_present:
      nominativo_base	= azienda
      nominativo_plus 	= '(Prot. '+sito+')'
      indirizzo_base	= indirizzo
      indirizzo_plus	= ''
      provincia		= provincia
      citta		= comune
      cap		= cap
      filename		= "RAC-Lotto45_Complete/"+documento

      url_servizio="https://raccomandata-ce.poste.it/rol-lc-business/rest/RecuperaIdRichiestaRestService/getIdRichiestaApplicationJson"
      data_for_id = json.dumps({"idServizio":"ROL", "numeroIdRichieste":1})
      output = io.BytesIO()
      c.setopt(c.URL, url_servizio)
      c.setopt(pycurl.HTTPHEADER, ['Accept: application/json', 'Accept-Encoding:gzip, deflate, br', 'Accept-Language:it-IT,it;q=0.8,en-US;q=0.6,en;q=0.4','X-Requested-With:XMLHttpRequest'])
      c.setopt(pycurl.POST, 1)
      c.setopt(pycurl.POSTFIELDS, data_for_id)
      c.setopt(pycurl.WRITEFUNCTION, output.write)
      try:
	c.perform()
      except ValueError:
	error.append({"riga":row_idx,"tipo":"Connection", "descrizione": "connessione fallita"})
	continue;
	

      try:
	id_richiesta=json.loads(output.getvalue(),'utf-8')['requestObject'][u'idrichieste'][0]
      except ValueError:
	error.append({"riga":row_idx,"tipo":"Problema ID richiesta ", "descrizione": id_richiesta})
      if not(id_richiesta):
	error.append({"riga":row_idx,"tipo":"Problema ID richiesta ", "descrizione": id_richiesta})


      method = 0

      url_upload = 'https://raccomandata-ce.poste.it/rol-lc-business/rest/UploadMultipleFilesService/UploadMultipleFiles'
      output0  = io.BytesIO()
      c.setopt(c.POST, 1)
      c.setopt(c.URL, url_upload)
      c.setopt(pycurl.WRITEFUNCTION,	output0.write)
      if not (os.path.isfile(filename)):
	error.append({"riga":row_idx,"tipo":"File non trovato", "descrizione": filename})
	continue;

      filesize = os.path.getsize(filename)
      lunghezza='Content-Length'+str(filesize)
      fin = open(filename, 'rb')

      c.setopt(
	  c.HTTPPOST,
	  [
	      ("attachment", (c.FORM_FILE, filename, c.FORM_CONTENTTYPE, "application/pdf")),
	      ("text",""),
	      ("document_id",str(id_richiesta)),
	      ("text_before_attachments", "false")
	  ]
      )

      c.setopt(pycurl.HTTPHEADER, [
	'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
	'Accept-Encoding:gzip, deflate, br',
	'Accept-Language:it-IT,it;q=0.8,en-US;q=0.6,en;q=0.4',
	'Connection:keep-alive',
	'Cache-Control:no-cache',
	'Content-Type:multipart/form-data',
	'Host:raccomandata-ce.poste.it',
	'Origin:https://raccomandata-ce.poste.it',
	'Pragma:no-cache',
	'Referer:https://raccomandata-ce.poste.it/rol-lc-business/rol-lc-business.do',
	'Upgrade-Insecure-Requests:1',
	'User-Agent:Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',

	])
      try:
	c.perform()
      except ValueError:
	error.append({"riga":row_idx,"tipo":"Connection", "descrizione": "connessione fallita"})
	continue;



      #invio ROL
      output5 = io.BytesIO()
      url_rol="https://raccomandata-ce.poste.it/rol-lc-business/rest/ROLRestService/InvioROL"
      c.setopt(c.URL, url_rol)
      c.setopt(pycurl.HTTPHEADER, ['Accept:application/json, text/javascript, */*; q=0.01',
			      'Accept-Encoding:gzip, deflate, br',
			      'Accept-Language:it-IT,it;q=0.8,en-US;q=0.6,en;q=0.4',
			      'Cache-Control:no-cache',
			      'Connection:keep-alive',
			      'Content-Length:630',
			      'Content-Type:application/json',
			      'Host:raccomandata-ce.poste.it',
			      'Origin:https://raccomandata-ce.poste.it',
			      'Pragma:no-cache',
			      'Referer:https://raccomandata-ce.poste.it/rol-lc-business/rol-lc-business.do',
			      'User-Agent:Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',
			      'X-Requested-With:XMLHttpRequest'
			      ])
      c.setopt(pycurl.POST, 1)
      c.setopt(pycurl.WRITEFUNCTION, output5.write)
      data_for_creation="""{"newIdRichiesta":"%s","oldIdRichiesta":null,"guid":null,"urlDocumento":"true?11111","utente":{"password":null,"nicknameAzienda":"divitelspa","tipoContratto":"PosteIt","idCliente":"","codiceCliente":"713.834.572","codiceFiscale":"rssvna69l47l840h","userId":"vania.rossetto.divitelspa","emailUtente":"divitelspa@bacheca-business.poste.it","emailAzienda":"divitelspa@bacheca-business.poste.it","centroDiFatturazione":null,"centroDiCosto":null,"tipoUtente":"B","idSender":"999980","senderSystem":"WEB","partner":"PCOM","tipoClient":"WSPCOMCE","partitaIva":"02852780242"},"mittente":{"provincia":"VI","cap":"36077","frazione":null,"stato":"italia","nominativo":"WIND TRE S.p.A.","complementoNominativo":null,"indirizzo":"CP N 19","complementoIndirizzo":null,"citta":"ALTAVILLA VICENTINA","zona":null,"tipoIndirizzo":{"_value_":"CASELLA POSTALE"},"casellaPostale":"N 19","ufficioPostale":"ALTAVILLA VICENTINA","forzaDestinazione":false,"otherAttributes":{}},"destinatari":[{"nominativo":{"nominativo":"%s","indirizzo":"%s","complementoNominativo":"%s","complementoIndirizzo":"%s","provincia":"%s","citta":"%s","cap":"%05d","stato":"Italia","tipoIndirizzo":{"_value_":"INDIRIZZOFISICO"}}}],"documento":{"md5":"md5","tipoDocumento":"pdf","nomeDocumento":"nomeDocumento"},"opzioni":{"opzionidiStampa":{"bw":"true","fronteRetro":"true","resolutionX":"300","resolutionY":"300","pageSize":{"_value_":"A4"}}},"ricevuta": {"provincia": "VI","cap": "36077","frazione":"","stato": "italia","nominativo": "WIND TRE S.p.A.","complementoNominativo": null,"indirizzo": "CP N 19","complementoIndirizzo": null,"citta": "ALTAVILLA VICENTINA","zona": null,"tipoIndirizzo": {"_value_": "CASELLA POSTALE"},"casellaPostale": "19","ufficioPostale": "ALTAVILLA VICENTINA","forzaDestinazione": false,"otherAttributes": {}}}""" % (id_richiesta,nominativo_base,indirizzo_base,nominativo_plus,indirizzo_plus,provincia,citta,cap)

     
      data_for_creation=json.dumps(json.loads(data_for_creation))
      c.setopt(pycurl.POSTFIELDS, data_for_creation)
      problema= 'Content-Length:'+str(len(data_for_creation))
      c.setopt(pycurl.HTTPHEADER, ['Accept:application/json, text/javascript, */*; q=0.01',
			      'Accept-Encoding:gzip, deflate, br',
			      'Accept-Language:it-IT,it;q=0.8,en-US;q=0.6,en;q=0.4',
			      'Cache-Control:no-cache',
			      'Connection:keep-alive',
			      problema,
			      'Content-Type:application/json',
			      'Host:raccomandata-ce.poste.it',
			      'Origin:https://raccomandata-ce.poste.it',
			      'Pragma:no-cache',
			      'Referer:https://raccomandata-ce.poste.it/rol-lc-business/rol-lc-business.do',
			      'User-Agent:Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',
			      'X-Requested-With:XMLHttpRequest'
			      ])
      c.setopt(pycurl.POSTFIELDSIZE, len(data_for_creation))
      try:
	c.perform()
      except ValueError:
	error.append({"riga":row_idx,"tipo":"Connection", "descrizione": "connessione fallita"})
	continue;
      j = json.loads(output5.getvalue())

      if j['error']:
	if "nominativo" in j['error']['description']:
	  error.append({"riga":row_idx,"tipo":"Lunghezza Destinatario", "descrizione": j['error']['description']})
	elif "indirizzo" in j['error']['description']:
	  error.append({"riga":row_idx,"tipo":"Lunghezza Indirizzo", "descrizione": j['error']['description']})
	elif "caratteri" in j['error']['description']:
	  error.append({"riga":row_idx,"tipo":"Caratteri on validi", "descrizione": j['error']['description']})
	else:
	  error.append({"riga":row_idx,"tipo":"Generico", "descrizione": j['error']['description']})
      else:
	imported.append({"riga":row_idx,"tipo":"Riga importata correttamente"})


      #print ">-----------------------------------------------------------------------"
      #print output5.getvalue()
      #print "------------------------------------------------------------------------"

    else:
      error.append({"riga":row_idx,"tipo":"Comune errato","descrizione":comune})

print ('-'*40)
print "Totale  : "+str(xl_sheet.nrows-1)
print "Righe OK: "+str(len(imported))
print "Errori  : "+str(len(error))
print ('-'*40)

if len(error)>0:
  #print "---------------------------------"    
  for er in error:
    try:
      print "Riga: "+str(er["riga"])+ " Errore("+er["tipo"]+") -->"+str(er["descrizione"]);
    except ValueError:
      print "Riga: "+str(er["riga"])+ " Errore("+er["tipo"]+")";
  #print "---------------------------------" 

print "Process ended("+str(datetime.now())+")"
