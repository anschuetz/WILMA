#!/usr/bin/python3
import pytz, requests, sys
from datetime import datetime, timedelta
from ics import Calendar
from importlib import reload
reload(sys)

try:
	from urllib2 import urlopen
except ImportError:
	from urllib.request import urlopen
url = "https://nextcloud.yourserver.de/remote.php/dav/public-calendars/PjstHCE6iXRtCHM9?export"

# aktuelle Zeit ermitteln
now_utc=datetime.now(pytz.utc)
tz = pytz.timezone('Europe/Berlin')
now=tz.normalize(now_utc)
#print(str(now))
# HTML-Datei vorbereiten
htmlkopf='\
        <html>\
                <head>\
                        <link rel="stylesheet" href="wilma.css">\
                        <meta charset="UTF-8">\
                        <meta http-equiv="refresh" content="5">\
                        <title>Anzeige</title>\
                </head>\
                <body>'
htmlbody='\
                        <div class="kopf">\
                        <span class="ueberschrift">MoRZ-WILMA</span>\
                        <span class="ueberschrift2"><span class="fett">W</span>ichtige <span class="fett">I</span>nformationen <span class="fett">L</span>esbar am <span class="fett">M</span>onitor <span class="fett">A</span>ngezeigt</span>\
                        <div class="zeit">'+str(now)+'</div>'
htmlfuss="\
                </body>\
        </html>"

inhalt=htmlkopf+htmlbody

# Kalender abfragen
c = Calendar(requests.get(url).text)
try:
        # Events durchlaufen
        for e in c.events:
           show = (e.end > now) & (now > e.begin)
           if show:
                 print("Erzeuge Event: "+e.name)
                 inhalt+='<div class="eintrag">'
              #   if show:
              #   inhalt+='<div class="ort">Wird gezeigt ab '+str(e.begin)+' bis '+str(e.end)+'- jetzt ist: '+str(now)+'</div>'
              #   else:
              #   inhalt+='<div class="ort">Wird nicht gezeigt seit '+str(e.end)+'- jetzt ist: '+str(now)+'</div>'
                 inhalt+='<div class="titel">' + e.name + '</div>'
                 if repr(e.location) != "None":
                        inhalt+='<div class="ort"> ('+e.location+') </div>'
                 if repr(e.description) != "None":
                        inhalt+='<div class="beschreibung">'+e.description+'</div>'
                 inhalt+="</div>"
except all:
        inhalt+='<div class="eintrag">'
        inhalt+='<div class="titel">etwas ist schief gelaufen bei der Abfrage</div>'
        inhalt+='</div>'

inhalt+=htmlfuss
try:
	datei = open("/var/www/html/wilma.html","w")
	datei.write(inhalt)
	datei.close()
except all:
	print("error opening file")
