#!/usr/bin/python3
#######################################################################
## MORZ-WILMA: Wichtige Informationen Leserlich am Monitor Angezeigt ##
##                                                                   ##
## jesko.anschuetz@linuxmuster.net - Juni 2018	                     ##
##                                                                   ##
#######################################################################
import pytz, requests, sys
from importlib import reload
from datetime import datetime, timedelta
from ics import Calendar
from urllib.request import urlopen
reload(sys)

url = "https://nextcloud.deinserver.de/remote.php/dav/public-calendars/PjstHCE6iXRtCHM9?export"
html_dateiname = "/var/www/html/wilma.html"
css_dateiname  = "wilma.css"
meta_refresh_rate = "5"
print("############# Programmstart ##############")
# aktuelle Zeit ermitteln
now_utc = datetime.now(pytz.utc)
tz      = pytz.timezone('Europe/Berlin')
now		= tz.normalize(now_utc)
print(now)

# HTML-Datei vorbereiten
htmlkopf='\
        <html>\
                <head>\
                        <link rel="stylesheet" href="'+css_dateiname+'">\
                        <meta charset="UTF-8">\
                        <meta http-equiv="refresh" content="'+meta_refresh_rate+'">\
                        <title>WILMA - Wichtige Informationen Leserlich am Monitor Angezeigt</title>\
                </head>\
                <body>'
htmlbody='\
                        <div class="kopf">\
                        <span class="ueberschrift">MoRZ-WILMA</span>\
                        <span class="ueberschrift2"><span class="fett">W</span>ichtige <span class="fett">I</span>nformationen <span class="fett">L</span>esbar am <span class="fett">M</span>onitor <span class="fett">A</span>ngezeigt</span>\
                        <div class="zeit">'+str(now)+'</div>'
htmlfuss = "\
                </body>\
        </html>"

inhalt = htmlkopf + htmlbody

# Kalender abfragen
c = Calendar(requests.get(url).text)
try:
        for t in c.todos:
                if t.percent < 100:
                        div = "eintragtodo"
                        print("------------------------------------------------------------")
                        print("-- {} -- Erzeuge Todo '{}', {}% erledigt".format(now, t.name, t.percent))
                        inhalt+='<div class="'+div+'">'
                        inhalt+='<div class="titel">' + t.name + ' (Prio:' + t.percent + ')' '</div>'
                        inhalt+="</div>"
except all:
        print("Todos lesen klappt nicht")
try:
        # Events durchlaufen
        for e in c.events:
           show = (e.end > now) & (now > e.begin)
           if show:
                 alter = abs(e.begin-now).seconds
                 if alter < 3600:
                     div = "eintragneu"
                 else:
                     div = "eintragalt"
                 print("------------------------------------------------------------")
                 print("-- {} -- Erzeuge Event '{}' Startzeit: {} Endzeit {}".format(now, e.name, e.begin, e.end))
                 print("Alter in Secunden: {}".format(str(abs(e.begin-now).seconds)))
                 inhalt+='<div class="'+div+'">'
                 inhalt+='<div class="titel">' + e.name + '</div>'
                 if repr(e.location) != "None":
                        inhalt+='<div class="ort">'+e.location+'</div>'
                 if repr(e.description) != "None":
                        inhalt+='<div class="beschreibung">'+e.description+'</div>'
                 inhalt+="</div>"
except all:
        inhalt+='<div class="eintrag">'
        inhalt+='<div class="titel">etwas ist schief gelaufen bei der Abfrage</div>'
        inhalt+='</div>'

inhalt+=htmlfuss
try:
	datei = open(html_dateiname,"w")
	datei.write(inhalt)
	datei.close()
except all:
	print("error opening/creating file")
