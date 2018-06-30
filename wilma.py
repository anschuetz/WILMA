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

debugflag = True
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
# Den Kopf:
htmlkopf ='<html>'
htmlkopf+=' <head>'
htmlkopf+='   <link rel="stylesheet" href="'+css_dateiname+'">'
htmlkopf+='   <meta charset="UTF-8">'
htmlkopf+='   <meta http-equiv="refresh" content="'+meta_refresh_rate+'">'
htmlkopf+='   <title>WILMA - Wichtige Informationen Leserlich am Monitor Angezeigt</title>'
htmlkopf+=' </head>'
htmlkopf+=' <body>'

# Die Einleitung vom Hauptteil
htmlbody ='    <div class="kopf">'
htmlbody+='    <span class="ueberschrift">MoRZ-WILMA</span>'
htmlbody+='    <span class="ueberschrift2">'
htmlbody+='        <span class="fett">W</span>ichtige '
htmlbody+='        <span class="fett">I</span>nformationen '
htmlbody+='        <span class="fett">L</span>esbar am '
htmlbody+='        <span class="fett">M</span>onitor '
htmlbody+='        <span class="fett">A</span>ngezeigt'
htmlbody+='    </span>'
htmlbody+='    <div class="zeit">'+str(now)+'</div>'

# Den Abschluss der Seite
htmlfuss =' </body>'
htmlfuss+='</html>'
inhalt = " "
# Kalender abfragen
try:
        c = Calendar(requests.get(url).text)
        if debugflag:
                        print(c)
        prio = 9
        # Nach Priorität geordnet ausgeben, beginnend mit der höchsten.
        while prio >= 0:
                for t in c.todos:
                        # falls keine Fälligkeitszeit angegeben ist, anzeigen, ansonsten prüfen ob noch aktuell
                        if repr(t.due) == "None":
                                vor_ende = True
                        else:
                                vor_ende = (t.due > now)
                        # falls keine Startzeit angegeben ist, anzeigen, ansonsten prüfen ob schon aktuell
                        if repr(t.begin) == "None":
                                nach_beginn = True
                        else:
                                nach_beginn = (now > t.begin)
                        # wenn noch nicht 100% UND Priorität stimmt und zwischen den Daten, dann anzeigen.
                        # Nextcloud speichert die Priorität falschrum (1 = hoch und 9 = niedrig. 0 = ohne)
                        if (t.percent < 100) & (((10-t.priority) % 10) == prio) & vor_ende & nach_beginn:
                                # css-Klasse passend zur Priorität auswählen
                                div_klasse = "eintragtodo"+str(prio)
                                if debugflag:
                                                print("------------------------------------------------------------")
                                                print("-- {} -- Erzeuge Todo '{}', {}% erledigt, Priorität-Task {}, Priorität korrigiert: {} ".format(now, t.name, t.percent, t.priority, prio))
                                inhalt+='   <div class="'+div_klasse+'">'
                                inhalt+='     <div class="titel">' + t.name
                                if debugflag:
                                                #nur im debugmodus die Priorität dazu schreiben
                                                inhalt+=' (Prio:' + str(t.priority) + '/'+ str(prio) +')'
                                inhalt+='     </div>'
                                # Nur bei gesetzter Beschreibung, diese auch ausgeben
                                if repr(t.description) != "None":
                                        inhalt+='<div class="beschreibung">'+t.description+'</div>'
                                inhalt+='   </div>'
                prio -= 1
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
                 if debugflag:
                        print("------------------------------------------------------------")
                        print("-- {} -- Erzeuge Event '{}' Startzeit: {} Endzeit {}".format(now, e.name, e.begin, e.end))
                        print("Alter in Secunden: {}".format(str(abs(e.begin-now).seconds)))
                 inhalt+='   <div class="'+div+'">'
                 inhalt+='     <div class="titel">' + e.name + '</div>'
                 if repr(e.location) != "None":
                        inhalt+='     <div class="ort">'+e.location+'</div>'
                 if repr(e.description) != "None":
                        inhalt+='     <div class="beschreibung">'+e.description+'</div>'
                 inhalt+="   </div>"
except all:
        inhalt+='<div class="eintrag">'
        inhalt+='<div class="titel">etwas ist schief gelaufen bei der Abfrage</div>'
        inhalt+='</div>'

# Seite zusammensetzen.

htmlseite = htmlkopf + htmlbody + inhalt + htmlfuss

try:
	datei = open(html_dateiname,"w")
	datei.write(htmlseite)
	datei.close()
except all:
	print("error opening/creating file")
