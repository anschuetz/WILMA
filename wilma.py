#!/usr/bin/python3
#######################################################################
## MORZ-WILMA: Wichtige Informationen Leserlich am Monitor Angezeigt ##
##                                                                   ##
## jesko.anschuetz@linuxmuster.net - Juni 2018	                     ##
##                                                                   ##
#######################################################################
import pytz, requests, sys, os, configparser
from importlib import reload
from datetime import datetime, timedelta
from ics import Calendar
from urllib.request import urlopen
from xlrd import open_workbook, xldate_as_tuple
reload(sys)
debugflag = False
#debugflag = True
# Config auslesen...
config = configparser.ConfigParser()
config.sections()
config.read('/home/pi/WILMA/wilma.ini')
excelDatei = config['WILMA']['excelDatei']
url = config['WILMA']['kalenderURL']
html_dateiname = config['WILMA']['htmlDatei']
css_dateiname = config['WILMA']['cssDatei']
meta_refresh_rate = config['WILMA']['metaRefreshRate']
datumZeile  = int(config['WILMA']['datumZeile'])
datumSpalte = int(config['WILMA']['datumSpalte'])
ersteZeileKlassen = int(config['WILMA']['ersteZeileKlassen'])
############# Hier konfigurieren falls nötig:
#excelDatei = "/home/shares/infodisplay/entschuldigung.xls"
#url = "https://nextcloud.morzgut.de/remote.php/dav/public-calendars/RAKfZs4TXrtqs9FK?export"
#html_dateiname = "/var/www/html/wilma.html"
#css_dateiname  = "wilma.css"
#meta_refresh_rate = "5"
############# ab hier gibts nichts zu ändern...

# Ein paar Variablen initialisieren:
excelDateiExistiert = False
try:
  f = open(excelDatei)
  f.close()
  unixTimeStamp = os.stat(excelDatei).st_mtime
  lastModified = datetime.fromtimestamp(unixTimeStamp).strftime('%d.%m.%Y um %H:%M:%S')
  fileDatum = datetime.fromtimestamp(unixTimeStamp).strftime('%d.%m.%Y')
  excelDateiExistiert = True
except:
  excelDateiExistiert = False
  lastModified = "n/a"
  fileDatum = "n/a"

if debugflag:
   print("############# Programmstart ##############")
# aktuelle Zeit ermitteln
now_utc = datetime.now(pytz.utc)
tz      = pytz.timezone('Europe/Berlin')
now		= tz.normalize(now_utc)
if debugflag:
   print(now)

# HTML-Datei vorbereiten
# Den Kopf:
htmlkopf ='<html>\n'
htmlkopf+=' <head>\n'
htmlkopf+='   <link rel="stylesheet" href="'+css_dateiname+'">\n'
htmlkopf+='   <meta charset="UTF-8">\n'
htmlkopf+='   <meta http-equiv="refresh" content="'+meta_refresh_rate+'">\n'
htmlkopf+='   <title>WILMA - Wichtige Informationen Leserlich am Monitor Angezeigt</title>\n'
htmlkopf+=' </head>\n'
htmlkopf+=' <body>\n'

# Die Einleitung vom Hauptteil
htmlbody ='    <div class="kopf">\n'
htmlbody+='    <span class="ueberschrift">MoRZ-WILMA</span>\n'
htmlbody+='    <span class="ueberschrift2">\n'
htmlbody+='        <span class="fett">W</span>ichtige \n'
htmlbody+='        <span class="fett">I</span>nformationen \n'
htmlbody+='        <span class="fett">L</span>esbar am \n'
htmlbody+='        <span class="fett">M</span>onitor \n'
htmlbody+='        <span class="fett">A</span>ngezeigt\n'
htmlbody+='    </span>\n'
htmlbody+='    <div class="zeit">'+str(now)+'</div>\n'

# Den Abschluss der Seite
htmlfuss ='   <div class="eintrag"><div class="beschreibung">\n'
htmlfuss+="<iframe src=\"http://localhost/fehler.html\" style=\"border:0px #FFFFFF none;\" name=\"errorlog\" scrolling=\"no\" frameborder=\"0\" align=aus marginheight=\"0px\" marginwidth=\"0px\" height=\"30\" width=\"640\"></iframe>\n"
htmlfuss+="   </div></div>\n"
htmlfuss+=' </body>\n'
htmlfuss+='</html>\n'

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
                        # wenn noch nicht 100% UND Priorität #stimmt und zwischen den Daten, dann anzeigen.
                        # Nextcloud speichert die Priorität falschrum (1 = hoch und 9 = niedrig. 0 = ohne)
			# Fehler abfangen, falls Variablen nicht gesetzt sind:
                        if isinstance(t.percent, int):
                           prozent = t.percent < 100
                        else:
                           prozent = True
                        if isinstance(t.priority, int):
                           prioritaetstimmt = ((10-t.priority) % 10) == prio
                        else:
                           prioritaetstimmt = 0 == prio

                        if (prozent) & prioritaetstimmt & vor_ende & nach_beginn:
				# css-Klasse passend zur Priorität auswählen
                                div_klasse = "eintragtodo"+str(prio)
                                if debugflag:
                                                print("------------------------------------------------------------")
                                                print("-- {} -- Erzeuge Todo '{}', {}% erledigt, Priorität-Task {}, Priorität korrigiert: {} ".format(now, t.name, t.percent, t.priority, prio))
                                inhalt+='   <div class="'+div_klasse+'">\n'
                                inhalt+='     <div class="titel">' + t.name
                                if debugflag:
                                                #nur im debugmodus die Priorität dazu schreiben
                                                inhalt+=' (Prio:' + str(t.priority) + '/'+ str(prio) +')'
                                inhalt+='     </div>\n'
                                # Nur bei gesetzter Beschreibung, diese auch ausgeben
                                if repr(t.description) != "None":
                                        inhalt+='<div class="beschreibung">'+t.description+'</div>\n'
                                inhalt+='   </div>\n'
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
                 inhalt+='   <div class="'+div+'">\n'
                 inhalt+='     <div class="titel">' + e.name + '</div>\n'
                 if repr(e.location) != "None":
                        inhalt+='     <div class="ort">'+e.location+'</div>\n'
                 if repr(e.description) != "None":
                        inhalt+='     <div class="beschreibung">'+e.description+'</div>\n'
                 inhalt+="   </div>"
except all:
        inhalt+='<div class="eintrag">\n'
        inhalt+='<div class="titel">etwas ist schief gelaufen bei der Abfrage</div>\n'
        inhalt+='</div>\n'


# die Entschuldtigungen
anzahlEntschuldigteSchueler=0

entschuldigungKopf ='<div class="entschuldigung">\n'
entschuldigung=""
# Modification-Time der Datei holen (Unix-Timestamp), in Datum/Zeit konvertieren und lesbar formatieren:

if excelDateiExistiert:
  with open_workbook(excelDatei, 'rb') as excelsheet:
    datemode = excelsheet.datemode
    tabelle = excelsheet.sheet_by_index(0)

    # Datum aus Exceldatei auslesen.

    xldate = (tabelle.cell(datumZeile,datumSpalte).value)
    datumTupel = xldate_as_tuple(xldate, datemode)
    datum = "{}.{}.{}".format( datumTupel[2],datumTupel[1],datumTupel[0] )
    #datum = fileDatum

    zeilen = []
    for zeilennummer in range(ersteZeileKlassen,tabelle.nrows):
        zeilen.append(tabelle.row_values(zeilennummer))

    entschuldigungKopf+='<table>\n'
    for zeile in zeilen:
       if (zeile[0] != "") & (zeile [1] != ""):
           entschuldigung+='<tr><td><b>{}</b></td><td>'.format(zeile[0])
           zeilenpuffer = ""
           for spalte in zeile:
              if spalte == "":
                continue
              zeilenpuffer+='{}, '.format(spalte)
              anzahlEntschuldigteSchueler+=1

           zeilenpuffer=zeilenpuffer.partition(", ")[2]
           anzahlEntschuldigteSchueler-=1
           entschuldigung+= zeilenpuffer.strip(", ")
           entschuldigung+='</td></tr>\n'
    entschuldigungKopf+='<tr><th colspan="8">Am {} sind insgesamt {} Schüler entschuldigt - zuletzt aktualisiert am {}</th>\n'.format(datum, anzahlEntschuldigteSchueler, lastModified)
    entschuldigungFuss='</table></div>\n'
# Seite zusammensetzen.

if anzahlEntschuldigteSchueler > 0:
   entschuldigungsblock = entschuldigungKopf + entschuldigung + entschuldigungFuss
else:
   entschuldigungsblock = '<div class="eintragtodo2"><div class="titel">Stand {} ist kein Schüler entschuldigt!</div></div>\n'.format(lastModified)

if excelDateiExistiert:
   pass
else:
   entschuldigungsblock = '<div class="eintragtodo9"><div class="titel">Entschuldigungsdatei muss an die richtige Stelle kopiert werden!!!</div></div>\n'

htmlseite = htmlkopf + htmlbody + inhalt + entschuldigungsblock + htmlfuss

try:
	datei = open(html_dateiname,"w")
	datei.write(htmlseite)
	datei.close()
except all:
	print("error opening/creating file")
