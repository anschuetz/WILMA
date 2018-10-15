#!/usr/bin/python3
# -*- coding: utf-8 -*-
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
configfile = "/home/pi/WILMA/wilma.ini"
debugflag = False
iframeFlag = False # Error.log unten auf der Seite einbinden... sollte nicht notwendig sein.
#debugflag = True
# konfiguration auslesen...
try:
        konfiguration = configparser.ConfigParser()
        konfiguration.sections()
        konfiguration.read(configfile)
        nameDerExceldatei = konfiguration['WILMA']['excelDatei']
        urlDesKalenders = konfiguration['WILMA']['kalenderURL']
        nameDerHTMLdatei = konfiguration['WILMA']['htmlDatei']
        nameDerCSSdatei = konfiguration['WILMA']['cssDatei']
        metaRefreshRate = konfiguration['WILMA']['metaRefreshRate']
        datumzelleZeile  = int(konfiguration['WILMA']['datumZeile'])
        datumzelleSpalte = int(konfiguration['WILMA']['datumSpalte'])
        ersteZeileKlassen = int(konfiguration['WILMA']['ersteZeileKlassen'])
except Exception as detail:
        print("Fehler in der Konfigurationsdatei {} beim Schlüssel {}".format(configfile, detail))
        exit(1)
def debugprint(message):
        if debugflag:
                print(message)

def fileExists(filename):
  try:
          f = open(filename)
          f.close()
          return True
  except IOError as detail:
          debugprint(detail)
          return False
  except Exception as detail:
          print("unbehandelter Fehler in Funktion 'fileExists': >>>", detail, "<<<")
          exit(1)

def getUnixTimestampFromFile(filename):
        try:
                return os.stat(filename).st_mtime
        except IOError as detail:
                print("File Not Found Error in getUnixTimestpamFromFile: ", detail)
                return 0
        except Exception as detail:
          print("unbehandelter Fehler in Funktion 'getUnixTimestampFromFile': >>>", detail, "<<<")
          exit(1)

def getDatumUhrzeitFromUnixTimestamp(timestamp):
        if timestamp == 0:
                return "kein korrektes Datum ermittelt"
        else:
                return datetime.fromtimestamp(timestamp).strftime('%d.%m.%Y um %H:%M:%S')     

def getDatumFromUnixTimestamp(timestamp):
        if timestamp == 0:
                return "kein korrektes Datum ermittelt"
        else:
                return datetime.fromtimestamp(timestamp).strftime('%d.%m.%Y')     

def jetzt(zeitzone):
        now_utc = datetime.now(pytz.utc)
        tz      = pytz.timezone(zeitzone)
        return tz.normalize(now_utc)

def createErrorHTML(was, detail):
        message = "{} klappt nicht: {}".format(was, detail)
        print(message)
        messagehtml ='<div class="eintrag">\n'
        messagehtml+='<div class="fehler">' + message + '</div>\n'
        messagehtml+='</div>\n'
        return messagehtml

# Ein paar Variablen initialisieren:
excelDateiExistiert = fileExists(nameDerExceldatei)
unixTimeStamp = getUnixTimestampFromFile(nameDerExceldatei)
lastModified = getDatumUhrzeitFromUnixTimestamp( unixTimeStamp )
fileDatum = getDatumUhrzeitFromUnixTimestamp( unixTimeStamp )

debugprint("############# Programmstart ##############")
# aktuelle Zeit ermitteln
now	= jetzt('Europe/Berlin')
debugprint(now)

# HTML-Datei vorbereiten
htmlkopf = ""
htmlbody = ""
htmlfuss = ""
inhalt = ""
htmlTodos = ""
htmlEvents = ""
htmlXLS = ""

# Den Kopf:
htmlkopf+='<html>\n'
htmlkopf+=' <head>\n'
htmlkopf+='   <link rel="stylesheet" href="'+nameDerCSSdatei+'">\n'
htmlkopf+='   <meta charset="UTF-8">\n'
htmlkopf+='   <meta http-equiv="refresh" content="'+metaRefreshRate+'">\n'
htmlkopf+='   <title>WILMA - Wichtige Informationen Leserlich am Monitor Angezeigt</title>\n'
htmlkopf+=' </head>\n'
htmlkopf+=' <body>\n'

# Die Einleitung vom Hauptteil
htmlbody+='    <div class="kopf">\n'
htmlbody+='    <span class="ueberschrift">MoRZ-WILMA</span>\n'
htmlbody+='    <span class="ueberschrift2">\n'
htmlbody+='        <span class="fett">W</span>ichtige \n'
htmlbody+='        <span class="fett">I</span>nformationen \n'
htmlbody+='        <span class="fett">L</span>esbar am \n'
htmlbody+='        <span class="fett">M</span>onitor \n'
htmlbody+='        <span class="fett">A</span>ngezeigt\n'
htmlbody+='    </span>\n'
htmlbody+='    <div class="zeit">Seite erzeugt am '+now.strftime('%d.%m.%Y um %H:%M')+ ' Wenn das länger als 2 Minuten her ist, stimmt etwas nicht!</div>\n'

# Den Abschluss der Seite

if iframeFlag:
        htmlfuss+='   <div class="eintrag"><div class="beschreibung">\n'
        htmlfuss+="<iframe src=\"fehler.html\" style=\"border:0px #FFFFFF none;\" name=\"errorlog\" scrolling=\"no\" frameborder=\"0\" align=aus marginheight=\"0px\" marginwidth=\"0px\" height=\"30\" width=\"640\"></iframe>"
        htmlfuss+="   </div></div>"
htmlfuss+=' </body>\n'
htmlfuss+='</html>\n'


# Kalender abfragen
try:
        c = Calendar(requests.get(urlDesKalenders).text)
        debugprint(c)
        debugprint(c.todos)
except Exception as detail:
        inhalt += createErrorHTML("Kalender abrufen", detail)
try:
        prio = 9
        # Nach Priorität geordnet ausgeben, beginnend mit der höchsten.
        while prio >= 0:
                for t in c.todos:
                        if repr(t.completed) == "None":
                                debugprint("---------------------\n  bearbeite {} \n ------------------".format(t))
                        else:
                                debugprint("completed Task {} ... skipping".format(repr(t)))
                                continue
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
                                debugprint("------------------------------------------------------------")
                                debugprint("-- {} -- Erzeuge Todo '{}', {}% erledigt, Priorität-Task {}, Priorität korrigiert: {} ".format(now, t.name, t.percent, t.priority, prio))
                                inhalt+='   <div class="'+div_klasse+'">\n'
                                inhalt+='     <div class="titel">' + t.name
                                if debugflag:
                                                #nur im debugmodus die Priorität dazu schreiben
                                                inhalt+=' (Prio:' + str(t.priority) + '/'+ str(prio) +')'
                                inhalt+='     </div>\n'
                                # Nur bei gesetzter Beschreibung, diese auch ausgeben
                                if repr(t.description) != "None":
                                        inhalt+='     <div class="beschreibung">'+t.description+'</div>\n'
                                inhalt+='   </div>\n'
                prio -= 1
except Exception as detail:
        inhalt += createErrorHTML("Todos lesen", detail)

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
                 
                 debugprint("------------------------------------------------------------")
                 debugprint("-- {} -- Erzeuge Event '{}' Startzeit: {} Endzeit {}".format(now, e.name, e.begin, e.end))
                 debugprint("Alter in Secunden: {}".format(str(abs(e.begin-now).seconds)))
                 inhalt+='   <div class="'+div+'">\n'
                 inhalt+='     <div class="titel">' + e.name + '</div>\n'
                 if repr(e.location) != "None":
                        inhalt+='     <div class="ort">'+e.location+'</div>\n'
                 if repr(e.description) != "None":
                        inhalt+='     <div class="beschreibung">'+e.description+'</div>\n'
                 inhalt+="   </div>"
except Exception as detail:
        inhalt += createErrorHTML("Events lesen", detail)


# die Entschuldtigungen
anzahlEntschuldigteSchueler=0

entschuldigungKopf ='<div class="entschuldigung">\n'
entschuldigung=""
# Modification-Time der Datei holen (Unix-Timestamp), in Datum/Zeit konvertieren und lesbar formatieren:

try:
  if excelDateiExistiert:
        with open_workbook(nameDerExceldatei, 'rb') as excelsheet:
                datemode = excelsheet.datemode
                tabelle = excelsheet.sheet_by_index(0)
                # Datum aus Exceldatei auslesen.
                xldate = (tabelle.cell(datumzelleZeile,datumzelleSpalte).value)
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

except Exception as detail: 
        entschuldigung += createErrorHTML("Entschuldigungen einbauen", detail)
        anzahlEntschuldigteSchueler = 1000
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
        datei = open(nameDerHTMLdatei,"w")
        datei.write(htmlseite)
        datei.close()
        debugprint(htmlseite)
except Exception as detail:
        print("HTML-Datei zum Schreiben öffnen", detail)
