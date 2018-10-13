# -*- coding: utf-8 -*-
debugflag = True

def fileExists(filename):
  try:
          f = open(filename)
          f.close()
          return True
  except IOError as detail:
          print(detail)
          return False
  except Exception as detail:
          print("unbehandelter Fehler in Funktion 'fileExists'", detail)
          exit(1)

def debugprint(message):
        if debugflag:
                print(message)



open("asf")
print("oberster try-block")
debugprint("Debugmodus an")
print(fileExists("testpy"))
print(fileExists("test.py"))
        