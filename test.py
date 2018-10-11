debugflag = True

def fileExists(filename):
  try:
          f = open(filename)
          f.close()
          return True
  except FileNotFoundError:
          return False
  except Exception as detail:
          print("unbehandelter Fehler in Funktion 'fileExists'", detail)
          exit(1)

def debugprint(message):
        if debugflag:
                print(message)
debugprint("Debugmodus an")
print(fileExists("testpy"))