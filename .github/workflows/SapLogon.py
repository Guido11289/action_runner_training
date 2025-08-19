import subprocess
import win32com.client
import time

@keyword('saplogin')
def saplogin(path):
    # Start SAP Logon
    proces = subprocess.Popen(path, shell=True, stdout=subprocess.PIPE, stdin=subprocess.PIPE)

    time.sleep(5)
    # subprocess.check_call([path], shell=True)
    # Wacht tot GUI geladen is
    SapGuiAuto = None
    for i in range(30):  # max 30 seconden wachten
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            break
        except:
            time.sleep(1)

    if not SapGuiAuto:
        raise Exception("SAP GUI scripting niet beschikbaar")

    application = SapGuiAuto.GetScriptingEngine
    if not application:
        raise Exception("Kon ScriptingEngine niet ophalen")

    return application


if __name__=='__main__':
  saplogin('C:\\Program Files\\SAP\\FrontEnd\\SAPGUI\\saplogon.exe')
