; Programma principale
; Descrizione: Utilizza l'event manager per coordinare le attività tra i diversi componenti del programma, inizializza e configura tutti i componenti dell'applicazione.
; Esempio:

#Requires AutoHotkey v2.0
#SingleInstance Force

#Include GlobalConstants.ahk
#Include Utils\EventManager.ahk
#Include Utils\ProcessManager.ahk
#Include Utils\DebugManager.ahk
#Include FLChecker.ahk
#Include GUI\MakeGUI.ahk
#Include SAP\SAP_Connection.ahk
#Include SAP\SAP_TableChecker.ahk
#Include SAP\SAP_UpLoadTable.ahk
#Include SAP\SAP_ControlAsset.ahk
#Include SAP\MakeFileUpload_CTRL_ASS.ahk
#Include SAP\SAP_TechObj.ahk
#Include SAP\MakeFileUpload_TECH_OBJ.ahk

Main()

Main() {
    ; Avvio dell'applicazione
    EventManager.Publish("ProcessStarted", {processId: "MainGUI", status: "Started", details: "", result: {}})
    app := MainApp()
    app.ShowMain
}