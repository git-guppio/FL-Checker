#Requires AutoHotkey v2.0


#Include GlobalConstants.ahk
#Include Utils\EventManager.ahk
#Include Utils\ProcessManager.ahk
#Include Utils\DebugManager.ahk

#Include SAP\MakeFileUpload_CTRL_ASS.ahk

result := MakeFileUpload_CTRL_ASS.MakeMapGuideline(G_CONSTANTS.file_FL_Bess) ; creo a partire dalla guideline un map che ha come chiave un pattern delle sole righe con lunghezza 4, 5, 6
if (result.success = true) { ; stampo i risultati
    myMap := result.value
    for key, data in myMap {
        ; data := {VALUE : "", SUB_VALUE : "", SUB_VALUE2 : "", TPLKZ : "", FLTYP : "", FLLEVEL : "", CODE_SEZ_PM : "", CODE_SIST : "", CODE_PARTE : "", TIPO_ELEM : ""}
        OutputDebug("key: " . key . "`n")
        OutputDebug("   data.VALUE: " . data.VALUE . "`n")
        OutputDebug("   data.SUB_VALUE: " . data.SUB_VALUE . "`n")
        OutputDebug("   data.SUB_VALUE2: " . data.SUB_VALUE2 . "`n")
        OutputDebug("   data.FLLEVEL: " . data.FLLEVEL . "`n")
        OutputDebug("   data.TPLKZ: " . data.TPLKZ . "`n")
        OutputDebug("   data.FLTYP: " . data.FLTYP . "`n")        
        OutputDebug("   data.CODE_SEZ_PM: " . data.CODE_SEZ_PM . "`n")
        OutputDebug("   data.CODE_SIST: " . data.CODE_SIST . "`n")
        OutputDebug("   data.CODE_PARTE: " . data.CODE_PARTE . "`n")                
        OutputDebug("   data.TIPO_ELEM: " . data.TIPO_ELEM . "`n")  
    }
}

fileName := "C:\Users\a259046\OneDrive - Enel Spa\SCRIPT AHK e VBA\Functional_Location\CheckFL_rev.3\Check_CTRL_ASS_Table.txt"

FL_Arr := GetFileContent(fileName)

result := MakeFileUpload_CTRL_ASS.TestGuideline(myMap, FL_Arr)
if (result.success = true) {
    WriteStringToFile(result.value, "C:\Users\a259046\OneDrive - Enel Spa\SCRIPT AHK e VBA\Functional_Location\CheckFL_rev.3\Test_CTRL_ASS_Table.txt", overwrite := true)
}

OutputDebug("-- Creato File --`n")                
              

    ; Funzione per leggere il file e estrarre i codici country e tecnologia
    ; Restituisce un map con i codici come chiave e la descrizione come valore.
    GetFileContent(fileName) {
        try {
            ; Legge il contenuto del file
            fileContent := FileRead(fileName)

            ; Divide il contenuto in linee
            lines := StrSplit(fileContent, "`n", "`r")

            ; Inizializza un array per i codici
            arr := []

            ; Estrae i codici paese
            for line in lines {
                if (line != "") {  ; Ignora le linee vuote
                    arr.Push(line) 
                }
            }
            return arr
        } catch Error as err {
            MsgBox("Errore nella lettura del file: " . FileName . " - " . err.Message, "Errore", 4112)
            return false ; in caso di errore restituisce un array vuoto
        }
    }

    ; Metodo: WriteStringToFile
    ; Descrizione: Scrive  stringa in un file di testo
    ; Parametri:
    ;   - param1: la variabile contenente la stringa che si desidera scrivere su file
    ;   - param2: il nome del file da scrivere
    ;   - param2: true -> sovrascrive il contenuto del file, false -> accoda il contenuto al file
    ; Restituisce:
    ;   - true se l'operazione di scrittura è andata a buon fine
    ;   - fasle altrimenti
    ; Esempio: WriteStringToFile(myString, "c:\pippo.txt", true)
    WriteStringToFile(content, fileName, overwrite := true) {
        try {
            if (overwrite) {
                ; Sovrascrive il contenuto se il file esiste, altrimenti crea un nuovo file
                FileObj := FileOpen(fileName, "w")
            } else {
                ; Aggiunge il contenuto alla fine del file se esiste, altrimenti crea un nuovo file
                FileObj := FileOpen(fileName, "a")
            }

            if (!FileObj) {
                throw Error("Impossibile aprire il file: " . fileName)
            }

            FileObj.Write(content)
            FileObj.Close()

            return true
        } catch as err {
            MsgBox("Errore durante la scrittura del file: " . err.Message, "Errore", 4112)
            return false
        }
    }