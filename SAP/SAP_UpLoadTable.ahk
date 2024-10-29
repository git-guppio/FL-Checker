; Classe: SAP_UpLoadTable
; Descrizione: Classe per il caricamento dei livelli delle FL non presenti nelle tabelle globali SAP
; Esempio:

class SAP_UpLoadTable {

    static __New() {
        SAP_UpLoadTable.InitializeVariables()
        SAP_UpLoadTable.SetupEventListeners()
    }

    Static InitializeVariables() {

    }

    Static SetupEventListeners() {
        EventManager.Subscribe("UpLoadFiles", (*) => SAP_UpLoadTable.UpLoadFiles())
    }

    ; Metodo: UpLoadLivello_2_SAP
    ; Descrizione: Esegue l'UpLoad in SAP del valore del secondo livello nella tabella globale
    ; Parametri:
    ; Restituisce:
    ;   - true se l'operazione va a buon fine
    ;   - false altrimenti
    ; Esempio: UpLoadSecondoLivello_SAP("USS8")
    static UpLoadFiles() {
        UpLoadFiles_result := { success: false, value: false, error: "", class: "SAP_UpLoadTable.ahk", function: "UpLoadFiles" }
        EventManager.Publish("ProcessStarted", {processId: UpLoadFiles_result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Caricamento tabelle globali SAP "}) ; Avvia l'indicatore di progresso
        UpLoadFiles_result.value := 0
        if (FileExist(G_CONSTANTS.file_FL_2_UpLoad)) { ; se il file non è stato creato allora non devo caricarlo
            EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: "Caricamento file 2° livello", result: {}})
            if (SAP_UpLoadTable.UpLoadLivello_2_SAP()) {
                EventManager.Publish("AddLV", {icon: "icon1", element: "FL_2_UpLoad.csv", text: "File di livello 2 caricato con successo"})
                EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: "Caricamento file 2° livello - OK", result: {}})
                UpLoadFiles_result.value := 1
            } else {
                EventManager.Publish("AddLV", {icon: "icon3", element: "FL_2_UpLoad.csv", text: "Errore nel caricamento del file di livello 2"})
                EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: "Errore nel caricamento del file di livello 2", result: {}})
                EventManager.Publish("PI_Stop", {inputValue: "Errore nel caricamento tabelle globali SAP"}) ; Ferma l'indicatore di progresso
            }
        }

        if !(FileExist(G_CONSTANTS.file_FL_n_UpLoad)) ; se il file non è stato creato allora non devo caricarlo
            return false
        else {
            if (SAP_UpLoadTable.UpLoadLivello_n_SAP()) {
                EventManager.Publish("AddLV", {icon: "icon1", element: "FL_n_UpLoad.csv", text: "File di livello 2 caricato con successo"})
                EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: "Caricamento file n livello - OK", result: {}})
                UpLoadFiles_result.value += 1
            } else {
                EventManager.Publish("AddLV", {icon: "icon3", element: "FL_n_UpLoad.csv", text: "Errore nel caricamento del file di livello n"})
                EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: "Errore nel caricamento del file di livello n", result: {}})
                EventManager.Publish("PI_Stop", {inputValue: "Errore nel caricamento tabelle globali SAP"}) ; Ferma l'indicatore di progresso
            }
        }
        OutputDebug("UpLoadFiles_result.value" . UpLoadFiles_result.value . "`n")
        if (UpLoadFiles_result.value > 0) {
            UpLoadFiles_result.success := true
            details_txt := "Effettutato caricamento di " . UpLoadFiles_result.value . (UpLoadFiles_result.value = 1 ? "file" : "files")
            EventManager.Publish("ProcessCompleted", {processId: UpLoadFiles_result.function, status: "Completed" , details: details_txt, result: UpLoadFiles_result})
            EventManager.Publish("PI_Stop", {inputValue: "Caricamento tabelle globali SAP - OK"}) ; Ferma l'indicatore di progresso
        }
        else if (UpLoadFiles_result.value = 0) {
            EventManager.Publish("ProcessError", {processId: UpLoadFiles_result.function, status: "Error", details: "Errore caricamento dei file", result: UpLoadFiles_result})
            EventManager.Publish("PI_Stop", {inputValue: "Errore caricamento tabelle globali SAP"}) ; Ferma l'indicatore di progresso            
        }
    }    

    ; Metodo: SeparateFilePathAndName
    ; Descrizione: Esegue il controllo della presenza nella tabelle SAP degli elementi appartenenti al 2 livelli delle FL
    ; Parametri:
    ;   - param1: stringa contenente il path ed il nome del file
    ; Restituisce un oggetto composto da:
    ;   - folderPath: una stringa contenente il path del file
    ;   - fileName: una stringa contenente il nome del file compresa la sua estensione
    ;   - isCSV: un valore booleano, true -> se il file ha estensione .csv, false -> altrimenti
    ; Esempio: SeparateFilePathAndName("C:\Users\a259046\OneDrive - Enel Spa\SCRIPT AHK e VBA\Functional_Location\CheckFL\FileUpLoad\FL_2_UpLoad.csv")
    Static SeparateFilePathAndName(fullPath) {
        ; Trova l'ultima occorrenza del carattere "\"
        lastBackslashPos := InStr(fullPath, "\", , -1)

        if (lastBackslashPos > 0) {
            ; Estrae il percorso della cartella
            folderPath := SubStr(fullPath, 1, lastBackslashPos - 1)

            ; Estrae il nome del file
            fileName := SubStr(fullPath, lastBackslashPos + 1)

            ; Verifica l'estensione del file
            fileExtension := SubStr(fileName, -4)
            isCSV := (StrLower(fileExtension) = ".csv")

            ; Restituisce un oggetto con percorso della cartella, nome del file e flag CSV
            return {folderPath: folderPath, fileName: fileName, isCSV: isCSV}
        } else {
            ; Se non trova "\", assume che sia solo un nome file
            fileExtension := SubStr(fullPath, -4)
            isCSV := (StrLower(fileExtension) = ".csv")
            return {folderPath: "", fileName: fullPath, isCSV: isCSV}
        }
    }

    ; Metodo: UpLoadLivello_2_SAP
    ; Descrizione: Esegue l'UpLoad in SAP del valore del secondo livello nella tabella globale
    ; Parametri:
    ; Restituisce:
    ;   - true se l'operazione va a buon fine
    ;   - false altrimenti
    ; Esempio: UpLoadSecondoLivello_SAP("USS8")
    Static UpLoadLivello_2_SAP() {
        result := this.SeparateFilePathAndName(G_CONSTANTS.file_FL_2_UpLoad)
        OutputDebug("Eseguo UpLoad file: " . result.folderPath . " - " result.fileName . " - " . result.isCSV . "`n")
        if !((result.isCSV) and (result.folderPath) and (result.fileName)) ; se il file ha estensione .csv ha un nome e un percorso allora procedo
            return false
        ; avvio una sessione SAP
        session := SAPConnection.GetSession()
        if (session) {
            try {
                session.findById("wnd[0]/tbar[0]/okcd").text := "/nZPM4R_UPL_FL_FILE"
                session.findById("wnd[0]").sendVKey(0)
                while session.Busy()
                {
                    sleep 500
                    OutputDebug("SAP is busy" . "`n")
                }
                ; seleziono il bottone <Tabella per 1 e 2 livello>
                session.findById("wnd[0]/usr/radR_BUT1").select
                ; seleziono il radio button <Con intestazione?>
                session.findById("wnd[0]/usr/chkP_INT").selected := true
                ; apro finestra dialogo per selezione file
                ;~ session.findById("wnd[0]/usr/ctxtP_FILE").setFocus
                session.findById("wnd[0]/usr/ctxtP_FILE").caretPosition := 0
                session.findById("wnd[0]").sendVKey(4)
                sleep 250
                ; imposto path e nome file
                session.findById("wnd[1]/usr/ctxtDY_PATH").text := result.folderPath
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text := result.fileName
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition := 15
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                ; eseguo upload del file
                session.findById("wnd[0]/tbar[1]/btn[8]").press
                while session.Busy()
                {
                    sleep 500
                    OutputDebug("SAP is busy" . "`n")
                }
                return true
            } catch as err {
                MsgBox("Errore nell'esecuzione dell'azione SAP: " err.Message, "Errore", 4112)
                return false
            }
        }
        else {
            MsgBox("Impossibile ottenere una sessione SAP valida.", "Errore", 4112)
            return false
        }
    }

    ; Metodo: UpLoadLivello_n_SAP
    ; Descrizione: Esegue l'UpLoad in SAP del valore dei livelli 3,4,5 e 6
    ; Parametri:
    ; Restituisce:
    ;   - true se l'operazione va a buon fine
    ;   - false altrimenti
    ; Esempio: UpLoadSecondoLivello_SAP("USS8")
    Static UpLoadLivello_n_SAP() {
        result := this.SeparateFilePathAndName(G_CONSTANTS.file_FL_n_UpLoad)
        OutputDebug("Eseguo UpLoad file: " . result.folderPath . " - " result.fileName . " - " . result.isCSV . "`n")
        if !((result.isCSV) and (result.folderPath) and (result.fileName)) ; se il file ha estensione .csv ha un nome e un percorso allora procedo
            return false
            ; avvio una sessione SAP
            session := SAPConnection.GetSession()
            if (session) {
                try {
                    session.findById("wnd[0]/tbar[0]/okcd").text := "/nZPM4R_UPL_FL_FILE"
                    session.findById("wnd[0]").sendVKey(0)
                    while session.Busy()
                    {
                        sleep 500
                        OutputDebug("SAP is busy" . "`n")
                    }
                    ; seleziono il bottone <Tabella per 3,4,5 e 6 livello
                    session.findById("wnd[0]/usr/radR_BUT2").select
                    ; seleziono il radio button <Con intestazione?>
                    session.findById("wnd[0]/usr/chkP_INT").selected := true
                    ; apro finestra dialogo per selezione file
                    ;~ session.findById("wnd[0]/usr/ctxtP_FILE").setFocus
                    session.findById("wnd[0]/usr/ctxtP_FILE").caretPosition := 0
                    session.findById("wnd[0]").sendVKey(4)
                    ; imposto path e nome file
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text := result.folderPath
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text := result.fileName
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition := 15
                    session.findById("wnd[1]/tbar[0]/btn[0]").press
                    ; eseguo upload del file
                    session.findById("wnd[0]/tbar[1]/btn[8]").press
                    while session.Busy()
                    {
                        sleep 500
                        OutputDebug("SAP is busy" . "`n")
                    }
                    return true
                } catch as err {
                    MsgBox("Errore nell'esecuzione dell'azione SAP: " err.Message, "Errore", 4112)
                    return false
                }
            }
            else {
                MsgBox("Impossibile ottenere una sessione SAP valida.", "Errore", 4112)
                return false
            }
    }
}    