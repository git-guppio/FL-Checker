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

    ; Metodo: UpLoadFiles
    ; Descrizione: Esegue l'UpLoad in SAP dei file csv presenti nella cartella <FileUpLoad>
    ; Parametri:
    ; Restituisce:
    ;   - true se l'operazione va a buon fine
    ;   - false altrimenti
    ; Esempio: UpLoadSecondoLivello_SAP("USS8")
    static UpLoadFiles() {
        ; variabili per memorizzare file di cui si esegue l'upload e l'esito del caricamento.
        Map_FileUploaded := map()
        esito := false

        UpLoadFiles_result := { success: false, value: false, error: "", class: "SAP_UpLoadTable.ahk", function: "UpLoadFiles" }
        EventManager.Publish("ProcessStarted", {processId: UpLoadFiles_result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Caricamento tabelle globali SAP "}) ; Avvia l'indicatore di progresso
        
        if (FileExist(G_CONSTANTS.file_FL_2_UpLoad)) { ; se il file esiste -> lo carico
            NameFile_Path := SAP_UpLoadTable.SeparateFilePathAndName(G_CONSTANTS.file_FL_2_UpLoad)          
            MsgBoxResult := MsgBox("Caricare il file " . NameFile_Path.fileName . "? (press Si or No)","UpLoad Global SAP table", 4132)
            if (MsgBoxResult = "Yes") {            
                EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: "Caricamento file 2° livello", result: {}})
                ; ---> carico file FL_2_UpLoad.csv

                if (SAP_UpLoadTable.UpLoadLivello_2_SAP()) {
                    msgInfo := "File di livello 2 caricato con successo"
                    EventManager.Publish("AddLV", {icon: "icon1", element: NameFile_Path.fileName, text: msgInfo})
                    EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: msgInfo, result: {}})
                    esito := true
                } else {
                    msgError := "Errore nel caricamento del file di livello 2"
                    EventManager.Publish("AddLV", {icon: "icon3", element: NameFile_Path.fileName, text: msgError})
                    EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: msgError, result: {}})
                    EventManager.Publish("PI_Stop", {inputValue: msgError}) ; Ferma l'indicatore di progresso
                    esito := false
                }
                Map_FileUploaded[NameFile_Path.fileName] := esito
            }
            else {
            
            }
        }

        if (FileExist(G_CONSTANTS.file_FL_n_UpLoad)) { ; se il file esiste -> lo carico
            NameFile_Path := SAP_UpLoadTable.SeparateFilePathAndName(G_CONSTANTS.file_FL_n_UpLoad)          
            MsgBoxResult := MsgBox("Caricare il file " . NameFile_Path.fileName . "? (press Si or No)","UpLoad Global SAP table", 4132)
            if (MsgBoxResult = "Yes") {          
                EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: "Caricamento file n° livello", result: {}})
                ; ---> carico file FL_n_UpLoad.csv
                if (SAP_UpLoadTable.UpLoadLivello_n_SAP()) {
                    msgInfo := "File di livello n caricato con successo"
                    EventManager.Publish("AddLV", {icon: "icon1", element: NameFile_Path.fileName, text: msgInfo})
                    EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: msgInfo, result: {}})
                    esito := true
                } else {
                    msgError := "Errore nel caricamento del file di livello n"
                    EventManager.Publish("AddLV", {icon: "icon3", element: NameFile_Path.fileName, text: msgError})
                    EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: msgError, result: {}})
                    EventManager.Publish("PI_Stop", {inputValue: msgError}) ; Ferma l'indicatore di progresso
                    esito := false
                }
                Map_FileUploaded[NameFile_Path.fileName] := esito
            }
            else {
            
            }                
        }

        if (FileExist(G_CONSTANTS.file_ZPMR_CTRL_ASS_UpLoad)) { ; se il file esiste -> lo carico
            NameFile_Path := SAP_UpLoadTable.SeparateFilePathAndName(G_CONSTANTS.file_ZPMR_CTRL_ASS_UpLoad)
            MsgBoxResult := MsgBox("Caricare il file " . NameFile_Path.fileName . "? (press Si or No)","UpLoad SAP CTRL ASS table", 4132)
            if (MsgBoxResult = "Yes") {
                EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: "Caricamento file CTRL_ASS", result: {}})                
                ; ---> carico file FL_n_UpLoad.csv
                if (SAP_UpLoadTable.UpLoadCTRL_ASS()) {
                    msgInfo := "File CTRL_ASS caricato con successo"
                    EventManager.Publish("AddLV", {icon: "icon1", element: NameFile_Path.fileName, text: msgInfo})
                    EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: msgInfo, result: {}})
                    esito := true
                } else {
                    msgError := "Errore nel caricamento del file CTRL_ASS"
                    EventManager.Publish("AddLV", {icon: "icon3", element: NameFile_Path.fileName, text: msgError})
                    EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: msgError, result: {}})
                    EventManager.Publish("PI_Stop", {inputValue: msgError}) ; Ferma l'indicatore di progresso
                    esito := false
                }
                Map_FileUploaded[NameFile_Path.fileName] := esito
            }
            else {
            
            }                
        }

        if (FileExist(G_CONSTANTS.file_ZPMR_TECH_OBJ_UpLoad)) {  ; se il file esiste -> lo carico
            NameFile_Path := SAP_UpLoadTable.SeparateFilePathAndName(G_CONSTANTS.file_ZPMR_TECH_OBJ_UpLoad)
            MsgBoxResult := MsgBox("Caricare il file " . NameFile_Path.fileName . "? (press Si or No)","UpLoad SAP TECH OBJ table", 4132)
            if (MsgBoxResult = "Yes") {
                ; ---> carico file FL_n_UpLoad.csv
                if (SAP_UpLoadTable.UpLoadTECH_OBJ()) {
                    msgInfo := "File TECH_OBJ caricato con successo"
                    EventManager.Publish("AddLV", {icon: "icon1", element: NameFile_Path.fileName, text: msgInfo})
                    EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: msgInfo, result: {}})
                    esito := true
                } else {
                    msgError := "Errore nel caricamento del file TECH_OBJ"
                    EventManager.Publish("AddLV", {icon: "icon3", element: NameFile_Path.fileName, text: msgError})
                    EventManager.Publish("ProcessProgress", {processId: UpLoadFiles_result.function, status: "In Progress", details: msgError, result: {}})
                    EventManager.Publish("PI_Stop", {inputValue: msgError}) ; Ferma l'indicatore di progresso
                    esito := false
                }
                Map_FileUploaded[NameFile_Path.fileName] := esito
            }
            else {
            
            }            
        }

        ; al termine del caricamento dei file         
        countFileOK := 0
        countFileNOK := 0
        for file, esito in Map_FileUploaded {
                msgInfo := "`t - " . file . " -> " . (esito=true ? " OK" : " NOK") . "`n"
                esito=true ? countFileOK++ : countFileNOK++
        }
        totFile:=countFileOK+countFileNOK
        UpLoadFiles_result.value := Map_FileUploaded
        OutputDebug("UpLoadFiles_result - Esito upload dei file: `n" . msgInfo . "`n")
        if (countFileOK > 0) {
            UpLoadFiles_result.success := true
            msgInfo := "Effettutato caricamento di " . countFileOK . "su" . totFile . (totFile>1 ? "files":"file")
            EventManager.Publish("ProcessCompleted", {processId: UpLoadFiles_result.function, status: "Completed" , details: msgInfo, result: UpLoadFiles_result})
            EventManager.Publish("PI_Stop", {inputValue: "Caricamento file in SAP concluso. Esito= " . countFileOK . "/" . totFile}) ; Ferma l'indicatore di progresso
        }
        else if (countFileOK = 0) {
            msgInfo := "Errore nel caricamento dei file in SAP - Nessun file caricato"
            EventManager.Publish("ProcessError", {processId: UpLoadFiles_result.function, status: "Error" , details: msgInfo, result: UpLoadFiles_result})
            EventManager.Publish("PI_Stop", {inputValue: msgInfo}) ; Ferma l'indicatore di progresso            
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
            SAPConnection.Disconnect()
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
                SAPConnection.Disconnect()
            }
            else {
                MsgBox("Impossibile ottenere una sessione SAP valida.", "Errore", 4112)
                return false
            }
    }

    ; Metodo: UpLoadCTRL_ASS
    ; Descrizione: Esegue l'UpLoad in SAP del file per l'aggiornamento della tabella CTRL_ASS
    ; Parametri:
    ; Restituisce:
    ;   - true se l'operazione va a buon fine
    ;   - false altrimenti
    ; Esempio: UpLoadSecondoLivello_SAP("USS8")
    Static UpLoadCTRL_ASS() {
        result := this.SeparateFilePathAndName(G_CONSTANTS.file_ZPMR_CTRL_ASS_UpLoad)
        OutputDebug("Eseguo UpLoad file: " . result.folderPath . " - " result.fileName . " - " . result.isCSV . "`n")
        if !((result.isCSV) and (result.folderPath) and (result.fileName)) ; se il file ha estensione .csv ha un nome e un percorso allora procedo
            return false
            ; avvio una sessione SAP
            session := SAPConnection.GetSession()
            if (session) {
                try {
                    session.findById("wnd[0]/tbar[0]/okcd").text := "/nZPM4R_UPL_FL_FILE"
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]/usr/radR_BUT3").setFocus
                    session.findById("wnd[0]/usr/radR_BUT3").select
                    session.findById("wnd[0]/usr/chkP_INT").selected := true
                    session.findById("wnd[0]/usr/ctxtP_FILE").setFocus
                    session.findById("wnd[0]/usr/ctxtP_FILE").caretPosition := 0
                    session.findById("wnd[0]").sendVKey(4)
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text := result.folderPath
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text := result.fileName
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition := 17
                    session.findById("wnd[1]/tbar[0]/btn[0]").press
                    sleep 500
                    ; OK carica file                    
                    session.findById("wnd[0]/tbar[1]/btn[8]").press
                    ;session.findById("wnd[0]/tbar[0]/btn[15]").press                    
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
                SAPConnection.Disconnect()
            }
            else {
                MsgBox("Impossibile ottenere una sessione SAP valida.", "Errore", 4112)
                return false
            }
    }     

    ; Metodo: UpLoadTECH_OBJ
    ; Descrizione: Esegue l'UpLoad in SAP del file per l'aggiornamento della tabella TECH_OBJ
    ; Parametri:
    ; Restituisce:
    ;   - true se l'operazione va a buon fine
    ;   - false altrimenti
    ; Esempio: UpLoadSecondoLivello_SAP("USS8")
    Static UpLoadTECH_OBJ() {
        result := this.SeparateFilePathAndName(G_CONSTANTS.file_ZPMR_TECH_OBJ_UpLoad)
        OutputDebug("Eseguo UpLoad file: " . result.folderPath . " - " result.fileName . " - " . result.isCSV . "`n")
        if !((result.isCSV) and (result.folderPath) and (result.fileName)) ; se il file ha estensione .csv ha un nome e un percorso allora procedo
            return false
            ; avvio una sessione SAP
            session := SAPConnection.GetSession()
            if (session) {
                try {
                    session.findById("wnd[0]/tbar[0]/okcd").text := "/nZPM4R_UPL_FL_FILE"
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]/usr/radR_BUT4").setFocus
                    session.findById("wnd[0]/usr/radR_BUT4").select
                    session.findById("wnd[0]/usr/chkP_INT").selected := true
                    session.findById("wnd[0]/usr/ctxtP_FILE").setFocus
                    session.findById("wnd[0]/usr/ctxtP_FILE").caretPosition := 0
                    session.findById("wnd[0]").sendVKey(4)
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text := result.folderPath
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text := result.fileName
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition := 17
                    session.findById("wnd[1]/tbar[0]/btn[0]").press
                    sleep 500
                    ; OK carica file
                    session.findById("wnd[0]/tbar[1]/btn[8]").press
                    ;session.findById("wnd[0]/tbar[0]/btn[15]").press
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
                SAPConnection.Disconnect()
            }
            else {
                MsgBox("Impossibile ottenere una sessione SAP valida.", "Errore", 4112)
                return false
            }
    }    
}    