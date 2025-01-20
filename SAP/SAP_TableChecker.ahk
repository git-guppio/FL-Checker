; Classe: TableChecker_SAP
; Descrizione: Classe per la verifica delle FL contenute nelle tabelle globali SAP
; Esempio:

class TableChecker_SAP {

    static resultArr := []

    static __New() {
        TableChecker_SAP.SetupEventListeners()
    }

    Static SetupEventListeners() {
        EventManager.Subscribe("VerificaFL_SAP", (data) => TableChecker_SAP.VerificaFL_SAP(data.flArray, data.flcountry, data.fltechnology))
        EventManager.Subscribe("ProcessStatusUpdated",(data) => TableChecker_SAP.ProcessStatusUpdated(data.processId, data.status, data.details, data.result))            
    }

    static ProcessStatusUpdated(processId, status, details := "", result:={}) {
/*         ; result contiene il risultato del processo se in stato "Completed"
        if (processID = "CheckData") and (status = "Completed") {
            resultArr := result.value
            outputdebug("TableChecker_SAP.resultArr: " . resultArr.length . "`n")
        }
        if (processID = "CheckCountry") and (status = "Completed") {
            flCountry := result.value
            outputdebug("TableChecker_SAP.flCountry: " . flCountry . "`n")
        }
        if (processID = "CheckTechnology") and (status = "Completed") {
            flTechnology := result.value
            outputdebug("TableChecker_SAP.flTechnology: " . flTechnology . "`n")
        } */        
    }

    ; Metodo: VerificaFL_SAP
    ; Descrizione: Esegue il controllo della presenza nella tabelle SAP dei livelli delle FL
    ; Parametri:
    ;   - param1: array contenente la lista delle FL da verificare
    ; Restituisce:
    ;   - True -> se gli elementi sono già presenti in SAP
    ;   - False -> se riscontra elementi non presenti in SAP
    ; Esempio: VerificaFL_SAP(fl_array)
    Static VerificaFL_SAP(arr, flCountry, flTechnology) {
        VerificaFL_SAP_result := { success: false, value: false, error: "", class: "SAP_TableChecker.ahk", function: "VerificaFL_SAP" }
        EventManager.Publish("ProcessStarted", {processId: VerificaFL_SAP_result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica tabelle globali SAP "}) ; Avvia l'indicatore di progresso
        EventManager.Publish("ProcessProgress", {processId: VerificaFL_SAP_result.function, details: "Verifica tabelle globali SAP in corso", result: {}})
        
        check := true
        ; Definisco le variabili per creare le stringhe che andranno scritte sui file x l'upload
        content_FL_2 := G_CONSTANTS.intestazione_FL_2
        content_FL_n := G_CONSTANTS.intestazione_FL_n
        map_arr := map() ; creo una map in cui inserire come chiave in numero di elementi della FL e come valore l'ultimo elemento
        map_arr := TableChecker_SAP.MakeMapLengthLastElementFL(arr)

        ;~ ; per Debug visualizza in basse alla chiave la collezione di livelli di FL
        ;~ for chiave, valore in map_arr {
                ;~ Debug("Valore chiave = " . chiave)
                ;~ for arrayElement in valore { ; valore è un array
                    ;~ for element in arrayElement
                        ;~ Debug (arrayElement)
                ;~ }
        ;~ }

        ; verifico la presenza dei livelli contenuti come chiave nel map()
        for key, value in map_arr {
            if (map_arr[key].Length = 0) ; se non ci sono elementi considero la prossima chiave
                continue
            if (key = 2) {
                OutputDebug("Verifico livello 2 della FL" . "`n")
                OutputDebug(value[1] . "`n")
                content_FL_2_temp := TableChecker_SAP.CheckSecondoLivello_SAP(value[1], flCountry, flTechnology)
                if (content_FL_2_temp) {
                    content_FL_2 .= content_FL_2_temp
                    EventManager.Publish("AddLV", {icon: "icon2", element: " Lev.2", text: "FL non presente in SAP:"})
                    EventManager.Publish("AddLV", {icon: "icon-1", element: "" , text: "[" . value[1] . "]"})
                    if (TableChecker_SAP.WriteStringToFile(content_FL_2, G_CONSTANTS.File_FL_2_UpLoad)) {
                        EventManager.Publish("AddLV", {icon: "icon4", element: "FL_2_UpLoad.csv", text: "File aggiornato"})
                        ;TableChecker_SAP.flChecker.mainGui.UpLoadBtn.Enabled := true
                    }
                    check := false
                }
                else {
                    EventManager.Publish("AddLV", {icon: "icon1", element: " Lev.2", text: "Tabella SAP aggiornata"})
                }
            }
            else if (key > 2) {
                OutputDebug("Verifico livello " . key . " della FL" . "`n")
                content_FL_n_temp := TableChecker_SAP.CheckLivelli_SAP(key, value, flCountry, flTechnology)
                if (content_FL_n_temp) {
                    content_FL_n .= content_FL_n_temp
                    EventManager.Publish("AddLV", {icon: "icon2", element: " Lev." . key, text: "FL non presente in SAP:"})
                    EventManager.Publish("AddLV", {icon: "icon-1", element: "" , text: "[" . TableChecker_SAP.Join(TableChecker_SAP.resultArr) . "]"})                    
                    if (TableChecker_SAP.WriteStringToFile(content_FL_n, G_CONSTANTS.File_FL_n_UpLoad)) {
                        EventManager.Publish("AddLV", {icon: "icon4", element: "FL_n_UpLoad.csv", text: "File aggiornato"})
                        ;TableChecker_SAP.flChecker.mainGui.UpLoadBtn.Enabled := true
                    }
                    check := false
                }
                else {
                    EventManager.Publish("AddLV", {icon: "icon1", element: " Lev." . key, text: "Tabella SAP aggiornata"})
                }
            }
        }
        if(check)
            EventManager.Publish("ProcessCompleted", {processId: "VerificaFL_SAP", status: "Completed", details: "Esecuzione completata con successo", result: check})
        else
            EventManager.Publish("ProcessError", {processId: "VerificaFL_SAP", status: "Error", details: "FL non presente in SAP:" . " - Line Number: " . A_LineNumber, result: {}})
    }

    ; Metodo: CheckSecondoLivello_SAP
    ; Descrizione: Esegue il controllo della presenza nella tabelle SAP degli elementi appartenenti al 2 livelli delle FL
    ; Parametri:
    ;   - param1: array contenente la lista delle FL da verificare
    ; Restituisce:
    ;   - Una stringa contenente gli elementi che devono essere inseriti in SAP nella forma definita nel file x upload
    ;   - False -> altrimenti
    ; Esempio: CheckSecondoLivello_SAP(fl_array)
    Static CheckSecondoLivello_SAP(FL_2, flCountry, flTechnology) {

        if !(flCountry and flTechnology and FL_2)
            return false
        ; avvio una sessione SAP
        session := SAPConnection.GetSession()
        if (session) {
            try {
                session.findById("wnd[0]/tbar[0]/okcd").text := "/nZPM4R_UP_CONTROL_F1"
                session.findById("wnd[0]").sendVKey(0)
                Sleep 1000
                ; imposto filtro su campo <Valore del Livello>
                ;~ ;session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, "VALUE"
                ;~ ;session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleColumn := "FLLEVEL"
                grid := session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
                grid.currentCellRow := -1
                grid.selectColumn("VALUE")
                session.findById("wnd[0]/tbar[1]/btn[29]").press
                sleep 250
                session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text := FL_2 ; il secondo livello è un singolo valore
                session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition := 4
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                while session.Busy()
                {
                    sleep 500
                    OutputDebug("SAP is busy" . "`n")
                }
                ; verifico il numero di risultati ottenuti
                rowCount := grid.RowCount
                ; esco dalla transazione in corso
                session.findById("wnd[0]/tbar[0]/okcd").text := "/n"
                session.findById("wnd[0]").sendVKey(0)
                ; elimino la connessione
                SAPConnection.Disconnect()
                OutputDebug("Risultati ottenuti 2: " . rowCount . "`n")
                ; se non sono stati rilevati elementi allora tutti gli elementi della FL devono essere caricata nella tabella
                if (rowCount = 0) {
                    ; creo il una stringa che poi verrà scritta in un file
                    TPLKZ := "Z-R" . flTechnology . "M"
                    FLTYP := flTechnology
                    FLLEVEL := "2"
                    LAND1 := flCountry
                    VALUE := FL_2
                    content := TPLKZ . ";" . FLTYP . ";" . FLLEVEL  . ";" . LAND1  . ";" . VALUE . "`r`n"
                    return content
                }
                else {
                    return false
                }
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

    ; Metodo: CheckLivelli_SAP
    ; Descrizione: Esegue il controllo della presenza nella tabelle SAP degli elementi appartenenti ai livelli maggiori di 2 della FL
    ; Parametri:
    ;   - param1: livello > 2 che si desidera analizzare (da 3 a 6)
    ;   - param2: array contenente la lista delle FL da verificare
    ; Restituisce:
    ;   - Una stringa contenente gli elementi che devono essere inseriti in SAP nella forma definita nel file x upload
    ;   - False -> altrimenti
    ; Esempio: CheckLivelli_SAP(fl_array, 3)
    Static CheckLivelli_SAP(level, arr, flCountry, flTechnology) {
        TableChecker_SAP.resultArr := [] ; array contenente gli elementi non presenti in SAP
        if (!(flCountry) and !(flTechnology) and !(arr))
            return false
        ; avvio una sessione SAP
        session := SAPConnection.GetSession()
        if (session) {
            try {
                session.findById("wnd[0]/tbar[0]/okcd").text := "/nZPM4R_FL2"
                session.findById("wnd[0]").sendVKey(0)
                while session.Busy()
                {
                    sleep 500
                    OutputDebug("SAP is busy" . "`n")
                }
                ; imposto filtro sul campo <Cat. sede tecnica>
                grid := session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
                grid.currentCellRow := -1
                grid.selectColumn("FLTYP")
                session.findById("wnd[0]/tbar[1]/btn[29]").press
                sleep 500
                session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text := flTechnology
                session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition := 4
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                while session.Busy()
                {
                    sleep 500
                    OutputDebug("SAP is busy" . "`n")
                }
                ; imposto filtro sul campo <Livello sede tecnica>
                ;~ grid.setCurrentCell -1, "FLLEVEL"
                grid.currentCellRow := -1
                grid.selectColumn("FLLEVEL")
                session.findById("wnd[0]/tbar[1]/btn[29]").press
                sleep 500
                session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").text := level
                session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").caretPosition = 1
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                while session.Busy()
                {
                    sleep 500
                    OutputDebug("SAP is busy" . "`n")
                }
                ; imposto filtro sul campo <Livello sede tecnica>
                grid.currentCellRow := -1
                grid.selectColumn("VALUE")
                session.findById("wnd[0]/tbar[1]/btn[29]").press
                ; apro finestra per inserire valori multipli
                session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN003_%_APP_%-VALU_PUSH").press
                sleep 500
                session.findById("wnd[2]/tbar[0]/btn[16]").press ;bidono precedente contenuto
                ; creo una stringa a partire dall'array contenuto nel map e la inserisco nella clipboard
                OutputDebug("Valori presenti nel livello " . level . ": " . arr.Length . "`n")
                TableChecker_SAP.ArrayToClipboard(arr) ; arr contiene tutti gli elementi presenti nel livello analizzato
                sleep 500
                session.findById("wnd[2]/tbar[0]/btn[24]").press ; incollo dalla clipboard
                sleep 500
                session.findById("wnd[2]/tbar[0]/btn[8]").press ; premo OK - chiudo finestra e torno alla schermata precedente
                sleep 500
                session.findById("wnd[1]/tbar[0]/btn[0]").press ; applico filtri
                ; ripristino il contenuto precedente della clipboard
                A_Clipboard := TableChecker_SAP.clipBoard
                while session.Busy()
                {
                    sleep 500
                    OutputDebug("SAP is busy" . "`n")
                }
                ; - verifico il numero di elementi trovati:
                ; 		- se è uguale al numero di elementi nell'array arr, allora non deve essere inserito nessun elemento in SAP (sono già presenti)
                ; 		- se non è uguale allora devo costruire un nuovo array contenente i valori che sono presenti in arr ma non sono presenti nei risultati
                ; 		- se è uguale a zero allora devono essere inseriti tutti gli elementi
                ; Determina il numero totale di righe nella tabella
                rowCount := grid.RowCount
                TableChecker_SAP.resultArr := []
                if (rowCount > 0) {
                    tabellaSAParray := [] ; tabella che contiene i dati già presenti nella tabella SAP del relativo livello
                    OutputDebug("Rilevo dati in tabella SAP - livello: " . level . "`n")
                    ; Determina il numero di righe visibili
                    visibleRows := grid.VisibleRowCount
                    ; Leggi tutte le righe, scorrendo quando necessario
                    currentRow := 0
                    while (currentRow < rowCount) {
                        ; Leggi le righe visibili
                        Loop visibleRows {
                            if (currentRow >= rowCount) {
                                break
                            }
                            try {
                                cellValue := grid.GetCellValue(currentRow, "VALUE")
                                OutputDebug("Riga: " . A_Index . " - valore: " . cellValue . " - " . currentRow . "`n")
                                if (cellValue != "") {
                                    if !TableChecker_SAP.HasValue(tabellaSAParray, cellValue) {
                                        tabellaSAParray.Push(cellValue)
                                    }
                                }
                            } catch as err {
                                MsgBox("Errore nel leggere la riga " . currentRow . ": " . err.Message, "Errore", 4112)
                                return false
                            }
                            currentRow++
                        }

                        ; Scorri alla prossima pagina di righe
                        if (currentRow < rowCount) {
                            grid.FirstVisibleRow := currentRow
                            Sleep(100)  ; Breve pausa per permettere il caricamento
                        }
                    }
                    ; esco dalla transazione in corso
                    session.findById("wnd[0]/tbar[0]/okcd").text := "/n"
                    session.findById("wnd[0]").sendVKey(0)
                    ; elimino la connessione
                    SAPConnection.Disconnect()
                    OutputDebug("Valori del livello " . level . " in SAP: " . tabellaSAParray.Length . "`n")
                    TableChecker_SAP.resultArr := TableChecker_SAP.ArrayDifference(arr, tabellaSAParray) ; creo un nuovo array contenente gli elementi presenti nel primo array ma non nel secondo
                    OutputDebug("Valori del livello: " . level . " da inserire in SAP: " . TableChecker_SAP.resultArr.length . "`n")
                }
                else if (rowCount = 0) { ; se il risultato del filtro non contiene elementi allora tutti gli elementi devono essere inseriti
                    TableChecker_SAP.resultArr := arr
                }
                if (TableChecker_SAP.resultArr.Length > 0) { ; se l'array contiene elementi
                    for element in TableChecker_SAP.resultArr {
                        TPLKZ := "Z-R" . flTechnology . "S"
                        FLTYP := (flTechnology = "H") ? "L" : flTechnology
                        FLLEVEL := level
                        VALUE := element
                        VALUETX := ""
                        REFLEVEL := ""
                        content .= TPLKZ . ";" . FLTYP . ";" . FLLEVEL  . ";" . VALUE  . ";" . VALUETX  . ";" . REFLEVEL . "`r`n"
                    }
                    OutputDebug("Content: " . content . "`n")
                    return content
                }
                else
                    return false

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

    ; Metodo: ArrayDifference
    ; Descrizione: Effettua la differenza tra il contenuto di due array
    ; Parametri:
    ;   - param1: array contenente un insieme di elementi
    ;   - param2: array contenente un insieme di elementi (presumibilmente sottoinsieme del primo)
    ; Restituisce:
    ; - un array contenente i valori presenti nel primo array meno gli elementi presenti nel secondo (che è un sottoinsieme del primo)
    ; - un array privo di elementi
    ; Esempio: ArrayDifference(fl_array, array)
    Static ArrayDifference(array1, array2) {
        result := []
        for _, item in array1 {
            if !TableChecker_SAP.HasValue(array2, item) {
                result.Push(item)
            }
        }
        return result
    }

    ; Metodo: ArrayToClipboard
    ; Descrizione: Scrive nella clipboard gli elementi contenuti nell'array
    ; Parametri:
    ;   - param1: array contenente la lista delle FL da copiare nella clipboard
    ; Restituisce:
    ;   - true se la clipboard è stata scritta
    ;   - fasle altrimenti
    ; Esempio: ArrayToClipboard(fl_array)
    Static ArrayToClipboard(arr) {
        TableChecker_SAP.clipBoard := A_Clipboard
        A_Clipboard := ""
        sleep 250
        ; Usa Join() per creare una stringa con elementi separati da newline
        result := TableChecker_SAP.Join(arr, "`r`n")
        ; Copia la stringa risultante nella clipboard
        A_Clipboard := result
        sleep 250
        ; Opzionale: ritorna la stringa risultante
        return A_Clipboard != "" ? true : false
    }

    ; Metodo: Join
    ; Descrizione: Crea una stringa contenente gli elementi dell'array separati dal delimitatore
    ; Parametri:
    ;   - param1: array da cui si vuole realizzare la stringa
    ;   - param2: delimitatore da inserire tra gli elementi dell'array
    ; Restituisce:
    ;   - una stringa conenente gli elementi dell'array separati dal delimitatore
    ; Esempio: Join(array, "-")
    Static Join(arr, delimiter := ", ") {
        result := ""
        for index, element in arr {
            if (index > 1)
                result .= delimiter
            result .= element
        }
        return result
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
    Static WriteStringToFile(content, fileName, overwrite := true) {
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

    ; Metodo: MakeMapLengthLastElementFL
    ; Descrizione: Crea una struttura di tipo map() con chiave il numero del livello della FL e come valore un array con tutti gli elementi di quel livello
    ; Parametri:
    ;   - param1: un array contenente la lista delle FL
    ; Restituisce:
    ;   - una struttura dati di tipo map()
    ;   - false se la struttura non contiene elementi
    ; Esempio: MakeMapLengthLastElementFL(array_FL)
    Static MakeMapLengthLastElementFL(arr) {
        map_array_FL := map()
        for element in arr {
            temp_array := StrSplit(element, "-")
            numberOfElement := temp_array.Length
            ; creo una struttura per inserire l'ultimo elemento della FL e collezionarli in base alla lunghezza della FL
            ; Se questa chiave non esiste ancora nella Map la crea e inserisco come valore un array
            if !map_array_FL.Has(numberOfElement) {
                map_array_FL[numberOfElement] := []
            }
            ; Aggiungi l'ultimo elemento della FL solo se non è già presente nell'array
            if !(TableChecker_SAP.HasValue(map_array_FL[numberOfElement], temp_array[numberOfElement]))
                map_array_FL[numberOfElement].Push(temp_array[numberOfElement])
        }
        ; verifico che siano stati inseriti dei valori
        if (map_array_FL.Count > 0)
            return map_array_FL
        else
            return false
    }

    ; Metodo: HasValue
    ; Descrizione: Verifica la presenza di un valore all'interno di un array
    ; Parametri:
    ;   - param1: un array
    ;   - param2: il valore da ricercare all'interno dell'array
    ; Restituisce:
    ;   - true se l'elemento è contenuto nell'array
    ;   - false altrimenti
    ; Esempio: HasValue(arr, "pippo")
    Static HasValue(arr, strValue) {
        for _, value in arr {
            if (value = strValue)
                return true
        }
        return false
    }
}