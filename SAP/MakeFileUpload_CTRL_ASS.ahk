#Requires AutoHotkey v2.0

class MakeFileUpload_CTRL_ASS {

    static __New() {
        MakeFileUpload_CTRL_ASS.InitializeVariables()
        MakeFileUpload_CTRL_ASS.SetupEventListeners()
    }

    Static InitializeVariables() {
        
    }

    Static SetupEventListeners() {
        EventManager.Subscribe("MakeFile_CTRL_ASS", (data) => MakeFileUpload_CTRL_ASS.MakeFile_CTRL_ASS(data.flArray, data.flTechnology, data.flInvType))
    }

; ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; * Realizzazione dei file per l'aggiornamento della Control Table Asset
;   Deve essere creato il file:
;   -   nomeFile_ZPMR_CTRL_ASS_UpLoad
;   
;   Da caricare attraverso la transazione: ZPM4R_UPL_FL_FILE
; ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    ; Funzione per costruire file per aggioranmento della tabella CTRL_ASS
    ; Viene confrontato ogni elemento dell'array prodotto dal metodo <VerificaControlAsset> con le guideline per il recupero delle informazioni necessarie alla costruzione dei file 
    ; da caricare a sistema.   
    ;~ Parametri:
    ;~ - un array contenente la lista degli elementi non presenti nella CTRL_ASS prodotta dalla verifica precedente
    ;~ Restituisce:
    ;~ - Due file pronti per il caricamento a sistema
    ;~ - False altrimenti

    static MakeFile_CTRL_ASS(arr, tech, invType:="") {
        MakeFileCTRL_ASS_result := { success: false, value: false, error: "", class: "MakeFileUpload_CTRL_ASS.ahk", function: "MakeFileCTRL_ASS" }
        EventManager.Publish("ProcessStarted", {processId: MakeFileCTRL_ASS_result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Crea file aggiornamento CTRL_ASS "}) ; Avvia l'indicatore di progresso 
            if (arr.Length = 0) {
                msgbox("Errore l'array indicato è vuoto!", "Errore", 4112)
                MakeFileCTRL_ASS_result.error := "Errore l'array indicato è vuoto!"
                EventManager.Publish("ProcessError", {processId: MakeFileCTRL_ASS_result.function, status: "Error", details: MakeFileCTRL_ASS_result.error, result: {}})
                return MakeFileCTRL_ASS_result  ; Se l'array è vuoto allora restituisco un errore
            }  
            arrFileGuideLine := []  ; creo un array per memorizzare i file da cui creare i pattern per verificare gli elementi presenti nell'array delle FL 
            MakeFileCTRL_ASS_result.success := true ; imposto a true e cambio in caso venga riscontrato almeno un errore

            ; Creo n array contenente tutti i file guideline coinvolti nella generazione delle FL,
            ; da utilizzare in seguito per creare un unico map() che ha per chiave il codice FL della guidelinee come valore un oggetto composto dagli element:
            ; {VALUE: "", SUB_VALUE: "", SUB_VALUE2: "", TPLKZ: "", FLTYP: "", FLLEVEL: "", CODE_SEZ_PM: "", CODE_SIST: "", CODE_PARTE: "", TIPO_ELEM: ""}
            try {
                EventManager.Publish("ProcessProgress", {processId: "MakeFile_CTRL_ASS", status: "In Progress", details: "Genero array contenente file guideline", result: {}})     
                if (tech = "S") { ; verifico gli impianti con tecnologia SOLAR
                    
                    arrFileGuideLine.Push(G_CONSTANTS.file_FL_Solar_Common)
                    arrFileGuideLine.Push(G_CONSTANTS.file_FL_S_SubStation)

                    ; in base alla tipologia di inverter creo il relativo array
                    switch invType
                    {
                        case 1:
                            arrFileGuideLine.Push(G_CONSTANTS.file_FL_Solar_CentralInv)
                            EventManager.Publish("DebugMsg",{msg: "MakeFile_CTRL_ASS: Selezionato Central Inverter", linenumber: A_LineNumber})
                        case 2:
                            arrFileGuideLine.Push(G_CONSTANTS.file_FL_Solar_StringInv)
                            EventManager.Publish("DebugMsg",{msg: "MakeFile_CTRL_ASS: Selezionato String Inverter", linenumber: A_LineNumber})
                        case 3:
                            arrFileGuideLine.Push(G_CONSTANTS.file_FL_Solar_InvModule)
                            EventManager.Publish("DebugMsg",{msg: "MakeFile_CTRL_ASS: Selezionato Inverter Module", linenumber: A_LineNumber})
                        default:
                            MakeFileCTRL_ASS_result.success := false
                            MakeFileCTRL_ASS_result.error := "Errore nella selezione tecnologia inverter"
                            EventManager.Publish("ProcessError", {processId: "SelectInverter", details: "Nessuna tipologia di inverter selezionato.", result: {}})
                            EventManager.Publish("AddLV", {icon: "icon2", element: "", text: "Errore codice inverte"})
                            EventManager.Publish("PI_Stop", {inputValue: "Errore nel codice inverter "}) ; Avvia l'indicatore di progresso 
                            return MakeFileCTRL_ASS_result
                    }
                }
                else if (tech = "E") { ; verifico gli impianti con tecnologia BESS
                    arrFileGuideLine.Push(G_CONSTANTS.file_FL_B_SubStation)
                    arrFileGuideLine.Push(G_CONSTANTS.file_FL_Bess)
                }
                else if (tech = "W") { ; verifico gli impianti con tecnologia WIND
                    arrFileGuideLine.Push(G_CONSTANTS.file_FL_W_SubStation)
                    arrFileGuideLine.Push(G_CONSTANTS.file_FL_Wind)
                }
                               
                ; genero una struttura map() a partire dalla linee guida della relativa tecnologia
                EventManager.Publish("ProcessProgress", {processId: "MakeFile_CTRL_ASS", status: "In Progress", details: "Genero map a partire dalle guideline", result: {}})     
                map_Guideline := MakeFileUpload_CTRL_ASS.MakeMapGuideline(arrFileGuideLine)
                MakeString_CTRL_ASS_result := MakeFileUpload_CTRL_ASS.MakeString_CTRL_ASS(map_Guideline.value, arr) 
                ; genera un oggetto contenente il risutlato del metodo MakeString_CTRL_ASS_result
                ; result := { success: false, value: false, error: "", class: "MakeFileUpload_CTRL_ASS.ahk", function: "MakeString_CTRL_ASS" }
                nomeFile_ZPMR_CTRL_ASS := MakeFileUpload_CTRL_ASS.GetFileName(G_CONSTANTS.file_ZPMR_CTRL_ASS_UpLoad)
                EventManager.Publish("ProcessProgress", {processId: "MakeFile_CTRL_ASS", status: "In Progress", details: "Scrivo file " . nomeFile_ZPMR_CTRL_ASS, result: {}})     
                if (MakeString_CTRL_ASS_result.success = true) { ; se la creazione del file è andata a buon fine allora scrivo il file su disco
                    WriteStringToFile_result := MakeFileUpload_CTRL_ASS.WriteStringToFile(MakeString_CTRL_ASS_result.value, G_CONSTANTS.file_ZPMR_CTRL_ASS_UpLoad)
                    if (WriteStringToFile_result = false) {
                        errorMsg := "Errore nelle generazione del file CTR_ASS" . " - LN: " . A_LineNumber
                        throw Error(errorMsg) 
                    }
                    else {
                        EventManager.Publish("AddLV", {icon: "icon4", element: nomeFile_ZPMR_CTRL_ASS, text: "File aggiornato"})
                        EventManager.Publish("ProcessCompleted", {processId: "MakeFile_CTRL_ASS", status: "Completed", details: "Esecuzione completata con successo", result: {}})
                        MakeFileCTRL_ASS_result.success := true
                        MakeFileCTRL_ASS_result.value := true
                        return MakeFileCTRL_ASS_result
                    }
                }
                else {
                    errorMsg := "Errore nelle generazione della stringa CTR_ASS" . " - LN: " . A_LineNumber
                    throw Error(errorMsg) 
                }                
            }    
            catch as err {
                MakeFileCTRL_ASS_result.error := "Errore: " . err.Message          
                EventManager.Publish("ProcessError", {processId: "MakeMapGuideline", status: "Error", details: MakeFileCTRL_ASS_result.error, result: {}})
                MsgBox(MakeFileCTRL_ASS_result.error, "Error", 4144)
                return MakeFileCTRL_ASS_result  
            }              
    }

    ; Funzione: MakeString_CTRL_ASS
    ; Descrizione: esamina ogni elemento presente nell'array contenente la lista delle FL non presenti nella CTRL_ASS e lo confronta con il pattern
    ; ricavato dalla guideline tramite la funzione MakeMapGuideline.
    ; Crea una stringa da utilizzare per creare il file di caricamento per l'aggiornamento della CTRL_ASS
    ; Parametri:
    ;   - param1:   il map() contenente come chiave il pattern delle FL ricavato a partire dalle guideline 
    ;   - param2:   array contenente la lista delle FL non presenti nella tabella deolle CTRL_ASS
    ; Restituisce:  Una stringa  per generare il file di caricamento per l'aggiornamento della CTRL_ASS. 
    ;               Considera solo le FL con lunghezza pari a 4, 5, 6.
    ; Esempio: 
    ; 
    static MakeString_CTRL_ASS(mapGuideline, FL_Arr) {
        result := { success: false, value: false, error: "", class: "MakeFileUpload_CTRL_ASS.ahk", function: "MakeString_CTRL_ASS" }
        content_CTRL_ASS := ""
        rowCount := 0
        for element in FL_Arr {        
            count := 0
            for key, data in mapGuideline { ; considero tutti gli elementi della Guideline, la chiave è il pattern da verificare con espressione regolare
                if MakeFileUpload_CTRL_ASS.IsValidString(element, key) {
                    ; considero solo le FL con lunghezza 4, 5, 6
                    LunghezzaFL := SubStr(element, -1)
                    patternFL := StrSplit(element, "_")
                    switch LunghezzaFL
                    {
                        case 4:
                            data.VALUE := patternFL[1]
                            data.SUB_VALUE := patternFL[2]
                        case 5:
                            data.VALUE := patternFL[1]
                            data.SUB_VALUE := patternFL[2]
                            data.SUB_VALUE2 :=patternFL[3]
                        case 6:
                            data.VALUE := patternFL[1]
                            data.SUB_VALUE := patternFL[2]
                            data.SUB_VALUE2 := patternFL[3]
                        default:
                            continue
                    }
                    myData := data
                    ;data := {VALUE : "", SUB_VALUE : "", SUB_VALUE2 : "", TPLKZ : "", FLTYP : "", FLLEVEL : "", CODE_SEZ_PM : "", CODE_SIST : "", CODE_PARTE : "", TIPO_ELEM : ""}                    
                    count++ ; per verificare che non ci sia più di una corrispondenza
                    }
            }
            Switch count
            {
            Case 0: ; nessuna corrispondenza trovata
                ;~ MakeFileUpload_CTRL_ASS.mainGui.gui.LV.Add("icon3", element, "Errore Guideline")
                ;~ MakeFileUpload_CTRL_ASS.mainGui.gui.LV.ModifyCol(1, "autoHdr")
                ;~ msgbox("Nessuna corrispondenza trovato nella guideline per: " . element)
                EventManager.Publish("AddLV", {icon: "icon3", element: element, text: "CTRL_ASS: Nessuna corrispondenza trovata in Guideline"})
                result.error := element . " - Nessuna corrispondenza trovata. `n"
            Case 1: ; una corrispondenza trovata
                ;EventManager.Publish("Debug",("MakeString_CTRL_ASS: l'elemento " . element . " è corretto."))
                content_CTRL_ASS .= myData.VALUE . ";" . myData.SUB_VALUE . ";" . myData.SUB_VALUE2 . ";" . myData.TPLKZ . ";" . myData.FLTYP . ";" . myData.FLLEVEL . ";" . myData.CODE_SEZ_PM . ";" . myData.CODE_SIST . ";" . myData.CODE_PARTE . ";" . myData.TIPO_ELEM . "`r`n"
                rowCount++
            Default:
                EventManager.Publish("AddLV", {icon: "icon2", element: element, text: "CTRL_ASS: Molteplici corrispondenze in Guideline"})
                ;~ msgbox("Molteplici corrispondenze trovato nella guideline per: " . element)
                result.error := element . " - Molteplici corrispondenze trovate. `n"
            }
        }
        OutputDebug("Elementi presenti in array: " . FL_Arr.length . " - Righe generate in CTRL_ASS: " . rowCount . "`n")
        if (result.error = "") {
            result.value := G_CONSTANTS.intestazione_CTRL_ASS .  content_CTRL_ASS
            result.success := true
        }
        return result
    }

    ; Funzione: MakeMapGuideline
    ; Descrizione: Legge il contenuto di un file guideline e crea un map() contenente gli elementi presenti nella tabella e necessari alla creazione dei file
    ; Parametri:
    ;   - param1: Il filename del file guideline
    ; Restituisce:  Un map() che ha come chiave l'espressione regolare relativa alla riga della linea guida e come valori un oggetto contenente gli elementi presenti nella tabella
    ;               e necessari alla costruzione dei file x upload.
    ;               Considera solo l FL con lunghezza pari a 4, 5, 6.
    ; Esempio: 
    ; 
    static MakeMapGuideline(arrFilename) {
        result := { success: false, value: false, error: "", class: "MakeFileUpload_CTRL_ASS.ahk", function: "MakeMapGuideline" }
        EventManager.Publish("ProcessStarted", {processId: "MakeMapGuideline", status: "Started", details: "Avvio funzione", result: {}})
        ; Inizializzo un map per contenere i codici FL e i relativi dati
        if(arrFilename.length > 0) { ; verifico che l'array file contenga elementi
            mapGuideline := map()
            for singleFile in arrFilename { ; per ogni file presente nell'array
                EventManager.Publish("ProcessProgress", {processId: "MakeMapGuideline", status: "In Progress", details: "Lettura file: " . singleFile, result: {}})     
                try {
                    ; Legge il contenuto del file
                    fileContent := FileRead(singleFile)
                    ; Divide il contenuto in linee
                    lines := StrSplit(fileContent, "`n", "`r")
                    ; Rimuove l'intestazione
                    lines.RemoveAt(1)
                } catch Error as err {
                    EventManager.Publish("ProcessError", {processId: "MakeMapGuideline", status: "Error", details: "Errore nella lettura del file: " . singleFile . " - " . err.Message " LN: " . A_LineNumber, result: {}})
                    MsgBox("Errore nella lettura del file: " . singleFile . " - " . err.Message, "Errore", 4112)
                    result.error := "Errore: " . err.Message
                    return result 
                }
                EventManager.Publish("ProcessProgress", {processId: "MakeMapGuideline", status: "In Progress", details: "Creo struttura map", result: {}})
                try {
                    ; genero un map a partire dal file delle regole delle guideline
                    MapRules := MakeFileUpload_CTRL_ASS.MakeMapRules(G_CONSTANTS.file_Rules)
                    ; Analizzo il contenuto del file
                    for line in lines {
                        data := {VALUE : "", SUB_VALUE : "", SUB_VALUE2 : "", TPLKZ : "", FLTYP : "", FLLEVEL : "", CODE_SEZ_PM : "", CODE_SIST : "", CODE_PARTE : "", TIPO_ELEM : ""}
                        if (line != "") {  ; Ignora le linee vuote
                            ; parts := StrSplit(line, "`t") ; utilizzato per i file di tipo .txt con separatore di elenco TAB
                            element := StrSplit(line, ";") ; utilizzato con file .csv con separatore di elenco ;
                            ; element è un array contenente gli elementi presenti nella linea "line"
                            ; il primo elemento è il codice FL
                            LunghezzaFL := StrSplit(element[1], "-").Length ; conto il numero di elementi per ricavare la lunghezza della FL
                            if (LunghezzaFL >= 4) and (LunghezzaFL <= 6) {
                                data.FLLEVEL := LunghezzaFL
                                patternFL := StrSplit(element[1], "-") ; creo un array contenente i diversi livelli della FL
                                technology := SubStr(patternFL[1], 3, 1) 
                                data.TPLKZ := "Z-R" . technology . "S"
                                data.FLTYP := technology
                                data.CODE_SEZ_PM := element[3]
                                data.CODE_SIST := element[4]
                                data.CODE_PARTE := element[5]
                                data.TIPO_ELEM := element[6]
                                ; i dati nelle tabelle sono nel seguente ordine                    
                                /*     3	            4	        5	            6
                                    AM Section  	AM Part	    AM Component	Element Type
                                    CODE_SEZ_PM	    CODE_SIST	CODE_PARTE	    TIPO_ELEM
                                */    
                                ; considero solo le FL con lunghezza 4, 5, 6
                                switch LunghezzaFL
                                {
                                    case 4:
                                        patternKey := ("^" . MakeFileUpload_CTRL_ASS.MakeSingeLevelPattern(patternFL[4], MapRules) . "_" . MakeFileUpload_CTRL_ASS.MakeSingeLevelPattern(patternFL[3], MapRules) . "_4$")
                                    case 5:
                                        patternKey := ("^" . MakeFileUpload_CTRL_ASS.MakeSingeLevelPattern(patternFL[5], MapRules) . "_" . MakeFileUpload_CTRL_ASS.MakeSingeLevelPattern(patternFL[4], MapRules) . "_" . MakeFileUpload_CTRL_ASS.MakeSingeLevelPattern(patternFL[3], MapRules) . "_5$")
                                    case 6:
                                        patternKey := ("^" . MakeFileUpload_CTRL_ASS.MakeSingeLevelPattern(patternFL[6], MapRules) . "_" . MakeFileUpload_CTRL_ASS.MakeSingeLevelPattern(patternFL[5], MapRules) . "_" . MakeFileUpload_CTRL_ASS.MakeSingeLevelPattern(patternFL[4], MapRules) . "_6$")
                                    default:
                                        continue
                                }
                                ; Verifico se esiste già una chiave con il valore trovato
                                if (!mapGuideline.Has(patternKey))
                                    mapGuideline[patternKey] := {}
                                else { ; se esiste già la stessa chiave allora genero un errore
                                    ;EventManager.Publish("AddLV", {icon: "icon3", element: patternKey, text: "MakeMapGuideline - Chiave già presente"})
                                    continue ; OK!
                                    ;throw Error("Chiave mapGuideline già esistente")
                                }
                                ; inserisco i valori nell'oggetto
                                mapGuideline[patternKey] := data ; assegno l'oggetto al map()
                            }
                        }
                    }
                }
                catch as err {
                    result.error := "Errore: " . err.Message          
                    EventManager.Publish("ProcessError", {processId: "MakeMapGuideline", status: "Error", details: result.error . " - LN: " . A_LineNumber, result: {}})
                    MsgBox(result.error, "Error", 4144)
                    return result  
                }                
            }                  
            NumeroElementiMap := mapGuideline.Count
            if (mapGuideline.Count = 0) {
                EventManager.Publish("ProcessError", {processId: "MakeMapGuideline", status: "Error", details: "Map privo di elementi. LN: " . A_LineNumber, result: {}})
                throw Error("Map privo di elementi") 
            }
            else {
/*                 ; test enum
                OutputDebug( "Ricavati " . mapGuideline.Count . " Pattern a partire da guideline:" . "`n")
                for key, value in mapGuideline {
                    OutputDebug(key . "`n")
                } */
                EventManager.Publish("ProcessCompleted", {processId: "MakeMapGuideline", status: "Completed", details: "Esecuzione completata con successo", result: {}})
                result.success := true
                result.value :=  mapGuideline
                return result
            }               
        }
        else {
            result.error := "MakeMapGuideline: Errore array file vuoto!"
            MsgBox("MakeMapGuideline: Errore array file vuoto!", "Errore", 4112)
            EventManager.Publish("ProcessError", {processId: "MakeMapGuideline", status: "Error", details: result.error " - LN: " . A_LineNumber, result: {}})
            return result 

        }

    }
    
    ; *
    ; Funzione: MakeSingeLevelPattern
    ; Descrizione: Crea il pattern di un singolo livello in base alle regole presenti nei file delle linee guida da utilizzare nell'espressione regolare
    ; Parametri:
    ;   - param1: Un singolo livello della FL ricavato dalla guideline
    ;   - param2: Il map contenente le regole da utilizzare
    ; Restituisce: Una stringa contenente l'espressione regolare da utilizzare per la verifica
    ; Esempio: 
    static MakeSingeLevelPattern(FL_singleLevel, MapRules) {
            ; sostituisco le occorrenze nella stringa
            ; sostituisco inizialmente le occorrenze di "nn" e "pp" per non avere errori nella sostituzione delle singole "n" o "p"
            FL_singleLevel := StrReplace(FL_singleLevel, "nn" , MapRules["nn"], true, , -1)
            FL_singleLevel := StrReplace(FL_singleLevel, "pp" , MapRules["pp"], true, , -1)
            for key, value in MapRules {
                FL_singleLevel := StrReplace(FL_singleLevel, key , value, true, , -1)
            }
        return FL_singleLevel
    
    }
    ; *
    ; Funzione: GetFL_Codes
    ; Descrizione: Legge il file delle guideline e restituisce un array contenente solamente i codici delle FL presenti nelle guideline (primo campo)
    ; Parametri:
    ;   - param1: Il nome del file delle guideline
    ; Restituisce: un array contenente solamente i codici delle FL presenti nelle guideline (primo campo)
    ; Esempio:
    static GetFL_Codes(filename) {
        try {
            ; Legge il contenuto del file
            fileContent := FileRead(filename)

            ; Divide il contenuto in linee
            lines := StrSplit(fileContent, "`n", "`r")
            ; Rimuove l'intestazione
            arrN_Campi := strsplit(lines[1], ";")
            N_Campi := arrN_Campi.Length
            lines.RemoveAt(1)
            ; Inizializza un array per i codici FL
            FL_Codes := []

            ; Estrae i codici paese
            for line in lines {
                if (line != "") {  ; Ignora le linee vuote
                    ; parts := StrSplit(line, "`t") ; utilizzato per i file di tipo .txt con separatore di elenco TAB
                    parts := StrSplit(line, ";") ; utilizzato con file .csv con separatore di elenco ;
                    if (parts.Length = N_Campi) { ; da verificare che tutte le FL siano associate ad una descrizione, ovvero composte da [0]FL [1]Descrizione
                        code := parts[1]  ; Prende il primo elemento
                        if (StrLen(code) >= 3) { ; se la lunghezza è almeno tre caratteri
                            FL_Codes.Push(code)
                        }
                    }
                    else {
                        MsgBox("Errore nella struttura del file: " . filename, "Errore", 4112)
                        return false
                    }
                }
            }
            return FL_Codes
        } catch Error as err {
            MsgBox("Errore nella lettura del file: " . filename . " - " . err.Message, "Errore", 4112)
            return false
        }
    }    
    
    ; *
    ; Funzione: MakeMapRules
    ; Descrizione: Legge il file contenente le regole e crea un map() con chiave i codici presenti nella guideline e come valori i codici da utilizzare nell'espressione regolare.
    ; Parametri:
    ;   - param1: Il nome del file contenente le regole per costruire le espressioni regolari.
    ; Restituisce: Un map() con chiave -> codice, valore -> valori x espressione regolare
    ; Esempio: map["nn"] := (?!00)[0-9]{2})
    static MakeMapRules(filename) {
        MapRules := Map()
        try {
            ; Legge il contenuto del file
            fileContent := FileRead(filename)

            ; Divide il contenuto in linee
            lines := StrSplit(fileContent, "`n", "`r")
            ; Ricavo analizzando l'intestazione il numero di campi in cui è composto il file
            Arr_N_Campi := StrSplit(lines[1], ";")
            N_Campi := Arr_N_Campi.Length
            ; Rimuove l'intestazione
            lines.RemoveAt(1)
            ; Estrae la codifica inserendo come chiave la lettera e come valore il pattern
            for line in lines {
                if (line != "") {  ; Ignora le linee vuote
                    ; parts := StrSplit(line, "`t") ; utilizzato per i file di tipo .txt con separatore di elenco TAB
                    parts := StrSplit(line, ";") ; utilizzato con file .csv con separatore di elenco ;
                    if (parts.Length = N_Campi) { ; da verificare che tutte le righe rispettino il numero di campi dell'intestazione
                        if (StrLen(parts[1]) = 1) or (StrLen(parts[1]) = 2) { ; verifico che il primo elemento contenga uno o due caratteri
                            MapRules[parts[1]] := parts[2]  ; Crea map chiave - valore
                        }
                        else {
                            MsgBox("Errore lunghezza codice chiave del file: " . FileName, "Errore", 4112)
                            return false
                        }
                    }
                    else {
                        MsgBox("Errore nella struttura chiave - valore del file: " . FileName, "Errore", 4112)
                        return false
                    }
                }
            }
            EventManager.Publish("Debug",("Struttura map rules creata"))
            return MapRules
        } catch Error as err {
            MsgBox("Errore nella lettura del file: " . filename . " - " . err.Message, "Errore", 4112)
            return false
        }
    }

    ; Funzione per verificare se una stringa è valida
    static IsValidString(str, pattern) {
        return RegExMatch(str, pattern) ? true : false
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

    ; Funzione che estrae il nome file con estensione da un percorso completo
    Static GetFileName(filePath) {
        ; Verifica che il parametro non sia vuoto
        if !filePath
            throw ValueError("Il percorso del file non può essere vuoto")
        
        ; Rimuove eventuali spazi iniziali e finali
        filePath := Trim(filePath)
        
        ; Trova l'ultima occorrenza di \ o /
        lastBackslash := InStr(filePath, "\", , -1)  ; Cerca da destra
        lastForwardSlash := InStr(filePath, "/", , -1)  ; Cerca da destra
        
        ; Usa il separatore trovato più a destra
        lastSeparator := Max(lastBackslash, lastForwardSlash)
        
        ; Se non trova separatori, restituisce il percorso originale
        if (lastSeparator = 0)
            return filePath
        
        ; Estrae il nome del file
        fileName := SubStr(filePath, lastSeparator + 1)
        
        return fileName
    }
;-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

}