#Requires AutoHotkey v2.0

class MakeFileUpload_CTRL_ASS {

; ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
; * Realizzazione dei file per l'aggiornamento della Control Table Asset
;   Deve essere creato il file:
;   -   file_ZPMR_CTRL_ASS_UpLoad
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

/*     static MakeFileUpDateCTRL_ASS(arr, tech, invType:="") {
        MakeFileCTRL_ASS_result := { success: false, value: false, error: "", class: "MakeFileUpload_CTRL_ASS.ahk", function: "MakeFileCTRL_ASS" }
        EventManager.Publish("ProcessStarted", {processId: MakeFileCTRL_ASS_result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Crea file aggiornamento CTRL_ASS "}) ; Avvia l'indicatore di progresso 
            if (arr.Length = 0) {
                msgbox("Errore l'array indicato è vuoto!", "Errore", 4112)
                MakeFileCTRL_ASS_result.error := "Errore l'array indicato è vuoto!"
                return MakeFileCTRL_ASS_result  ; Se l'array è vuoto allora restituisco un errore
            }  
        
            MakeFileCTRL_ASS_result.success := true ; imposto a true e cambio in caso venga riscontrato almeno un errore

            if (tech = "S") { ; verifico gli impianti con tecnologia SOLAR
                ; Creo dei map che hanno per chiave il codice FL della guidelinee come valore un oggetto composto dagli element:
                ; {VALUE: "", SUB_VALUE: "", SUB_VALUE2: "", TPLKZ: "", FLTYP: "", FLLEVEL: "", CODE_SEZ_PM: "", CODE_SIST: "", CODE_PARTE: "", TIPO_ELEM: ""}
                

                pattern_array_FL_Solar_Common := MakeFileUpload_CTRL_ASS.MakePattern_array(MakeFileUpload_CTRL_ASS.GetFL_Codes(G_CONSTANTS.file_FL_Solar_Common))
                pattern_array_FL_S_SubStation := MakeFileUpload_CTRL_ASS.MakePattern_array(MakeFileUpload_CTRL_ASS.GetFL_Codes(G_CONSTANTS.file_FL_S_SubStation))

                ; in base alla tipologia di inverter creo il relativo array
                switch invType
                {
                    case 1:
                        pattern_array_FL_Solar_Technology := MakeFileUpload_CTRL_ASS.MakePattern_array(MakeFileUpload_CTRL_ASS.GetFL_Codes(G_CONSTANTS.file_FL_Solar_CentralInv))
                        EventManager.Publish("DebugMsg",{msg: "Selezionato Central Inverter", linenumber: A_LineNumber})
                    case 2:
                        pattern_array_FL_Solar_Technology := MakeFileUpload_CTRL_ASS.MakePattern_array(MakeFileUpload_CTRL_ASS.GetFL_Codes(G_CONSTANTS.file_FL_Solar_StringInv))
                        EventManager.Publish("DebugMsg",{msg: "Selezionato String Inverter", linenumber: A_LineNumber})
                    case 3:
                        pattern_array_FL_Solar_Technology := MakeFileUpload_CTRL_ASS.MakePattern_array(MakeFileUpload_CTRL_ASS.GetFL_Codes(G_CONSTANTS.file_FL_Solar_InvModule))
                        EventManager.Publish("DebugMsg",{msg: "Selezionato Inverter Module", linenumber: A_LineNumber})
                    default:
                        MakeFileCTRL_ASS_result.success := false
                        MakeFileCTRL_ASS_result.error := "Errore nella selezione tecnologia inverter"
                        EventManager.Publish("ProcessError", {processId: "SelectInverter", details: "Nessuna tipologia di inverter selezionato.", result: {}})
                        EventManager.Publish("AddLV", {icon: "icon2", element: "", text: "Programma interrotto dall'utente"})
                        EventManager.Publish("PI_Stop", {inputValue: "Errore nella selezione tecnologia inverter "}) ; Avvia l'indicatore di progresso 
                        return MakeFileCTRL_ASS_result
                }
                ; Controllo i valori al terzo livello
                EventManager.Publish("PI_Start", {inputValue: "Verifica guideline "}) ; Avvia l'indicatore di progresso
                for element in arr { ; per ogni riga
                    FL_Levels := StrSplit(trim(element), "-") ; scompongo la riga nei singoli livelli
                    if (FL_Levels.Length > 2) {
                        if ((FL_Levels[3] = "00") or (FL_Levels[3] = "9Z") or (FL_Levels[3] = "ZZ")) { ; Common
                            if !(MakeFileUpload_CTRL_ASS.TestGuideline_array(pattern_array_FL_Solar_Common, element)) { ; confronto l'elemento con le linee guida
                                MakeFileCTRL_ASS_result.error := "Errore guideline Common."
                                MakeFileCTRL_ASS_result.success := false
                                EventManager.Publish("AddLV", {icon: "icon3", element: " " . element, text: result.error})
                                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})
                            }
                        }
                        else if FL_Levels[3] = "0A" { ; Substation
                            if !(MakeFileUpload_CTRL_ASS.TestGuideline_array(pattern_array_FL_S_SubStation, element)) { ; confronto l'elemento con le linee guida
                                MakeFileCTRL_ASS_result.error := "Errore guideline Substation."        
                                MakeFileCTRL_ASS_result.success := false
                                EventManager.Publish("AddLV", {icon: "icon3", element: " " . element, text: result.error})
                                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})
                            }
                        }
                        else { ; in tutti gli altri casi devo considerare il tipo di tecnologia degli inverter
                            if !(MakeFileUpload_CTRL_ASS.TestGuideline_array(pattern_array_FL_Solar_Technology, element)) { ; confronto l'elemento con le linee guida
                                result.error := "Errore guideline Inverter."                
                                result.success := false
                                EventManager.Publish("AddLV", {icon: "icon3", element: " " . element, text: result.error})
                                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})
                            }
                        }
                    }
                }
                if (MakeFileCTRL_ASS_result.success = true) {
                    EventManager.Publish("AddLV", {icon: "icon1", element: "", text: "Controllo guideline OK"})
                    EventManager.Publish("ProcessCompleted", {processId: MakeFileCTRL_ASS_result.function, status: "Completed", details: "Esecuzione completata con successo", result: {}})
                }
                return MakeFileCTRL_ASS_result
            }
            else if (tech = "E") { ; verifico gli impianti con tecnologia BESS
                ; Creo degli array a partire dai file
                pattern_array_FL_B_SubStation := MakeFileUpload_CTRL_ASS.MakePattern_array(MakeFileUpload_CTRL_ASS.GetFL_Codes(G_CONSTANTS.file_FL_B_SubStation))
                pattern_array_FL_Bess_Technology := MakeFileUpload_CTRL_ASS.MakePattern_array(MakeFileUpload_CTRL_ASS.GetFL_Codes(G_CONSTANTS.file_FL_Bess))
                ; Controllo i valori al terzo livello
                for element in Arr { ; per ogni riga
                    FL_Levels := StrSplit(trim(element), "-") ; scompongo la riga nei singoli livelli
                    if (FL_Levels.Length > 2) {
                        if FL_Levels[3] = "0A" { ; Substation
                            if !(MakeFileUpload_CTRL_ASS.TestGuideline_array(pattern_array_FL_B_SubStation, element)) { ; confronto l'elemento con le linee guida
                                result.error := "Errore guideline Substation."
                                result.success := false
                                EventManager.Publish("AddLV", {icon: "icon3", element: " " . element, text: result.error})
                                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})             
                            }
                        }
                        else { ; in tutti gli altri casi devo considerare il tipo di tecnologia degli inverter
                            if !(MakeFileUpload_CTRL_ASS.TestGuideline_array(pattern_array_FL_Bess_Technology, element)) { ; confronto l'elemento con le linee guida
                                result.error := "Errore guideline Bess."                
                                result.success := false
                                EventManager.Publish("AddLV", {icon: "icon3", element: " " . element, text: result.error})
                                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})                                                  
                            }
                        }
                    }
                }
                if (result.success = true) {
                    EventManager.Publish("AddLV", {icon: "icon1", element: "", text: "Controllo guideline OK"})
                    EventManager.Publish("ProcessCompleted", {processId: result.function, status: "Completed", details: "Esecuzione completata con successo", result: {}})
                }
                return result
            }
            else if (tech = "W") { ; verifico gli impianti con tecnologia WIND
                ; Creo degli array a partire dai file
                pattern_array_FL_W_SubStation := MakeFileUpload_CTRL_ASS.MakePattern_array(MakeFileUpload_CTRL_ASS.GetFL_Codes(G_CONSTANTS.file_FL_W_SubStation))
                pattern_array_FL_Wind_Technology := MakeFileUpload_CTRL_ASS.MakePattern_array(MakeFileUpload_CTRL_ASS.GetFL_Codes(G_CONSTANTS.file_FL_Wind))
                ; Controllo i valori al terzo livello
                for element in Arr { ; per ogni riga
                    FL_Levels := StrSplit(trim(element), "-") ; scompongo la riga nei singoli livelli
                    if (FL_Levels.Length > 2) {
                        if FL_Levels[3] = "0A" { ; Substation
                            if !(MakeFileUpload_CTRL_ASS.TestGuideline_array(pattern_array_FL_W_SubStation, element)) { ; confronto l'elemento con le linee guida
                                result.error := "Errore guideline Substation."
                                result.success := false
                                EventManager.Publish("AddLV", {icon: "icon3", element: " " . element, text: result.error})
                                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})             
                            }
                        }
                        else { ; in tutti gli altri casi devo considerare il tipo di tecnologia degli inverter
                            if !(MakeFileUpload_CTRL_ASS.TestGuideline_array(pattern_array_FL_Wind_Technology, element)) { ; confronto l'elemento con le linee guida
                                result.error := "Errore guideline Bess."                
                                result.success := false
                                EventManager.Publish("AddLV", {icon: "icon3", element: " " . element, text: result.error})
                                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})                                                  
                            }
                        }
                    }
                }
                if (result.success = true) {
                    EventManager.Publish("AddLV", {icon: "icon1", element: "", text: "Controllo guideline OK"})
                    EventManager.Publish("ProcessCompleted", {processId: result.function, status: "Completed", details: "Esecuzione completata con successo", result: {}})
                }
                return result
            }            
    } */

    ; Funzione: TestGuideline_array
    ; Descrizione: esamina ogni elemento presente nell'array contenente la lista delle FL non presenti nella CTRL_ASS e lo confronta con il pattern
    ; ricavato dalla guideline tramite la funzione MakeMapGuideline
    ; Parametri:
    ;   - param1:   array contenente la lista delle FL non presenti nella tabella deolle CTRL_ASS
    ;   - param2:   il map() contenente come chiave il pattern delle FL ricavato a partire dalle guideline 
    ; Restituisce:  Un map() che ha come chiave l'espressione regolare relativa alla riga della linea guida e come valori un oggetto contenente gli elementi presenti 
    ;               nella tabella e necessari alla costruzione dei file x upload.
    ;               Considera solo le FL con lunghezza pari a 4, 5, 6.
    ; Esempio: 
    ; 
    static TestGuideline(mapGuideline, FL_Arr) {
        result := { success: false, value: false, error: "", class: "MakeFileUpload_CTRL_ASS.ahk", function: "TestGuideline" }
        content_CTRL_ASS := ""
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
                    count += 1 ; per verificare che non ci sia più di una corrispondenza
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
                ;EventManager.Publish("Debug",("TestGuideline: l'elemento " . element . " è corretto."))
                content_CTRL_ASS .= myData.VALUE . ";" . myData.SUB_VALUE . ";" . myData.SUB_VALUE2 . ";" . myData.TPLKZ . ";" . myData.FLTYP . ";" . myData.FLLEVEL . ";" . myData.CODE_SEZ_PM . ";" . myData.CODE_SIST . ";" . myData.CODE_PARTE . ";" . myData.TIPO_ELEM . "`r`n"
            Default:
                EventManager.Publish("AddLV", {icon: "icon2", element: element, text: "CTRL_ASS: Molteplici corrispondenze in Guideline"})
                ;~ msgbox("Molteplici corrispondenze trovato nella guideline per: " . element)
                result.error := element . " - Molteplici corrispondenze trovate. `n"
            }
        }
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
    static MakeMapGuideline(filename) {
        result := { success: false, value: false, error: "", class: "MakeFileUpload_CTRL_ASS.ahk", function: "MakeMapGuideline" }
        EventManager.Publish("ProcessStarted", {processId: "MakeMapGuideline", status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("ProcessProgress", {processId: "MakeMapGuideline", status: "In Progress", details: "Lettura file: " . filename, result: {}})     
        try {
            ; Legge il contenuto del file
            fileContent := FileRead(filename)
            ; Divide il contenuto in linee
            lines := StrSplit(fileContent, "`n", "`r")
            ; Rimuove l'intestazione
            lines.RemoveAt(1)
        } catch Error as err {
            EventManager.Publish("ProcessError", {processId: "MakeMapGuideline", status: "Error", details: "Errore nella lettura del file: " . filename . " - " . err.Message " LN: " . A_LineNumber, result: {}})
            MsgBox("Errore nella lettura del file: " . filename . " - " . err.Message, "Errore", 4112)
            return false
        }
        EventManager.Publish("ProcessProgress", {processId: "MakeMapGuideline", status: "In Progress", details: "Creo struttura map", result: {}})
        try {
            ; genero un map a partire dal file delle regole delle guideline
            MapRules := MakeFileUpload_CTRL_ASS.MakeMapRules(G_CONSTANTS.file_Rules)
            ; Inizializzo un map per contenere i codici FL e i relativi dati
            mapGuideline := map()
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
                            EventManager.Publish("AddLV", {icon: "icon3", element: patternKey, text: "MakeMapGuideline - Chiave già presente"})
                            throw Error("Chiave mapGuideline già esistente")
                        }
                        ; inserisco i valori nell'oggetto
                        mapGuideline[patternKey] := data ; assegno l'oggetto al map()
                    }
                }
            }
            NumeroElementiMap := mapGuideline.Count
            if (mapGuideline.Count = 0) {
                EventManager.Publish("ProcessError", {processId: "MakeMapGuideline", status: "Error", details: "Map privo di elementi. LN: " . A_LineNumber, result: {}})
                throw Error("Map privo di elementi", "Verifica tabella CTRL_ASS", 4132) 
            }
            else {
                EventManager.Publish("ProcessCompleted", {processId: "MakeMapGuideline", status: "Completed", details: "Esecuzione completata con successo", result: {}})
                result.success := true
                result.value :=  mapGuideline
                return result
            }
        }                
        catch as err {
            result.error := "Errore: " . err.Message          
            EventManager.Publish("ProcessError", {processId: "MakeMapGuideline", status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})
            MsgBox(result.error, "Error", 4144)
            return result  
        }
    }
    
    ; *
    ; Funzione: MakeSingeLevelPattern
    ; Descrizione: Crea il pattern di un singolo livello in base alle regole presenti nei file delle linee guida da utilizzare nell'espressione regolare
    ; Parametri:
    ;   - param1: Un singolo livello della FL
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
    ; Funzione: MakePattern_array
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
            ; Rimuove l'intestazione
            lines.RemoveAt(1)
            ; Estrae la codifica inserendo come chiave la lettera e come valore il pattern
            for line in lines {
                if (line != "") {  ; Ignora le linee vuote
                    ; parts := StrSplit(line, "`t") ; utilizzato per i file di tipo .txt con separatore di elenco TAB
                    parts := StrSplit(line, ";") ; utilizzato con file .csv con separatore di elenco ;
                    if (parts.Length = 2) { ; da verificare che tutte i codici siano associati ad un pattern
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
;-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

}