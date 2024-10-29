#Requires AutoHotkey v2.0

#Include SAP_Connection.ahk

class SAP_ControlAsset {

    static __New() {
        SAP_ControlAsset.InitializeVariables()
        SAP_ControlAsset.SetupEventListeners()
    }

    Static InitializeVariables() {
        SAP_ControlAsset.DebugMode := 0 ; 0 -> leggo la tabella da SAP; 1 -> leggo la tabella da file
    }

    Static SetupEventListeners() {
        EventManager.Subscribe("VerificaFL_SAP_ControlAsset", (data) => SAP_ControlAsset.VerificaControlAsset(data.flArray, data.flcountry, data.fltechnology, data.flInverterTechnology))
    }

    ; Metodo: VerificaControlAsset
    ; Descrizione: Esegue la verifica degli elementi delle FL nelle tabelle SAP dei control asset
    ; Parametri:
    ; Restituisce: 
    ;   - un array contenente gli elementi non presenti nella CTRL_ASS se esistono
    ;   - un array vuoto altrimenti
    ; Esempio:
    static VerificaControlAsset(flArray, flcountry, fltechnology, flInverterTechnology) {
        VerificaControlAsset_result := { success: false, value: false, error: "", class: "SAP_ControlAsset.ahk", function: "VerificaControlAsset" }
        EventManager.Publish("ProcessStarted", {processId: VerificaControlAsset_result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica control asset"}) ; Avvia l'indicatore di progresso         
        
        mapControlAsset := map()
        data := {lunghezza: 0, conc: ""}
        Lista_FL_Conc := []
        ; analizzo gli elementi presenti nell'array FL
        EventManager.Publish("ProcessProgress", {processId: VerificaControlAsset_result.function, status: "In Progress", details: "Analizza elementi FL", result: {}})
        for element in flArray {
            temp_array := StrSplit(element, "-")
            numberOfElement := temp_array.Length
            ; i valori da controllare sono solo quelli >= 4
            if (numberOfElement >= 4) {
                ;OutputDebug("element: " . element . " - temp_array.Length: " . temp_array.Length . "`n")
                switch numberOfElement
                {
                case 4: concatena := SAP_ControlAsset.Concatena(temp_array, 2)
                case 5: concatena := SAP_ControlAsset.Concatena(temp_array, 3)
                case 6: concatena := SAP_ControlAsset.Concatena(temp_array, 3)
                }
                ; memorizzo i valori nel map
                ; preparo una struttura dati per memorizzare il risultato del confronto
                mapControlAsset[element] := {lunghezza: numberOfElement, conc: concatena, check: ""}
                Lista_FL_Conc.Push(concatena)
            }
        }
        ; se l'array contiene elementi allora scrivo il contenuto in un file
        if (Lista_FL_Conc.Length > 0) {
            ; Scrivo il contenuto dell'array in un file.
            SAP_ControlAsset.WriteArrayToFile(Lista_FL_Conc,  A_ScriptDir . "\Lista_FL_Conc.txt")
            OutputDebug("Elementi concatenati a partire dalle FL: " . Lista_FL_Conc.Length . "`n")    
        }
        else { ; se non contiene elementi concludo 
            VerificaControlAsset_result.error := "Nessun elemento da verificare"
            return VerificaControlAsset_result
        }

        if (SAP_ControlAsset.DebugMode = 0) { ; 0 -> leggo la tabella da SAP; 1 -> leggo la tabella da file
            OutputDebug("Prelevo i dati da SAP")
            ; estraggo i dati dalla tabella SAP filtrando in base alla tecnologia
            EventManager.Publish("ProcessProgress", {processId: VerificaControlAsset_result.function, status: "In Progress", details: "Estraggo tabella CTRL_ASS da SAP", result: {}})
            CTRL_ASS_Table := SAP_ControlAsset.EstraiTableControlAsset(fltechnology) ; restituisce un array contenente un array per ogni per ogni riga della tabella.
            if !(CTRL_ASS_Table) {
                VerificaControlAsset_result.error := "Errore estrazione tabella CTRL_ASS"
                EventManager.Publish("ProcessError", {processId: VerificaControlAsset_result.function, status: "Error", details: VerificaControlAsset_result.error . " - Line Number: " . A_LineNumber, result: {}})                   
                return VerificaControlAsset_result                
            }
            ; ricerco i valori delle intestazioni delle colonne nel file
            Index_CTRL_ASS_LivelloSedeTecnica := SAP_ControlAsset.TableHasIndex(CTRL_ASS_Table, "Liv.Sede")
            Index_CTRL_ASS_Valore_Livello := SAP_ControlAsset.TableHasIndex(CTRL_ASS_Table, "Valore Livello")
            Index_CTRL_ASS_Valore_Liv_Superiore_1 := SAP_ControlAsset.TableHasIndex(CTRL_ASS_Table, "Valore Liv. Superiore")
            Index_CTRL_ASS_Valore_Liv_Superiore_2 := SAP_ControlAsset.TableHasIndex(CTRL_ASS_Table, "Valore Liv. Superiore")
            Index_CTRL_ASS_Valore_StructureIndicator := SAP_ControlAsset.TableHasIndex(CTRL_ASS_Table, "Str. ")
            Index_CTRL_ASS_Valore_FL_Category := SAP_ControlAsset.TableHasIndex(CTRL_ASS_Table, "C")

        }
        else if (SAP_ControlAsset.DebugMode = 1) { ; 0 -> leggo la tabella da SAP; 1 -> leggo la tabella da file
            OutputDebug("Prelevo i dati da file")
            ; estraggo i dati da file
            EventManager.Publish("ProcessProgress", {processId: VerificaControlAsset_result.function, status: "In Progress", details: "Leggo file tabella CTRL_ASS", result: {}})
            ;CTRL_ASS_Table := SAP_ControlAsset.CreaArrayDaFile("C:\Users\a259046\OneDrive - Enel Spa\SCRIPT AHK e VBA\Functional_Location\CheckFL_rev.3\SAP\ZMPR_CTRL_ASS.txt")
            CTRL_ASS_Table := SAP_ControlAsset.CreaArrayDaFile("C:\Users\a259046\OneDrive - Enel Spa\SCRIPT AHK e VBA\Functional_Location\CheckFL_rev.3\SAP\ZMPR_CTRL_ASS_R4Q.txt")
            if !(CTRL_ASS_Table) {
                VerificaControlAsset_result.error := "Errore lettura file tabella CTRL_ASS"
                EventManager.Publish("ProcessError", {processId: VerificaControlAsset_result.function, status: "Error", details: VerificaControlAsset_result.error . " - Line Number: " . A_LineNumber, result: {}})                   
                EventManager.Publish("PI_Stop", {inputValue: "Errore verifica control asset"}) ; Ferma l'indicatore di progresso
                return VerificaControlAsset_result                
            }             

            ; ricerco i valori delle intestazioni delle colonne nel file
            Index_CTRL_ASS_LivelloSedeTecnica := SAP_ControlAsset.TableHasIndex(CTRL_ASS_Table, "Liv.Sede")
            Index_CTRL_ASS_Valore_Livello := SAP_ControlAsset.TableHasIndex(CTRL_ASS_Table, "Valore Livello")
            Index_CTRL_ASS_Valore_Liv_Superiore_1 := SAP_ControlAsset.TableHasIndex(CTRL_ASS_Table, "Valore Liv. Superiore")
            Index_CTRL_ASS_Valore_Liv_Superiore_2 := SAP_ControlAsset.TableHasIndex(CTRL_ASS_Table, "Valore Liv. Superiore")
            Index_CTRL_ASS_Valore_StructureIndicator := SAP_ControlAsset.TableHasIndex(CTRL_ASS_Table, "Str. ")
            Index_CTRL_ASS_Valore_FL_Category := SAP_ControlAsset.TableHasIndex(CTRL_ASS_Table, "C")

            ; Filtro la tabella in base al tipo di tecnologia.
            ; Esamino i campi Index_CTRL_ASS_Valore_FL_Category e Index_CTRL_ASS_Valore_StructureIndicator
            EventManager.Publish("ProcessProgress", {processId: VerificaControlAsset_result.function, status: "In Progress", details: "Filtro tabella CTRL_ASS in base alla tecnologia", result: {}})
            CTRL_ASS_Table := SAP_ControlAsset.FilterTabByTechnology(CTRL_ASS_Table, fltechnology, Index_CTRL_ASS_Valore_StructureIndicator, Index_CTRL_ASS_Valore_FL_Category)
        }

        OutputDebug("-- INDICI DELLA COLONNE PRESENTI NEL FILE --`n")
        OutputDebug("Index_CTRL_ASS_LivelloSedeTecnica " . Index_CTRL_ASS_LivelloSedeTecnica . "`n")
        OutputDebug("Index_CTRL_ASS_Valore_Livello " . Index_CTRL_ASS_Valore_Livello . "`n")
        OutputDebug("Index_CTRL_ASS_Valore_Liv_Superiore_1 " . Index_CTRL_ASS_Valore_Liv_Superiore_1 . "`n")
        OutputDebug("Index_CTRL_ASS_Valore_Liv_Superiore_2 " . Index_CTRL_ASS_Valore_Liv_Superiore_2 . "`n")
        OutputDebug("Index_CTRL_ASS_FL_Category " . Index_CTRL_ASS_Valore_FL_Category . "`n")
        OutputDebug("Index_CTRL_ASS_StructureIndicator " . Index_CTRL_ASS_Valore_StructureIndicator . "`n")        
        
        ; verifico che le intestazioni cercate abbiano un valore
        if !(Index_CTRL_ASS_LivelloSedeTecnica 
            AND Index_CTRL_ASS_Valore_Livello 
            AND  Index_CTRL_ASS_Valore_Liv_Superiore_1 
            AND Index_CTRL_ASS_Valore_Liv_Superiore_2
            AND  Index_CTRL_ASS_Valore_FL_Category 
            AND Index_CTRL_ASS_Valore_StructureIndicator) {
                VerificaControlAsset_result.error := "Errore intestazioni tabella CTRL_ASS"
                EventManager.Publish("ProcessError", {processId: VerificaControlAsset_result.function, status: "Error", details: VerificaControlAsset_result.error . " - Line Number: " . A_LineNumber, result: {}})                   
                EventManager.Publish("PI_Stop", {inputValue: "Errore verifica control asset"}) ; Ferma l'indicatore di progresso
                return VerificaControlAsset_result
            }

        ; La tabella CTRL_ASS ha lo stesso nome per due intestazioni, verifico quella corretta
        ; La colonna indicata come <Index_CTRL_ASS_Valore_Liv_Superiore_2> deve essere vuota per il livello sede tecnica = 4
        ; result := {column: 0, rows: 0, value: ""}
        Data := SAP_ControlAsset.TableHasValue(CTRL_ASS_Table, , column:=Index_CTRL_ASS_LivelloSedeTecnica, value:= "4") ; verifico in quale riga è contenuto il valore 4 per la colonna <Liv.Sede>
        if (Data.rows != 0) { ; ricavo la riga che contiene il valore 4
            ; verifico se Index_CTRL_ASS_Valore_Liv_Superiore_2 contiene un valore vuoto
            Data := SAP_ControlAsset.TableHasValue(CTRL_ASS_Table, Data.rows, column:=Index_CTRL_ASS_Valore_Liv_Superiore_2, )
            if (Data.value != "") { ; allora il valore non è corretto, devo invertire gli indici
                temp := Index_CTRL_ASS_Valore_Liv_Superiore_1
                Index_CTRL_ASS_Valore_Liv_Superiore_1 := Index_CTRL_ASS_Valore_Liv_Superiore_2
                Index_CTRL_ASS_Valore_Liv_Superiore_2 := temp
            }
        }

        EventManager.Publish("ProcessProgress", {processId: VerificaControlAsset_result.function, status: "In Progress", details: "Concateno valori per effettuare controllo", result: {}})
        ; a partire dalla tabella estratta da SAP o letta da file e filtrata creo un array con i valori concatenati
        SAP_ConcArr := SAP_ControlAsset.MakeConcTable(  CTRL_ASS_Table, 
                                                        Index_CTRL_ASS_LivelloSedeTecnica, 
                                                        Index_CTRL_ASS_Valore_Livello, 
                                                        Index_CTRL_ASS_Valore_Liv_Superiore_1, 
                                                        Index_CTRL_ASS_Valore_Liv_Superiore_2)
        
        OutputDebug(SAP_ConcArr.Length . " elementi in Tabella SAP_CTRL_ASS`n")    
        ; *** Scrivo il contenuto dell'array in un file.
        SAP_ControlAsset.WriteArrayToFile(SAP_ConcArr,  A_ScriptDir . "\SAP_CTRL_ASS_" . fltechnology . ".txt")

        temp_result_1 := SAP_ControlAsset.Check_CTRL_ASS_Table_slow(SAP_ConcArr, mapControlAsset)

        temp_result_2 := SAP_ControlAsset.Check_CTRL_ASS_Table(SAP_ConcArr, mapControlAsset)

        EventManager.Publish("ProcessCompleted", {processId: VerificaControlAsset_result.function, status: "Completed", details: "Esecuzione completata con successo", result: temp_result_2.value})
        EventManager.Publish("PI_Stop", {inputValue: "Verifica tabella Control Asset completata"}) ; Ferma l'indicatore di progresso

        VerificaControlAsset_result.success := true
        VerificaControlAsset_result.value := temp_result_2
        return VerificaControlAsset_result
    }

    ; Funzione: Check_CTRL_ASS_Table_slow
    ; Descrizione: Verifica ogni elemento presente nel map costruito a partire dalle FL con tutti gli elementila tabella CTRL_ASS presente in SAP
    ; Parametri:
    ;   - param1: un array contenente i valori presenti in SAP
    ;   - param2: un map costruito a partire dalle FL inserite nel tool e avente come chiave il codice della FL e come valore un oggetto {lunghezza: numberOfElement, conc: concatena, check: ""}
    ; Restituisce:  Un array costituito dagli elementi non presenti in SAP
    ; Esempio:
    ; Risultato:
    static Check_CTRL_ASS_Table_slow(SAP_ConcArr, mapControlAsset) {    
        ; costruisco un map contenente gli elementi della tabella SAP in cui la chiave identifica la lunghezza della FL mentre il valore è un array contenente tutte le FL di quella lunghezza
        Check_CTRL_ASS_Table_result := { success: false, value: false, error: "", class: "SAP_ControlAsset.ahk", function: "Check_CTRL_ASS_Table_Slow" }
        EventManager.Publish("ProcessStarted", {processId: Check_CTRL_ASS_Table_result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica tabella CTRL_ASS "}) ; Avvia l'indicatore di progresso 
        ; {lunghezza: numberOfElement, conc: concatena, check: ""}
        Element_NOK := []
        Count := 0
        for element, data in mapControlAsset {
            if (SAP_ControlAsset.ArrayHasValue(SAP_ConcArr, data.conc) ) {
                data.check := "OK"
            }
            else {
                data.check := "NOK"
                EventManager.Publish("AddLV", {icon: "icon3", element: element, text: "Non presente in control asset table"})
                count += 1
            }            
        }
        EventManager.Publish("ProcessProgress", {processId: Check_CTRL_ASS_Table_result.function, status: "In Progress", details: count . " elementi non presenti in tabella CTRL_ASS", result: {}})
        OutputDebug(count . " elementi non presenti in tabella CTRL_ASS")

        ; se tutte le FL sono presenti nella tabella di controllo SAP 
        arrResult := []
        if (count = 0) {
            EventManager.Publish("AddLV", {icon: "icon1", element: "", text: "Tabella control asset aggiornata"})
        }
        else {
            ; Memorizzo la lista delle FL non presenti nelle tabbelle CTRL_ASS in un file
            for element, data in mapControlAsset {
                if (data.check = "NOK") {
                    arrResult.Push(data.conc)
                }
            }
        }     

        Check_CTRL_ASS_Table_result.success := true
        if (arrResult.Length > 0) {
            ; Scrivo il contenuto dell'array in un file.
            SAP_ControlAsset.WriteArrayToFile(arrResult,  A_ScriptDir . "\Check_CTRL_ASS_Table_slow.txt")
            Check_CTRL_ASS_Table_result.value := arrResult
        }

        EventManager.Publish("ProcessCompleted", {processId: Check_CTRL_ASS_Table_result.function, status: "Completed", details: "Esecuzione completata con successo", result: Check_CTRL_ASS_Table_result.value})
        EventManager.Publish("PI_Stop", {inputValue: "Verifica tabella Control Asset completata"}) ; Ferma l'indicatore di progresso
        return Check_CTRL_ASS_Table_result
    }

    ; Funzione: Check_CTRL_ASS_Table
    ; Descrizione: Confronta gli elementi presenti nel Map costruito a partire dalle FL con gli elementi di pari lunghezza presenti in SAP 
    ; Parametri:
    ;   - param1: un array contenente i valori presenti in SAP
    ;   - param2: un map costruito a partire dalle FL inserite nel tool e avente come chiave il codice della FL e come valore un oggetto {lunghezza: numberOfElement, conc: concatena, check: ""}
    ; Restituisce:  Un array costituito dagli elementi non presenti in SAP
    ; Esempio:
    ; Risultato:
    static Check_CTRL_ASS_Table(arr_SAP_CTRL_ASS_Table, map_FL) {    
        ; costruisco un map contenente gli elementi della tabella SAP in cui la chiave identifica la lunghezza della FL mentre il valore è un array contenente tutte le FL di quella lunghezza
        Check_CTRL_ASS_Table_result := { success: false, value: false, error: "", class: "SAP_ControlAsset.ahk", function: "Check_CTRL_ASS_Table" }
        EventManager.Publish("ProcessStarted", {processId: Check_CTRL_ASS_Table_result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica tabella CTRL_ASS "}) ; Avvia l'indicatore di progresso         
        map_SAP_CTRL_ASS := map()
        Count := 0
        if (arr_SAP_CTRL_ASS_Table.Length = 0) or (map_FL.count = 0) {
            msgbox("Errore l'array indicato è vuoto!", "Errore", 4112)
            Check_CTRL_ASS_Table_result.error := "Errore parametri non validi"
            return Check_CTRL_ASS_Table_result  ; Se l'array è vuoto allora restituisco un errore
        }  
        else {
            ; creo una struttura map con chiave la lunghezza della FL e come valore un array contenente tutti gli elementi di quella lunghezza    
            for element in arr_SAP_CTRL_ASS_Table { ; considero ogni elemento della FL
                ; Considero l'ultimo carattere per ottenere la lunghezza della FL
                FL_lenght := SubStr(element, -1)
                ;OutputDebug("Elemento: " . element . " - lunghezza: " . dashCount "`n")
                ; Se non esiste ancora un array per questo conteggio, crealo
                if (!map_SAP_CTRL_ASS.Has(FL_lenght))
                    map_SAP_CTRL_ASS[FL_lenght] := []
                ; Aggiungi la linea all'array corrispondente
                map_SAP_CTRL_ASS[FL_lenght].Push(element)
            }
            ; verifico la presenza degli elementi contenuti nel map_FL nel map_SAP_CTRL_ASS
            for key, value in map_FL {
                value.check := false
                if !(map_SAP_CTRL_ASS.Has(string(value.lunghezza))) {
                    OutputDebug(value.lunghezza . " - Lunghezza non presente nella SAP CTRL_ASS")
                    continue   
                }                      
                for item in map_SAP_CTRL_ASS[string(value.lunghezza)] { ; scansionon il contenuto dell'array
                    if (value.conc = item) {
                        ;OutputDebug("CTRL_ASS: " . item " = map_FL: " value.conc)
                        value.check := true
                        break
                    }
                }
            }
            ; creo un array con i soli elementi non presenti nella SAP_CTRL_ASS
            arrResult := []
            for key, value in map_FL {
                if (value.check = false) {
                    arrResult.Push(value.conc)
                    EventManager.Publish("AddLV", {icon: "icon3", element: key, text: "Non presente in control asset table"})
                }
            }
            if (arrResult.Length = 0) {
                EventManager.Publish("AddLV", {icon: "icon1", element: "", text: "Tabella control asset aggiornata"})
                EventManager.Publish("ProcessProgress", {processId: Check_CTRL_ASS_Table_result.function, status: "In Progress", details: "Tutti gli elementi sono presenti nella tabella CTRL_ASS", result: {}})
                OutputDebug("Tutti gli elementi sono presenti nella tabella CTRL_ASS")
            }
            else {
                EventManager.Publish("ProcessProgress", {processId: Check_CTRL_ASS_Table_result.function, status: "In Progress", details: arrResult.Length . " elementi non presenti in tabella CTRL_ASS", result: {}})
                OutputDebug(arrResult.Length . " elementi NON presenti in tabella CTRL_ASS")
                ; Scrivo il contenuto dell'array in un file.
                SAP_ControlAsset.WriteArrayToFile(arrResult,  A_ScriptDir . "\Check_CTRL_ASS_Table.txt")
                Check_CTRL_ASS_Table_result.value := arrResult
            }        
        }      
        Check_CTRL_ASS_Table_result.success := true
        EventManager.Publish("ProcessCompleted", {processId: Check_CTRL_ASS_Table_result.function, status: "Completed", details: "Esecuzione completata con successo", result: Check_CTRL_ASS_Table_result.value})
        EventManager.Publish("PI_Stop", {inputValue: "Verifica tabella Control Asset completata"}) ; Ferma l'indicatore di progresso
        return Check_CTRL_ASS_Table_result  
    }

    ; Funzione: FilterTabByTechnology
    ; Descrizione: Crea una nuova tabella contenente solo gli elementi in base alla tecnologia
    ; Parametri:
    ;   - param1: La tabella di partenza
    ;   - param2: La tecnologia utilizzata
    ;   - param3: La colonna contenente il valore StructureIndicator composto da 5 caratteri (Z-RES) dobbiamo prendere il 4 carattere.
    ;   - param4: La colonna contenente il valore FL_Category
    ; Restituisce:  Una tabella filtrata 
    ; Esempio:
    ; Risultato:
    static FilterTabByTechnology(CTRL_ASS_Table, fltechnology, Index_StructureIndicator, Index_Valore_FL_Category) {
        Table_Filtered := []
        for rows in CTRL_ASS_Table {
            Table_row := A_index
            ; copio la riga di intestazione
            if(Table_row = 1) {
                Table_Filtered.Push(CTRL_ASS_Table[1])
                continue
            }            
            ; Verifico il codice StructureIndicator               
            code_StructureIndicator := trim(CTRL_ASS_Table[Table_row][Index_StructureIndicator]) ; cerco in tutta la tabella
            code_FL_Category := trim(CTRL_ASS_Table[Table_row][Index_Valore_FL_Category]) ; cerco in tutta la tabella
            ; verifico che il codice lunghezza sia composto da 5 caratteri
            if (RegExMatch(code_StructureIndicator, "^Z-R[SWBGHJE][MS]$")) 
                and (RegExMatch(code_FL_Category, "^[SWBGHJE]$"))
                and (RegExMatch(fltechnology, "^[SWBGHJE]$")) {
                    T_Code := (subStr(code_StructureIndicator, 4 , 1))
                    if ((T_Code = fltechnology) and (code_FL_Category = fltechnology)) {
                        ;OutputDebug("Check code - OK")
                        Table_Filtered.Push(CTRL_ASS_Table[Table_row])
                    }                
                }

        }
        return Table_Filtered
    }

    ; Funzione: WriteArrayToFile
    ; Descrizione: Scrive il contenuto di un array in un file di testo
    ; Parametri:
    ;   - param1: l'array di cui scrivere il contenuto
    ;   - param2: Il percorso del file da scrivere
    ; Restituisce:  Un file in formato .txt riportante il contenuto dell'array
    ; Esempio:   WriteArrayToFile(SAP_ConcArr,  A_ScriptDir . "\SAP_CTRL_ASS.txt")   
    ; Risultato:    
    static WriteArrayToFile(arr, filePath) {
        if !(arr) {
            throw Error("array vuoto")
        }
        try {
            ; Apre il file in modalità scrittura, sovrascrivendo il contenuto esistente
            fileObj := FileOpen(filePath, "w")
            
            ; Scrive ogni elemento dell'array in una nuova riga
            for index, value in arr {
                fileObj.WriteLine(value)
            }
            
            ; Chiude il file
            fileObj.Close()
            
            return true  ; Operazione riuscita
        } catch as err {
            MsgBox("Errore durante la scrittura del file: " . err.Message, "Errore", 16)
            return false  ; Operazione fallita
        }
    }

    ; Funzione: MakeConcTable
    ; Descrizione:  Crea a partire da una tabella un array contenente gli elementi delle FL concatenati in base ai seguenti criteri:
    ;               - Se il livello della sede tecnica è 4 -> considera solo gli ultimi 2 elementi + il livello della sede tecnica
    ;               - Se il livello della sede tecnica è 5 oppure 6 allora considera gli ultimi 3 elementi + il livello della sede tecnica
    ; Parametri:
    ;   - param1: tabella su cui eseguire elaborazione
    ;   - param2: indice del livello sede tecnica
    ;   - param3: Valore livello
    ;   - param4: Valore livello superiore_1
    ;   - param4: Valore livello superiore_2
    ; Restituisce: Un array contenente gli elementi delle FL concatenati + il livello della sede tecnica
    static MakeConcTable(table, Index_LivelloSedeTecnica, Index_Valore_Livello, Index_Valore_Liv_Superiore_1, Index_Valore_Liv_Superiore_2) {
        ConcArr := []
        for rows in table {  
            row := A_index
                if (trim(table[row][Index_LivelloSedeTecnica]) == "4") {
                    element := trim(table[row][Index_Valore_Livello])
                                . "_" . trim(table[row][Index_Valore_Liv_Superiore_1]) 
                                . "_" . trim(table[row][Index_LivelloSedeTecnica])

                    ConcArr.push(element)
                }
                else if ((trim(table[row][Index_LivelloSedeTecnica]) == "5") or (trim(table[row][Index_LivelloSedeTecnica]) == "6")) {
                    element := trim(table[row][Index_Valore_Livello]) 
                                . "_" . trim(table[row][Index_Valore_Liv_Superiore_1]) 
                                . "_" . trim(table[row][Index_Valore_Liv_Superiore_2]) 
                                . "_" . trim(table[row][Index_LivelloSedeTecnica])

                    ConcArr.push(element)
                }
        }
        return ConcArr
    }

    ; Funzione: TableHasIndex
    ; Descrizione: Verifica se un intesatazione è contenuta in una tabella. Se una intestazioni compare più volte, 
    ; la funzione memorizza i valori ricercati e se già trovati cerca un nuova occorrenza.
    ; Parametri:
    ;   - param1: tabella in cui ricercare l'intestazione
    ;   - param2: intestazione da ricercare
    ; Restituisce: L'indice dell' array in cui è contenuto l' elemento
    ; Esempio:      TableHasIndex(CTRL_ASS_Table, "Liv.Sede")   
    ; Risultato:    8
    static TableHasIndex(table:=false, index:=false) {
        ; dato che la tabella ha delle intestazioni con lo stesso nome, memorizzo i valori già trovati in un array in modo da non 
        ; restituirli di nuovo
        static arr := []
        ; se eseguo la funzione senza argomenti resetto l'array 
        if ((table = false) and (index = false)) {
            arr := []
            return false
        }
            
        for rows in table {
            row := A_index
            if (row = 1) { ; cerco nella riga di intestazione
                for element in rows {
                    column := A_index
                        ; Se trovo ricerco la stessa intestazione verifico che non abbia già memorizzato il valore
                        if(table[row][column] = index) and !(SAP_ControlAsset.ArrayHasValue(arr, column)) {
                            arr.Push(column)
                            return column
                        }
                }
            }
            else if (row > 1) ; se termino la riga di intestazione senza trovare il valore allora 
                return false

        }
    }

    ; Funzione: ArrayHasValue
    ; Descrizione: Verifica se un elemento è un elemento è contentuto in un array.
    ; Parametri:
    ;   - param1: array in cui ricercare l' elemento
    ;   - param2: valore da ricercare.
    ; Restituisce: L'indice dell' array in cui è contenuto l' elemento
    ; Esempio:      arr = ["ITS", "0PMS", "00", "MF"]
    ;               ArrayHasValue(arr, "00")       
    ; Risultato:    3
    static ArrayHasValue(arr, Value) {
            for element in arr {
                if (element = Value)
                    return A_Index
            }
        return false
    }

    ; Funzione: TableHasValue
    ; Descrizione:  Ricerca dati all'interno di una tabella, in base ai parametri forniti gestisce diversi tipi di ricerca.
    ;               1) (row = 0) and (column = 0) and (value != "") -> ricerca un valore nell' intera tabella
    ;               2) (row = 0) and (column != 0) and (value != "") -> cerca il valore solo nella colonna indicata e in tutte le righe
    ;               3) (row != 0) and (column != 0) and (value = "") -> restituisca il valore della cella indicata
    ;               4) (row != 0) and (column = 0) and (value != "") -> cerco il valore nella riga indicata in tutte le colonne
    ; Parametri:
    ;   - param1: la tabella su cui eseguire le ricerche
    ;   - param2: il numero di riga in cui ricercare l'elemento
    ;   - param3: il numero di colonna in cui ricercare l'elemento
    ;   - param4: il valore da ricercare
    ; Restituisce: un oggetto composto da result := {column: 0, rows: 0, value: ""}
    static TableHasValue(table, row:=0, column:=0, value:="") {
        result := {column: 0, rows: 0, value: ""}
        if (row = 0) and (column = 0) and (value != "") { ; se sono entrambi 0 allora ricerco il valore in tutta la tabella
            ; e restituisco riga colonna e valore se esistono
            for rows in table {
                Table_row := A_index
                for element in rows {
                    Table_column := A_index                    
                    if(table[Table_row][Table_column] = Trim(value)) { ; cerco in tutta la tabella
                        result := {column: Table_column, rows: Table_row, value: value}
                    }
                }
            }
        }
        else if (row = 0) and (column != 0) and (value != "") { ; cerco il valore solo nella colonna indicata in tutte le righe
            for rows in table {
                Table_row := A_index
                if(trim(table[Table_row][column]) = Trim(value)) { ; scansiono le righe della tabella mantenendo la colonna costante
                    result := {column: column, rows: Table_row, value: value}
                    break ; se trovo il valore termino il ciclo
                }
            }
        }
        else if (row != 0) and (column != 0) and (value = "") { ; restituisco il valore della cella indicata
                value := trim(table[row][column])
                result := {column: column, rows: row, value: value}
            }
        else if (row != 0) and (column = 0) and (value != "") { ; cerco il valore nella riga indicata in tutte le colonne
            columnCount := table[1].length
            loop columnCount {  
                    Table_column := A_index                
                    if(table[row][Table_column] = Trim(value)) { 
                        result := {column: Table_column, rows: row, value: value}
                        break ; se trovo il valore interrompo la ricerca
                    }
                }
        }
        return result
    }
        
    ; Funzione: Concatena
    ; Descrizione: Crea una stringa contenente tutti gli elementi pressenti nell'array concatenati
    ; Parametri:
    ;   - param1: array contenente gli elementi da concatenare
    ;   - param2: il numero di elementi da concatenare a partire dall'ultimo
    ; Restituisce: una stringa composta dagli elementi indicati + la lunghezza dell' array
    ; Esempio:      arr = ["ITS", "0PMS", "00", "MF"]
    ;               Concatena(arr, 3)        
    ; Risultato:    MF_00_4
    static Concatena(arr, numElementi) {
        risultato := ""
        lunghezzaArray := arr.Length
        
        ; Assicuriamoci che numElementi non sia maggiore della lunghezza dell'array
        numElementi := Min(numElementi, lunghezzaArray)
        
        ; Iteriamo dall'ultimo elemento verso il primo
        Loop numElementi {
            indice := lunghezzaArray - A_Index + 1
            
            ; Aggiungiamo l'elemento al risultato
            risultato .= arr[indice]
            
            ; Aggiungiamo il separatore appropriato
            if (A_Index < numElementi) {
                risultato .= "_"
            } else if (numElementi == 1) {
                ;return risultato
            }
        }
        
        return risultato . "_" . lunghezzaArray
    }

    ; Funzione: EstraiTableControlAsset
    ; Descrizione: Estrae i dati relativi alla tabella Control Asset in SAP utilizzando la transazione ZPMR_CTRL_ASS
    ; Parametri: Nessuno
    ; Restituisce: Copia la tabella nella clipboard
    Static EstraiTableControlAsset(fltechnology) {
        ; avvio una sessione SAP
        session := SAPConnection.GetSession()
        if (session) {
            try {
                Temp_Clipboard := A_Clipboard ; memorizzo il contenuto della clipboard
                A_Clipboard := ""
                session.findById("wnd[0]/tbar[0]/okcd").text := "/nZPMR_CTRL_ASS"
                session.findById("wnd[0]").sendVKey(0)
                ; filtro in base alla tecnologia
                session.findById("wnd[0]/usr/txtI4-LOW").text := "Z-R" . fltechnology . "S"
                session.findById("wnd[0]/usr/txtI5-LOW").text := fltechnology
                ; modifico il numero massimo di risultati
                session.findById("wnd[0]/usr/txtMAX_SEL").text := "9999999"
                session.findById("wnd[0]").sendVKey(0)
                ; avvio la transazione
                session.findById("wnd[0]/tbar[1]/btn[8]").press
                while session.Busy()
                    {
                        sleep 500
                        OutputDebug("SAP is busy" . "`n")
                    }
                ; esporto i valori nella clipboard
                session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]").select
                while session.Busy()
                    {
                        sleep 500
                        OutputDebug("SAP is busy" . "`n")
                    }                
                session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select
                session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                while session.Busy()
                    {
                        sleep 500
                        OutputDebug("SAP is busy" . "`n")
                    } 
                if !ClipWait(5) ; attendo fino a 5 secondi per verificare che i dati siano copiati nella clipboard
                    {
                        throw Error("Errore nella scrittura della clipBoard.")
                    }
                ; memorizzo il contenuto della clipboard in un' array
                resultArray := SAP_ControlAsset.CreaArrayDaClipboard()
                ; conto il numero di elementi per confrontarlo con quello dei valori presenti in tabella
                numeroDiElementArray := resultArray.Length
                grid := session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
                ; verifico il numero di risultati ottenuti
                tableRowCount := grid.RowCount + 1 ; aggiungo la riga di intestazione che non viene conteggiata
                OutputDebug("Numero di elementi array: " . numeroDiElementArray . "`n")
                OutputDebug("Numero di righe tabella SAP: " . tableRowCount . "`n")
                if (numeroDiElementArray != tableRowCount)
                    throw Error("Errore nell'estrazione dei dati.")
                else
                    return resultArray   
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

    ; Funzione: CreaArrayDaClipboard
    ; Descrizione: Crea un array a partire dal contenuto della Clipboard
    ; Parametri:Nessuno
    ; Restituisce: un'array contenente la tabella degli asset estratta da SAP, vengono eliminate le righe contenenti "---------------"
    static CreaArrayDaClipboard() {
        try {
            ; Legge il contenuto del file
            clipboardContent := A_Clipboard
            ; Divide il contenuto in linee
            lines := StrSplit(clipboardContent, "`n", "`r")
            ; Inizializza un array per i codici FL
            CTRL_ASS := [] ; ogni elemento dell'array è un array contenente gli elementi della riga
            ; Estrae i codici paese
            for line in lines {
                if (line != "") and !InStr(line, "-----------") { ; rimuovo le righe vuote e le righe composte da trattini
                    parts := StrSplit(line, "|")
                    if (parts.Length >= 2) { ; da verificare che tutte le FL siano costituite da più campi
                        CTRL_ASS.Push(parts)
                    }
                    else {
                        MsgBox("Errore nel contenuto della clipBoard.", "Errore", 4112)
                        return false
                    }
                }
            }
            return CTRL_ASS
        } catch Error as err {
            MsgBox("Errore nel contenuto della clipBoard. - " . err.Message, "Errore", 4112)
            return false
        }
    }
    ; Funzione: CreaArrayDaFile
    ; Descrizione: Crea un array a partire dal contenuto di un file
    ; Parametri:    filePath -> il nome del file da leggere
    ; Restituisce: un'array contenente la tabella degli asset estratta da SAP, vengono eliminate le righe contenenti "---------------"
    static CreaArrayDaFile(filePath) {
        try {
            ; Verifica che il file esista
            if !FileExist(filePath) {
                MsgBox("Il file non esiste: " . filePath, "Errore", 4112)
                return false
            }
    
            ; Legge il contenuto del file
            fileContent := FileRead(filePath)
            if (fileContent = "") {
                MsgBox("Il file è vuoto: " . filePath, "Errore", 4112)
                return false
            }
    
            ; Divide il contenuto in linee
            lines := StrSplit(fileContent, "`n", "`r")
            
            ; Inizializza un array per i codici FL
            CTRL_ASS := []
    
            ; Estrae i codici paese
            for line in lines {
                if (line != "") and !InStr(line, "-----------") { ; rimuovo le righe vuote e le righe composte da trattini
                    parts := StrSplit(line, "|")
                    if (parts.Length >= 2) { ; da verificare che tutte le FL siano associate ad una descrizione
                        CTRL_ASS.Push(parts)
                    }
                    else {
                        MsgBox("Errore nella struttura del file. Riga non valida: " . line, "Errore", 4112)
                        return false
                    }
                }
            }
    
            if (CTRL_ASS.Length = 0) {
                MsgBox("Nessun dato valido trovato nel file.", "Errore", 4112)
                return false
            }
    
            return CTRL_ASS
    
        } catch Error as err {
            MsgBox("Errore durante la lettura del file: " . err.Message, "Errore", 4112)
            return false
        }
    }

}