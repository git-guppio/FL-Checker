#Requires AutoHotkey v2.0

#Include SAP_Connection.ahk

class SAP_TechnicalObject {

    static __New() {
        SAP_TechnicalObject.InitializeVariables()
        SAP_TechnicalObject.SetupEventListeners()
    }

    Static InitializeVariables() {
        SAP_TechnicalObject.DebugMode := 0 ; 0 -> leggo la tabella da SAP; 1 -> leggo la tabella da file
    }

    Static SetupEventListeners() {
        EventManager.Subscribe("VerificaTechnicalObject", (data) => SAP_TechnicalObject.VerificaTechnicalObject(data.flArray, data.flcountry, data.fltechnology))
    }

    ; Metodo: VerificaTechnicalObject
    ; Descrizione: Esegue la verifica degli elementi delle FL nelle tabelle SAP dei technical object
    ; Parametri:
    ; Restituisce: 
    ;   - un array contenente gli elementi non presenti nella TECH_OBJ se esistono
    ;   - un array vuoto altrimenti
    ; Esempio:
    static VerificaTechnicalObject(flArray, flcountry, fltechnology) {
        VerificaTechnicalObject_result := { success: false, value: false, error: "", class: "SAP_TechnicalObject.ahk", function: "VerificaTechnicalObject" }
        EventManager.Publish("ProcessStarted", {processId: VerificaTechnicalObject_result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica tecnical object"}) ; Avvia l'indicatore di progresso         
        
        mapControlAsset := map()
        data := {lunghezza: 0, conc: ""}
        Lista_FL_Conc := []
        ; analizzo gli elementi presenti nell'array FL
        EventManager.Publish("ProcessProgress", {processId: VerificaTechnicalObject_result.function, status: "In Progress", details: "Analizza elementi FL", result: {}})
        for element in flArray {
            temp_array := StrSplit(element, "-")
            numberOfElement := temp_array.Length
            ; i valori da controllare sono solo quelli >= 4
            if (numberOfElement >= 4) {
                ;OutputDebug("element: " . element . " - temp_array.Length: " . temp_array.Length . "`n")
                switch numberOfElement
                {
                case 4: concatena := SAP_TechnicalObject.Concatena(temp_array, 2)
                case 5: concatena := SAP_TechnicalObject.Concatena(temp_array, 3)
                case 6: concatena := SAP_TechnicalObject.Concatena(temp_array, 3)
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
            SAP_TechnicalObject.WriteArrayToFile(Lista_FL_Conc,  A_ScriptDir . "\Lista_FL_Conc.txt")
            OutputDebug("Elementi concatenati a partire dalle FL: " . Lista_FL_Conc.Length . "`n")    
        }
        else { ; se non contiene elementi concludo 
            VerificaTechnicalObject_result.error := "Nessun elemento da verificare"
            return VerificaTechnicalObject_result
        }

        if (SAP_TechnicalObject.DebugMode = 0) { ; 0 -> leggo la tabella da SAP; 1 -> leggo la tabella da file
            OutputDebug("Prelevo i dati da SAP `n")
            ; estraggo i dati dalla tabella SAP filtrando in base alla tecnologia
            EventManager.Publish("ProcessProgress", {processId: VerificaTechnicalObject_result.function, status: "In Progress", details: "Estraggo tabella TECH_OBJ da SAP", result: {}})
            TECH_OBJ_Table := SAP_TechnicalObject.EstraiTableTechnicalObject(fltechnology) ; restituisce un array contenente un array per ogni per ogni riga della tabella.
            if !(TECH_OBJ_Table) {
                VerificaTechnicalObject_result.error := "Errore estrazione tabella TECH_OBJ"
                EventManager.Publish("ProcessError", {processId: VerificaTechnicalObject_result.function, status: "Error", details: VerificaTechnicalObject_result.error . " - Line Number: " . A_LineNumber, result: {}})                   
                return VerificaTechnicalObject_result                
            }

            OutputDebug("-- Verifico intestazione tabella TECH_OBJ --`n")
                for element in TECH_OBJ_Table[1] {
                    OutputDebug(element . "`t")            
                }
                OutputDebug("`n")
                        
            ; ricerco i valori delle intestazioni delle colonne nel file
            Index_TECH_OBJ_LivelloSedeTecnica := SAP_TechnicalObject.TableHasIndex(TECH_OBJ_Table, "Liv.Sede")
            Index_TECH_OBJ_Valore_Livello := SAP_TechnicalObject.TableHasIndex(TECH_OBJ_Table, "Valore Livello")
            Index_TECH_OBJ_Valore_Liv_Superiore_1 := SAP_TechnicalObject.TableHasIndex(TECH_OBJ_Table, "Valore Liv. Superiore")
            Index_TECH_OBJ_Valore_Liv_Superiore_2 := SAP_TechnicalObject.TableHasIndex(TECH_OBJ_Table, "Valore Liv. Superiore")
            Index_TECH_OBJ_Valore_StructureIndicator := SAP_TechnicalObject.TableHasIndex(TECH_OBJ_Table, "Str. ")
            Index_TECH_OBJ_Valore_FL_Category := SAP_TechnicalObject.TableHasIndex(TECH_OBJ_Table, "C")

        }
        else if (SAP_TechnicalObject.DebugMode = 1) { ; 0 -> leggo la tabella da SAP; 1 -> leggo la tabella da file
            OutputDebug("Prelevo i dati da file")
            ; estraggo i dati da file
            EventManager.Publish("ProcessProgress", {processId: VerificaTechnicalObject_result.function, status: "In Progress", details: "Leggo file tabella TECH_OBJ", result: {}})
            ;TECH_OBJ_Table := SAP_TechnicalObject.CreaArrayDaFile("C:\Users\a259046\OneDrive - Enel Spa\SCRIPT AHK e VBA\Functional_Location\CheckFL_rev.3\SAP\ZMPR_TECH_OBJ.txt")
            TECH_OBJ_Table := SAP_TechnicalObject.CreaArrayDaFile("C:\Users\a259046\OneDrive - Enel Spa\SCRIPT AHK e VBA\Functional_Location\CheckFL_rev.3\SAP\ZMPR_TECH_OBJ.txt")
            if !(TECH_OBJ_Table) {
                VerificaTechnicalObject_result.error := "Errore lettura file tabella TECH_OBJ"
                EventManager.Publish("ProcessError", {processId: VerificaTechnicalObject_result.function, status: "Error", details: VerificaTechnicalObject_result.error . " - Line Number: " . A_LineNumber, result: {}})                   
                EventManager.Publish("PI_Stop", {inputValue: "Errore verifica technical object"}) ; Ferma l'indicatore di progresso
                return VerificaTechnicalObject_result                
            }    
                     
            OutputDebug("-- Verifico intestazione tabella TECH_OBJ --`n")
                for element in TECH_OBJ_Table[1] {
                    OutputDebug(element . "`t")            
                }
                OutputDebug("`n")

            ; ricerco i valori delle intestazioni delle colonne nel file
            Index_TECH_OBJ_LivelloSedeTecnica := SAP_TechnicalObject.TableHasIndex(TECH_OBJ_Table, "Liv.Sede")
            Index_TECH_OBJ_Valore_Livello := SAP_TechnicalObject.TableHasIndex(TECH_OBJ_Table, "Valore Livello")
            Index_TECH_OBJ_Valore_Liv_Superiore_1 := SAP_TechnicalObject.TableHasIndex(TECH_OBJ_Table, "Valore Liv. Superiore")
            Index_TECH_OBJ_Valore_Liv_Superiore_2 := SAP_TechnicalObject.TableHasIndex(TECH_OBJ_Table, "Valore Liv. Superiore")
            Index_TECH_OBJ_Valore_StructureIndicator := SAP_TechnicalObject.TableHasIndex(TECH_OBJ_Table, "Str. ")
            Index_TECH_OBJ_Valore_FL_Category := SAP_TechnicalObject.TableHasIndex(TECH_OBJ_Table, "C")

            ; Filtro la tabella in base al tipo di tecnologia.
            ; Esamino i campi Index_TECH_OBJ_Valore_FL_Category e Index_TECH_OBJ_Valore_StructureIndicator
            EventManager.Publish("ProcessProgress", {processId: VerificaTechnicalObject_result.function, status: "In Progress", details: "Filtro tabella TECH_OBJ in base alla tecnologia", result: {}})
            TECH_OBJ_Table := SAP_TechnicalObject.FilterTabByTechnology(TECH_OBJ_Table, fltechnology, Index_TECH_OBJ_Valore_StructureIndicator, Index_TECH_OBJ_Valore_FL_Category)
        }

        OutputDebug("-- INDICI DELLA COLONNE PRESENTI NEL FILE --`n")
        OutputDebug("Index_TECH_OBJ_LivelloSedeTecnica " . Index_TECH_OBJ_LivelloSedeTecnica . "`n")
        OutputDebug("Index_TECH_OBJ_Valore_Livello " . Index_TECH_OBJ_Valore_Livello . "`n")
        OutputDebug("Index_TECH_OBJ_Valore_Liv_Superiore_1 " . Index_TECH_OBJ_Valore_Liv_Superiore_1 . "`n")
        OutputDebug("Index_TECH_OBJ_Valore_Liv_Superiore_2 " . Index_TECH_OBJ_Valore_Liv_Superiore_2 . "`n")
        OutputDebug("Index_TECH_OBJ_FL_Category " . Index_TECH_OBJ_Valore_FL_Category . "`n")
        OutputDebug("Index_TECH_OBJ_StructureIndicator " . Index_TECH_OBJ_Valore_StructureIndicator . "`n")        
        
        ; verifico che le intestazioni cercate abbiano un valore
        if !(Index_TECH_OBJ_LivelloSedeTecnica 
            AND Index_TECH_OBJ_Valore_Livello 
            AND  Index_TECH_OBJ_Valore_Liv_Superiore_1 
            AND Index_TECH_OBJ_Valore_Liv_Superiore_2
            AND  Index_TECH_OBJ_Valore_FL_Category 
            AND Index_TECH_OBJ_Valore_StructureIndicator) {
                VerificaTechnicalObject_result.error := "Errore intestazioni tabella TECH_OBJ"
                EventManager.Publish("ProcessError", {processId: VerificaTechnicalObject_result.function, status: "Error", details: VerificaTechnicalObject_result.error . " - Line Number: " . A_LineNumber, result: {}})                   
                EventManager.Publish("PI_Stop", {inputValue: "Errore verifica technical object"}) ; Ferma l'indicatore di progresso
                return VerificaTechnicalObject_result
            }

        ; La tabella TECH_OBJ ha lo stesso nome per due intestazioni, verifico quella corretta
        ; La colonna indicata come <Index_TECH_OBJ_Valore_Liv_Superiore_2> deve essere vuota per il livello sede tecnica = 4
        ; result := {column: 0, rows: 0, value: ""}
        Data := SAP_TechnicalObject.TableHasValue(TECH_OBJ_Table, , column:=Index_TECH_OBJ_LivelloSedeTecnica, value:= "4") ; verifico in quale riga è contenuto il valore 4 per la colonna <Liv.Sede>
        if (Data.rows != 0) { ; ricavo la riga che contiene il valore 4
            ; verifico se Index_TECH_OBJ_Valore_Liv_Superiore_2 contiene un valore vuoto
            Data := SAP_TechnicalObject.TableHasValue(TECH_OBJ_Table, Data.rows, column:=Index_TECH_OBJ_Valore_Liv_Superiore_2, )
            if (Data.value != "") { ; allora il valore non è corretto, devo invertire gli indici
                temp := Index_TECH_OBJ_Valore_Liv_Superiore_1
                Index_TECH_OBJ_Valore_Liv_Superiore_1 := Index_TECH_OBJ_Valore_Liv_Superiore_2
                Index_TECH_OBJ_Valore_Liv_Superiore_2 := temp
            }
        }

        EventManager.Publish("ProcessProgress", {processId: VerificaTechnicalObject_result.function, status: "In Progress", details: "Concateno valori per effettuare controllo", result: {}})
        ; a partire dalla tabella estratta da SAP o letta da file e filtrata creo un array con i valori concatenati
        SAP_ConcArr := SAP_TechnicalObject.MakeConcTable(  TECH_OBJ_Table, 
                                                        Index_TECH_OBJ_LivelloSedeTecnica, 
                                                        Index_TECH_OBJ_Valore_Livello, 
                                                        Index_TECH_OBJ_Valore_Liv_Superiore_1, 
                                                        Index_TECH_OBJ_Valore_Liv_Superiore_2)
        
        OutputDebug(SAP_ConcArr.Length . " elementi in Tabella SAP_TECH_OBJ`n")    
        ; *** Scrivo il contenuto dell'array in un file.
        SAP_TechnicalObject.WriteArrayToFile(SAP_ConcArr,  A_ScriptDir . "\SAP_TECH_OBJ_" . fltechnology . ".txt")

       ; temp_result_1 := SAP_TechnicalObject.Check_TECH_OBJ_Table_slow(SAP_ConcArr, mapControlAsset)

        temp_result := SAP_TechnicalObject.Check_TECH_OBJ_Table(SAP_ConcArr, mapControlAsset)

        EventManager.Publish("ProcessCompleted", {processId: VerificaTechnicalObject_result.function, status: "Completed", details: "Esecuzione completata con successo", result: temp_result})
        EventManager.Publish("PI_Stop", {inputValue: "Verifica tabella technical object completata"}) ; Ferma l'indicatore di progresso

        VerificaTechnicalObject_result.success := true
        VerificaTechnicalObject_result.value := temp_result
        return VerificaTechnicalObject_result
    }

    ; Funzione: Check_TECH_OBJ_Table
    ; Descrizione: Confronta gli elementi presenti nel Map costruito a partire dalle FL con gli elementi di pari lunghezza presenti in SAP 
    ; Parametri:
    ;   - param1: un array contenente i valori presenti in SAP
    ;   - param2: un map costruito a partire dalle FL inserite nel tool e avente come chiave il codice della FL e come valore un oggetto {lunghezza: numberOfElement, conc: concatena, check: ""}
    ; Restituisce:  Un array costituito dagli elementi non presenti in SAP
    ; Esempio:
    ; Risultato:
    static Check_TECH_OBJ_Table(arr_SAP_TECH_OBJ_Table, map_FL) {    
        ; costruisco un map contenente gli elementi della tabella SAP in cui la chiave identifica la lunghezza della FL mentre il valore è un array contenente tutte le FL di quella lunghezza
        Check_TECH_OBJ_Table_result := { success: false, value: false, error: "", class: "SAP_TechnicalObject.ahk", function: "Check_TECH_OBJ_Table" }
        EventManager.Publish("ProcessStarted", {processId: Check_TECH_OBJ_Table_result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica tabella TECH_OBJ "}) ; Avvia l'indicatore di progresso         
        map_SAP_TECH_OBJ := map()
        Count := 0
        if (arr_SAP_TECH_OBJ_Table.Length = 0) or (map_FL.count = 0) {
            msgbox("Errore l'array indicato è vuoto!", "Errore", 4112)
            Check_TECH_OBJ_Table_result.error := "Errore parametri non validi"
            return Check_TECH_OBJ_Table_result  ; Se l'array è vuoto allora restituisco un errore
        }  
        else {
            ; creo una struttura map con chiave la lunghezza della FL e come valore un array contenente tutti gli elementi di quella lunghezza    
            for element in arr_SAP_TECH_OBJ_Table { ; considero ogni elemento della FL
                ; Considero l'ultimo carattere per ottenere la lunghezza della FL
                FL_lenght := SubStr(element, -1)
                ;OutputDebug("Elemento: " . element . " - lunghezza: " . dashCount "`n")
                ; Se non esiste ancora un array per questo conteggio, crealo
                if (!map_SAP_TECH_OBJ.Has(FL_lenght))
                    map_SAP_TECH_OBJ[FL_lenght] := []
                ; Aggiungi la linea all'array corrispondente
                map_SAP_TECH_OBJ[FL_lenght].Push(element)
            }
            ; verifico la presenza degli elementi contenuti nel map_FL nel map_SAP_TECH_OBJ
            for key, value in map_FL {
                value.check := false
                if !(map_SAP_TECH_OBJ.Has(string(value.lunghezza))) {
                    OutputDebug(value.lunghezza . " - Lunghezza non presente nella SAP TECH_OBJ")
                    continue   
                }                      
                for item in map_SAP_TECH_OBJ[string(value.lunghezza)] { ; scansionon il contenuto dell'array
                    if (value.conc = item) {
                        ;OutputDebug("TECH_OBJ: " . item " = map_FL: " value.conc)
                        value.check := true
                        break
                    }
                }
            }
            ; creo un array con i soli elementi non presenti nella SAP_TECH_OBJ
            arrResult := []
            for key, value in map_FL {
                if (value.check = false) {
                    arrResult.Push(value.conc)
                    EventManager.Publish("AddLV", {icon: "icon3", element: value.conc, text: "Non presente in technical object table"})
                }
            }
            if (arrResult.Length = 0) {
                EventManager.Publish("AddLV", {icon: "icon1", element: "", text: "Tabella technical object aggiornata"})
                EventManager.Publish("ProcessProgress", {processId: Check_TECH_OBJ_Table_result.function, status: "In Progress", details: "Tutti gli elementi sono presenti nella tabella TECH_OBJ", result: {}})
                OutputDebug("Tutti gli elementi sono presenti nella tabella TECH_OBJ")
                Check_TECH_OBJ_Table_result.value := arrResult
            }
            else {
                EventManager.Publish("ProcessProgress", {processId: Check_TECH_OBJ_Table_result.function, status: "In Progress", details: arrResult.Length . " elementi non presenti in tabella TECH_OBJ", result: {}})
                OutputDebug(arrResult.Length . " elementi NON presenti in tabella TECH_OBJ `n")
                ; Scrivo il contenuto dell'array in un file.
                SAP_TechnicalObject.WriteArrayToFile(arrResult,  A_ScriptDir . "\Check_TECH_OBJ_Table.txt")
                Check_TECH_OBJ_Table_result.value := arrResult
            }        
        }      
        Check_TECH_OBJ_Table_result.success := true
        EventManager.Publish("ProcessCompleted", {processId: Check_TECH_OBJ_Table_result.function, status: "Completed", details: "Esecuzione completata con successo", result: Check_TECH_OBJ_Table_result})
        EventManager.Publish("PI_Stop", {inputValue: "Verifica tabella technical object completata"}) ; Ferma l'indicatore di progresso
        return Check_TECH_OBJ_Table_result
    }

    ; Funzione: FilterTabByTechnology
    ; Descrizione:  A partire dal file crea una nuova tabella contenente solo gli elementi in base alla tecnologia
    ;               N.B. non viene utilizzato per i dati scaricati da SAP in quanto sono già filtrati nella transazione
    ; Parametri:
    ;   - param1: La tabella di partenza
    ;   - param2: La tecnologia utilizzata
    ;   - param3: La colonna contenente il valore StructureIndicator composto da 5 caratteri (Z-RES) dobbiamo prendere il 4 carattere.
    ;   - param4: La colonna contenente il valore FL_Category
    ; Restituisce:  Una tabella filtrata 
    ; Esempio:
    ; Risultato:
    static FilterTabByTechnology(TECH_OBJ_Table, fltechnology, Index_StructureIndicator, Index_Valore_FL_Category) {
        Table_Filtered := []
        for rows in TECH_OBJ_Table {
            Table_row := A_index
            ; copio la riga di intestazione
            if(Table_row = 1) {
                Table_Filtered.Push(TECH_OBJ_Table[1])
                continue
            }            
            ; Verifico il codice StructureIndicator               
            code_StructureIndicator := trim(TECH_OBJ_Table[Table_row][Index_StructureIndicator]) ; cerco in tutta la tabella
            code_FL_Category := trim(TECH_OBJ_Table[Table_row][Index_Valore_FL_Category]) ; cerco in tutta la tabella
            ; verifico che il codice lunghezza sia composto da 5 caratteri
            if (RegExMatch(code_StructureIndicator, "^Z-R[SWBGHJE][MS]$")) 
                and (RegExMatch(code_FL_Category, "^[SWBGHJE]$"))
                and (RegExMatch(fltechnology, "^[SWBGHJE]$")) {
                    T_Code := (subStr(code_StructureIndicator, 4 , 1))
                    if ((T_Code = fltechnology) and (code_FL_Category = fltechnology)) {
                        ;OutputDebug("Check code - OK")
                        Table_Filtered.Push(TECH_OBJ_Table[Table_row])
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
    ; Esempio:   WriteArrayToFile(SAP_ConcArr,  A_ScriptDir . "\SAP_TECH_OBJ.txt")   
    ; Risultato:    
    static WriteArrayToFile(arr, filePath) {
    try {        
        if !(arr) {
            throw Error("array vuoto")
        }

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
    ; Esempio:      TableHasIndex(TECH_OBJ_Table, "Liv.Sede")   
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
            
        intestazione := table[1] ; array contenente la lista delle intestazioni
        for element in intestazione {
            column := A_index
                ; Ricerco le intestazioni e verificao che non abbia già memorizzato il valore
                if(element = index) and !(SAP_ControlAsset.ArrayHasValue(arr, column)) {
                    arr.Push(column)
                    return column
                }
        }
        ; se termino la riga di intestazione senza trovare il valore allora 
        return false
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

    ; Funzione: EstraiTableTechnicalObject
    ; Descrizione: Estrae i dati relativi alla tabella technical object in SAP utilizzando la transazione ZPMR_TECH_OBJ
    ; Parametri: Nessuno
    ; Restituisce: Copia la tabella nella clipboard
    Static EstraiTableTechnicalObject(fltechnology) {
        result := { success: false, value: false, error: "", class: "SAP_TechObj.ahk", function: "EstraiTableTechnicalObject" }
        EventManager.Publish("ProcessStarted", {processId: result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Estrai tabella CTRL_ASS da SAP"}) ; Avvia l'indicatore di progresso         
        ; avvio una sessione SAP
        session := SAPConnection.GetSession()
        if (session) {
            try {
                Temp_Clipboard := A_Clipboard ; memorizzo il contenuto della clipboard
                EventManager.Publish("ProcessProgress", {processId: result.function, status: "In Progress", details: "Avvio transazione SE16 per tabella ZPM4R_GL_T_FL", result: {}})
                session.findById("wnd[0]/tbar[0]/okcd").text := "/nSE16"
                session.findById("wnd[0]").sendVKey(0)
                session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text := "ZPM4R_GL_T_FL"
                session.findById("wnd[0]").sendVKey(0)
                sleep 500
                ; filtro in base alla tecnologia                
                session.findById("wnd[0]/usr/ctxtI4-LOW").text := "Z-R" . fltechnology . "S"
                session.findById("wnd[0]/usr/ctxtI5-LOW").text := fltechnology
                ; filtro in base al livello della FL
                session.findById("wnd[0]/usr/btn%_I6_%_APP_%-VALU_PUSH").press
                sleep 500
                session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text := "4"
                session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text := "5"
                session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text := "6"
                session.findById("wnd[1]/tbar[0]/btn[8]").press
                sleep 500          
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
                sleep 500    
                EventManager.Publish("ProcessProgress", {processId: result.function, status: "In Progress", details: "Memorizzo valori in clipboard", result: {}})                       
                ; esporto i valori nella clipboard
                session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]").select
                while session.Busy()
                    {
                        sleep 500
                        OutputDebug("SAP is busy" . "`n")
                    }
                sleep 500                          
                session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select
                session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus
                A_Clipboard := ""
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                while session.Busy()
                    {
                        sleep 500
                        OutputDebug("SAP is busy" . "`n")
                    } 
                ; Impostiamo un timeout per evitare loop infiniti
                maxWaitTime := 10000  ; millisecondi (2 secondi)
                startTime := A_TickCount

                ; Ciclo che attende che la clipboard contenga dati
                while (A_Clipboard = "") {
                    if (A_TickCount - startTime > maxWaitTime) {
                        EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: "Timeout: La clipboard non è stata riempita entro " . maxWaitTime . " ms", result: {}})
                        MsgBox("Timeout: La clipboard non è stata riempita entro " . maxWaitTime . " ms")
                        OutputDebug("Timeout: La clipboard non è stata riempita entro " . maxWaitTime . " ms")
                        return false
                    }
                    Sleep(100)  ; Pausa di 50ms per non sovraccaricare la CPU
                }
                EventManager.Publish("ProcessProgress", {processId: result.function, status: "In Progress", details: "Valori memorizzati in clipboard", result: {}})
                OutputDebug("Clipboard riempita con successo!`n")          
                ; memorizzo il contenuto della clipboard in un' array
                resultArray := SAP_TechnicalObject.CreaArrayDaClipboard()
                ; ripristino il contenuto della Clipboard
                A_Clipboard := Temp_Clipboard                
                ; conto il numero di elementi per confrontarlo con quello dei valori presenti in tabella
                numeroDiElementArray := resultArray.Length
                grid := session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
                ; verifico il numero di risultati ottenuti
                tableRowCount := grid.RowCount + 1 ; aggiungo la riga di intestazione che non viene conteggiata
                OutputDebug("Numero di elementi array: " . numeroDiElementArray . "`n")
                OutputDebug("Numero di righe tabella SAP: " . tableRowCount . "`n")
                EventManager.Publish("ProcessProgress", {processId: result.function, status: "In Progress", details: "Check: n. elementi SAP = " .  tableRowCount . " - n. elementi array = " . numeroDiElementArray, result: {}})
                if (numeroDiElementArray != tableRowCount)
                    throw Error("Errore nell'estrazione dei dati.")
                else
                    EventManager.Publish("ProcessCompleted", {processId: result.function, status: "Completeds", details: "Esecuzione completata con successo", result: resultArray})
                    return resultArray   
            } catch as err {
                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: "Errore nell'esecuzione dell'azione SAP: " err.Message, result: {}})
                MsgBox("Errore nell'esecuzione dell'azione SAP: " err.Message, "Errore", 4112)
                return false
            }
        }
        else {
            EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: "Impossibile ottenere una sessione SAP valida.", result: {}})
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
            ; Ottiene il numero di campi dalla riga di intestazione (header) [seconda riga]
            expectedFields := StrSplit(lines[2], "|").Length
            ; Inizializza un array per i codici FL
            TECH_OBJ := [] ; ogni elemento dell'array è un array contenente gli elementi della riga
            ; Estrae i codici paese
            for line in lines {
                if (line != "") and !InStr(line, "-----------") { ; rimuovo le righe vuote e le righe composte da trattini
                    parts := StrSplit(line, "|")
                    if (parts.Length = expectedFields) { ; verifico che tutte le righe siano costituite dallo stesso numero di campi contenuti nell'intestazione
                        TECH_OBJ.Push(parts)
                    }
                    else {
                        MsgBox("Errore nel contenuto della clipBoard.", "Errore", 4112)
                        return false
                    }
                }
            }
            return TECH_OBJ
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
            ; Ottiene il numero di campi contenuti nell'intestazione [riga 2]
            expectedFields := StrSplit(lines[2], "|").Length
            ; Inizializza un array per i codici FL
            TECH_OBJ := []
    
            ; Estrae i codici paese
            for line in lines {
                if (line != "") and !InStr(line, "-----------") { ; rimuovo le righe vuote e le righe composte da trattini
                    parts := StrSplit(line, "|")
                    if (parts.Length = expectedFields) { ; da verificare che tutte le FL siano associate ad una descrizione
                        TECH_OBJ.Push(parts)
                    }
                    else {
                        MsgBox("Errore nella struttura del file. Riga non valida: " . line, "Errore", 4112)
                        return false
                    }
                }
            }
    
            if (TECH_OBJ.Length = 0) {
                MsgBox("Nessun dato valido trovato nel file.", "Errore", 4112)
                return false
            }
    
            return TECH_OBJ
    
        } catch Error as err {
            MsgBox("Errore durante la lettura del file: " . err.Message, "Errore", 4112)
            return false
        }
    }

}