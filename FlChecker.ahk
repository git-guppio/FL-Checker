#Requires AutoHotkey v2.0

; Classe: FLChecker
; Descrizione: Classe per la verifica dei dati inseriti nella GUI
; Esempio:


class FLChecker {

    static status := ""

    static __New() {
        FLChecker.InitializeVariables()
        FLChecker.SetupEventListeners()
    }

    Static InitializeVariables() {
        FLChecker.inverterTechnology := ""
        FLChecker.VerificaFL_SAP := false
        FLChecker.UpLoadFiles_SAP := true

/*         ; Definizione dei percorsi dei file
		G_CONSTANTS.file_Country := A_ScriptDir . "\Config\country.txt"
		G_CONSTANTS.file_Tech := A_ScriptDir . "\Config\Technology.txt"
		G_CONSTANTS.file_Rules := A_ScriptDir . "\Config\Rules.txt"
		G_CONSTANTS.file_Mask := A_ScriptDir . "\Config\Mask_FL.txt"
        ; Guideline
		G_CONSTANTS.file_FL_Wind:= A_ScriptDir . "\Config\Wind_FL_GuideLine.txt"
		G_CONSTANTS.file_FL_Bess:= A_ScriptDir . "\Config\Bess_FL_GuideLine.txt"
		G_CONSTANTS.file_FL_Solar_Common:= A_ScriptDir . "\Config\Solar_FL_Common_GuideLine.txt"
		G_CONSTANTS.file_FL_WSB_SubStation:= A_ScriptDir . "\Config\WSB_FL_SubStation_Guideline.txt"
		G_CONSTANTS.file_FL_Solar_CentralInv:= A_ScriptDir . "\Config\Solar_FL_CentrealInv_GuideLine.txt"
		G_CONSTANTS.file_FL_Solar_StringInv:= A_ScriptDir . "\Config\Solar_FL_StringInv_GuideLine.txt"
		G_CONSTANTS.file_FL_Solar_InvModule:= A_ScriptDir . "\Config\Solar_FL_InvModule_GuideLine.txt"
		G_CONSTANTS.file_FL_2_UpLoad:= A_ScriptDir . "\FileUpLoad\FL_2_UpLoad.csv"
		G_CONSTANTS.file_FL_n_UpLoad:= A_ScriptDir . "\FileUpLoad\FL_n_UpLoad.csv" */
    }

    Static SetupEventListeners() {
        EventManager.Subscribe("CheckDataRequest", (data) => FLChecker.CheckFL(data.flArray))
        EventManager.Subscribe("ProcessStatusUpdated",(data) => FLChecker.ProcessStatusUpdated(data.processId, data.status, data.details, data.result))            
    }

    static ProcessStatusUpdated(processId, status, details := "", result:={}) {
        ; leggo il valore selezionato dalla GUI della classe MakeInvTecGUI
        if (processID = "SelectInverter") and (status = "Completed") {
            FLChecker.inverterTechnology := result.value
            (FLChecker.inverterTechnology >= 0) ? FLChecker.inverterTechnology : 0
            outputdebug("FLChecker.inverterTechnology: " . FLChecker.inverterTechnology . "`n")
        }
        else if (processID = "SelectInverter") and (status != "Completed")
            FLChecker.inverterTechnology := "" 
        
        ; leggo il valore di ritorno dal controllo tabelle globali in SAP
        if (processID = "VerificaFL_SAP") and (status = "Completed") {
            FLChecker.VerificaFL_SAP := 1
            outputdebug("FLChecker.VerificaFL_SAP: " . FLChecker.VerificaFL_SAP . "`n")
        }
        else if (processID = "VerificaFL_SAP") and (status = "Error") {
            FLChecker.VerificaFL_SAP := -1
            outputdebug("FLChecker.VerificaFL_SAP: " . FLChecker.VerificaFL_SAP . "`n")
        }
        else if (processID = "VerificaFL_SAP")
            FLChecker.VerificaFL_SAP := 0
        
        ; leggo il valore di ritorno dal caricamento file tabelle globali SAP
        if (processID = "UpLoadFiles") and ((status = "Completed") or (status = "Error")) {
            FLChecker.UpLoadFiles_SAP := result.value   ; può assumere valori 0, 1 o 2 in base a quanti file sono stati caricati
            outputdebug("FLChecker.UpLoadFiles: " . result.value . "`n")
        }
        else if (processID = "UpLoadFiles")
            FLChecker.UpLoadFiles_SAP := -1 
        
        ; leggo il valore di ritorno dal controllo delle CTRL_ASS in SAP
        if (processID = "VerificaControlAsset") and ((status = "Completed") or (status = "Error")) {
            FLChecker.VerificaControlAsset := result.value   ; può assumere valori 0, 1 o 2 in base a quanti file sono stati caricati
            outputdebug("FLChecker.UpLoadFiles: " . result.value . "`n")
        }
        else if (processID = "UpLoadFiles")
            FLChecker.UpLoadFiles_SAP := -1 

    }

    ; Funzione attivata dal tasto <check> presente nella GUI principale
    ; Legge il contenuto del controllo edit e crea un array eliminando le righe vuote.
    ; memorizza il risultato nell'array <array_FL>
    Static CheckFL(data) {
        CheckFL_result := { success: false, value: false, error: "", class: "FlChecker.ahk", function: "CheckFL" }
        EventManager.Publish("ProcessStarted", {processId: "CheckFL", status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifico FL "})        
        try {
            ; Trasformo il contenuto della FL list in un array e verifico che aderisca alla maschera generica
            CheckData_result := FLChecker.CheckData(data)
            ;result := { success: false, value: false, error: "", class: "FlChecker.ahk", function: "CheckFL" }
            if(CheckData_result.success = false) {
                EventManager.Publish("ProcessError", {processId: "CheckFL", status: "Error", details: CheckData_result.error . " - Line Number: " . A_LineNumber, result: {}})
                EventManager.Publish("PI_Stop", {inputValue: "Errore nel contenuto della lista!"}) ; Ferma l'indicatore di progresso
                throw Error("Contenuto della lista non valido.") 
            }
            ; Eseguo controllo sul codice country
            CheckCountry_result := FLChecker.CheckCountry(CheckData_result.value, G_CONSTANTS.file_Country)
            if (CheckCountry_result.success = false) {
                EventManager.Publish("ProcessError", {processId: "CheckFL", status: "Error", details: CheckCountry_result.error . " - Line Number: " . A_LineNumber, result: {}})
                EventManager.Publish("PI_Stop", {inputValue: "Errore nella verifica della country!"}) ; Ferma l'indicatore di progresso
                throw Error("Errore nella verifica della country.")
            }
            ; Eseguo controllo sul codice tecnologia
            CheckTechnology_result := FLChecker.CheckTechnology(CheckData_result.value, G_CONSTANTS.file_Tech)
            if (CheckTechnology_result.success = false) {
                EventManager.Publish("ProcessError", {processId: "CheckFL", status: "Error", details: CheckTechnology_result.error . " - Line Number: " . A_LineNumber, result: {}})
                EventManager.Publish("PI_Stop", {inputValue: "Errore nella verifica della tecnologia!"}) ; Ferma l'indicatore di progresso
                throw Error("Errore nella verifica della tecnologia.")
            }

            CheckMask_result := FLChecker.CheckMask(CheckData_result.value, CheckTechnology_result.value)
            if (CheckMask_result.success = false) {
                EventManager.Publish("ProcessError", {processId: "CheckMask", status: "Error", details: CheckMask_result.error . " - Line Number: " . A_LineNumber, result: {}})
                EventManager.Publish("PI_Stop", {inputValue: "Errore nella verifica della maschera!"}) ; Ferma l'indicatore di progresso
                throw Error("Errore nella verifica della maschera.")                
            }

            CheckParent_result := FLChecker.CheckParent_new(CheckData_result.value)
            if (CheckParent_result.success = false) {
                EventManager.Publish("ProcessError", {processId: CheckParent_result.function, status: "Error", details: CheckParent_result.error . " - Line Number: " . A_LineNumber, result: {}})
                EventManager.Publish("PI_Stop", {inputValue: "Errore nella verifica dei parent!"}) ; Ferma l'indicatore di progresso
                throw Error("Errore nella verifica dei parent.")                
            }
            
            Check_2_Livello_result := FLChecker.Check_2_Livello(CheckData_result.value)
            if (Check_2_Livello_result.success = false) {
                EventManager.Publish("ProcessError", {processId: Check_2_Livello_result.function, status: "Error", details: Check_2_Livello_result.error . " - Line Number: " . A_LineNumber, result: {}})
                EventManager.Publish("PI_Stop", {inputValue: "Errore nella verifica del 2° livello!"}) ; Ferma l'indicatore di progresso
                throw Error("Errore nella verifica dei parent.")                
            }

            Check_Duplicati_result := FLChecker.Check_Duplicati(CheckData_result.value)
            if (Check_Duplicati_result.success = false) {
                EventManager.Publish("ProcessError", {processId: Check_Duplicati_result.function, status: "Error", details: Check_Duplicati_result.error . " - Line Number: " . A_LineNumber, result: {}})
                EventManager.Publish("PI_Stop", {inputValue: "Errore nella verifica dei duplicati!"}) ; Ferma l'indicatore di progresso
                throw Error("Errore nella verifica dei duplicati.")                
            }

            CheckGuideline_result := FLChecker.CheckGuideline(CheckData_result.value, CheckTechnology_result.value)
            if (CheckGuideline_result.success = false) {
                EventManager.Publish("ProcessError", {processId: CheckGuideline_result.function, status: "Error", details: CheckGuideline_result.error . " - Line Number: " . A_LineNumber, result: {}})
                EventManager.Publish("PI_Stop", {inputValue: "Errore nella verifica delle guideline!"}) ; Ferma l'indicatore di progresso
                throw Error("Errore nella verifica delle guideline.")                
            }

            MsgBoxResult := MsgBox("Check global table in SAP? (press Si or No)","Check SAP table", 4132)
            if (MsgBoxResult = "Yes") {
                ; se tutti i controlli precedenti vanno a buon fine allora procedo con
                ; il verificare se i diversi livelli delle FL sono già presenti in SAP
                EventManager.Publish("VerificaFL_SAP", {flArray: CheckData_result.value, flcountry: CheckCountry_result.value , fltechnology: CheckTechnology_result.value}) ; Invia una richiesta
            }
            else {
                FLChecker.VerificaFL_SAP := 1
/*                 EventManager.Publish("PI_Stop", {inputValue: "Verifica SAP FL interrotta"}) ; Ferma l'indicatore di progresso
                EventManager.Publish("ProcessCompleted", {processId: "CheckFL", status: "Completed", details: "Esecuzione interrotta dall'utente", result: {}})
                CheckFL_result.error := "Esecuzione interrotta dall'utente"
                CheckFL_result.success := true
                EventManager.Publish("AddLV", {icon: "icon2", element: "", text: CheckFL_result.error})
                return CheckFL_result */
            }
            ; attendo l'esito dell'evento VerificaFL_SAP
            ;if(FLChecker.WaitForConditionWithTimeout(FLChecker.VerificaFL_SAP, G_CONSTANTS.timeoutSeconds)) {
            while(!FLChecker.VerificaFL_SAP) {
                OutputDebug("FLChecker.VerificaFL_SAP " . FLChecker.VerificaFL_SAP . "`n")
                Sleep(100)
            }
            if (FLChecker.VerificaFL_SAP = -1) {
                throw Error("Errore nella verifica delle tabelle global in SAP")        
            }
            MsgBoxResult := MsgBox("Check global table asset in SAP? (press Si or No)","Check SAP asset table", 4132)
            if (MsgBoxResult = "Yes") {
                ; richiedo la verifica in SAP delle tabelle control asset
                EventManager.Publish("VerificaFL_SAP_ControlAsset", {flArray: CheckData_result.value, flcountry: CheckCountry_result.value , fltechnology: CheckTechnology_result.value, flInverterTechnology: FLChecker.inverterTechnology}) ; Invia una richiesta   
            }
            else {
                EventManager.Publish("PI_Stop", {inputValue: "Verifica SAP asset table interrotta"}) ; Ferma l'indicatore di progresso
                EventManager.Publish("ProcessCompleted", {processId: "CheckFL", status: "Completed", details: "Esecuzione interrotta dall'utente", result: {}})
                CheckFL_result.error := "Esecuzione interrotta dall'utente"
                CheckFL_result.success := true
                EventManager.Publish("AddLV", {icon: "icon2", element: "", text: CheckFL_result.error})
                return CheckFL_result
            }            


        }
        catch as err {
            CheckFL_result.error := "Errore: " . err.Message          
            EventManager.Publish("PI_Stop", {inputValue: CheckFL_result.error}) ; Ferma l'indicatore di progresso fornendo indicazioni sull'errore     
            EventManager.Publish("ProcessError", {processId: "CheckFL", status: "Error", details: CheckFL_result.error . " - Line Number: " . A_LineNumber, result: {}})
            MsgBox(CheckFL_result.error, "Error", 4144)
            return CheckFL_result      
        }
    } 

    static HasValue(arr, strValue) {
        for _, value in arr {
            if (value = strValue)
                return true
        }
        return false
    }

    static WaitForConditionWithTimeout(&condition, timeoutSeconds) {
        startTime := A_TickCount
        timeout := timeoutSeconds * 1000
    
        while (!condition) {  ; Qui, 'condition' riflette il valore della variabile originale
            if (A_TickCount - startTime > timeout) {
                return false
            }
            Sleep(100)
        }
    
        return true
    }



    ; Metodo per la pulizia dell'array.
    ; Elimina le righe vuote se presenti
    static FilterArray(arr) {
        filteredArr := []
        for element in arr {
            if (Trim(element) != "") {
                filteredArr.Push(Trim(element))
            }
        }
        return filteredArr
    }

    ; Funzione per verificare se una stringa è valida
    static IsValidString(str, pattern) {
        return RegExMatch(str, pattern) ? true : false
    }

    static CheckData(data) {
        result := { success: false, value: false, error: "", class: "FlChecker.ahk", function: "CheckFL" }
        EventManager.Publish("ProcessStarted", {processId: "CheckData", status: "Started", details: "Avvio funzione", result: {}})
        ; Dividiamo il contenuto in un array, considerando vari separatori
        tempArray := StrSplit(data, ["`r`n", "`n", "`r"])

        ; Filtriamo l'array per rimuovere elementi vuoti o solo spazi
        filteredArray := FLChecker.FilterArray(tempArray)

        ; Verifichiamo che ci siano elementi dopo il filtraggio
        if (filteredArray.Length = 0) {
            EventManager.Publish("ProcessError", {processId: "CheckData", status: "Error", details: "Contenuto della lista non valido. LN: " . A_LineNumber, result: {}})
            MsgBox("Errore: Nessuna FL valida trovata nell'input.", "Errore", 4112)
            return result
        }

        ; verifico che gli elementi presenti rispettino le maschere definite
        ; utilizzo una maschera generica, dato che ancora non ho rilevato la tecnologia
        maskPattern := "^(?:([A-Z0-9]{3})(?:-([A-Z0-9]{4})(?:-([A-Z0-9]{2})(?:-([A-Z0-9]{2,3})(?:-([A-Z0-9]{2,3})(?:-([A-Z0-9]{2}))?)?)?)?)?)?$"
        count := 0
        for element in filteredArray {
            if !(FLChecker.IsValidString(element, maskPattern)) {
                EventManager.Publish("AddLV", {icon: "icon2", element: element, text: "Elemento FL non valido"})
                count += 1
            }
        }
        if (count = 0) {
            result.success := true
            result.value := filteredArray
            EventManager.Publish("ProcessCompleted", {processId: "CheckData", details: "Funzione CheckData completata", result: result})
            return result
        }
        else {
            EventManager.Publish("ProcessError", {processId: "CheckData", details: "Contenuto della lista non valido. LN: " . A_LineNumber, result: {}})
            result.error := "Contenuto della lista non valido"
            return result
        }  
    }

    ; Funzione per leggere il file e estrarre i codici della FL
    ; Restituisce un array contenente i codici delle FL
    static GetFL_Codes(filename) {
        try {
            ; Legge il contenuto del file
            fileContent := FileRead(filename)

            ; Divide il contenuto in linee
            lines := StrSplit(fileContent, "`n", "`r")
            ; Rimuove l'intestazione
            lines.RemoveAt(1)
            ; Inizializza un array per i codici FL
            FL_Codes := []

            ; Estrae i codici paese
            for line in lines {
                if (line != "") {  ; Ignora le linee vuote
                    ; parts := StrSplit(line, "`t") ; utilizzato per i file di tipo .txt con separatore di elenco TAB
                    parts := StrSplit(line, ";") ; utilizzato con file .csv con separatore di elenco ;
                    if (parts.Length >= 2) { ; da verificare che tutte le FL siano associate ad una descrizione, ovvero composte da [0]FL [1]Descrizione
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

    ; Funzione per leggere il file e estrarre i codici country e tecnologia
    ; Restituisce un map con i codici come chiave e la descrizione come valore.
    static GetFileContent(fileName) {
        try {
            ; Legge il contenuto del file
            fileContent := FileRead(fileName)

            ; Divide il contenuto in linee
            lines := StrSplit(fileContent, "`n", "`r")

            ; Rimuove l'intestazione
            lines.RemoveAt(1)

            ; Inizializza un array per i codici
            mapCodes := map()

            ; Estrae i codici paese
            for line in lines {
                if (line != "") {  ; Ignora le linee vuote
                    ; parts := StrSplit(line, "`t") ; utilizzato per i file di tipo .txt con separatore di elenco TAB
                    parts := StrSplit(line, ";") ; utilizzato con file .csv con separatore di elenco ;
                    if (parts.Length = 2) { ; verifico che ogni elemento dell'array sia costituito da 2 elementi (codice -> descrizione)
                        code := parts[1]  ; Prende il primo elemento
                        if (StrLen(code) <= 2) { ; verifico che i codici siano composti al massimo di 2 caratteri
                            mapCodes[code] := parts[2]
                        }
                        else {
                            MsgBox("Errore lunghezza codici del file: " . FileName, "Errore", 4112)
                        }
                    }
                    else {
                        MsgBox("Errore nella struttura del file: " . FileName, "Errore", 4112)
                    }
                }
            }
            return mapCodes
        } catch Error as err {
            MsgBox("Errore nella lettura del file: " . FileName . " - " . err.Message, "Errore", 4112)
            return false ; in caso di errore restituisce un array vuoto
        }
    }

    ; Verifica il codice della country indicato nella FL con i codici contenuti nel file
    ; Restituisce:
    ; - la country indicata nella FL se valida
    ; - false altrimenti
    static CheckCountry(arr, FileName) {
        result := { success: false, value: false, error: "", class: "FlChecker.ahk", function: "CheckCountry" }
        EventManager.Publish("ProcessStarted", {processId: "CheckCountry", status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica country "}) ; Avvia l'indicatore di progresso
        mapCountry := Map()
        tempMap := Map()
        if (arr.Length = 0) {
            msgbox("Errore l'array indicato è vuoto!", "Errore", 4112)
            result.error := "Errore l'array indicato è vuoto!"
            return result  ; Se l'array è vuoto allora restituisco un errore
        }        
        ; verifico se la country è contenuta nel file
        mapCountry := FLChecker.GetFileContent(FileName)
        if MapCountry.count != 0 {
            for element in arr {
                FL_Levels := StrSplit(element, "-") ; considero il primo elemento dell'array e creo un nuovo array contenete i diversi livelli della FL
                country := SubStr(FL_Levels[1], 1, 2) ; considero i primi due caratteri
                ;~ Verifico che i codici country siano tutti uguali
                if tempMap.Has(country)
                    tempMap[country]++
                else
                    tempMap[country] := 1
            }
            ; verifico il numero di codici riscontrati
            CountryError := tempMap.Count
            if (CountryError = 1) { ; un solo codice presente
                ; verifico se è contenuto nel map <mapCountry>
                ; se viene generato un errore allora la chiave non esiste dunque il codice country non è contenuto nel map
                try {
                    value := mapCountry[country]
                }
                catch UnsetError as err {
                    result.error := "Errore codifica country"
                    EventManager.Publish("AddLV", {icon: "icon3", element: country, text: result.error})
                    EventManager.Publish("ProcessError", {processId: "CheckCountry", status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})
                    return result
                } 
                    result.success := true
                    result.value := country                    
                    EventManager.Publish("AddLV", {icon: "icon1", element: country, text: "Codifica country = " . value})
                    EventManager.Publish("ProcessCompleted", {processId: "CheckCountry", status: "Completed", details: "Esecuzione completata con successo", result: result})
                    ;EventManager.Publish("Debug",("Check country: OK"))
                    return result

            }
            else if (CountryError > 1) {
                country := ""
                for key, value in tempMap {
                    country .= key . "[" . value . "] "
                }
                result.error := "Codice country non univoco"
                EventManager.Publish("AddLV", {icon: "icon2", element: country, text: result.error})
                EventManager.Publish("ProcessError", {processId: "CheckCountry", status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})
                return result
            }
        }
        else {
            result.error := "Errore nella creazione del map mapCountry"
            msgbox("Errore nella creazione del map mapCountry", "Errore", 4112)
            EventManager.Publish("ProcessError", {processId: "CheckCountry", status: "Error", details: "Errore nella creazione del map mapCountry. Line Number: " . A_LineNumber, result: {}})
            return result  ; Se l'array è vuoto allora restituisco un errore
        }
    }

    ; Verifica il codice della tecnologia indicato nella FL
    ; Restituisce:
    ; - il tipo di tecnologia indicato nella FL se valido
    ; - false altrimenti
    static CheckTechnology(arr, FileName) {
        result := { success: false, value: false, error: "", class: "FlChecker.ahk", function: "CheckTechnology" }
        EventManager.Publish("ProcessStarted", {processId: "CheckTechnology", status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica tecnologia "}) ; Avvia l'indicatore di progresso        
        mapTech := map()
        tempMap := Map()
        if (arr.Length = 0) {
            msgbox("Errore l'array indicato è vuoto!", "Errore", 4112)
            result.error := "Errore l'array indicato è vuoto!"
            return result  ; Se l'array è vuoto allora restituisco un errore
        }
        else {
            ; verifico se la tecnologia è contenuta nel file
            mapTech := FLChecker.GetFileContent(FileName)
            if mapTech.count != 0 {
                for element in arr {
                    FL_Levels := StrSplit(element, "-") ; considero ogni elemento dell'array e creo un nuovo array contenete i diversi livelli della FL
                    Tech := SubStr(FL_Levels[1], -1) ; considero l'ultimo carattere del primo livello della FL
                    if tempMap.Has(Tech)
                        tempMap[Tech]++
                    else
                        tempMap[Tech] := 1
                }
                ; verifico il numero di codici riscontrati
                TechError := tempMap.Count
                if (TechError = 1) {
                    ; verifico se è contenuto nel map <mapCountry>
                    ; se viene generato un errore allora la chiave non esiste dunque il codice country non è contenuto nel map
                    try {
                        value := mapTech[Tech]
                        result.success := true
                        result.value := Tech                        
                        EventManager.Publish("AddLV", {icon: "icon1", element: Tech, text: "Codifica tecnologia = " . value})
                        EventManager.Publish("ProcessCompleted", {processId: "CheckTechnology", status: "Completed", details: "Esecuzione completata con successo", result: result})
                        ;EventManager.Publish("Debug",("Check country: OK"))
                        return result
                    } catch UnsetError as err {
                        result.error := "Codice tecnologia non valido"
                        EventManager.Publish("AddLV", {icon: "icon3", element: Tech, text: result.error})
                        EventManager.Publish("ProcessError", {processId: "CheckTechnology", status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})                   
                        return result
                    }
                }
                else if (TechError > 1) {
                    Tech := ""
                    for key, value in tempMap {
                        Tech .= key . "[" . value . "] "
                    }
                    result.error := "Codice tecnologia non univoco"
                    EventManager.Publish("AddLV", {icon: "icon3", element: Tech, text: result.error})
                    EventManager.Publish("ProcessError", {processId: "CheckTechnology", status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})                     
                    return result
                }
            }
        }
    }

    ;~ Funzione per verificare che ogni sede tecnica rispecchi la maschera definita per quella tecnologia
    ;~ Resitutisce:
    ;~ - True se non riscontra errori
    ;~ - False altrimenti
    ;~ - in caso di errori viene riportato nell LV la sede tecnica non conforme con il relativo errore
    static CheckMask(arr, tech) {
        result := { success: false, value: false, error: "", class: "FlChecker.ahk", function: "CheckMask" }
        EventManager.Publish("ProcessStarted", {processId: "CheckMask", status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica maschera "}) ; Avvia l'indicatore di progresso 
        count := 0
        if (arr.Length = 0) {
            msgbox("Errore l'array indicato è vuoto!", "Errore", 4112)
            result.error := "Errore l'array indicato è vuoto!"
            return result  ; Se l'array è vuoto allora restituisco un errore
        }
        if(tech != false) {
            map_Mask := Map()
            map_Mask := FLChecker.GetFL_Mask(G_CONSTANTS.file_Mask)
            ; Ricavo a partire dal file la maschera corrispondente alla tecnologia
            mask :=  map_Mask[tech]
            ;~ EventManager.Publish("Debug",(mask))
            maskPattern := FLChecker.MakeMaskPattern(mask)
            for element in arr {
                if !(FLChecker.IsValidString(element, maskPattern)) {
                    EventManager.Publish("AddLV", {icon: "icon3", element: element, text: "Non coerente con la maschera"})
                    count += 1
                }
            }
        }
        else {
            msgbox("Errore codifica country!", "Errore", 4112)
            result.error := "Errore codifica country!"
            return result  ; Se l'array è vuoto allora restituisco un errore
        }                    
        if (count = 0) {
            result.success := true
            EventManager.Publish("AddLV", {icon: "icon1", element: "", text: mask})
            EventManager.Publish("ProcessCompleted", {processId: "CheckMask", status: "Completed", details: "Esecuzione completata con successo", result: {}})
            return result
        }
        else {
            result.error := "Elementi non coerenti con la maschera."
            EventManager.Publish("ProcessError", {processId: "CheckMask", status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})                       
            return result
        }
    }

    static MakeMaskPattern(mask) {
            ;XXX-XXXX-XX-XX-XX-XX
            ;~ pattern := "^(?:([A-Z0-9]{index_1})(?:-([A-Z0-9]{index_2})(?:-([A-Z0-9]{index_3})(?:-([A-Z0-9]{index_4})(?:-([A-Z0-9]{index_5})(?:-([A-Z0-9]{index_6}))?)?)?)?)?)?$"
            pattern := "^([A-Z0-9]{index_1})$|^([A-Z0-9]{index_1}-[A-Z0-9]{index_2})$|^([A-Z0-9]{index_1}-[A-Z0-9]{index_2}-[A-Z0-9]{index_3})$|^([A-Z0-9]{index_1}-[A-Z0-9]{index_2}-[A-Z0-9]{index_3}-[A-Z0-9]{index_4})$|^([A-Z0-9]{index_1}-[A-Z0-9]{index_2}-[A-Z0-9]{index_3}-[A-Z0-9]{index_4}-[A-Z0-9]{index_5})$|^([A-Z0-9]{index_1}-[A-Z0-9]{index_2}-[A-Z0-9]{index_3}-[A-Z0-9]{index_4}-[A-Z0-9]{index_5}-[A-Z0-9]{index_6})$"
            ; verifico che la maschera contenga un valore coerente
            if(StrLen(mask) >= 20) and (StrLen(mask) <= 21) {
                ; conto le occorrenze delle X contenute tra i caratteri "-" nella maschera
                xCount := 0
                index_value := 1
                Loop parse mask {
                    char := A_LoopField
                    if (char = "X") {
                        xCount += 1
                    }
                    if (char = "-") or (StrLen(mask) = A_Index) {
                        index := "index_" . index_value
                        pattern := StrReplace(pattern, index , xCount, true)
                        index_value += 1
                        xCount := 0
                    }
                }
                EventManager.Publish("Debug",(pattern))
                return pattern
            }
            else {
                msgbox("Errore nella lunghezza della maschera FL", "Errore", 4112)
                return false
            }
    }

    ; Funzione per leggere il file e estrarre le maschere in base alla tecnologia
    ; Restituisce:
    ;   - un map contenente come chiave la tecnologia e come valore la maschera
    ;   - false in caso di errore
    static GetFL_Mask(filename) {
        try {
            ; Legge il contenuto del file
            fileContent := FileRead(filename)

            ; Divide il contenuto in linee
            lines := StrSplit(fileContent, "`n", "`r")
            ; Rimuove l'intestazione
            lines.RemoveAt(1)
            ; Inizializza un array per i codici FL
            Map_Mask := Map()

            ; Estrae i codici paese
            for line in lines {
                if (line != "") {  ; Ignora le linee vuote
                    ; parts := StrSplit(line, "`t") ; utilizzato per i file di tipo .txt con separatore di elenco TAB
                    parts := StrSplit(line, ";") ; utilizzato con file .csv con separatore di elenco ;
                    if (parts.Length = 3) { ; verifico che la struttura del file sia rispettata (3 colonne)
                        if (StrLen(parts[1]) = 5) { ; verifico che la stringa che utilizzerò come chiave sia lunga 5 caratteri
                            ; utilizzo il 4 carattere indicante la tecnologia come chiave.
                            Map_Mask[SubStr(parts[1], 4, 1)] := parts[2] ; Prende il primo elemento e lo uso come chiave, il secondo come valore
                        }
                        else {
                            msgbox("Errore nella lunghezza della chiave del file: " . filename, "Errore", 4112)
                            return false
                        }
                    }
                    else {
                        msgbox("Errore nella struttura del file: " . filename, "Errore", 4112)
                        return false
                    }
                }
            }
            return Map_Mask
        } catch Error as err {
            MsgBox("Errore nella lettura del file: " . filename . " - " . err.Message, "Errore", 4112)
            return false
        }
    }


    ;~ Funzione per verificare che ogni livello della FL possieda un livello genitore
    ;~ Restituisce:
    ;~ - True se entrambi i controlli vanno a buon fine
    ;~ - False altrimenti
    static CheckParent_new(arr) {
        result := { success: false, value: false, error: "", class: "FlChecker.ahk", function: "CheckParent_new" }
        EventManager.Publish("ProcessStarted", {processId: result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica dei parent "}) ; Avvia l'indicatore di progresso         
        Count := 0
        MapLenghtFL := Map()
        if (arr.Length = 0) {
            msgbox("Errore l'array indicato è vuoto!", "Errore", 4112)
            result.error := "Errore l'array indicato è vuoto!"
            return result  ; Se l'array è vuoto allora restituisco un errore
        }        
        ; creo una struttura map con chiave la lunghezza della FL e come valore un array contenente tutti gli elementi di quella lunghezza    
        for element in arr { ; considero ogni elemento della FL
            ; Conta il numero di occorrenze di "-"
            dashCount := StrSplit(element, "-").Length
            ;OutputDebug("Elemento: " . element . " - lunghezza: " . dashCount "`n")
            ; Se non esiste ancora un array per questo conteggio, crealo
            if (!MapLenghtFL.Has(dashCount))
                MapLenghtFL[dashCount] := []
            ; Aggiungi la linea all'array corrispondente
            MapLenghtFL[dashCount].Push(element)
        }    
        count := 0                    
        loop 7 { ; creo un loop di 6 iterazioni
            key := 7 - A_Index ; creo un indice da 6  a 0
            if (key > 2) { ; controllo fino al secondo livello - escludo il primo livello dal controllo
                if !(MapLenghtFL.Has(key))
                    continue                     
                arr :=  MapLenghtFL[key]
                for item in arr {
                    lastDashPos := InStr(item, "-", , -1)  ; Trova la posizione dell'ultimo trattino
                    if (lastDashPos > 1) { 
                        item_levelUp := SubStr(item, 1, lastDashPos - 1)  ; Estrae la sottostringa fino all'ultimo trattino escluso
                        ;OutputDebug("Elemento: " . item . " - levelUp: " . item_levelUp . "`n")
                        ; verifico se item_levelUp è contenuto nel map con key = key - 1
                        key_levelUp := key - 1
                        if !(MapLenghtFL.Has(key_levelUp)) { ; se non esiste il livello 
                            EventManager.Publish("AddLV", {icon: "icon3", element: item, text: "Parent " . item_levelUp . " mancante"})                              
                            count += 1 
                        }
                        else if (FlChecker.HasValue(MapLenghtFL[key_levelUp], item_levelUp)) {
                            ;OutputDebug("OK" . "`n")
                        }    
                        else {
                            EventManager.Publish("AddLV", {icon: "icon3", element: item, text: "Parent " . item_levelUp . " mancante"})
                            MapLenghtFL[key_levelUp].push(item_levelUp) ; inserisco il valore nel map superiore per essere controlato alla prossima iterazione                         
                            count += 1
                        }
                    }
                }
            }
        }
        if (count = 0) {
            EventManager.Publish("AddLV", {icon: "icon1", element: "", text: "Parent OK"})
            EventManager.Publish("ProcessCompleted", {processId: result.function, status: "Completed", details: "Esecuzione completata con successo", result: {}})               
            result.success := true
            return result 
        }
        else {
            result.error := "Elementi parent mancanti."
            EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})                  
            return result            
        }
    }




    ;~ Funzione per verificare che ogni livello della FL possieda un livello genitore
    ;~ Restituisce:
    ;~ - True se entrambi i controlli vanno a buon fine
    ;~ - False altrimenti
    static CheckParent(arr) {
        result := { success: false, value: false, error: "", class: "FlChecker.ahk", function: "CheckParent" }
        EventManager.Publish("ProcessStarted", {processId: result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica dei parent "}) ; Avvia l'indicatore di progresso         
        Count := 0
        MapParent := Map()
        if (arr.Length = 0) {
            msgbox("Errore l'array indicato è vuoto!", "Errore", 4112)
            result.error := "Errore l'array indicato è vuoto!"
            return result  ; Se l'array è vuoto allora restituisco un errore
        }
            for outerElement in arr { ; considero ogni elemento della FL
                ;~ EventManager.Publish("Debug",(element))
                Errore := ""
                ; conto quanti "-" sono contenuti nella FL
                RegExReplace(outerElement, "-", "", &ReplacementCount)
                DashCount := ReplacementCount
                ; se contiene almeno  il terzo livello procedo con il controllo altrimenti si tratta della radice della FL
                    ChildElement := outerElement
                    while (DashCount >= 2) {
                        ultimoDash := InStr(ChildElement, "-" , , -1, )
                        ParentElement := SubStr(ChildElement, 1 , ultimoDash - 1)
                        ;~ EventManager.Publish("Debug",(ParentElement))
                        for InnerElement in arr {
                            if (trim(ParentElement) = trim(InnerElement)) {
                                Count += 1 ; l'elemento parent esiste
                            }
                        }
                        if !(count) and !(MapParent.Has(ChildElement)) {
                            MapParent[ChildElement] := ParentElement
                        }
                        ChildElement := ParentElement
                        RegExReplace(ChildElement, "-", "", &ReplacementCount)
                        DashCount := ReplacementCount
                        Count := 0
                    }
            }
            ParentError := MapParent.Count
            if (ParentError) {
                for key, value in MapParent {
                    EventManager.Publish("AddLV", {icon: "icon3", element: key, text: "Parent " . value . " mancante"})
                }
                result.error := "Elementi parent mancanti."
                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})                        
                return result
            }
            else {
                EventManager.Publish("AddLV", {icon: "icon1", element: "", text: "Parent OK"})
                EventManager.Publish("ProcessCompleted", {processId: result.function, status: "Completed", details: "Esecuzione completata con successo", result: {}})
                result.success := true
                return result
            }
    }

    ;~ Funzione per verificare che tutti gli elementi del secondo livello siano uguali
    ;~ Restituisce:
    ;~ - True se tutti sono uguali
    ;~ - False altrimenti
    static Check_2_Livello(arr) {
        result := { success: false, value: false, error: "", class: "FlChecker.ahk", function: "Check_2_Livello" }
        EventManager.Publish("ProcessStarted", {processId: result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica 2° livello "}) ; Avvia l'indicatore di progresso 
        countMap := map()
            if (arr.Length = 0) {
                msgbox("Errore l'array indicato è vuoto!", "Errore", 4112)
                result.error := "Errore l'array indicato è vuoto!"
                return result  ; Se l'array è vuoto allora restituisco un errore
            }
            for elemento in arr {
                parts := StrSplit(elemento, "-")
                if (parts.Length >= 2) { ; verifico che la FL sia composta da almeno due elementi (es. USS-USS8)
                    ;~ EventManager.Publish("Debug",("parts[2] " . parts[2]))
                    if countMap.Has(parts[2])
                        countMap[parts[2]]++
                    else
                        countMap[parts[2]] := 1
                }
            }
            ; verifico il numero di codici riscontrati
            count := countMap.Count ; verifico il numero di elementi riscontrati
            if (count = 1) {
                result.success := true
                EventManager.Publish("AddLV", {icon: "icon1", element:  " Lev. 2", text: "Codice univoco"})
                EventManager.Publish("ProcessCompleted", {processId: result.function, status: "Completed", details: "Esecuzione completata con successo", result: {}})
                return result
            }
            else if (count > 1) {
                result.error := "Codice Lev. 2 non univoco"
                EventManager.Publish("AddLV", {icon: "icon3", element:  " Lev. 2", text: result.error})
                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})
                return result
            }
            else {
                result.error := "Errore verifica codice Lev. 2"
                EventManager.Publish("AddLV", {icon: "icon3", element:  " Lev. 2", text: result.error})
                return result
            }
    }

    ;~ Funzione per verificare che non esistano elementi duplicati
    ;~ Parametri:
    ;~ - un array contenente la lista delle FL
    ;~ Restituisce:
    ;~ - True se non esistono elementi duplicati
    ;~ - False altrimenti
    static Check_Duplicati(arr) {
        result := { success: false, value: false, error: "", class: "FlChecker.ahk", function: "Check_Duplicati" }
        EventManager.Publish("ProcessStarted", {processId: result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica duplicati "}) ; Avvia l'indicatore di progresso 
            countMap := map()
                if (arr.Length = 0) {
                    msgbox("Errore l'array indicato è vuoto!", "Errore", 4112)
                    result.error := "Errore l'array indicato è vuoto!"
                    return result  ; Se l'array è vuoto allora restituisco un errore
                }
                for elemento in arr {
                        if countMap.Has(elemento)
                            countMap[elemento]++
                        else
                            countMap[elemento] := 1
                    }
                ; verifico che ogni chiave abbia un solo numero di elementi
                for fl, count in countMap {
                    if (count > 1) {
                        result.error := "Codice duplicato."
                        EventManager.Publish("AddLV", {icon: "icon3", element: fl, text: result.error})
                    }                
                    else if (count != 1) {
                        result.error := "Anomalia conteggio duplicati."
                        EventManager.Publish("AddLV", {icon: "icon3", element: fl, text: result.error})
                    }
                }
                if (result.error = "") {
                    result.success := true
                    EventManager.Publish("AddLV", {icon: "icon1", element: "", text: "Controllo duplicati OK"})
                    EventManager.Publish("ProcessCompleted", {processId: result.function, status: "Completed", details: "Esecuzione completata con successo", result: {}})
                    return result
                }
                else {
                    result.error := "Errore controllo duplicati."
                    EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})                   
                    return result
                }
        }

    ; Funzione per verificare la coerenza con le Guideline
    ; Viene creato un unico aray con tutti i pattern e confrontato il singolo elemento della FL con tutte le occorrenze
    ;~ Parametri:
    ;~ - un array contenente la lista delle FL
    ;~ Restituisce:
    ;~ - True se tutti gli elementi della FL soddisfano i criteri della Guideline
    ;~ - False altrimenti, vengono riportati nella LV i valori che hanno generato errore.
    static CheckGuideline(arr, tech) {
        result := { success: false, value: false, error: "", class: "FlChecker.ahk", function: "CheckGuideline" }
        EventManager.Publish("ProcessStarted", {processId: result.function, status: "Started", details: "Avvio funzione", result: {}})
        EventManager.Publish("PI_Start", {inputValue: "Verifica guideline "}) ; Avvia l'indicatore di progresso 
            if (arr.Length = 0) {
                msgbox("Errore l'array indicato è vuoto!", "Errore", 4112)
                result.error := "Errore l'array indicato è vuoto!"
                return result  ; Se l'array è vuoto allora restituisco un errore
            }  
        
        result.success := true ; imposto a true e cambio in caso venga riscontrato almeno un errore

            if (tech = "S") { ; verifico gli impianti con tecnologia SOLAR
                ; Creo degli array a partire dai file
                pattern_array_FL_Solar_Common := FLChecker.MakePattern_array(FLChecker.GetFL_Codes(G_CONSTANTS.file_FL_Solar_Common))
                pattern_array_FL_S_SubStation := FLChecker.MakePattern_array(FLChecker.GetFL_Codes(G_CONSTANTS.file_FL_S_SubStation))

                EventManager.Publish("ShowInv", {}) ; mostro menu per selezione della tecnologia solar
                ; Loop di attesa per il risultato
                EventManager.Publish("PI_Start", {inputValue: "Attesa selezione tecnologia inverter "}) ; Avvia l'indicatore di progresso 
                while (FLChecker.inverterTechnology = "") {
                    Sleep(100)
                }
                ; in base alla tipologia di inverter creo il relativo array
                switch FLChecker.inverterTechnology
                {
                    case 1:
                        pattern_array_FL_Solar_Technology := FLChecker.MakePattern_array(FLChecker.GetFL_Codes(G_CONSTANTS.file_FL_Solar_CentralInv))
                        EventManager.Publish("DebugMsg",{msg: "Selezionato Central Inverter", linenumber: A_LineNumber})
                    case 2:
                        pattern_array_FL_Solar_Technology := FLChecker.MakePattern_array(FLChecker.GetFL_Codes(G_CONSTANTS.file_FL_Solar_StringInv))
                        EventManager.Publish("DebugMsg",{msg: "Selezionato String Inverter", linenumber: A_LineNumber})
                    case 3:
                        pattern_array_FL_Solar_Technology := FLChecker.MakePattern_array(FLChecker.GetFL_Codes(G_CONSTANTS.file_FL_Solar_InvModule))
                        EventManager.Publish("DebugMsg",{msg: "Selezionato Inverter Module", linenumber: A_LineNumber})
                    default:
                        result.success := false
                        result.error := "Errore nella selezione tecnologia inverter"
                        EventManager.Publish("ProcessError", {processId: "SelectInverter", details: "Nessuna tipologia di inverter selezionato.", result: result})
                        EventManager.Publish("AddLV", {icon: "icon2", element: "", text: "Programma interrotto dall'utente"})
                        EventManager.Publish("PI_Stop", {inputValue: "Errore nella selezione tecnologia inverter "}) ; Avvia l'indicatore di progresso 
                        return result
                }
                ; Controllo i valori al terzo livello
                EventManager.Publish("PI_Start", {inputValue: "Verifica guideline "}) ; Avvia l'indicatore di progresso
                for element in arr { ; per ogni riga
                    FL_Levels := StrSplit(trim(element), "-") ; scompongo la riga nei singoli livelli
                    if (FL_Levels.Length > 2) {
                        if ((FL_Levels[3] = "00") or (FL_Levels[3] = "9Z") or (FL_Levels[3] = "ZZ")) { ; Common
                            if !(FLChecker.TestGuideline_array(pattern_array_FL_Solar_Common, element)) { ; confronto l'elemento con le linee guida
                                result.error := "Errore guideline Common."
                                result.success := false
                                EventManager.Publish("AddLV", {icon: "icon3", element: " " . element, text: result.error})
                                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: result})                  
                            }
                        }
                        else if FL_Levels[3] = "0A" { ; Substation
                            if !(FLChecker.TestGuideline_array(pattern_array_FL_S_SubStation, element)) { ; confronto l'elemento con le linee guida
                                result.error := "Errore guideline Substation."        
                                result.success := false
                                EventManager.Publish("AddLV", {icon: "icon3", element: " " . element, text: result.error})
                                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: result})                                                 
                            }
                        }
                        else { ; in tutti gli altri casi devo considerare il tipo di tecnologia degli inverter
                            if !(FLChecker.TestGuideline_array(pattern_array_FL_Solar_Technology, element)) { ; confronto l'elemento con le linee guida
                                result.error := "Errore guideline Inverter."                
                                result.success := false
                                EventManager.Publish("AddLV", {icon: "icon3", element: " " . element, text: result.error})
                                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: result})                                       
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
            else if (tech = "E") { ; verifico gli impianti con tecnologia BESS
                ; Creo degli array a partire dai file
                pattern_array_FL_B_SubStation := FLChecker.MakePattern_array(FLChecker.GetFL_Codes(G_CONSTANTS.file_FL_B_SubStation))
                pattern_array_FL_Bess_Technology := FLChecker.MakePattern_array(FLChecker.GetFL_Codes(G_CONSTANTS.file_FL_Bess))
                ; Controllo i valori al terzo livello
                for element in Arr { ; per ogni riga
                    FL_Levels := StrSplit(trim(element), "-") ; scompongo la riga nei singoli livelli
                    if (FL_Levels.Length > 2) {
                        if FL_Levels[3] = "0A" { ; Substation
                            if !(FLChecker.TestGuideline_array(pattern_array_FL_B_SubStation, element)) { ; confronto l'elemento con le linee guida
                                result.error := "Errore guideline Substation."
                                result.success := false
                                EventManager.Publish("AddLV", {icon: "icon3", element: " " . element, text: result.error})
                                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})             
                            }
                        }
                        else { ; in tutti gli altri casi devo considerare il tipo di tecnologia degli inverter
                            if !(FLChecker.TestGuideline_array(pattern_array_FL_Bess_Technology, element)) { ; confronto l'elemento con le linee guida
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
                pattern_array_FL_W_SubStation := FLChecker.MakePattern_array(FLChecker.GetFL_Codes(G_CONSTANTS.file_FL_W_SubStation))
                pattern_array_FL_Wind_Technology := FLChecker.MakePattern_array(FLChecker.GetFL_Codes(G_CONSTANTS.file_FL_Wind))
                ; Controllo i valori al terzo livello
                for element in Arr { ; per ogni riga
                    FL_Levels := StrSplit(trim(element), "-") ; scompongo la riga nei singoli livelli
                    if (FL_Levels.Length > 2) {
                        if FL_Levels[3] = "0A" { ; Substation
                            if !(FLChecker.TestGuideline_array(pattern_array_FL_W_SubStation, element)) { ; confronto l'elemento con le linee guida
                                result.error := "Errore guideline Substation."
                                result.success := false
                                EventManager.Publish("AddLV", {icon: "icon3", element: " " . element, text: result.error})
                                EventManager.Publish("ProcessError", {processId: result.function, status: "Error", details: result.error . " - Line Number: " . A_LineNumber, result: {}})             
                            }
                        }
                        else { ; in tutti gli altri casi devo considerare il tipo di tecnologia degli inverter
                            if !(FLChecker.TestGuideline_array(pattern_array_FL_Wind_Technology, element)) { ; confronto l'elemento con le linee guida
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
    }

        ; Confronta un singolo elemento dell FL con l'array contenente tutti i pattern della guideline
        static TestGuideline_array(arr_Guideline, element) {
            count := 0
            for pattern in arr_Guideline { ; considero tutti gli elementi della Guideline
                if FLChecker.IsValidString(element, pattern) {
                    count += 1 ; per verificare che non ci sia più di una corrispondenza
                    }
            }
            Switch count
            {
            Case 0: ; nessuna corrispondenza trovata
                ;~ FLChecker.mainGui.gui.LV.Add("icon3", element, "Errore Guideline")
                ;~ FLChecker.mainGui.gui.LV.ModifyCol(1, "autoHdr")
                ;~ msgbox("Nessuna corrispondenza trovato nella guideline per: " . element)
                return false
            Case 1: ; una corrispondenza trovata
                ;EventManager.Publish("Debug",("TestGuideline: l'elemento " . element . " è corretto."))
                return true
            Default:
                EventManager.Publish("AddLV", {icon: "icon2", element: element, text: "Molteplici corrispondenze in Guideline"})
                ;~ msgbox("Molteplici corrispondenze trovato nella guideline per: " . element)
                return false
            }
        }


    ; Funzione: MakePattern_array
    ; Descrizione: Crea un array contenente tutti i pattern pressenti nei file delle linee guida da utilizzare nell'espressione regolare
    ; Parametri:
    ;   - param1: Descrizione del primo parametro
    ;   - param2: Descrizione del secondo parametro
    ; Restituisce: Descrizione di ciò che la funzione restituisce
    ; Esempio: NomeFunzione("esempio", 42)
    static MakePattern_array(arr) {
        new_arr := []
        MapRules := FLChecker.MakeMapRules(G_CONSTANTS.file_Rules)
        for pattern in arr {
            ; verifica la presenza dei codici presenti nel file rules e li sostituisce con le stringhe
            ; sostituisco inizialmente le occorrenze di "nn" per non avere errori nella sostituzione delle singole "n"
            pattern := StrReplace(pattern, "nn" , MapRules["nn"], true, , -1)
            pattern := StrReplace(pattern, "pp" , MapRules["pp"], true, , -1)
            for key, value in MapRules {
                pattern := StrReplace(pattern, key , value, true, , -1)
            }
            new_arr.push("^" . pattern . "$")
        }
        return new_arr
    }

    ; Funzione: MakePattern_array
    ; Descrizione: Legge il file e crea un map() con chiave i codici presenti nella guideline e come valori i codici da utilizzare nell'espressione regolare.
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
}