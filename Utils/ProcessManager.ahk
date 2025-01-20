#Requires AutoHotkey v2.0

class ProcessManager {
    static processes := Map()

    ; Ottiene i tick correnti ad alta precisione
    static GetHighResolutionTime() {
        DllCall("QueryPerformanceCounter", "Int64*", &currentTime := 0)
        return currentTime
    }
    
    ; Converte i tick in millisecondi
    static ConvertTicksToMS(startTicks, endTicks) {
        ; Ottiene la frequenza del contatore ad alte prestazioni
        DllCall("QueryPerformanceFrequency", "Int64*", &frequency := 0)
        ; Calcola i millisecondi
        return (endTicks - startTicks) * 1000 / frequency
    }

    ; Formatta il tempo in minuti, secondi e millisecondi
    static FormatExecutionTime(milliseconds) {
        minutes := Floor(milliseconds / 60000)
        seconds := Floor((milliseconds - (minutes * 60000)) / 1000)
        ms := Round(Mod(milliseconds, 1000), 2)  ; Arrotonda a 2 decimali
        
        timeStr := ""
        if (minutes > 0)
            timeStr .= minutes . " min "
        if (seconds > 0 || minutes > 0)
            timeStr .= seconds . " sec "
        timeStr .= ms . " ms"
        
        return timeStr
    }

    static __New() {
        ProcessManager.SetupEventListeners()
    }

    static SetupEventListeners() {
        EventManager.Subscribe("ProcessStarted", (data) => ProcessManager.UpdateProcessStatus(data.processId, "Started", data.details, data.result))
        EventManager.Subscribe("ProcessProgress", (data) => ProcessManager.UpdateProcessStatus(data.processId, "In Progress", data.details, data.result))
        EventManager.Subscribe("ProcessCompleted", (data) => ProcessManager.UpdateProcessStatus(data.processId, "Completed", data.details, data.result))
        EventManager.Subscribe("ProcessError", (data) => ProcessManager.UpdateProcessStatus(data.processId, "Error", data.details, data.result))
    }

    static UpdateProcessStatus(processId, status, details := "", result := {}) {

        executionTimeMS := ""

        if (!ProcessManager.processes.Has(processId)) {
            ProcessManager.processes[processId] := {status: "", details: "", history: []}
        }
        
        ProcessManager.processes[processId].status := status
        ProcessManager.processes[processId].details := details

        ; Ottiene il timestamp corrente ad alta precisione
        currentTicks := this.GetHighResolutionTime()

        if (status == "Completed") {
            startTicks := 0
            latestStartTicks := 0
            
            ; Cerca il timestamp di Start più recente
            for entry in ProcessManager.processes[processId].history {
                if (entry.status == "Started") {
                    startTicks := entry.ticks
                    if (startTicks > latestStartTicks)
                        latestStartTicks := startTicks
                }
            }
            
            if (startTicks != 0) {
                ; Calcola il tempo di esecuzione in millisecondi
                executionTimeMS := this.ConvertTicksToMS(startTicks, currentTicks)
                details .= " (Tempo di esecuzione: " . this.FormatExecutionTime(executionTimeMS) . ")"
            }
        }

        ; Memorizza sia i tick che il timestamp leggibile
        ProcessManager.processes[processId].history.Push({
            timestamp: A_Now,
            ticks: currentTicks,
            status: status,
            details: details
        })

        EventManager.Publish("ProcessStatusUpdated", {
            processId: processId,
            status: status,
            details: details,
            result: result,
            executionTime: executionTimeMS
        })
    }
}
    
/*     ; Esempio di utilizzo
    ExampleUsage() {
        processId := "TestProcess"
        
        ; Avvia il processo
        EventManager.Publish("ProcessStarted", {
            processId: processId,
            details: "Inizio elaborazione"
        })
        
        ; Simula un'elaborazione
        Sleep(1234)  ; Attende 1.234 secondi
        
        ; Aggiorna il progresso
        EventManager.Publish("ProcessProgress", {
            processId: processId,
            details: "Elaborazione in corso"
        })
        
        Sleep(2345)  ; Attende altri 2.345 secondi
        
        ; Completa il processo
        EventManager.Publish("ProcessCompleted", {
            processId: processId,
            details: "Elaborazione completata"
        })
        
        ; Mostra la storia del processo
        process := ProcessManager.processes[processId]
        historyText := ""
        for entry in process.history {
            historyText .= "Status: " . entry.status . "`n"
            historyText .= "Details: " . entry.details . "`n`n"
        }
        MsgBox(historyText)

        ; Simulazione di processi
        EventManager.Publish("ProcessStarted", {processId: "Process1", details: "Inizializzazione"})
        Sleep(1000)
        EventManager.Publish("ProcessProgress", {processId: "Process1", details: "In corso"})
        Sleep(1000)
        EventManager.Publish("ProcessCompleted", {processId: "Process1", details: "Operazione completata con successo", result: {}})

        EventManager.Publish("ProcessStarted", {processId: "Process2", details: "Avvio download"})
        Sleep(500)
        EventManager.Publish("ProcessError", {processId: "Process2", details: "Errore di rete"})

        ; Recupero e visualizzazione dello stato dei processi
        MsgBox("Stato Process1: " . ProcessManager.GetProcessStatus("Process1").status)
        MsgBox("Stato Process2: " . ProcessManager.GetProcessStatus("Process2").status)

        ; Visualizzazione di tutti i processi
        allProcesses := ProcessManager.GetAllProcesses()
        for processId, processInfo in allProcesses {
            MsgBox("Processo: " . processId . "`nStato: " . processInfo.status . "`nDettagli: " . processInfo.details)
        }

        ; Pulizia
        ProcessManager.ClearAllProcesses()
} */