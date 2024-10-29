#Requires AutoHotkey v2.0

class ProcessManager {
    static processes := Map()

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
        if (!ProcessManager.processes.Has(processId)) {
            ProcessManager.processes[processId] := {status: "", details: "", history: []}
        }
        
        ProcessManager.processes[processId].status := status
        ProcessManager.processes[processId].details := details

        if (status == "Completed") {
            startTime := ""
            tempStartTime := 0
            for entry in ProcessManager.processes[processId].history {
                if (entry.status == "Started") { ; devo prendere la più recente
                    startTime := entry.timestamp
                    if (startTime > tempStartTime)
                        tempStartTime := startTime
                }
            }
            
            if (startTime != "") {
                executionTime := DateDiff(A_Now, startTime, "Seconds")
                details .= " (Tempo di esecuzione: " . executionTime . " secondi)"
            }
        }

        ProcessManager.processes[processId].history.Push({timestamp: A_Now, status: status, details: details})

        EventManager.Publish("ProcessStatusUpdated", {processId: processId, status: status, details: details, result: result})
    }

    static GetProcessStatus(processId) {
        return ProcessManager.processes.Has(processId) ? ProcessManager.processes[processId] : {processId: processId, status: "Unknown", details: "Process not found", history: []}
    }

    static GetAllProcesses() {
        return ProcessManager.processes
    }

    static ClearProcessHistory(processId) {
        if (ProcessManager.processes.Has(processId)) {
            ProcessManager.processes.Delete(processId)
            ;EventManager.Publish("ProcessCleared", {processId: processId})
        }
    }

    static ClearAllProcesses() {
        ProcessManager.processes := Map()
        ;EventManager.Publish("AllProcessesCleared")
    }
}

; Esempio di utilizzo
ExampleUsage() {
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
}