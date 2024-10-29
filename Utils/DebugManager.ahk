#Requires AutoHotkey v2.0

class DebugManager {
    static DEBUG_LOG_FILE := A_ScriptDir . "\debug_log.txt"
    static DEBUG_MODE := 2  ; 0 = disattivato, 1 = OutputDebug, 2 = File di log

    static __New() {
        ; Questo metodo viene chiamato automaticamente quando la classe viene caricata
        DebugManager.SetupEventListeners()
    }

    static SetupEventListeners() {
        EventManager.Subscribe("DebugMsg", (data) => DebugManager.DebugMsg(data.msg, data.linenumber))
        EventManager.Subscribe("SetDebug", (data) => DebugManager.SetDebugMode(data.mode))
        EventManager.Subscribe("ProcessMsg", (data) => DebugManager.ProcessMsg(data.processId, data.status, data.details))
        EventManager.Subscribe("ProcessStatusUpdated",(data) => DebugManager.ProcessStatusUpdated(data.processId, data.status, data.details, data.result))
    }


    static ProcessStatusUpdated(processId, status, details := "", result:={}) {
        ; Logica per aggiornare lo stato
        DebugManager.ProcessMsg(processId, status, details)
        return status
    }

    static IsCompiled() {
        return A_IsCompiled
    }

    static ProcessMsg(processId, status, details:="") {
        if (DebugManager.DEBUG_MODE = 0)
            return

            debugMsg := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") . " - " . processId . " - " . status . " - " .  details . "`n"

        if (DebugManager.DEBUG_MODE = 1) {
            OutputDebug(debugMsg)
        } else if (DebugManager.DEBUG_MODE = 2) {
            try {
                FileAppend(debugMsg, DebugManager.DEBUG_LOG_FILE)
            } catch as err {
                MsgBox("Errore nella scrittura del file di log: " . err.Message, "Errore Debug", 16)
            }
        }
    }

    static DebugMsg(msg, lineNumber := 0) {
        if (DebugManager.DEBUG_MODE = 0)
            return

        lineInfo := lineNumber ? " (Line " . lineNumber . ")" : ""
        debugMsg := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") . " - " . A_ScriptName . lineInfo . ": " . msg . "`n"

        if (DebugManager.DEBUG_MODE = 1) {
            OutputDebug(debugMsg)
        } else if (DebugManager.DEBUG_MODE = 2) {
            try {
                FileAppend(debugMsg, DebugManager.DEBUG_LOG_FILE)
            } catch as err {
                MsgBox("Errore nella scrittura del file di log: " . err.Message, "Errore Debug", 16)
            }
        }
    }

    static SetDebugMode(mode) { ; 0 = disattivato, 1 = OutputDebug, 2 = File di log
        DebugManager.DEBUG_MODE := mode
        modeStr := mode = 0 ? "Disattivato" : mode = 1 ? "OutputDebug/MsgBox" : "File di log"
        DebugManager.DebugMsg("Modalità di debug cambiata a: " . modeStr, A_LineNumber)
    }

    static ClearDebugLogFile() {
        if (FileExist(DebugManager.DEBUG_LOG_FILE)) {
            try {
                FileDelete(DebugManager.DEBUG_LOG_FILE)
                MsgBox("File di log cancellato")
            } catch as err {
                MsgBox("Errore nella cancellazione del file di log: " . err.Message, "Errore Debug", 16)
            }
        }
    }
}