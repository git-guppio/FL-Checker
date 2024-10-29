; Classe: SAP_Connection
; Descrizione: Classe contenente metodi per gestire una connessione SAP
; Esempio: session := SAPConnection.GetSession()

class SAPConnection {
    static _oSAP := ""
    static connection := ""
    static application := ""
    static session := ""
    static timeout := 30  ; timeout configurabile


    static Connect() {
        if (this.IsConnected()) {
            return true
        }

        if (this.ConnectToExistingSession()) {
            return true
        }

        return this.OpenNewSession()
    }    

    static ConnectToExistingSession() {
        if (WinExist("ahk_class SAP_FRONTEND_SESSION")) {
            try {
                WinActivate
                this.application := ComObject("SAPGUI").GetScriptingEngine
                this.session := this.application.ActiveSession
                return true
            } catch as err {
                this.ShowError("Errore connessione a sessione SAP esistente: " . err.Message)
                return false
            }
        }
        return false
    }    

    static OpenNewSession() {
        try {
            this._oSAP := ComObject("SAPGUI")
        } catch {
            if (!this.LaunchSAP()) {
                return false
            }
        }

        try {
            this.application := this._oSAP.GetScriptingEngine
            this.connection := this.application.OpenConnection("0116 E4E - R4P - Power Generation", true)
            if (WinWaitActive("SAP Easy Access", , this.timeout)) {
                this.session := this.application.ActiveSession
                return true
            } else {
                this.ShowError("Timeout nell'apertura di SAP Easy Access")
                return false
            }
        } catch as err {
            this.ShowError("Errore nell'apertura di una nuova sessione SAP: " . err.Message)
            return false
        }
    }
    
    static LaunchSAP() {
        Run '"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"'
        if (WinWaitActive("SAP Logon", , this.timeout)) {
            try {
                this._oSAP := ComObject("SAPGUI")
                return true
            } catch as err {
                this.ShowError("Errore nell'inizializzazione di SAP dopo il lancio: " . err.Message)
                return false
            }
        } else {
            this.ShowError("Timeout nel lancio di SAP Logon")
            return false
        }
    }

    static IsConnectionFastEnough() {
        if (!this.IsConnected()) {
            this.ShowError("Non connesso a SAP")
            return false
        }

        try {
            isLowSpeed := this.session.Info.LowSpeed()
            if (isLowSpeed) {
                this.ShowError("La connessione SAP è troppo lenta per l'esecuzione ottimale degli script")
                return false
            }
            return true
        } catch as err {
            this.ShowError("Errore nel controllo della velocità di connessione SAP: " . err.Message)
            return false
        }
    }

    static ShowError(message) {
        MsgBox(message, "Errore SAP", 16)
    }

    static IsConnected() {
        return (this.session != "")
    }

    static Disconnect() {
        this._oSAP := ""
        this.connection := ""
        this.application := ""
        this.session := ""
    }

    static GetSession() {
        if (this.Connect()) {
            return this.session
        }
        return ""
    }
}