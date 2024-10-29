; Classe: EventManager
; Descrizione: Classe per la gestione degli eventi
; Esempio:

class EventManager {
    static events := Map()

    ; Metodo per sottoscrivere a un evento
    static Subscribe(eventName, func) {
        if (!this.events.Has(eventName)) {
            this.events[eventName] := []
        }
        this.events[eventName].Push(func)
    }

    ; Metodo per pubblicare un evento
    static Publish(eventName, data := "") {
        if (this.events.Has(eventName)) {
            for func in this.events[eventName] {
                func.Call(data)
            }
        }
    }

    ; Metodo per cancellare la sottoscrizione a un evento
    static Unsubscribe(eventName, callback) {
        if (this.events.Has(eventName)) {
            newCallbacks := []
            for cb in this.events[eventName] {
                if (cb != callback) {
                    newCallbacks.Push(cb)
                }
            }
            this.events[eventName] := newCallbacks
        }
    }   
}