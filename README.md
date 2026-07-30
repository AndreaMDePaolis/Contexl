# Contexl addin per excel  
## aiutoFiscale

## 🚀 Installazione

1. Scarica dalla sezione **Releases** il file dell'add-in adatto alla tua versione di Excel:
   - `Contexl-Addin-x64.xll` → Excel **64 bit** (la più diffusa)
   - `Contexl-Addin-x86.xll` → Excel **32 bit**
2. Copia il file in una cartella stabile del tuo PC, ad esempio `C:\Add-in Excel\`.
3. Apri Excel e vai su **File → Opzioni → Componenti aggiuntivi**.
4. In fondo alla finestra, nel menu **Gestisci**, scegli **Componenti aggiuntivi di Excel** e premi **Vai…**.
5. Premi **Sfoglia…**, seleziona il file `.xll` e conferma con **OK**.    

💡 **Modo rapido:** puoi anche fare doppio clic sul file `.xll`. Le funzioni saranno disponibili solo per la sessione di Excel aperta in quel momento.   
> ⚠️ Se dopo aver seguito le indicazioni, Windows blocca il file e [appare una messaggio come questo](blocco.png), fai clic destro sul file scaricato → **Proprietà** → spunta **Annulla blocco** → **OK**.   

**Come verificare che funzioni:** in una cella qualsiasi scrivi `=FISCO.IVA(1000)` e premi Invio. Se compare `220`, l'installazione è riuscita 🎉   

**Per chi non ha installato .NET 10**, può scaricare il runtime (pochi MB) dal sito di Microsoft per la [versione a 64 bit](https://dotnet.microsoft.com/it-it/download/dotnet/thank-you/runtime-desktop-10.0.10-windows-x64-installer) o [qui a 32bit](https://dotnet.microsoft.com/it-it/download/dotnet/thank-you/runtime-desktop-10.0.10-windows-x86-installer).   

⚠️ **ATTENZIONE**   
Nel file Excel fornito come esempio, le formule sono volutamente non pre-calcolate: i valori delle funzioni personalizzate possono essere prodotti solo da Excel con l'add-in caricato. All'apertura vedrai #NOME? finché l'.xll non è attivo — poi basta Ctrl + Alt + F9 e tutto si popola.
