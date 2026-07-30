# Contexl addin per excel  
## aiutoFiscale
Scegliere la versione a 32 bit o a 64 bit a seconda della versione di Excel installata nel sistema.   
Per installare l'add-in, una volta scaricata la versione adatta, apri **Excel** -> **File** -> **Opzioni** -> **Componenti aggiuntivi**.   
In basso a fianco a `Gestisci:` verificare che sia abilitata la voce `Componenti aggiuntivi di Excel` e cliccare sul pulsante `Vai...`   
Nella finestra che si apre, clicca sul pulsante `Sfoglia...` e cerca il file xll nella cartella dove si è scaricato e quindi selezionarlo.   
Chiudi la finestra cliccando sul pulsante `Ok`.   

💡 **Modo rapido:** puoi anche fare doppio clic sul file `.xll`. Le funzioni saranno disponibili solo per la sessione di Excel aperta in quel momento.   
> ⚠️ Se dopo il download Windows blocca il file, fai clic destro → **Proprietà** → spunta **Annulla blocco** → **OK**.

**Come verificare che funzioni:** in una cella qualsiasi scrivi `=FISCO.IVA(1000)` e premi Invio. Se compare `220`, l'installazione è riuscita 🎉   

**Per chi non ha installato .NET 10**, può scaricare il runtime (pochi MB) dal sito di Microsoft per la [versione a 64 bit](https://dotnet.microsoft.com/it-it/download/dotnet/thank-you/runtime-desktop-10.0.10-windows-x64-installer) o [qui a 32bit](https://dotnet.microsoft.com/it-it/download/dotnet/thank-you/runtime-desktop-10.0.10-windows-x86-installer).   

⚠️ **ATTENZIONE**   
Nel file Excel fornito come esempio, le formule sono volutamente non pre-calcolate: i valori delle funzioni personalizzate possono essere prodotti solo da Excel con l'add-in caricato. All'apertura vedrai #NOME? finché l'.xll non è attivo — poi basta Ctrl + Alt + F9 e tutto si popola.
