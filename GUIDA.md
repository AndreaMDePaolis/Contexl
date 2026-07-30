# 📊 AiutoFiscale Contexl Add-in per Excel

**Funzioni personalizzate per Excel** dedicate a chi lavora ogni giorno con contabilità, fatturazione e amministrazione: calcolo e scorporo dell'**IVA** (Italia e principali paesi UE), calcolo dell'**IRPEF**, generazione e controllo del **codice fiscale**, verifica di **partite IVA** e **IBAN**, gestione dei **calendari fiscali** e conversioni di testo in **Base64**.

Una volta installato l'add-in, le funzioni si usano **esattamente come le formule native di Excel**: basta digitare `=` e il nome della funzione.

---

## 📑 Indice

- [📊 AiutoFiscale Contexl Add-in per Excel](#-aiutofiscale-contexl-add-in-per-excel)
  - [📑 Indice](#-indice)
  - [🚀 Installazione](#-installazione)
  - [🔍 Come trovare le funzioni in Excel](#-come-trovare-le-funzioni-in-excel)
  - [⚡ Elenco rapido delle funzioni](#-elenco-rapido-delle-funzioni)
  - [💶 Funzioni fiscali](#-funzioni-fiscali)
    - [FISCO.IVA](#fiscoiva)
    - [FISCO.IRPEF](#fiscoirpef)
    - [PARTITA.IVA](#partitaiva)
    - [IVA.PROVINCIA.CODICE](#ivaprovinciacodice)
    - [IVA.PROVINCIA.NOME](#ivaprovincianome)
    - [IVA.PROVINCIA.MATRICE](#ivaprovinciamatrice)
  - [🆔 Funzioni per il codice fiscale](#-funzioni-per-il-codice-fiscale)
    - [CODICE.FISCALE](#codicefiscale)
    - [CF\_CHECK](#cf_check)
    - [CF\_COMUNE](#cf_comune)
    - [📥 Cartella di esempio pronta all'uso](#-cartella-di-esempio-pronta-alluso)
    - [📋 Esempio completo: foglio anagrafica](#-esempio-completo-foglio-anagrafica)
      - [Struttura del foglio](#struttura-del-foglio)
      - [Formule da inserire nella riga 2 (poi trascinare verso il basso)](#formule-da-inserire-nella-riga-2-poi-trascinare-verso-il-basso)
      - [Consigli per l'uso quotidiano](#consigli-per-luso-quotidiano)
    - [📋 Esempio: anagrafica clienti e fornitori](#-esempio-anagrafica-clienti-e-fornitori)
      - [Formule da inserire nella riga 2](#formule-da-inserire-nella-riga-2)
  - [📅 Funzioni per il calendario fiscale](#-funzioni-per-il-calendario-fiscale)
    - [FISCO.MESEFISCALE](#fiscomesefiscale)
    - [FISCO.QUADRIMESTRE](#fiscoquadrimestre)
    - [FISCO.ANNOFISCALE](#fiscoannofiscale)
    - [FISCO.SETTIMANAFISCALE](#fiscosettimanafiscale)
  - [🏦 Funzioni bancarie](#-funzioni-bancarie)
    - [BANCA.IBAN](#bancaiban)
  - [🔤 Funzioni di testo](#-funzioni-di-testo)
    - [STRINGA.BASE64](#stringabase64)
    - [BASE64.STRINGA](#base64stringa)
  - [🌍 Tabella delle aliquote IVA](#-tabella-delle-aliquote-iva)
  - [🗺 Elenco province e codici ufficio IVA](#-elenco-province-e-codici-ufficio-iva)
  - [❓ Domande frequenti](#-domande-frequenti)
  - [📜 Licenza](#-licenza)

---

## 🚀 Installazione

1. Scarica dalla sezione **Releases** il file dell'add-in adatto alla tua versione di Excel:
   - `Contexl-Addin-x64.xll` → Excel **64 bit** (la più diffusa)
   - `Contexl-Addin-x86.xll` → Excel **32 bit**
2. Copia il file in una cartella stabile del tuo PC, ad esempio `C:\Downloads\`.
3. Apri Excel e vai su **File → Opzioni → Componenti aggiuntivi**.
4. In fondo alla finestra, nel menu **Gestisci**, scegli **Componenti aggiuntivi di Excel** e premi **Vai…**.
5. Premi **Sfoglia…**, seleziona il file `.xll` e conferma con **OK**.

> 💡 **Modo rapido:** puoi anche fare doppio clic sul file `.xll`. Le funzioni saranno disponibili solo per la sessione di Excel aperta in quel momento.

> ⚠️ Se dopo il download Windows blocca il file, fai clic destro → **Proprietà** → spunta **Annulla blocco** → **OK**.

**Come verificare che funzioni:** in una cella qualsiasi scrivi `=FISCO.IVA(1000)` e premi Invio. Se compare `220`, l'installazione è riuscita 🎉

---

## 🔍 Come trovare le funzioni in Excel

Le funzioni sono raggruppate in tre categorie visibili nella finestra **Inserisci funzione** (`Maiusc + F3`):

| Categoria | Funzioni incluse |
| --------- | ---------------- |
| **Fisco** | IVA, IRPEF, partite IVA, codice fiscale, province, calendario fiscale |
| **Banca** | Verifica IBAN |
| **Testo** | Codifica e decodifica Base64 |

Selezionando una funzione dall'elenco, Excel mostra la **descrizione di ogni argomento**, così non serve ricordare a memoria i parametri.

Gli argomenti indicati **tra parentesi quadre** (es. `[operazione]`) sono **facoltativi**: se li ometti, l'add-in usa il valore predefinito.

---

## ⚡ Elenco rapido delle funzioni

| Funzione | Cosa fa | Esempio veloce |
| -------- | ------- | -------------- |
| `FISCO.IVA` | Calcola, scorpora o aggiunge l'IVA | `=FISCO.IVA(1000)` → `220` |
| `FISCO.IRPEF` | Calcola l'IRPEF lorda | `=FISCO.IRPEF(30000)` → `7140` |
| `PARTITA.IVA` | Verifica una partita IVA italiana | `=PARTITA.IVA(A1)` → `VERO` |
| `IVA.PROVINCIA.CODICE` | Provincia → codice ufficio IVA | `=IVA.PROVINCIA.CODICE("Trento")` → `022` |
| `IVA.PROVINCIA.NOME` | Codice ufficio IVA → provincia | `=IVA.PROVINCIA.NOME("022")` → `Trento` |
| `IVA.PROVINCIA.MATRICE` | Elenco completo province + codici | `=IVA.PROVINCIA.MATRICE()` |
| `CODICE.FISCALE` | Calcola il codice fiscale di una persona | `=CODICE.FISCALE("Mario";"Rossi";A1;"M";"H501")` |
| `CF_CHECK` | Riporta un codice fiscale omocodico alla forma base | `=CF_CHECK(A1)` |
| `CF_COMUNE` | Nome del comune → codice catastale | `=CF_COMUNE("Roma")` → `H501` |
| `FISCO.MESEFISCALE` | Numero del mese nell'anno fiscale | `=FISCO.MESEFISCALE(9;7)` → `3` |
| `FISCO.QUADRIMESTRE` | Periodo di 3 mesi di appartenenza | `=FISCO.QUADRIMESTRE(5)` → `2` |
| `FISCO.ANNOFISCALE` | Anno fiscale di una data | `=FISCO.ANNOFISCALE(A1;DATA(2026;7;1))` |
| `FISCO.SETTIMANAFISCALE` | Numero della settimana | `=FISCO.SETTIMANAFISCALE(OGGI())` |
| `BANCA.IBAN` | Verifica la correttezza di un IBAN | `=BANCA.IBAN(A1)` → `VERO` |
| `STRINGA.BASE64` | Testo → Base64 | `=STRINGA.BASE64("Ciao")` → `Q2lhbw==` |
| `BASE64.STRINGA` | Base64 → testo | `=BASE64.STRINGA("Q2lhbw==")` → `Ciao` |

---

## 💶 Funzioni fiscali

### FISCO.IVA

Calcola l'IVA, la **scorpora** da un importo lordo oppure restituisce il **totale con IVA inclusa**. Funziona sia su un singolo importo sia su un **intero intervallo di celle**.

```excel
=FISCO.IVA(importo; [operazione]; [stato]; [aliquota])
```

| Argomento | Obbligatorio | Valori ammessi | Predefinito |
| --------- | :----------: | -------------- | ----------- |
| **importo** | ✅ | Un numero oppure un intervallo di celle | – |
| **operazione** | ❌ | `0` = calcola l'IVA · `1` = scorpora l'IVA · `2` = totale con IVA inclusa | `0` |
| **stato** | ❌ | `0` = Italia · `1` = Francia · `2` = Germania · `3` = Spagna · `4` = Austria · `5` = Belgio · `6` = Paesi Bassi · `7` = Danimarca · `8` = Finlandia | `0` (Italia) |
| **aliquota** | ❌ | `0` = ordinaria · `1` = ridotta · `2` = seconda ridotta · `3` = minima | `0` |

**Esempi pratici**

```excel
=FISCO.IVA(1000)              →  220        ' IVA ordinaria italiana (22%)
=FISCO.IVA(1220; 1)           → 1000        ' Scorporo: dal lordo all'imponibile
=FISCO.IVA(1000; 2)           → 1220        ' Totale con IVA inclusa
=FISCO.IVA(1000; 0; 0; 1)     →  100        ' IVA italiana al 10%
=FISCO.IVA(1000; 0; 2)        →  190        ' IVA tedesca ordinaria (19%)
=FISCO.IVA(B2:B50; 2)         → matrice     ' Totali lordi di un'intera colonna
```

> 💡 **Su un intervallo:** scrivi la formula in una sola cella (es. `=FISCO.IVA(B2:B50)`) e nelle versioni recenti di Excel i risultati si espandono automaticamente su tutte le righe.

> ⚠️ Se scegli un'aliquota **non prevista** per quel paese (es. la "minima" in Germania) la funzione restituisce `#VALORE!`. Consulta la [tabella delle aliquote](#-tabella-delle-aliquote-iva) per sapere quali sono disponibili.

---

### FISCO.IRPEF

Calcola l'**IRPEF lorda** su un reddito imponibile, potendo scegliere tra il sistema di aliquote in vigore **dal 2026** e quello **precedente**.

```excel
=FISCO.IRPEF(reddito; [vecchiaIRPEF])
```

| Argomento | Obbligatorio | Valori ammessi | Predefinito |
| --------- | :----------: | -------------- | ----------- |
| **reddito** | ✅ | Reddito imponibile (numero) | – |
| **vecchiaIRPEF** | ❌ | `0` = sistema dal 2026 · `1` = vecchio sistema a 4 aliquote | `0` |

**Scaglioni applicati**

| Sistema dal 2026 | Aliquota | | Vecchio sistema | Aliquota |
| ---------------- | :------: |-| --------------- | :------: |
| fino a 28.000 €  | 23 % | | fino a 15.000 €   | 23 % |
| 28.001 – 50.000 €| 35 % | | 15.001 – 28.000 € | 25 % |
| oltre 50.000 €   | 43 % | | 28.001 – 50.000 € | 35 % |
|                  |      | | oltre 50.000 €    | 43 % |

**Esempi pratici**

```excel
=FISCO.IRPEF(30000)        →  7.140,00 €    ' Sistema 2026
=FISCO.IRPEF(30000; 1)     →  7.400,00 €    ' Vecchio sistema
=FISCO.IRPEF(60000)        → 18.440,00 €    ' Sistema 2026
```

Il risultato è già **arrotondato a 2 decimali**.

> ℹ️ La funzione calcola l'imposta **lorda**: non tiene conto di detrazioni, deduzioni o addizionali regionali e comunali.

---

### PARTITA.IVA

Verifica se una **partita IVA italiana** è formalmente corretta (11 cifre e cifra di controllo valida). Restituisce `VERO` o `FALSO`.

```excel
=PARTITA.IVA(partitaIVA)
```

**Esempi pratici**

```excel
=PARTITA.IVA("00743110157")               → VERO / FALSO
=PARTITA.IVA(A2)                          ' controlla la P.IVA presente in A2
=SE(PARTITA.IVA(A2); "✅ Valida"; "❌ Da verificare")
```

**Caso d'uso tipico:** controllo massivo dell'anagrafica clienti/fornitori. Copia la formula accanto alla colonna delle partite IVA e applica un filtro sui `FALSO` per individuare subito i dati da correggere.

> ⚠️ Se la partita IVA in Excel è memorizzata come **numero**, gli zeri iniziali vengono persi e il controllo fallisce. Formatta la colonna come **Testo** prima di inserire i dati.

> ℹ️ Il controllo è **formale**: conferma che il numero sia costruito correttamente, non che sia effettivamente attivo presso l'Agenzia delle Entrate.

---

### IVA.PROVINCIA.CODICE

Restituisce il **codice dell'ufficio IVA** corrispondente a una provincia italiana.

```excel
=IVA.PROVINCIA.CODICE(nome)
```

**Esempi pratici**

```excel
=IVA.PROVINCIA.CODICE("Trento")     → 022
=IVA.PROVINCIA.CODICE("Milano I")   → 015
=IVA.PROVINCIA.CODICE("Verona")     → 023
```

> ℹ️ Alcune grandi città hanno **più uffici**: vanno indicati come `"Milano I"`, `"Milano II"`, `"Roma I"`, `"Roma II"`, `"Firenze I"`, `"Genova II"`, ecc. Il nome non è sensibile a maiuscole/minuscole. Se la provincia non viene trovata, il risultato è `#VALORE!`.

---

### IVA.PROVINCIA.NOME

Operazione inversa della precedente: dal codice dell'ufficio IVA ottieni il **nome della provincia**.

```excel
=IVA.PROVINCIA.NOME(codice)
```

**Esempi pratici**

```excel
=IVA.PROVINCIA.NOME("022")   → Trento
=IVA.PROVINCIA.NOME("058")   → Roma I
=IVA.PROVINCIA.NOME("001")   → Torino
```

**Caso d'uso tipico:** risalire alla provincia di attribuzione leggendo le cifre dalla **8ª alla 10ª** di una partita IVA.

```excel
=IVA.PROVINCIA.NOME(STRINGA.ESTRAI(A2; 8; 3))
```

> ⚠️ Il codice va sempre scritto **tra virgolette** (`"022"`, non `22`), perché contiene zeri iniziali.

---

### IVA.PROVINCIA.MATRICE

Restituisce l'**elenco completo** delle province italiane con i relativi codici ufficio IVA: utile per creare tabelle di riferimento, elenchi a discesa o ricerche con `CERCA.VERT`.

```excel
=IVA.PROVINCIA.MATRICE()
```

Il risultato occupa **3 colonne**:

| Colonna | Contenuto |
| ------- | --------- |
| 1 | Nome della provincia |
| 2 | Codice dell'ufficio IVA |
| 3 | Riga descrittiva riassuntiva |

**Come usarla**

- **Excel Microsoft 365 / 2021 e successivi:** scrivi la formula in una cella; la tabella si espande da sola su tutte le righe.
- **Versioni precedenti:** seleziona un'area di **97 righe × 3 colonne**, scrivi la formula e conferma con `Ctrl + Maiusc + Invio`.

**Esempio di ricerca**

```excel
=CERCA.VERT("Bergamo"; IVA.PROVINCIA.MATRICE(); 2; FALSO)   → 016
```

---

## 🆔 Funzioni per il codice fiscale

Tre funzioni che lavorano insieme: `CF_COMUNE` trova il codice del comune, `CODICE.FISCALE` genera il codice fiscale e `CF_CHECK` normalizza i codici omocodici.

### CODICE.FISCALE

Calcola il **codice fiscale italiano** di una persona fisica a partire dai dati anagrafici.

```excel
=CODICE.FISCALE(nome; cognome; dataNascita; sesso; codiceLuogo; [indiceOmocodia])
```

| Argomento | Obbligatorio | Descrizione |
| --------- | :----------: | ----------- |
| **nome** | ✅ | Nome della persona |
| **cognome** | ✅ | Cognome della persona |
| **dataNascita** | ✅ | Data di nascita (cella con data Excel oppure data scritta come testo) |
| **sesso** | ✅ | `"M"` oppure `"F"` |
| **codiceLuogo** | ✅ | Codice catastale del comune di nascita, formato `A999` … `Z999` (es. `H501` per Roma) |
| **indiceOmocodia** | ❌ | `0` = codice base (predefinito) · da `1` a `127` = varianti omocodiche |

**Esempi pratici**

```excel
=CODICE.FISCALE("Mario"; "Rossi"; DATA(1980;5;12); "M"; "H501")
=CODICE.FISCALE(B2; A2; C2; D2; E2)                      ' dati presi da colonne separate
=CODICE.FISCALE(B2; A2; C2; D2; CF_COMUNE(F2))           ' comune ricavato dal nome
=CODICE.FISCALE("Mario"; "Rossi"; DATA(1980;5;12); "M"; "H501"; 1)   ' 1ª variante omocodica
```

**Cos'è l'omocodia:** quando due persone otterrebbero lo stesso codice fiscale, l'Agenzia delle Entrate assegna una **variante** sostituendo alcune cifre con lettere. L'argomento `indiceOmocodia` permette di generare queste varianti; nell'uso normale va lasciato a `0` o semplicemente omesso.

> ⚠️ La data di nascita deve essere una **data vera** (cella formattata come data o funzione `DATA()`), non un testo tipo `12-05-80` con formato ambiguo. In caso di dato non valido la funzione restituisce `#VALORE!`.

> ℹ️ Se il campo **sesso** non contiene `"F"`, viene considerato maschile.

**Caso d'uso tipico:** ricostruire i codici fiscali di un'anagrafica dipendenti o clienti importata da un gestionale, per poi confrontarli con quelli già presenti e individuare eventuali errori di digitazione.

---

### CF_CHECK

Ricalcola un codice fiscale in forma **normalizzata**, rimuovendo l'omocodia: le lettere inserite al posto delle cifre vengono riportate ai numeri originali.

```excel
=CF_CHECK(codiceFiscale)
```

| Argomento | Obbligatorio | Descrizione |
| --------- | :----------: | ----------- |
| **codiceFiscale** | ✅ | Codice fiscale di 16 caratteri |

**Esempi pratici**

```excel
=CF_CHECK(A2)                          ' restituisce il codice fiscale in forma base
=SE(CF_CHECK(A2)=CF_CHECK(B2); "Stessa persona"; "Persone diverse")
```

**Caso d'uso tipico:** confrontare due anagrafiche in cui la stessa persona può comparire una volta con il codice base e una volta con una variante omocodica. Normalizzando entrambi i codici con `CF_CHECK` i duplicati diventano immediatamente riconoscibili.

> ⚠️ Se il codice non ha 16 caratteri o contiene caratteri non ammessi, il risultato è `#VALORE!`. La funzione restituisce il codice normalizzato, non `VERO`/`FALSO`.

---

### CF_COMUNE

Restituisce il **codice catastale** (codice Belfiore) di un comune italiano partendo dal suo nome. È il valore da usare nell'argomento `codiceLuogo` di `CODICE.FISCALE`.

```excel
=CF_COMUNE(comune)
```

| Argomento | Obbligatorio | Descrizione |
| --------- | :----------: | ----------- |
| **comune** | ✅ | Nome del comune di nascita |

**Esempi pratici**

```excel
=CF_COMUNE("Roma")                     → H501
=CF_COMUNE("Milano")                   → F205
=CF_COMUNE(A2)                         ' comune indicato in A2
```

**Da usare in combinazione con `CODICE.FISCALE`:**

```excel
=CODICE.FISCALE(B2; A2; C2; D2; CF_COMUNE(E2))
```

> ⚠️ Il nome del comune deve corrispondere a quello ufficiale presente nell'archivio incorporato nell'add-in. Se non viene trovato, il risultato è `#VALORE!`: controlla eventuali abbreviazioni, apostrofi o denominazioni bilingue (es. i comuni dell'Alto Adige).

> 💡 L'archivio dei comuni è **incluso nell'add-in**: non serve alcuna connessione a Internet né un file esterno di appoggio.

---

### 📥 Cartella di esempio pronta all'uso

Nella sezione **Releases** trovi il file **`Fisco-AddIn-Esempi.xlsx`**: una cartella di lavoro con tutte le formule già impostate, pronta da aprire e modificare.

| Foglio | Contenuto |
| ------ | --------- |
| **Istruzioni** | Installazione, legenda dei colori e indice dei fogli |
| **Anagrafica Persone** | `CF_COMUNE`, `CODICE.FISCALE`, `CF_CHECK` |
| **Clienti Fornitori** | `PARTITA.IVA`, `IVA.PROVINCIA.NOME`, `BANCA.IBAN` |
| **Calcoli IVA** | `FISCO.IVA`: registro fatture, scorporo e confronto tra i 9 paesi |
| **IRPEF** | `FISCO.IRPEF`: confronto tra sistema 2026 e precedente |
| **Calendario Fiscale** | `FISCO.MESEFISCALE`, `FISCO.QUADRIMESTRE`, `FISCO.ANNOFISCALE`, `FISCO.SETTIMANAFISCALE` |
| **Province IVA** | `IVA.PROVINCIA.CODICE`, `IVA.PROVINCIA.NOME`, `IVA.PROVINCIA.MATRICE` |
| **Utilità Testo** | `STRINGA.BASE64`, `BASE64.STRINGA` |

**Legenda dei colori:** le celle **gialle con testo blu** sono dati che puoi modificare liberamente; le celle **grigie** contengono le formule dell'add-in.

> ⚠️ All'apertura potresti vedere `#NOME?` in tutte le celle con formula: significa semplicemente che **l'add-in non è ancora caricato**. Installalo come descritto sopra e premi `Ctrl + Alt + F9` per ricalcolare l'intera cartella.

---

### 📋 Esempio completo: foglio anagrafica

Uno schema pronto da replicare per **generare e controllare i codici fiscali** di un'intera anagrafica dipendenti o clienti.

#### Struttura del foglio

| | A | B | C | D | E | F | G | H | I |
|-|---|---|---|---|---|---|---|---|---|
| **1** | Cognome | Nome | Data nascita | Sesso | Comune nascita | Cod. catastale | CF calcolato | CF da gestionale | Esito |
| **2** | Rossi | Mario | 12/05/1980 | M | Roma | `H501` | `RSSMRA80E12H501U` | RSSMRA80E12H501U | ✅ Coincide |
| **3** | Bianchi | Anna | 03/09/1992 | F | Milano | `F205` | `BNCNNA92P43F205K` | BNCNNA92P43F205K | ✅ Coincide |
| **4** | Verdi | Luca | 27/11/1975 | M | Torino | `L219` | `VRDLCU75S27L219B` | VRDLCU75S27L219B | ✅ Coincide |
| **5** | Esposito | Giulia | 15/02/1988 | F | Napoli | `F839` | `SPSGLI88B55F839R` | SPSGL188B55F839R | ⚠️ Da verificare |

> Le colonne con sfondo grigio (**F**, **G**, **I**) contengono formule; le altre sono dati inseriti manualmente o importati dal gestionale.

#### Formule da inserire nella riga 2 (poi trascinare verso il basso)

```excel
F2  =SE.ERRORE(CF_COMUNE(E2); "❓ Comune non trovato")

G2  =SE.ERRORE(CODICE.FISCALE(B2; A2; C2; D2; F2); "❓ Dati incompleti")

I2  =SE(H2=""; "—";
     SE(SE.ERRORE(CF_CHECK(H2); "")=G2; "✅ Coincide"; "⚠️ Da verificare"))
```

**Cosa fa ogni formula**

| Cella | Funzione usata | Risultato |
| ----- | -------------- | --------- |
| **F** | `CF_COMUNE` | Traduce il nome del comune nel codice catastale richiesto dal calcolo |
| **G** | `CODICE.FISCALE` | Genera il codice fiscale usando i dati anagrafici e il codice della colonna F |
| **I** | `CF_CHECK` | Normalizza il codice presente nel gestionale (rimuovendo l'eventuale omocodia) e lo confronta con quello calcolato |

> 💡 Il passaggio da `CF_CHECK` nella colonna I è importante: senza normalizzazione, una persona con codice **omocodico** risulterebbe erroneamente "da verificare" pur essendo corretta.

#### Consigli per l'uso quotidiano

- **Formatta come tabella** (`Ctrl + T`): le formule si estendono in automatico alle nuove righe inserite.
- **Colonna C:** assicurati che le date siano riconosciute da Excel (allineate a destra). Se sono allineate a sinistra sono testo e vanno convertite con **Dati → Testo in colonne → Fine**.
- **Colonne G e H:** formattale come **Testo** e attiva l'allineamento a sinistra, così i codici restano leggibili e confrontabili.
- **Filtro rapido sugli errori:** applica un filtro alla colonna I e seleziona solo `⚠️ Da verificare` e `❓` per isolare in un attimo le righe da correggere.
- **Formattazione condizionale:** evidenzia in rosso le celle della colonna I diverse da `✅ Coincide` per un colpo d'occhio immediato.

---

### 📋 Esempio: anagrafica clienti e fornitori

Stessa logica applicata ai **dati fiscali e bancari** delle aziende, per validare partite IVA e IBAN prima di emettere fatture o disporre bonifici.

| | A | B | C | D | E | F |
|-|---|---|---|---|---|---|
| **1** | Ragione sociale | Partita IVA | P.IVA valida | Provincia ufficio | IBAN | IBAN valido |
| **2** | Alfa S.r.l. | `00743110157` | VERO | Milano I | IT60X054281110100000012345 | FALSO |
| **3** | Beta S.p.A. | `01234567890` | FALSO | — | IT60X0542811101000000123456 | VERO |

#### Formule da inserire nella riga 2

```excel
C2  =PARTITA.IVA(B2)

D2  =SE.ERRORE(IVA.PROVINCIA.NOME(STRINGA.ESTRAI(B2; 8; 3)); "—")

F2  =BANCA.IBAN(E2)
```

**Cosa fa ogni formula**

| Cella | Funzione usata | Risultato |
| ----- | -------------- | --------- |
| **C** | `PARTITA.IVA` | Verifica la correttezza formale della partita IVA (11 cifre + cifra di controllo) |
| **D** | `IVA.PROVINCIA.NOME` | Ricava l'ufficio IVA di attribuzione dalle cifre 8-10 della partita IVA |
| **F** | `BANCA.IBAN` | Controlla lunghezza, caratteri e cifra di controllo internazionale dell'IBAN |

> ⚠️ Formatta la colonna **B** come **Testo** prima di incollare le partite IVA: se Excel le interpreta come numeri, gli zeri iniziali vengono persi e il controllo restituisce `FALSO` anche su codici corretti.

---

## 📅 Funzioni per il calendario fiscale

Utilissime quando l'azienda ha un **esercizio non coincidente con l'anno solare** (es. inizio a luglio) o per raggruppare i dati nelle tabelle pivot.

### FISCO.MESEFISCALE

Restituisce il **numero progressivo del mese** all'interno dell'anno fiscale (da 1 a 12).

```excel
=FISCO.MESEFISCALE(mese; [mese_inizio])
```

| Argomento | Obbligatorio | Valori ammessi | Predefinito |
| --------- | :----------: | -------------- | ----------- |
| **mese** | ✅ | Numero del mese, da 1 a 12 | – |
| **mese_inizio** | ❌ | Mese di apertura dell'esercizio, da 1 a 12 | `1` (gennaio) |

**Esempi pratici**

```excel
=FISCO.MESEFISCALE(9; 7)        → 3     ' Settembre è il 3° mese se l'esercizio parte a luglio
=FISCO.MESEFISCALE(7; 7)        → 1     ' Luglio è il 1° mese
=FISCO.MESEFISCALE(6; 7)        → 12    ' Giugno è l'ultimo mese
=FISCO.MESEFISCALE(MESE(A2); 7) ' partendo da una data in A2
```

---

### FISCO.QUADRIMESTRE

Raggruppa i mesi in **blocchi da 3** e restituisce un numero da **1 a 4**, tenendo conto del mese di inizio dell'esercizio.

```excel
=FISCO.QUADRIMESTRE(mese; [inizio])
```

| Argomento | Obbligatorio | Valori ammessi | Predefinito |
| --------- | :----------: | -------------- | ----------- |
| **mese** | ✅ | Numero del mese, da 1 a 12 | – |
| **inizio** | ❌ | Mese di apertura dell'esercizio, da 1 a 12 | `1` (gennaio) |

**Esempi pratici**

```excel
=FISCO.QUADRIMESTRE(2)          → 1     ' Gen-Feb-Mar
=FISCO.QUADRIMESTRE(5)          → 2     ' Apr-Mag-Giu
=FISCO.QUADRIMESTRE(11)         → 4     ' Ott-Nov-Dic
=FISCO.QUADRIMESTRE(9; 7)       → 1     ' Con esercizio da luglio, settembre è nel 1° periodo
```

**Caso d'uso tipico:** creare una colonna di appoggio da usare come campo di raggruppamento nelle tabelle pivot per i report periodici.

---

### FISCO.ANNOFISCALE

Indica a quale **anno fiscale** appartiene una determinata data.

```excel
=FISCO.ANNOFISCALE(data; [data_inizio])
```

| Argomento | Obbligatorio | Descrizione |
| --------- | :----------: | ----------- |
| **data** | ✅ | La data da classificare |
| **data_inizio** | ❌ | La data di apertura dell'anno fiscale |

**Esempi pratici**

```excel
=FISCO.ANNOFISCALE(A2; DATA(2026;7;1))    ' Esercizio 1 luglio – 30 giugno
=FISCO.ANNOFISCALE(DATA(2026;9;15); DATA(2026;7;1))
```

> 💡 **Consiglio:** indica **sempre** il secondo argomento con la data di apertura del tuo esercizio. Ometterlo porta a risultati poco intuitivi. Se il tuo esercizio coincide con l'anno solare, usa semplicemente la funzione nativa `=ANNO(A2)`.

---

### FISCO.SETTIMANAFISCALE

Restituisce il **numero della settimana** a cui appartiene una data. La settimana inizia sempre di **lunedì** (convenzione europea).

```excel
=FISCO.SETTIMANAFISCALE(data; [data_inizio])
```

| Argomento | Obbligatorio | Descrizione |
| --------- | :----------: | ----------- |
| **data** | ✅ | La data di cui vuoi conoscere la settimana |
| **data_inizio** | ❌ | Se indicata, il conteggio riparte da quella data (giorno e mese di apertura dell'esercizio). Se omessa, si usa la numerazione standard ISO 8601 |

**Esempi pratici**

```excel
=FISCO.SETTIMANAFISCALE(OGGI())                         ' Settimana ISO corrente
=FISCO.SETTIMANAFISCALE(A2)                             ' Settimana ISO della data in A2
=FISCO.SETTIMANAFISCALE(A2; DATA(2026;7;1))             ' Settimana dell'esercizio che apre l'1 luglio
```

> ℹ️ Se la data di inizio non cade di lunedì, i giorni iniziali formano una **prima settimana più corta**, numerata comunque come settimana 1.

---

## 🏦 Funzioni bancarie

### BANCA.IBAN

Verifica se un **codice IBAN** è formalmente valido, controllando lunghezza, caratteri ammessi e la cifra di controllo internazionale (MOD 97). Restituisce `VERO` o `FALSO`.

```excel
=BANCA.IBAN(IBAN)
```

**Esempi pratici**

```excel
=BANCA.IBAN("IT60X0542811101000000123456")     → VERO / FALSO
=BANCA.IBAN("IT60 X054 2811 1010 0000 0123 456")   ' spazi e trattini ammessi
=SE(BANCA.IBAN(A2); "✅ OK"; "❌ IBAN errato")
```

**Caso d'uso tipico:** controllare i **codici IBAN dei fornitori o dei dipendenti** prima di generare un flusso di bonifici, evitando disposizioni scartate dalla banca.

> ℹ️ Puoi incollare l'IBAN con **spazi, trattini o lettere minuscole**: la funzione normalizza automaticamente il testo. Vengono accettati IBAN di qualunque paese (da 15 a 34 caratteri).

> ⚠️ Il controllo è **matematico e formale**: verifica che il codice sia scritto correttamente, non che il conto corrente esista davvero.

---

## 🔤 Funzioni di testo

### STRINGA.BASE64

Converte un testo nella sua **codifica Base64** (codifica UTF-8), formato spesso richiesto da portali web, tracciati telematici e integrazioni informatiche.

```excel
=STRINGA.BASE64(testo)
```

**Esempi pratici**

```excel
=STRINGA.BASE64("Ciao")                → Q2lhbw==
=STRINGA.BASE64(A2)                    ' codifica il contenuto di A2
```

---

### BASE64.STRINGA

Operazione inversa: riporta una stringa **Base64** al testo originale.

```excel
=BASE64.STRINGA(base64)
```

**Esempi pratici**

```excel
=BASE64.STRINGA("Q2lhbw==")            → Ciao
=BASE64.STRINGA(A2)                    ' decodifica il contenuto di A2
```

> ⚠️ Se il testo fornito non è una stringa Base64 valida, la funzione restituisce `#VALORE!`.

---

## 🌍 Tabella delle aliquote IVA

Valori utilizzabili nell'argomento **aliquota** di `FISCO.IVA`:

| Paese | Codice stato | `0` ordinaria | `1` ridotta | `2` seconda ridotta | `3` minima |
| ----- | :----------: | :-----------: | :---------: | :-----------------: | :--------: |
| 🇮🇹 Italia | `0` | 22 % | 10 % | 5 % | 4 % |
| 🇫🇷 Francia | `1` | 20 % | 10 % | 5,5 % | 2,1 % |
| 🇩🇪 Germania | `2` | 19 % | 7 % | — | — |
| 🇪🇸 Spagna | `3` | 21 % | 10 % | 4 % | — |
| 🇦🇹 Austria | `4` | 20 % | 13 % | 10 % | — |
| 🇧🇪 Belgio | `5` | 21 % | 12 % | 6 % | — |
| 🇳🇱 Paesi Bassi | `6` | 21 % | 9 % | — | — |
| 🇩🇰 Danimarca | `7` | 25 % | — | — | — |
| 🇫🇮 Finlandia | `8` | 25,5 % | 14 % | 10 % | — |

I trattini indicano aliquote **non previste** per quel paese: richiedendole, la funzione restituisce `#VALORE!`.

---

## 🗺 Elenco province e codici ufficio IVA

L'add-in include **97 voci** (tutte le province storiche italiane, con gli uffici multipli per Brescia, Firenze, Genova, Milano, Napoli e Roma). Puoi consultare l'elenco completo direttamente in Excel con:

```excel
=IVA.PROVINCIA.MATRICE()
```

Alcuni esempi: Torino `001` · Milano I `015` · Bergamo `016` · Trento `022` · Verona `023` · Venezia `027` · Bologna `037` · Firenze I `048` · Roma I `058` · Napoli I `063` · Bari `072` · Palermo `082` · Cagliari `092`

---

## ❓ Domande frequenti

**Le funzioni non compaiono quando digito `=FISCO...`**  
Verifica che l'add-in sia spuntato in **File → Opzioni → Componenti aggiuntivi → Componenti aggiuntivi di Excel → Vai…**. Se non compare nell'elenco, ripeti l'installazione con **Sfoglia…**.

**Ottengo `#NOME?`**  
Excel non riconosce la funzione: significa che l'add-in non è caricato in quella sessione. Riapri Excel o ricarica il file `.xll`.

**Ottengo `#VALORE!`**  
Uno degli argomenti non è valido: ad esempio un testo al posto di un numero, un mese fuori dall'intervallo 1-12, un'aliquota non prevista per quel paese, una provincia scritta in modo diverso da quello dell'elenco oppure un comune non riconosciuto da `CF_COMUNE`.

**Il codice fiscale calcolato non coincide con quello sulla tessera sanitaria**  
Nella maggior parte dei casi dipende da un **comune di nascita** o da una **data** non corretti, oppure dal fatto che la persona ha un codice **omocodico**: in questo caso usa `CF_CHECK` sul codice ufficiale per confrontarlo con quello base generato dall'add-in.

**Posso usare le funzioni in una cartella condivisa con colleghi?**  
Sì, ma **ogni utente deve avere l'add-in installato** sul proprio PC. Chi non lo ha vedrà `#NOME?` al posto dei risultati. Se devi inviare il file all'esterno, copia i risultati e incollali come **Valori** (`Ctrl + Alt + V` → Valori).

**Le funzioni rallentano il file?**  
No: sono calcolate in memoria e non richiedono connessione a Internet, quindi restano rapide anche su migliaia di righe.

**Le aliquote e gli scaglioni sono aggiornati?**  
Aliquote IVA e scaglioni IRPEF sono incorporati nell'add-in. In caso di modifiche normative verrà pubblicata una nuova versione nella sezione **Releases**: sostituisci il file `.xll` e riavvia Excel.

---

## 📜 Licenza

Distribuito con licenza **MIT**. Vedi il file `LICENSE` per i dettagli.

---

*Le funzioni forniscono un supporto al calcolo e ai controlli formali: non sostituiscono la verifica di un professionista abilitato.*
