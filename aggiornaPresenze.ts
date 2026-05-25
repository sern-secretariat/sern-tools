function main(
  workbook: ExcelScript.Workbook,
  jsonTurni: string,
  jsonAssenze: string
) {
  let userIds: string[] = [
    "04fe1e3f-1f9e-41f1-90b3-a0f2d0270c58",
    "b1ab83a9-5f86-4729-9b45-2b24aed7f5cc",
    "d28776ec-51f1-4ca4-ae71-596b574beb71",
    "dfef8343-9d87-48ff-b801-8617b6d7651c",
    "4b01936b-9645-4a10-b475-0be99f872c0a",
    "89f1124d-4fb6-48b0-b620-54972c41069d"
  ];
  let userCognomi: string[] = [
    "Catellani",
    "Mereu",
    "Avanzi",
    "Roli",
    "Gravina",
    "Ricci"
  ];

  let torIds: string[] = [
    "TOR_ab185598-929c-4170-a55a-64bf632c4d8b",
    "TOR_85c5b979-4ac1-47e6-a069-b29b8132ec32",
    "TOR_318f7ab2-b9a0-4aba-ad16-9b2a4ff90a6d",
    "TOR_18ac3bae-1855-4ade-8ca6-35f4f2e8cef8",
    "TOR_326820e6-c4b1-4ef8-8707-cc9dbf874ab2",
    "TOR_0e4deda3-c76b-4f1a-bfac-08692b8e6ee1"
  ];
  let torColonne: string[] = [
    "Ferie",
    "Malattia",
    "Permessi",
    "Permessi",
    "CongParentale",
    "Festivita"
  ];

  let tabella = workbook.getTable("TabellaPresenze");
  if (!tabella) {
    console.log("TabellaPresenze non trovata");
    return;
  }

  let colonne = tabella.getHeaderRowRange().getValues()[0];
  let righe = tabella.getRangeBetweenHeaderAndTotal().getValues();

  let iChiave = -1; let iDip = -1; let iMese = -1; let iGiorno = -1;
  let iOre = -1; let iMal = -1; let iPerm = -1; let iFerie = -1;
  let iCong = -1; let iBuoni = -1;
  for (let c = 0; c < colonne.length; c++) {
    let nome = colonne[c].toString();
    if (nome === "Chiave") iChiave = c;
    if (nome === "Dipendente") iDip = c;
    if (nome === "Mese") iMese = c;
    if (nome === "NumGiorno") iGiorno = c;
    if (nome === "OreOrdinarie") iOre = c;
    if (nome === "Malattia") iMal = c;
    if (nome === "Permessi") iPerm = c;
    if (nome === "Ferie") iFerie = c;
    if (nome === "CongParentale") iCong = c;
    if (nome === "BuoniPasto") iBuoni = c;
  }

  let chiavi: string[] = [];
  for (let r = 0; r < righe.length; r++) {
    chiavi.push(righe[r][iChiave].toString());
  }

  console.log("Righe in tabella: " + righe.length.toString());

  function getCognome(userId: string): string {
    for (let i = 0; i < userIds.length; i++) {
      if (userIds[i] === userId) return userCognomi[i];
    }
    return "";
  }

  function trovaRiga(chiave: string): number {
    for (let i = 0; i < chiavi.length; i++) {
      if (chiavi[i] === chiave) return i;
    }
    return -1;
  }

  function calcolaOre(startISO: string, endISO: string): number {
    let start = new Date(startISO);
    let end = new Date(endISO);
    let ore = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
    return Math.round(ore * 2) / 2;
  }

  function oreGiornoIntero(cognome: string, giornoSett: number): number {
    if (cognome === "Avanzi" && giornoSett === 1) return 6;
    return 8;
  }

  function dataEffettiva(isoStr: string): string {
    let d = new Date(isoStr);
    if (d.getUTCHours() >= 20) d.setUTCDate(d.getUTCDate() + 1);
    let anno = d.getUTCFullYear();
    let mese = d.getUTCMonth() + 1;
    let giorno = d.getUTCDate();
    let meseStr = mese < 10 ? "0" + mese.toString() : mese.toString();
    let giornoStr = giorno < 10 ? "0" + giorno.toString() : giorno.toString();
    return anno.toString() + "-" + meseStr + "-" + giornoStr;
  }

  let oggi = new Date();
  let meseCorrente = oggi.getMonth() + 1;
  let annoCorrente = oggi.getFullYear();
  let mesePrecedente = meseCorrente - 1;
  let annoPrecedente = annoCorrente;
  if (mesePrecedente === 0) { mesePrecedente = 12; annoPrecedente = annoCorrente - 1; }

  let meseTarget1 = annoPrecedente.toString() + "-" + (mesePrecedente < 10 ? "0" + mesePrecedente.toString() : mesePrecedente.toString());
  let meseTarget2 = annoCorrente.toString() + "-" + (meseCorrente < 10 ? "0" + meseCorrente.toString() : meseCorrente.toString());

  console.log("Mesi target: " + meseTarget1 + " e " + meseTarget2);

  let bodyRange = tabella.getRangeBetweenHeaderAndTotal();

  let righeAzzerate = 0;
  for (let r = 0; r < righe.length; r++) {
    let meseRiga = righe[r][iMese].toString();
    if (meseRiga === meseTarget1 || meseRiga === meseTarget2) {
      bodyRange.getCell(r, iOre).setValue(0);
      bodyRange.getCell(r, iMal).setValue(0);
      bodyRange.getCell(r, iPerm).setValue(0);
      bodyRange.getCell(r, iFerie).setValue(0);
      bodyRange.getCell(r, iCong).setValue(0);
      if (iBuoni >= 0) bodyRange.getCell(r, iBuoni).setValue(0);
      righe[r][iOre] = 0; righe[r][iMal] = 0; righe[r][iPerm] = 0;
      righe[r][iFerie] = 0; righe[r][iCong] = 0;
      righeAzzerate++;
    }
  }
  console.log("Righe azzerate: " + righeAzzerate.toString());

  let turni: object[] = JSON.parse(jsonTurni);
  let turniScritti = 0;
  let turniIgnorati = 0;
  let turnoChiavi: string[] = [];
  let turnoOreRaw: number[] = [];
  let turnoCognomi: string[] = [];

  for (let t = 0; t < turni.length; t++) {
    let turno = turni[t] as Record<string, unknown>;
    let userId = turno["userId"] as string;
    if (!userId) { turniIgnorati++; continue; }
    let cognome = getCognome(userId);
    if (cognome === "") { turniIgnorati++; continue; }
    let shift = turno["sharedShift"] as Record<string, unknown>;
    if (!shift) { turniIgnorati++; continue; }
    let startDT = shift["startDateTime"] as string;
    let endDT = shift["endDateTime"] as string;
    if (!startDT || !endDT) { turniIgnorati++; continue; }
    let dataParte = startDT.substring(0, 10);
    let chiave = cognome + "-" + dataParte;
    let rigaIdx = trovaRiga(chiave);
    if (rigaIdx === -1) { turniIgnorati++; continue; }
    let ore = calcolaOre(startDT, endDT);
    if (ore <= 0) { turniIgnorati++; continue; }
    let trovato = false;
    for (let k = 0; k < turnoChiavi.length; k++) {
      if (turnoChiavi[k] === chiave) {
        turnoOreRaw[k] = turnoOreRaw[k] + ore;
        trovato = true;
        break;
      }
    }
    if (!trovato) {
      turnoChiavi.push(chiave);
      turnoOreRaw.push(ore);
      turnoCognomi.push(cognome);
    }
    turniScritti++;
  }

  for (let k = 0; k < turnoChiavi.length; k++) {
    let chiave = turnoChiavi[k];
    let oreRaw = turnoOreRaw[k];
    let cognome = turnoCognomi[k];
    let oreNette = oreRaw > 4 ? oreRaw - 0.5 : oreRaw;
    if (cognome === "Catellani" || cognome === "Roli") {
      console.log("TURNO " + cognome + " " + chiave.substring(chiave.length - 10) + " raw:" + oreRaw.toString() + " netto:" + oreNette.toString());
    }
    let rigaIdx = trovaRiga(chiave);
    if (rigaIdx >= 0) {
      bodyRange.getCell(rigaIdx, iOre).setValue(oreNette);
      righe[rigaIdx][iOre] = oreNette;
    }
  }

  console.log("Turni scritti: " + turniScritti.toString() + ", ignorati: " + turniIgnorati.toString());

  let assenze: object[] = JSON.parse(jsonAssenze);
  let assenzeScritte = 0;
  let assenzeIgnorate = 0;

  for (let a = 0; a < assenze.length; a++) {
    let assenza = assenze[a] as Record<string, unknown>;
    let userId = assenza["userId"] as string;
    if (!userId) { assenzeIgnorate++; continue; }
    let cognome = getCognome(userId);
    if (cognome === "") { assenzeIgnorate++; continue; }
    let timeOff = assenza["sharedTimeOff"] as Record<string, unknown>;
    if (!timeOff) { assenzeIgnorate++; continue; }
    let startDT = timeOff["startDateTime"] as string;
    let endDT = timeOff["endDateTime"] as string;
    let reasonId = timeOff["timeOffReasonId"] as string;
    if (!startDT || !endDT || !reasonId) { assenzeIgnorate++; continue; }
    let tipoAssenza = "";
    for (let i = 0; i < torIds.length; i++) {
      if (torIds[i] === reasonId) { tipoAssenza = torColonne[i]; break; }
    }
    if (tipoAssenza === "" || tipoAssenza === "Festivita") { assenzeIgnorate++; continue; }
    let colTarget = -1;
    if (tipoAssenza === "Ferie") colTarget = iFerie;
    if (tipoAssenza === "Malattia") colTarget = iMal;
    if (tipoAssenza === "Permessi") colTarget = iPerm;
    if (tipoAssenza === "CongParentale") colTarget = iCong;
    if (colTarget === -1) { assenzeIgnorate++; continue; }
    let startDate = new Date(startDT);
    let endDate = new Date(endDT);
    let diffOre = (endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60);
    if (diffOre < 24) {
      let dataParte3 = startDT.substring(0, 10);
      let chiave3 = cognome + "-" + dataParte3;
      let rigaIdx3 = trovaRiga(chiave3);
      if (rigaIdx3 === -1) { assenzeIgnorate++; continue; }
      let ore = calcolaOre(startDT, endDT);
      if (ore <= 0) { assenzeIgnorate++; continue; }
      if (ore > 8) ore = 8;
      let oreEsistenti = parseFloat(righe[rigaIdx3][colTarget].toString());
      if (isNaN(oreEsistenti)) oreEsistenti = 0;
      bodyRange.getCell(rigaIdx3, colTarget).setValue(oreEsistenti + ore);
      righe[rigaIdx3][colTarget] = oreEsistenti + ore;
      assenzeScritte++;
    } else if (diffOre >= 24 && diffOre <= 25) {
      let dataEff = dataEffettiva(startDT);
      let chiaveSingola = cognome + "-" + dataEff;
      let rigaIdxS = trovaRiga(chiaveSingola);
      if (rigaIdxS === -1) {
        chiaveSingola = cognome + "-" + startDT.substring(0, 10);
        rigaIdxS = trovaRiga(chiaveSingola);
      }
      if (rigaIdxS === -1) { assenzeIgnorate++; continue; }
      let dataCheck = new Date(dataEff + "T00:00:00Z");
      let giornoSett = dataCheck.getUTCDay();
      if (giornoSett === 0 || giornoSett === 6) { assenzeIgnorate++; continue; }
      let oreAss = oreGiornoIntero(cognome, giornoSett);
      let oreEsistenti = parseFloat(righe[rigaIdxS][colTarget].toString());
      if (isNaN(oreEsistenti)) oreEsistenti = 0;
      bodyRange.getCell(rigaIdxS, colTarget).setValue(oreEsistenti + oreAss);
      righe[rigaIdxS][colTarget] = oreEsistenti + oreAss;
      assenzeScritte++;
    } else {
      let currentDate = new Date(startDate.getTime());
      if (currentDate.getUTCHours() >= 20) currentDate.setUTCDate(currentDate.getUTCDate() + 1);
      currentDate.setUTCHours(0, 0, 0, 0);
      let endCheck = new Date(endDate.getTime());
      if (endCheck.getUTCHours() >= 20) endCheck.setUTCDate(endCheck.getUTCDate() + 1);
      endCheck.setUTCHours(0, 0, 0, 0);
      while (currentDate.getTime() < endCheck.getTime()) {
        let giornoSett = currentDate.getUTCDay();
        if (giornoSett === 0 || giornoSett === 6) { currentDate.setUTCDate(currentDate.getUTCDate() + 1); continue; }
        let anno = currentDate.getUTCFullYear();
        let mese = currentDate.getUTCMonth() + 1;
        let giorno = currentDate.getUTCDate();
        let meseStr = mese < 10 ? "0" + mese.toString() : mese.toString();
        let giornoStr = giorno < 10 ? "0" + giorno.toString() : giorno.toString();
        let chiave2 = cognome + "-" + anno.toString() + "-" + meseStr + "-" + giornoStr;
        let rigaIdx2 = trovaRiga(chiave2);
        if (rigaIdx2 >= 0) {
          let oreAssenza = oreGiornoIntero(cognome, giornoSett);
          let oreEsistenti = parseFloat(righe[rigaIdx2][colTarget].toString());
          if (isNaN(oreEsistenti)) oreEsistenti = 0;
          bodyRange.getCell(rigaIdx2, colTarget).setValue(oreEsistenti + oreAssenza);
          righe[rigaIdx2][colTarget] = oreEsistenti + oreAssenza;
          assenzeScritte++;
        }
        currentDate.setUTCDate(currentDate.getUTCDate() + 1);
      }
    }
  }

  console.log("Assenze scritte: " + assenzeScritte.toString() + ", ignorate: " + assenzeIgnorate.toString());

  if (iBuoni >= 0) {
    let buoniScritti = 0;
    for (let r = 0; r < righe.length; r++) {
      let meseRiga = righe[r][iMese].toString();
      if (meseRiga !== meseTarget1 && meseRiga !== meseTarget2) continue;
      let dipendente = righe[r][iDip].toString();
      if (dipendente === "Sofia Ricci") continue;
      let oreOrd = parseFloat(righe[r][iOre].toString());
      if (isNaN(oreOrd)) oreOrd = 0;
      if (oreOrd >= 4) { bodyRange.getCell(r, iBuoni).setValue(8); buoniScritti++; }
      else bodyRange.getCell(r, iBuoni).setValue(0);
    }
    console.log("Buoni pasto scritti: " + buoniScritti.toString());
  }

  console.log("Completato!");
}
