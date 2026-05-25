function main(workbook: ExcelScript.Workbook) {
  let oggi = new Date();
  let meseCorrente = oggi.getMonth() + 1;
  let annoCorrente = oggi.getFullYear();
  let mesePrecedente = meseCorrente - 1;
  let annoPrecedente = annoCorrente;
  if (mesePrecedente === 0) { mesePrecedente = 12; annoPrecedente = annoCorrente - 1; }
  let mesiDaElaborare: number[] = [mesePrecedente, meseCorrente];
  let anniDaElaborare: number[] = [annoPrecedente, annoCorrente];
  let nomiMesi = ["", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];
  let tabella = workbook.getTable("TabellaPresenze");
  if (!tabella) { console.log("TabellaPresenze non trovata"); return; }
  let colonne = tabella.getHeaderRowRange().getValues()[0];
  let tutteLeRighe = tabella.getRangeBetweenHeaderAndTotal().getValues();
  let iMese = -1; let iDip = -1; let iGiorno = -1; let iOre = -1;
  let iMal = -1; let iPerm = -1; let iFerie = -1; let iCong = -1;
  let iBuoni = -1; let iIndTir = -1;
  for (let c = 0; c < colonne.length; c++) {
    let nome = colonne[c].toString();
    if (nome === "Mese") iMese = c;
    if (nome === "Dipendente") iDip = c;
    if (nome === "NumGiorno") iGiorno = c;
    if (nome === "OreOrdinarie") iOre = c;
    if (nome === "Malattia") iMal = c;
    if (nome === "Permessi") iPerm = c;
    if (nome === "Ferie") iFerie = c;
    if (nome === "CongParentale") iCong = c;
    if (nome === "BuoniPasto") iBuoni = c;
    if (nome === "IndTirocinio") iIndTir = c;
  }
  let tabellaRimb = workbook.getTable("TabellaRimborsi");
  let rimborsiRighe: (string | number | boolean)[][] = [];
  let iRDip = -1; let iRTipo = -1; let iRData = -1;
  let iRImporto = -1; let iRProgetto = -1;
  let iRProgDate = -1; let iRTravDate = -1; let iRViaggioLungo = -1;
  if (tabellaRimb) {
    let rimColonne = tabellaRimb.getHeaderRowRange().getValues()[0];
    rimborsiRighe = tabellaRimb.getRangeBetweenHeaderAndTotal().getValues();
    for (let c = 0; c < rimColonne.length; c++) {
      let nome = rimColonne[c].toString();
      if (nome === "Dipendente") iRDip = c;
      if (nome === "TipoRichiesta") iRTipo = c;
      if (nome === "Data") iRData = c;
      if (nome === "Importo") iRImporto = c;
      if (nome === "Progetto") iRProgetto = c;
      if (nome === "GiorniProgrammaDate") iRProgDate = c;
      if (nome === "GiorniViaggioDate") iRTravDate = c;
      if (nome === "Viaggio_lungo") iRViaggioLungo = c;
    }
    console.log("TabellaRimborsi: " + rimborsiRighe.length.toString() + " righe");
  }
  function parsaSingolaData(dataStr: string): { giorno: number; mese: number; anno: number } | null {
    dataStr = dataStr.trim();
    let spazio = dataStr.indexOf(" ");
    if (spazio > 0) dataStr = dataStr.substring(0, spazio);
    let parti = dataStr.split("/");
    if (parti.length !== 3) return null;
    let giorno = parseInt(parti[0]);
    let mese = parseInt(parti[1]);
    let anno = parseInt(parti[2]);
    if (anno < 100) anno = anno + 2000;
    if (isNaN(giorno) || isNaN(mese) || isNaN(anno)) return null;
    if (giorno < 1 || giorno > 31 || mese < 1 || mese > 12) return null;
    return { giorno: giorno, mese: mese, anno: anno };
  }
  function parsaListaDate(strDate: string): { giorno: number; mese: number; anno: number }[] {
    let risultati: { giorno: number; mese: number; anno: number }[] = [];
    if (!strDate || strDate.toString().trim() === "") return risultati;
    let parti = strDate.toString().split(",");
    for (let p = 0; p < parti.length; p++) {
      let parsed = parsaSingolaData(parti[p]);
      if (parsed !== null) risultati.push(parsed);
    }
    return risultati;
  }
  function meseStr(mese: number, anno: number): string {
    return anno.toString() + "-" + (mese < 10 ? "0" + mese.toString() : mese.toString());
  }
  for (let m = 0; m < mesiDaElaborare.length; m++) {
    let meseNum = mesiDaElaborare[m];
    let annoNum = anniDaElaborare[m];
    let meseTarget = meseStr(meseNum, annoNum);
    let nomeFoglio = nomiMesi[meseNum] + " " + annoNum.toString();
    let wsTarget = workbook.getWorksheet(nomeFoglio);
    if (!wsTarget) { console.log("Foglio non trovato: " + nomeFoglio); continue; }
    let datiFiltrati: (string | number | boolean)[][] = [];
    for (let r = 0; r < tutteLeRighe.length; r++) {
      if (tutteLeRighe[r][iMese].toString() === meseTarget) datiFiltrati.push(tutteLeRighe[r]);
    }
    let rimborsiMese: (string | number | boolean)[][] = [];
    for (let r = 0; r < rimborsiRighe.length; r++) {
      let riga = rimborsiRighe[r];
      let tipo = iRTipo >= 0 ? riga[iRTipo].toString() : "";
      let appartiene = false;
      if (tipo === "Mission") {
        let progDateStr = iRProgDate >= 0 ? riga[iRProgDate].toString() : "";
        let travDateStr = iRTravDate >= 0 ? riga[iRTravDate].toString() : "";
        let dateList = parsaListaDate(progDateStr + "," + travDateStr);
        for (let d = 0; d < dateList.length; d++) {
          if (meseStr(dateList[d].mese, dateList[d].anno) === meseTarget) { appartiene = true; break; }
        }
      } else {
        let dataStr = iRData >= 0 ? riga[iRData].toString() : "";
        if (dataStr !== "" && dataStr !== "0") {
          let parsed = parsaSingolaData(dataStr);
          if (parsed !== null) {
            if (meseStr(parsed.mese, parsed.anno) === meseTarget) appartiene = true;
          } else {
            let serial = parseFloat(dataStr);
            if (!isNaN(serial) && serial > 40000) {
              let excelDate = new Date((serial - 25569) * 86400 * 1000);
              if (meseStr(excelDate.getUTCMonth() + 1, excelDate.getUTCFullYear()) === meseTarget) appartiene = true;
            }
          }
        }
      }
      if (appartiene) rimborsiMese.push(riga);
    }
    console.log(nomeFoglio + ": " + datiFiltrati.length.toString() + " presenze, " + rimborsiMese.length.toString() + " rimborsi");
    if (datiFiltrati.length === 0 && rimborsiMese.length === 0) continue;
    let usedRange = wsTarget.getUsedRange();
    let values = usedRange.getValues();
    let numGiorni = new Date(annoNum, meseNum, 0).getDate();
    let colTotale = numGiorni + 1;
    for (let i = 0; i < values.length; i++) {
      let cellA = values[i][0] !== null && values[i][0] !== undefined ? values[i][0].toString().trim() : "";
      let cellAupper = cellA.toUpperCase();
      let cellAlower = cellA.toLowerCase();
      if (cellA === "ORE ORDINARIE" || cellA === "MALATTIA" || cellA === "PERMESSI" || cellA === "FERIE" ||
        cellAupper.indexOf("CONGEDO") >= 0 || cellAupper.indexOf("ALLOWANCE") >= 0) {
        for (let col = 1; col <= colTotale; col++) wsTarget.getCell(i, col).setValue("");
      }
      if (cellAlower.indexOf("rimborso") >= 0 || cellAlower.indexOf("buoni") >= 0 ||
        cellAlower.indexOf("indennit") >= 0 || cellAlower.indexOf("progetto") >= 0 ||
        cellAlower.indexOf("altri") >= 0) {
        for (let col = 1; col <= colTotale; col++) wsTarget.getCell(i, col).setValue("");
      }
    }
    let nomiTrovati: string[] = [];
    let righeOre: number[] = []; let righeMal: number[] = []; let righePerm: number[] = [];
    let righeFerie: number[] = []; let righeCong: (number | null)[] = [];
    let righeAllowance: (number | null)[] = []; let righeRimbKm: number[] = [];
    let righeAltriRimb: number[] = []; let righeProgetto: number[] = [];
    let righeBuoni: number[] = []; let righeIndTir: (number | null)[] = [];
    for (let i = 0; i < values.length; i++) {
      let cellA = values[i][0] !== null && values[i][0] !== undefined ? values[i][0].toString().trim() : "";
      if (cellA === "ORE ORDINARIE") {
        let nome = i > 0 && values[i - 1][0] !== null && values[i - 1][0] !== undefined ? values[i - 1][0].toString().trim() : "";
        if (nome === "" || nome === "TOTALE") continue;
        let rigaExcel = i + 1;
        let rigaCongedo: number | null = null;
        let offsetDopoCong = 4;
        if (i + 4 < values.length) {
          let check = values[i + 4][0] !== null ? values[i + 4][0].toString().trim().toUpperCase() : "";
          if (check.indexOf("CONGEDO") >= 0) { rigaCongedo = rigaExcel + 4; offsetDopoCong = 5; }
        }
        let rigaAllowance: number | null = null;
        if (i + offsetDopoCong < values.length) {
          let checkAll = values[i + offsetDopoCong][0] !== null ? values[i + offsetDopoCong][0].toString().trim().toUpperCase() : "";
          if (checkAll.indexOf("ALLOWANCE") >= 0) { rigaAllowance = rigaExcel + offsetDopoCong; offsetDopoCong++; }
        }
        let rigaIndTir: number | null = null;
        if (i + offsetDopoCong + 3 < values.length) {
          let checkTir = values[i + offsetDopoCong + 3][0] !== null ? values[i + offsetDopoCong + 3][0].toString().trim().toLowerCase() : "";
          if (checkTir.indexOf("tirocinio") >= 0) rigaIndTir = rigaExcel + offsetDopoCong + 3;
        }
        nomiTrovati.push(nome);
        righeOre.push(rigaExcel); righeMal.push(rigaExcel + 1); righePerm.push(rigaExcel + 2);
        righeFerie.push(rigaExcel + 3); righeCong.push(rigaCongedo); righeAllowance.push(rigaAllowance);
        righeRimbKm.push(rigaExcel + offsetDopoCong); righeAltriRimb.push(rigaExcel + offsetDopoCong + 1);
        righeProgetto.push(rigaExcel + offsetDopoCong + 2); righeBuoni.push(rigaExcel + offsetDopoCong + 3);
        righeIndTir.push(rigaIndTir);
      }
    }
    console.log("Dipendenti in griglia: " + nomiTrovati.length.toString());
    for (let d = 0; d < nomiTrovati.length; d++) {
      let nome = nomiTrovati[d];
      let totOre = 0; let totMal = 0; let totPerm = 0; let totFerie = 0; let totCong = 0;
      let totBuoni = 0; let totIndTir = 0;
      for (let r = 0; r < datiFiltrati.length; r++) {
        if (datiFiltrati[r][iDip].toString() !== nome) continue;
        let giorno = parseInt(datiFiltrati[r][iGiorno].toString());
        let col = giorno;
        let ore = parseFloat(datiFiltrati[r][iOre].toString());
        if (ore > 0) { wsTarget.getCell(righeOre[d] - 1, col).setValue(ore); totOre += ore; }
        let mal = parseFloat(datiFiltrati[r][iMal].toString());
        if (mal > 0) { wsTarget.getCell(righeMal[d] - 1, col).setValue(mal); totMal += mal; }
        let perm = parseFloat(datiFiltrati[r][iPerm].toString());
        if (perm > 0) { wsTarget.getCell(righePerm[d] - 1, col).setValue(perm); totPerm += perm; }
        let fer = parseFloat(datiFiltrati[r][iFerie].toString());
        if (fer > 0) { wsTarget.getCell(righeFerie[d] - 1, col).setValue(fer); totFerie += fer; }
        let cong = parseFloat(datiFiltrati[r][iCong].toString());
        if (cong > 0 && righeCong[d] !== null) { wsTarget.getCell(righeCong[d]! - 1, col).setValue(cong); totCong += cong; }
        let bp = datiFiltrati[r][iBuoni];
        if (bp !== "" && bp !== null && bp !== undefined) totBuoni += parseFloat(bp.toString());
        let it = datiFiltrati[r][iIndTir];
        if (it !== "" && it !== null && it !== undefined) totIndTir += parseFloat(it.toString());
      }
      if (totOre > 0) wsTarget.getCell(righeOre[d] - 1, colTotale).setValue(totOre);
      if (totMal > 0) wsTarget.getCell(righeMal[d] - 1, colTotale).setValue(totMal);
      if (totPerm > 0) wsTarget.getCell(righePerm[d] - 1, colTotale).setValue(totPerm);
      if (totFerie > 0) wsTarget.getCell(righeFerie[d] - 1, colTotale).setValue(totFerie);
      if (totCong > 0 && righeCong[d] !== null) wsTarget.getCell(righeCong[d]! - 1, colTotale).setValue(totCong);
      if (totBuoni > 0) wsTarget.getCell(righeBuoni[d] - 1, colTotale).setValue(totBuoni);
      if (totIndTir > 0 && righeIndTir[d] !== null) wsTarget.getCell(righeIndTir[d]! - 1, colTotale).setValue(totIndTir);
      let totRimbKm = 0; let totAltriRimb = 0; let totAllowance = 0; let progetti: string[] = [];
      for (let r = 0; r < rimborsiMese.length; r++) {
        let rigaRimb = rimborsiMese[r];
        if (iRDip < 0 || rigaRimb[iRDip].toString() !== nome) continue;
        let tipo = iRTipo >= 0 ? rigaRimb[iRTipo].toString() : "";
        let progetto = iRProgetto >= 0 ? rigaRimb[iRProgetto].toString() : "";
        if (progetto !== "" && progetti.indexOf(progetto) === -1) progetti.push(progetto);
        if (tipo === "Mileage Reimbursement") {
          let imp = iRImporto >= 0 ? rigaRimb[iRImporto] : "";
          if (imp !== "" && imp !== null && imp !== undefined) {
            let importo = parseFloat(imp.toString().replace(",", "."));
            if (!isNaN(importo) && importo > 0) totRimbKm += importo;
          }
        } else if (tipo === "Expense Reimbursement") {
          let imp = iRImporto >= 0 ? rigaRimb[iRImporto] : "";
          if (imp !== "" && imp !== null && imp !== undefined) {
            let importo = parseFloat(imp.toString().replace(",", "."));
            if (!isNaN(importo) && importo > 0) {
              let dataStr = iRData >= 0 ? rigaRimb[iRData].toString() : "";
              let giornoDato = 0;
              let parsed = parsaSingolaData(dataStr);
              if (parsed !== null) {
                giornoDato = parsed.giorno;
              } else {
                let serial = parseFloat(dataStr);
                if (!isNaN(serial) && serial > 40000) {
                  giornoDato = new Date((serial - 25569) * 86400 * 1000).getUTCDate();
                }
              }
              if (giornoDato > 0 && giornoDato <= numGiorni) {
                let esistente = wsTarget.getCell(righeAltriRimb[d] - 1, giornoDato).getValue();
                let esistenteNum = esistente !== "" && esistente !== null ? parseFloat(esistente.toString()) : 0;
                if (isNaN(esistenteNum)) esistenteNum = 0;
                wsTarget.getCell(righeAltriRimb[d] - 1, giornoDato).setValue(esistenteNum + importo);
              }
              totAltriRimb += importo;
            }
          }
        } else if (tipo === "Mission") {
          if (righeAllowance[d] !== null) {
            let vlRaw = iRViaggioLungo >= 0 ? rigaRimb[iRViaggioLungo].toString().toLowerCase().trim() : "";
            let valoreViaggio = (vlRaw === "yes") ? 1 : 0.5;
            let progDateStr = iRProgDate >= 0 ? rigaRimb[iRProgDate].toString() : "";
            let giorniProg = parsaListaDate(progDateStr);
            for (let gp = 0; gp < giorniProg.length; gp++) {
              let d2 = giorniProg[gp];
              if (meseStr(d2.mese, d2.anno) === meseTarget) {
                wsTarget.getCell(righeAllowance[d]! - 1, d2.giorno).setValue(1);
                totAllowance += 77.47;
              }
            }
            let travDateStr = iRTravDate >= 0 ? rigaRimb[iRTravDate].toString() : "";
            let giorniTrav = parsaListaDate(travDateStr);
            for (let gt = 0; gt < giorniTrav.length; gt++) {
              let dt = giorniTrav[gt];
              if (meseStr(dt.mese, dt.anno) !== meseTarget) continue;
              let giaProg = false;
              for (let gp = 0; gp < giorniProg.length; gp++) {
                if (giorniProg[gp].giorno === dt.giorno && meseStr(giorniProg[gp].mese, giorniProg[gp].anno) === meseTarget) { giaProg = true; break; }
              }
              if (!giaProg) {
                wsTarget.getCell(righeAllowance[d]! - 1, dt.giorno).setValue(valoreViaggio);
                totAllowance += valoreViaggio * 77.47;
              }
            }
          }
        }
      }
      if (totRimbKm > 0) wsTarget.getCell(righeRimbKm[d] - 1, colTotale).setValue(totRimbKm);
      if (totAltriRimb > 0) wsTarget.getCell(righeAltriRimb[d] - 1, colTotale).setValue(totAltriRimb);
      if (totAllowance > 0 && righeAllowance[d] !== null) wsTarget.getCell(righeAllowance[d]! - 1, colTotale).setValue(totAllowance);
      if (progetti.length > 0) wsTarget.getCell(righeProgetto[d] - 1, colTotale).setValue(progetti.join(", "));
      console.log(nome + ": ore=" + totOre.toString() + " allowance=" + totAllowance.toString() + " km=" + totRimbKm.toString() + " altri=" + totAltriRimb.toString());
    }
  }
  console.log("Completato!");
}
