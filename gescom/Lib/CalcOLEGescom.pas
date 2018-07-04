unit CalcOLEGescom;

interface

function GCCalcOLEEtat(sf, sp: string): variant;
function GCGetBasesEdt(sf, sp: string): Variant;
function GCGetEchesEdt(sf, sp: string): Variant;
function GCGetPiecePrecedenteEdt(sp: string): Variant;
function GCGetPortsEdt(sf, sp: string): Variant;
function GCGetPortsEdtTot(sf, sp: string): Variant;
procedure initvariable;
function GCDimToGen(sp: string): Variant;
function GCGetDimValR(sp: string): Variant;

const stPiecePrecedenteSauv: string = '';
  CodeArticle: string = ''; // permet de sauvegarder le code article pour les fonctions sur les tarifs
  CodeDevise: string = ''; // permet de sauvegarder le code devise pour les fonctions sur les tarifs
  CodeArtDispo: string = ''; // permet de sauvegarder le code article pour la fonction GCRechDispoArticle
  DateTarif: TDateTime = 0; // permet de sauvegarder la date pour les fonctions sur les tarifs
  DateDernierDoc: string = ''; // Fct GCGetDernierDoc : svgde date dernier document
implementation

uses SysUtils, Classes, HEnt1, UTOB, EntGC, UtilTarif, HDimension, TarifUtil,
  factutil, SaisUtil, StockUtil, UtilPGI,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  {$IFNDEF GCGC}
  CalcOle,
  {$ENDIF}
  {$IFDEF CCS3}
  {$IFDEF GRC}CalcOleGRC, {$ENDIF}
  {$IFDEF FOS5}CalcOLEFO, {$ENDIF}
  {$ELSE}
  {$IFDEF GRC}CalcOleGRC, {$ENDIF}
  {$IFDEF FOS5}CalcOLEFO, {$ENDIF}
  {$IFDEF MODE}UTofMBOArreteStock, {$ENDIF}
  {$IFDEF AFFAIRE}CalcOLEAffaire, {$ENDIF}
  {$ENDIF}
  HCtrls, ParamSoc, UtilArticle, Ent1;

procedure initvariable;
begin
  DateTarif := 0;
  CodeArticle := '';
  CodeArtDispo := '';
  CodeDevise := '';
end;

function GCReadParam(var St: string): string;
var i: Integer;
begin
  i := Pos(';', St);
  if i <= 0 then i := Length(St) + 1;
  result := Copy(St, 1, i - 1);
  Delete(St, 1, i);
end;

function GCCalculePrixArticleEdt(sp: string): Variant;
var st, stPrix: string;
  TobA: TOB;
begin
  Result := '';
  TobA := Tob.Create('ARTICLE', nil, -1);
  St := sp;

  try
    TobA.PutValue('GA_PAHT', StrToFloat(GCReadParam(st)));
  except
    TobA.PutValue('GA_PAHT', 0);
  end;
  try
    TobA.PutValue('GA_PRHT', StrToFloat(GCReadParam(st)));
  except
    TobA.PutValue('GA_PRHT', 0);
  end;
  try
    TobA.PutValue('GA_DPA', StrToFloat(GCReadParam(st)));
  except
    TobA.PutValue('GA_DPA', 0);
  end;
  try
    TobA.PutValue('GA_PMAP', StrToFloat(GCReadParam(st)));
  except
    TobA.PutValue('GA_PMAP', 0);
  end;
  try
    TobA.PutValue('GA_DPR', StrToFloat(GCReadParam(st)));
  except
    TobA.PutValue('GA_DPR', 0);
  end;
  try
    TobA.PutValue('GA_PMRP', StrToFloat(GCReadParam(st)));
  except
    TobA.PutValue('GA_PMRP', 0);
  end;
  try
    TobA.PutValue('GA_PVHT', StrToFloat(GCReadParam(st)));
  except
    TobA.PutValue('GA_PVHT', 0);
  end;
  try
    TobA.PutValue('GA_PVTTC', StrToFloat(GCReadParam(st)));
  except
    TobA.PutValue('GA_PVTTC', 0);
  end;
  TobA.PutValue('GA_CALCPRIXHT', GCReadParam(st));
  TobA.PutValue('GA_CALCPRIXTTC', GCReadParam(st));
  TobA.PutValue('GA_ARRONDIPRIX', GCReadParam(st));
  try
    TobA.PutValue('GA_COEFCALCHT', StrToFloat(GCReadParam(st)));
  except
    TobA.PutValue('GA_COEFCALCHT', 0);
  end;
  try
    TobA.PutValue('GA_COEFCALCTTC', StrToFloat(GCReadParam(st)));
  except
    TobA.PutValue('GA_COEFCALCTTC', 0);
  end;
  stPrix := GCReadParam(st);
  TobA.PutValue('GA_CALCAUTOHT', 'X');
  TobA.PutValue('GA_CALCAUTOTTC', 'X');

  CalculePrixArticle(TobA);
  if stPrix = 'TTC' then Result := TobA.GetValue('GA_PVTTC')
  else Result := TobA.GetValue('GA_PVHT');
  TobA.Free;
end;

function NumMois(Date: TDateTime): integer;
var Annee, Mois, Jour: Word;
begin
  DecodeDate(Date, Annee, Mois, Jour);
  Result := Mois;
end;

function FindTOBBase(Nature, Souche, CatTaxe: string; Numero, Indice: integer): TOB;
var TOBR: TOB;
  Q: TQuery;
begin
  TOBR := VH_GC.TOBGCB.FindFirst(['GPB_NATUREPIECEG', 'GPB_SOUCHE', 'GPB_NUMERO', 'GPB_INDICEG', 'GPB_CATEGORIETAXE'], [Nature, Souche, Numero, Indice,
    CatTaxe], True);
  if TOBR <> nil then TOBR := TOBR.Parent else
  begin
    TOBR := TOB.Create('', VH_GC.TOBGCB, -1);
    Q := OpenSQL('SELECT * FROM PIEDBASE WHERE GPB_NATUREPIECEG="' + Nature + '" AND GPB_SOUCHE="' + Souche + '" AND GPB_NUMERO=' + IntToStr(Numero) +
      ' AND GPB_INDICEG=' + IntToStr(Indice) + ' AND GPB_CATEGORIETAXE="' + CatTaxe + '"', True,-1,'',true);
    TOBR.LoadDetailDB('PIEDBASE', '', '', Q, False, True);
    Ferme(Q);
  end;
  Result := TOBR;
end;

function GCGetBasesEdt(sf, sp: string): Variant;
var Nature, Souche, St, Champ, CatTaxe: string;
  Numero, Indice, ii: integer;
  TOBB: TOB;
begin
  Result := '';
  St := sp;
  Champ := GCReadParam(st);
  ii := ValeurI(GCReadParam(st));
  Nature := GCReadParam(St);
  Souche := GCReadParam(St);
  Numero := ValeurI(GCReadParam(St));
  Indice := ValeurI(GCReadParam(St));
  CatTaxe := 'TX' + sf[Length(sf)];
  TOBB := FindTOBBase(Nature, Souche, CatTaxe, Numero, Indice);
  if TOBB = nil then Exit;
  if ii > TOBB.Detail.Count then Exit;
  if not TOBB.Detail[ii - 1].FieldExists('GPB_' + Champ) then Exit;
  Result := TOBB.Detail[ii - 1].GetValue('GPB_' + Champ);
end;

function FindTOBEche(Nature, Souche: string; Numero, Indice: integer): TOB;
var TOBR: TOB;
  Q: TQuery;
begin
  TOBR := VH_GC.TOBGCE.FindFirst(['GPE_NATUREPIECEG', 'GPE_SOUCHE', 'GPE_NUMERO', 'GPE_INDICEG'], [Nature, Souche, Numero, Indice], True);
  if TOBR <> nil then TOBR := TOBR.Parent else
  begin
    TOBR := TOB.Create('', VH_GC.TOBGCE, -1);
    Q := OpenSQL('SELECT * FROM PIEDECHE WHERE GPE_NATUREPIECEG="' + Nature + '" AND GPE_SOUCHE="' + Souche + '" AND GPE_NUMERO=' + IntToStr(Numero) +
      ' AND GPE_INDICEG=' + IntToStr(Indice), True,-1,'',true);
    TOBR.LoadDetailDB('PIEDECHE', '', '', Q, False, True);
    Ferme(Q);
  end;
  Result := TOBR;
end;

function GCGetEchesEdt(sf, sp: string): Variant;
var Nature, Souche, St, Champ, Dol: string;
  Numero, Indice, ii: integer;
  TOBE: TOB;
begin
  Result := '';
  St := sp;
  Champ := GCReadParam(st);
  ii := ValeurI(GCReadParam(st));
  Nature := GCReadParam(St);
  Souche := GCReadParam(St);
  Numero := ValeurI(GCReadParam(St));
  Indice := ValeurI(GCReadParam(St));
  Dol := GCReadParam(St);
  TOBE := FindTOBEche(Nature, Souche, Numero, Indice);
  if TOBE = nil then Exit;
  if ii > TOBE.Detail.Count then Exit;
  if not TOBE.Detail[ii - 1].FieldExists('GPE_' + Champ) then Exit;
  Result := TOBE.Detail[ii - 1].GetValue('GPE_' + Champ);
  if ((Champ = 'MODEPAIE') and (Dol = '$')) then Result := RechDom('TTMODEPAIE', Result, False);
end;

function FindTOBPort(Nature, Souche: string; Numero, Indice: integer): TOB;
var TOBR: TOB;
  Q: TQuery;
begin
  TOBR := TOB.Create('', nil, -1);
  Q := OpenSQL('SELECT * FROM PIEDPORT WHERE GPT_NATUREPIECEG="' + Nature + '" AND GPT_SOUCHE="' + Souche + '" AND GPT_NUMERO=' + IntToStr(Numero) +
    ' AND GPT_INDICEG=' + IntToStr(Indice), True,-1,'',true);
  TOBR.LoadDetailDB('PIEDPORT', '', '', Q, False, True);
  Ferme(Q);
  Result := TOBR;
end;

function GCGetPiecePrecedenteEdt(sp: string): Variant;
var st, stPiecePrecedente, stLigne, stClePiecePrecedente: string;
  stClePiecePrecedenteSauv, stLigneSauv: string;
  iRead, iInd: integer;
  dd, mm, yy: Word;
begin
  Result := '';
  St := sp;
  iInd := ValeurI(GCReadParam(st));
  stPiecePrecedente := St;
  if iInd = 9 then
  begin
    if stPiecePrecedente <> '' then
    begin
      stLigne := stPiecePrecedente;
      for iRead := 1 to 5 do // recuperation n° ligne courante
      begin
        ReadTokenSt(stLigne);
      end;
      stClePiecePrecedente := stPiecePrecedente;
      for iRead := 1 to 5 do // récupération cle piece precedente courante
      begin
        if iRead > 1 then Result := Result + ';';
        Result := Result + ReadTokenSt(stClePiecePrecedente);
      end;
      stClePiecePrecedente := Result;
      stLigneSauv := stPiecePrecedenteSauv;
      for iRead := 1 to 5 do // récupération n° ligne sauvée
      begin
        ReadTokenSt(stLigneSauv);
      end;
      Result := '';
      stClePiecePrecedenteSauv := StPiecePrecedenteSauv;
      for iRead := 1 to 5 do // récupération cle piece precedente sauv
      begin
        if iRead > 1 then Result := Result + ';';
        Result := Result + ReadTokenSt(stClePiecePrecedenteSauv);
      end;
      stClePiecePrecedenteSauv := Result;
      if (stClePiecePrecedenteSauv <> stClePiecePrecedente) or
        (stLigneSauv >= stLigne) then
      begin
        stPiecePrecedenteSauv := stPiecePrecedente;
        Result := 'X';
      end else Result := '-';
    end else Result := '-';
  end else
    if iInd = 0 then
  begin
    if stPiecePrecedente <> '' then
    begin
      for iRead := 1 to 5 do
      begin
        if iRead > 1 then Result := Result + ';';
        Result := Result + ReadTokenSt(stPiecePrecedente);
      end;
    end else Result := '';
  end else
  begin
    for iRead := 1 to iInd do
    begin
      st := ReadTokenst(stPiecePrecedente);
    end;
    if iInd = 1 then
    begin
      Result := '';
      if St <> '' then
      begin
        dd := StrToInt(Copy(st, 1, 2));
        mm := StrToInt(Copy(st, 3, 2));
        yy := StrToInt(Copy(st, 5, 4));
        Result := Encodedate(yy, mm, dd);
      end;
    end else Result := st;
  end;
end;

function GCGetPortsEdt(sf, sp: string): Variant;
var Nature, Souche, St, Champ: string;
  Numero, Indice, ii: integer;
  TOBP: TOB;
  QQ: TQuery;
  EnHT: Boolean;
begin
  Result := '';
  St := sp;
  Champ := GCReadParam(st);
  ii := ValeurI(GCReadParam(st));
  Nature := GCReadParam(St);
  Souche := GCReadParam(St);
  Numero := ValeurI(GCReadParam(St));
  Indice := ValeurI(GCReadParam(St));
  //Dol:=GCReadParam(St) ;
  TOBP := FindTOBPort(Nature, Souche, Numero, Indice);
  if TOBP = nil then Exit;
  if ii > TOBP.Detail.Count then
  begin
    TOBP.Free;
    Exit
  end;
  QQ := OpenSQL('SELECT GP_FACTUREHT FROM PIECE WHERE GP_NATUREPIECEG="' + Nature + '" AND GP_SOUCHE="' + Souche + '" AND GP_NUMERO=' + IntToStr(Numero) +
    ' AND GP_INDICEG=' + IntToStr(Indice), True,-1,'',true);
  EnHT := ((QQ.Fields[0].AsString) = 'X');
  Ferme(QQ);
  if not TOBP.Detail[ii - 1].FieldExists('GPT_' + Champ) then
  begin
    if Champ = 'CODEPDF' then Champ := 'CODEPORT';
    if Champ = 'LIBELLEPDF' then Champ := 'LIBELLE';
    if Champ = 'TAUXPDF' then Champ := 'POURCENT';
    if Champ = 'BASEPDF' then
      if EnHT then Champ := 'BASEHTDEV' else Champ := 'BASETTCDEV';
    if Champ = 'MONTANTPDF' then
      if EnHT then Champ := 'TOTALHTDEV' else Champ := 'TOTALTTCDEV';
    if (TOBP.Detail[ii - 1].Getvalue('GPT_FRAISREPARTIS') = 'X') or (TOBP.Detail[ii - 1].Getboolean('GPT_RETENUEDIVERSE')) then
    begin
      TOBP.Free;
      Exit
    end;
    if not TOBP.Detail[ii - 1].FieldExists('GPT_' + Champ) then
    begin
      TOBP.Free;
      Exit
    end;
  end;
  Result := TOBP.Detail[ii - 1].GetValue('GPT_' + Champ);
  TOBP.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 11/12/2001
Modifié le ... : 11/12/2001
Description .. : Fct appelable depuis edtion
Suite ........ : qui permet d'avoir le mtt total des ports
Suite ........ : (et non pasligne par ligne  comme la fct GcGetPortsEdt)
Suite ........ : Recoit MOntantPDf ou BasePDF et les info necessaire sà la
Suite ........ : souche
Suite ........ : recoit en plus de cette fct GP_FACTUREHT pour eviter une
Suite ........ : requête SQL en plus.
Suite ........ :
Suite ........ : Fct faite à partir de GcGetPortsEDt
Mots clefs ... : PORT;EDITION;
*****************************************************************}

function GCGetPortsEdtTot(sf, sp: string): Variant;
var Nature, Souche, St, Champ, EntHt: string;
  Numero, Indice, ii: integer;
  QQ: TQuery;
  EnHT: Boolean;
begin
  Result := '';
  St := sp;
  Champ := GCReadParam(st);
  Nature := GCReadParam(St);
  Souche := GCReadParam(St);
  Numero := ValeurI(GCReadParam(St));
  Indice := ValeurI(GCReadParam(St));
  EntHt := GCReadParam(St);
  if (EntHt = 'X') then EnHt := True else EnHt := False;
  QQ := OpenSQL('SELECT sum(GPT_TOTALHTDEV),Sum(GPT_BASEHTDEV),sum(GPT_TOTALTTCDEV),Sum(GPT_BASETTCDEV) FROM PIEDPORT WHERE GPT_NATUREPIECEG="' + Nature +
    '" AND GPT_SOUCHE="' + Souche + '" AND GPT_NUMERO=' + IntToStr(Numero) + ' AND GPT_INDICEG=' + IntToStr(Indice), True,-1,'',true);
  if QQ.EOF then Exit;
  ii := 99;
  if Champ = 'BASEPDF' then
    if EnHT then ii := 1 else ii := 3;
  if Champ = 'MONTANTPDF' then
    if EnHT then ii := 0 else ii := 2;
  if ii > 3 then Exit; // cas où mauvais paramètre
  Result := QQ.Fields[ii].AsFloat;
  Ferme(QQ);
end;

function GCGetDimTitreR(sp: string): Variant;
var St: string;
  CodeGrille: array[1..5] of string;
  Indice: integer;
begin
  Result := '';
  St := sp;
  for indice := 1 to 5 do CodeGrille[Indice] := GCReadParam(st);
  Indice := ValeurI(GCReadParam(st));
  Result := GCGetTitreDimRemplie(CodeGrille[1], CodeGrille[2], CodeGrille[3], CodeGrille[4], CodeGrille[5], Indice);
end;

function GCGetDimValR(sp: string): Variant;
var St: string;
  CodeGrille: array[1..5] of string;
  CodeDim: array[1..5] of string;
  Indice: integer;
begin
  Result := '';
  St := sp;
  for indice := 1 to 5 do CodeGrille[Indice] := GCReadParam(st);
  for indice := 1 to 5 do CodeDim[Indice] := GCReadParam(st);
  Indice := ValeurI(GCReadParam(st));
  Result := GCGetCodeDimRemplie(CodeGrille[1], CodeGrille[2], CodeGrille[3], CodeGrille[4], CodeGrille[5],
    CodeDim[1], CodeDim[2], CodeDim[3], CodeDim[4], CodeDim[5], Indice);
end;

function GCGetDimORLIValR(sp: string): Variant;
var St: string;
  CodeGrille: array[1..5] of string;
  CodeDim: array[1..5] of string;
  Indice: integer;
begin
  Result := '';
  St := sp;
  for indice := 1 to 5 do CodeGrille[Indice] := GCReadParam(st);
  for indice := 1 to 5 do CodeDim[Indice] := GCReadParam(st);
  Indice := ValeurI(GCReadParam(st));
  Result := GCGetCodeDimORLIRemplie(CodeGrille[1], CodeGrille[2], CodeGrille[3], CodeGrille[4], CodeGrille[5],
    CodeDim[1], CodeDim[2], CodeDim[3], CodeDim[4], CodeDim[5], Indice);
end;

function GCGetDimTitre(sp: string): Variant; // Param : Indice de la dimension
var St: string;
  Indice: integer;
begin
  Result := '';
  St := sp;
  Indice := ValeurI(GCReadParam(st));
  Result := GCGetTitreDim(Indice);
end;

function GCGetDimVal(sp: string): Variant; // Param : Code grille + Code dim + indice
var St: string;
  CodeGrille: string;
  CodeDim: string;
  Indice: integer;
begin
  Result := '';
  St := sp;
  CodeGrille := GCReadParam(st);
  CodeDim := GCReadParam(st);
  Indice := ValeurI(GCReadParam(st));
  if (trim(CodeGrille) <> '') and (trim(CodeDim) <> '') then Result := GCGetCodeDim(CodeGrille, CodeDim, Indice);
end;

function GCDimToGen(sp: string): Variant;
var StBlanc, Article: string;
  longueur: integer;
begin
  Article := Trim(sp);
  longueur := length(Article);
  StBlanc := '                  ';
  Article := Article + Copy(StBlanc, 1, 18 - longueur) + '               ' + 'X';
  Result := Article;
end;

function GCRechPieceTiers(sp: string): Variant;
var St, CodeTiers, NaturePiece: string;
  Nombre: integer;
  Q: TQuery;
begin
  Result := 0;
  St := sp;
  CodeTiers := GCReadParam(st);
  NaturePiece := GCReadParam(st);
  Q := OpenSQL('SELECT T_AUXILIAIRE,COUNT(GP_NUMERO)AS NOMBRE FROM TIERS ' +
    'LEFT JOIN PIECE ON GP_TIERS=T_TIERS WHERE T_TIERS="' +
    CodeTiers + '" and GP_NATUREPIECEG="' + NaturePiece + '" GROUP BY T_AUXILIAIRE', true,-1,'',true);
  if not Q.EOF then
  begin
    Nombre := Q.Findfield('NOMBRE').AsInteger;
    result := Nombre;
  end else result := 0;
  Ferme(Q);
end;

function GCRechVteCliDat(sp: string): Variant;
var St, CodeTiers, NaturePiece: string;
  DateDeb, DateFin: TDateTime;
  Mtvte: double;
  Q: TQuery;
begin
  Result := 0;
  St := sp;
  CodeTiers := GCReadParam(st);
  DateDeb := StrToFloat(GCReadParam(st));
  DateFin := StrToFloat(GCReadParam(st));
  NaturePiece := GCReadParam(st);
  Q := OpenSQL('SELECT GP_TIERS,SUM(GP_TOTALTTC)AS MTVTE FROM PIECE ' +
    'WHERE GP_TIERS="' + CodeTiers + '" and GP_NATUREPIECEG="' +
    NaturePiece + '" and GP_DATEPIECE >="' +
    usdatetime(DateDeb) + '" and GP_DATEPIECE <="' +
    usdatetime(DateFin) + '" GROUP BY GP_TIERS', true,-1,'',true);
  if not Q.EOF then
  begin
    Mtvte := Q.Findfield('MTVTE').AsFloat;
    result := Mtvte;
  end else result := 0;
  Ferme(Q);
end;

function GCRechRegCliDat(sp: string): Variant;
var St, CodeTiers, NaturePiece, SansTypReg: string;
  DateDeb, DateFin: TDateTime;
  Mtreg: double;
  Q: TQuery;
begin
  Result := 0;
  St := sp;
  CodeTiers := GCReadParam(st);
  DateDeb := StrToFloat(GCReadParam(st));
  DateFin := StrToFloat(GCReadParam(st));
  NaturePiece := GCReadParam(st);
  SansTypReg := GCReadParam(st);
  Q := OpenSQL('SELECT GP_TIERS,SUM(GPE_MONTANTECHE)AS MTREG FROM PIECE ' +
    'left join PIEDECHE on GP_NATUREPIECEG = GPE_NATUREPIECEG and ' +
    'GP_SOUCHE = GPE_SOUCHE and GP_INDICEG = GPE_INDICEG and ' +
    'GP_NUMERO = GPE_NUMERO LEFT JOIN MODEPAIE ON GPE_MODEPAIE = MP_MODEPAIE WHERE GP_TIERS="' +
    CodeTiers + '" and GP_NATUREPIECEG="' + NaturePiece + '" AND MP_TYPEMODEPAIE<>"' +
    SansTypReg + '" and GP_DATEPIECE >="' + usdatetime(DateDeb) + '" and GP_DATEPIECE <="' +
    usdatetime(DateFin) + '" GROUP BY GP_TIERS', true,-1,'',true);
  if not Q.EOF then
  begin
    Mtreg := Q.Findfield('MTREG').AsFloat;
    result := Mtreg;
  end else result := 0;
  Ferme(Q);
end;

// recherche du libelle d'une famille article

function GCRechTitrefam(sp: string): Variant;
var St: string;
  indice: integer;
begin
  Result := 0;
  St := sp;
  Indice := ValeurI(GCReadParam(st));
  Result := RechDom('GCLIBFAMILLE', 'LF' + InttoStr(Indice), False);
end;

// recherche du libelle d'une zone libre

function GCRechTitreZoneLibre(sp: string): Variant;
var St, TitreZone: string;
begin
  Result := '';
  St := sp;
//  TitreZone := rechdom('GCZONELIBRE', St, False);
	TitreZone := RechDomZoneLibre (St, False);

  if Length(TitreZone) > 0 then if (copy(TitreZone, 1, 2) = '.-') then TitreZone := '';
  Result := TitreZone;
end;

// recherche du libelle de la catégorie des taxes

function GCRechTitreCategTaxe(sp: string): Variant;
var St, TitreCateg: string;
begin
  Result := '';
  St := sp;
  TitreCateg := rechdom('GCCATEGORIETAXE', St, False);
  if Length(TitreCateg) > 0 then if (copy(TitreCateg, 1, 2) = '.-') then TitreCateg := '';
  Result := TitreCateg;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 01/08/2001
Modifié le ... :   /  /
Description .. : Spécifique mode: Tarif détail
Suite ........ : Retourne le tarif spécifique en prix d'un article (période promo)
Mots clefs ... : TARIF;PRIX
*****************************************************************}

function GCRechTarifSpec(sp: string): Variant;
var St, CodeArt, TarifArticle, CodeDepot, RegimePrix,
  InfoRem, TypeTarifEtab, CodeTiers, TarifTiers,
    CodeDevPrecedente, NatureType, VenteAchat, NomChamp, NomChampDev: string;
  DatePrecedente: TDateTime;
  QQ, Q: TQuery;
  Prix: Double;
  TOBRechTarif, TOBMode: TOB;
  EnHT: Boolean;
begin
  St := sp;
  CodeArt := GCReadParam(st);
  TarifArticle := GCReadParam(st);
  CodeDepot := GCReadParam(st);
  DatePrecedente := DateTarif;
  DateTarif := StrToFloat(GCReadParam(st));
  RegimePrix := GCReadParam(st);
  CodeDevPrecedente := CodeDevise;
  CodeDevise := GCReadParam(st);
  CodeTiers := GCReadParam(st);
  TarifTiers := GCReadParam(st);
  if RegimePrix = 'TTC' then
  begin
    VenteAchat := 'VEN';
    NatureType := 'VTE';
    EnHT := False;
    NomChamp := 'ET_TYPETARIF';
    NomChampDev := 'ET_DEVISE';
  end else
  begin
    VenteAchat := 'ACH';
    NatureType := 'ACH';
    EnHT := True;
    NomChamp := 'ET_TYPETARIFACH';
    NomChampDev := 'ET_DEVISEACH';
  end;
  TOBMode := VH_GC.TOBEdt.FindFirst(['_TARFMODE'], ['TarfMode'], false);
  if (TOBMode = nil) or (TOBMode.Detail.count = 0) or (CodeDevise <> CodeDevPrecedente) then
  begin
    TOBMode := TOB.Create('', nil, -1);
    TOBMode.AddChampSup('_TARFMODE', False);
    TOBMode.PutValue('_TARFMODE', 'TarfMode');
    TOBMode.AddChampSup('_TYPETARIFETAB', False);
    QQ := OpenSQL('Select ' + NomChamp + ',' + NomChampDev + ' from ETABLISS Where ET_ETABLISSEMENT="' + CodeDepot + '"', True,-1,'',true);
    if not QQ.EOF then
    begin
      if QQ.FindField(NomChamp).AsString <> '' then TypeTarifEtab := QQ.FindField(NomChamp).AsString;
      TOBMode.PutValue('_TYPETARIFETAB', TypeTarifEtab);
      if CodeDevise = '' then CodeDevise := QQ.FindField(NomChampDev).AsString;
    end else
    begin
      TypeTarifEtab := '...';
      TOBMode.PutValue('_TYPETARIFETAB', TypeTarifEtab);
    end;
    Ferme(QQ);
    Q := OpenSQL('Select GFM_TARFMODE,GFM_TYPETARIF,GFM_NATURETYPE,GFM_DATEDEBUT,GFM_PROMO from TARIFMODE where gfm_typetarif in ("' +
      TOBMode.GetValue('_TYPETARIFETAB') + '","...") and GFM_Naturetype="' + NatureType + '" order by GFM_TYPETARIF DESC,GFM_DATEDEBUT DESC', True,-1,'',true);
    TOBMode.LoadDetailDB('TARIFMODE', '', '', Q, False);
    if TOBMode <> nil then TOBMode.Changeparent(VH_GC.TOBEdt, -1);
    ferme(Q);
  end;
  if CodeDevise = '' then CodeDevise := V_PGI.DevisePivot;
  TOBRechTarif := VH_GC.TOBEdt.FindFirst(['_TARFART'], ['TarfArt'], false);
  if (TOBRechTarif <> nil) then
  begin
    if ((Copy(CodeArticle, 1, 18)) <> (Copy(CodeArt, 1, 18))) or (DateTarif <> DatePrecedente) then
    begin
      TOBRechTarif.Free;
      TOBRechTarif := CreerTobTarifArt(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat, DateTarif, EnHT);
      if TOBRechTarif <> nil then TOBRechTarif.Changeparent(VH_GC.TOBEdt, -1);
      CodeArticle := CodeArt;
    end;
  end else
  begin
    if (TOBRechTarif = nil) then
    begin
      TOBRechTarif := CreerTobTarifArt(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat, DateTarif, EnHT);
      if TOBRechTarif <> nil then TOBRechTarif.Changeparent(VH_GC.TOBEdt, -1);
      CodeArticle := CodeArt;
    end;
  end;
  InfoRem := RechTarifSpec(TOBRechTarif, TOBMode, CodeArt, CodeDepot, CodeDevise, TOBMode.GetValue('_TYPETARIFETAB'), NatureType);
  Prix := Valeur(ReadTokenSt(InfoRem));
  Result := Prix;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 01/08/2001
Modifié le ... :   /  /
Description .. : Spécifique mode: Tarif détail
Suite ........ : Retourne le prix de vente d'un article
Mots clefs ... : TARIF;PRIX
*****************************************************************}

function GCRechTarifPxVente(sp: string): Variant;
var St, CodeArt, TarifArticle, CodeDepot, RegimePrix,
  ArrondiP, InfoRem, TypeTarifEtab, CodeTiers, TarifTiers,
    CodeDevPrecedente, NatureType, VenteAchat, NomChamp, NomChampDev: string;
  DatePrecedente: TDateTime;
  QQ, Q: TQuery;
  Remise, Prix, PrixBase: Double;
  TOBRechTarif, TOBMode: TOB;
  DEV: RDEVISE;
  EnHT: Boolean;
begin
  St := sp;
  CodeArt := GCReadParam(st);
  TarifArticle := GCReadParam(st);
  CodeDepot := GCReadParam(st);
  DatePrecedente := DateTarif;
  DateTarif := StrToFloat(GCReadParam(st));
  RegimePrix := GCReadParam(st);
  CodeDevPrecedente := CodeDevise;
  CodeDevise := GCReadParam(st);
  PrixBase := Valeur(GCReadParam(St));
  CodeTiers := GCReadParam(st);
  TarifTiers := GCReadParam(st);
  if RegimePrix = 'TTC' then
  begin
    VenteAchat := 'VEN';
    NatureType := 'VTE';
    EnHT := False;
    NomChamp := 'ET_TYPETARIF';
    NomChampDev := 'ET_DEVISE';
  end else
  begin
    VenteAchat := 'ACH';
    NatureType := 'ACH';
    EnHT := True;
    NomChamp := 'ET_TYPETARIFACH';
    NomChampDev := 'ET_DEVISEACH';
  end;
  TOBMode := VH_GC.TOBEdt.FindFirst(['_TARFMODE'], ['TarfMode'], false);
  if (TOBMode = nil) or (TOBMode.Detail.count = 0) or (CodeDevise <> CodeDevPrecedente) then
  begin
    TOBMode := TOB.Create('', nil, -1);
    TOBMode.AddChampSup('_TARFMODE', False);
    TOBMode.PutValue('_TARFMODE', 'TarfMode');
    TOBMode.AddChampSup('_TYPETARIFETAB', False);
    QQ := OpenSQL('Select ' + NomChamp + ',' + NomChampDev + ' from ETABLISS Where ET_ETABLISSEMENT="' + CodeDepot + '"', True,-1,'',true);
    if not QQ.EOF then
    begin
      if QQ.FindField(NomChamp).AsString <> '' then TypeTarifEtab := QQ.FindField(NomChamp).AsString;
      TOBMode.PutValue('_TYPETARIFETAB', TypeTarifEtab);
      if CodeDevise = '' then CodeDevise := QQ.FindField(NomChampDev).AsString;
    end else
    begin
      TypeTarifEtab := '...';
      TOBMode.PutValue('_TYPETARIFETAB', TypeTarifEtab);
    end;
    Ferme(QQ);
    Q := OpenSQL('Select GFM_TARFMODE,GFM_TYPETARIF,GFM_NATURETYPE,GFM_DATEDEBUT,GFM_PROMO from TARIFMODE where gfm_typetarif in ("' +
      TOBMode.GetValue('_TYPETARIFETAB') + '","...") and GFM_Naturetype="' + NatureType + '" order by GFM_TYPETARIF DESC,GFM_DATEDEBUT DESC', True,-1,'',true);
    TOBMode.LoadDetailDB('TARIFMODE', '', '', Q, False);
    if TOBMode <> nil then TOBMode.Changeparent(VH_GC.TOBEdt, -1);
    ferme(Q);
  end;
  if CodeDevise = '' then CodeDevise := V_PGI.DevisePivot;
  TOBRechTarif := VH_GC.TOBEdt.FindFirst(['_TARFART'], ['TarfArt'], false);
  if (TOBRechTarif <> nil) then
  begin
    if ((Copy(CodeArticle, 1, 18)) <> (Copy(CodeArt, 1, 18))) or (DateTarif <> DatePrecedente) then
    begin
      TOBRechTarif.Free;
      TOBRechTarif := CreerTobTarifArt(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat, DateTarif, EnHT);
      if TOBRechTarif <> nil then TOBRechTarif.Changeparent(VH_GC.TOBEdt, -1);
      CodeArticle := CodeArt;
    end;
  end else
  begin
    if (TOBRechTarif = nil) then
    begin
      TOBRechTarif := CreerTobTarifArt(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat, DateTarif, EnHT);
      if TOBRechTarif <> nil then TOBRechTarif.Changeparent(VH_GC.TOBEdt, -1);
      CodeArticle := CodeArt;
    end;
  end;
  InfoRem := RechTarifSpec(TOBRechTarif, TOBMode, CodeArt, CodeDepot, CodeDevise, TOBMode.GetValue('_TYPETARIFETAB'), NatureType);
  Prix := Valeur(ReadTokenSt(InfoRem));
  if Prix = 0 then
  begin
    if CodeTiers = '' then CodeTiers := 'ETAT';
    if TarifTiers = '' then TarifTiers := 'ETAT';
    InfoRem := '';
    InfoRem := ChercheMieRem(TOBRechTarif, TobMode, CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDepot, CodeDevise, TOBMode.GetValue('_TYPETARIFETAB'),
      NatureType);
    Remise := Valeur(ReadTokenSt(InfoRem));
    ArrondiP := ReadTokenSt(InfoRem);
    if remise <> 0 then
    begin
      Prix := PrixBase * (1 - (Remise / 100));
      Prix := ArrondirPrix(ArrondiP, Prix);
    end;
  end;
  DEV.Code := CodeDevise;
  GetInfosDevise(DEV);
  if Prix = 0 then Prix := PrixBase;
  Result := Arrondi(Prix, DEV.Decimale);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 01/08/2001
Modifié le ... :   /  /
Description .. : Spécifique mode: Tarif détai
Suite ........ : Retourne une remise d'un article
Mots clefs ... : TARIF;REMISE
*****************************************************************}

function GCRechRemise(sp: string): Variant;
var St, CodeArt, TarifArticle, CodeDepot, RegimePrix,
  InfoRem, TypeTarifEtab, CodeTiers, TarifTiers,
    CodeDevPrecedente, NatureType, VenteAchat, NomChamp, NomChampDev: string;
  DatePrecedente: TDateTime;
  Q, QQ: TQuery;
  Remise: Double;
  TOBRechTarif, TOBMode: TOB;
  EnHT: Boolean;
begin
  St := sp;
  CodeArt := GCReadParam(st);
  TarifArticle := GCReadParam(st);
  CodeDepot := GCReadParam(st);
  DatePrecedente := DateTarif;
  DateTarif := StrToFloat(GCReadParam(st));
  RegimePrix := GCReadParam(st);
  CodeDevPrecedente := CodeDevise;
  CodeDevise := GCReadParam(st);
  CodeTiers := GCReadParam(st);
  TarifTiers := GCReadParam(st);
  if RegimePrix = 'TTC' then
  begin
    VenteAchat := 'VEN';
    NatureType := 'VTE';
    EnHT := False;
    NomChamp := 'ET_TYPETARIF';
    NomChampDev := 'ET_DEVISE';
  end else
  begin
    VenteAchat := 'ACH';
    NatureType := 'ACH';
    EnHT := True;
    NomChamp := 'ET_TYPETARIFACH';
    NomChampDev := 'ET_DEVISEACH';
  end;
  TOBMode := VH_GC.TOBEdt.FindFirst(['_TARFMODE'], ['TarfMode'], false);
  if (TOBMode = nil) or (TOBMode.Detail.count = 0) or (CodeDevise <> CodeDevPrecedente) then
  begin
    TOBMode := TOB.Create('', nil, -1);
    TOBMode.AddChampSup('_TARFMODE', False);
    TOBMode.PutValue('_TARFMODE', 'TarfMode');
    TOBMode.AddChampSup('_TYPETARIFETAB', False);
    QQ := OpenSQL('Select ' + NomChamp + ',' + NomChampDev + ' from ETABLISS Where ET_ETABLISSEMENT="' + CodeDepot + '"', True,-1,'',true);
    if not QQ.EOF then
    begin
      if QQ.FindField(NomChamp).AsString <> '' then TypeTarifEtab := QQ.FindField(NomChamp).AsString;
      TOBMode.PutValue('_TYPETARIFETAB', TypeTarifEtab);
      if CodeDevise = '' then CodeDevise := QQ.FindField(NomChampDev).AsString;
    end else
    begin
      TypeTarifEtab := '...';
      TOBMode.PutValue('_TYPETARIFETAB', TypeTarifEtab);
    end;
    Ferme(QQ);
    Q := OpenSQL('Select GFM_TARFMODE,GFM_TYPETARIF,GFM_NATURETYPE,GFM_DATEDEBUT,GFM_PROMO from TARIFMODE where gfm_typetarif in ("' +
      TOBMode.GetValue('_TYPETARIFETAB') + '","...") and GFM_Naturetype="' + NatureType + '" order by GFM_TYPETARIF DESC,GFM_DATEDEBUT DESC', True,-1,'',true);
    TOBMode.LoadDetailDB('TARIFMODE', '', '', Q, False);
    if TOBMode <> nil then TOBMode.Changeparent(VH_GC.TOBEdt, -1);
    ferme(Q);
  end;
  if CodeDevise = '' then CodeDevise := V_PGI.DevisePivot;
  TOBRechTarif := VH_GC.TOBEdt.FindFirst(['_TARFART'], ['TarfArt'], false);
  if (TOBRechTarif <> nil) then
  begin
    if ((Copy(CodeArticle, 1, 18)) <> (Copy(CodeArt, 1, 18))) or (DateTarif <> DatePrecedente) then
    begin
      TOBRechTarif.Free;
      TOBRechTarif := CreerTobTarifArt(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat, DateTarif, EnHT);
      if TOBRechTarif <> nil then TOBRechTarif.Changeparent(VH_GC.TOBEdt, -1);
      CodeArticle := CodeArt;
    end;
  end else
  begin
    if (TOBRechTarif = nil) then
    begin
      if CodeTiers = '' then CodeTiers := 'ETAT';
      if TarifTiers = '' then TarifTiers := 'ETAT';
      TOBRechTarif := CreerTobTarifArt(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat, DateTarif, EnHT);
      if TOBRechTarif <> nil then TOBRechTarif.Changeparent(VH_GC.TOBEdt, -1);
      CodeArticle := CodeArt;
    end;
  end;
  InfoRem := ChercheMieRem(TOBRechTarif, TobMode, CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDepot, CodeDevise, TOBMode.GetValue('_TYPETARIFETAB'),
    NatureType);
  Remise := Valeur(ReadTokenSt(InfoRem));
  Result := Remise;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 01/08/2001
Modifié le ... :   /  /
Description .. : Spécifique mode: Tarif détail
Suite ........ : Retourne un prix de base (tarif ou prix article) sur lequel sera
Suite ........ : appliqué une remise
Mots clefs ... : TARIF
*****************************************************************}

function GCRechTarifBAS(sp: string): Variant;
var St, CodeArt, CodeDepot, TarifArticle, CodeGen, RegimePrix,
  CodeTiers, TarifTiers, TypeTarifEtab,
    CodeDevPrecedente, NatureType, VenteAchat, NomChamp, NomChampDev: string;
  Q, QQ: TQuery;
  TOBMode, TOBRechTarif, TOBM: Tob;
  PrixBase, PrixArt, PrixArtTTC, PrixArtHT, Prix: Double;
  DatePrecedente: TDateTime;
  DEV: RDEVISE;
  EnHT: Boolean;
begin
  Result := 0;
  St := sp;
  CodeArt := GCReadParam(st);
  TarifArticle := GCReadParam(st);
  CodeDepot := GCReadParam(st);
  DatePrecedente := DateTarif;
  DateTarif := StrToFloat(GCReadParam(st));
  RegimePrix := Trim(GCReadParam(st));
  CodeDevPrecedente := CodeDevise;
  CodeDevise := Trim(GCReadParam(st));
  PrixArtTTC := Valeur(GCReadParam(St));
  PrixArtHT := Valeur(GCReadParam(St));
  CodeTiers := GCReadParam(st);
  TarifTiers := GCReadParam(st);
  if RegimePrix = 'TTC' then
  begin
    VenteAchat := 'VEN';
    NatureType := 'VTE';
    EnHT := False;
    NomChamp := 'ET_TYPETARIF';
    NomChampDev := 'ET_DEVISE';
    PrixArt := PrixArtTTC;
  end else
  begin
    VenteAchat := 'ACH';
    NatureType := 'ACH';
    EnHT := True;
    NomChamp := 'ET_TYPETARIFACH';
    NomChampDev := 'ET_DEVISEACH';
    PrixArt := PrixArtHT;
  end;
  CodeGen := CodeArticleUnique(Copy(CodeArt, 1, 18), '', '', '', '', '');
  TOBMode := VH_GC.TOBEdt.FindFirst(['_TARFMODE'], ['TarfMode'], false);
  if (TOBMode = nil) or (TOBMode.Detail.count = 0) or (CodeDevise <> CodeDevPrecedente) then
  begin
    TOBMode := TOB.Create('', nil, -1);
    TOBMode.AddChampSup('_TARFMODE', False);
    TOBMode.PutValue('_TARFMODE', 'TarfMode');
    TOBMode.AddChampSup('_TYPETARIFETAB', False);
    QQ := OpenSQL('Select ' + NomChamp + ',' + NomChampDev + ' from ETABLISS Where ET_ETABLISSEMENT="' + CodeDepot + '"', True,-1,'',true);
    if not QQ.EOF then
    begin
      if QQ.FindField(NomChamp).AsString <> '' then TypeTarifEtab := QQ.FindField(NomChamp).AsString;
      TobMode.PutValue('_TYPETARIFETAB', TypeTarifEtab);
      if CodeDevise = '' then CodeDevise := QQ.FindField(NomChampDev).AsString;
    end else
    begin
      TypeTarifEtab := '...';
      TobMode.PutValue('_TYPETARIFETAB', TypeTarifEtab);
    end;
    Ferme(QQ);
    Q := OpenSQL('Select GFM_TARFMODE,GFM_TYPETARIF,GFM_DATEDEBUT,GFM_NATURETYPE,GFM_PROMO from TARIFMODE where gfm_typetarif in ("' +
      TOBMode.GetValue('_TYPETARIFETAB') + '","...") and GFM_Naturetype="' + NatureType + '" order by GFM_TYPETARIF DESC,GFM_DATEDEBUT DESC', True,-1,'',true);
    TOBMode.LoadDetailDB('TARIFMODE', '', '', Q, False);
    if TOBMode <> nil then TOBMode.Changeparent(VH_GC.TOBEdt, -1);
    ferme(Q);
  end;
  if CodeDevise = '' then CodeDevise := V_PGI.DevisePivot;
  TOBRechTarif := VH_GC.TOBEdt.FindFirst(['_TARFART'], ['TarfArt'], false);
  if (TOBRechTarif <> nil) then
  begin
    if ((Copy(CodeArticle, 1, 18)) <> (Copy(CodeArt, 1, 18))) or (DateTarif <> DatePrecedente) then
    begin
      TOBRechTarif.Free;
      TOBRechTarif := CreerTobTarifArt(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat, DateTarif, EnHT);
      if TOBRechTarif <> nil then TOBRechTarif.Changeparent(VH_GC.TOBEdt, -1);
      CodeArticle := CodeArt;
    end;
  end else
  begin
    if (TOBRechTarif = nil) then
    begin
      TOBRechTarif := CreerTobTarifArt(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat, DateTarif, EnHT);
      if TOBRechTarif <> nil then TOBRechTarif.Changeparent(VH_GC.TOBEdt, -1);
      CodeArticle := CodeArt;
    end;
  end;
  PrixBase := RechPrixTarifBase(TOBRechTarif, TOBMode, CodeArt, CodeDepot, CodeDevise, TOBMode.GetValue('_TYPETARIFETAB'), NatureType);
  DEV.Code := CodeDevise;
  GetInfosDevise(DEV);
  if PrixBase = 0 then
  begin
    TOBM := TOBMode.FindFirst(['GFM_DEVISE'], [CodeDevise], False);
    if TOBM <> nil then PrixArt := PrixArt * TOBM.GetValue('GFM_COEF');
    Prix := PrixArt;
  end else Prix := PrixBase;
  Result := Arrondi(Prix, DEV.Decimale);
end;

function RechLibDimensions(sp: string): Variant;
var stDim: string;
  i_ind: integer;
  Sep: string;
  GrilleDim, CodeDim, LibDim: string;
begin
  Result := '';
  Sep := '';
  stDim := '';
  for i_ind := 1 to MaxDimension do
  begin
    GrilleDim := GCReadParam(sp);
    CodeDim := GCReadParam(sp);
    if GrilleDim <> '' then
    begin
      LibDim := GCGetCodeDim(GrilleDim, CodeDim, i_ind);
      if LibDim <> '' then stDim := stDim + Sep + LibDim;
    end;
    Sep := ' - ';
  end;
  Result := stDim;
end;

{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 07/05/2002
Modifié le ... :   /  /
Description .. : Fct qui permet depuis un état d'utiliser la fct PlusDate de
Suite ........ : l'AGL
Mots clefs ... : PLUSDATE
*****************************************************************}

function GCPlusDAte(sp: string): Variant;
var St: string;
  interv: integer;
  Date: TdateTime;
begin
  Result := 0;
  Date := StrTofloat(GcReadParam(SP));
  Interv := strtoInt(GCReadParam(sP));
  st := GCReadParam(sp);
  result := PlusDate(Date, Interv, St);
end;

// recherche du mois pour etat des ventes

function GCRechMois(sp: string): Variant;
var St: string;
  Moisdeb, NbMois, i, mois: integer;
begin
  Result := 0;
  St := sp;
  Moisdeb := strtoint(GCReadParam(st));
  NbMois := strtoint(GCReadParam(st));
  i := strtoint(GCReadParam(st));
  mois := Moisdeb - Nbmois + i;
  if mois < 1 then
    mois := mois + 12;
  if mois > 12 then
    mois := mois - 12;
  result := ShortMonthNames[mois];
end;

function GCGetNumOrdre(sp: string): Variant;
var i: integer;
  TOBN, TOBT: TOB;
  Find: Boolean;
  StRech: string;
begin
  Result := -1;
  TOBN := nil;
  Find := False;
  {rechercher une fille dont un champsup s'appelle du nom de la fonction}
  for i := 0 to VH_GC.TOBEdt.Detail.Count - 1 do
  begin
    TOBN := VH_GC.TOBEdt.Detail[i];
    if TOBN.FieldExists('GCNUMORDRE') then
    begin
      Find := True;
      Break;
    end;
  end;
  {Si pas trouvée la créer}
  if not Find then
  begin
    TOBN := TOB.Create('', VH_GC.TOBEdt, -1);
    TOBN.AddChampSup('GCNUMORDRE', False);
  end;
  {Recherche dans le détail de la fille d'un tob qui contient déjà la chaine de recherche}
  StRech := uppercase(Trim(Sp));
  TOBT := TOBN.FindFirst(['CHAMPRECHERCHE'], [StRech], True);
  if TOBT <> nil then Result := TOBT.GetIndex + 1 else
  begin
    {Si pas trouvée la créer}
    TOBT := TOB.Create('', TOBN, -1);
    TOBT.AddChampSup('CHAMPRECHERCHE', False);
    TOBT.PutValue('CHAMPRECHERCHE', StRech);
    Result := TOBN.Detail.Count;
  end;
  {Le résultat est le numéro d'ordre}
end;

function GCVideNumOrdre: Variant;
var i: integer;
  TOBN: TOB;
  Find: Boolean;
begin
  Result := True;
  TOBN := nil;
  Find := False;
  {rechercher une fille dont un champsup s'appelle du nom de la fonction}
  for i := 0 to VH_GC.TOBEdt.Detail.Count - 1 do
  begin
    TOBN := VH_GC.TOBEdt.Detail[i];
    if TOBN.FieldExists('GCNUMORDRE') then
    begin
      Find := True;
      Break;
    end;
  end;
  if ((Find) and (TOBN <> nil)) then TOBN.ClearDetail;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Corinne TARDY
Créé le ...... : 27/08/2001
Modifié le ... :   /  /
Description .. : Spécifique mode: Permet de décomposer un champ qui serait le regroupement de plusieurs,
Suite ........ : renvoi la valeur voulu. Il suffit d'indiquer le nom de la valeur à extraire
Paramètres ... : Colonne souhaité ; Type de la colonne ; chaine à décomposer
Mots clefs ... :
*****************************************************************}

function GCRechValeurColonne(sp: string): Variant;
var
  Chaine, Valeur, TypeColonne: string;
  NumColonne, i_ind1: integer;
begin
  Result := 0;
  // Récupération du numero de la colonne voulu
  NumColonne := StrToInt(GCReadParam(sp));
  // Récupération du type de la colonne, Surtout si c'est une date
  TypeColonne := GCReadParam(sp);
  Chaine := sp;
  // Je récupère la valeur de la colonne voulu
  for i_ind1 := 0 to NumColonne - 1 do
    Valeur := ReadTokenSt(Chaine);
  // Si c'est une date je transforme la chaine en date
  if (TypeColonne = 'Date') or (TypeColonne = 'DATE') then
    result := EvalueDate(Valeur)
  else result := Valeur;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Corinne TARDY
Créé le ...... : 19/09/2001
Modifié le ... :   /  /
Description .. : Spécifique mode:Permet de renvoyer la quantité disponible pour
Suite ........ : un article (dimensionné ou générique) avec ou sans établissement
Paramètres ... : Type de l'article,L'article,L'établissement
Mots clefs ... : DISPO;QTE
*****************************************************************}

function GCRechDispoArticle(sp: string): Variant;
var TypeArticle, Article, Depot, CodeArt, st: string;
  TOBRechDispo: TOB;
begin
  Result := 0;
  st := sp;
  TypeArticle := GCReadParam(st);
  CodeArt := CodeArtDispo; // Ceci me permet de conservé le code article précédent
  Article := GCReadParam(st);
  Depot := GCReadParam(st);
  CodeArtDispo := Trim(Copy(Article, 1, 18));

  TOBRechDispo := VH_GC.TOBEdt.FindFirst(['_DISPOART'], ['DispoArt'], false);
  if (TOBRechDispo <> nil) then
  begin // Si la TOB n'est pas vide je ne la recrée qui si le code article est différent
    if (CodeArtDispo <> CodeArt) then
    begin
      TOBRechDispo.free;
      // Création de la tob dispo
      TOBRechDispo := CreerTobDispoArt(CodeArtDispo);
      if TOBRechDispo <> nil then TOBRechDispo.Changeparent(VH_GC.TOBEdt, -1);
    end;
  end
  else
  begin // Si la TOB est vide je la crée
    if (TOBRechDispo = nil) then
    begin
      TOBRechDispo.free;
      // Création de la tob dispo
      TOBRechDispo := CreerTobDispoArt(CodeArtDispo);
      if TOBRechDispo <> nil then TOBRechDispo.Changeparent(VH_GC.TOBEdt, -1);
    end;
  end;
  // Je calcul la quantité à retourné suivant les cas
  if (Depot = '') then
  begin
    if (TypeArticle = 'GEN') then // somme de toutes les quantités pour un code article
      Result := TOBRechDispo.Somme('GQ_PHYSIQUE', ['GA_CODEARTICLE'], [CodeArtDispo], False, False)
    else // somme de toutes les quantités pour un article dimensionné
      Result := TOBRechDispo.Somme('GQ_PHYSIQUE', ['GA_ARTICLE'], [Article], False, False);
  end
  else
  begin
    if (TypeArticle = 'GEN') then // somme de toutes les quantités pour un code article et un dépot
      Result := TOBRechDispo.Somme('GQ_PHYSIQUE', ['GA_CODEARTICLE', 'GQ_DEPOT'], [CodeArtDispo, Depot], False, False)
    else // somme de toutes les quantités pour un article dimensionné et un dépot
      Result := TOBRechDispo.Somme('GQ_PHYSIQUE', ['GA_ARTICLE', 'GQ_DEPOT'], [Article, Depot], False, False);
  end;
end;

function GCGetDevisePiece(sp: string): Variant;
var st, Devise, SaisieContre: string;
begin
  st := sp;
  Devise := GCReadParam(st);
  SaisieContre := GCReadParam(st);
  if SaisieContre = 'X' then
  begin
    if VH^.TenueEuro then Result := V_PGI.DeviseFongible else Result := 'EUR';
  end
  else Result := Devise;
end;

function GCCodeDate(sp: string): Variant;
var st, stDate, stCodage: string;
  stMois: char;
begin
  st := sp;
  stDate := GCReadParam(st);
  stCodage := GCReadParam(st);
  if stCodage = '1' then
  begin // 13/12/2001 -> 13L (format JJ+X avec X=A..L pour le mois 1..12)
    stMois := chr(64 + (StrToInt(copy(stDate, 4, 2))));
    Result := copy(stDate, 1, 2) + stMois;
  end
  else Result := TraduireMemoire('Codage non défini');
end;

function GCGetEncours(sp: string): Variant;
var st, stDB, stTiers, stNature, stSql: string;
  QQ: TQuery;
begin
  st := sp;
  stDB := GCReadParam(st); // Nom de la base de données
  stTiers := GCReadParam(st); // Code tiers
  stNature := GCReadParam(st); // Nature de pièce

  stSql := '@@select SUM(GP_TOTALTTC) from ';
  if stDB <> '' then stSql := stSql + stDB + '.dbo.';
  stSql := stSql + 'PIECE where GP_TIERS="' + stTiers + '" and GP_NATUREPIECEG="' + stNature + '" and GP_VIVANTE="X"';

  QQ := OpenSQL(stSql, True,-1,'',true);
  if not QQ.Eof then Result := QQ.Fields[0].AsFloat else Result := 0.0;
  Ferme(QQ);

end;

{Paramètres : tokenString = 'NomBase;CodeTiers;NaturePièce;CodeChamp'
  avec NomBase : Nom de la BDD interrogée (peut-être <> BDD courante sur SQLSERVER,
       CodeTiers : Code du tiers interrogé,
       NaturePièce : Nature de pîèce interrogée,
       CodeChamp : '1' retourne le numéro de la pièce,
                   '2' retourne la date de la pièce.
 Attention : le 1er appel doit utiliser le CodeChamp '1' !
 Restriction : interrogation autre base possible sur SQLSERVER uniquement !
}

function GCGetDernierDoc(sp: string): Variant;
var st, stDB, stTiers, stNature, stChamp, stSql: string;
  QQ: TQuery;
begin
  st := sp;
  stDB := GCReadParam(st); // Nom de la base de données
  stTiers := GCReadParam(st); // Code tiers
  stNature := GCReadParam(st); // Nature de pièce
  stChamp := GCReadParam(st); // Champ retourné

  if stChamp = '1' then
  begin
    if stDB <> '' then stSql := '@@' else stSql := '';
    stSql := stSql + 'select GP_DATEPIECE, max(GP_NUMERO) from ';
    if stDB <> '' then stSql := stSql + stDB + '.dbo.';
    stSql := stSql + 'PIECE where GP_TIERS="' + stTiers + '" and GP_NATUREPIECEG="' + stNature
      + '" and GP_DATEPIECE = (select max(GP_DATEPIECE) from ';
    if stDB <> '' then stSql := stSql + stDB + '.dbo.';
    stSql := stSql + 'PIECE where GP_TIERS="' + stTiers + '" and GP_NATUREPIECEG="' + stNature
      + '") group by GP_DATEPIECE';
    QQ := OpenSQL(stSql, True,-1,'',true);
    if not QQ.Eof then
    begin
      DateDernierDoc := QQ.Fields[0].AsString;
      Result := QQ.Fields[1].AsString;
    end
    else
    begin
      DateDernierDoc := ' ';
      Result := ' '
    end;
    Ferme(QQ);
  end
  else if stChamp = '2' then Result := DateDernierDoc;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Didier Carret
Créé le ...... : 04/06/2002
Modifié le ... : 17/02/2003 Option exclusion écarts d'inventaire
Modifié le ... : 28/03/2003 Option mvts du jour de la clôture
Description .. : Paramètre = tokenString :
Suite ........ : - Statut de l'article, obligatoire
Suite ........ : - Identifiant article, obligatoire ou code article
Suite ........ : - Code article, obligatoire ou identifiant article
Suite ........ : - Date arrêté de stock au format JJ/MM/AAAA, obligatoire
Suite ........ : - Dépôt, à défaut utilisation du dépôt par défaut
Suite ........ : - Méthode, 1 = calcul à partir de la 1ère clôture postérieure
Suite ........ : (par défaut)
Suite ........ :           2 = calcul à partir du stock actuel
Suite ........ : - Mouvements du jour de l'arrêté inclus "-"/"X", par défaut
Suite ........ : "-"
Suite ........ : - Mouvements du jour de la clôture de stock inclus "-"/"X", par défaut
Suite ........ : "-"
Suite ........ : - Ecarts d'inventaire inclus "-"/"X", par défaut "-"
Mots clefs ... : STOCK;DATE;ARRETE
*****************************************************************}
{$IFDEF MODE}

function GCStockDate(sp: string): string;
var st, stStatutArt, stArticle, stCodeArticle, stJJMMAAAA, stDepot, stDepart, stMvtJDinclus, stMvtJFinclus, stEcartINVinclus: string;
  Depart, iTob: integer;
  QteStock: integer;
  MvtJDinclus, MvtJFinclus, EcartINVinclus: boolean;
  TobArreteStock: Tob;
begin
  st := sp;
  stStatutArt := GCReadParam(st);
  stArticle := GCReadParam(st);
  stCodeArticle := GCReadParam(st);
  stJJMMAAAA := GCReadParam(st);
  stDepot := GCReadParam(st);
  if stDepot = '' then stDepot := '???';
  stDepart := GCReadParam(st);
  if stDepart = '' then Depart := 1 else Depart := StrToInt(StDepart);
  stMvtJDinclus := GCReadParam(st);
  MvtJDinclus := (stMvtJDinclus = 'X');
  stMvtJFinclus := GCReadParam(st);
  MvtJFinclus := (stMvtJFinclus = 'X');
  stEcartINVinclus := GCReadParam(st);
  EcartINVinclus := (stEcartINVinclus = 'X');
  TobArreteStock := TOB.Create('DISPO', nil, -1);
  QteStock := 0;

  if RecalculStockDate(stStatutArt, stArticle, stCodeArticle, stJJMMAAAA, stDepot, Depart, MvtJDinclus, MvtJFinclus, EcartINVinclus, TobArreteStock) then
  begin
    for iTob := 0 to TobArreteStock.Detail.Count - 1 do
      QteStock := QteStock + TobArreteStock.Detail[iTob].GetValue('GQ_PHYSIQUE');
    Result := IntToStr(QteStock);
  end
  else Result := '?';
  TobArreteStock.Free;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Corinne TARDY
Créé le ...... : 08/07/2002
Modifié le ... :   /  /
Description .. : Permet de renvoyer une chaine de caractère contenant l'ensemble
Suite ........ : des évènements pour une journée de vente
Paramètres ... : le numéro de l'évènement
Mots clefs ... :
*****************************************************************}

function GCRechEvenement(sp: string): string;
var st, stEvt, numEvt: string;
  QEVT: TQuery;
begin
  st := sp;
  result := '';
  numEvt := GCReadParam(st);
  stEvt := 'select CC_LIBELLE from joursetabevt '
    + 'left join choixcod on CC_TYPE="GFE" and CC_CODE=GET_CODEEVENT '
    + 'where GET_NOEVENT="' + numEvt + '"';
  QEVT := OpenSQL(stEvt, True,-1,'',true);
  while not QEVT.EOF do
  begin
    if result <> '' then result := result + ', ' else result := 'Evènements : ';
    result := result + QEVT.Findfield('CC_LIBELLE').asstring;
    QEVT.Next;
  end;
end;

function GCCalcOLEEtat(sf, sp: string): Variant;
begin
  {$IFNDEF GCGC}
  if copy(sf, 1, 2) <> 'GC' then
  begin
    Result := CalcOLEEtat(sf, sp);
    Exit;
  end;
  {$ENDIF}
  {$IFDEF GRC}
  if copy(sf, 1, 2) = 'RT' then
  begin
    Result := RTCalcOLEEtat(sf, sp);
    Exit;
  end;
  {$ENDIF}
  {$IFDEF FOS5}
  if copy(sf, 1, 2) = 'FO' then Result := FOCalcOLEEtat(sf, sp);
  {$ENDIF}
  if sf = 'SEMAINE' then Result := NumSemaine(StrToFloat(GCReadParam(sp)))
  else if sf = 'PLUSDATE' then Result := GCPlusDate(sp)
  else if sf = 'MOIS' then Result := NumMois(StrToFloat(GCReadParam(sp)))
  else if sf = 'GCECHE' then Result := GCGetEchesEdt(sf, sp)
  else if sf = 'GCNUMORDRE' then Result := GCGetNumOrdre(sp)
  else if sf = 'GCVIDENUMORDRE' then Result := GCVideNumOrdre
  else if copy(sf, 1, 8) = 'GCPDFTOT' then Result := GCGetPortsEdtTot(sf, sp)
  else if copy(sf, 1, 5) = 'GCPDF' then Result := GCGetPortsEdt(sf, sp)
  else if copy(sf, 1, 6) = 'GCBASE' then Result := GCGetBasesEdt(sf, sp)
  else if copy(sf, 1, 11) = 'GCDIMTITRER' then Result := GCGetDimTitreR(sp)
  else if copy(sf, 1, 10) = 'GCDIMTITRE' then Result := GCGetDimTitre(sp)
  else if copy(sf, 1, 9) = 'GCDIMVALR' then Result := GCGetDimValR(sp)
  else if copy(sf, 1, 13) = 'GCDIMORLIVALR' then Result := GCGetDimORLIValR(sp)
  else if copy(sf, 1, 8) = 'GCDIMVAL' then Result := GCGetDimVal(sp)
  else if copy(sf, 1, 10) = 'GCDIMTOGEN' then Result := GCDimToGen(sp)
  else if copy(sf, 1, 13) = 'GCEXISTEPIECE' then Result := GCRechPieceTiers(sp)
  else if copy(sf, 1, 15) = 'GCRECHVTECLIDAT' then Result := GCRechVteCliDat(sp)
  else if copy(sf, 1, 15) = 'GCRECHREGCLIDAT' then Result := GCRechRegCliDat(sp)
  else if copy(sf, 1, 12) = 'GCRECHTITFAM' then Result := GCRechTitreFam(sp)
  else if copy(sf, 1, 15) = 'GCRECHTARIFSPEC' then Result := GCRechTarifSpec(sp)
  else if copy(sf, 1, 11) = 'GCPIECEPREC' then Result := GCGetPiecePrecedenteEdt(sp)
  else if copy(sf, 1, 10) = 'GCRECHMOIS' then Result := GCRechMois(sp)
  else if copy(sf, 1, 12) = 'GCRECHLIBDIM' then Result := RechLibDimensions(sp)
  else if copy(sf, 1, 12) = 'GCRECHTARBAS' then Result := GCRechTarifBAS(sp)
  else if copy(sf, 1, 12) = 'GCRECHREMISE' then Result := GCRechRemise(sp)
  else if copy(sf, 1, 18) = 'GCRECHTARIFPXVENTE' then Result := GCRechTarifPxVente(sp)
  else if copy(sf, 1, 18) = 'GCRECHTITZONELIBRE' then Result := GCRechTitreZoneLibre(sp)
  else if copy(sf, 1, 18) = 'GCRECHTITCATEGTAXE' then Result := GCRechTitreCategTaxe(sp)
  else if copy(sf, 1, 18) = 'GCRECHVALCOLONNE' then Result := GCRechValeurColonne(sp)
  else if copy(sf, 1, 14) = 'GCRECHDISPOART' then Result := GCRechDispoArticle(sp)
  else if copy(sf, 1, 16) = 'GCGETDEVISEPIECE' then Result := GCGetDevisePiece(sp)
  else if copy(sf, 1, 10) = 'GCCODEDATE' then Result := GCCodeDate(sp)
  else if copy(sf, 1, 12) = 'GCGETENCOURS' then Result := GCGetEncours(sp)
  else if copy(sf, 1, 15) = 'GCGETDERNIERDOC' then Result := GCGetDernierDoc(sp)
  else if copy(sf, 1, 13) = 'GCCALCPRIXART' then Result := GCCalculePrixArticleEdt(sp)
    {$IFDEF MODE}
  else if copy(sf, 1, 11) = 'GCSTOCKDATE' then Result := GCStockDate(sp)
    {$ENDIF}
  else if copy(sf, 1, 15) = 'GCRECHEVENEMENT' then Result := GCRechEvenement(sp)

    {$IFDEF CCS3}
  else Result := '';
  {$ELSE}
    {$IFDEF AFFAIRE}
  else Result := AFCalcOLEEtat(sf, sp); // si la fx n'est pas trouvée, on regarde sur affaire
  {$ENDIF}
  {$ENDIF}

end;

end.
