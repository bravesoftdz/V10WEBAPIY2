unit GpLightSeria;

interface

Uses     Classes, Sysutils, Messages,
         Dialogs, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
         MaineAGL, eFichList,
{$ELSE}
         db,dbTables,Fe_Main,fichlist,
{$ENDIF}
         HCtrls, UTob, wJetons, HDB;

function  SerialiseGpLight : boolean;
function  InitCircuit : boolean;
function  InitDetailCircuit : boolean;
function  InitPhase : boolean;
function  InitItineraire : boolean;
function  InitPhaseItineraire : boolean;
function  InitSite : boolean;
function  InitNatureTravail : boolean;
function  InitUnites : boolean;
function  InitJetons : boolean;
function  InitAutomatisme : boolean;
function  InitNomenclature : boolean;

function CreateEntete(TobEntete : TOB; TobWNT : TOB) : boolean;
function CreateLigne(TobWNT, TobLignes : TOB) : boolean;

implementation

function  SerialiseGpLight : boolean;
begin
//Result := False;
Result := InitCircuit;
if Result then
  Result := InitDetailCircuit;
if Result then
  Result := InitPhase;
if Result then
  Result := InitItineraire;
if Result then
  Result := InitPhaseItineraire;
if Result then
  Result := InitSite;
if Result then
  Result := InitNatureTravail;
if Result then
  Result := InitUnites;
if Result then
  Result := InitJetons;
if Result then
  Result := InitNomenclature;
end;
//------------------------------------------------------------------------//
function  InitCircuit : boolean;
var stSQL, Depot : string;
    TobDepot : TOB;
    ind1 : integer;

begin
  Result := True;
  // Initialisation de la table CIRCUIT : un circuit AGC par dépot defini dans la base
  stSQL := 'Delete from QCIRCUIT';
  ExecuteSQL(stSQL);
  TobDepot := TOB.Create('DEPOTS', nil, -1);
  stSQL := 'Select GDE_DEPOT from DEPOTS';
  TobDepot.LoadDetailDBFromSQL('DEPOTS', stSQL);
  for ind1 := 0 to TobDepot.Detail.Count - 1 do
  begin
      Depot := TobDepot.Detail[ind1].GetValue('GDE_DEPOT');
      stSQL := 'Insert into QCIRCUIT (QCI_CTX, QCI_CIRCUIT, QCI_CIRCUITLIB, QCI_CODITI, QCI_PRIVILEGIE, QCI_DATEMODIF)' +
               ' values ("0", "' + Depot + '", "Assemblage", "AGC", "-", "' + UsDateTime(Date) + '")';
      Result := (ExecuteSQL(stSQL) = 1);
      if not Result then Break;
  end;
  TobDepot.Free;
end;
//------------------------------------------------------------------------//
function  InitDetailCircuit : boolean;
var stSQL : string;
    TobCircuit : TOB;
    ind1 : integer;

begin
  Result := True;
  // Initialisation de la table Detail CIRCUIT : une ligne de detail par circuit
  stSQL := 'Delete from QDETCIRC';
  ExecuteSQL(stSQL);
  TobCircuit := TOB.Create('QCIRCUIT', nil, -1);
  stSQL := 'Select QCI_CTX, QCI_CIRCUIT from QCIRCUIT';
  TobCircuit.LoadDetailDBFromSQL('QCIRCUIT', stSQL);
  for ind1 := 0 to TobCircuit.Detail.Count - 1 do
  begin
      stSQL := 'Insert into QDETCIRC (QDE_CTX, QDE_CIRCUIT, QDE_OPECIRC, QDE_POLE, QDE_SITE, QDE_GRP) ' +
               'values ("' + TobCircuit.Detail[ind1].GetValue('QCI_CTX') +
               '", "' + TobCircuit.Detail[ind1].GetValue('QCI_CIRCUIT') +
               '", "110", "%", "' + TobCircuit.Detail[ind1].GetValue('QCI_CIRCUIT') + '", "%")';
      Result := (ExecuteSQL(stSQL) = 1);
      if not Result then Break;
  end;
  TobCircuit.Free;
end;
//------------------------------------------------------------------------//
function  InitPhase : boolean;
var stSQL : string;

begin
//  Result := True;
  // Initialisation de la table PHASE : une phase AGC par defaut
  stSQL := 'Delete from QPHASE';
  ExecuteSQL(stSQL);
  stSQL := 'Insert into QPHASE (QPH_CTX, QPH_PHASE, QPH_PHASELIB) ' +
           'values ("0", "AGC", "Assemblage")';
  Result := (ExecuteSQL(stSQL) = 1);
end;
//------------------------------------------------------------------------//
function  InitItineraire : boolean;
var stSQL : string;

begin
//  Result := True;
  // Initialisation de la table ITINERAIRE : un Itineraire AGC standard
  stSQL := 'Delete from QITI';
  ExecuteSQL(stSQL);
  stSQL := 'Insert into QITI (QIT_CTX, QIT_CODITI, QIT_CODITILIB) ' +
           'values ("0", "AGC", "Assemblage")';
  Result := (ExecuteSQL(stSQL) = 1);
end;
//------------------------------------------------------------------------//
function  InitPhaseItineraire : boolean;
var stSQL : string;

begin
//  Result := True;
  // Initialisation de la table PHASEITI : une phase 110 Par defaut sur l'itineraire AGC
  stSQL := 'Delete from QPHASEITI';
  ExecuteSQL(stSQL);
  stSQL := 'Insert into QPHASEITI (QP_CTX, QP_CODITI, QP_OPEITI, QP_PHASE, QP_OPEITILIB) ' +
           'values ("0", "AGC", "110", "AGC", "Assemblage")';
  Result := (ExecuteSQL(stSQL) = 1);
end;
//------------------------------------------------------------------------//
function  InitSite : boolean;
var stSQL, Depot : string;
    TobDepot : TOB;
    ind1 : integer;

begin
  Result := True;
  // Initialisation de la table Site : un Site par dépot defini dans la base
  stSQL := 'Delete from QSITE';
  ExecuteSQL(stSQL);
  TobDepot := TOB.Create('DEPOTS', nil, -1);
  stSQL := 'Select GDE_DEPOT from DEPOTS';
  TobDepot.LoadDetailDBFromSQL('DEPOTS', stSQL);
  for ind1 := 0 to TobDepot.Detail.Count - 1 do
  begin
      Depot := TobDepot.Detail[ind1].GetValue('GDE_DEPOT');
      stSQL := 'Insert into QSITE (QSI_CTX, QSI_SITE, QSI_SITELIB, QSI_DEPOT)' +
               ' values ("0", "' + Depot + '", "Atelier Assemblage", "' + Depot + '")';
      Result := (ExecuteSQL(stSQL) = 1);
      if not Result then Break;
  end;
  TobDepot.Free;
end;
//------------------------------------------------------------------------//
function  InitNatureTravail : boolean;
var stSQL : string;

begin
//  Result := True;
  // Initialisation de la table NATURETRAVAIL : une nature de travail FAB obligatoire
  stSQL := 'Delete from WNATURETRAVAIL';
  ExecuteSQL(stSQL);
  stSQL := 'Insert into WNATURETRAVAIL (WNA_NATURETRAVAIL, WNA_LIBELLE, WNA_NATUREPIECEA, WNA_NATUREPIECEB, ' +
           'WNA_WITHMOD, WNA_WITHVAL, WNA_WITHPER, WNA_ACTIF, WNA_ATTWOL, WNA_PHYWOL, WNA_RESWOB, WNA_PHYWOB) ' +
           'values ("FAB", "Fabrication", "WAP", "WBP", "X", "X", "-", "X", "APR", "EPR", "RPR", "SPR")';
  Result := (ExecuteSQL(stSQL) = 1);
end;
//------------------------------------------------------------------------//
function  InitUnites : boolean;
var stSQL : string;

begin
//  Result := True;
  // Initialisation des champs GA_UNITEPROD, GA_UNITECONSO pour tous les articles existants
  stSQL := 'Update ARTICLE set GA_UNITEPROD="UNI", GA_UNITECONSO="UNI", GA_MODECONSO="LAN"';
  Result := (ExecuteSQL(stSQL) > 0);
  stSQL := 'Update ARTICLE set GA_QUALIFUNITESTO="UNI" where GA_QUALIFUNITESTO="" and GA_TENUESTOCK="X"';
  Result := (ExecuteSQL(stSQL) >= 0);
end;
//------------------------------------------------------------------------//
function  InitJetons : boolean;
var stSQL : string;

begin
  Result := True;
  // Initialisation de la table JETONS : Tous partent de 0
//  stSQL := 'Update WJETONS set WJT_JETON=1';
//  Result := (ExecuteSQL(stSQL) > 0);
  stSQL := 'Delete from WJETONS';
  ExecuteSQL(stSQL);
end;
//------------------------------------------------------------------------//
function  InitAutomatisme : boolean;
var stSQL : string;
    LastIdent : integer;
begin
  Result := True;
  // Initialisation de la table AUTOMATISME : un enreg pour 'FAB' avec les etapes 4, 7 et 8 positionnées
  stSQL := 'Delete from WORDREAUTO';
  ExecuteSQL(stSQL);

  LastIdent := wSetJeton('WOA');
  stSQL := 'Insert into WORDREAUTO (WOA_CODEAUTO, WOA_IDENTIFIANT, WOA_NATURETRAVAIL, ' +
           'WOA_BOOLLIBRE4, WOA_BOOLLIBRE7, WOA_BOOLLIBRE8) values (' +
           '"WOL", ' + IntToStr(LastIdent) + ', "FAB", "X", "X", "X")';
  ExecuteSQL(stSQL);
end;
//------------------------------------------------------------------------//
function  InitNomenclature : boolean;
var stSQL, Entete : string;
    TobEntete, TobLignes, TobWNT : TOB;
    ind1 : integer;

begin
  Result := True;
  // Initialisation de la table Site : un Site par dépot defini dans la base
  stSQL := 'Delete from WNOMETET';
  ExecuteSQL(stSQL);
  stSQL := 'Delete from WNOMELIG';
  ExecuteSQL(stSQL);
  stSQL := 'Delete from WARTNAT';
  ExecuteSQL(stSQL);
  TobEntete := TOB.Create('NOMENENT', nil, -1);
  TobLignes := TOB.Create('NOMENLIG', nil, -1);
  TobWNT    := TOB.Create('WNOMETET', nil, -1);
  stSQL := 'Select * from NOMENENT left join ARTICLE on GA_ARTICLE=GNE_ARTICLE where GA_TYPENOMENC="ASS"';
  TobEntete.LoadDetailDBFromSQL('NOMENENT', stSQL);
  for ind1 := 0 to TobEntete.Detail.Count - 1 do
  begin
      Entete := TobEntete.Detail[ind1].GetValue('GNE_NOMENCLATURE');
      stSQL := 'Select * from NOMENLIG where GNL_NOMENCLATURE="' + Entete + '"';
      if CreateEntete(TobEntete.Detail[ind1], TobWNT) then
      begin
        TobLignes.ClearDetail;
        TobLignes.LoadDetailDBFromSQL('NOMENLIG', stSQL);
        if TobLignes.Detail.Count > 0 then
        begin
          if not CreateLigne(TobWNT, TobLignes) then
          begin
            Result := False;
            Exit;
          end;
        end;
      end;
  end;
  TobEntete.Free;
end;
//------------------------------------------------------------------------//
function CreateEntete(TobEntete : TOB; TobWNT : TOB) : boolean;
var QCode : TQuery;
    CodeMajeur, stSQL : string;
    LastIdent : integer;
    TOBA : TOB;

begin
  QCode:=OpenSQL('SELECT WNT_MAJEUR FROM WNOMETET WHERE WNT_NATURETRAVAIL="FAB" AND WNT_ARTICLE="'
                 + TobEntete.GetValue('GNE_ARTICLE') + '" Order By WNT_MAJEUR Desc', True);
  if not QCode.Eof then
     CodeMajeur := QCode.FindField('WNT_MAJEUR').AsString;
  Ferme(QCode);
  if Trim(CodeMajeur) = '' then
      CodeMajeur := 'A'
      else
      CodeMajeur := Chr(Ord(CodeMajeur[1]) + 1);
  TOBA := TOB.Create('', nil, -1);
  QCode:=OpenSQL('SELECT GA_UNITEPROD, GA_QUALIFUNITESTO FROM ARTICLE ' +
                 'WHERE GA_ARTICLE="' + TobEntete.GetValue('GNE_ARTICLE') + '"', True);
  TOBA.SelectDB('', QCode);
  Ferme(QCode);
  LastIdent := wSetJeton('WNT');
  stSQL := 'Insert into WNOMETET (WNT_NATURETRAVAIL, WNT_ARTICLE, WNT_CODEARTICLE, WNT_MAJEUR, WNT_IDENTIFIANT, ' +
           'WNT_CODITI, WNT_ETATREV, WNT_UNITELOT, WNT_QUALIFUNITESTO, WNT_QLOTSAIS, WNT_QLOTSTOC, ' +
           'WNT_COEFLOT) values ("FAB", ' +
           '"' + TobEntete.GetValue('GNE_ARTICLE') + '", ' +
           '"' + Copy(TobEntete.GetValue('GNE_ARTICLE'), 1, 18) + '", ' +
           '"' + CodeMajeur + '", ' +
           IntToStr(LastIdent) + ', "AGC", "MOD", ' +
           '"' + TOBA.GetValue('GA_UNITEPROD') + '", ' +
           '"' + TOBA.GetValue('GA_UNITEPROD') + '", 1, 1, 1)';
  TOBA.Free;
  Result := (ExecuteSQL(stSQL) = 1);
  if not Result then Exit;
  QCode := OpenSQL('Select * from WNOMETET where WNT_IDENTIFIANT=' + IntToStr(LastIdent), True);
  TobWNT.SelectDB('', QCode);
  Ferme(QCode);
// on enregistre l'identifiant GP dans la nomenclature
  TobEntete.PutValue('GNE_IDENTIFIANTWNT', LastIdent);
  stSQL := 'Update NOMENENT set GNE_IDENTIFIANTWNT=' + IntToStr(LastIdent) + ' where GNE_NOMENCLATURE="' +
           TobEntete.GetValue('GNE_NOMENCLATURE') + '"';
  Result := (ExecuteSQL(stSQL) = 1);
  if not Result then Exit;
// on force la tenue en stock de l'article
  stSQL := 'Update ARTICLE set GA_TENUESTOCK="X" where GA_ARTICLE="' + TobEntete.GetValue('GNE_ARTICLE') + '"';
  Result := (ExecuteSQL(stSQL) = 1);
  if not Result then Exit;
// on crée le lien article - nature de travail
  LastIdent := wSetJeton('WAN');
  stSQL := 'Select WAN_NATURETRAVAIL from WARTNAT Where WAN_ARTICLE="' + TobEntete.GetValue('GNE_ARTICLE') + '"';
  if not (ExisteSQL(stSQL)) then
  begin
    stSQL := 'Insert into WARTNAT (WAN_NATURETRAVAIL, WAN_ARTICLE, WAN_CODEARTICLE, WAN_IDENTIFIANT, ' +
             'WAN_CODITI, WAN_CIRCUIT, WAN_MISEENPROD, WAN_WBMEMO) values ("FAB", ' +
             '"' + TobEntete.GetValue('GNE_ARTICLE') + '", ' +
             '"' + Copy(TobEntete.GetValue('GNE_ARTICLE'), 1, 18) + '", ' +
             IntToStr(LastIdent) +
             ', "AGC", "AGC", "DEC", "-")';
    Result := (ExecuteSQL(stSQL) = 1);
    if not Result then Exit;
  end;
end;
//------------------------------------------------------------------------//
function CreateLigne(TobWNT, TobLignes : TOB) : boolean;
var QCode : TQuery;
    CodeArt, stSQL, stSQLDeb : string;
    ind1 : integer;
    TOBA : TOB;

begin
  Result := False;
  stSQLDeb := 'Insert into WNOMELIG (WNL_NATURETRAVAIL, WNL_ARTICLE, WNL_CODEARTICLE, WNL_MAJEUR, WNL_LIENNOME, ' +
              'WNL_CIRCUIT, WNL_IDENTIFIANT, WNL_PHASE, WNL_OPEITI, WNL_OPEITIAPPRO, WNL_COMPOSANT, WNL_CODECOMPOSANT, ' +
              'WNL_TYPECOMPOSANT, WNL_QLIENSAIS, WNL_QLIENSTOC, WNL_QLOTSAIS, WNL_MODECONSO, ' +
              'WNL_COEFLIEN, WNL_COEFPFIXE, WNL_QPFIXESAIS, WNL_QPFIXESTOC, ' +
              'WNL_QPPERSAIS, WNL_QPPERSTOC, WNL_COEFPPER, ' +
              'WNL_QPERIODESAIS, WNL_QPERIODESTOC, WNL_COEFPERIODE, WNL_PERTEPROP, ' +
              'WNL_UNITELIEN, WNL_QUALIFUNITESTO, WNL_UNITEPPER, WNL_UNITEPERIODE, WNL_UNITEPFIXE) values (';
  for ind1 := 0 to TobLignes.Detail.Count - 1 do
  begin
    CodeArt := TobLignes.Detail[ind1].GetValue('GNL_ARTICLE');
    TOBA := TOB.Create('ARTICLE', nil, -1);
    QCode := OpenSQL('SELECT GA_UNITEPROD FROM ARTICLE WHERE GA_ARTICLE="' + CodeArt + '"', True);
    TOBA.SelectDB('', QCode);
    Ferme(QCode);
    stSQL := stSQLDeb + '"FAB", "' + Trim(TobWNT.GetValue('WNT_ARTICLE'));
    stSQL := stSQL + '", "' + Copy(Trim(TobWNT.GetValue('WNT_ARTICLE')), 1, 18);
    stSQL := stSQL + '", "' + Trim(TobWNT.GetValue('WNT_MAJEUR'));
    stSQL := stSQL + '", ' + IntToStr((ind1 + 1) * 1000) + ', "", ' + IntToStr(wSetJeton('WNL'));
    stSQL := stSQL + ', "AGC", "110", "110", ' + '"' + CodeArt + '", "' + Copy(CodeArt, 1, 18) + '", "';
    stSQL := stSQL + '", ' + FloatToStr(TobLignes.Detail[ind1].GetValue('GNL_QTE'));
    stSQL := stSQL + ', ' + FloatToStr(TobLignes.Detail[ind1].GetValue('GNL_QTE'));
    stSQL := stSQL + ', 1, "LAN", 1, 1, 0, 0, 0, 0, 1, 0, 0, 1, 0';
    stSQL := stSQL + ', "' + TOBA.GetValue('GA_UNITEPROD');
    stSQL := stSQL + '", "' + TOBA.GetValue('GA_UNITEPROD');
    stSQL := stSQL + '", "' + TOBA.GetValue('GA_UNITEPROD');
    stSQL := stSQL + '", "' + TOBA.GetValue('GA_UNITEPROD');
    stSQL := stSQL + '", "' + TOBA.GetValue('GA_UNITEPROD') + '")';
    TOBA.Free;
    Result := (ExecuteSQL(stSQL) = 1);
    if not Result then Break;
  end;
end;

end.
