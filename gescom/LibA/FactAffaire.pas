unit FactAffaire;

interface
uses SysUtils, Hent1, HCtrls, AGLInit,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$ENDIF}
  UTOB, ParamSoc, UTilFonctionCalcul,
  entGC, FactComm, FactUtil, UtilPGI, factcalc,
  SAISUTIL {RDEVISE}, factTob, 

  //MajTable, StdCtrls, ComCtrls, HQry,
  //DBCtrls, GalEnv, PGIEnv, MajHalleyUtil, Forms, HStatus, MajStruc,
  UtilRessource,Dicobtp,uEntCommun,UtilTOBPiece;

// Lien d'une ligne pièce avec une tâche
procedure AFLignePieceToTache(pStCodeAffaire, pStCodeTiers, pStLignePieceAff: string; pAction: TActionFiche);

// AFORMULEVAR - Formules à éléments variables
procedure LoadLesAFFormuleVar(pStFormule: string; TobOrigine: TOB; var TOBFormuleVar, TOBFormuleVarQte: TOB);
procedure DupplicLesAFFormuleVar(TobOrigine, TOBFormuleVarQte, TOBFormuleVarQte_O: TOB);
procedure DupplicLesAFRevision(TobOrigine, TOBRevision, TOBRevision_O: TOB);
function EvaluationAFFormuleVar(pStOrigine, pStFormule, pStAction: string; TobFormuleVar, TobFormuleVarQte, TOBL: TOB): double;
procedure ValideLesAFFormuleVar(TobOrigine, TobOrigine_O, TOBFormuleVarQte: TOB; pNouvelle: boolean = False);
procedure LoadDescriptifAFFormuleVar(TOBFormuleVar: TOB);

function DetruitAffaireComplement(TOBVarQte, TOBRev: TOB): boolean;
function SetDefinitiveNumberAffaire(TOBPiece, TobRevision, TobVariable: TOB; NumDef: integer): boolean;
procedure InverseQteLigneArticle(NewNature: string; TOBPiece_ori, TOBPiece,TOBBases,TOBEches,TobPorcs: TOB);
procedure RepriseQuantiteVar(NewNature: string; TOBPiece_ori, TOBPiece,TobVariable: TOB);

procedure ValideLesVariables(TOBPiece, TOBPiece_ori, TOBVariable: TOB);
procedure ValideLesRevisions(TOBPiece, TOBPiece_ori, TOBRevision: TOB);
procedure G_ChargeLesRevisions(TOBPiece, TOBRevision: TOB); //Affaire-ONYX
procedure G_ChargeLesVariables(TOBPiece, TOBVariable: TOB); //Affaire-ONYX
procedure RenumeroteVariableRev(TobGenere, TobVariable, TobRevision: TOB);
procedure ChargeExceptionDocAffaire(NaturePiece: string; TobAffaire: Tob);

implementation

{**********************************************************************
Auteur  ...... : AB
Description .. : Lien d'une ligne piéce avec une tache
Mots clefs ... : Code Affaire,Code tiers,identifiant ligne
Mots clefs ... : TobOrigine : TobPiece ou TobActivité
Mots clefs ... : TobFormuleVar : description de la formule
Mots clefs ... : TobFormuleVarQte :enregistrement des valeurs saisies
************************************************************************}

procedure AFLignePieceToTache(pStCodeAffaire, pStCodeTiers, pStLignePieceAff: string; pAction: TActionFiche);
var vStArg, vStCle: string;
begin
  vStCle := findEtreplace(pStLignePieceAff, ';', '|', true);
  vStArg := 'ATA_AFFAIRE:' + pStCodeAffaire + ';ATA_TIERS:' + pStCodeTiers + ';ATA_IDENTLIGNE:' + vStCle + ';';
  vStArg := vStArg + ActionToString(pAction);
  AGLLanceFiche('AFF', 'AFTACHES', '', '', vStArg);
end;

{**********************************************************************
Auteur  ...... : AB
Créé le ...... : 10/04/2003
Description .. : Gestion des Formules à éléments variables des quantités
Mots clefs ... : Code formule : FORMULEVAR
Mots clefs ... : TobOrigine : TobPiece ou TobActivité
Mots clefs ... : TobFormuleVar : description de la formule
Mots clefs ... : TobFormuleVarQte :enregistrement des valeurs saisies
************************************************************************}

procedure LoadDescriptifAFFormuleVar(TOBFormuleVar: TOB);
var QQ: TQuery;
  ind: integer;
  StFormuleVar, stSql: string;
  TobLVar, TOBLVarDet, TOBFormuleVarDetail: TOB;
begin
  if (not (ctxAffaire in V_PGI.PGIContexte) and not (ctxGCAFF in V_PGI.PGIContexte))
    or not GetParamSoc('SO_AFVARIABLES') or (TOBFormuleVar = nil) then exit;

  TOBFormuleVar.cleardetail;
  stSql := 'SELECT AVF_FORMULEVAR,AVF_FORMVARLIB,AVF_FORMVARDEF,AVF_FORMULETARIF,AVF_MESURE,';
  stSql := stSql + 'AVF_QTETARIFPRINC,AVF_QTETARIFFINALE,AVF_QTETARIFFOR,AVF_FORCOMPTEUR,AVF_QTEARRONDI';
  stSql := stSql + ' FROM AFORMULEVAR WHERE AVF_FORMVARACTIVE="X"';
  QQ := OpenSQl(stSql, True,-1, '', True);
  if not QQ.EOF then
    TOBFormuleVar.LoadDetailDB('AFORMULEVAR', '', '', QQ, False);
  Ferme(QQ);
  {  for ind := 0 to TOBFormuleVar.detail.count - 1 do
    begin
      StFormuleVar := TOBFormuleVar.detail[ind].GetValue('AVF_FORMULEVAR');
      stSql := 'SELECT AVD_FORMULEVAR, AVD_RANG, AVD_MESURE,AVD_LIBMESURE, AVD_LIBMESURECOURT,';
      stSql := stSql + ' AVD_VALDEF,AVD_LPVIS, AVD_LPENABLE, AVD_ACTVIS, AVD_ACTENABLE,AVD_PLAVIS,AVD_PLAENABLE';
      stSql := stSql + ' FROM AFORMULEVARDET WHERE AVD_FORMULEVAR = "' + StFormuleVar + '"';
      QQ := OpenSQl(stSql, True);
      if not QQ.EOF then
        TOBFormuleVar.detail[ind].LoadDetailDB('AFORMULEVARDET', '', '', QQ, False);
      Ferme(QQ);
    end;   }
  if TOBFormuleVar.detail.count = 0 then exit;
  TOBFormuleVarDetail := TOB.Create('', nil, -1);
  stSql := 'SELECT AVD_FORMULEVAR, AVD_RANG, AVD_MESURE,AVD_LIBMESURE, AVD_LIBMESURECOURT,';
  stSql := stSql + ' AVD_VALDEF,AVD_LPVIS, AVD_LPENABLE, AVD_ACTVIS, AVD_ACTENABLE,AVD_PLAVIS,AVD_PLAENABLE';
  stSql := stSql + ' FROM AFORMULEVAR,AFORMULEVARDET';
  stSql := stSql + ' WHERE AVD_FORMULEVAR = AVF_FORMULEVAR AND AVF_FORMVARACTIVE="X" ';
  stSql := stSql + ' ORDER BY AVD_FORMULEVAR,AVD_RANG';
  QQ := OpenSQl(stSql, True,-1, '', True);
  if not QQ.EOF then
    TOBFormuleVarDetail.LoadDetailDB('AFORMULEVARDET', '', '', QQ, False);
  Ferme(QQ);
  for ind := TOBFormuleVarDetail.Detail.Count - 1 downto 0 do
  begin
    TOBLVarDet := TOBFormuleVarDetail.Detail[ind];
    StFormuleVar := TOBLVarDet.GetValue('AVD_FORMULEVAR');
    TobLVar := TOBFormuleVar.FindFirst(['AVF_FORMULEVAR'], [StFormuleVar], False);
    if TobLVar <> nil then TOBLVarDet.ChangeParent(TobLVar, -1);
  end;
  TOBFormuleVarDetail.Free;
  for ind := 0 to TOBFormuleVar.detail.count - 1 do
    TOBFormuleVar.Detail[ind].detail.Sort('AVD_RANG');
end;

procedure LoadLesAFFormuleVar(pStFormule: string; TobOrigine: TOB; var TOBFormuleVar, TOBFormuleVarQte: TOB);
var QQ: TQuery;
  stSql: string;
  ind : integer;
  TobLQte,TobL :Tob;
begin
  if (not (ctxAffaire in V_PGI.PGIContexte) and not (ctxGCAFF in V_PGI.PGIContexte))
    or not GetParamSoc('SO_AFVARIABLES') or (TobOrigine = nil) then exit;

  if TOBFormuleVar = nil then TOBFormuleVar := TOB.Create('AFORMULEVAR', nil, -1);
  if (TobOrigine.NomTable = 'ACTIVITE') then
  begin
    if pStFormule <> '' then
    begin
      if not TOBFormuleVar.SelectDB('"' + pStFormule + '"', nil) then exit;
      stSql := 'SELECT AVD_FORMULEVAR, AVD_RANG, AVD_MESURE,AVD_LIBMESURE, AVD_LIBMESURECOURT,';
      stSql := stSql + ' AVD_VALDEF,AVD_LPVIS, AVD_LPENABLE, AVD_ACTVIS, AVD_ACTENABLE,AVD_PLAVIS, AVD_PLAENABLE';
      stSql := stSql + ' FROM AFORMULEVARDET  ';
      stSql := stSql + ' WHERE AVD_FORMULEVAR="' + pStFormule + '"';
      QQ := OpenSQL(stSql, TRUE,-1, '', True);
      TOBFormuleVar.LoadDetailDB('AFORMULEVARDET', '', '', QQ, false);
      Ferme(QQ);
    end;
    if TOBFormuleVarQte = nil then
      TOBFormuleVarQte := TOB.Create('AFORMULEVARQTE', nil, -1);
    TobFormuleVarQte.InitValeurs(False);
    if TobOrigine.GetValue('ACT_NUMLIGNEUNIQUE') > 0 then
    begin
      stSql := 'SELECT * FROM AFORMULEVARQTE ';
      stSql := stSql + ' WHERE AVV_ORIGVAR = "ACT" AND AVV_AFFAIRE = "' + TobOrigine.GetValue('ACT_AFFAIRE') + '"';
      stSql := stSql + ' AND AVV_TYPEACTIVITE = "' + TobOrigine.GetValue('ACT_TYPEACTIVITE') + '"';
      stSql := stSql + ' AND AVV_NUMLIGNEUNIQUE = ' + IntToStr(TobOrigine.GetValue('ACT_NUMLIGNEUNIQUE'));
      QQ := OpenSQL(stSql, TRUE,-1, '', True);
      if not QQ.EOF then
        TOBFormuleVarQte.SelectDB('', QQ);
      Ferme(QQ);
    end
    else
    begin
      if TobOrigine.GetValue('ACT_NUMORDRE') > 0 then
      begin
        stSql := 'SELECT * FROM AFORMULEVARQTE ';
        stSql := stSql + ' WHERE AVV_ORIGVAR="' + VH_GC.AFNatAffaire + '"';
        stSql := stSql + ' AND AVV_AFFAIRE="' + TobOrigine.GetValue('ACT_AFFAIRE') + '"';
        stSql := stSql + ' AND AVV_NATUREPIECEG="' + VH_GC.AFNatAffaire + '"';
        stSql := stSql + ' AND AVV_NUMORDRE=' + IntToStr(TobOrigine.GetValue('ACT_NUMORDRE'));
        QQ := OpenSQL(stSql, TRUE,-1, '', True);
        if not QQ.EOF then
          TOBFormuleVarQte.SelectDB('', QQ);
        Ferme(QQ);
      end;
      if TOBFormuleVar.GetValue('AVF_FORCOMPTEUR') = 'X' then
      begin
        stSql := 'SELECT AVV_QTE01 FROM AFORMULEVARQTE '; //recherche de la valeur 1 saisie du compteur
        stSql := stSql + ' WHERE AVV_ORIGVAR = "ACT" AND AVV_AFFAIRE = "' + TobOrigine.GetValue('ACT_AFFAIRE') + '"';
        stSql := stSql + ' AND AVV_TYPEACTIVITE = "' + TobOrigine.GetValue('ACT_TYPEACTIVITE') + '"';
        stSql := stSql + ' AND AVV_FORMULEVAR = "' + pStFormule + '"';
        stSql := stSql + ' AND AVV_DATEACTIVITE <= "' + UsDateTime(TobOrigine.GetValue('ACT_DATEACTIVITE')) + '"';
        stSql := stSql + ' ORDER BY AVV_RANGVAR DESC ';
        QQ := OpenSQL(stSql, TRUE, 1, '', True);
        if not QQ.EOF then
        begin
          QQ.First;
          TOBFormuleVarQte.PutValue('AVV_QTE02', QQ.Fields[0].AsFloat);
        end;
        Ferme(QQ);
      end;
    end;
  end
  else if (TobOrigine.NomTable = 'PIECE') then
  begin
    LoadDescriptifAFFormuleVar(TOBFormuleVar); //charge le descriptif et le détail des formules
    TOBFormuleVarQte.cleardetail;
    stSql := 'SELECT * FROM AFORMULEVARQTE ';
    stSql := stSql + ' WHERE AVV_ORIGVAR="' + TobOrigine.GetValue('GP_NATUREPIECEG') + '"';
    stSql := stSql + ' AND AVV_SOUCHE="' + TobOrigine.GetValue('GP_SOUCHE') + '"';
    stSql := stSql + ' AND AVV_NATUREPIECEG="' + TobOrigine.GetValue('GP_NATUREPIECEG') + '"';
    stSql := stSql + ' AND AVV_NUMERO=' + IntToStr(TobOrigine.GetValue('GP_NUMERO'));
    QQ := OpenSQL(stSql, TRUE,-1, '', True);
    if not QQ.EOF then
      TOBFormuleVarQte.LoadDetailDB('AFORMULEVARQTE', '', '', QQ, False);
    Ferme(QQ);
    TOBFormuleVarQte.AddChampSupValeur('AUTORISEQTE', '-', False);

    for ind := 0 to TobFormuleVarQte.detail.count - 1 do
    begin
      TobLQte := TobFormuleVarQte.detail[ind];
      if (TobLQte = nil) then continue;
      TobL := TobOrigine.FindFirst(['GL_NUMORDRE'], [TobLQte.GetValue('AVV_NUMORDRE')], False);
      if (TobL <> nil) and (TobLQte.GetValue('AVV_QTECALCUL')<>0) then
      begin
        TobL.PutValue('GL_QTEFACT', TobLQte.GetValue('AVV_QTECALCUL'));
        if  (GetInfoParPiece(TobOrigine.GetValue('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X') then  // gm 10/10/03
          InverseX(TOBL, 'GL_QTEFACT');
      end;
    end;
  end;
end;

{============================ Dupplique les formules variables ==================================}

procedure DupplicLesAFFormuleVar(TobOrigine, TOBFormuleVarQte, TOBFormuleVarQte_O: TOB);
var ind: integer;
  TobLQte, TobL, TobLQte_O: TOB;
begin
  if (not (ctxAffaire in V_PGI.PGIContexte) and not (ctxGCAFF in V_PGI.PGIContexte))
    or not GetParamSoc('SO_AFVARIABLES') then exit;

  for ind := 0 to TobFormuleVarQte.detail.count - 1 do
  begin
    TobLQte := TobFormuleVarQte.detail[ind];
    TobLQte_O := TobFormuleVarQte_O.detail[ind];
    if (TobLQte = nil) or (TobLQte_O = nil) then continue;
    TobL := TobOrigine.FindFirst(['ANCIENNUMORDRE'], [TobLQte_O.GetValue('AVV_NUMORDRE')], False);
    if TobL <> nil then
      TobLQte.PutValue('AVV_NUMORDRE', TobL.GetValue('GL_NUMORDRE'));
  end;
end;
{============================ Dupplique les révisions ==================================}

procedure DupplicLesAFRevision(TobOrigine, TOBRevision, TOBRevision_O: TOB);
var ind: integer;
  TobLRev, TobL, TobLRev_O: TOB;
begin
  //  if (not (ctxAffaire in V_PGI.PGIContexte) and not (ctxGCAFF in V_PGI.PGIContexte))
  //    or not GetParamSoc('SO_AFVREVISION') then exit;
  if not (ctxAffaire in V_PGI.PGIContexte) and not (ctxGCAFF in V_PGI.PGIContexte) then exit;

  for ind := 0 to TOBRevision.detail.count - 1 do
  begin
    TobLRev := TOBRevision.detail[ind];
    TobLRev_O := TOBRevision_O.detail[ind];
    if (TobLRev = nil) or (TobLRev_O = nil) then continue;
    TobL := TobOrigine.FindFirst(['ANCIENNUMORDRE'], [TobLRev_O.GetValue('AFR_NUMORDRE')], False);
    if TobL <> nil then
      TobLRev.PutValue('AFR_NUMORDRE', TobL.GetValue('GL_NUMORDRE'));
  end;
end;

{**********************************************************************
Auteur  ...... : AB
Créé le ...... : 10/04/2003
Description .. : Enregistre la saisie des formules à éléments variables
Mots clefs ... : Origine = LIG :ligne pièce; ACT :Activité; PLA :Planning
Mots clefs ... : TobOrigine : TobPiece ou TobActivité
Mots clefs ... : TobOrigine_O : TobPiece_O ou TobligneActivité en base
Mots clefs ... : TobFormuleVarQte :enregistrement des valeurs saisies
Mots clefs ... : Nouvelle ligne pour l'activité à insérer
************************************************************************}

procedure ValideLesAFFormuleVar(TobOrigine, TobOrigine_O, TOBFormuleVarQte: TOB; pNouvelle: boolean = False);
var ind: integer;
  stSql: string;
  TobL, TobLQte, TOBFormuleVar: TOB;
begin
  if (not (ctxAffaire in V_PGI.PGIContexte) and not (ctxGCAFF in V_PGI.PGIContexte))
    or (not GetParamSoc('SO_AFVARIABLES')) or (TobOrigine = nil) then exit;

  if (TobOrigine.NomTable = 'PIECE') then
  begin
    if (TOBFormuleVarQte = nil) then LoadLesAFFormuleVar('', TobOrigine, TOBFormuleVar, TOBFormuleVarQte);
    stSql := 'DELETE FROM AFORMULEVARQTE WHERE AVV_ORIGVAR = "' + TobOrigine.GetValue('GP_NATUREPIECEG') + '"';
    stSql := stSql + ' AND AVV_SOUCHE="' + TobOrigine.GetValue('GP_SOUCHE') + '" AND AVV_NUMERO = ' + IntToStr(TobOrigine.GetValue('GP_NUMERO'));
    ExecuteSQL(stSql);
    if TOBFormuleVarQte.detail.count > 0 then
    begin
      for ind := 0 to TobOrigine.detail.count - 1 do
      begin
        TobL := TobOrigine.detail[ind];
        TobLQte := TobFormuleVarQte.FindFirst(['AVV_NUMORDRE'], [TobL.GetValue('GL_NUMORDRE')], False);
        if TobLQte <> nil then
        begin
          if TobL.GetValue('GL_FORMULEVAR') = '' then TobLQte.free
          else TobLQte.PutValue('AVV_AFFAIRE', TobL.GetValue('GL_AFFAIRE'));
        end;
      end;
      for ind := 0 to TOBFormuleVarQte.detail.count - 1 do
      begin
        TOBFormuleVarQte.detail[ind].PutValue('AVV_ORIGVAR', TobOrigine.GetValue('GP_NATUREPIECEG'));
        TOBFormuleVarQte.detail[ind].PutValue('AVV_RANGVAR', ind + 1);
        TOBFormuleVarQte.detail[ind].PutValue('AVV_NUMERO', TobOrigine.GetValue('GP_NUMERO'));
        TOBFormuleVarQte.detail[ind].PutValue('AVV_SOUCHE', TobOrigine.GetValue('GP_SOUCHE'));
        TOBFormuleVarQte.detail[ind].PutValue('AVV_NATUREPIECEG', TobOrigine.GetValue('GP_NATUREPIECEG'));
      end;
      TOBFormuleVarQte.InsertDB(nil);
    end;
  end
  else if (TobOrigine.NomTable = 'ACTIVITE') then
  begin
    if (TobOrigine_O <> nil) and
      (pNouvelle or ((TobOrigine_O.GetValue('ACT_FORMULEVAR') <> '') and (TobOrigine.GetValue('ACT_FORMULEVAR') = ''))) then
    begin
      stSql := 'DELETE FROM AFORMULEVARQTE WHERE AVV_ORIGVAR="ACT" AND AVV_AFFAIRE="' + TobOrigine_O.GetValue('ACT_AFFAIRE') + '"';
      stSql := stSql + ' AND AVV_RANGVAR=' + IntToStr(TobOrigine_O.GetValue('ACT_NUMLIGNEUNIQUE'));
      stSql := stSql + ' AND AVV_NUMERO=0';
      ExecuteSQL(stSql);
      if TobOrigine.GetValue('ACT_FORMULEVAR') = '' then exit;
    end;
    if (TOBFormuleVarQte = nil) then exit;
    TOBFormuleVarQte.PutValue('AVV_DATEACTIVITE', TobOrigine.GetValue('ACT_DATEACTIVITE'));
    if pNouvelle or (TOBFormuleVarQte.GetValue('AVV_AFFAIRE') = '') then
    begin
      TOBFormuleVarQte.PutValue('AVV_ORIGVAR', 'ACT');
      TOBFormuleVarQte.PutValue('AVV_AFFAIRE', TobOrigine.GetValue('ACT_AFFAIRE'));
      TOBFormuleVarQte.PutValue('AVV_RANGVAR', TobOrigine.GetValue('ACT_NUMLIGNEUNIQUE'));
      TOBFormuleVarQte.PutValue('AVV_NUMLIGNEUNIQUE', TobOrigine.GetValue('ACT_NUMLIGNEUNIQUE'));
      TOBFormuleVarQte.PutValue('AVV_TYPEACTIVITE', TobOrigine.GetValue('ACT_TYPEACTIVITE'));
      TOBFormuleVarQte.PutValue('AVV_NUMERO', 0);
      if not TOBFormuleVarQte.InsertDB(nil) then
        V_PGI.IoError := oeSaisie;
    end else
    begin
      if not TOBFormuleVarQte.UpdateDB then
        V_PGI.IoError := oeUnknown;
    end;
  end;
end;

{***********************************************************************************
Auteur  ...... : AB
Créé le ...... : 10/04/2003
Description .. : Saisie des valeurs des formules à éléments variables
Mots clefs ... : Origine = LIG :ligne pièce; ACT :Activité;GEN:Génération de facture
Mots clefs ... : Code formule : FORMULEVAR ;Code Affaire;
Mots clefs ... : Action : CONSULT-MODIF-SUPPRIME
Mots clefs ... : TobFormuleVar : description de la formule
Mots clefs ... : TobFormuleVarQte :enregistrement des valeurs saisies
Mots clefs ... : TOBL: ligne pièce ou ligne activité
************************************************************************************}

function EvaluationAFFormuleVar(pStOrigine, pStFormule, pStAction: string; TobFormuleVar, TobFormuleVarQte, TOBL: TOB): double;
var RF, RFQte: R_FonctCal;
  StTitre, StFormule, StFormuleTarif, StOrigvar, StFormuleImp, StFormuleImp2: string;
  TobLVar, TobLQte, TobLQteAct: TOB;
  ind: integer;
  NewFormule, changeFormule, bSilence: boolean;

  procedure InitFonctCal(Rang: integer; TobDet: tob);
  var StMesure, StLibMesure: string;
    dqte: double;
  begin
    StMesure := TobDet.getvalue('AVD_MESURE');
    StLibMesure := TobDet.getvalue('AVD_LIBMESURE');
    StFormule := FindEtReplace(StFormule, '[' + StMesure + ']', '[' + StLibMesure + ']', false);
    if StFormuleTarif <> '' then
      StFormuleTarif := FindEtReplace(StFormuleTarif, '[' + StMesure + ']', '[' + StLibMesure + ']', false);
    RF.VarLibelle[Rang] := AdapteVariable(StLibMesure);
    if (pStOrigine = 'LIG') then
    begin
      RF.affichable[Rang] := (TobDet.getvalue('AVD_LPVIS') = 'X');
      RF.modifiable[Rang] := (TobDet.getvalue('AVD_LPENABLE') = 'X');
    end
    else if (pStOrigine = 'ACT') then
    begin
      RF.affichable[Rang] := (TobDet.getvalue('AVD_ACTVIS') = 'X');
      RF.modifiable[Rang] := (TobDet.getvalue('AVD_ACTENABLE') = 'X');
    end
    else if (pStOrigine = 'PLA') then
    begin
      RF.affichable[Rang] := (TobDet.getvalue('AVD_PLAVIS') = 'X');
      RF.modifiable[Rang] := (TobDet.getvalue('AVD_PLAENABLE') = 'X');
    end
    else if (pStOrigine = 'GEN') and (TobDet.getvalue('AVD_ACTENABLE') = 'X') then
    begin
      dqte := TobLQteAct.GetValue('AVV_QTE' + Format('%2.2d', [Rang + 1]));
      TobLQte.PutValue('AVV_QTE' + Format('%2.2d', [Rang + 1]), dqte);
    end;
    if NewFormule or ChangeFormule then
      RF.VarValeur[Rang] := TobDet.getvalue('AVD_VALDEF')
    else
      RF.VarValeur[Rang] := TobLQte.GetValue('AVV_QTE' + Format('%2.2d', [Rang + 1]));

    if (TobLVar.GetValue('AVF_FORCOMPTEUR') = 'X') and (Rang = 1) then
    begin
      RF.VarValeur[Rang] := TobLQte.GetValue('AVV_QTE' + Format('%2.2d', [Rang + 1]));
      RF.modifiable[Rang] := False;
    end;
    TobLQte.PutValue('AVV_UNIT' + Format('%2.2d', [Rang + 1]), StMesure);
  end;

  procedure CalcQteTarif;
  var RFTarif: R_FonctCal;
    ind1, ind2: integer;
  begin
    TobLQte.PutValue('AVV_QUALIFTARIF', TobLVar.GetValue('AVF_MESURE'));
    if TobLVar.GetValue('AVF_QTETARIFPRINC') = 'X' then
    begin
      TobLQte.PutValue('AVV_QTETARIF', RFQte.VarValeur[0]);
    end else
      if TobLVar.GetValue('AVF_QTETARIFFINALE') = 'X' then
    begin
      TobLQte.PutValue('AVV_QTETARIF', RFQte.resultat);
    end else
      if (TobLVar.GetValue('AVF_QTETARIFFOR') = 'X') and (StFormuleTarif <> '') then
    begin
      RFTarif := RecupVariableFormule(GFormuleVersFormatPolonais(StFormuleTarif));
      for ind1 := 0 to High(RFTarif.VarLibelle) - 1 do
      begin
        for ind2 := 0 to High(RFQte.VarLibelle) - 1 do
        begin
          if (RFTarif.VarLibelle[ind1] = '') or (RFQte.VarLibelle[ind2] = '') or (RFTarif.VarLibelle[ind1] <> RFQte.VarLibelle[ind2])
            then continue;
          RFTarif.VarValeur[ind1] := RFQte.VarValeur[ind2];
        end;
      end;
      RFTarif := EvaluationFormule('', RFTarif, True);
      TobLQte.PutValue('AVV_QTETARIF', RFTarif.Resultat);
    end else
    begin
      TobLQte.PutValue('AVV_QTETARIF', 0);
      TobLQte.PutValue('AVV_QUALIFTARIF', '');
    end;
  end;

  function ExtraireSousChaine(Chaine: string): string;
  var nbOuvPar, i_ind1: integer;
  begin
    Delete(Chaine, 1, 1);
    Delete(Chaine, 1, pos('(', Chaine));
    Delete(Chaine, 1, pos('(', Chaine));
    Result := '';
    nbOuvPar := 1;
    i_ind1 := 1;
    while (i_ind1 <= Length(Chaine)) and (nbOuvPar <> 0) do
    begin
      if Chaine[i_ind1] = '(' then
        Inc(nbOuvPar)
      else if Chaine[i_ind1] = ')' then
        Dec(nbOuvPar);
      Inc(i_ind1);
    end;
    if nbOuvPar = 0 then
    begin
      Result := Copy(Chaine, 1, i_ind1 - 2);
    end;
  end;
begin
  Result := -1;
  bSilence := False;
  TobLVar := TOBFormuleVar;
  TobLQte := TobFormuleVarQte;
  StOrigvar := pStOrigine;
  if (pStOrigine = 'LIG') then
  begin
    TobLQte := TobFormuleVarQte.FindFirst(['AVV_NUMORDRE'], [TobL.GetValue('GL_NUMORDRE')], False);
    if (pStAction = 'SUPPRIME') then
    begin
      if TobLQte <> nil then TobLQte.free;
      exit;
    end;
    if (TobLQte = nil) then
    begin
      if (pStAction = 'CONSULT') then
      begin
        Result := -1;
        exit;
      end;
      TobLQte := TOB.Create('AFORMULEVARQTE', TobFormuleVarQte, -1);
      TobLQte.InitValeurs(False);
      TobLQte.putvalue('AVV_NUMORDRE', TobL.GetValue('GL_NUMORDRE'));
      NewFormule := true;
    end;
    TobLVar := TOBFormuleVar.FindFirst(['AVF_FORMULEVAR'], [pStFormule], False);
  end
  else if (pStOrigine = 'GEN') then //génération de factures avec des Lignes affaires liées
  begin
    bSilence := true;
    TobLVar := TOBFormuleVar.FindFirst(['AVF_FORMULEVAR'], [pStFormule], False);
    TobLQte := TobFormuleVarQte.detail[0];
    TobLQteAct := TobFormuleVarQte.detail[1];
  end;

  if TobLVar = nil then exit;
  if not NewFormule and (TobLQte.GetValue('AVV_FORMULEVAR') <> TobLVar.GetValue('AVF_FORMULEVAR')) then
    ChangeFormule := true;

  RF := initRecord;
  StFormule := TobLVar.GetValue('AVF_FORMVARDEF');
  StFormuleTarif := TobLVar.GetValue('AVF_FORMULETARIF');
  if (pStAction = 'CONSULT') then StTitre := 'ACTION=CONSULTATION;' + TobLVar.GetValue('AVF_FORMVARLIB')
  else StTitre := TobLVar.GetValue('AVF_FORMVARLIB');
  for ind := 0 to TobLVar.detail.count - 1 do
    InitFonctCal(ind, TobLVar.detail[ind]);
  RF.Formule := GFormuleVersFormatPolonais(StFormule);
  if TobLVar.GetValue('AVF_QTEARRONDI') = 'X' then
    RF.Formule := 'ARRONDI(' + RF.Formule + ';' + IntToStr(V_PGI.OkDecQ) + ')';
  RF.Affichage := StFormule;
  RFQte := EvaluationFormule(StTitre, RF, bSilence);
  if RFQte.annule then exit;
  Result := RFQte.Resultat;
  if TobLVar.GetValue('AVF_QTEARRONDI') = 'X' then
    StFormuleImp := ExtraireSousChaine(RFQte.Expression)
  else StFormuleImp := RFQte.Expression;
  StFormuleImp2 := StFormuleImp;
  CalcQteTarif;
  TobLQte.PutValue('AVV_FORMULEVAR', TobLVar.GetValue('AVF_FORMULEVAR'));
  TobLQte.PutValue('AVV_QTECALCUL', RFQte.Resultat);
  if (pStOrigine = 'LIG') then
  begin
    TobLQte.PutValue('AVV_NUMORDRE', TobL.GetValue('GL_NUMORDRE'));
    TobL.putvalue('GL_QTETARIF', TobLQte.GetValue('AVV_QTETARIF'));
    TobL.putvalue('GL_QUALIFQTETARIF', TobLVar.GetValue('AVF_MESURE'));
  end else
    if (pStOrigine = 'GEN') then
  begin
    TobL.putvalue('GL_QTETARIF', TobLQte.GetValue('AVV_QTETARIF'));
    TobL.putvalue('GL_QUALIFQTETARIF', TobLVar.GetValue('AVF_MESURE'));
  end else
    if (pStOrigine = 'ACT') then
  begin
    TobLQte.PutValue('AVV_NUMORDRE', TobL.GetValue('ACT_NUMORDRE'));
    TobL.putvalue('ACT_QTETARIF', TobLQte.GetValue('AVV_QTETARIF'));
    TobL.putvalue('ACT_UNITETARIF', TobLVar.GetValue('AVF_MESURE'));
  end;
  for ind := 0 to 9 do
    if RFQte.VarLibelle[ind] = '' then
    begin
      TobLQte.PutValue('AVV_QTE' + Format('%2.2d', [ind + 1]), 0);
      TobLQte.PutValue('AVV_UNIT' + Format('%2.2d', [ind + 1]), '');
    end else
    begin
      TobLQte.PutValue('AVV_QTE' + Format('%2.2d', [ind + 1]), RFQte.VarValeur[ind]);
      StFormuleImp := FindEtReplace(StFormuleImp, RFQte.VarLibelle[ind], TobLVar.Detail[ind].GetValue('AVD_LIBMESURECOURT'), false);
      StFormuleImp2 := FindEtReplace(StFormuleImp2, ' ' + RFQte.VarLibelle[ind], '', false);
    end;
  TobLQte.PutValue('AVV_FORMULEIMP', StFormuleImp);
  TobLQte.PutValue('AVV_FORMULEIMP2', StFormuleImp2);
end;


function DetruitAffaireComplement(TOBVarQte, TOBRev: TOB): boolean;
begin
  Result := True;
  //Affaire-ONYX
  if TOBVarQte <> nil then if TOBVarQte.Detail.Count > 0 then
    begin
      Result := TOBVarQte.DeleteDB;
      if not Result then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
    end;
  if TOBRev <> nil then if TOBRev.Detail.Count > 0 then
    begin
      Result := TOBRev.DeleteDB;
      if not Result then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
    end;
end;

function SetDefinitiveNumberAffaire(TOBPiece, TobRevision, TobVariable: TOB; NumDef: integer): boolean;
var SoucheG: String3;
  i, iNum: integer;
  TOBB: TOB;
begin
  Result := False;
  iNum := 0;
  SoucheG := TOBPiece.GetValue('GP_SOUCHE');
  if SoucheG = '' then Exit;
  if NumDef <= 0 then exit;
  Result := True;
  if NumDef <> TOBPiece.GetValue('GP_NUMERO') then
  begin
    TOBPiece.PutValue('GP_NUMERO', NumDef);
    for i := 0 to TOBPiece.Detail.Count - 1 do
    begin
      if i = 0 then iNum := TOBPiece.Detail[i].GetNumChamp('GL_NUMERO');
      TOBB := TOBPiece.Detail[i];
      TOBB.PutValeur(iNum, NumDef);
    end;

    if TOBRevision <> nil then for i := 0 to TOBRevision.Detail.Count - 1 do
      begin
        if i = 0 then iNum := TOBRevision.Detail[i].GetNumChamp('AFR_NUMERO');
        TOBB := TOBRevision.Detail[i];
        TOBB.PutValeur(iNum, NumDef);
      end;
    if TOBVariable <> nil then for i := 0 to TOBVariable.Detail.Count - 1 do
      begin
        if i = 0 then iNum := TOBVariable.Detail[i].GetNumChamp('AVV_NUMERO');
        TOBB := TOBVariable.Detail[i];
        TOBB.PutValeur(iNum, NumDef);
      end;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Merieux
Créé le ...... : 08/07/2003
Modifié le ... :   /  /
Description .. : Pour passer d'une facture à un avoir on inverse le sens des
Suite ........ : qté
Mots clefs ... : GIGA;FACTURATION
*****************************************************************}

procedure InverseQteLigneArticle(NewNature: string; TOBPiece_ori, TOBPiece,TOBBases,TOBEches,TobPorcs: TOB);
//var i: integer;
//  TOBL: TOB;
begin
  if (NewNature = 'AVC') and (TobPiece_ori.GetValue('GP_NATUREPIECEG') = 'FAC') then
  begin
    InverseLesPieces(TOBPiece, 'PIECE');
    InverseLesPieces(TOBBases, 'PIEDBASE');
    InverseLesPieces(TOBEches, 'PIEDECHE');
    InverseLesPieces(TOBPorcs, 'PIEDPORT');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Merieux
Créé le ...... : 12/08/2003
Modifié le ... :   /  /    
Description .. : Fonction qui permet d'alimenter la qté 
Suite ........ : facturé(GL_QTEFACT) avec la Qté non arrondi qu'on a 
Suite ........ : stocké dans les variables
Mots clefs ... : GIGA;VARIABLES
*****************************************************************}
procedure RepriseQuantiteVar(NewNature: string; TOBPiece_ori, TOBPiece,TobVariable : TOB);
var i: integer;
    TOBL,TobLqte : TOB;
begin
  if ((NewNature = 'AVC') and (TobPiece_ori.GetValue('GP_NATUREPIECEG') = 'FAC')) or
      ((NewNature = 'FAC') and (TobPiece_ori.GetValue('GP_NATUREPIECEG') = 'FPR')) then
  begin
    if (Tobvariable.detail.count <> 0) then
    begin
      for i := 0 to TOBPiece.Detail.Count - 1 do
      begin
        TOBL := TOBPiece.Detail[i];
        if (TOBL.GetValue('GL_TYPELIGNE') = 'ART') then
        begin
          TobLQte := TobVariable.FindFirst(['AVV_NUMORDRE'], [TobL.GetValue('GL_NUMORDRE')], False);
          if TobLQte <> nil then
            TobL.PutValue('GL_QTEFACT',TobLqte.GetValue('AVV_QTECALCUL'));
         end;
      end; // for
    end;
  end;
end;

procedure ValideLesRevisions(TOBPiece, TOBPiece_ori, TOBRevision: TOB);
var i, Numero, Indice: integer;
  Nature, Souche: string;
  TOBP: TOB;
begin
  if TOBRevision = nil then Exit;
  if TOBRevision.Detail.Count <= 0 then Exit;
  Nature := TOBPiece.GetValue('GP_NATUREPIECEG');
  Souche := TOBPiece.GetValue('GP_SOUCHE');
  Numero := TOBPiece.GetValue('GP_NUMERO');
  Indice := TOBPiece.GetValue('GP_INDICEG');
  for i := 0 to TOBRevision.Detail.Count - 1 do
  begin
    TOBP := TOBRevision.Detail[i];
    TOBP.PutValue('AFR_NATUREPIECEG', Nature);
    TOBP.PutValue('AFR_SOUCHE', Souche);
    TOBP.PutValue('AFR_NUMERO', Numero);
    TOBP.PutValue('AFR_INDICEG', Indice);
  end;
  if not TOBRevision.InsertDB(nil) then V_PGI.IoError := oeUnknown;
end;

procedure ValideLesVariables(TOBPiece, TOBPiece_ori, TOBVariable: TOB);
var i, Numero: integer;
  Nature, Souche: string;
  TOBP: TOB;
begin
  if TOBVariable = nil then Exit;
  if TOBVariable.Detail.Count <= 0 then Exit;
  Nature := TOBPiece.GetValue('GP_NATUREPIECEG');
  Souche := TOBPiece.GetValue('GP_SOUCHE');
  Numero := TOBPiece.GetValue('GP_NUMERO');
  for i := 0 to TOBVariable.Detail.Count - 1 do
  begin
    TOBP := TOBVariable.Detail[i];
    TOBP.PutValue('AVV_NATUREPIECEG', Nature);
    TOBP.PutValue('AVV_SOUCHE', Souche);
    TOBP.PutValue('AVV_NUMERO', Numero);
    TOBP.PutValue('AVV_ORIGVAR', Nature)
  end;
  if not TOBVariable.InsertDB(nil) then V_PGI.IoError := oeUnknown;
end;

procedure G_ChargeLesRevisions(TOBPiece, TOBRevision: TOB); //Affaire-ONYX
var Q: TQuery;
  CD: R_CleDoc;
begin
  CD := TOB2CleDoc(TOBPiece);
  Q := OpenSQL('SELECT * FROM AFREVISION WHERE ' + WherePiece(CD, ttdRevision, False), True,-1, '', True);
  if not Q.EOF then TOBRevision.LoadDetailDB('AFREVISION', '', '', Q, True, True);
  Ferme(Q);
end;

procedure G_ChargeLesVariables(TOBPiece, TOBVariable: TOB); //Affaire-ONYX
var Q: TQuery;
  CD: R_CleDoc;
  ind :integer;
  TobL,TobLQte :tob;
begin
  CD := TOB2CleDoc(TOBPiece);
  Q := OpenSQL('SELECT * FROM AFORMULEVARQTE WHERE ' + WherePiece(CD, ttdVariable, False), True,-1, '', True);
  if not Q.EOF then TOBVariable.LoadDetailDB('AFORMULEVARQTE', '', '', Q, True, True);
  Ferme(Q);
  for ind := 0 to TOBVariable.detail.count - 1 do
  begin
    TobLQte := TOBVariable.detail[ind];
    if (TobLQte = nil) then continue;
    TobL := TOBPiece.FindFirst(['GL_NUMORDRE'], [TobLQte.GetValue('AVV_NUMORDRE')], False);
    if (TobL <> nil) and (TobLQte.GetValue('AVV_QTECALCUL')<>0) then
    begin
      TobL.PutValue('GL_QTEFACT', TobLQte.GetValue('AVV_QTECALCUL'));
    end;
  end;
end;

procedure RenumeroteVariableRev(TobGenere, TobVariable, TobRevision: TOB);
var TobVariable_O, TobRevision_O: TOB;
begin
  // ne pas faire ce traitement pour les regroupements d'achat
  // mais seulment pour les FAC (transformation de FPR en FAC)
  if (Tobgenere.GetValue('GP_NATUREPIECEG') <> 'FAC') then exit;

  TOBVariable_O := TOB.Create('', nil, -1);
  TOBRevision_O := TOB.Create('', nil, -1);
  TOBVariable_O.Dupliquer(TOBVariable, True, True);
  TOBRevision_O.Dupliquer(TOBRevision, True, True);
  DupplicLesAFFormuleVar(TOBGenere, TOBVariable, TOBVariable_O);
  DupplicLesAFRevision(TOBGenere, TOBRevision, TOBRevision_O);
  TOBVariable_O.free;
  TOBRevision_O.free;
end;

procedure ChargeExceptionDocAffaire(NaturePiece: string; TobAffaire: Tob);
var Q: TQuery;
  StSql: string;
begin
  if GetparamSoc('SO_AFPARPIECEAFF') and (TobAffaire.GetValue('AFF_AFFAIRE') <> '') then
  begin
    StSql := 'SELECT * FROM AFFAIREPIECE WHERE API_AFFAIRE="' + TobAffaire.GetValue('AFF_AFFAIRE') + '"';
    if trim(NaturePiece) <> '' then
      StSql := StSql + ' AND API_NATUREPIECEG="' + NaturePiece + '"';
    Q := OpenSQL(StSql, true,-1, '', True);
    if not Q.EOF then
      TobAffaire.LoadDetailDB('AFFAIREPIECE', '', '', Q, False);
    Ferme(Q);
  end;
end;

end.
