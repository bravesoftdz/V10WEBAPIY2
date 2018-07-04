{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... : 11/09/2001
Description .. : TOM gestion des ventilations des rémunéations pour les
Suite ........ : écritures comptables
Mots clefs ... : PAIE;GENERCOMPTA
*****************************************************************
PT1 SB 29/11/2001 V563 On force la validation avant la Duplication
PT2 SB 30/11/2001 V563 Dysfonctionnement test code existant sur Duplication
PT3 SB 13/12/2001 V570 Fiche de bug n° 279
                       Test code existant ne test pas bon numéro de dossier
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT4 PH 03/06/2004 V_500 FQ 11060 Controle duplication
PT5 PH 06/07/2004 V_50 FQ 11001 Contrôle des champs saisis
PT6 PH 08/11/2004 V_60 Duplication DOS en STD
PT7 SB 10/06/2005 V_65 CWAS : Chargement tablette compatibilite CWAS
PT8 GGS 06/02/2007 V_80 Journal d'événements
PT8-2 GGS 26/04/2007 V_80 Revue gestion Trace
PT8-3 GGS 03/05/2007 V_80 Revue gestion Trace pour duplication
PT9 FC 06/11/2007 V_80 Correction bug journal des évènement quand création (lib Modification au lieu de Création)
}
unit UTOMVENTIREMPAIE;

interface
uses Windows, StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Fe_Main, Fiche,
{$ELSE}
  MaineAgl, eFiche, UtileAGL,
{$ENDIF}
  Grids, HCtrls, HEnt1, EntPaie, HMsgBox, UTOM, UTOB, Vierge, P5Def,
  AGLInit,
  PGOutils,PgOutils2,
  PAIETOM;                        //PT8
type
  TOM_VENTIREMPAIE = class(PGTOM) //PT8 TOM devient PGTOM
  private
    PGPredef, PGNodossier, PGRem, mode, DerniereCreate: string;
    CEG, STD, DOS, LectureSeule: Boolean;
    Loaded, OnFerme: boolean;
    MsgError: integer;
    Trace: TStringList;                                            //PT8
    LeStatut:TDataSetState; //PT9
    procedure DupliquerVentilRem(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
  end;

implementation



procedure TOM_VENTIREMPAIE.OnArgument(Arguments: string);
var
  st: string;
  Btn: TToolBarButton97;
begin
  inherited;
  st := Trim(Arguments);
  PGPredef := '';
  PGNodossier := '';
  PGPredef := ReadTokenSt(st); // Recup Predefini
  PGNodossier := ReadTokenSt(st); // Recup Nodossier
  PGRem := ReadTokenSt(st); // Recup remunération
  Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
  if btn <> nil then Btn.OnClick := DupliquerVentilRem;
  // DEB PT4
  if Pos('CREATION', St) > 1 then
  begin
    if PGPredef = 'CEG' then PGPredef := 'STD';
    SetControlEnabled('BDUPLIQUER', False);
  end
  else SetControlEnabled('BDUPLIQUER', True);
  SetControlVisible('BINSERT', FALSE);
  {
  SetControlEnabled('BDUPLIQUER', False)
  else SetControlEnabled('BDUPLIQUER', True)}
// FIN PT4
//PT8-2  Trace := TStringList.Create ;         //PT8
end;


procedure TOM_VENTIREMPAIE.OnChangeField(F: TField);
var
  Pred, Rubrique: string;
begin
  inherited;
  if mode = 'DUPLICATION' then exit;
  MsgError := 0;
  if Loaded then
  begin
    Loaded := FALSE;
    exit;
  end;
  if (F.FieldName = 'PVS_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PVS_PREDEFINI');
    Rubrique := (GetField('PVS_RUBRIQUE'));
    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas ventiler de rubrique de rémunération prédéfini CEGID', 'Accès refusé');
      SetControlProperty('PVS_PREDEFINI', 'Value', 'DOS');
      SetField('PVS_PREDEFINI', 'DOS');
      MsgError := 1;
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez ventiler de rubrique de rémunération prédéfini Standard', 'Accès refusé');
      SetControlProperty('PVS_PREDEFINI', 'Value', 'DOS');
      MsgError := 1;
      SetField('PVS_PREDEFINI', 'DOS');
    end;
    PGPredef := GetField('PVS_PREDEFINI');
  end;

  if (ds.state in [dsBrowse]) then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
    end;
    SetControlEnabled('BInsert', True);
    SetControlEnabled('PVS_PREDEFINI', False);
    SetControlEnabled('PVS_RUBRIQUE', False);
  end;
    if (Getfield('PVS_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end;
  if (Getfield('PVS_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end;
  if (Getfield('PVS_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), (DOS = False));
    SetControlEnabled('BDelete', DOS);
  end;
  PaieConceptPlanPaie(Ecran);
end;

procedure TOM_VENTIREMPAIE.OnLoadRecord;
begin
  inherited;
  if not (DS.State in [dsInsert]) then DerniereCreate := '';
  if mode = 'DUPLICATION' then exit;
  Loaded := True;
  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;
  SetControlEnabled('BInsert', True);

  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFiche(Ecran), False);
    //   SetControlEnabled('PVS_PREDEFINI',True);
    //   SetControlEnabled('PVS_RUBRIQUE',True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
  end;
  // DEB PT4
  if DS.State in [dsInsert] then
  begin
    if (Getfield('PVS_PREDEFINI') = 'DOS') or (Getfield('PVS_PREDEFINI') = 'STD') then
      SetControlEnabled('PVS_PREDEFINI', true);
  end
  else SetControlEnabled('PVS_PREDEFINI', False);
  SetControlEnabled('PVS_RUBRIQUE', False);
  // FIN PT4
  if (Getfield('PVS_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end;
  if (Getfield('PVS_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end;
  if (Getfield('PVS_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), (DOS = False));
    SetControlEnabled('BDelete', DOS);
  end;
  PaieConceptPlanPaie(Ecran);
end;

procedure TOM_VENTIREMPAIE.OnNewRecord;
begin
  inherited;
  if mode = 'DUPLICATION' then exit;
  Loaded := TRUE;
  // DEB PT4
  if PGPredef <> '' then
  begin
    setControlEnabled('PVS_PREDEFINI', TRUE);
    SetField('PVS_PREDEFINI', PGPredef);
  end;
  // FIN PT4
  if PGNodossier <> '' then SetField('PVS_NODOSSIER', PGNodossier);
  if PGRem <> '' then SetField('PVS_RUBRIQUE', PGRem);

end;

procedure TOM_VENTIREMPAIE.OnUpdateRecord;
var Predef : String;
begin
  inherited;
  OnFerme := False;
  LeStatut := DS.State; //PT9
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PVS_RUBRIQUE')
  else
    if (DerniereCreate = GetField('PVS_RUBRIQUE')) then OnFerme := True; // le bug arrive on se casse !!!
  if mode = 'DUPLICATION' then exit;
  if (DS.State = dsinsert) then
  begin
    if (GetField('PVS_PREDEFINI') <> 'DOS') then
      SetField('PVS_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PVS_NODOSSIER', PgRendNoDossier());
  end;
  // PT4 SetField('PVS_PREDEFINI', GetField('PVS_PREDEFINI'));
  LastError := MsgError;
  // DEB PT5
  if (GetField('PVS_RACINE1') = '') and (GetField('PVS_RACINE2') = '') then
  begin
    PgiBox('Vous devez renseigner au moins une racine', Ecran.caption);
    LastError := 1;
  end;
  if (GetField('PVS_RACINE1') <> '') then
  begin
    if (GetField('PVS_PARTSALPAT1') = '') then PgiBox('Vous devez renseigner le type d''alimentation 1', Ecran.caption);
    if (GetField('PVS_SENS1') = '') then PgiBox('Vous devez renseigner le sens 1', Ecran.caption);
    if (GetField('PVS_PARTSALPAT1') = '') or (GetField('PVS_SENS1') = '') then LastError := 2;
  end;
  if (GetField('PVS_RACINE2') <> '') then
  begin
    if (GetField('PVS_PARTSALPAT2') = '') then PgiBox('Vous devez renseigner le type d''alimentation 2', Ecran.caption);
    if (GetField('PVS_SENS2') = '') then PgiBox('Vous devez renseigner le sens 2', Ecran.caption);
    if (GetField('PVS_PARTSALPAT2') = '') or (GetField('PVS_SENS2') = '') then LastError := 2;
  end;
  // FIN PT5
  PreDef := GetField('PVS_PREDEFINI');
  if (DS.State in [dsInsert]) then
  begin
    if (PreDef = 'CEG') AND (NOT CEG) then
    begin
    PgiBox('Vous n''êtes pas autorisé à créer une ventilation CEGID', Ecran.caption);
    LastError := 1;
    end
    else
    if (PreDef = 'STD') AND (NOT STD) then
    begin
    PgiBox('Vous n''êtes pas autorisé à créer une ventilation standard', Ecran.caption);
    LastError := 1;
    end
    else
    if (PreDef = 'DOS') AND (NOT DOS) then
    begin
    PgiBox('Vous n''êtes pas autorisé à créer une ventilation dossier', Ecran.caption);
    LastError := 1;
    end;
  end;
end;

procedure TOM_VENTIREMPAIE.OnAfterUpdateRecord;
var
  even: boolean;          //PT8
begin
  inherited;
      Trace := TStringList.Create ;         //PT8-2
      even := IsDifferent(dernierecreate,PrefixeToTable(TFFiche(Ecran).TableName),'PVS_RUBRIQUE',TFFiche(Ecran).LibelleName,Trace,TFFiche(Ecran),LeStatut);  //PT8 //PT9
      FreeAndNil (Trace);  //PT8-2  Trace.free;
      if OnFerme then Ecran.Close;
end;

procedure TOM_VENTIREMPAIE.DupliquerVentilRem(Sender: TObject);
var
{$IFNDEF EAGLCLIENT}
  Code: THDBEdit;
{$ELSE}
  Code: THEdit;
{$ENDIF}
  Rubrique : string;
  Champ: array[1..3] of Hstring;
  Valeur: array[1..3] of variant;
  Ok: Boolean;
  St: string;
begin
  // PT4  TFFiche(Ecran).BValider.Click; //PT1 @@@@@@
  mode := 'DUPLICATION';
  Rubrique := GetField('PVS_RUBRIQUE');
  AglLanceFiche('PAY', 'CODE', '', '', 'VENT;' + Rubrique + '; ;4');
  if PGCodeDupliquer <> '' then
  begin
    Champ[1] := 'PVS_PREDEFINI';
    Valeur[1] := PGCodePredefini;
    // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    Champ[2] := 'PVS_NODOSSIER';
    Valeur[2] := PgRendNoDossier();
    if PGCodePredefini = 'DOS' then Valeur[2] := PgRendNoDossier() //PT3
    else Valeur[2] := '000000'; //PT3
    Champ[3] := 'PVS_RUBRIQUE';
    Valeur[3] := PGCodeDupliquer; //PT2 StrToInt
    Ok := RechEnrAssocier('VENTIREMPAIE', Champ, Valeur);
    if Ok = False then //Test si code existe ou non
    begin
{$IFNDEF EAGLCLIENT}
      Code := THDBEdit(GetControl('PVS_RUBRIQUE'));
{$ELSE}
      Code := THEdit(GetControl('PVS_RUBRIQUE'));
{$ENDIF}
      if (code <> nil) then DupliquerPaie(TFFiche(Ecran).TableName, Ecran);
      SetField('PVS_RUBRIQUE', PGCodeDupliquer);
      SetField('PVS_PREDEFINI', PGCodePredefini);
      // PT4     AccesFicheDupliquer(TFFiche(Ecran), PGCodePredefini, NoDossier, LectureSeule);  @@@@
      // PT6 Duplication DOS en STD
      if GetField ('PVS_PREDEFINI') = 'DOS' then SetField('PVS_NODOSSIER', PgRendNoDossier())
      else SetField('PVS_NODOSSIER', '000000');
      // FIN PT6
      SetControlEnabled('PVS_RUBRIQUE', False);
      SetControlEnabled('PVS_PREDEFINI', False);
//PT8-3
      if Assigned(Trace) then FreeAndNil(Trace);
      Trace := TStringList.Create;
      st := 'Duplication de la rubrique '+Rubrique;
      Trace.add (st);
      st := 'Création de la rubrique '+ GetField('PVS_RUBRIQUE');
      Trace.add (st);
      EnDupl := 'OUI';
      IsDifferent(dernierecreate, PrefixeToTable(TFFiche(Ecran).TableName), TFFiche(Ecran).CodeName, TFFiche(Ecran).LibelleName, Trace, TFFiche(Ecran));
      FreeAndNil(Trace);
      EnDupl := 'NON';
//Fin PT8-3
      TFFiche(Ecran).Bouge(nbPost);
      {$IFNDEF EAGLCLIENT}
      ChargementTablette(TFFiche(Ecran).TableName, '');
      {$ELSE}
      ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), '');   //PT7
      {$ENDIF}
      end
    else
      HShowMessage('5;Ventilation des rémunérations :;La duplication est impossible, la ventilation existe déjà.;W;O;O;O;;;', '', '');
  end;
  mode := '';

end;



initialization
  registerclasses([TOM_VENTIREMPAIE]);
end.

