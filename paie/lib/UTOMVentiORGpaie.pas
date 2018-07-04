{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... : 11/09/2001
Description .. : TOM gestion des ventilations des organismes pour les
Suite ........ : écritures comptables
Mots clefs ... : PAIE;GENERCOMPTA
*****************************************************************
PT1 SB 29/11/2001 V563 On force la validation avant la Duplication
PT2 SB 05/12/2001 V563 Fiche de bug n°364 Affectation du prédefini
PT3 SB 13/12/2001 V570 Fiche de bug n° 279
                       Test code existant ne test pas bon numéro de dossier
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT4 PH 12/05/2004 V_50 FQ 11001 Initialisation champ PREDEFINI
PT5 PH 06/07/2004 V_50 FQ 11001 Contrôle des champs saisis
PT5 PH 23/12/2004 V_60 FQ xxxxx Bouton création invisible
PT6 SB 10/06/2005 V_65 CWAS : Chargement tablette compatibilite CWAS
PT7 GGS 06/02/2007 V_80 GGS : Journal d'événements
PT7-2 GGS 26/04/2007 V_80 GGS : revu gestion Trace
PT7-3 GGS 30/05/2007 V_80 GGS : revu gestion Trace Duplication
PT8 FC 06/11/2007 V_80 Correction bug journal des évènement quand création (lib Modification au lieu de Création)
PT9 FC 04/01/2008 V_81 FQ 15081 Un utilisateur qui n'est pas réviseur ne peut plus modifier le prédéfini DOS
}
unit UTOMVENTIORGPAIE;

interface
uses Windows, StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Fe_Main, Fiche,
  {$ELSE}
  MaineAgl, eFiche, UtileAGL,
  {$ENDIF}
  Grids, HCtrls, HEnt1, EntPaie, HMsgBox, UTOM, UTOB, Vierge, P5Def,
  AGLInit, Pgoutils,PgOutils2,
  PAIETOM;          //PT7
type
  TOM_VENTIORGPAIE = class(PGTOM)       //PT7 Tom devient PGTOM
  private
    PGPredef, PGNodossier, PGOrg, mode, DerniereCreate: string;
    CEG, STD, DOS, LectureSeule, OnFerme: Boolean;
    Trace: TStringList;                                            //PT7
    LeStatut:TDataSetState; //PT8
    procedure DupliquerVentilOrg(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
  end;

implementation

procedure TOM_VENTIORGPAIE.OnArgument(Arguments: string);
var
  st: string;
  Btn: TToolBarButton97;
begin
  inherited;
  st := Trim(Arguments);
  PGPredef := '';
  PGNodossier := '';
  PGOrg := ReadTokenSt(st); // Recup organisme
  Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
  if btn <> nil then Btn.OnClick := DupliquerVentilOrg;
  if Pos('CREATION', St) > 1 then SetControlEnabled('BDUPLIQUER', False)
  else SetControlEnabled('BDUPLIQUER', True);
  SetControlVisible('BINSERT', FALSE); // PT5
//PT7-2  Trace := TStringList.Create ;         //PT7
end;

procedure TOM_VENTIORGPAIE.OnChangeField(F: TField);
var
  Pred, Rubrique: string;
begin
  inherited;
  if (F.FieldName = 'PVO_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PVO_PREDEFINI');
    Rubrique := GetField('PVO_TYPORGANISME'); //PT2  Nom de champ erroné : PVO_RUBRIQUE
    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez ventiler d''organisme prédéfini CEGID', 'Accès refusé');
      SetControlProperty('PVO_PREDEFINI', 'Value', 'DOS');
    end
    else
      if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez ventiler d''organisme prédéfini Standard', 'Accès refusé');
      SetControlProperty('PVO_PREDEFINI', 'Value', 'DOS');
    end;
  end;

  if (ds.state in [dsBrowse]) then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
    end;
    SetControlEnabled('BInsert', True);
    SetControlEnabled('PVO_PREDEFINI', False);
    SetControlEnabled('PVO_TYPORGANISME', False); //PT2  Nom de champ erroné : PVO_RUBRIQUE
  end;
end;

procedure TOM_VENTIORGPAIE.OnLoadRecord;
begin
  inherited;
  if not (DS.State in [dsInsert]) then DerniereCreate := '';
  if mode = 'DUPLICATION' then exit;
  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;

  SetControlEnabled('BInsert', True);
  SetControlEnabled('PVO_PREDEFINI', False); //PT2
  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('PVO_PREDEFINI', True);
    //   SetControlEnabled('PVO_TYPORGANISME',True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
  end;
  SetControlEnabled('PVO_TYPORGANISME', False); //PT2 Nom de champ erroné : PVO_RUBRIQUE
  if (Getfield('PVO_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end
  else
    if (Getfield('PVO_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end
  else
    if (Getfield('PVO_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), (DOS = False)); //PT9
    SetControlEnabled('BDelete', DOS);
  end;
  PaieConceptPlanPaie(Ecran);
end;

procedure TOM_VENTIORGPAIE.OnNewRecord;
begin
  inherited;
  if PGOrg <> '' then SetField('PVO_TYPORGANISME', PGOrg);
// PT4 PH 12/05/2004 V_50 FQ 11001 Initialisation champ PREDEFINI
  SetField ('PVO_PREDEFINI', 'DOS' );
end;

procedure TOM_VENTIORGPAIE.OnUpdateRecord;
var Predef : String;
begin
  inherited;
  LeStatut := DS.State; //PT8
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PVO_TYPORGANISME')
  else
    if (DerniereCreate = GetField('PVO_TYPORGANISME')) then OnFerme := True; // le bug arrive on se casse !!!
  if mode = 'DUPLICATION' then exit;
  if (DS.State = dsinsert) then
  begin
    if (GetField('PVO_PREDEFINI') <> 'DOS') then
      SetField('PVO_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PVO_NODOSSIER', PgRendNoDossier());
  end;
  // DEB PT5
  if (GetField ('PVO_RACINE1') = '') AND (GetField ('PVO_RACINE2') = '') then
  begin
    PgiBox ('Vous devez renseigner au moins une racine', Ecran.caption);
    LastError := 1;
  end ;
  if (GetField ('PVO_RACINE1') <> '') then
  begin
    if (GetField ('PVO_PARTSALPAT1') = '') then PgiBox ('Vous devez renseigner le type d''alimentation 1', Ecran.caption);
    if (GetField ('PVO_SENS1') = '') then PgiBox ('Vous devez renseigner le sens 1', Ecran.caption);
    if (GetField ('PVO_PARTSALPAT1') = '') OR (GetField ('PVO_SENS1') = '') then LastError := 2;
  end ;
  if (GetField ('PVO_RACINE2') <> '') then
  begin
    if (GetField ('PVO_PARTSALPAT2') = '') then PgiBox ('Vous devez renseigner le type d''alimentation 2', Ecran.caption);
    if (GetField ('PVO_SENS2') = '') then PgiBox ('Vous devez renseigner le sens 2', Ecran.caption);
    if (GetField ('PVO_PARTSALPAT2') = '') OR (GetField ('PVO_SENS2') = '') then LastError := 2;
  end ;
  // FIN PT5
  PreDef := GetField('PVO_PREDEFINI');
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


procedure TOM_VENTIORGPAIE.DupliquerVentilOrg(Sender: TObject);
var
  Code: THEdit;
  Organisme, NoDossier: string;
  Champ: array[1..3] of Hstring;
  Valeur: array[1..3] of variant;
  st: string;
begin
  TFFiche(Ecran).BValider.Click; //PT1
  mode := 'DUPLICATION';
  Organisme := GetField('PVO_TYPORGANISME');
  AglLanceFiche('PAY', 'CODE', '', '', 'VENT;' + Organisme + '; ;3');
  if PGCodeDupliquer <> '' then
  begin
    Champ[1] := 'PVO_PREDEFINI';
    Valeur[1] := PGCodePredefini;
    Champ[2] := 'PVO_NODOSSIER';
    // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    Valeur[2] := PgRendNoDossier();
    if PGCodePredefini = 'DOS' then Valeur[2] := PgRendNoDossier() //PT3
    else Valeur[2] := '000000'; //PT3
    Champ[3] := 'PVO_TYPORGANISME';
    Valeur[3] := PGCodeDupliquer;
    if RechEnrAssocier('VENTIORGPAIE', Champ, Valeur) = False then //Test si code existe ou non
    begin
      Code := THEdit(GetControl('PVO_TYPORGANISME'));
      if (code <> nil) then DupliquerPaie(TFFiche(Ecran).TableName, Ecran);
      SetField('PVO_TYPORGANISME', PGCodeDupliquer);
      SetField('PVO_PREDEFINI', PGCodePredefini);
      AccesFicheDupliquer(TFFiche(Ecran), PGCodePredefini, NoDossier, LectureSeule);
      SetField('PVO_NODOSSIER', NoDossier);
      SetControlEnabled('PVO_TYPORGANISME', False);
      SetControlEnabled('PVO_PREDEFINI', False);

//PT7-3
      if Assigned(Trace) then FreeAndNil(Trace);
      Trace := TStringList.Create;
      st := 'Duplication de la ventilation organisme '+Organisme;
      Trace.add (st);
      st := 'Création de la ventilation organisme '+ GetField('PVO_TYPORGANISME');
      Trace.add (st);
      EnDupl := 'OUI';
      IsDifferent(dernierecreate, PrefixeToTable(TFFiche(Ecran).TableName), TFFiche(Ecran).CodeName, TFFiche(Ecran).LibelleName, Trace, TFFiche(Ecran)); //PT44
      FreeAndNil(Trace);
      EnDupl := 'NON';
//Fin PT7-3
      TFFiche(Ecran).Bouge(nbPost);
      {$IFNDEF EAGLCLIENT}
      ChargementTablette(TFFiche(Ecran).TableName, '');
      {$ELSE}
      ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), ''); //PT6
      {$ENDIF}
    end
    else
      HShowMessage('5;Ventilation des organismes :;La duplication est impossible, la ventilation existe déjà.;W;O;O;O;;;', '', '');
  end;
  mode := '';
end;

procedure TOM_VENTIORGPAIE.OnAfterUpdateRecord;
var
  even: boolean;          //PT7
begin
  inherited;
      Trace := TStringList.Create ;         //PT7-2
      even := IsDifferent(dernierecreate,PrefixeToTable(TFFiche(Ecran).TableName),'PVO_TYPORGANISME',TFFiche(Ecran).LibelleName,Trace,TFFiche(Ecran),LeStatut);  //PT7 //PT8
      FreeAndNil (Trace);  //PT7-2   Trace.free;
  if OnFerme then Ecran.Close;
end;

initialization
  registerclasses([TOM_VENTIORGPAIE]);
end.

