{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/05/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PBSINDICATEURS (PBSINDICATEURS)
Mots clefs ... : TOM;PBSINDICATEURS
*****************************************************************}
{
PT1 10/02/2006 SB V650 FQ 12902 Refonte indicateur absence
PT2 02/05/2006 SB V_65 FQ 12480 Nouveaux indicateurs
PT3 07/12/2006 SB V_70 FQ 13639 Modification des tablettes
}
unit PBSINDICATEURS_TOM;

interface

uses
  Controls,
  Classes,
  sysutils,
  {$IFNDEF EAGLCLIENT}
  db,
  Fiche,
  Fe_main,
  DbCtrls,
  {$ELSE}
  Utob,
  eFiche,
  MainEagl,
  UtileAgl,
  {$ENDIF}
  HMsgBox,
  UTOM,
  HTB97,
  HCtrls;


type
  TOM_PBSINDICATEURS = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
    Private
    mode, DerniereCreate : String;
    OnFerme, LectureSeule, CEG, STD, DOS,SaisieValeur : Boolean;
    procedure DupliquerIndicateur ( sender : TObject);
  end;

implementation

uses pgoutils,pgoutils2,p5def;

procedure TOM_PBSINDICATEURS.OnNewRecord;
begin
  inherited;
  if mode = 'DUPLICATION' then exit;
  SetField('PBI_TYPINDICATBS', '998');
  SetField('PBI_PREDEFINI', 'DOS');
  SetFocusControl('PBI_INDICATEURBS');
end;

procedure TOM_PBSINDICATEURS.OnDeleteRecord;
begin
  inherited;

  If ExisteSql('SELECT * FROM PBSINDDETSEL WHERE ##PIL_PREDEFINI## PIL_INDICATEURBS="'+GetField ('PBI_INDICATEURBS')+'"') then
    Begin
    LastError := 1;
    PgiBox('Suppression impossible. Cet indicateur est rattaché à une présentation.',Ecran.caption);
    exit;
    End;

  If ExisteSql('SELECT * FROM BILANSOCIAL WHERE PBC_INDICATEURBS="'+GetField ('PBI_INDICATEURBS')+'"') then
    Begin
    LastError := 1;
    PgiBox('Suppression impossible. Cet indicateur est calculé dans le bilan social.',Ecran.caption);
    exit;
    End; 

end;

procedure TOM_PBSINDICATEURS.OnUpdateRecord;
begin
  inherited;
  OnFerme := False;
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PBI_INDICATEURBS')
  else
    if (DerniereCreate = GetField('PBI_INDICATEURBS')) then OnFerme := True; // le bug arrive on se casse !!!

  if mode = 'DUPLICATION' then exit;
  if (GetField ('PBI_LESVALEURS')= '') and (SaisieValeur) then
  begin
    LastError := 1;
    PGIBox('Vous devez saisir des valeurs',Ecran.Caption);
    SetFocusControl('PBI_LESVALEURS');
    exit;
  end;
  if GetField ('PBI_TYPINDICATBS')= '' then
  begin
    LastError := 1;
    PGIBox('Vous devez saisir un type d''indicateur',Ecran.Caption);
    SetFocusControl('PBI_TYPINDICATBS');
    exit;
  end;
end;

procedure TOM_PBSINDICATEURS.OnAfterUpdateRecord;
begin
  inherited;
  if OnFerme then Ecran.Close;
end;

procedure TOM_PBSINDICATEURS.OnLoadRecord;
begin
  inherited;
  if not (DS.State in [dsInsert]) then DerniereCreate := '';
  if mode = 'DUPLICATION' then exit;

  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;
  if (Getfield('PBI_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran)  , (CEG = False));
    SetControlEnabled('BDelete'      , CEG);
  end
  else
    if (Getfield('PBI_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end
  else
    if (Getfield('PBI_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('BDelete', DOS);
  end;

  SetControlEnabled('BInsert', True);
  SetControlEnabled('PBI_PREDEFINI', False);
  SetControlEnabled('PBI_INDICATEURBS', False);
  SetControlEnabled('BDUPLIQUER', True);

  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('PBI_PREDEFINI', True);
    SetControlEnabled('PBI_INDICATEURBS', True);
    SetControlEnabled('BDUPLIQUER', False);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
  end;
end;

procedure TOM_PBSINDICATEURS.OnChangeField(F: TField);
var
  st, LaTablette: string;
  Nb: Integer;
begin
  inherited;

  if F.FieldName = 'PBI_TYPINDICATBS' then
  begin
    St := GetControlText(F.FieldName);
    if St = '' then exit;
    SetControlEnabled('PBI_FOURCHETTEMINI', FALSE);
    SetControlEnabled('PBI_FOURCHETTEMAXI', FALSE);
    SetControlEnabled('PBI_ARRONDI'       , FALSE);
//    SetControlProperty('TBRUPTURE','TabVisible',True);
    Nb := StrToInt(St);
    SaisieValeur := True;
    if nb < 24 then SetControlEnabled('PBI_ARRONDI', TRUE);
    case Nb of
      1..3: LaTablette := 'PGCUMULPAIEDUP';  { PT2 } { PT3 }
      5..8: LaTablette := 'PGREMUNERATIONDUP'; { PT2 } { PT3 }
      10..12: LaTablette := 'PGCOTISATIONDUP'; { PT3 }
      20..23: LaTablette := 'PGBASECOTISATIONDUP'; { PT3 }
      24 : begin LaTablette := 'PGMOTIFABSENCELIB';        SaisieValeur := False; End;  { PT1 }
      25 : begin LaTablette := 'PGTYPEABSDUP';   { PT3 }   SaisieValeur := False; End;  { PT1 }
      26 : begin LaTablette := 'PGMOTIFABSENCELIB';        SaisieValeur := False; End;  { PT1 }
      27 : begin LaTablette := 'PGTYPEABSDUP';   { PT3 }   SaisieValeur := False; End;  { PT1 }
      30..31:
        begin
          SetControlEnabled('PBI_FOURCHETTEMINI', TRUE);
          SetControlEnabled('PBI_FOURCHETTEMAXI', TRUE);
          SaisieValeur := False;
        end;
      32..33:
        begin
          SetControlEnabled('PBI_FOURCHETTEMINI', TRUE);
          SetControlEnabled('PBI_FOURCHETTEMAXI', TRUE);
          SetControlEnabled('PBI_ARRONDI', TRUE);
          SaisieValeur := False;
        end;  
      35..36 : begin LaTablette := 'PGMOTIFABSENCELIB';    SaisieValeur := False; End;  { PT1 }
      37..38 : begin LaTablette := 'PGTYPEABSDUP'; { PT3 } SaisieValeur := False; End;  { PT1 }
      50     : LaTablette := 'PGMOTIFSORTIEDUP';   { PT3 }
      60..65 : LaTablette := 'PGTYPECONTRATDUP';   { PT3 }
      200..206 :
        Begin
        SetControlEnabled('PBI_ARRONDI', TRUE);
        LaTablette := 'PGCUMULPAIEDUP';  { PT3 }
//        SetControlProperty('TBRUPTURE','TabVisible',False);
        End;
      400: LaTablette := 'TZGENERAL';  // Comptes généraux ou racine des comptes généraux
    end;
    if LaTablette <> '' then
    begin
      SetControlEnabled('PBI_LESVALEURS', TRUE);
      SetControlProperty('PBI_LESVALEURS', 'Datatype', LaTablette);
    end
    else
    begin
      SaisieValeur := False;
      SetControlEnabled('PBI_LESVALEURS', FALSE);
      SetControlProperty('PBI_LESVALEURS', 'Datatype', '');
      if (mode='') and (GetField('PBI_FOURCHETTEMINI')<>0) then SetField('PBI_FOURCHETTEMINI', 0);
      if (mode='') and (GetField('PBI_FOURCHETTEMAXI')<>0) then SetField('PBI_FOURCHETTEMAXI', 0);
    end;
  end;

if mode = 'DUPLICATION' then exit;



  if F.FieldName = 'PBI_INDICATEURBS' then
    Begin
    If (length(GetField('PBI_INDICATEURBS')) > 0) AND (DS.State = dsinsert) then
      if (TestRubrique(GetField('PBI_PREDEFINI'), GetField('PBI_INDICATEURBS'), 0) = False)  then
      begin
      st := MesTestRubrique('BILAN', GetField('PBI_PREDEFINI'), 0);
      PgiBox( st , Ecran.caption);
      SetFocusControl('PBI_INDICATEURBS');
      if GetField('PBI_INDICATEURBS') <> '' then SetField('PBI_INDICATEURBS','');
      end;
    End;

   if (F.FieldName = 'PBI_PREDEFINI') then
      Begin
      If (DS.State = dsinsert) and (GetField('PBI_INDICATEURBS')<>'') then
          SetField('PBI_INDICATEURBS',GetField('PBI_INDICATEURBS'));
      End;

if (ds.state in [dsBrowse]) then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
    end;
    SetControlEnabled('BInsert', True);
    SetControlEnabled('PBI_PREDEFINI', False);
    SetControlEnabled('PBI_INDICATEURBS', False);
    SetControlEnabled('BDUPLIQUER', True);
  end;

end;

procedure TOM_PBSINDICATEURS.OnArgument(S: string);
Var Btn : TToolBarButton97;
    Num : integer;
begin
  inherited;
  Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
  if btn <> nil then Btn.OnClick := DupliquerIndicateur;
  for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PBI_BTRAVAILN' + IntToStr(Num)), GetControl('TPBI_BTRAVAILN' + IntToStr(Num)));
  VisibiliteStat(GetControl('PBI_BCODESTAT'), GetControl('TPBI_BCODESTAT'));
  VisibiliteChampLibreSal(IntToStr(Num), GetControl('PBI_BLIBREPCMB' + IntToStr(Num)), GetControl('TPBI_BLIBREPCMB' + IntToStr(Num)))
  end;
end;

procedure TOM_PBSINDICATEURS.OnClose;
begin
  inherited;
end;

procedure TOM_PBSINDICATEURS.OnCancelRecord;
begin
  inherited;
end;

procedure TOM_PBSINDICATEURS.DupliquerIndicateur(sender: TObject);
var
  AncValInd, NoDossier : string;
  Valeur: array[1..3] of variant;
  Ok: Boolean;
begin
//PBI_INDICATEURBS
  TFFiche(Ecran).BValider.Click;
  mode := 'DUPLICATION';
  AncValInd := GetField('PBI_INDICATEURBS');
  AglLanceFiche('PAY', 'CODE', '', '', 'BILAN;' + AncValInd + '; ;0');
  if PGCodeDupliquer <> '' then
  begin
    Valeur[1] := PGCodePredefini;
    if PGCodePredefini = 'DOS' then Valeur[2] := PgRendNoDossier()
    else Valeur[2] := '000000';
    Valeur[3] := PGCodeDupliquer;
    Ok := RechEnrAssocier('PBSINDICATEURS',['PBI_PREDEFINI','PBI_NODOSSIER','PBI_INDICATEURBS'], Valeur);
    if Ok = False then //Test si code existe ou non
    begin  //PBI_PREDEFINI,PBI_NODOSSIER,PBI_INDICATEURBS  
      DupliquerPaie(TFFiche(Ecran).TableName, TFFiche(Ecran));
      SetField('PBI_INDICATEURBS', PGCodeDupliquer);
      { DEB PT17-1 Refonte pour intégration de la duplication des éléments associés }
      { Duplication de la gestion associée }
      SetField('PBI_PREDEFINI', PGCodePredefini);
      AccesFicheDupliquer(TFFiche(Ecran), PGCodePredefini, NoDossier, LectureSeule);
      SetField('PBI_NODOSSIER', NoDossier);
      SetControlEnabled('BInsert', True);
      SetControlEnabled('PBI_PREDEFINI', False);
      SetControlEnabled('PBI_INDICATEURBS', False);
      SetControlEnabled('BDUPLIQUER', True);
      TFFiche(Ecran).Bouge(nbPost); //Force enregistrement
      PgiInfo('Duplication effectuée.', Ecran.caption); //PT18-3
    end
    else
      PgiBox('La duplication est impossible, l''indicateur existe déjà.',Ecran.caption);
  end;
  mode := '';
end;


initialization
  registerclasses([TOM_PBSINDICATEURS]);
end.

