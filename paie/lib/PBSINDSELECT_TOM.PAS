{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/05/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PBSINDSELECT (PBSINDSELECT)
Mots clefs ... : TOM;PBSINDSELECT
*****************************************************************}
{
PT1 12/06/2001 SB FQ 13562 Contrôle saisie présentation
}
unit PBSINDSELECT_TOM;

interface

uses
  StdCtrls,
  Controls,
  Classes,
  sysutils,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fiche,
  Fe_main,
  DbCtrls,
  Utob,  
  {$ELSE}
  eFiche,
  MainEagl,
  Utob,
  UtileAgl,
  {$ENDIF}
  HMsgBox,
  UTOM,
  HTB97,
  HCtrls,
  HEnt1,
  HSysMenu;

type
  TOM_PBSINDSELECT = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
  private
    mode, DerniereCreate : String;
    OnFerme, LectureSeule, CEG, STD, DOS: boolean;
    procedure DupliquerPresentation(Sender: TObject);
    procedure ClickBtIndPres(Sender: TObject);
  end;

implementation

uses pgoutils2,pgoutils;

procedure TOM_PBSINDSELECT.OnNewRecord;
begin
  inherited;
  if mode = 'DUPLICATION' then exit;
  SetField('PID_PREDEFINI', 'STD');
  SetField('PID_PGPERIODICITE', 'A');
  SetField('PID_DOSSIER', '000000');
end;

procedure TOM_PBSINDSELECT.OnDeleteRecord;
var
  st: string;
begin
  inherited;

  If ExisteSql('SELECT * FROM BILANSOCIAL WHERE PBC_BSPRESENTATION="'+GetField ('PID_BSPRESENTATION')+'"') then
  Begin
  LastError := 1;
  PgiBox('Suppression impossible. Cette présentation est calculé dans le bilan social.',Ecran.caption);
  exit;
  End;         

  st := 'DELETE FROM PBSINDDETSEL WHERE PIL_PREDEFINI="'+GetField('PID_PREDEFINI')+'" '+
  'AND PIL_NODOSSIER="' + PgRendNoDossier() + '" '+
  'AND PIL_BSPRESENTATION=' +GetField('PID_BSPRESENTATION');
  ExecuteSQL(st);
end;

procedure TOM_PBSINDSELECT.OnUpdateRecord;
begin
  inherited;
  OnFerme := False;
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PID_BSPRESENTATION')
  else
    if (DerniereCreate = GetField('PID_BSPRESENTATION')) then OnFerme := True; // le bug arrive on se casse !!!
  if mode = 'DUPLICATION' then exit;
  { DEB PT1 }
  if GetField('PID_BSPRESENTATION')='' then
    Begin
    LastError := 1;
    PgiBox('Vous devez renseigner : la présentation.',TFFiche(Ecran).Caption);
    End;
  { FIN PT1 }
end;

procedure TOM_PBSINDSELECT.OnAfterUpdateRecord;
begin
  inherited;
if OnFerme then Ecran.Close;
end;

procedure TOM_PBSINDSELECT.OnLoadRecord;
begin
  inherited;
  if (not (DS.state in [dsInsert])) then
  begin
    DerniereCreate := '';
  end;

  if mode = 'DUPLICATION' then exit;

  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;
  if (Getfield('PID_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran)  , (CEG = False));
    SetControlEnabled('BDelete'      , CEG);
  end
  else
    if (Getfield('PID_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete'  , STD);
  end
  else
    if (Getfield('PID_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('BDelete', DOS);
  end;

  SetControlEnabled('BInsert', True);
  SetControlEnabled('PID_PREDEFINI', False);
  SetControlEnabled('PID_BSPRESENTATION', False);
  SetControlEnabled('BDUPLIQUER', True);

  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('PID_PREDEFINI', True);
    SetControlEnabled('PID_BSPRESENTATION', True);
    SetControlEnabled('BDUPLIQUER', False);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
  end;

end;

procedure TOM_PBSINDSELECT.OnChangeField(F: TField);
var
nb : integer;
St, Mes : string;
Modif : Boolean;
begin
  inherited;
  if mode = 'DUPLICATION' then exit;
  if (GetField('PID_PREDEFINI') = 'STD') or (GetField('PID_PREDEFINI') = 'CEG') then SetField('PID_DOSSIER', '000000')
  else SetField('PID_DOSSIER', PgRendNoDossier());

  if (F.FieldName = 'PID_BSPRESENTATION') AND (GetField('PID_BSPRESENTATION')<>'') then
    Begin
    Modif := False;
    St := GetField('PID_BSPRESENTATION');
    if Length(st)<3 then
       Begin
       nb := 3 - Length(St);
       st := UpperCase(StringOfChar('0',nb) + st);
       Modif := True;
       End;
    if (DS.State = dsinsert) then
    Begin
    If (length(st) = 3)  and (not IsNumeric(St[3])) then
      Begin
      PgiBox('Le code présentation doit terminer par un chiffre.',Ecran.caption);
      St := '';
      Modif := True;
      SetFocusControl('PID_BSPRESENTATION');
      End
    else
      If (length(st) = 3)  and (IsNumeric(St[3])) and (GetField('PID_PREDEFINI') <> '') then
         Begin
         if (TestRubrique(GetField('PID_PREDEFINI'), st, 0) = False) or (St = '000') then
           begin
           Mes := MesTestRubrique('BILAN', GetField('PID_PREDEFINI'), 0);
           PgiBox( mes , Ecran.caption);
           SetFocusControl('PID_BSPRESENTATION');
           St := '';
           Modif := True;
           end;
         End;
    End;
    if Modif  then SetField('PID_BSPRESENTATION',UpperCase(St));
    End;

    if (F.FieldName = 'PID_PREDEFINI') then
      Begin
      If (DS.State = dsinsert) and (GetField('PID_BSPRESENTATION')<>'') then
          SetField('PID_BSPRESENTATION',GetField('PID_BSPRESENTATION'));
      End;


if (ds.state in [dsBrowse]) then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
    end;
    SetControlEnabled('BInsert', True);
    SetControlEnabled('PID_PREDEFINI', False);
    SetControlEnabled('PID_BSPRESENTATION', False);
    SetControlEnabled('BDUPLIQUER', True);
  end;
end;

procedure TOM_PBSINDSELECT.OnArgument(S: string);
Var Btn : TToolBarButton97;
    Edit : THEdit;
begin
  inherited;
  Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
  if btn <> nil then Btn.OnClick := DupliquerPresentation;

  Edit := THEdit(GetControl('BTINDPRES'));
  if Assigned(Edit) then  Edit.OnClick := ClickBtIndPres;
end;

procedure TOM_PBSINDSELECT.DupliquerPresentation(Sender: TObject);
var
  TOB_GestAssoc, T_Present, T: TOB;
  i: integer;
  AncValPres, St, NoDossier : string;
  Valeur: array[1..3] of variant;
  Ok: Boolean;
  Q: TQuery;
begin
  //PID_PREDEFINI ,PID_NODOSSIER,PID_BSPRESENTATION
  TFFiche(Ecran).BValider.Click;
  mode := 'DUPLICATION';
  AncValPres := GetField('PID_BSPRESENTATION');
  AglLanceFiche('PAY', 'CODE', '', '', 'BILAN;' + AncValPres + '; ;3');
  if PGCodeDupliquer <> '' then
  begin
    Valeur[1] := PGCodePredefini;
    if PGCodePredefini = 'DOS' then Valeur[2] := PgRendNoDossier()
    else Valeur[2] := '000000';
    Valeur[3] := PGCodeDupliquer;
    Ok := RechEnrAssocier('PBSINDSELECT',['PID_PREDEFINI','PID_NODOSSIER','PID_BSPRESENTATION'], Valeur);
    if Ok = False then //Test si code existe ou non
    begin
      //PIL_PREDEFINI,PIL_NODOSSIER,PIL_BSPRESENTATION,PIL_INDICATEURBS
      TOB_GestAssoc := TOB.Create('La Cotisation originale', nil, -1);
      st := 'SELECT * FROM PBSINDDETSEL WHERE ##PIL_PREDEFINI## PIL_BSPRESENTATION="' + AncValPres + '"';
      Q := OpenSql(st, TRUE);
      TOB_GestAssoc.LoadDetailDB('PBSINDDETSEL', '', '', Q, FALSE);
      ferme(Q);
      T_Present := TOB.Create('La présentation dupliquée', nil, -1);
      DupliquerPaie(TFFiche(Ecran).TableName, TFFiche(Ecran));
      SetField('PID_BSPRESENTATION', PGCodeDupliquer);
      { Duplication de la gestion associée }
      try
        BeginTrans;
        for i := 0 to TOB_GestAssoc.Detail.Count - 1 do
          if ((TOB_GestAssoc.Detail[i].GetValue('PIL_BSPRESENTATION')) = AncValPres) then
          begin
            T := TOB_GestAssoc.Detail[i];
            if T <> nil then
            begin
              T.PutValue('PIL_BSPRESENTATION', PGCodeDupliquer);
              T.PutValue('PIL_PREDEFINI', PGCodePredefini);
              T.PutValue('PIL_NODOSSIER', PgRendNoDossier());
            end;
          end;
        T_Present.Dupliquer(Tob_GestAssoc, TRUE, TRUE, FALSE);
        T_Present.InsertDB(nil, False);
        TOB_GestAssoc.free;
        T_Present.free;
        CommitTrans;
      except
        Rollback;
        PGIBox('Une erreur est survenue lors de la duplication des éléments associés à la présentation.', Ecran.caption);
      end;
      SetField('PID_PREDEFINI', PGCodePredefini);
      AccesFicheDupliquer(TFFiche(Ecran), PGCodePredefini, NoDossier, LectureSeule);
      SetField('PID_NODOSSIER', NoDossier);
      SetControlEnabled('BInsert', True);
      SetControlEnabled('PID_PREDEFINI', False);
      SetControlEnabled('PID_BSPRESENTATION', False);
      SetControlEnabled('BDUPLIQUER', True);
      TFFiche(Ecran).Bouge(nbPost); //Force enregistrement
      PgiInfo('Duplication effectuée.', Ecran.caption);
    end
    else
      PgiBox('La duplication est impossible, la présentation existe déjà.',Ecran.caption);
  end;
  mode := '';
end;




procedure TOM_PBSINDSELECT.ClickBtIndPres(Sender: TObject);
begin
AglLanceFiche('PAY','MUL_BSINDICPRES','','',GetField ('PID_BSPRESENTATION')+';'+GetField ('PID_PREDEFINI')+';'+GetField ('PID_LIBELLE'));
end;

initialization
  registerclasses([TOM_PBSINDSELECT]);
end.

