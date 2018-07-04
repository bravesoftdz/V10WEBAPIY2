{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 11/07/2002
Modifié le ... :
Description .. : Source TOM de la TABLE : EXERFORMATION (EXERFORMATION)
Mots clefs ... : TOM;EXERFORMATION
*****************************************************************
21/02/2003 modifs pour Norme développement
PT 1 : 17/05/2004 JL V_50 Nouveau champ : taux de charge prévisionnel
PT 2 : 30/06/2006 JL V_70 13333 ajout du taux de charge dans maj salaire anim
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
PT 3 : 17/10/2007 FL V_80 Emanager / Gestion des prévisionnels Emanager
}

unit UTOMExerFormation;

interface

uses StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DBCtrls,
  {$ELSE}
  eFiche, eFichList, MaineAgl,
  {$ENDIF}
  sysutils, HCtrls, HEnt1, HMsgBox, UTOM, UTob, EntPaie;

type
  TOM_EXERFORMATION = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    MajBudget, Cloture: Boolean;
    TypeSaisie: string;
    procedure GestionOuverturePlan(Sender: TObject);
    procedure GestionOuvertureBudget(Sender: TObject);
  end;

implementation

Uses ParamSoc;

procedure TOM_EXERFORMATION.OnClose;
begin
  inherited;
        AvertirTable('PGEXERFORMATION');
end;

procedure TOM_EXERFORMATION.OnNewRecord;
var
  Q: TQuery;
  aa, jj, mm: Word;
  i: Integer;
begin
  inherited;
  if TypeSaisie = 'EXERCICE' then
  begin
    if not ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION') then SetField('PFE_ENCOURS', 'X')
    else SetField('PFE_ENCOURS', '-');
    SetField('PFE_ACTIF', '-');
    SetField('PFE_CLOTURE', '-');
    SetField('PFE_CLOTUREPREV', '-');
    SetField('PFE_OUVREPREV', '-');
    Q := OpenSQL('SELECT MAX(PFE_MILLESIME) AS MILLESIME FROM EXERFORMATION', True);
    if not Q.Eof then
    begin
      if Q.FindField('MILLESIME').AsString <> '' then
      begin
        SetField('PFE_MILLESIME', IntToStr(Q.FindField('MILLESIME').AsInteger + 1));
        Setfield('PFE_DATEDEBUT', EncodeDate(Q.FindField('MILLESIME').AsInteger + 1, 01, 01));
        Setfield('PFE_DATEFIN', EncodeDate(Q.FindField('MILLESIME').AsInteger + 1, 12, 31));
      end
      else
      begin
        DecodeDate(Date, aa, mm, jj);
        SetField('PFE_MILLESIME', IntToStr(aa));
        Setfield('PFE_DATEDEBUT', EncodeDate(aa, 01, 01));
        Setfield('PFE_DATEFIN', EncodeDate(aa, 12, 31));
      end;
    end;
    Ferme(Q);
  end;
  for i := 1 to 8 do
  begin
    SetField('PFE_COUTFIXEFOR' + IntToStr(i), 0);
  end;
  for i := 1 to 8 do
  begin
    SetField('PFE_OPCACTFIXE' + IntToStr(i), -1);
  end;
  SetField('PFE_MASSESAL', 0);
  SetField('PFE_MASSESALCDD', 0);
  SetField('PFE_TAUXCHARGEC', 1);
  SetField('PFE_TAUXCHARGENC', 1);
  SetField('PFE_TAUXBUDGET', 1);
end;

procedure TOM_EXERFORMATION.OnDeleteRecord;
var
  Mess: string;
begin
  inherited;
  Mess := '';
  if ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_MILLESIME="' + GetField('PFE_MILLESIME') + '"') then
    Mess := '#13#10- Des stages inscrits au budget';
  if ExisteSQL('SELECT PFS_MILLESIME FROM FRAISSALFORM WHERE PFS_MILLESIME="' + GetField('PFE_MILLESIME') + '"') then
    Mess := Mess + '#13#10- Des salaires catégoriels';
  if ExisteSQL('SELECT PFP_MILLESIME FROM FRAISSALPLAF WHERE PFP_MILLESIME="' + GetField('PFE_MILLESIME') + '"') then
    Mess := Mess + '#13#10- Des plafonds pour les frais';
  if ExisteSQL('SELECT PFF_MILLESIME FROM FORFAITFORM WHERE PFF_MILLESIME="' + GetField('PFE_MILLESIME') + '"') then
    Mess := Mess + '#13#10- Des forfaits';
  if Mess <> '' then
  begin
    PGIBox('Impossible de supprimer cet exercice car il existe :' + Mess, Ecran.Caption);
    LastError := 1;
    Exit;
  end;
end;

procedure TOM_EXERFORMATION.OnUpdateRecord;
var
  Q: TQuery;
  Mess: string;
begin
  inherited;
  if TypeSaisie = 'EXERCICE' then
  begin
    Mess := '';
    if ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_MILLESIME<>"' + GetField('PFE_MILLESIME') + '" ' +
      'AND  ((PFE_DATEDEBUT<="' + UsDateTime(GetField('PFE_DATEDEBUT')) + '" AND PFE_DATEFIN>="' + UsDateTime(GetField('PFE_DATEDEBUT')) + '") ' +
      'OR (PFE_DATEDEBUT<="' + UsDateTime(GetField('PFE_DATEFIN')) + '" AND PFE_DATEFIN>="' + UsDateTime(GetField('PFE_DATEFIN')) + '"))') then
      Mess := '- Les dates de début et/ou de fin sont comprises dans un autre exercice.';
    if GetField('PFE_DATEFIN') < GetField('PFE_DATEDEBUT') then
      Mess := Mess + '#13#10 - La date de fin d''exercice ne peut être infèrieur à la date de début d''exercice.';
    if Mess <> '' then
    begin
      LastError := 1;
      PGIBox(Mess, Ecran.Caption);
    end;
  end;
  if Typesaisie = 'SALAIREANIM' then
  begin
    MajBudget := False;
    if not ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE ((PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-") OR (PFE_ACTIF="X" AND PFE_CLOTURE="-")) AND PFE_MILLESIME="' +
      GetField('PFE_MILLESIME') + '"') then
    begin
      PGIBox('Modfication impossible car l''exercice et/ou le budget ne sont pas ouverts', Ecran.Caption);
      LastError := 1;
    end;
    if ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-" AND PFE_MILLESIME="' + GetField('PFE_MILLESIME') + '"') then
    begin
      case PGIAskCancel('Voulez-vous mettre à jour le budget avec le nouveau taux horaire ?', Ecran.Caption) of
        mrYes: MajBudget := True;
        mrCancel:
          begin
            LastError := 1;
            exit;
          end;
      end;
    end;
  end;
end;

procedure TOM_EXERFORMATION.OnAfterUpdateRecord;
var
  Q: TQuery;
  TobBudgete: Tob;
  i: Integer;
  NbHeure, Montant,TxAnim : Double;
  NbAnim: Integer;
begin
  inherited;
  if TypeSaisie = 'EXERCICE' then
  begin
    if GetField('PFE_CLOTUREPREV') = 'X' then Cloture := True
    else Cloture := False;
  end;
  if Typesaisie = 'SALAIREANIM' then
  begin
    if MajBudget = True then
    begin
      Q := OpenSQL('SELECT PFE_TAUXBUDGET FROM EXERFORMATION WHERE PFE_MILLESIME="' + GetField('PFE_MILLESIME') + '"', True);   //PT 2
      If Not Q.eof then TxAnim := Q.FindField('PFE_TAUXBUDGET').AsFloat
      else TxAnim := 1;
      Ferme(Q);
      Q := OpenSQL('SELECT * FROM STAGE WHERE PST_MILLESIME="' + GetField('PFE_MILLESIME') + '"', True);
      TobBudgete := Tob.Create('STAGE', nil, -1);
      TobBudgete.LoadDetailDB('STAGE', '', '', Q, False);
      Ferme(Q);
      for i := 0 to TobBudgete.Detail.Count - 1 do
      begin
        NbHeure := TobBudgete.Detail[i].GetValue('PST_DUREESTAGE');
        NbAnim := TobBudgete.Detail[i].GetValue('PST_NBANIM');
        Montant := GetField('PFE_SALAIREANIM');

        TobBudgete.Detail[i].PutValue('PST_COUTSALAIR', Arrondi(TxAnim * Montant * NbAnim * NbHeure, 2)); //PT 2
        TobBudgete.Detail[i].UpdateDB(False);
      end;
      TobBudgete.Free;
    end;
  end;
end;

procedure TOM_EXERFORMATION.OnLoadRecord;
begin
  inherited;
  if TypeSaisie = 'EXERCICE' then
  begin
    if GetField('PFE_CLOTUREPREV') = 'X' then Cloture := True
    else Cloture := False;
  end;
end;

procedure TOM_EXERFORMATION.OnChangeField(F: TField);
begin
  inherited;
end;

procedure TOM_EXERFORMATION.OnArgument(S: string);
var
  {$IFNDEF EAGLCLIENT}
  Check: TDBCheckBox;
  {$ELSE}
  Check: TCheckBox;
  {$ENDIF}
begin
  inherited;
  TypeSaisie := S;
  {$IFNDEF EAGLCLIENT}
  Check := TDBCheckBox(GetControl('PFE_ACTIF'));
  {$ELSE}
  Check := TCheckBox(GetControl('PFE_ACTIF'));
  {$ENDIF}
  if Check <> nil then Check.OnClick := GestionOuverturePlan;
  {$IFNDEF EAGLCLIENT}
  Check := TDBCheckBox(GetControl('PFE_CLOTURE'));
  {$ELSE}
  Check := TCheckBox(GetControl('PFE_CLOTURE'));
  {$ENDIF}
  if Check <> nil then Check.OnClick := GestionOuverturePlan;
  {$IFNDEF EAGLCLIENT}
  Check := TDBCheckBox(GetControl('PFE_OUVREPREV'));
  {$ELSE}
  Check := TCheckBox(GetControl('PFE_OUVREPREV'));
  {$ENDIF}
  if Check <> nil then Check.OnClick := GestionOuvertureBudget;
  {$IFNDEF EAGLCLIENT}
  Check := TDBCheckBox(GetControl('PFE_CLOTUREPREV'));
  {$ELSE}
  Check := TCheckBox(GetControl('PFE_CLOTUREPREV'));
  {$ENDIF}
  if Check <> nil then Check.OnClick := GestionOuvertureBudget;
  if VH_Paie.PGForPrevisionnel = False then
  begin
        SetControlVisible('PFE_OUVREPREV',False);
        SetControlVisible('PFE_CLOTUREPREV',False);
        SetControlVisible('TPFE_TAUXBUDGET',False);
        SetControlVisible('PFE_TAUXBUDGET',False);
  end;

  If GetParamSocSecur('SO_IFDEFCEGID', False) Then SetControlVisible('PFE_ENCOURS', True); //PT3
end;

procedure TOM_EXERFORMATION.GestionOuverturePlan(Sender: TObject);
begin
  {$IFNDEF EAGLCLIENT}
  if TDBCheckBox(Sender) = nil then Exit;
  if TDBCheckBox(Sender).Name = 'PFE_ACTIF' then
    {$ELSE}
  if TCheckBox(Sender) = nil then Exit;
  if TCheckBox(Sender).Name = 'PFE_ACTIF' then
    {$ENDIF}
  begin
    if GetCheckBoxState('PFE_ACTIF') = CbChecked then
    begin
      Setfield('PFE_CLOTURE', '-');
      SetControlEnabled('PFE_CLOTURE', False);
    end
    else SetControlEnabled('PFE_CLOTURE', True);
  end;
  {$IFNDEF EAGLCLIENT}
  if TDBCheckBox(Sender).Name = 'PFE_CLOTURE' then
    {$ELSE}
  if TCheckBox(Sender).Name = 'PFE_CLOTURE' then
    {$ENDIF}
  begin
    if GetcheckBoxState('PFE_CLOTURE') = CbChecked then
    begin
      Setfield('PFE_ACTIF', '-');
      SetControlEnabled('PFE_ACTIF', False);
    end
    else SetControlEnabled('PFE_ACTIF', True);
  end;
end;

procedure TOM_EXERFORMATION.GestionOuvertureBudget(Sender: TObject);
begin
  {$IFNDEF EAGLCLIENT}
  if TDBCheckBox(Sender) = nil then Exit;
  if TDBCheckBox(Sender).Name = 'PFE_OUVREPREV' then
    {$ELSE}
  if TCheckBox(Sender) = nil then Exit;
  if TCheckBox(Sender).Name = 'PFE_OUVREPREV' then
    {$ENDIF}
  begin
    if GetCheckBoxState('PFE_OUVREPREV') = CbChecked then
    begin
      Setfield('PFE_CLOTUREPREV', '-');
      SetControlEnabled('PFE_CLOTUREPREV', False);
    end
    else SetControlEnabled('PFE_CLOTUREPREV', True);
  end;
  {$IFNDEF EAGLCLIENT}
  if TDBCheckBox(Sender).Name = 'PFE_CLOTUREPREV' then
    {$ELSE}
  if TCheckBox(Sender).Name = 'PFE_CLOTUREPREV' then
    {$ENDIF}
  begin
    if getCheckBoxState('PFE_CLOTUREPREV') = CbChecked then
    begin
      Setfield('PFE_OUVREPREV', '-');
      SetControlEnabled('PFE_OUVREPREV', False);
    end
    else SetControlEnabled('PFE_OUVREPREV', True);
  end;
end;

initialization
  registerclasses([TOM_EXERFORMATION]);
end.

