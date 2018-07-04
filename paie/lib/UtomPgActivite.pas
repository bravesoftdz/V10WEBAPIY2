{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 06/05/2003
Modifié le ... : 06/05/2003
Description .. : Gestion des activités de la paie
Mots clefs ... : PAIE;ACTIVITE
*****************************************************************}
{
PT1 08/02/2005 SB V_60 FQ 11863 Affectation du no dossier
PT2 30/08/2007 PH V_80 FQ 14713 Création activité et concept plan de paie
PT3 01/10/2007 FC V_80 Concepts plan de paie
}
unit UtomPgActivite;

interface

uses
  SysUtils,
  classes
{$IFDEF EAGLCLIENT}
  , eFiche
  , Utob
{$ELSE}
  , Fiche
  , db
{$ENDIF}
  , Utom
  , HMsgBox
  , PgOutils2
  , HEnt1;

type
  TOM_PGACTIVITE = class(TOM)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnNewRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
  end;

implementation

uses PgOutils, P5def;

const TexteMessage: array[1..3] of string = (
             {1}'Vous devez renseigner un code activité numérique.',
             {2}'Vous devez renseigner le type predefini.',
             {3}''
    );

{ TOM_PGACTIVITE }

procedure TOM_PGACTIVITE.OnAfterUpdateRecord;
begin
  inherited;
  SetControlEnabled('PAC_ACTIVITE', False);
  SetControlEnabled('PAC_PREDEFINI', False);
end;

procedure TOM_PGACTIVITE.OnArgument(Arguments: string);
begin
  inherited;
  SetControlProperty('PAC_ACTIVITE', 'EditMask', '999');
end;

procedure TOM_PGACTIVITE.OnChangeField(F: TField);
var
  Act, Pred, Mes, vide, TempAct: string;
  CEG, STD, DOS, OKRub: Boolean;
begin
  inherited;
  if (F.FieldName = 'PAC_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Act := GetField('PAC_ACTIVITE');
    Pred := GetField('PAC_PREDEFINI');
    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer d''activité prédéfini CEGID.', 'Accès refusé');
      Pred := 'STD'; // PT2
      SetControlProperty('PAC_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer d''activité prédéfini Standard.', 'Accès refusé');
      Pred := '';
      SetControlProperty('PAC_PREDEFINI', 'Value', Pred);
    end;
    if (Trim(Act) <> '') and (Pred <> '') then
    begin
      OKRub := TestRubrique(Pred, Act, 100);
      if (OkRub = False) or (Act = '000') then
      begin
        Mes := MesTestRubrique('ACT', Pred, 100);
        HShowMessage('2;Code Erroné: ' + Act + ' ;' + mes + ';W;O;O;;;', '', '');
        SetField('PAC_ACTIVITE', vide);
        if Pred <> GetField('PAC_PREDEFINI') then SetField('PAC_PREDEFINI', pred);
        SetFocusControl('PAC_ACTIVITE');
        exit;
      end;
    end;
    if Pred <> GetField('PAC_PREDEFINI') then SetField('PAC_PREDEFINI', pred);
  end;

  if (F.FieldName = 'PAC_ACTIVITE') then
  begin
    Act := Getfield('PAC_ACTIVITE');
    if Act = '' then exit;
    if ((isnumeric(Act)) and (Trim(Act) <> '')) then
    begin
      TempAct := ColleZeroDevant(strtoint(trim(Act)), 3);
      if (DS.State = dsinsert) and (TempAct <> '') and (GetField('PAC_PREDEFINI') <> '') then
      begin
        OKRub := TestRubrique(GetField('PAC_PREDEFINI'), TempAct, 100);
        if (OkRub = False) or (Act = '000') then
        begin
          Mes := MesTestRubrique('ACT', GetField('PAC_PREDEFINI'), 100);
          HShowMessage('2;Code Erroné: ' + TempAct + ' ;' + Mes + ';W;O;O;;;', '', '');
          TempAct := '';
        end;
      end;
      if TempAct <> Act then
      begin
        SetField('PAC_ACTIVITE', TempAct);
        SetFocusControl('PAC_ACTIVITE');
      end;
    end;
  end;
end;

procedure TOM_PGACTIVITE.OnLoadRecord;
var
  CEG, STD, DOS: boolean;
begin
  inherited;
  AccesPredefini('TOUS', CEG, STD, DOS);
  if (Getfield('PAC_PREDEFINI') = 'CEG') then
  begin
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end
  else
    if (Getfield('PAC_PREDEFINI') = 'STD') then
    begin
      PaieLectureSeule(TFFiche(Ecran), (STD = False));
      SetControlEnabled('BDelete', STD);
    end;
//PT3 On ne gère pas les DOS
//    else
//      if (Getfield('PAC_PREDEFINI') = 'DOS') then
//      begin
//        PaieLectureSeule(TFFiche(Ecran), (DOS = False));
//        SetControlEnabled('BDelete', True);
//      end;

// DEB PT2  Remis les lignes au bon endroit
  //DEB PT3
  if (CEG = False) and (STD = False) then
    SetControlEnabled('BInsert', False)
  else
    SetControlEnabled('BInsert', True);
  //FIN PT3
  SetControlEnabled('PAC_PREDEFINI', False);
  SetControlEnabled('PAC_ACTIVITE', False);
  if DS.State in [dsInsert] then
  begin
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('PAC_PREDEFINI', True);
    SetControlEnabled('PAC_ACTIVITE', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
  end;
// FIN PT2
  PaieConceptPlanPaie(Ecran); // PT46
end;

procedure TOM_PGACTIVITE.OnNewRecord;
begin
  inherited;
  SetField ('PAC_PREDEFINI','STD');
end;

procedure TOM_PGACTIVITE.OnUpdateRecord;
begin
  inherited;
  if not IsNumeric(GetField('PAC_ACTIVITE')) then
    LastError := 1
  else
    if GetField('PAC_PREDEFINI') = '' then
      LastError := 2;
  if LastError > 0 then PgiBox(TexteMessage[LastError], Ecran.caption);
{ DEB PT1 }
  if (DS.State = dsinsert) then
  begin
    if (GetField('PAC_PREDEFINI') <> 'DOS') then
      SetField('PAC_NODOSSIER', '000000')
    else
      SetField('PAC_NODOSSIER', PgRendNoDossier);
  end;
{ FIN PT1 }
end;



initialization
  registerclasses([TOM_PGACTIVITE]);
end.

