{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 13/10/2006
Modifié le ... :   /  /    
Description .. : Gestion de la fiches des indicateurs calculés
Mots clefs ... : PAIE;BILAN
*****************************************************************}
{
PT1 13/10/2006 SB V65 FQ 13564 Gestion des boutons en mode vérrouillé
}
unit PBSBILANSOCIAL_TOM;

interface

uses
classes,
SysUtils,

{$IFNDEF EAGLCLIENT}
Fiche,
db,
{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
eFiche,Utob,
{$ENDIF}

HEnt1,
HMsgBox,
HCtrls,
Utom;


type
  TOM_BILANSOCIAL = class(TOM)
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
    DerniereCreate : String;
    OnFerme  : Boolean;
  end;

implementation
uses pgoutils,pgoutils2,p5def;


{ TOM_BILANSOCIAL }

procedure TOM_BILANSOCIAL.OnAfterUpdateRecord;
begin
  inherited;
if OnFerme then Ecran.Close;

if (LastError = 0) and  (not (DS.State in [dsInsert])) then
  Begin
  SetControlEnabled('PBC_INDICATEURBS',False);
  SetControlEnabled('PBC_DATEDEBUT',False);
  SetControlEnabled('PBC_DATEFIN',False);
  SetControlEnabled('PBC_BSPRESENTATION',False);
  SetControlEnabled('PBC_INCREM',False);
  End;

if GetField('PBC_PGINDICATION')='VER' then  SetControlEnabled('BDELETE',False);   { PT1 }

end;

procedure TOM_BILANSOCIAL.OnArgument(S: string);
var Num : integer;
begin
  inherited;
//VisibiliteOrganisation(TFFiche(Ecran));
for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PBC_TRAVAILN' + IntToStr(Num)), GetControl('TPBC_TRAVAILN' + IntToStr(Num)));
  VisibiliteStat(GetControl('PBC_CODESTAT'), GetControl('TPBC_CODESTAT'));
  VisibiliteChampLibreSal(IntToStr(Num), GetControl('PBC_LIBREPCMB' + IntToStr(Num)), GetControl('TPBC_LIBREPCMB' + IntToStr(Num)))
  end;

end;

procedure TOM_BILANSOCIAL.OnCancelRecord;
begin
  inherited;

end;

procedure TOM_BILANSOCIAL.OnChangeField(F: TField);
begin
  inherited;

end;

procedure TOM_BILANSOCIAL.OnClose;
begin
  inherited;

end;

procedure TOM_BILANSOCIAL.OnDeleteRecord;
begin
  inherited;

end;

procedure TOM_BILANSOCIAL.OnLoadRecord;
begin
  inherited;
if not (DS.State in [dsInsert]) then DerniereCreate := '';

PaieLectureSeule(TFFiche(Ecran)  , (GetField('PBC_PGINDICATION')='VER') );
SetControlEnabled('PBC_PGINDICATION',True);
SetControlEnabled('BValider',True);     { PT1 }
SetControlEnabled('BInsert' ,True);     { PT1 }


if not (DS.State in [dsInsert]) then
  Begin
  SetControlEnabled('PBC_INDICATEURBS',False);
  SetControlEnabled('PBC_DATEDEBUT',False);
  SetControlEnabled('PBC_DATEFIN',False);
  SetControlEnabled('PBC_BSPRESENTATION',False);
  SetControlEnabled('PBC_INCREM',False);
  End;
end;

procedure TOM_BILANSOCIAL.OnNewRecord;
begin
  inherited;
SetField('PBC_DATEDEBUT' ,DateToStr(DEBUTANNEE(Date)));
SetField('PBC_DATEFIN'   ,DateToStr(FINANNEE(Date)));
SetField('PBC_NODOSSIER' ,PgRendNoDossier);
SetField('PBC_PGINDICATION', 'PER');
SetControlEnabled('PBC_PGINDICATION',False);  
end;

procedure TOM_BILANSOCIAL.OnUpdateRecord;
Var St : string;
    Q : TQuery;
begin
  inherited;
  OnFerme := False;
  If (IsValidDate(GetField('PBC_DATEDEBUT'))) And  (IsValidDate(GetField('PBC_DATEFIN'))) then
    if (DS.State in [dsInsert]) then
       DerniereCreate := DateToStr(GetField('PBC_DATEDEBUT'))+';'+
                         DateToStr(GetField('PBC_DATEFIN'))+';'+
                         GetField('PBC_BSPRESENTATION')+';'+
                         GetField('PBC_INDICATEURBS')
  else
    if (DerniereCreate = DateToStr(GetField('PBC_DATEDEBUT'))+';'+
                         DateToStr(GetField('PBC_DATEFIN'))+';'+
                         GetField('PBC_BSPRESENTATION')+';'+
                         GetField('PBC_INDICATEURBS')) then OnFerme := True; // le bug arrive on se casse !!!
  //PBC_INDICATEURBS,PBC_DATEDEBUT,PBC_DATEFIN,PBC_BSPRESENTATION,PBC_INCREM
  If not IsValidDate(GetField('PBC_DATEDEBUT')) then
  Begin
  LastError := 1;
  PgiBox('Vous devez saisir une date de début valide.',Ecran.Caption);
  exit;
  End;

  If not IsValidDate(GetField('PBC_DATEFIN')) then
  Begin
  LastError := 1;
  PgiBox('Vous devez saisir une date de fin valide.',Ecran.Caption);
  exit;
  End;

  If (GetField('PBC_BSPRESENTATION')='') then
  Begin
  LastError := 1;
  PgiBox('Vous devez saisir une présentation.',Ecran.Caption);
  exit;
  End;

  If (GetField('PBC_ETABLISSEMENT')='') then
  Begin
  LastError := 1;
  PgiBox('Vous devez saisir l''établissement.',Ecran.Caption);
  exit;
  End;

  If (GetField('PBC_CATBILAN')='') then
  Begin
  LastError := 1;
  PgiBox('Vous devez saisir une catégorie.',Ecran.Caption);
  exit;
  End;

  If (DS.State = DsInsert) then
    Begin
    St := 'SELECT MAX(PBC_INCREM) AS MAXINCREM FROM BILANSOCIAL '+
          'WHERE PBC_INDICATEURBS="'+GetField('PBC_INDICATEURBS')+'" '+
          'AND PBC_DATEDEBUT="'+UsDateTime(GetField('PBC_DATEDEBUT'))+'" '+
          'AND PBC_DATEFIN="'+UsDateTime(GetField('PBC_DATEFIN'))+'" '+
          'AND PBC_BSPRESENTATION="'+GetField('PBC_BSPRESENTATION')+'" ';
    Q := OpenSql(St,True);
    If not Q.Eof then
          SetField('PBC_INCREM',Q.FindField('MAXINCREM').Asinteger + 1)
    else
          SetField('PBC_INCREM',1);
    Ferme(Q);
    End;
end;

initialization
  registerclasses([TOM_BILANSOCIAL]);
end.
 
