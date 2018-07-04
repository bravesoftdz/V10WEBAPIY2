{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/05/2003
Modifié le ... : 22/05/2003
Description .. : Source TOM de la TABLE : TYPEREMISE (TYPEREMISE)
Mots clefs ... : TOM;TYPEREMISE
*****************************************************************}
unit TYPEREMISE_TOM;

interface

uses
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$ENDIF}
  HEnt1, UTOM, UTob;

type
  TOM_TYPEREMISE = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  private
    procedure MetEnErreur(NomChamp: string; NoMsg: integer);
  end;

implementation

const
  // Libellés des messages d'erreur
  TexteMessage: array[1..2] of string = (
    {1}'La date de début est supérieure à la date de fin.',
    {2}'Le % maximum de remise est incorrect.'
    );

  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 22/05/2003
  Modifié le ... : 22/05/2003
  Description .. : Met un champ en erreur et affiche le message associé
  Mots clefs ... :
  *****************************************************************}

procedure TOM_TYPEREMISE.MetEnErreur(NomChamp: string; NoMsg: integer);
begin
  if NomChamp <> '' then SetFocusControl(NomChamp);
  LastError := NoMsg;
  if (NoMsg >= Low(TexteMessage)) and (NoMsg <= High(TexteMessage)) then
    LastErrorMsg := TexteMessage[NoMsg];
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/05/2003
Modifié le ... : 22/05/2003
Description .. : OnNewRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_TYPEREMISE.OnNewRecord;
begin
  inherited;
  SetField('GTR_DATEDEBUT', iDate1900);
  SetField('GTR_DATEFIN', iDate2099);
  SetField('GTR_IMPRIMABLE', 'X');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/05/2003
Modifié le ... : 22/05/2003
Description .. : OnDeleteRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_TYPEREMISE.OnDeleteRecord;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/05/2003
Modifié le ... : 22/05/2003
Description .. : OnUpdateRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_TYPEREMISE.OnUpdateRecord;
var
  DateDeb, DateFin: TDateTime;
  Pourc: integer;
begin
  inherited;
  // Contrôle de la période d'utilisation
  DateDeb := GetField('GTR_DATEDEBUT');
  DateFin := GetField('GTR_DATEFIN');
  if DateDeb > DateFin then
  begin
    MetEnErreur('GTR_DATEFIN', 1);
    Exit;
  end;
  // Contrôle des % maximun de remise
  Pourc := GetField('GTR_REMPOURMAX');
  if (Pourc < 0) or (Pourc > 100) then
  begin
    MetEnErreur('GTR_REMPOURMAX', 2);
    Exit;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/05/2003
Modifié le ... : 22/05/2003
Description .. : OnAfterUpdateRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_TYPEREMISE.OnAfterUpdateRecord;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/05/2003
Modifié le ... : 22/05/2003
Description .. : OnLoadRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_TYPEREMISE.OnLoadRecord;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/05/2003
Modifié le ... : 22/05/2003
Description .. : OnChangeField
Mots clefs ... :
*****************************************************************}

procedure TOM_TYPEREMISE.OnChangeField(F: TField);
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/05/2003
Modifié le ... : 22/05/2003
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOM_TYPEREMISE.OnArgument(S: string);
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/05/2003
Modifié le ... : 22/05/2003
Description .. : OnClose
Mots clefs ... :
*****************************************************************}

procedure TOM_TYPEREMISE.OnClose;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/05/2003
Modifié le ... : 22/05/2003
Description .. : OnCancelRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_TYPEREMISE.OnCancelRecord;
begin
  inherited;
end;

initialization
  registerclasses([TOM_TYPEREMISE]);
end.
