{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/02/2003
Modifié le ... : 27/02/2003
Description .. : Source TOF de la FICHE : MFOUSERCHG
Mots clefs ... : TOF;MFOUSERCHG
*****************************************************************}
unit MFOUSERCHG_TOF;

interface

uses Classes, sysutils,
  {$IFNDEF EAGLCLIENT}
  dbtables,
  {$ENDIF}
  HCtrls, UTOF, LicUtil, UTOB;

type
  TOF_MFOUSERCHG = class(TOF)
  private
    NoErreur: integer;
  public
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  end;

implementation

uses
  Vierge;

const
  // Libellés des messages d'erreur
  TexteMessage: array[1..3] of string = (
    {1}'Vous devez renseigner un nom d''utilisateur.',
    {2}'Mot de passe incorrect !',
    {3}'Utilisateur inconnu !'
    );

  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 27/02/2003
  Modifié le ... : 27/02/2003
  Description .. : OnNew
  Mots clefs ... : FO
  *****************************************************************}

procedure TOF_MFOUSERCHG.OnNew;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/02/2003
Modifié le ... : 27/02/2003
Description .. : OnDelete
Mots clefs ... : FO
*****************************************************************}

procedure TOF_MFOUSERCHG.OnDelete;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/02/2003
Modifié le ... : 27/02/2003
Description .. : OnUpdate
Mots clefs ... : FO
*****************************************************************}

procedure TOF_MFOUSERCHG.OnUpdate;
var NomUsr, PassWord, Groupe, sSql: string;
  QQ: TQuery;
begin
  inherited;
  NomUsr := GetControlText('US_ABREGE');
  PassWord := GetControlText('US_PASSWORD');
  if NomUsr = '' then
  begin
    NoErreur := 1;
    SetFocusControl('US_ABREGE');
    Exit;
  end;
  // vérification de l'utilisateur
  Groupe := '';
  sSql := 'SELECT US_PASSWORD,US_GROUPE FROM UTILISAT WHERE US_ABREGE="' + NomUsr + '"';
  QQ := OpenSQL(sSql, False);
  if not QQ.EOF then
  begin
    if (CryptageSt(PassWord) = QQ.FindField('US_PASSWORD').AsString) or
      (PassWord = DayPass(Date)) then
    begin
      Groupe := QQ.FindField('US_GROUPE').AsString;
    end else
    begin
      NoErreur := 2;
      SetControlText('US_PASSWORD', '');
      SetFocusControl('US_PASSWORD');
    end;
  end else
  begin
    NoErreur := 3;
    SetControlText('US_ABREGE', '');
    SetFocusControl('US_ABREGE');
  end;
  Ferme(QQ);
  // lecture n° du groupe d'utilisateurs
  if Groupe <> '' then
  begin
    sSql := 'SELECT UG_NUMERO FROM USERGRP WHERE UG_GROUPE="' + Groupe + '"';
    QQ := OpenSQL(sSql, False);
    if not QQ.EOF then
      TFVierge(Ecran).Retour := QQ.FindField('UG_NUMERO').AsString;
    Ferme(QQ);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/02/2003
Modifié le ... : 27/02/2003
Description .. : OnLoad
Mots clefs ... : FO
*****************************************************************}

procedure TOF_MFOUSERCHG.OnLoad;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/02/2003
Modifié le ... : 27/02/2003
Description .. : OnArgument
Mots clefs ... : FO
*****************************************************************}

procedure TOF_MFOUSERCHG.OnArgument(S: string);
begin
  inherited;
  NoErreur := 0;
  TFVierge(Ecran).Retour := '';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/02/2003
Modifié le ... : 27/02/2003
Description .. : OnClose
Mots clefs ... : FO
*****************************************************************}

procedure TOF_MFOUSERCHG.OnClose;
begin
  inherited;
  if NoErreur <> 0 then
  begin
    LastError := NoErreur;
    if (NoErreur >= Low(TexteMessage)) and (NoErreur <= High(TexteMessage)) then
      LastErrorMsg := TexteMessage[NoErreur];
    NoErreur := 0;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/02/2003
Modifié le ... : 27/02/2003
Description .. : OnDisplay
Mots clefs ... : FO
*****************************************************************}

procedure TOF_MFOUSERCHG.OnDisplay();
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 27/02/2003
Modifié le ... : 27/02/2003
Description .. : OnCancel
Mots clefs ... : FO
*****************************************************************}

procedure TOF_MFOUSERCHG.OnCancel();
begin
  inherited;
end;

initialization
  registerclasses([TOF_MFOUSERCHG]);
end.
