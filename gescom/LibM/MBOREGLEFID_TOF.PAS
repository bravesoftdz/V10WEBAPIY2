{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : Source TOF de la FICHE : MBOREGLEFID ()
Suite ........ : Saisie du paramétrage de la fidélité client
Mots clefs ... : TOF;MBOREGLEFID
*****************************************************************}
unit MBOREGLEFID_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  db, dbtables, FE_Main,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF,
  UTOB, uTableFiltre, SaisieList, HTB97, AglInit;

type
  TOF_MBOREGLEFID = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    TF: TTableFiltre;
    procedure BNEWPROGClick(Sender: TObject);
    procedure TREEENTETEDblClick(Sender: TObject);
    procedure DoSetNavigate(Sender: TObject);
  end;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : Création d'un nouveau programme de fidélité
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOREGLEFID.BNEWPROGClick(Sender: TObject);
var
  CodeProg, Args: string;
begin
  if (Ecran = nil) or not (Ecran is TFSaisieList) or
    (TFSaisieList(Ecran).TypeAction = taConsult) then
    Exit;

  Args := ActionToString(taCreat) + ';MONOFICHE';
  CodeProg := AGLLanceFiche('MBO', 'MBOPROGRAMMEFID', '', '', 'ACTION=CREATION;MONOFICHE');
  if CodeProg <> '' then
    TF.RefreshEntete(CodeProg);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : Modification d'un programme de fidélité
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOREGLEFID.TREEENTETEDblClick(Sender: TObject);
var
  CodeProg, Args: string;
begin
  if (Ecran <> nil) and (Ecran is TFSaisieList) and
    (TFSaisieList(Ecran).TypeAction <> taConsult) then
    Args := ActionToString(taModif)
  else
    Args := ActionToString(taConsult);
  Args := Args + ';MONOFICHE';

  CodeProg := TF.TOBFiltre.GetValue('GFO_CODEFIDELITE');
  CodeProg := AGLLanceFiche('MBO', 'MBOPROGRAMMEFID', '', CodeProg, Args);
  if CodeProg <> '' then
    TF.RefreshEntete(CodeProg);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOREGLEFID.DoSetNavigate(Sender: TObject);
begin
  SetControlEnabled('BNEWPROG', TF.CanReload);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnNew
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOREGLEFID.OnNew;
begin
  inherited;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnDelete
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOREGLEFID.OnDelete;
var
  sSql: string;
begin
  inherited;
  sSql := 'DELETE PARREGLEFID'
    + ' WHERE GFS_CODEFIDELITE="' + GetControlText('GFO_CODEFIDELITE') + '"'
    + ' AND GFS_REGLEFIDELITE="' + GetControlText('GFO_CODEFIDELITE') + '"';
  ExecuteSQL(sSql);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnUpdate
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOREGLEFID.OnUpdate;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnLoad
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOREGLEFID.OnLoad;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOREGLEFID.OnArgument(S: string);
var
  Ctrl: TControl;
begin
  inherited;
  if (Ecran <> nil) and (Ecran is TFSaisieList) then
  begin
    TF := TFSaisieList(Ecran).LeFiltre;
    TF.OnSetNavigate := DoSetNavigate;
  end;

  Ctrl := GetControl('TreeEntete');
  if (Ctrl <> nil) and (Ctrl is THTreeView) then
    THTreeView(Ctrl).OnDblClick := TREEENTETEDblClick;

  Ctrl := GetControl('BNEWPROG');
  if (Ctrl <> nil) and (Ctrl is TToolBarButton97) then
  begin
    if (TFSaisieList(Ecran).TypeAction = taConsult) then
      SetControlVisible('BNEWPROG', False)
    else TToolBarButton97(Ctrl).OnClick := BNEWPROGClick;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnClose
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOREGLEFID.OnClose;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnDisplay
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOREGLEFID.OnDisplay();
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/04/2003
Modifié le ... : 10/04/2003
Description .. : OnCancel
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOREGLEFID.OnCancel();
begin
  inherited;
end;

initialization
  registerclasses([TOF_MBOREGLEFID]);
end.
