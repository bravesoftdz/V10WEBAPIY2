{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : Source TOF de la FICHE : MFOPARAMPOSTE ()
Mots clefs ... : TOF;MFOPARAMPOSTE
*****************************************************************}
Unit MFOPARAMPOSTE_TOF ;

Interface

Uses
  StdCtrls, Classes,
{$IFNDEF EAGLCLIENT}
  dbtables,
{$ENDIF}
  UTOB, HCtrls, HEnt1, UTOF;

Type
  TOF_MFOPARAMPOSTE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

uses
  FODefi, FOUtil;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : OnNew
Mots clefs ... :
*****************************************************************}
procedure TOF_MFOPARAMPOSTE.OnNew ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : OnDelete
Mots clefs ... :
*****************************************************************}
procedure TOF_MFOPARAMPOSTE.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : OnUpdate
Mots clefs ... :
*****************************************************************}
procedure TOF_MFOPARAMPOSTE.OnUpdate ;
var
  CaisseMono: boolean;
begin
  Inherited ;
  if GetControlVisible('CAISSEMONOPOSTE') then
    CaisseMono := (GetCheckBoxState('CAISSEMONOPOSTE') = cbChecked)
  else
    CaisseMono := False;
  if CaisseMono <> FOCaisseIsMonoPoste then
    SaveSynRegKey(CAISSEMONOPOSTE, CaisseMono, True) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : OnLoad
Mots clefs ... :
*****************************************************************}
procedure TOF_MFOPARAMPOSTE.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}
procedure TOF_MFOPARAMPOSTE.OnArgument (S : String ) ;
var
  sSql: string;
  QQ: TQuery;
  Visu, Ok: boolean;
begin
  Inherited ;
  Visu := False;
{$IFDEF EAGLCLIENT}
  Ok := False;      // pas de caisse monoposte en CWAS
{$ELSE}
 {$IFDEF GESCOM}
  Ok := False;      // pas de caisse monoposte en Vente Comptoir
 {$ELSE}
  Ok := True;
 {$ENDIF}
{$ENDIF}
  if Ok then
  begin
    sSql := 'SELECT COUNT(*) FROM PARCAISSE WHERE GPK_FERME="-"';
    QQ := OpenSQL(sSql, True);
    if (not QQ.EOF) and (QQ.Fields[0].AsInteger = 1) then Visu := True;
    Ferme(QQ);
  end;
  SetControlVisible('CAISSEMONOPOSTE', Visu) ;
  SetControlChecked('CAISSEMONOPOSTE', FOCaisseIsMonoPoste);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : OnClose
Mots clefs ... :
*****************************************************************}
procedure TOF_MFOPARAMPOSTE.OnClose ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : OnDisplay
Mots clefs ... :
*****************************************************************}
procedure TOF_MFOPARAMPOSTE.OnDisplay () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 01/08/2003
Modifié le ... : 01/08/2003
Description .. : OnCancel
Mots clefs ... :
*****************************************************************}
procedure TOF_MFOPARAMPOSTE.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_MFOPARAMPOSTE ] ) ;
end.
