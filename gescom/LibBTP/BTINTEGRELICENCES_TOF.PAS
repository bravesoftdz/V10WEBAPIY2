{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/06/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTINTEGRELICENCES ()
Mots clefs ... : TOF;BTINTEGRELICENCES
*****************************************************************}
Unit BTINTEGRELICENCES_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF ;

Type
  TOF_BTINTEGRELICENCES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure MOMFICFInd (sender : TObject);
  end ;

Implementation

uses TntDialogs,UdroitUtilisation;

procedure TOF_BTINTEGRELICENCES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTINTEGRELICENCES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTINTEGRELICENCES.OnUpdate ;
begin
  Inherited ;
  if PGIAsk ('ATTENTION : Vous allez intégrer un fichier de licence.#13#10 Confirmez-vous le traitement ?')=Mryes then
  begin
    SetInfoActivation (GetControlText('NOMFIC')); 
  end;
end ;

procedure TOF_BTINTEGRELICENCES.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTINTEGRELICENCES.OnArgument (S : String ) ;
begin
  Inherited ;
  THEdit(GetControl('NOMFIC')).OnElipsisClick := MOMFICFInd;
end ;

procedure TOF_BTINTEGRELICENCES.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTINTEGRELICENCES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTINTEGRELICENCES.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTINTEGRELICENCES.MOMFICFInd(sender: TObject);
var Emplace : string;
    XX : THOpenDialog;
begin
  XX := THOpenDialog.Create(application);
  XX.Filter := 'fichiers Licences (*LIC)|*.LIC';
  if XX.Execute then
  begin
    SetControlText('NOMFIC',XX.FileName); 
  end;
  XX.Free;
end;

Initialization
  registerclasses ( [ TOF_BTINTEGRELICENCES ] ) ; 
end.
