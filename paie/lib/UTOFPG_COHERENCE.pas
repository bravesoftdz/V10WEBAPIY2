{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/08/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : COHERENCE
Mots clefs ... :
*****************************************************************}
unit UTOFPG_COHERENCE;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFDEF EAGLCLIENT}
  MaineAgl,
{$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  Vierge,
  UTOF;

type
  TOF_PGCOHERENCE = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    procedure BValiderClick(Sender: TObject);
  end;


implementation

uses Buttons, HStatus;

procedure TOF_PGCOHERENCE.OnNew;
begin
  inherited;
end;

procedure TOF_PGCOHERENCE.OnDelete;
begin
  inherited;
end;

procedure TOF_PGCOHERENCE.OnUpdate;
begin
  inherited;
end;

procedure TOF_PGCOHERENCE.OnLoad;
begin
  inherited;
end;

procedure TOF_PGCOHERENCE.OnArgument(S: string);
var b: tBitBtn;
begin
  inherited;
  TFVierge(Ecran).Retour := '';
  b := TBitBtn(GetControl('BVALIDER'));
  if b <> nil then b.OnClick := BValiderClick;
end;

procedure TOF_PGCOHERENCE.BValiderClick(Sender: TObject);
begin
  if GetControlText('FOUI') <> 'OUI' then
  begin
    PgiError('Cette opération est impossible : le mot de passe est incorrect');
    Exit;
  end;
  TFVierge(Ecran).Retour := 'X';
end;

procedure TOF_PGCOHERENCE.OnClose;
begin
  inherited;
end;

initialization
  registerclasses([TOF_PGCOHERENCE]);
end.

