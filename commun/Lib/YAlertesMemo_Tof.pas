{***********UNITE************************************************* 
Auteur  ...... : Garnier MN
Créé le ...... : 24/01/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ALERTESMEMO ()
Mots clefs ... : TOF;ALERTESMEMO
*****************************************************************}
Unit YAlertesMemo_Tof ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     forms, uTob,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,wTof,extctrls,wCommuns,HRichOLE, HTB97,Vierge ;

Type
  TOF_ALERTESMEMO = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  Private
    TobAlertes: tob;
    procedure JAILU_OnClick(Sender: TObject);
    procedure BFerme_OnClick(Sender: TObject);
  end ;

Implementation

uses
  UtilPGI,YAlertesConst
  ;

procedure TOF_ALERTESMEMO.OnArgument (S : String ) ;
var Memo : TRichEdit;
    ThePanel:TPanel;
begin
inherited ;
  TFVierge (ecran) .Retour := 'X';
  TobAlertes := Tob(GetArgumentInteger(S, 'TOBALERTES'   ));
  if not Assigned (TobAlertes) then exit;
  if not Assigned (GetControl('JAILU')) then exit;
  if not Assigned (GetControl('BValider')) then exit;

  if TobAlertes.GetValue('YAL_VALIDATION') = 'S' then
    SetControlVisible('JAILU',false)
  else
    begin
    SetControlVisible('BValider',false);
    SetControlVisible('BFerme',false);
    end;
  TCheckBox(GetControl('JAILU')).OnClick := JAILU_OnClick;

  ThePanel:=TPanel(GetControl('PAN_ALERTE'));
  Memo := TRichEdit.Create(ThePanel) ;
  Memo.Parent:=ThePanel;
  Memo.Align:=alClient;
  Memo.ReadOnly:=true;
  StringTorich(Memo,TobAlertes.GetValue('YAL_BLOCNOTE'));
  Ecran.Caption := 'Alerte '+Rechdom('YMODEALERTE',TobAlertes.GetString('YAL_MODEBLOCAGE'),false);
  UpdateCaption (ecran);
  if TobAlertes.GetString('YAL_MODEBLOCAGE') = Interrogative then
    begin
    if TobAlertes.GetValue('YAL_VALIDATION') = 'S' then
      SetControlVisible('BFerme',true);
    if Assigned(GetControl('BFerme')) then
      TToolBarButton97(GetControl('BFerme')).OnClick := BFerme_OnClick;
    end;
end;

procedure TOF_ALERTESMEMO.BFerme_OnClick(Sender: TObject);
begin
  TFVierge (ecran) .Retour := '-';
end;

procedure TOF_ALERTESMEMO.JAILU_OnClick(Sender: TObject);
begin
if TCheckBox(GetControl('JAILU')).checked=true then
   begin
   if TobAlertes.GetString('YAL_MODEBLOCAGE') <> Blocage then
     SetControlVisible('BValider',true);
   if TobAlertes.GetString('YAL_MODEBLOCAGE') <> SansBlocage then
     SetControlVisible('BFerme',true);
   end
else
   begin
   SetControlVisible('BValider',false);
   SetControlVisible('BFerme',false);
   end
end;

procedure TOF_ALERTESMEMO.OnClose ;
begin
  if (TobAlertes.GetValue('YAL_VALIDATION') <> 'S') and
    ( not TCheckBox(GetControl('JAILU')).checked ) then
    LastError := 1 ;
end;

Initialization
  registerclasses ( [ TOF_ALERTESMEMO ] ) ;
end.

