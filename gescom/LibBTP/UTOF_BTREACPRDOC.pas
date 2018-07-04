{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 08/01/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTREACPRDOC ()
Mots clefs ... : TOF;BTREACPRDOC
*****************************************************************}
Unit UTOF_BTREACPRDOC ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     AglInit,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     uTob,
     maineagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOB,
     UTOF ;

Type
  TOF_BTREACPRDOC = Class (TOF)
  private
    TOBRESULT,TOBpiece : TOB;
    DPA,DPR,DPV : TcheckBox;
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

procedure DefiniModeReactualisation (var Mode : integer; PVBloque : string; TOBpiece : TOB);

Implementation
uses factutil;

procedure DefiniModeReactualisation (var Mode : integer; PVBloque : string; TOBPiece : TOB);
var TOBR : TOB;
		PRBLOQUE : string;
BEGIN
	if IsExisteFraisChantier(TOBpiece) then PRBLOQUE :='X' else PRBLOQUE := '-';
  TOBR := TOB.create ('THE RESULT',nil,-1);
  TRY
    TOBR.AddChampSupValeur ('MODE',-1);
    TOBR.AddChampSupValeur ('PRBLOQUE',PRBLOQUE);
    TOBR.AddChampSupValeur ('PVBLOQUE',PVBloque);
    TOBR.data := TOBpiece;
    TheTOB := TOBR;
    AGLLanceFiche('BTP', 'BTREACPRDOC', '', '', '');
    if (TheTob <> nil) and (TheTOB.GetValue('MODE')>0) then Mode :=TheTOB.GetValue('MODE');
  FINALLY
    TheTOB := nil;
    TOBR.free;
  END;
END;

procedure TOF_BTREACPRDOC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTREACPRDOC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTREACPRDOC.OnUpdate ;
var result : integer;
begin
  Inherited ;
  result := 0;
  if DPA.Checked then result := result + 1;
  if DPR.Checked then result := result + 2;
  if DPV.Checked then result := result + 4;
  if result > 0 then TOBRESULT.Putvalue ('MODE',result);
  TheTOB := TOBRESULT;
end ;

procedure TOF_BTREACPRDOC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTREACPRDOC.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBResult := laTOB;
  TOBPiece := TOB(Tobresult.data);
  DPA := TcheckBox(getControl('CDPA'));
  DPR := TcheckBox(getControl('CDPR'));
  DPV := TcheckBox(getControl('CPV'));
  if laTob.getValue('PVBLOQUE')='X' then DPV.Enabled := false;
  if laTob.getValue('PRBLOQUE')='X' then DPR.Enabled := false;
  if (TOBPiece.getValue('GP_COEFFC')<>0) or (TOBPiece.getValue('GP_COEFFR')<>0) then DPR.enabled := false;
end ;

procedure TOF_BTREACPRDOC.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTREACPRDOC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTREACPRDOC.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTREACPRDOC ] ) ;
end.

