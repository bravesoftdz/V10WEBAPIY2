{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTECHANGECPTAS1 ()
Mots clefs ... : TOF;BTECHANGECPTAS1
*****************************************************************}
Unit BTECHANGECPTAS1_TOF ;

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
{$ENDIF}
     uTob, 
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     Utilline,
     BTPARAMSOCGEN_TOF,
     UTOF ;

Type
  TOF_BTECHANGECPTAS1 = Class (TOF_BTPARAMSOCGEN)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	function VerifInfos : boolean;
  end ;

Implementation

procedure TOF_BTECHANGECPTAS1.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTECHANGECPTAS1.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTECHANGECPTAS1.OnUpdate ;
begin
  Inherited ;
  if not verifInfos then
  begin
  	TForm(Ecran).ModalResult:=0;
    exit;
  end;
	StockeInfos(Ecran, LaTob);
end ;

procedure TOF_BTECHANGECPTAS1.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTECHANGECPTAS1.OnArgument (S : String ) ;
var TheTOBValue : TOB;
begin
  Inherited ;
  chargeEcran(ecran, LaTob);
  TheTOBValue := LaTOb.findFirst(['SOC_NOM'],['SO_CPLIENGAMME'],true);
  if (TheTOBValue <> nil ) and (TheTOBValue.GetValue('SOC_DATA')='') then
  begin
  	THLabel(GetControl('LSO_CPPOINTAGESX')).visible := true;
    THValComboBox(GetControl('SO_CPPOINTAGESX')).visible := true;
  end;
  if (TheTOBValue <> nil ) and (TheTOBValue.GetValue('SOC_DATA')='S1') then
  begin
    THCheckBox(GetControl('SO_CPMODESYNCHRO')).Checked := false;
    THCheckBox(GetControl('SO_CPMODESYNCHRO')).visible := false;
  end;
end ;

procedure TOF_BTECHANGECPTAS1.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTECHANGECPTAS1.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTECHANGECPTAS1.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_BTECHANGECPTAS1.VerifInfos: boolean;
begin
	result := false;
  result := true;
end;

Initialization
  registerclasses ( [ TOF_BTECHANGECPTAS1 ] ) ; 
end.

