{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/01/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MULSECTION ()
Mots clefs ... : TOF;MULSECTION
*****************************************************************}
Unit VISUSECTION_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$ifdef eAGLClient}
  MenuOLX,MaineAGL,eTablette,
{$else}
  Tablette,FE_Main,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}
{$endif eAGLClient}

{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     Ent1,
     Mul ; 

Type
  TOF_VISUSECTION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    procedure AxeOnChange(sender: TObject);
    procedure FListeDblClick(sender: TObject);
  end ;

Implementation

procedure TOF_VISUSECTION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_VISUSECTION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_VISUSECTION.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_VISUSECTION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_VISUSECTION.AxeOnChange ( sender : TObject ) ;
var Axe : string ;
begin
  if GetControlText('S_AXE') = '' then
    begin
    SetControlProperty('S_SECTION','DATATYPE', '' ) ;
    SetControlProperty('S_SECTION','ELIPSISBUTTON', false ) ;
    end
  else
    begin
    Axe := GetControlText('S_AXE') ;
    SetControlProperty('S_SECTION','ELIPSISBUTTON', true ) ;
    Case Axe[2] of
     '1' : SetControlProperty('S_SECTION','DATATYPE', 'TZSECTION' ) ;
     '2' : SetControlProperty('S_SECTION','DATATYPE', 'TZSECTION2' ) ;
     '3' : SetControlProperty('S_SECTION','DATATYPE', 'TZSECTION3' ) ;
     '4' : SetControlProperty('S_SECTION','DATATYPE', 'TZSECTION4' ) ;
     '5' : SetControlProperty('S_SECTION','DATATYPE', 'TZSECTION5' ) ;
     end ;
    end ;
end ;

procedure TOF_VISUSECTION.FListeDblClick ( sender : TObject ) ;
var S : string ;
begin
  S := TFMul(Ecran).Q.FindField('S_AXE').AsString +';'+ TFMul(Ecran).Q.FindField('S_SECTION').AsString ;
  AGLLanceFiche('CP','CPSECTION','' , S,'' ) ;
end ;

procedure TOF_VISUSECTION.OnArgument (S : String ) ;
begin
  Inherited ;
  LibellesTableLibre(TTabSheet(GetControl('PZLIBRE')),'TS_TABLE','S_TABLE','S') ;
  THValComboBox(GetControl('S_AXE')).onChange := AxeOnChange ;
  TFMul(Ecran).FListe.OnDblClick := FListeDblClick ;
end ;

procedure TOF_VISUSECTION.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_VISUSECTION ] ) ;
end.
