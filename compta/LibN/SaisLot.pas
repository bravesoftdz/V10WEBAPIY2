{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 09/02/2004
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit SaisLot;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFDEF EAGLCLIENT}
  UTOb,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}

  StdCtrls, Mask, Hctrls, HTB97, ExtCtrls, HSysMenu, Ent1, HEnt1, hmsgbox, SaisUtil ;

Function SaisieCodeLot(Var CodeLot : String ; smp : tSuiviMP = smpAucun ;Abrege : String = '' ; Libre : String = '') : Boolean ;

type
  TFSaisLot = class(TForm)
    HMTrad: THSystemMenu;
    Panel1: TPanel;
    Dock: TDock97;
    PanelBouton: TToolWindow97;
    BOuvrir: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    HLabel3: THLabel;
    E_NOMLOT: THCritMaskEdit;
    HM: THMsgBox;
    PrefE_NOMLOT: TEdit;
    procedure FormShow(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure E_NOMLOTElipsisClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    CodeLot,CodeLotResult : String ;
    Abrege,Libre : String ;
    PreparationLot : Boolean ;
    Smp : tSuiviMP ;
  public

  end;

implementation

{$R *.DFM}

Uses MulSMPUtil,LookUp ;

Function SaisieCodeLot(Var CodeLot : String ; smp : tSuiviMP = smpAucun ;Abrege : String = '' ; Libre : String = '') : Boolean ;
var X : TFSaislot ;
    ii : Integer ;
BEGIN
X:=TFSaislot.Create(Application) ;
 Try
  X.CodeLot:=CodeLot ; X.Abrege:=Abrege ; X.Libre:=Libre ;
  X.PreparationLot:=Abrege<>'' ; X.Smp:=Smp ;
  ii:=X.ShowModal ;
  Result:=(ii=mrOk) And (X.CodeLotResult<>'') ;
  If Result Then
    BEGIN
    CodeLot:=X.CodeLotResult ;
    END ;
 Finally
  X.Free ;
 End ;
SourisNormale ;
END ;

procedure TFSaisLot.FormShow(Sender: TObject);
Var pref,St : String ;
begin
E_NOMLOT.Text:=CodeLot ;
If PreparationLot And (smp<>smpAucun) Then
  BEGIN
  Pref:=AttribSoucheLotPreparation(smp) ;
  E_NOMLOT.Plus:=AttribPlus(smp) ;
  PrefE_NOMLOT.Text:=Pref ;
  If CodeLot<>'' Then
    BEGIN
    St:=Copy(CodeLot,1,3) ;
    If St=Pref Then E_NOMLOT.Text:=Copy(CodeLot,4,Length(CodeLot)-3) ;
    END ;
  END ;
(*
If Abrege<>'' Then
  BEGIN
  E_NOMLOT.SelStart:=Length(CodeLot) ;
  E_NOMLOT.SelLength:=0 ;
  END ;
*)
end;

procedure TFSaisLot.BOuvrirClick(Sender: TObject);
Var Q : TQuery ;
    OkOk : Boolean ;
    SQL : String ;
    i : Integer ;
begin
(**)
OkOk:=TRUE ;
If E_NOMLOT.Text='' Then BEGIN HM.Execute(0,Caption,'') ; E_NOMLOT.SetFocus ; BOuvrir.ModalResult:=mrNone ; Exit ; END ;
CodeLotResult:=PrefE_NOMLOT.Text+E_NOMLOT.Text ;
SQL:='SELECT * FROM CHOIXEXT WHERE YX_TYPE="CNL" AND YX_CODE="'+CodeLotResult+'"' ;
Q:=OpenSQL(SQL,TRUE) ;
If Not Q.Eof Then i:=1 Else i:=2 ;
Ferme(Q) ;
If HM.Execute(i,Caption,'')<>mrYes Then OkOk:=FALSE ;
BOuvrir.ModalResult:=mrNone ;
If Not OkOk Then BEGIN CodeLotResult:='' ; Exit ; END ;
BOuvrir.ModalResult:=mrOk ;
CreerCodeLot(CodeLotResult,TRUE,Abrege,Libre) ;
end;

procedure TFSaisLot.E_NOMLOTElipsisClick(Sender: TObject);
begin
If LookupCombo(E_NOMLOT) And PreparationLot And (E_NOMLOT.Text<>'') Then E_NOMLOT.Text:=Copy(E_NOMLOT.Text,4,Length(E_NOMLOT.Text)-3) ;

end;

procedure TFSaisLot.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if Shift<>[] then Exit ;
if Key=VK_F10 then BEGIN Key:=0 ; BOuvrirClick(Nil) ; ModalResult:=mrOk ; END else
 if Key=VK_ESCAPE then BEGIN Key:=0 ; ModalResult:=mrCancel ; END ;
end;

end.
