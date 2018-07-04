{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 09/02/2004
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit DetruitLot;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFDEF EAGLCLIENT}
  UTOb,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  StdCtrls, Mask, Hctrls, HTB97, ExtCtrls, HSysMenu, Ent1, HEnt1, hmsgbox, SaisUtil ;

Function DetruitLeLot(smp : tSuiviMP) : Boolean ;

type
  TFKillLot = class(TForm)
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
    { Déclarations privées }
    Smp : tSuiviMP ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses MulSMPUtil,LookUp ;

Function DetruitLELot(smp : tSuiviMP) : Boolean ;
var X : TFKillLot ;
    ii : Integer ;
BEGIN
X:=TFKillLot.Create(Application) ;
 Try
  X.Smp:=Smp ;
  ii:=X.ShowModal ;
  Result:=(ii=mrOk)  ;
  If Result Then
    BEGIN
    END ;
 Finally
  X.Free ;
 End ;
SourisNormale ;
END ;

procedure TFKillLot.FormShow(Sender: TObject);
Var
  pref : String ;
begin
If (smp<>smpAucun) Then
  BEGIN
  Pref:=AttribSoucheLotPreparation(smp) ;
  E_NOMLOT.Plus:=AttribPlus(smp) ;
  PrefE_NOMLOT.Text:=Pref ;
  END ;
(*
If Abrege<>'' Then
  BEGIN
  E_NOMLOT.SelStart:=Length(CodeLot) ;
  E_NOMLOT.SelLength:=0 ;
  END ;
*)
end;

procedure TFKillLot.BOuvrirClick(Sender: TObject);
Var Q : TQuery ;
    OkOk : Boolean ;
    SQL : String ;
    i : Integer ;
    CodeLotResult : String ;
begin
(**)
OkOk:=TRUE ;
If E_NOMLOT.Text='' Then BEGIN HM.Execute(0,Caption,'') ; E_NOMLOT.SetFocus ; BOuvrir.ModalResult:=mrNone ; Exit ; END ;
CodeLotResult:=PrefE_NOMLOT.Text+E_NOMLOT.Text ;
SQL:='SELECT * FROM CHOIXEXT WHERE YX_TYPE="CNL" AND YX_CODE="'+CodeLotResult+'"' ;
Q:=OpenSQL(SQL,TRUE) ;
If Q.Eof Then i:=1 Else i:=2 ;
Ferme(Q) ;
If i=2 Then
  BEGIN
  SQL:='SELECT E_NOMLOT FROM ECRITURE WHERE E_NOMLOT="'+CodeLotResult+'" AND (E_ETATLETTRAGE="AL" Or E_ETATLETTRAGE="PL")' ;
  Q:=OpenSQL(SQL,TRUE) ;
  If Not Q.Eof Then i:=3 ;
  Ferme(Q) ;
  END ;
If HM.Execute(i,Caption,'')<>mrYes Then OkOk:=FALSE ;
If (i=1) Or (i=3) Then Exit ;
BOuvrir.ModalResult:=mrNone ;
If Not OkOk Then BEGIN CodeLotResult:='' ; Exit ; END ;
BOuvrir.ModalResult:=mrOk ;
SQL:='DELETE FROM CHOIXEXT WHERE YX_TYPE="CNL" AND YX_CODE="'+CodeLotResult+'"' ;
ExecuteSQL(SQL) ;
end;

procedure TFKillLot.E_NOMLOTElipsisClick(Sender: TObject);
begin
If LookupCombo(E_NOMLOT) And (E_NOMLOT.Text<>'') Then E_NOMLOT.Text:=Copy(E_NOMLOT.Text,4,Length(E_NOMLOT.Text)-3) ;
end;

procedure TFKillLot.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if Shift<>[] then Exit ;
if Key=VK_F10 then BEGIN Key:=0 ; BOuvrirClick(Nil) ; ModalResult:=mrOk ; END else
 if Key=VK_ESCAPE then BEGIN Key:=0 ; ModalResult:=mrCancel ; END ;
end;

end.
