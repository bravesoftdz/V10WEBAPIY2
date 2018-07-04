unit ReinitCA;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hcompte, Ent1, HEnt1, Mask, Hctrls, HSysMenu,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} 
  Buttons, ExtCtrls, hmsgbox, CpteSAV ;

Procedure ReinitCodeAccept ;

type
  TFInitCA = class(TForm)
    HMTrad: THSystemMenu;
    Panel1: TPanel;
    TFDateCpta1: THLabel;
    FDateCpta1: TMaskEdit;
    TFDateCpta2: TLabel;
    FDateCpta2: TMaskEdit;
    TFEtab: THLabel;
    FEtab: THValComboBox;
    TFNumPiece1: THLabel;
    FNumPiece1: TMaskEdit;
    FNumPiece2: TMaskEdit;
    TFExercice: THLabel;
    FExercice: THValComboBox;
    HPB: TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    Label1: TLabel;
    HMess: THMsgBox;
    TCPTEDEBUT: THLabel;
    CPTEDEBUT: THCpteEdit;
    TCPTEFIN: THLabel;
    CPTEFIN: THCpteEdit;
    procedure FExerciceChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
  private
    { Déclarations privées }
    Procedure LanceTraitement ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  HStatus;


Procedure ReinitCodeAccept ;
var TT : TFInitCA;
BEGIN
TT:=TFInitCA.Create(Application) ;
 Try
  TT.ShowModal ;
 Finally
  TT.Free ;
 End ;
END ;


procedure TFInitCA.FExerciceChange(Sender: TObject);
begin
If FExercice.ItemIndex>0 Then ExoToDates(FExercice.Value,FDateCpta1,FDateCpta2) ;
end;

procedure TFInitCA.FormShow(Sender: TObject);
begin
FExercice.Value:=VH^.Encours.Code ;
end;

Procedure TFInitCA.LanceTraitement ;
Var St : String ;
    Exo : TExoDate ;
BEGIN
St:='' ;
FillChar(Exo,SizeOf(Exo),#0) ;
if FExercice.ItemIndex>0 then BEGIN (*St:=St+' AND E_EXERCICE="'+FExercice.Value+'" ' ;*) Exo.Code:=FExercice.Value ; END ;
St:=St+' And E_DATECOMPTABLE>="'+USDATE(FDateCpta1)+'" ' ;
St:=St+' And E_DATECOMPTABLE<="'+USDATE(FDateCpta2)+'" ' ;
Exo.Deb:=StrToDate(FDateCpta1.Text) ; Exo.Fin:=StrToDate(FDateCpta2.Text) ;
if FEtab.ItemIndex>0 then St:=St+' AND E_ETABLISSEMENT="'+FEtab.Value+'" ' ;
St:=St+' And E_NUMEROPIECE>='+FNumPiece1.Text+' and E_NUMEROPIECE<='+FNumPiece2.Text+' ' ;
If CpteDebut.Text<>'' Then St:=St+' AND E_AUXILIAIRE>="'+CpteDebut.Text+'" ' ;
If CpteFin.Text<>'' Then St:=St+' AND E_AUXILIAIRE<="'+CpteFin.Text+'" ' ;

InitCodeAccept(St,Exo) ;
END ;

procedure TFInitCA.BValiderClick(Sender: TObject);
Var i : Integer ;
begin
i:=HMess.Execute(0,'','') ;
If i<>mrYes Then Exit ;
Try
 BeginTrans ;
 LanceTraitement ;
 CommitTrans ;
Except
 RollBack ;
End ;
end;

end.
