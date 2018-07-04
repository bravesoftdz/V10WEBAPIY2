unit InfFicSI;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Ent1, Hent1, Hctrls, StdCtrls, HTB97, {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF}, HStatus ;

Procedure InfoFicSISCO(StFichier : String) ;
Function InfoFicSISCORecup(StFichier : String) : Boolean ;

type
  TFInfoFicSISCO = class(TForm)
    Dock: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label9: TLabel;
    Label8: TLabel;
    SActivite: TEdit;
    SNumDossier: TEdit;
    SNomSoc: TEdit;
    SAdr1Soc: TEdit;
    SAdr2Soc: TEdit;
    SAdr3Soc: TEdit;
    SDateDeb: TEdit;
    SDateFin: TEdit;
    SLgCpt: TEdit;
    SCarCli: TEdit;
    SCarFour: TEdit;
    SAna: TEdit;
    SLgAna: TEdit;
    TTenue: TLabel;
    TStatutExo: TLabel;
    procedure FormShow(Sender: TObject);
    procedure SNumDossierKeyPress(Sender: TObject; var Key: Char);
  private
    { Déclarations privées }
    StFichier : String ;
    OkOk : Boolean ;
    procedure Init ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Procedure InfoFicSISCO(StFichier : String) ;
var FFR: TFInfoFicSISCO;
begin
FFR:=TFInfoFicSISCO.Create(Application) ;
FFR.StFichier:=StFichier ;
try
 FFR.ShowModal ;
finally
 FFR.Free ;
end ;
Screen.Cursor:=SyncrDefault ;
end ;

Function InfoFicSISCORecup(StFichier : String) : Boolean ;
var FFR: TFInfoFicSISCO;
begin
Result:=FALSE ;
FFR:=TFInfoFicSISCO.Create(Application) ;
FFR.StFichier:=StFichier ;
try
  FFR.ShowModal ;
finally
 If FFR.ModalResult=mrOk Then Result:=TRUE;
 FFR.Free ;
end ;
Screen.Cursor:=SyncrDefault ;
end ;

Function ConvertDate(St : String) : String ;
Var D,M,Y : Word ;
    DD : tDateTime ;
BEGIN
Result:='' ;
if Trim(St)='' then Exit ;
Y:=StrToInt(Copy(St,5,2)) ; M:=StrToInt(Copy(St,3,2)) ; D:=StrToInt(Copy(St,1,2)) ;
If Y<90 Then Y:=2000+Y Else Y:=1900+y ;
DD:=EncodeDate(Y,M,D) ;
Result:=DateToStr(DD) ;
END ;

procedure TFInfoFicSISCO.Init ;
Var Fichier : TextFile ;
    St : String ;
    Ok00,Ok01,Ok02,Ok03,Ok04,Ok50 : Boolean ;
    Monnaie : String ;
begin
AssignFile(Fichier,StFichier) ; {$i-} Reset(Fichier) ; {$i+}
If Not Eof(Fichier) THen
  BEGIN
  Readln(Fichier,St) ; If Copy(St,1,11)<>'***DEBUT***' Then Exit ;
  END ;
InitMove(1000,'') ;
Ok00:=FALSE ; Ok01:=FALSE ; Ok02:=FALSE ; Ok03:=FALSE ; Ok04:=FALSE ; Ok50:=FALSE ;
While Not Eof(Fichier) Do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ;
  If (Copy(St,1,2)='00') Then
    BEGIN
    Ok00:=TRUE ;
    SNumDossier.Text:=Copy(St,3,5) ;
    SDateDeb.Text:=ConvertDate(Copy(St,8,6)) ;
    SDateFin.Text:=ConvertDate(Copy(St,14,6)) ;
    SLgCpt.Text:=Copy(St,21,2) ;
    Monnaie:=Trim(Copy(St,36,3)) ;
    If Monnaie='EUR' Then TTenue.Caption:='Tenu en EURO' ;
    END ;
  If (Copy(St,1,2)='01') Then
    BEGIN
    Ok01:=TRUE ;
    SActivite.Text:=Copy(St,33,32) ;
    SNomSoc.Text:=Copy(St,3,30) ;
    END ;
  If (Copy(St,1,2)='02') Then
    BEGIN
    Ok02:=TRUE ;
    SAdr1Soc.Text:=Copy(St,3,32) ;
    SAdr2Soc.Text:=Copy(St,35,32) ;
    SAdr3Soc.Text:=Copy(St,67,32) ;
    END ;
  If (Copy(St,1,2)='03') Then
    BEGIN
    Ok03:=TRUE ;
    SCarCli.Text:=Copy(St,56,1) ;
    SCarFour.Text:=Copy(St,57,1) ;
    SAna.Text:=Copy(St,54,1) ;
    If SAna.Text='N' Then Ok50:=TRUE ;
    END ;
  If (Copy(St,1,2)='04') Then
    BEGIN
    Ok04:=TRUE ;
    If StrToInt(Copy(St,42,1))=0 Then TStatutExo.Caption:='Exercice ouvert' ;
    END ;
  If (Copy(St,1,2)='50') Then
    BEGIN
    Ok50:=TRUE ;
    SLgAna.Text:=Copy(St,44,2) ;
    END ;
  If Ok00 And Ok01 And Ok02 And Ok03 And Ok04 And Ok50 Then Break ;
  END ;
System.Close(Fichier) ;
FiniMove ;
If VH^.RecupSISCOPGI Then BValider.visible:=TRUE ;
end;


procedure TFInfoFicSISCO.FormShow(Sender: TObject);
begin
Caption:=Caption+StFichier ; Init ;

end ;
procedure TFInfoFicSISCO.SNumDossierKeyPress(Sender: TObject;
  var Key: Char);
begin
Key:=#0 ;
end;

end.
