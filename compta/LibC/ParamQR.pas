unit ParamQR;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, ExtCtrls, Ent1, HEnt1, HCtrls,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  HStatus,
  HTB97, StdCtrls, Hcompte, Formule;

Const pipe='|' ;
Type tParamCol = Record
                 OkOk : Boolean ;
                 Titre : String ;
                 Titre1,Titre2 : String ;
                 Value1,Value2 : String ;
                 DebitPos1,DebitPos2 : Boolean ;
                 Formule1,Formule2 : String ;
                 End ;

Type tCHAMPTABLE = Array[0..9] Of String ;
Type tCHAMPTABLE11 = Array[0..11] Of String ;

Type tCritEdtSup = Record
                   PC : Array[1..4] Of tParamCol ;
                   CHAMPTABLE : tCHAMPTABLE ;
                   End ;

Function ParamCol(Var St : String ; NumC : Integer ; CodeTableNat,LibTableNat : String) : Boolean ;

type
  TFParamQR = class(TForm)
    HMTrad: THSystemMenu;
    HMess: THMsgBox;
    Dock: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Panel1: TPanel;
    HTC1: TLabel;
    TC1: TEdit;
    HTC11: TLabel;
    TC11: TEdit;
    HTC12: TLabel;
    TC12: TEdit;
    TG_TABLE11: THLabel;
    G_TABLE11: THCpteEdit;
    TG_TABLE12: THLabel;
    G_TABLE12: THCpteEdit;
    CBDEB1: TCheckBox;
    CBDEB2: TCheckBox;
    TFSC11: TLabel;
    FSC11: TEdit;
    TFSC12: TLabel;
    FSC12: TEdit;
    ToolbarButton971: TToolbarButton97;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure G_TABLE11Exit(Sender: TObject);
    procedure ToolbarButton971Click(Sender: TObject);
    procedure G_TABLE11Change(Sender: TObject);
    procedure FSC11Change(Sender: TObject);
  private
    { Déclarations privées }
    OkValide : Boolean ;
    StParam : String ;
    CodeTableNat,LibTableNat : String ;
    NumC : Integer ;
    procedure DecodeStParam ;
    procedure EncodeStParam ;
    function  Load_Sais ( St : String ) : Variant ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Function ParamCol(Var St : String ; NumC : Integer ; CodeTableNat,LibTableNat : String) : Boolean ;
var FFR: TFParamQR;
begin
Result:=FALSE ;
FFR:=TFParamQR.Create(Application) ;
try
 FFR.OkValide:=FALSE ; FFR.StParam:=St ; FFR.CodeTableNat:=CodeTableNat ;
 FFR.LibTableNat:=LibTableNat ; FFR.NumC:=NumC ;
 FFR.ShowModal ;
finally
 Result:=FFR.OkValide ; If Result Then St:=FFR.StParam ;
 FFR.Free ;
end ;
Screen.Cursor:=SyncrDefault ;
end ;

procedure TFParamQR.BValiderClick(Sender: TObject);
begin
OkValide:=TRUE ;
EncodeStParam ;
end;

procedure TFParamQR.DecodeStParam ;
Var St,St1 : String ;
BEGIN
St:=Trim(StParam) ; If St='' Then Exit ;
If St<>'' Then BEGIN St1:=ReadTokenPipe(St,Pipe) ; TC1.Text:=St1 ; END ;
If St<>'' Then BEGIN St1:=ReadTokenPipe(St,Pipe) ; G_TABLE11.Text:=St1 ; END ;
If St<>'' Then BEGIN St1:=ReadTokenPipe(St,Pipe) ; TC11.Text:=St1 ; END ;
If St<>'' Then BEGIN St1:=ReadTokenPipe(St,Pipe) ; G_TABLE12.Text:=St1 ; END ;
If St<>'' Then BEGIN St1:=ReadTokenPipe(St,Pipe) ; TC12.Text:=St1 ; END ;
If St<>'' Then BEGIN St1:=ReadTokenPipe(St,Pipe) ; CBDEB1.Checked:=St1='X' ; END ;
If St<>'' Then BEGIN St1:=ReadTokenPipe(St,Pipe) ; CBDEB2.Checked:=St1='X' ; END ;
If St<>'' Then BEGIN St1:=ReadTokenPipe(St,Pipe) ; FSC11.Text:=St1 ; END ;
If St<>'' Then BEGIN St1:=ReadTokenPipe(St,Pipe) ; FSC12.Text:=St1 ; END ;
END ;

procedure TFParamQR.EncodeStParam ;
BEGIN
StParam:=TC1.Text+Pipe+G_TABLE11.Text+pipe+TC11.Text+pipe+G_TABLE12.Text+Pipe+TC12.Text+pipe ;
If CBDEB1.Checked Then StParam:=StParam+'X'+Pipe Else StParam:=StParam+'-'+Pipe ;
If CBDEB2.Checked Then StParam:=StParam+'X'+Pipe Else StParam:=StParam+'-'+Pipe ;
StParam:=StParam+FSC11.Text+Pipe+FSC12.Text+Pipe ;
END ;

procedure TFParamQR.FormShow(Sender: TObject);
begin
Caption:=HMess.Mess[NumC-1] ;
TG_TABLE11.Caption:='&'+LibTableNat ; G_TABLE11.ZoomTable:=NatureTotz(CodeTableNat) ;
TG_TABLE12.Caption:='&'+LibTableNat ; G_TABLE12.ZoomTable:=NatureTotz(CodeTableNat) ;
DecodeStParam ;
end;

procedure TFParamQR.G_TABLE11Exit(Sender: TObject);
Var Q : TQuery ;
    St : String ;
    TE : TEdit ;
begin
If thCpteEdit(Sender).Name='G_TABLE11' Then BEGIN St:=G_TABLE11.Text ; TE:=TC11 ; END Else
  If thCpteEdit(Sender).Name='G_TABLE12' Then BEGIN St:=G_TABLE12.Text ; TE:=TC12 ; END Else Exit ;
If TE.Text='' Then
  BEGIN
  Q:=OpenSQL('Select NT_LIBELLE FROM NATCPTE WHERE NT_TYPECPTE="'+CodeTableNat+'" AND NT_NATURE="'+St+'" ',TRUE) ;
  If Not Q.Eof Then TE.Text:=Q.Fields[0].AsString ;
  Ferme(Q) ;
  END ;
end;

function TFParamQR.Load_Sais ( St : String ) : Variant ;
Var V    : Variant ;
BEGIN
V:=#0 ; Result:=V ;
St:=uppercase(Trim(St)) ; if St='' then Exit ;
If St='11' Then V:=11.0 Else
If St='12' Then V:=12.0 Else
If St='21' Then V:=21.0 Else
If St='22' Then V:=22.0 Else Exit ;
Load_Sais:=V ;
END ;


procedure TFParamQR.ToolbarButton971Click(Sender: TObject);
Var St : String ;
    stF : string;
begin
StF := FSC11.Text;
//St:=GFormule(StF,Load_Sais,Nil,1) ;
St:='** '+St+' **' ;
end;

procedure TFParamQR.G_TABLE11Change(Sender: TObject);
begin
If thCpteEdit(Sender).Name='G_TABLE11' Then
  BEGIN
  If thCpteEdit(Sender).Text<>'' Then FSC11.Text:='' ;
  END Else
If thCpteEdit(Sender).Name='G_TABLE12' Then
  BEGIN
  If thCpteEdit(Sender).Text<>'' Then FSC12.Text:='' ;
  END ;
end;

procedure TFParamQR.FSC11Change(Sender: TObject);
begin
If tEdit(Sender).Name='FSC11' Then
  BEGIN
  If tEdit(Sender).Text<>'' Then G_TABLE11.Text:='' ;
  END Else
If tEdit(Sender).Name='FSC12' Then
  BEGIN
  If tEdit(Sender).Text<>'' Then G_TABLE12.Text:='' ;
  END ;
end;

end.
