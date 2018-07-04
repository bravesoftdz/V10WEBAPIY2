unit SousSect;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  Buttons,
{$IFDEF EAGLCLIENT}
  uTob,
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DB,
{$ENDIF}
  Hctrls,
  ExtCtrls,
  hmsgbox,
  HSysMenu,
  ent1,
  HEnt1,
  HCompte,
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  ComCtrls  ,UentCommun
  ;

Function ChoisirSSection(fb : tFichierBase ; Section : String ; AvecJoker : Boolean ; Action : TActionFiche) : String ;

type
  TFChoixSSec = class(TForm)
    PBouton: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    Panel2: TPanel;
    TSC1: TLabel;
    SC1: THCpteEdit;
    LibSC1: THLabel;
    SC2: THCpteEdit;
    TSC2: TLabel;
    LibSC2: THLabel;
    TSC3: TLabel;
    SC3: THCpteEdit;
    LibSC3: THLabel;
    TSC4: TLabel;
    SC4: THCpteEdit;
    LibSC4: THLabel;
    TSC5: TLabel;
    SC5: THCpteEdit;
    LibSC5: THLabel;
    TSC6: TLabel;
    SC6: THCpteEdit;
    LibSC6: THLabel;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BFermeClick(Sender: TObject);
  private
    Action : TActionFiche ;
    AvecJoker : Boolean ;
    fb : tFichierBase ;
    OkSection : Boolean ;
    Section : String ;
    SSection : Array[1..6] Of String ;
    NbVisible : Integer ;
    Procedure Invisible(i : Integer ; T : TLabel ; Lib : THLabel ; SC : THCpteEdit) ;
    public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}


Function ChoisirSSection(fb : tFichierBase ; Section : String ; AvecJoker : Boolean ; Action : TActionFiche) : String ;
var
  FChoixSSec: TFChoixSSec ;
BEGIN
Result:=Section ;
FChoixSSec:=TFChoixSSec.Create(Application) ;
try
  FChoixSSec.fb:=fb ;
  FChoixSSec.Section:=Section ;
  FChoixSSec.OkSection:=TRUE ;
  FChoixSSec.AvecJoker:=AvecJoker ;
  FChoixSSec.Action:=Action ;
  FChoixSSec.ShowModal ;
  finally
  If Action in [taModif,taCreat] Then
     BEGIN
     if FChoixSSec.OkSection then Result:=FChoixSSec.Section Else Result:='' ;
     END ;
  FChoixSSec.free ;
  end ;
//SourisNormale ;
Screen.Cursor:=SyncrDefault ;
END ;

Procedure TFChoixSSec.Invisible(i : Integer ; T : TLabel ; Lib : THLabel ; SC : THCpteEdit) ;
BEGIN
If VH^.SousPlanAxe[fb,i].Code<>'' Then Inc(NbVisible) Else
   BEGIN
   T.Visible:=FALSE ; Lib.Visible:=FALSE ; SC.Visible:=FALSE ;
   END ;
END ;

procedure TFChoixSSec.FormShow(Sender: TObject);
Var i : Integer ;
    HCpt : TComponent ;
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
SC1.ZoomTable:=TZoomTable(Ord(tzAxe1SP1)+((Ord(fb)-Ord(fbAxe1))*6)) ;
SC2.ZoomTable:=TZoomTable(Ord(tzAxe1SP2)+((Ord(fb)-Ord(fbAxe1))*6)) ;
SC3.ZoomTable:=TZoomTable(Ord(tzAxe1SP3)+((Ord(fb)-Ord(fbAxe1))*6)) ;
SC4.ZoomTable:=TZoomTable(Ord(tzAxe1SP4)+((Ord(fb)-Ord(fbAxe1))*6)) ;
SC5.ZoomTable:=TZoomTable(Ord(tzAxe1SP5)+((Ord(fb)-Ord(fbAxe1))*6)) ;
SC6.ZoomTable:=TZoomTable(Ord(tzAxe1SP6)+((Ord(fb)-Ord(fbAxe1))*6)) ;
TSC1.Caption:='&'+VH^.SousPlanAxe[fb,1].Lib ; TSC2.Caption:='&'+VH^.SousPlanAxe[fb,2].Lib ;
TSC3.Caption:='&'+VH^.SousPlanAxe[fb,3].Lib ; TSC4.Caption:='&'+VH^.SousPlanAxe[fb,4].Lib ;
TSC5.Caption:='&'+VH^.SousPlanAxe[fb,5].Lib ; TSC6.Caption:='&'+VH^.SousPlanAxe[fb,6].Lib ;
LibSC1.Caption:='' ; LibSC2.Caption:='' ; LibSC3.Caption:='' ;
LibSC4.Caption:='' ; LibSC5.Caption:='' ; LibSC6.Caption:='' ;
NbVisible:=0 ;
Invisible(1,TSC1,LibSC1,SC1) ; Invisible(2,TSC2,LibSC2,SC2) ; Invisible(3,TSC3,LibSC3,SC3) ;
Invisible(4,TSC4,LibSC4,SC4) ; Invisible(5,TSC5,LibSC5,SC5) ; Invisible(6,TSC6,LibSC6,SC6) ;
If NbVisible>4 Then Self.Height:=160 Else If NbVisible>2 Then Self.Height:=130 Else Self.Height:=100 ;
Fillchar(SSection,SizeOf(SSection),#0) ;
For i:=1 To NbVisible Do
  BEGIN
  SSEction[i]:=Copy(Section,VH^.SousPlanAxe[fb,i].Debut,VH^.SousPlanAxe[fb,i].Longueur) ;
  HCpt:=Self.FindComponent('SC'+IntToStr(i)) ;
  If HCpt<>NIL Then
     BEGIN
     THCpteEdit(HCpt).Text:=SSEction[i] ;
     THCpteEdit(HCpt).ExisteH ;
     If Action=TaConsult Then THCpteEdit(HCpt).Enabled:=FALSE ;
     END ;
  END ;
If NbVisible>0 Then SC1.SetFocus ;
end;


procedure TFChoixSSec.BValiderClick(Sender: TObject);
Var i,j : Integer ;
    HCpt : TComponent ;
    C : Char ;
    Q : TQuery ;
    OkOk,OkFerme : Boolean ;
begin
OkFerme:=TRUE ;
Section:='' ;
If Action=taConsult Then Exit ;
For i:=1 To NbVisible Do
   BEGIN
   HCpt:=Self.FindComponent('SC'+IntToStr(i)) ;
   If HCpt<>NIL Then
      BEGIN
      SSEction[i]:=THCpteEdit(HCpt).Text ;
      If Trim(SSection[i])='' Then
         BEGIN
         SSection[i]:='' ;
         If AvecJoker Then C:='?' Else C:=' ' ;
         For j:=1 To VH^.SousPlanAxe[fb,i].Longueur Do SSection[i]:=SSection[i]+C ;
         END ;
      Section:=Section+SSection[i] ;
      END ;
   END ;
If Action in [tacreat,taModif] Then
   BEGIN
   Q:=OpenSQL('SELECT * FROM SECTION WHERE S_AXE="'+fbToAxe(fb)+'" AND S_SECTION="'+Section+'" ',TRUE) ;
   OkOK:=Not Q.Eof ;
   Ferme(Q) ;
   If Not OkOk Then
      BEGIN
      Case MsgBox.execute(0,'','') of
        mrYes : BEGIN
                If Assigned(ProcZoomSection) Then
                  ProcZoomSection(Nil,fbToAxe(fb),Section,taCreatOne,0) ;
                END ;
        mrNo  : OkFerme:=FALSE ;
        mrCancel : OkFerme:=FALSE ;
        end ;
      END ;
   END ;
If OkFerme Then Close ;
end;

procedure TFChoixSSec.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : boolean ;
begin
Vide:=(Shift=[]) ;
if Vide then
   BEGIN
   Case Key of
        VK_F10 : BEGIN Key:=0 ; BValiderClick(Nil) ; END ;
        VK_RETURN :BEGIN
                   Key:=0 ;              
                   FindNextControl(ActiveControl,True,True,False).SetFocus ;
                   END ;
     END ;
   END ;
end;

procedure TFChoixSSec.BFermeClick(Sender: TObject);
begin
OkSection:=FALSE ;
end;

end.
