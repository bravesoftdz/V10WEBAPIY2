{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 31/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit RevBuRea;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, ExtCtrls, Grids, Hctrls, StdCtrls, Buttons, ComCtrls,
  Ent1, Hent1,
{$IFDEF EAGLCLIENT}
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  HStatus, Menus, HTB97, HPanel, UiUtil, ImgList, UTob,
  UObjFiltres, {JP 30/06/06}
  UtilPGI, HImgList ;

Procedure RevisionBudgetaireRealiser ;

Type TInfoJal = Class
      ExoDeb  : String ;
      ExoFin  : String ;
      Ax      : String ;
      Gatt    : String ;
      Satt    : String ;
      CptGen  : String ;
      CptGen2 : String ;
      CptSec  : String ;
      CptSec2 : String ;
      Souche  : String ;
      PerDeb  : TDateTime ;
      PerFin  : TDateTime ;
      NatJal  : String ;
      Categor : String ;
     End ;

Type TInfoEcr = Class
       Montant : Double ;
       End ;
type
  TFRevBuRea = class(TForm)
    Pages: TPageControl;
    PParamS: TTabSheet;
    Bevel1: TBevel;
    FListe: THGrid;
    HMTrad: THSystemMenu;
    HMG: THMsgBox;
    PParamD: TTabSheet;
    Bevel2: TBevel;
    TCbBud: TLabel;
    BudS: THValComboBox;
    TCbEtab: TLabel;
    EtabS: THValComboBox;
    NatS: THValComboBox;
    TCbNatS: TLabel;
    TPerdebS: TLabel;
    PerDebS: THValComboBox;
    TPerFinS: TLabel;
    PerFinS: THValComboBox;
    CbPS: THValComboBox;
    TBudD: TLabel;
    BudD: THValComboBox;
    TNatDSup: TLabel;
    NatDSup: THValComboBox;
    TEtabD: TLabel;
    EtabD: THValComboBox;
    PerFinD: THValComboBox;
    TPerFinD: TLabel;
    PerDebD: THValComboBox;
    TPerDebD: TLabel;
    CbPD: THValComboBox;
    HM: THMsgBox;
    NatDInf: THValComboBox;
    TNatDInf: TLabel;
    Image: THImageList;
    RgSens: TRadioGroup;
    PopZ: TPopupMenu;
    POPS: TPopupMenu;
    BCpy: THBitBtn;
    BcpyLig: THBitBtn;
    BcolLig: THBitBtn;
    Pcrit: TTabSheet;
    TCbCouple1: TLabel;
    CbCouple1: THValComboBox;
    CbCouple2: THValComboBox;
    TCbCouple2: TLabel;
    RgTyp: TRadioGroup;
    Bevel4: TBevel;
    HLabel4: THLabel;
    Resol: THValComboBox;
    NbE: TLabel;
    Bevel3: TBevel;
    NbEcr: TLabel;
    PFiltres: TToolWindow97;
    FFiltres: THValComboBox;
    BCherche: TToolbarButton97;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    Dock: TDock97;
    Dock971: TDock97;
    HPB: TToolWindow97;
    BAnnuler: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    BBudDown: THBitBtn;
    BBudUp: THBitBtn;
    procedure BAideClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BudSChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure NatSChange(Sender: TObject);
    procedure PerDebSChange(Sender: TObject);
    procedure PerFinSChange(Sender: TObject);
    procedure PerDebDChange(Sender: TObject);
    procedure PerFinDChange(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure BValiderClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure FListeCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure BAnnulerClick(Sender: TObject);
    procedure BBudUpClick(Sender: TObject);
    procedure BBudDownClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure BCpyClick(Sender: TObject);
    procedure BcpyLigClick(Sender: TObject);
    procedure BcolLigClick(Sender: TObject);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure RgTypClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    ObjFiltre : TObjFiltre; {JP 30/06/06}

    WMinX,WMinY    : Integer ;
    OldRow,LeSens : Integer ;
    NumPieceSup,NumPieceInf : LongInt ;
    CollerLigne : HTStrings ;

    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Procedure PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Function  ChargeInfoJalBud(StCod : String) : TInfoJal ;
    Procedure GetInfoJal(Sender : TObject) ;
    Procedure DetruitObjCombo(Sender : TObject) ;
    Procedure ChageInfoPeriode(Pref : String ; Obj : TInfoJal) ;
    Procedure InitLaFListe ;
    Procedure ChargeLesPeriodes(Sender : TObject) ;
    Function  ChercheBudgete(StJal,StRub : String ; D1,D2 : TDateTime ; Exo,Etab,NatB : String ; Var Cbg,Cbs,Laval : String) : Double ;
    Function  ControleCompteOk : Boolean ;
    Function  ControleCpteAttOk : Boolean ;
    Function  ControlePeriodeOk : Boolean ;
    Function  ControleCategorOk : Boolean ;
    Procedure RunLaCreationAuProrata ;
    Procedure QuelCompteBudget(Var BudG,BudS : String ; Lig : Integer) ;
    Function  ChercheNumPiece : LongInt ;
    Function  ControleNatureOk : Boolean ;
    Function  FormateResultat(D,C : Double ; Var Laval : String) : Double ;
    Function  ChercheDebCred(Var Ecar : Double ; ARow,ACol : Integer) : Byte ;
    Procedure EnvoiVisuPiece(Nat : String ; NumP : LongInt) ;
    Procedure EcritLenreg(Exo,Nat,Bg,Bs : String ; Dcpta,DModif : TDateTime ;
                          D,C : Double ; NPiece : LongInt) ;
    Procedure ChargeCroisement(StJal,StCat : String) ;
    Function  ChercheCompte(CRubG,CRubS : String ; Var CbG,CbS : String) : Boolean ;
    procedure GlyphEtCouleurs(NbGlyphs : Integer ; I : TButtonState97 ; FontColor : TColor ; BmpSource : TBitMap ; var BitMap : TBitMap) ;
  public
    { Déclarations publiques }

    {JP 28/06/06 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc}
    procedure GereEtablissement;
    {JP 28/06/06 : FQ 16149 : on s'assure que le filtre coincide avec les restrictions utilisateurs sur l'établissement}
    procedure ControlEtab;
    procedure OnApresChangementFiltre;
  end;


implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  Calcole, CpteUtil, SaisUtil,
  eSaisBud,
{$IFDEF EAGLCLIENT}
  UtileAGL,
{$ELSE}
  PrintDBG,
{$ENDIF}
  UtilEdt;

Procedure RevisionBudgetaireRealiser ;
var FRevBuRea : TFRevBuRea ;
    PP : THPanel ;
BEGIN
FRevBuRea:=TFRevBuRea.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FRevBuRea.ShowModal ;
    Finally
     FRevBuRea.Free;
    End ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FRevBuRea,PP) ;
   FRevBuRea.Show ;
   END ;
END ;

procedure TFRevBuRea.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFRevBuRea.WMGetMinMaxInfo(var MSG: Tmessage);
BEGIN with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end; END ;

procedure TFRevBuRea.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
  {JP 30/06/06 : Mise à jour de la gestion des filtres}
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, 'REVBUDREA');
  ObjFiltre.ApresChangementFiltre := OnApresChangementFiltre; 

  WMinX:=Width ;
  WMinY:=Height ;
  CollerLigne:=HTStringList.Create ;
end;

Procedure TFRevBuRea.DetruitObjCombo(Sender : TObject) ;
Var i : Integer ;
BEGIN
for i:=0 to THValComboBox(Sender).Values.Count-1 do
    if THValComboBox(Sender).Values.Objects[i]<>Nil then
       TObject(THValComboBox(Sender).Values.Objects[i]).Free ;
END ;

procedure TFRevBuRea.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  {JP 30/06/06 : Mise à jour de la gestion des filtres}
  ObjFiltre.Free;

FListe.VidePile(True) ; DetruitObjCombo(BudS) ; DetruitObjCombo(BudD) ;
PurgePopup(POPS) ; PurgePopup(POPZ) ; CollerLigne.Free ;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFRevBuRea.PerDebSChange(Sender: TObject);
begin
if(PerDebS.Value='') Or (PerFinS.Value='') then Exit ;
if StrToDate(PerDebS.Value)>StrToDate(PerFinS.Value) then
   PerFinS.Value:=DateToStr(FinDeMois(StrToDate(PerDebS.Value))) ;
end;

procedure TFRevBuRea.PerFinSChange(Sender: TObject);
begin
if(PerFinS.Value='') Or (PerDebS.Value='') then Exit ;
if StrToDate(PerFinS.Value)<StrTodate(PerDebS.Value) then
   PerDebS.Value:=DateToStr(DebutDeMois(StrToDate(PerDebS.Value))) ;
end;

procedure TFRevBuRea.PerDebDChange(Sender: TObject);
begin
if(PerDebD.Value='') Or (PerFinD.Value='') then Exit ;
if StrToDate(PerDebD.Value)>StrToDate(PerFinD.Value) then
   PerFinD.Value:=DateToStr(FinDeMois(StrToDate(PerDebD.Value))) ;
end;

procedure TFRevBuRea.PerFinDChange(Sender: TObject);
begin
if(PerFinD.Value='') Or (PerDebD.Value='') then Exit ;
if StrToDate(PerFinD.Value)<StrTodate(PerDebD.Value) then
   PerDebD.Value:=DateToStr(DebutDeMois(StrToDate(PerDebD.Value))) ;
end;

procedure TFRevBuRea.FormShow(Sender: TObject);
begin
  {JP 30/06/06 : Mise à jour de la gestion des filtres}
  ObjFiltre.Charger;

  Pages.ActivePage:=Pages.Pages[0] ; BMenuZoom.Enabled:=False ;
  Resol.Value:='C' ;
  FListe.GetCellCanvas:=GetCellCanvas ; FListe.PostDrawCell:=PostDrawCell ;

  {JP 30/06/06 : La tablette est TTETABLISSEMENT !!!!
  EtabS.Value:=V_PGI.CodeSociete ; EtabD.Value:=V_PGI.CodeSociete ;}
  GereEtablissement;

// FQ 11644
If (BudS.Values.Count<=0) Or (BudD.Values.Count<=0) Then BEGIN HM.Execute(11,'','') ; Exit; END ;
if ((Not V_PGI.OutLook) and (BudS.Values[0]='')) then BEGIN HM.Execute(11,'','') ; Exit ; END ;
if BudS.Values[0]<>'' then BudS.Value:=BudS.Values[0] ;
if BudD.Values[0]<>'' then BudD.Value:=BudD.Values[0] ;
NatS.Value:='INI' ; NatDSup.Value:='DM1' ; NatDInf.Value:='DM2' ;

InitLaFListe ;
end;

Procedure TFRevBuRea.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if (ACol=0) And (ARow<>0) And (Pos(':',FListe.Cells[0,ARow])>0) then
   FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style+[fsBold]  else
   FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style-[fsBold] ;
END ;

procedure TFRevBuRea.GlyphEtCouleurs(NbGlyphs : Integer ; I : TButtonState97 ; FontColor : TColor ; BmpSource : TBitMap ; var BitMap : TBitMap) ;
const
  ROP_DSPDxax = $00E20746;
var MonoBMP        : TBitMap ;
    DestDC         : HDC ;
    IWidth,IHeight : integer ;
    IRect,ORect    : TRect ;
BEGIN
IWidth := BmpSource.Width div NbGlyphs ; IHeight := BmpSource.Height;
BitMap:=TBitmap.Create ;
Bitmap.Width := IWidth ; Bitmap.Height := IHeight;
IRect:=Rect(0, 0, IWidth, IHeight) ;
ORect:=Rect(Ord(I) * IWidth, 0, (Ord(I) + 1) * IWidth, IHeight) ;
Case I of
  bsup      :BEGIN
             BitMap.Canvas.Brush.Color := FontColor ;
             BitMap.Canvas.BrushCopy(IRect, BmpSource, ORect, BmpSource.TransparentColor) ;
             END ;
  bsDisabled:BEGIN
             MonoBmp := TBitmap.Create;
             With BitMap.Canvas do
                 BEGIN
                 { Change les couleurs blanches et grises au couleurs clBtnHighlight et clBtnShadow }
                 CopyRect(IRect, BmpSource.Canvas, ORect) ;
                 MonoBmp.Width := IWidth ;
                 MonoBmp.Height := IHeight ;
                 MonoBmp.Monochrome := True ;

                 { Blanc -> clBtnHighlight }
                 BmpSource.Canvas.Brush.Color := clWhite ;
                 MonoBmp.Canvas.CopyRect(IRect, BmpSource.Canvas, ORect) ;
                 Brush.Color := clBtnHighlight ;
                 DestDC := Handle ;
                 SetTextColor(DestDC, clBlack) ;
                 SetBkColor(DestDC, clWhite) ;
                 BitBlt(DestDC, 0, 0, IWidth, IHeight,
                        MonoBmp.Canvas.Handle, 0, 0, ROP_DSPDxax) ;

                 { gris -> clBtnShadow }
                 BmpSource.Canvas.Brush.Color := clGray ;
                 MonoBmp.Canvas.CopyRect(IRect, BmpSource.Canvas, ORect) ;
                 Brush.Color := clBtnShadow ;
                 DestDC := Handle ;
                 SetTextColor(DestDC, clBlack) ;
                 SetBkColor(DestDC, clWhite) ;
                 BitBlt(DestDC, 0, 0, IWidth, IHeight,
                        MonoBmp.Canvas.Handle, 0, 0, ROP_DSPDxax);
                 { Transparent -> FontColor }
                 BmpSource.Canvas.Brush.Color := ColorToRGB(BmpSource.TransparentColor);
                 MonoBmp.Canvas.CopyRect(IRect, BmpSource.Canvas, ORect);
                 Brush.Color := FontColor ;
                 DestDC := Handle ;
                 SetTextColor(DestDC, clBlack) ;
                 SetBkColor(DestDC, clWhite) ;
                 BitBlt(DestDC, 0, 0, IWidth, IHeight,
                 MonoBmp.Canvas.Handle, 0, 0, ROP_DSPDxax) ;
                 MonoBmp.Free ;
                 END ;
             END ;
  END ;
END ;

Procedure TFRevBuRea.PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
Var R : TRect ;
    BitMap : TBitMap ;
BEGIN
if (Pos(':',FListe.Cells[0,ARow])>0) And (ARow<>0) And (ACol<>0) then
    BEGIN
    Canvas.Brush.Color := Fliste.FixedColor ; ;
    Canvas.Brush.Style := bsBDiagonal ;
    Canvas.Pen.Color   := Fliste.FixedColor ;
    Canvas.Pen.Mode    := pmCopy ;
    Canvas.Pen.Style   := psClear ;
    Canvas.Pen.Width   := 1 ;
    R:=FListe.CellRect(ACol,ARow) ;
    Canvas.Rectangle(R.Left,R.Top,R.Right+1,R.Bottom+1) ;
    END ;
if (Pos(':',FListe.Cells[0,ARow])>0) And (ACol=0) then
    BEGIN
    R:=FListe.CellRect(ACol,ARow) ;
    R.Left:=R.Right-20 ; R.Top:=R.Top+1 ; R.Right:=R.Left+16 ;R.Bottom:=R.Top+16 ;
    BitMap:=TBitMap.Create ;
    if FListe.RowHeights[ARow+1]=18 then Image.GetBitMap(0,BitMap)
                                    else Image.GetBitMap(1,BitMap) ;
    GlyphEtCouleurs(1,bsup,FListe.FixedColor,BitMap,BitMap) ;
    FListe.Canvas.StretchDraw(R,BitMap) ;
    BitMap.Free ;
    END ;
END ;

procedure TFRevBuRea.FListeRowEnter(Sender: TObject; Ou: Longint;
  var Cancel: Boolean; Chg: Boolean);
begin
if Pos(HMG.Mess[2],FListe.Cells[0,Ou])>0 then FListe.Options:=FListe.Options+[goEditing]
                                         else FListe.Options:=FListe.Options-[goEditing] ;
if FListe.RowHeights[Ou]=0 then
  BEGIN
  if OldRow>FListe.Row then FListe.Row:=FListe.Row-1
  else BEGIN
       if FListe.Row+1>FListe.RowCount-1 then FListe.Row:=OldRow
                                         else FListe.Row:=FListe.Row+1 ;
       END ;
  END else OldRow:=FListe.Row ;
end;

Procedure TFRevBuRea.ChargeCroisement(StJal,StCat : String) ;
Var QLoc : TQuery ;
    Sql : String ;
BEGIN
CbCouple1.Values.Clear ; CbCouple1.Items.Clear ;
CbCouple2.Values.Clear ; CbCouple2.Items.Clear ;
if StCat<>'' then StJal:=StCat ;
Case RgTyp.ItemIndex of
     0 : Sql:='Select CX_COMPTE,CX_SECTION,BG_RUB,BS_RUB From CROISCPT '+
              'Left join BUDGENE on CX_COMPTE=BG_BUDGENE '+
              'Left join BUDSECT on CX_SECTION=BS_BUDSECT '+
              'Where CX_TYPE="BUD" And CX_JAL="'+StJal+'"' ;
     1 : Sql:='Select DISTINCT CX_COMPTE,BG_RUB From CROISCPT '+
              'Left join BUDGENE on CX_COMPTE=BG_BUDGENE '+
              'Where CX_TYPE="BUD" And CX_JAL="'+StJal+'"' ;
  End ;
QLoc:=OpenSql(Sql,True) ;
While Not QLoc.Eof do
  BEGIN
  Case RgTyp.ItemIndex of
       0 : BEGIN
           CbCouple1.Items.Add(QLoc.Fields[0].AsString+' / '+QLoc.Fields[1].AsString) ;
           CbCouple1.Values.Add(QLoc.Fields[2].AsString+' / '+QLoc.Fields[3].AsString) ;
           END ;
       1 : BEGIN
           CbCouple1.Items.Add(QLoc.Fields[0].AsString) ;
           CbCouple1.Values.Add(QLoc.Fields[1].AsString) ;
           END ;
      End ;
  QLoc.Next ;
  END ;
Ferme(QLoc) ;
CbCouple2.Values.Assign(CbCouple1.Values) ;
CbCouple2.Items.Assign(CbCouple1.Items) ;
if CbCouple1.Items.Count>0 then
   BEGIN
   CbCouple1.ItemIndex:=0 ; CbCouple2.ItemIndex:=CbCouple2.Items.Count-1 ;
   END ;
END ;

Function TFRevBuRea.ChargeInfoJalBud(StCod : String) : TInfoJal ;
Var QLoc : TQuery ;
    X : TInfoJal ;
    i : Integer ;
BEGIN
Result:=Nil ;
if StCod='' then Exit ;
X:=TInfoJal.Create ;
QLoc:=OpenSql('Select * From BUDJAL Where BJ_BUDJAL="'+StCod+'"',True) ;
for i:=0 to QLoc.FieldCount-1 do
   BEGIN
   if QLoc.Fields[i].FieldName='BJ_EXODEB' then X.ExoDeb:=QLoc.Fields[i].AsString else
   if QLoc.Fields[i].FieldName='BJ_EXOFIN' then X.ExoFin:=QLoc.Fields[i].AsString else
   if QLoc.Fields[i].FieldName='BJ_AXE' then X.Ax:=QLoc.Fields[i].AsString else
   if QLoc.Fields[i].FieldName='BJ_GENEATTENTE' then X.Gatt:=QLoc.Fields[i].AsString else
   if QLoc.Fields[i].FieldName='BJ_SECTATTENTE' then X.Satt:=QLoc.Fields[i].AsString else
   if QLoc.Fields[i].FieldName='BJ_BUDGENES' then X.CptGen:=QLoc.Fields[i].AsString else
   if QLoc.Fields[i].FieldName='BJ_BUDSECTS' then X.CptSec:=QLoc.Fields[i].AsString else
   if QLoc.Fields[i].FieldName='BJ_BUDGENES2' then X.CptGen2:=QLoc.Fields[i].AsString else
   if QLoc.Fields[i].FieldName='BJ_BUDSECTS2' then X.CptSec2:=QLoc.Fields[i].AsString else
   if QLoc.Fields[i].FieldName='BJ_COMPTEURNORMAL' then X.Souche:=QLoc.Fields[i].AsString else
   if QLoc.Fields[i].FieldName='BJ_PERDEB' then X.PerDeb:=QLoc.Fields[i].AsDateTime else
   if QLoc.Fields[i].FieldName='BJ_PERFIN' then X.PerFin:=QLoc.Fields[i].AsDateTime else
   if QLoc.Fields[i].FieldName='BJ_NATJAL' then X.NatJal:=QLoc.Fields[i].AsString else
   if QLoc.Fields[i].FieldName='BJ_CATEGORIE' then X.Categor:=QLoc.Fields[i].AsString ;
   END ;
Ferme(QLoc) ; Result:=X ;
END ;

Procedure TFRevBuRea.GetInfoJal(Sender : TObject) ;
Var LaVal,Pref : String ;
    UnCb : THvalComboBox ;
BEGIN
UnCb:=THValComboBox(Sender) ;
If UnCb.Values.Count<=0 Then Exit ;
if UnCb=Nil then Exit ;
if UnCB.ItemIndex=-1 then Exit ; 
LaVal:=UnCb.Value ; Pref:=Copy(UnCb.Name,Length(UnCb.Name),1) ;
if UnCb.Values.Objects[UnCb.Values.IndexOf(LaVal)]=Nil then
   UnCb.Values.Objects[UnCb.Values.IndexOf(LaVal)]:=ChargeInfoJalBud(LaVal) ;
if UnCb.Values.Objects[UnCb.Values.IndexOf(LaVal)]<>Nil then
   BEGIN
   ChageInfoPeriode(Pref,TInfoJal(UnCb.Values.Objects[UnCb.Values.IndexOf(LaVal)])) ;
   if Pref='S' then NatSChange(Nil) ;
   END ;
if Pref='S' then
   BEGIN
   if TInfoJal(UnCb.Values.Objects[UnCb.Values.IndexOf(LaVal)]).NatJal='CHA' then RgSens.ItemIndex:=0
                                                                             else RgSens.ItemIndex:=1 ;
   ChargeCroisement(UnCb.Value,TInfoJal(UnCb.Values.Objects[UnCb.Values.IndexOf(LaVal)]).Categor) ;
   END ;
END ;

Procedure TFRevBuRea.ChageInfoPeriode(Pref : String ; Obj : TInfoJal) ;
Var DDeb,DFin,DTemp : TDateTime ;
    LDat : String ;
BEGIN
DDeb:=Obj.PerDeb ; DFin:=Obj.PerFin ;
THValComboBox(FindComponent('PerDeb'+Pref)).Values.Clear ; THValComboBox(FindComponent('PerDeb'+Pref)).Items.Clear ;
THValComboBox(FindComponent('PerFin'+Pref)).Values.Clear ; THValComboBox(FindComponent('PerFin'+Pref)).Items.Clear ;
Repeat
 DTemp:=FindeMois(DDeb) ;
 LDat:=FormatDateTime('mmmm yyyy',DDeb) ; LDat:=FirstMajuscule(LDat) ;
 THValComboBox(FindComponent('PerDeb'+Pref)).Items.Add(LDat) ;
 THValComboBox(FindComponent('PerDeb'+Pref)).Values.Add(DateToStr(DDeb)) ;
 LDat:=FormatDateTime('mmmm yyyy',DTemp) ; LDat:=FirstMajuscule(LDat) ;
 THValComboBox(FindComponent('PerFin'+Pref)).Items.Add(LDat) ;
 THValComboBox(FindComponent('PerFin'+Pref)).Values.Add(DateToStr(DTemp)) ;
 DDeb:=PlusMois(DDeb,1) ;
Until DTemp>=DFin ;
if THValComboBox(FindComponent('PerDeb'+Pref)).Values.Count>0 then
   THValComboBox(FindComponent('PerDeb'+Pref)).Value:=THValComboBox(FindComponent('PerDeb'+Pref)).Values[0] ;
if THValComboBox(FindComponent('PerFin'+Pref)).Values.Count>0 then
   THValComboBox(FindComponent('PerFin'+Pref)).Value:=THValComboBox(FindComponent('PerFin'+Pref)).Values[THValComboBox(FindComponent('PerFin'+Pref)).Values.Count-1] ;
END ;

procedure TFRevBuRea.BudSChange(Sender: TObject);
begin GetInfoJal(Sender) ; end;

Procedure TFRevBuRea.InitLaFListe ;
Var i : Integer ;
BEGIN
FListe.VidePile(True) ;
FListe.ColCount:=PerFinS.ItemIndex-PerDebS.ItemIndex+2 ;
for i:=PerDebS.ItemIndex to PerFinS.ItemIndex do
    BEGIN
    FListe.Cells[(i-PerDebS.ItemIndex)+1,0]:=PerDebS.Items[i] ;
    FListe.ColWidths[(i-PerDebS.ItemIndex)+1]:=100 ;
    FListe.ColAligns[(i-PerDebS.ItemIndex)+1]:=taRightJustify ;
    END ;
END ;

procedure TFRevBuRea.NatSChange(Sender: TObject);
Var QLoc : TQuery ;
    i : Integer ;
    Sql : String ;
begin
Sql:='Select Distinct BE_NUMEROPIECE From BUDECR Where BE_BUDJAL="'+BudS.Value+'" ' ;
if NatS.Value<>'' then Sql:=Sql+'And BE_NATUREBUD="'+NatS.Value+'"';
QLoc:=OpenSql(Sql,True) ;
i:=0 ;
While Not QLoc.Eof do BEGIN Inc(i) ; QLoc.Next ; END ;
Ferme(QLoc) ;
NbEcr.Caption:=IntToStr(i) ;
if i>1 then NbE.Caption:=HMG.Mess[4] else NbE.Caption:=HMG.Mess[3] ;
end;

Procedure TFRevBuRea.ChargeLesPeriodes(Sender : TObject) ;
Var i,j,k : Integer ;
    Pref : String ;
BEGIN
Pref:=Copy(THValComboBox(Sender).Name,Length(THValComboBox(Sender).Name),1) ;
THValComboBox(Sender).Values.Clear ; THValComboBox(Sender).Items.Clear ;
j:=THValComboBox(FindComponent('PerDeb'+Pref)).ItemIndex ;
k:=THValComboBox(FindComponent('PerFin'+Pref)).ItemIndex ;
for i:=j to k do
   BEGIN
   THValComboBox(Sender).Values.Add(THValComboBox(FindComponent('PerDeb'+Pref)).Values[i]) ;
   THValComboBox(Sender).Items.Add(THValComboBox(FindComponent('PerDeb'+Pref)).Items[i]) ;
   END ;
END ;

Function OnPeutSurCumul(Exo : String ; D1,D2 : TDateTime) : Boolean ;
Var DateBut : TDateTime ;
BEGIN
Result:=False ;
if Exo=VH^.Encours.Code then DateBut:=VH^.EnCours.DateButoirBudgete else
   if Exo=VH^.Suivant.Code then DateBut:=VH^.Suivant.DateButoirBudgete else Exit ;
if D1>DateBut then Exit ;
if DebutDeMois(D1)<>D1 then Exit ;
if D2>DateBut then Exit ;
if FindeMois(D2)<>D2 then Exit ;
Result:=True ;
END ;

Function TFRevBuRea.ChercheCompte(CRubG,CRubS : String ; Var CbG,CbS : String) : Boolean ;
Var St,StC : String ;
    i : Integer ;
BEGIN
Result:=False ; CbG:='' ; CbS:='' ; St:=CRubG+' / '+CRubS ;
i:=CbCouple1.Values.IndexOf(St) ;
if i<0 then Exit ;
if (i<CbCouple1.ItemIndex) or (i>CbCouple2.ItemIndex) then Exit ;
StC:=CbCouple1.Items[i] ;
CbG:=Copy(StC,1,Pos(' / ',StC)-1) ; CbS:=Copy(StC,Pos(' / ',StC)+3,Length(StC)) ;
Result:=True ;
END ;

Function TFRevBuRea.ChercheBudgete(StJal,StRub : String ; D1,D2 : TDateTime ; Exo,Etab,NatB : String ; Var Cbg,Cbs,Laval : String) : Double ;
Var CptRubG,CptRubS : String ;
    i : Integer ;
    D,C : Double ;
    QBud : TQuery ;
    lStSQL : String ;
BEGIN
Laval:='' ; CbG:='' ; CbS:='' ; CptRubG:='' ; CptRubS:='' ; D:=0 ; C:=0 ;
Case RgTyp.ItemIndex of
     0 : BEGIN
         CptRubG:=Copy(StRub,4+Length(StJal),(Pos(':',StRub))-(4+Length(StJal))) ;
         CptRubS:=Copy(StRub,Pos(':',StRub)+1,Length(StRub)-Pos(':',StRub)) ;
         if Not ChercheCompte(CptRubG,CptRubS,CbG,CbS) then BEGIN Result:=0 ; Exit ; END ;
         END ;
     1 : BEGIN
         CptRubG:=Copy(StRub,4+Length(StJal),Length(StRub)) ;
         i:=CbCouple1.Values.IndexOf(CptRubG) ;
         if (i<0) or (i<CbCouple1.ItemIndex) or (i>CbCouple2.ItemIndex) then BEGIN Result:=0 ; Exit ; END ;
         CbG:=CbCouple1.Items[i] ; CbS:=TInfoJal(BudS.Values.Objects[BudS.ItemIndex]).Satt ;
         END ;
   End ;

  // Remplacement des requêtes paramétrées
  lStSQL := 'SELECT SUM(BE_DEBIT) DEB, SUM(BE_CREDIT) CRED FROM BUDECR '
              + 'WHERE BE_BUDJAL="'         + StJal
              + '" AND BE_EXERCICE="'       + Exo
              + '" AND BE_DATECOMPTABLE>="' + UsDateTime(D1)
              + '" AND BE_DATECOMPTABLE<="' + UsDateTime(D2)
              + '" AND BE_QUALIFPIECE="'    + 'N'
              + '" AND BE_ETABLISSEMENT="'  + Etab
              + '" AND BE_BUDGENE="'        + CbG
              + '" AND BE_BUDSECT="'        + CbS + '"' ;
  if NatB<>'' then
     lstSql := lStSql + ' And BE_NATUREBUD="'      + NatB + '"' ;
  QBud := OpenSQL( lStSQL , True ) ;
  if not QBud.Eof then
    begin
    D := QBud.FindField('DEB').AsFloat ;
    C := QBud.FindField('CRED').AsFloat ;
    end ;
  Ferme(QBud) ;
  // Fin remplacement des requêtes paramétrées

  Result:=FormateResultat(D,C,Laval) ;
END ;

Function TFRevBuRea.FormateResultat(D,C : Double ; Var Laval : String) : Double ;
Var DebitPos : Boolean ;
    Decim : Integer ;
BEGIN
DebitPos:=True ; Result:=0 ;
Case LeSens of
    0 : BEGIN Result:=D-C ; DebitPos:=True ;  END ;
    1 : BEGIN Result:=C-D ; DebitPos:=False ; END ;
   End ;
Reevaluation(D,C,Resol.Value,0) ;
if Resol.Value='C' then Decim:=V_PGI.OkDecV else Decim:=0 ;
Laval:=Trim(PrintEcart(D,C,Decim,DebitPos)) ;
if Laval='' then Laval:='0' ;
END ;

procedure TFRevBuRea.BChercheClick(Sender: TObject);
Var QLoc : TQuery ;
    Etab,Dev,BNat,RubAno,Exo,LavalBud,LavalRea,LavalEca : String ;
    D1,D2 : TDateTime ;
    TResult : TabloExt ;
    St : String ;
    MontantRea,MontantBud,Ecart : Double ;
    i : Integer ;
    Cptbg,Cptbs,SurQuoi : String ;
    X : TInfoEcr ;
    lInNbEnreg : Integer ;
begin
InitLaFListe ; BMenuZoom.Enabled:=False ;
LeSens:=RgSens.ItemIndex ;
if CbCouple1.Items.Count<=0 then Exit ;
ChargeLesPeriodes(THValComboBox(CbPS)) ;
Case RgTyp.ItemIndex of
     0 : SurQuoi:='G/S;' ;
     1 : SurQuoi:='CBG;' ;
   End ;
BBudUp.Enabled:=False ; BBudDown.Enabled:=False ;
Etab:=EtabS.Value ; Dev:='' ; BNat:=NatS.Value ;
St:='' ; SQuelTyp(WhatTypeEcr('TOU',True,cbUnchecked),St) ; RubAno:=St ;
// Modif SBO : On compte les enregistrements à part
lInNbEnreg := 0 ;
QLoc:=OpenSql('SELECT COUNT(*) TOTAL FROM RUBRIQUE WHERE RB_BUDJAL="'+BudS.Value+'" AND '+
              'RB_NATRUB="BUD" AND RB_TABLELIBRE="-" AND RB_FAMILLES="'+SurQuoi+'"',True) ;
if not QLoc.Eof then
  lInNbEnreg := QLoc.FindField('TOTAL').AsInteger ;
Ferme(QLoc) ;
InitMove(lInNbEnreg,'') ;
// Fin Modif SBO
QLoc:=OpenSql('Select * from RUBRIQUE Where RB_BUDJAL="'+BudS.Value+'" And '+
              'RB_NATRUB="BUD" And RB_TABLELIBRE="-" And RB_FAMILLES="'+SurQuoi+'" Order by RB_RUBRIQUE',True) ;
While Not QLoc.Eof do
   BEGIN
   MoveCur(False) ;
   for i:=0 to CbPS.Values.Count-1 do
      BEGIN
      D1:=StrToDate(CbPS.Values[i]) ; D2:=FinDeMois(D1) ; Exo:=QuelExoDt(D1) ;
      Cptbg:='' ; Cptbs:='' ;
      MontantBud:=ChercheBudgete(BudS.Value,QLoc.FindField('RB_RUBRIQUE').AsString,D1,D2,Exo,Etab,BNat,Cptbg,Cptbs,LavalBud) ;
      if(Cptbg='') or (Cptbs='') then
         BEGIN
         if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ; Break ;
         END ;
      if i=0 then
         BEGIN
         FListe.Cells[0,FListe.RowCount-1]:=Cptbg+' : '+Cptbs ;
         FListe.RowCount:=FListe.RowCount+3 ;
         FListe.Cells[0,FListe.RowCount-3]:='    '+HMG.Mess[0] ;
         FListe.Cells[0,FListe.RowCount-2]:='    '+HMG.Mess[1] ;
         FListe.Cells[0,FListe.RowCount-1]:='    '+HMG.Mess[2] ;
         END ;
      Fillchar(TResult,SizeOf(TResult),0) ;
      MontantRea:=GetCumul('BUDREA',QLoc.FindField('RB_RUBRIQUE').AsString,'',RubAno,Etab,Dev,Exo,D1,D2,False,True,Nil,TResult,FALSE) ;
      if (MontantRea<>0) and (TResult[2]<>0) then MontantRea:=FormateResultat(0,TResult[2],LavalRea)
                                             else MontantRea:=FormateResultat(TResult[1],0,LavalRea) ;
      X:=TInfoEcr.Create ; X.Montant:=MontantBud ;
      FListe.Cells[i+1,FListe.RowCount-3]:=LavalBud ; FListe.Objects[i+1,FListe.RowCount-3]:=X ;

      X:=TInfoEcr.Create ; X.Montant:=MontantRea ;
      FListe.Cells[i+1,FListe.RowCount-2]:=LavalRea ; FListe.Objects[i+1,FListe.RowCount-2]:=X ;

      X:=TInfoEcr.Create ;
      if RgSens.ItemIndex=0 then Ecart:=FormateResultat(MontantRea,MontantBud,LavalEca)
                            else Ecart:=FormateResultat(MontantBud,MontantRea,LavalEca) ;
      X.Montant:=Ecart ;
      FListe.Cells[i+1,FListe.RowCount-1]:=LavalEca ; FListe.Objects[i+1,FListe.RowCount-1]:=X ;
      END ;
   if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount+1 ;
   QLoc.Next ;
   END ;
FiniMove ; Ferme(QLoc) ;
if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
if FListe.RowCount>=4 then BEGIN FListe.Row:=4 ; OldRow:=FListe.Row ; FListe.SetFocus ; END ;
end;

procedure TFRevBuRea.BValiderClick(Sender: TObject);
begin
if FListe.Cells[0,FListe.RowCount-1]='' then Exit ;
if BudS.Value<>BudD.Value then if Not ControleCompteOk then Exit ;
if Not ControleCpteAttOk then Exit ;
if Not ControlePeriodeOk then Exit ;
if Not ControleNatureOk  then Exit ;
if Not ControleCategorOk then Exit ;
if HM.Execute(5,'','')<>mrYes then Exit ;
BeginTrans ; RunLaCreationAuProrata ; CommitTrans ;
end;

Function TFRevBuRea.ControleCompteOk : Boolean ;
Var StGS,StSS,St : String ;
    XS,XD : TInfoJal ;
    i : Integer ;
BEGIN
Result:=False ;
XS:=TInfoJal(BudS.Values.Objects[BudS.Values.IndexOf(BudS.Value)]) ;
XD:=TInfoJal(BudD.Values.Objects[BudD.Values.IndexOf(BudD.Value)]) ;
if XD.CptGen2<>'' then XD.CptGen:=XD.CptGen+XD.CptGen2 ;
if XD.CptSec2<>'' then XD.CptSec:=XD.CptSec+XD.CptSec2 ;
if Length(XD.CptGen)<>Length(XS.CptGen) then BEGIN HM.Execute(0,'','') ;  Exit ; END ;
if Length(XD.CptSec)<>Length(XS.CptSec) then BEGIN HM.Execute(1,'','') ;  Exit ; END ;
StGS:=XS.CptGen ; StSS:=XS.CptSec ;
While StGS<>'' do
   BEGIN
   St:=ReadTokenSt(StGS) ; i:=Pos(St,XD.CptGen) ;
   if i<=0 then
      BEGIN
      HM.Execute(0,'','') ; Exit ;
      END else
      BEGIN
      if XD.CptGen[i+Length(St)]<>';' then BEGIN HM.Execute(0,'','') ;  Exit ; END ;
      END ;
   END ;
While StSS<>'' do
   BEGIN
   St:=ReadTokenSt(StSS) ; i:=Pos(St,XD.CptSec) ;
   if i<=0 then
      BEGIN
      HM.Execute(1,'','') ;  Exit ;
      END else
      BEGIN
      if XD.CptSec[i+Length(St)]<>';' then BEGIN HM.Execute(1,'','') ;  Exit ; END ;
      END ;
   END ;
Result:=True ;
END ;

Function TFRevBuRea.ControleCpteAttOk : Boolean ;
BEGIN
Result:=False ;
if TInfoJal(BudS.Values.Objects[BudS.Values.IndexOf(BudS.Value)]).CptGen<>TInfoJal(BudD.Values.Objects[BudD.Values.IndexOf(BudD.Value)]).CptGen then
   BEGIN HM.Execute(3,'','') ; Exit ; END ;
if TInfoJal(BudS.Values.Objects[BudS.Values.IndexOf(BudS.Value)]).CptSec<>TInfoJal(BudD.Values.Objects[BudD.Values.IndexOf(BudD.Value)]).CptSec then
   BEGIN HM.Execute(4,'','') ; Exit ; END ;
Result:=True ;
END ;

Function TFRevBuRea.ControlePeriodeOk : Boolean ;
BEGIN
Result:=False ;
ChargeLesPeriodes(THValComboBox(CbPD)) ;
if CbPS.Items.Count<>CbPD.Items.Count then
   BEGIN if HM.Execute(2,'','')<>mrYes then Exit ; END ;
Result:=True ;
END ;

Function TFRevBuRea.ControleNatureOk : Boolean ;
BEGIN
Result:=False ;
if NatS.Value=NatDSup.Value    then BEGIN if HM.Execute(6,'','')<>mrYes then Exit ; END ;
if NatS.Value=NatDInf.Value    then BEGIN if HM.Execute(7,'','')<>mrYes then Exit ; END ;
if NatDInf.Value=NatDSup.Value then BEGIN if HM.Execute(8,'','')<>mrYes then Exit ; END ;
Result:=True ;
END ;

Function TFRevBuRea.ControleCategorOk : Boolean ;
BEGIN
Result:=False ;
if TInfoJal(BudS.Values.Objects[Buds.ItemIndex]).Categor<>TInfoJal(BudD.Values.Objects[BudD.ItemIndex]).Categor then BEGIN HM.Execute(10,'','') ; Exit ; END ;
Result:=True ;
END ;

Procedure TFRevBuRea.QuelCompteBudget(Var BudG,BudS : String ; Lig : Integer) ;
Var St : String ;
BEGIN
BudG:='' ; BudS:='' ;
St:=FListe.Cells[0,Lig] ;
if Pos(':',St)<=0 then Exit ;
BudG:=Trim(Copy(St,1,Pos(':',St)-1)) ; BudS:=Trim(Copy(St,Pos(':',St)+1,50)) ;
END ;

Function TFRevBuRea.ChercheNumPiece : LongInt ;
Var MM : String17 ;
    Sh : String ;
    NextNum : LongInt ;
BEGIN
MM:='' ; Sh:=TInfoJal(BudD.Values.Objects[BudD.ItemIndex]).Souche ;
NextNum:=GetNum(EcrBud,Sh,MM,0) ; Result:=NextNum ;
SetIncNum(EcrBud,Sh,NextNum,0) ;
END ;

Function TFRevBuRea.ChercheDebCred(Var Ecar : Double ; ARow,ACol : Integer) : Byte ;
BEGIN
Result:=0 ;
Ecar:=Valeur(FListe.Cells[ACol,ARow+3]) ;
if Ecar=0 then Result:=0 else
   if Ecar>0 then Result:=1 else
      if Ecar<0 then Result:=2 ;
(*Result:=0 ;
Budg:=TInfoEcr(FListe.Objects[ACol,ARow+1]).Montant ;
Real:=TInfoEcr(FListe.Objects[ACol,ARow+2]).Montant ;
Ecar:=Valeur(FListe.Cells[ACol,ARow+3]) ;
if (Ecar=0) or (Budg=Real) then Result:=0 else
   if Ecar>Budg then Result:=1 else
      if Ecar<Budg then Result:=2 ;*)
END ;

Procedure TFRevBuRea.RunLaCreationAuProrata ;
Var BudG,BudS,Nature,Exo : String ;
    i,j,k :Integer ;
    NextNumPiece : LongInt ;
    NowFutur,DCpta : TDateTime ;
    D,C,SomUp,SomDown,Ecar : Double ;
BEGIN
InitMove(FListe.RowCount-1,'') ; NowFutur:=NowH ; BMenuZoom.Enabled:=True ;
NumPieceSup:=0 ; NumPieceInf:=0 ;
(*
for i:=1 to FListe.RowCount-1 do
   BEGIN
   MoveCur(False) ;
   if((Pos(':',FListe.Cells[0,i])>0) and (FListe.RowHeights[i+1]=0)) or
     (Pos(':',FListe.Cells[0,i])<=0) then Continue ;
   QuelCompteBudget(BudG,BudS,i) ; SomUp:=0 ; SomDown:=0 ;
   for j:=1 to FListe.ColCount-1 do
       BEGIN
       k:=ChercheDebCred(Ecar,i,j) ;
       if k=0 then Continue else
         if k=1 then SomUp:=SomUp+Ecar else
            if k=2 then SomDown:=SomDown+Ecar ;
       END ;
   if Somup<>0 then
      BEGIN
      if LeSens=0 then BEGIN D:=Somup/CbPD.Items.Count ; C:=0 ; END
                  else BEGIN D:=0 ; C:=Somup/CbPD.Items.Count ; END ;
      if NumPieceSup=0 then NumPieceSup:=ChercheNumPiece ;
      NextNumPiece:=NumPieceSup ; Nature:=NatDSup.Value ;
      for j:=0 to CbPD.Values.Count-1 do
          BEGIN
          DCpta:=StrToDate(CbPD.Values[j]) ; Exo:=QuelExoDt(Dcpta) ;
          EcritLenreg(Exo,Nature,BudG,BudS,Dcpta,NowFutur,D,C,NextNumPiece) ;
          END ;
      END ;
   if SomDown<>0 then
      BEGIN
      if LeSens=0 then BEGIN D:=SomDown/CbPD.Items.Count ; C:=0 ; END
                  else BEGIN D:=0 ; C:=SomDown/CbPD.Items.Count ; END ;
      if NumPieceInf=0 then NumPieceInf:=ChercheNumPiece ;
      NextNumPiece:=NumPieceInf ; Nature:=NatDInf.Value ;
      for j:=0 to CbPD.Values.Count-1 do
          BEGIN
          DCpta:=StrToDate(CbPD.Values[j]) ; Exo:=QuelExoDt(Dcpta) ;
          EcritLenreg(Exo,Nature,BudG,BudS,Dcpta,NowFutur,D,C,NextNumPiece) ;
          END ;
      END ;
   END ;
   *)
  {JP 24/04/07 : FQ 19935 : demande d'OG = ne pas proratiser les montants sur la période}
  for i:=1 to FListe.RowCount-1 do begin
    MoveCur(False) ;
    if ((Pos(':',FListe.Cells[0,i])>0) and (FListe.RowHeights[i+1]=0)) or
       (Pos(':',FListe.Cells[0,i])<=0) then Continue;
    QuelCompteBudget(BudG,BudS,i) ; SomUp:=0 ; SomDown:=0 ;
    for j := 1 to FListe.ColCount - 1 do begin
      {Recherche de l'écart et de son sens}
      k := ChercheDebCred(Ecar, i, j);
      {L'écart est à 0, on passe}
      if k = 0 then Continue;

      if Ecar <> 0 then begin
        if LeSens = 0 then begin
          D := Ecar;
          C := 0;
        end
        else begin
          D := 0;
          C := Ecar;
        end;
        if NumPieceSup = 0 then NumPieceSup := ChercheNumPiece;
        NextNumPiece := NumPieceSup;
        if k = 2 then Nature := NatDInf.Value
                 else Nature := NatDSup.Value;
        DCpta := StrToDate(CbPD.Values[j - 1]);
        Exo := QuelExoDt(Dcpta) ;
        EcritLenreg(Exo, Nature, BudG, BudS, Dcpta, NowFutur, D, C, NextNumPiece) ;
      end;
    end;
  end;
     (*
     if SomDown<>0 then
        BEGIN
        if LeSens=0 then BEGIN D:=SomDown/CbPD.Items.Count ; C:=0 ; END
                    else BEGIN D:=0 ; C:=SomDown/CbPD.Items.Count ; END ;
        if NumPieceInf=0 then NumPieceInf:=ChercheNumPiece ;
        NextNumPiece:=NumPieceInf ; Nature:=NatDInf.Value ;
        for j:=0 to CbPD.Values.Count-1 do
            BEGIN
            DCpta:=StrToDate(CbPD.Values[j]) ; Exo:=QuelExoDt(Dcpta) ;
            EcritLenreg(Exo,Nature,BudG,BudS,Dcpta,NowFutur,D,C,NextNumPiece) ;
            END ;
        END ;
     END ;
        *)

FiniMove ;
BBudUp.Enabled:=(NumPieceSup<>0) ;
BBudDown.Enabled:=(NumPieceInf<>0) ;
END ;

Procedure TFRevBuRea.EcritLenreg(Exo,Nat,Bg,Bs : String ; Dcpta,DModif : TDateTime ;
                                 D,C : Double ; NPiece : LongInt) ;
var
  TLoc : Tob;
BEGIN
  TLoc := Tob.Create('BUDECR', nil, -1);
  TLoc.InitValeurs;
  TLoc.SetString('BE_BUDJAL', BudD.Value);
  TLoc.SetString('BE_EXERCICE', Exo);
  TLoc.SetDateTime('BE_DATECOMPTABLE', Dcpta);
  TLoc.SetDateTime('BE_DATEMODIF', DModif);
  TLoc.SetString('BE_AXE', TInfoJal(BudD.Values.Objects[BudD.ItemIndex]).Ax);
  TLoc.SetInteger('BE_NUMEROPIECE', NPiece);
  TLoc.SetDouble('BE_DEBIT', D);
  TLoc.SetDouble('BE_CREDIT', C);
  TLoc.SetString('BE_NATUREBUD', Nat);
  TLoc.SetString('BE_BUDGENE', BG);
  TLoc.SetString('BE_BUDSECT', BS);
  TLoc.SetString('BE_QUALIFPIECE', 'N');
  TLoc.SetString('BE_TYPESAISIE', 'G');
  TLoc.SetString('BE_RESOLUTION', 'F');
  TLoc.SetString('BE_ETABLISSEMENT', EtabD.Value);
  TLoc.InsertDB(nil);
  TLoc.Free;
END ;
// Gestion évènement de la liste ******************************************

procedure TFRevBuRea.FListeDblClick(Sender: TObject);
Var i,Arow : Integer ;
begin
Arow:=FListe.Row ;
if (Arow=0) or (Pos(':',Fliste.Cells[0,Arow])<=0) then Exit ;
if FListe.RowHeights[Arow+1]<>0 then for i:=1 to 3 do FListe.RowHeights[Arow+i]:=0
                                else for i:=1 to 3 do FListe.RowHeights[Arow+i]:=18 ;
end;

procedure TFRevBuRea.FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Var ACol,ARow : LongInt ;
begin
FListe.MouseToCell(X,Y,ACol,ARow) ;
if (ARow>=1) And (ACol=0) then FListe.Row:=ARow ;
//if Pos(':',FListe.Cells[ACol,ARow])>0 then
end;

procedure TFRevBuRea.BAnnulerClick(Sender: TObject);
Var i,ARow : Integer ;
    Mount,Bidon : Double ;
    LaValeur : String ;
begin
if FListe.Cells[0,FListe.RowCount-1]='' then Exit ;
ARow:=FListe.Row ;
While Pos(HMG.Mess[2],FListe.Cells[0,ARow])<=0 do Inc(ARow) ;
for i:=1 to FListe.ColCount-1 do
   BEGIN
   Mount:=TInfoEcr(FListe.Objects[i,Arow]).Montant ; LaValeur:='' ; Bidon:=0 ;
   if RgSens.ItemIndex=0 then FormateResultat(Mount,Bidon,LaValeur) else FormateResultat(Bidon,Mount,LaValeur) ;
   FListe.Cells[i,Arow]:=LaValeur ;
   END ;
end;

procedure TFRevBuRea.FListeCellExit(Sender: TObject; var ACol,ARow: Longint; var Cancel: Boolean);
Var St : String ;
begin
if (Pos(HMG.Mess[2],FListe.Cells[0,ARow])<=0) or (ACol=0) then Exit ;
St:=FListe.Cells[ACol,ARow] ;
if (Pos('-',St)=1) or (Pos('+',St)=1) then St:=Copy(St,2,Length(St)) ;
if Not IsNumeric(St) then BEGIN HM.Execute(9,'','') ; Cancel:=True ; Exit ; END ;
FListe.Cells[ACol,ARow]:=StrfMontant(Valeur(FListe.Cells[ACol,ARow]),15,2,'',True) ;
end;

procedure TFRevBuRea.BBudUpClick(Sender: TObject);
begin EnvoiVisuPiece(NatDSup.Value,NumPieceSup) ; end;

procedure TFRevBuRea.BBudDownClick(Sender: TObject);
begin EnvoiVisuPiece(NatDInf.Value,NumPieceInf) ; end;

Procedure TFRevBuRea.EnvoiVisuPiece(Nat : String ; NumP : LongInt) ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select * from BUDECR Where BE_BUDJAL="'+BudD.Value+'" And BE_NATUREBUD="'+Nat+'" '+
              'And BE_NUMEROPIECE='+IntToStr(NumP)+' And BE_QUALIFPIECE="N"',True) ;
if Not QLoc.Eof then TrouveEtLanceSaisBud(QLoc,taConsult) ;
Ferme(QLoc) ;
END ;

procedure TFRevBuRea.BImprimerClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
var
  T, F : Tob;
  i,j : Integer;

  procedure Add(L : Array of String);
  var
    i : Integer;
  begin
    for i := 0 to Length(L)-1 do
      F.AddChampSup(L[i], False);
  end;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  SourisSablier;
  T := TOB.Create('non', nil, -1);

  // Contenu de la liste
  for i := 0 to FListe.RowCount-2 do begin
    F := TOB.Create('Ligne '+IntToStr(i), T, -1);
    // Contenu des lignes
    Add(['CH1', 'CH2', 'CH3', 'CH4', 'CH5', 'CH6', 'CH7', 'CH8', 'CH9', 'CH10', 'CH11', 'CH12', 'CH13']);
    for j := 0 to FListe.ColCount-1 do
      F.SetString('CH'+IntToStr(j+1), Trim(FListe.Cells[j, i+1]));
    for j := FListe.ColCount to 12 do
      F.SetString('CH'+IntToStr(j+1), '');
  end;

  // Titres des colonnes
  F := T.Detail[0];
  Add(['L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'L9', 'L10', 'L11', 'L12']);
  for j := 1 to FListe.ColCount-1 do
    F.SetString('L'+IntToStr(j), FListe.Cells[j, 0]);
  for j := FListe.ColCount to 12 do
    F.SetString('L'+IntToStr(j), '');

  SourisNormale;
  LanceEtatTob('E','CST','CBR',T, True, False, False, nil, '', Caption, False, 0, '', 0, '');
  T.Free;
{$ELSE}
  PrintDBGrid(FListe,Nil,Caption,'') ;
{$ENDIF}
end;

procedure TFRevBuRea.POPSPopup(Sender: TObject);
begin
if FListe.Cells[0,FListe.RowCount-1]='' then
   BEGIN
   BCpy.Enabled:=False ; BcpyLig.Enabled:=False ; BcolLig.Enabled:=False ;
   END else
   BEGIN
   BCpy.Enabled:=True ; BcpyLig.Enabled:=True ; BcolLig.Enabled:=True ;
   END ;
InitPopUp(Self) ;
end ;

procedure TFRevBuRea.BCpyClick(Sender: TObject);
Var ARow,i : Integer ;
begin
if FListe.Cells[0,FListe.RowCount-1]='' then Exit ;
ARow:=FListe.Row ;
While Pos(':',FListe.Cells[0,ARow])<=0 do Dec(ARow) ;
for i:=1 to FListe.ColCount-1 do FListe.Cells[i,ARow+3]:=StrfMontant(Valeur(FListe.Cells[i,ARow+2]),15,2,'',True) ;
end;

procedure TFRevBuRea.BcpyLigClick(Sender: TObject);
Var i : Integer ;
begin
if Pos(':',FListe.Cells[0,FListe.Row])>0 then Exit ;
CollerLigne.Clear ;
for i:=1 to FListe.ColCount-1 do CollerLigne.Add(FListe.Cells[i,FListe.Row]) ;
end;

procedure TFRevBuRea.BcolLigClick(Sender: TObject);
Var i : Integer ;
begin
if CollerLigne.Count<=0 then Exit ;
if Pos(HMG.Mess[2],FListe.Cells[0,FListe.Row])<=0 then Exit ;
for i:=1 to FListe.ColCount-1 do FListe.Cells[i,Fliste.Row]:=StrfMontant(Valeur(CollerLigne.Strings[i-1]),15,2,'',True) ;
end;

procedure TFRevBuRea.RgTypClick(Sender: TObject);
begin
TCbCouple1.Caption:=HMG.Mess[5+RgTyp.ItemIndex] ;
ChargeCroisement(BudS.Value,TInfoJal(BudS.Values.Objects[BudS.ItemIndex]).Categor) ;
end;

procedure TFRevBuRea.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ;
end;

procedure TFRevBuRea.BFermeClick(Sender: TObject);
begin
  Close;
  if IsInside(Self) then
    CloseInsidePanel(Self);
end;

{---------------------------------------------------------------------------------------}
procedure TFRevBuRea.ControlEtab;
{---------------------------------------------------------------------------------------}
var
  Eta : string;
begin
  if not Assigned(EtabS) then Exit;
  {S'il n'y a pas de gestion des établissement, logiquement, on ne force pas l'établissement !!!}
  if not VH^.EtablisCpta then Exit;

  Eta := EtabForce;
  {S'il y a une restriction utilisateur et qu'elle ne correspond pas au contenu de la combo ...}
  if (Eta <> '') and (Eta <> EtabS.Value) then begin
    {... on affiche l'établissement des restrictions}
    EtabS.Value := Eta;
    {... on désactive la zone}
    EtabS.Enabled := False;
  end;

  if not Assigned(EtabD) then Exit;
  {S'il y a une restriction utilisateur et qu'elle ne correspond pas au contenu de la combo ...}
  if (Eta <> '') and (Eta <> EtabD.Value) then begin
    {... on affiche l'établissement des restrictions}
    EtabD.Value := Eta;
    {... on désactive la zone}
    EtabD.Enabled := False;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TFRevBuRea.GereEtablissement;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(EtabS) then begin
    {Si l'on ne gère pas les établissement ...}
    if not VH^.EtablisCpta  then begin
      {... on affiche l'établissement par défaut}
      EtabS.Value := VH^.EtablisDefaut;
      {... on désactive la zone}
      EtabS.Enabled := False;
    end

    {On gère l'établisement, donc ...}
    else begin
      {... On commence par regarder les restrictions utilisateur}
      PositionneEtabUser(EtabS);
      {... s'il n'y a pas de restrictions, on reprend le paramSoc
       JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
      if EtabS.Value = '' then begin
        {... on affiche l'établissement par défaut
        EtabS.Value := VH^.EtablisDefaut;
        {... on active la zone
        EtabS.Enabled := True;
      end;}
    end;
  end;

  if Assigned(EtabD) then begin
    {Si l'on ne gère pas les établissement ...}
    if not VH^.EtablisCpta  then begin
      {... on affiche l'établissement par défaut}
      EtabD.Value := VH^.EtablisDefaut;
      {... on désactive la zone}
      EtabD.Enabled := False;
    end

    {On gère l'établisement, donc ...}
    else begin
      {... On commence par regarder les restrictions utilisateur}
      PositionneEtabUser(EtabD);
      {... s'il n'y a pas de restrictions, on reprend le paramSoc
       JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
      if EtabD.Value = '' then begin
        {... on affiche l'établissement par défaut
        EtabD.Value := VH^.EtablisDefaut;
        {... on active la zone
        EtabD.Enabled := True;
      end;}
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TFRevBuRea.OnApresChangementFiltre;
{---------------------------------------------------------------------------------------}
begin
  ControlEtab;
end;

end.
