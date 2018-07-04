unit LP_View;

interface

uses Windows, Menus, Hctrls, Dialogs, ImgList, Controls, LP_Base, Forms,
  Spin, StdCtrls, ExtCtrls, HTB97, Classes, LP_Impri, MC_Erreur, UTob
{$IFDEF eAGLClient}
{$ELSE}
  , DBTables//, ColorGrd //XMG 05/04/04 ne sert a rien....
{$ENDIF}
  ;

type

  TFLPPreView = class(TForm)
    TopDock: TDock97;
    LeftDock: TDock97;
    RightDock: TDock97;
    BottomDOck: TDock97;
    GeneralBox: TToolbar97;
    Fonds: TScrollBox;
    LPBase: TLPBase;
    BClose: TToolbarButton97;
    ImagesBtns: TImageList;
    BAide: TToolbarButton97;
    PagesBox: TToolbar97;
    ZoomBox: TToolbar97;
    BFitPage: TToolbarButton97;
    BFitLarge: TToolbarButton97;
    BFit100: TToolbarButton97;
    TZoom: THLabel;
    FZoom: TSpinEdit;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    FInfoPanel: TPanel;
    FTotPage: TLabel;
    FPageCourant: THNumEdit;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BRecherche: TToolbarButton97;
    BExport: TToolbarButton97;
    Find: TFindDialog;
    HMTrad: THMainMenu;
    procedure FormResize(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BFitPageClick(Sender: TObject);
    procedure BFitLargeClick(Sender: TObject);
    procedure BFit100Click(Sender: TObject);
    procedure FPageCourantExit(Sender: TObject);
    procedure FZoomChange(Sender: TObject);
    procedure BExportClick(Sender: TObject);
    procedure BRechercheClick(Sender: TObject);
    procedure FindFind(Sender: TObject);
    procedure FindClose(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BCloseClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FondsMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
  private
    BtnGros       : Boolean ; 
    Fprinter      : TLPPrinter ;
    FProgress     : TFLPProgress ;
    Pages         : TList ;
    PageFind,
    LigneFind,
    CharFind,
    OriginWidth,
    OriginHeight  : integer ;
    Marges        : TRect ;
    TextSel       : TLPChamp ;
    InFacturette  : Boolean ;
    InValidation  : Boolean ;
    InCheque      : Boolean ;
    FLastErrorImp : TImprimErreurs ;
    LaPage        : TList ;
    FIsExport     : Boolean ; //XMG 30/03/04
    Procedure VideVideListe (Laliste : TObject) ;
    Procedure LirePage ( NumPage : Integer ; AutoVisible : Boolean ) ;
    Procedure AnnuleSelect ;
    Function  CherchePage(NPage : Integer ) : TList ;
    Function  ChercheLigne(NPage,NLIgne: Integer ) : TList ;
    function  Combienlignes : Integer ;
    Function  RecupereValeursImpr ( Lg : TLPChamp ) : String ;
    Function  ImprimeLigne(Ligne : TList) : TLPSiValeurZero ;
    Procedure ResizeBtns ;
    Procedure ChargePage (UnePage : TList) ;
  public
    PrinterPort    : String ;
    PrinterParams  : String ;
    Imprimante     : String ;
    IsImpControl   : Boolean ;
    InitImp        : Boolean ;
    InPreview      : Boolean ;
    InTestMateriel : Boolean ;
    Procedure InitView ( ALPBase : TLPBase ) ;
    function  NewPage : Boolean ;
    Function  AddChamp( Lgn : Integer ; Champ : TControl) : Boolean ;
    Function  Addlignes (NLignes : Integer ) : Boolean ;
    Function  Imprime (AvecProgress : Boolean ; UnExemplaire : Boolean ; var Err : TMC_Err )  : Boolean ;
    Procedure CalcelImpression ;

    property LastError    : TImprimErreurs read FLastErrorimp ;

    Property IsExport     : Boolean read FIsExport write FIsExport ; //XMG 30/03/04
    end;

  TRuptures = Record
              Champ : String ;
              Valeur : Variant ;
              end ;

  TLPImprime = Class ( TComponent)
     Private
      Pixel           : TLPPixel ;
      FProgress       : TFLPProgress ;
      FForm           : TForm ;
      LpBase          : TLPBase ;
      LPParam         : TLPParam ;
      QPrin,QSbDet    : TQuery ;
      BandeCourant    : TLPBande ;
      InPreView       : Boolean ;
      FPreView        : TFLPPreView ;
      MaxLignePage    : Integer ;
      LignePage       : Integer ;
      Fport           : String ;
      FParams         : String ;
      Fimpr           : String ;
      FInitImp        : Boolean ;
      TicketCoupe     : Boolean ;
      InFacturette    : Boolean ;
      InValidation    : Boolean ;
      InCheque        : Boolean ;
      TobRupt,
      UneTOB          : TOB ;
      FSQLBased       : Boolean ;
      isEof           : Boolean ;
      IdxTOB          : Integer ;
      Totaliser       : TOB ;
      FRuptures       : Array of TRuptures ;
      FNbRuptures     : Integer ;
      IsImpControl    : Boolean ;
      FWhereLocal     : String ;
      FInTestMateriel : Boolean ;
      ChampsRequete   : TStringList ;
      ChampsAliases   : TStringList ; //XMG 07/04/04
      IsVerifSQL      : Boolean ; //XMG 25/06/03
      ErrorFormule    : String ; //XMG 25/06/03
      Procedure RajouteWhere ( Where : String ) ;
      Function  TraiteBande ( Bande : TLPBande ; Var Err : TMC_Err ) : Boolean ;
      Function  TraiteSubBande ( NombreSubBnd : Integer ; TypeSubBnd : TLPTypeBandes ; Var Err : TMC_Err ) : Boolean ;
      Function RenvoieNomchampRequete(Champ : String ) : String ; //XMG 07/04/04
      Function  PrendChamp ( Champ : String ) : Variant ;
      Procedure ChargeVariables ;
      Function ChercheChamp(Champ : String ; Var Evalue : Boolean ) : String ; //XMG 25/06/03
      Function  TestCondition ( Bande : TLPBande ) : Boolean ;
      Procedure InitPreView ;
      Function  NouvellePage : Boolean ;
      Function  SautLignes ( NLignes : Integer ) : Boolean ;
      Function  WritelnLP ( St : String ) : Boolean ;
      Function ChargeparamsQ ( SQLOri : String ) : String  ; //XMG 25/06/03
      Function  TraiteFonction ( var Formule : String ) : Boolean ; //XMG 25/06/03
      Function  TraiteFormule ( LaMasque,Formule : String ; Var Resultat : String ) : Boolean ; //XMG 25/06/03
      Function  InsertChampTEX ( QueChamp : String ) : Boolean ;
      Procedure RechercheTotaux ;
      Procedure ChargeRuptures ;
      Function  TesteRuptures  : Integer ;
      Function  TraiteDesRupt(DesRupt : Integer ; Tb : TLPTypeBandes ; Var Err : TMC_Err ; Final : Boolean = FALSE ) : Boolean ;
     protected
      Function GereGFormule ( Formule : String ; Var BienPasse : Boolean ) : String ; //XMG 25/06/03
      function  VerifParam ( Var Err : TMC_Err ) : Boolean ; virtual ;
      Function GetLastError : TImprimErreurs ;
      Function  ChargeChampsrequete( Champ : String ) : Variant ;
      Procedure RechercheChampsRequete ( Parent : TWinControl ) ;
      Procedure AjouteChampRequete ( ch : String ) ;
      Procedure ParcourtRequete ( St : String ) ;
      Procedure GardeLesAliases ( NomChamp : String ) ; //XMG 07/04/04
      Function verifiechamps(SQL : String ) : String ;
     Public
      Constructor Create (AOwner : TComponent) ; override ;
      Destructor destroy ; override ;
      Function InitLP (TypeEtat,NatEtat,CodeEtat,Where,Imprimante,Port,Params : String ; Initialise,Preview,SQLBased : Boolean ; LaTOB : TOB ; Var Err : TMC_Err ; AIsImpControl : Boolean = FALSE ) : Boolean ;
      Function Imprime (Var Err : TMC_Err ) : Boolean ;
      Procedure ForceCancel ;

      property LastError      : TImprimErreurs read GetLastError ;
      Property InTestMateriel : Boolean read FInTestMateriel write FInTestMateriel ;
     End ;

Function ImprimeLP (TypeEtat,NatEtat,CodeEtat,Where,Imprimante,Port,Params : String ; Initialise,PreView,SQLBased : Boolean ; UneTOB : TOB ; Var Err : TMC_Err ; IsImpControl : Boolean = FALSE ; IsTestMateriel : Boolean = FALSE  ) : Boolean ;
Function Init_ImprLP (Imprimante,Port,Params : String ; parle : boolean ) : Boolean ;
Function ExporteLP(TypeEtat,NatEtat,CodeEtat,Where,NomFichier : String ; PreView,SQLBased : Boolean ; UneTOB : TOB ; Var Err : TMC_Err ) : Boolean ;


implementation

USes Graphics, Math, hent1, Sysutils, HMsgBox,
{$IFNDEF EXPORTASCII}
     MC_Admin,TR_Base,
{$ENDIF EXPORTASCII} //XMG 30/03/04
     ED_Tools, MC_Lib, Formule,
     ComCtrls, SQLParts, ParamSoc, HDebug, //XMG 25/06/03
{$IFDEF EAGLCLIENT}
     UtileAGL
{$ELSE}
     uEdtComp, Db
{$ENDIF}
     ;

Const Zero  : Integer = 0 ;
      Somme : String = 'SOMME(' ;
      Absol : String = 'ABSOLUE(' ;

{$R *.DFM}
Procedure TFLPPreView.InitView ( ALPBase : TLPBase ) ;
Begin
LPBase.assign(ALPBase) ;
LPBase.Units:=lunone ;
LPBase.Color:=clwhite ;
LPBase.pixel.asPixel:=LPbase.MargeLeft   ; Marges.left:=LPbase.Pixel.AsChar[FALSE,lpTailleDef]  ; LPBase.Margeleft:=0 ;
LPBase.pixel.asPixel:=LPbase.MargeTop    ; Marges.Top:=LPbase.pixel.AsChar[TRUE,LPTailleDef]    ; LPBase.MargeTop:=0 ;//Aucune importance
LPBase.pixel.asPixel:=LPbase.MargeRight  ; Marges.Right:=LPbase.pixel.AsChar[FALSE,lpTailleDef] ; LPBase.MargeRight:=0 ;//Aucune Importance
LPBase.pixel.asPixel:=LPbase.MargeBottom ; Marges.Bottom:=LPbase.pixel.AsChar[TRUE,lpTailleDef] ; LPBase.MargeBottom:=0 ;//Aucune imnportance
OriginWidth:=LPBase.Width ;
OriginHeight:=LPbase.tailleligne ;  
if not LPBase.Rouleau then OriginHeight:=LPBase.Height ; //??? 
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.FormResize(Sender: TObject);
begin
LpBase.Left:=maxintValue([8,(LpBase.parent.ClientWidth-LPBase.Width-8) div 2]) ;
LpBase.Top:=maxintValue([8,(LpBase.parent.ClientHeight-LPBase.Height-8) div 2]) ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.FormCreate(Sender: TObject);
begin
Fprogress:=TFLPProgress.Create(Self) ;
FProgress.Libelle:=TraduireMemoire(MC_MsgErrDefaut(1026)) ;
Pages:=TList.Create ;
TextSel:=nil ;
InFacturette:=FALSE ;
InValidation:=FALSE ;
InCheque:=FALSE ;
IsImpControl:=FALSE ;
BtnGros:=FALSE ;
FLastErrorIMP:=ieUnknown ;
LaPage:=TList.Create ;
IsExport:=FALSE ; //XMG 30/03/04
end;
//////////////////////////////////////////////////////////////////////////////////
Procedure TFLPPreView.VideVideListe (Laliste : TObject ) ;
Begin
//XMG 24/02/04 début
   If (LaListe is TList) and (TList(Laliste).Count>0) then begin
      while TList(Laliste).Count>0 do Begin
         if TObject(TList(LaListe)[TList(Laliste).Count-1]) is TList then
            VideVideListe(TList(LaListe)[TList(Laliste).Count-1])
         else
            TObject(TList(LaListe)[TList(Laliste).Count-1]).Free ;
         TList(LaListe).Delete(TList(Laliste).Count-1) ;
      End ;
   End ;
   LaListe.Free ;
{If (LaListe is TList) and (TList(Laliste).Count>0) and (TObject(TList(LaListe)[0]) is TList) then VideVideListe(TList(LaListe)[0])
                                                                                             else LaListe.Free ;}
//XMG 24/02/03 fin
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.FormDestroy(Sender: TObject);
{var i,j,k        : integer ;
    PageC,LIgneC : TList ;}
begin
Videliste(LaPage) ;
LaPage.Free ;
VideVideListe(Pages) ;
{For i:=0 to Pages.Count-1 do
    Begin
    PageC:=TList(Pages[i]) ;
    for j:=0 to PageC.count-1 do
        Begin
        LigneC:=TList(PageC[j]) ;
        for k:=0 to LigneC.count-1 do
            TObject(LigneC[k]).Free ;
        LIgneC.Free ;
        End ;
    PageC.Free ;
    End ;
Pages.Free ;}
FProgress.Free ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPPreView.CherchePage(NPage : Integer ) : TList ;
Begin
Result:=nil ;
if (Npage<>-1) and (not (NPage in [1..Pages.Count])) then exit ;
if (NPage=-1) then
   Begin
   if (Pages.Count>0) then Result:=TList(Pages.Items[Pages.Count-1]) else
      Begin
      Result:=TList.Create ;
      Pages.Add(Result) ;
      End ;
   End else Result:=TList(Pages.Items[NPage-1]) ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPPreView.ChercheLigne(NPage,NLIgne: Integer ) : TList ;
Var PageC,LIgneC : TList ;
    i            : Integer ;
Begin
Result:=nil ;
PageC:=CherchePage(NPage) ;
if PageC=nil then exit ;
if (NLigne>PageC.Count) or ((NLigne=-1) and (PageC.Count=0)) then
   for i:=PageC.Count to NLigne-1 do
      Begin
      LigneC:=TList.Create ;
      PageC.Add(LigneC) ;
      End ;
if (NLigne<0) then
   BEgin
   if NLigne<-1 then
      Begin
      LigneC:=TList.Create ;
      PageC.Add(LigneC) ;
      End ;
   Result:=TList(PageC.Items[PageC.Count-1]) ;
   End else Result:=TList(PageC.Items[NLigne-1]) ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function TFLPPreview.NewPage : Boolean ;
var pageC : TList ;
Begin
pageC:=TList.Create ;
result:=(pages.Add(PageC)>-1) ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPPreview.Addlignes (NLignes : Integer ) : Boolean ;
Var PageC,LIgneC : TList ;
    i            : Integer ;
Begin
Result:=(NLignes=0) ;
if NLIgnes=0 then exit ;
PageC:=CherchePage(-1) ;
if (PageC=nil) or (PageC.Count+NLignes>(LPBase.height div LpBase.tailleligne)) then exit ;
for i:=PageC.Count to PageC.Count+NLignes-1 do
  Begin
  LigneC:=TList.Create ;
  result:=(PageC.Add(LigneC)>-1) ;
  End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPPreView.AddChamp ( Lgn : Integer ; Champ : TControl) : Boolean ;
Var LIgneC : TList ;
    Ctrl   : TControl ;

Begin
Result:=FALSE ;
LigneC:=ChercheLigne(-1,Lgn) ;
if LigneC=nil then exit ;
Ctrl:=Nil ;
if (Champ is TLPChamp) then
   Begin
   Ctrl:=TLPChamp.Create(Self) ;
   with TLPCHamp(Ctrl) do
     Begin
     Parent:=nil ;
     Assign(Champ) ;
     Visible:=(PrinterVisible) and (not InVisibleParZero) ;
     Colonne:=Colonne+Marges.left ;
     if Lgn>0 then Ligne:=Lgn ; //XMG 23/06/03
     End ;
   End else
if Champ is TLPImage then
   Begin
   Ctrl:=TLPImage.Create(Self) ;
   with TLPImage(Ctrl) do
     Begin
     Parent:=nil ;
     Assign(Champ) ;
     Visible:=(PrinterVisible) ;
     Colonne:=Colonne+Marges.left ;
     if Lgn>0 then Ligne:=Lgn ; //XMG 23/06/03
     ImprimeImage ;
     End ;
   End else ;
if assigned(Ctrl) then Result:=(LigneC.Add(Ctrl)>-1) ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPPreView.ResizeBtns ;
Var Coef : double ;
    i    : integer ;
    B    : TToolbarButton97 ;
Begin
BtnGros:=not BtnGros ;
if BtnGros then Coef:=1.5 else Coef:=1/1.5 ;
for i:=0 to componentCount-1 do
    if Components[i] is TToolBarButton97 then
       Begin
       B:=TToolBarButton97(Components[i]) ;
       B.Left:=Round(B.Left*Coef) ;
       B.Top:=Round(B.Top*Coef) ;
       B.Height:=Round(B.Height*Coef) ;
       B.Width:=Round(B.Width*Coef) ;
       End ;

with Fonds do
  Begin
  with VertScrollBar do
   Begin
   Size:=40*ord(BtnGros) ;
   ButtonSize:=40*ord(BtnGros) ;
   End ;
  With HorzScrollBar do
   Begin
   Size:=40*ord(BtnGros) ;
   ButtonSize:=40*ord(BtnGros) ;
   End ;
  End ; 
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.FormShow(Sender: TObject);
begin
BImprimer.Visible:=Not IsExport ; //XMG 30/03/04 
if V_PGI.ZoomForm then ResizeBtns ;
FtotPage.Caption:=inttostr(pages.Count) ;
{$IFNDEF EXPORTASCII}
BImprimer.Enabled:=(UsesTPV(V_MC.GetDispositif(mcPrinter),mcPrinter)) ;
BExport.Enabled:=(isExport) or (UsesTPV(V_MC.GetDispositif(mcPrinter),mcPrinter)) ; //XMG 02/04/04
{$ELSE}
BExport.Enabled:=TRUE ;
{$ENDIF EXPORTASCII} //XMG 30/03/04
PagesBox.Visible:=Pages.Count>1 ;
if Not PagesBox.Visible then GeneralBox.DockPos:=ZoomBox.DockPos+ZoomBox.Width ;
LirePage(1,TRUE) ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPPreView.ImprimeLigne(Ligne : TList) : TLPSiValeurZero ;
Var i  : Integer ;
    Lg : TLPChamp ;
Begin
Result:=szImprimer ;
i:=0 ;
while (i<=Ligne.Count-1) and (Result=szImprimer) do
  Begin
  if TObject(Ligne.items[i]) is TLPChamp then
    Begin
    lg:=TLPChamp(Ligne.items[i]) ;
    if (lg.SiZero<>szImprimer) then
       begin
       if (isNumeric(Lg.Value)) then
          Begin
          if (valeur(lg.Value)=0) then Result:=Lg.SiZero ;
          End else if (Trim(lg.Value)='') then Result:=Lg.SiZero ;
       End ;
    End ;
  inc(i) ;
  End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPPreview.LirePage ( NumPage : Integer ; AutoVisible : Boolean ) ;
Var i,j,LignesSupprimees,LC : Integer ;
    Lg                      : TLPChamp ;
    Im                      : TLPImage ;
    PageC,LigneC            : TList ;
Begin
LPBase.Hide ;
if not (NumPage in [1..Pages.Count]) then exit ;
While LpBase.ControlCount>0 do LPBase.COntrols[0].Free ;
PageC:=Cherchepage(Numpage) ;
if PageC=nil then Exit ;
//Etablir la longeur et l'hauter de la page
LPbase.Width:=OriginWidth ;


Fprogress.SetMax(pageC.Count) ;
FProgress.SetValue(0) ;
FProgress.Libelle:=TraduireMemoire(MC_MsgErrDefaut(1026)) ;
FProgress.Show ; LignesSupprimees:=0 ;
if LPBase.Rouleau then
   Begin
   LPBase.Height:=LPBase.TailleLigne*PageC.Count ;
   OriginHeight:=LPBase.Height ; //???
   End ;
For i:=0 to PageC.Count-1 do
  Begin
  FProgress.SetValue(FProgress.Progress+1) ;
  LigneC:=TList(PageC.items[i]) ;
  if FProgress.Canceled then break;
  LC:=i-LignesSupprimees+1 ; 
  case Imprimeligne(LigneC) of
    szImprimer      : Begin
                      for j:=0 to LigneC.Count-1 do
                        if TObject(LigneC.items[j]) is TLPChamp then
                          Begin
                          Lg:=TLPChamp.Create(Self) ;
                          Lg.Parent:=LPBase ;
                          lg.HideSelection:=FALSE ;
                          lg.Ctl3D:=FALSE ;
                          Lg.BorderStyle:=bsNone ;
                          lg.Assign(TLPChamp(LigneC.items[j])) ;
                          lg.Ligne:=LC ; 
                          lg.TextBloque:=TRUE ;
                          //lg.text:=lg.Value ;
                          End else
                        if TObject(LigneC.items[j]) is TLPImage then
                          Begin
                          Im:=TLPImage.Create(Self) ;
                          Im.Parent:=LPBase ;
                          Im.Assign(TLPImage(LigneC.items[j])) ;
                          Im.Ligne:=LC ; 
                          End ;
                      LPBase.RealigneLigne(LC) ;
                      End ;
    szSupprimeLigne : Begin
                      Inc(LignesSupprimees);
                      if LPBase.Rouleau then LpBase.Height:=LpBase.Height-LPBase.TailleLigne ;
                      End ;
    szLigneEnBlanc  : ;
    End ;
  End ;
FProgress.Hide ;
if AutoVisible then LPBase.Show ;
FPageCourant.Value:=NumPage ;

BFirst.enabled:=(NumPage>1) ;
BPrev.Enabled:=(NumPage>1) ;
BNext.Enabled:=(NumPage<Pages.Count) ;
BLast.Enabled:=(NumPage<Pages.Count) ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.BFirstClick(Sender: TObject);
begin
LirePage(1,TRUE) ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.BPrevClick(Sender: TObject);
begin
LirePage(round(FPageCourant.value-1),TRUE) ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.BNextClick(Sender: TObject);
begin
LirePage(round(FPageCourant.value+1),TRUE) ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.BLastClick(Sender: TObject);
begin
LirePage(Pages.Count,TRUE) ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.BImprimerClick(Sender: TObject);
var err : TMC_Err ;
Begin
if (not Imprime(TRUE,TRUE,Err)) and (err.Code<>0) then PGIBox(Err.Libelle,Caption) ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function TFLPPreview.Combienlignes : Integer ;
//XMG 11/06/03 début
var ii,jj,OldLigne : Integer ;
    Pg             : TList ;
    ZN             : TLP_ZoneImp ;
Begin
Result:=0 ;
for ii:=0 to Pages.Count-1 do
    Begin
    oldLigne:=-1 ;
    pg:=Tlist(Pages.items[ii]) ;
    for jj:=0 to pg.count-1 do
       Begin
       zn:=TLP_ZoneImp(pg.items[jj]) ;
       if zn.Ligne<>OldLigne then
          Begin
          inc(result) ;
          OldLigne:=zn.Ligne ;
          End ;
       End ;
    End ;
//XMG 11/06/03 fin
End ;
//////////////////////////////////////////////////////////////////////////////////
function LaPageSort (Index1, Index2 : Pointer ): Integer;
//Tri par Ligne et Colonne Descendant
Var ll1,cc1,ll2,cc2 : Longint ;
    Zn              : TLP_ZoneImp ;
Begin
Zn:=TLP_ZoneImp(Index1) ;
ll1:=Zn.Ligne ; cc1:=Zn.Colonne ;

Zn:=TLP_ZoneImp(Index2) ;
ll2:=Zn.Ligne ; cc2:=Zn.Colonne ;

Result:=(ll1-ll2) ;
if result=0 then Result:=(cc2-cc1) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TFLPPreView.ChargePage (UnePage : TList) ;
//XMG 25/06/03 début
var ii,jj,la,Max,lng,LgnSupprimes : Longint ;
    LigneC                        : TList ;
    im                            : TLPImage ;
    lg                            : TLPChamp ;
    Val,ExtraInfo                 : String ;
    Sl                            : TStringList ;
    Seq                           : TLP_SetSqcesPossibles ;
    CC                            : TLPSiValeurZero ;
    Zn                            : TLP_ZoneImp ;
    IsSpecial                     : Boolean ;
Begin
VideListe(LaPage ) ;
SL:=TStringList.Create ; LgnSupprimes:=0 ;
For jj:=0 to UnePage.Count-1 do
  Begin
  LigneC:=TList(UnePage[jj]) ;
  CC:=Imprimeligne(LigneC) ; inc(LgnSupprimes,ord(cc=szSupprimeLigne)) ;
  if (CC<>szSupprimeLigne) then
     for ii:=0 to LigneC.Count-1 do
        Begin
        ExtraInfo:='' ;
        if (TObject(LigneC[ii]) is TLPChamp) then
           Begin
           Lg:=TLPChamp(LigneC[ii]) ;
           IsSpecial:=(IsSpecialChamp(Lg)) ;
           if (lg.PrinterVisible) or IsSpecial then
              Begin
              la:=0 ;
              Sl.Clear ; Val:='' ; Lng:=1 ; Max:=1 ;
              if CC=szImprimer then
                 Begin
                 if IsSpecial then Val:=RecupereValeursImpr(Lg) else
                    Begin
                    if lg.IsCodeBarre then
                       Begin
                       Val:=Lg.Value ;
                       ExtraInfo:=Lg.Mask ; Max:=Lg.NbrLignes ;//XMG 10/07/03
                       End else
                       begin
                       SL.Text:=Lg.GetJustifiedValue ;  //XMG 25/06/03
                       Lng:=Lg.NbrChars ; Max:=Lg.NbrLignes ;
                       End ;
                    End ;
                 Seq:=lg.Police ;
                 End ;
              while la<max do
                Begin
                if SL.Count>0 then Val:=SL[la] ;
                Val:=Fprinter.DoTranslate(Val) ; //XMG 30/03/04
                Zn:=TLP_ZoneImp.Create ;
                with Zn do
                  Begin
                  TypeZone:=tzChamp ;
                  Ligne:=lg.Ligne+la-LgnSupprimes ;
                  NbrLignes:=Max ; //lg.NbrLignes ;
                  Colonne:=lg.Colonne ;
                  Long:=Lng ;
                  Taille:=Length(Val) ;
                  Police:=Seq ;
                  Donnees:=Val ;
                  Extra:=ExtraInfo ;
                  End ;
                LaPage.add(Zn) ;
                inc(la) ;
                End ;
              End ; //if (lg.PrinterVisible....
           End else
        if TObject(LigneC[ii]) is TLPImage then
           Begin
           Im:=TLPImage(LigneC[ii]) ;
           if im.PrinterVisible then
              Begin
              Seq:=im.Police ;
              for la:=1 to Im.Largeur do
                  Begin
                  Val:=Im.ImagetoSt(la-1) ;
                  Zn:=TLP_ZoneImp.Create ;
                  with Zn do
                    Begin
                    TypeZone:=tzImage ;
                    Ligne:=im.Ligne+la-1 ;
                    NbrLignes:=im.Largeur ;
                    Colonne:=im.Colonne ;
                    Long:=Im.Longeur ;
                    Taille:=Length(Val) ;
                    Police:=Seq ;
                    Donnees:=Val ;
                    Extra:=ExtraInfo ;
                    End ;
                  LaPage.add(Zn) ;
                  End ;
              End ;
           End ;
        End ; //For ii:=0...
  End ; //for JJ:=0...
SL.Free ;
LaPage.Sort(LaPageSort) ;
End ;
//XMG 25/06/03
//////////////////////////////////////////////////////////////////////////////////
Function TFLpPreview.Imprime (AvecProgress : Boolean ; UnExemplaire : Boolean ; var Err : TMC_Err )  : Boolean ;
Var NPage,jj,Colonne,Longeur,NbEx,lgVD,OldLigne,Ligne : Integer ;
    St                                                : String ;
    Zn                                                : TLP_ZoneImp ;
    ErrP                                              : TMC_Err ;
Const LigneDemo = 5 ; //XMG 30/03/04 On doit faire apparaitre le texte Demo toutes les LigneDemo ....
begin
result:=FALSE ;
err.Code:=0 ; Err.Libelle:='' ;
Enabled:=FALSE ;
if (Not Inpreview) and (Fprogress.Canceled) then exit ;
if UnExemplaire then NbEx:=1 else NbEx:=V_PGI.DefaultDocCopies ;
if NbEx < 1 then NbEx:=1 else if NbEx > 9 then NbEx:=9 ;
While NbEx > 0 do
   Begin
   InFacturette:=FALSE ;
   InValidation:=FALSE ;
   InCheque:=FALSE ;
   FPrinter:=TLPPrinter.Create(Self,imprimante) ;
   try
      FPrinter.InTestMateriel:=InTestMateriel ;
      Result:=FPrinter.ChargePortetParams(PrinterPort,PrinterParams,err) ;
      if Result then
         Begin
         repeat
           Result:=TRUE ;
           Fprinter.Initialise:=InitImp ;
           if AvecProgress then
              Begin
              Fprogress.SetMax(Combienlignes) ;
              FProgress.SetValue(0) ;
              FProgress.Libelle:=TraduireMemoire(MC_MsgErrDefaut(1027)) ;
              FProgress.Show ;
              End ;
           if Result then Result:=FPrinter.OuvrirImprimante(Err) ;
           if (result) and (not FProgress.Canceled) then
              Begin
              NPage:=0 ; FPrinter.Demarre ;
              while (Npage<=Pages.Count-1) and (Result) and (Not FProgress.Canceled) do
                Begin
                ChargePage(TList(Pages.items[NPage])) ;
                jj:=0 ; LgVD:=0 ; OldLigne:=-1 ;
                while (jj<=LaPage.Count) and (Result) and (Not FProgress.Canceled) do
                  Begin
                  //XMG 11/06/03 début
                  Zn:=nil ; Ligne:=-2 ;
                  if jj<=Lapage.Count-1 then
                     Begin
                     zn:=LaPage[jj] ;
                     Ligne:=Zn.Ligne ;
                     End ;
                  if (jj>=LaPage.Count) or (OldLigne<>Ligne) then
                     Begin
                     if OldLigne>-1 then Result:=FPrinter.EcrireLigne(St,OldLigne+lgVD) ;
                     if (V_PGI.VersionDemo) and ((Ligne mod LigneDemo = 0) or (Ligne-OldLigne>=LigneDemo)) and (Ligne>0) then //XMG 30/03/04
                        Begin
                        FPrinter.InitNewLigne ;
                        Result:=FPrinter.EcrireLigne(FPrinter.DoTranslate(MC_MsgErrDefaut(1077)),Ligne+lgVD) ; //XMG 30/03/04
                        inc(lgVD) ;
                        End ;
                     if AvecProgress then FProgress.SetValue(FProgress.Progress+1) ;
                     if jj<=LaPage.Count-1 then
                        Begin
                        OldLigne:=Ligne ;
                        St:=Format_String(' ',FPrinter.Largeur) ;
                        FPrinter.InitNewLigne ;
                        End ;
                     End ;
                  if jj<=LaPage.Count-1 then
                     Begin
                     Colonne:=Zn.Colonne ;
                     Longeur:=Zn.Long ;
                     Longeur:=MinIntValue([FPrinter.Largeur-Colonne,Longeur]) ;
                     Delete(St,Colonne,Longeur) ;
                     Insert(Zn.Donnees,St,Colonne) ;
                     FPrinter.FormatPolice(zn) ; //.Police,Colonne,LongData,) ;
                     End ;
                  //XMG 11/06/03 fin
                  inc(jj) ;
                  End ;//While jj:=...
                if result then Result:=FPrinter.SautPage ;
                inc(NPage) ;
                End ; //while i:=..
              if Fprogress.Canceled then
                 Begin
                 FPrinter.InitNewLigne ;
                 FPrinter.ecrireligne(FPrinter.Sqce[spLF,TRUE],FPrinter.LastLigne+1) ;
                 FPrinter.ecrireligne(FPrinter.Sqce[spLF,TRUE],FPrinter.LastLigne+1) ;
                 FPrinter.ecrireligne(FPrinter.DoTranslate(traduirememoire(MC_MsgErrDefaut(1031))),FPrinter.LastLigne+1) ; //XMG 30/03/04
                 FPrinter.ecrireligne(FPrinter.Sqce[spLF,TRUE],FPrinter.LastLigne+1) ;
                 FPrinter.ecrireligne(FPrinter.Sqce[spLF,TRUE],FPrinter.LastLigne+1) ;
                 End ;
              if (InFacturette) or (InValidation) or (InCheque) then
                 Begin //Force la fermeture du mode facturette, validation ou Chéque
                 FPrinter.InitNewLigne ;
                 if InFacturette then Fprinter.EcrireLigne(FPrinter.Sqce[spDesactivateSlip,TRUE],FPrinter.Lastligne+1) ;
                 if InValidation then Fprinter.EcrireLigne(FPrinter.Sqce[spDesactivateValidation,TRUE],FPrinter.Lastligne+1) ;
                 if InCheque     then Fprinter.EcrireLigne(FPrinter.Sqce[spFinCheque,TRUE],FPrinter.Lastligne+1) ;
                 End ;
              End ; //If result..
           if (FPrinter.isConnected) and (Not FPrinter.fermeImprimante(Errp)) then
              if Result then
                 Begin
                 Result:=FALSE ;
                 Err:=Errp;
                 End ;
         if AvecProgress then
            Begin
            Fprogress.Hide ;
            activate ;
            End ;
         until ((Result) or ((Not Result) and (FPrinter.LastError<>ieErrorPaper)) or (Fprogress.Canceled)) ;
         End else PGIBox(Err.Libelle,LPMessageCap) ;
    Finally
      FlastErrorImp:=FPrinter.LastError ;
      FPrinter.Free ;
      FPrinter:=Nil ;
      Enabled:=TRUE ;
    End ;
 if Result then Dec(NbEx) else NbEx:=-1 ;  
 End ;
if visible and canfocus then Setfocus ; 
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPPreview.CalcelImpression ;
Begin
Fprogress.ForceCancel ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.BFitPageClick(Sender: TObject);
Var M,M2 : Integer ;
begin
M:=LPBase.Parent.Width-16 ;
M:=Round(M*100/OriginWidth) ;
M2:=LPBase.Parent.height-8 ;
M2:=Round(M2*100/OriginHeight) ;
M:=minintValue([M,M2]) ;
FZoom.Value:=M ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.BFitLargeClick(Sender: TObject);
Var M : Integer ;
begin
M:=LPBase.Parent.Width-16 ;
M:=Round(M*100/OriginWidth) ;
FZoom.Value:=M ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.BFit100Click(Sender: TObject);
begin
FZoom.Value:=100 ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.FPageCourantExit(Sender: TObject);
begin
LirePage(round(FPageCourant.value),TRUE) ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.FZoomChange(Sender: TObject);
begin
LpBase.Width:=OriginWidth ;
lirePage(Round(FPageCourant.value),FALSE) ;
LPbase.scaleBy(FZoom.Value,100) ;
FormResize(nil) ;
LPBase.Show ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.BExportClick(Sender: TObject);
Var OldPort   : String ;
    Err       : TMC_Err ; 
begin
OldPort:=printerPort ;
PrinterPort:='FTX' ;

if (not Imprime(TRUE,TRUE,Err)) and (Err.Code<>0) then PGIBox(Err.Libelle,Caption) ;

PrinterPort:=OldPort ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.BRechercheClick(Sender: TObject);
begin
PageFind:=-1 ; Find.Execute ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.FindFind(Sender: TObject);
Var PageC,LIgneC : TList ;
    lg           : TLPChamp ;
    WholeW,
    MatchC  : Boolean ;
    i,
    Supprimes,
    PageF,
    LigneF,
    CharF,
    LblTop,    
    LblLeft,
    hp,hl   : Integer ;
    St,StCh : String ;
begin
LigneF:=-1 ; supprimes:=0 ;
MatchC:=(frMatchCase in Find.Options) ; WholeW:=(frWholeWord in Find.Options) ;
StCH:=Find.FindText ; if MatchC then StCH:=Uppercase(StCH) ;
if PageFind=-1 then Begin PageF:=round(FPageCourant.value) ; LigneFind:=-1 ; CharFind:=-1 ; End else PageF:=PageFind ;
If frdown in find.options then hp:=Pages.Count-1
   Else Begin PageF:=-PageF ; hp:=0 ; End ;
CharF:=CharFind ;
While PageF<=hp do
 Begin
 PageC:=CherchePage(abs(PageF)) ;
 if LigneFind=-1 then LigneF:=1 else LigneF:=LigneFind ;
 If frdown in find.options then hl:=pageC.Count-1
    Else Begin LigneF:=-LigneF ; hl:=0 ; End ;
 While LigneF<=hl do
   Begin
   LigneC:=chercheligne(abs(PageF),abs(LigneF)) ;
   //**gestionar el frdown in find options
   while CharF<=LigneC.count do
      Begin
      Lg:=TLPChamp(LigneC.items[CharF]) ;
      St:=Lg.Text ;
      if MatchC then St:=Uppercase(St) ;
      if WholeW then
         Begin
         CharF:=Pos(' '+StCh+' ',St) ; If CharF>0 then inc(CharF) ;
         if CharF<=0 then CharF:=Pos(StCh+' ',St) ;
         if CharF<=0 then Begin CharF:=Pos(' '+StCh,St) ; If CharF>0 then inc(CharF) ; End ;
         if CharF<=0 then CharF:=ord(StCh=St) ;
         End else CharF:=Pos(StCh,St) ;
      if CharF>0 then Break ;
      Inc(LigneF) ;
      End ;
   if CharF>0 then Break ;
   End ;
 Inc(PageF) ;
 End ;
if CharF>0 then
   Begin
   LigneF:=Abs(LigneF) ; PageF:=Abs(PageF) ;
   AnnuleSelect ;
   if PageF+1<>Round(FPageCourant.Value) then LirePage(PageF+1,TRUE) ;
   LblTop:=LPBase.TailleLigne*LigneF ;
   LblLeft:=LPBase.TailleChar*(CharF+Supprimes) ;
   For i:=0 to LPBase.ControlCount-1 do
     Begin
     if (LPBase.COntrols[i] is TLPChamp) then
        Begin
        TextSel:=TLPChamp(LPBase.Controls[i]) ;
        if (TextSel.Top=Lbltop) and (TextSel.Left<=LblLeft)  and (TextSel.left+TextSel.Width>=LblLeft) then
           Begin
           TextSel.selStart:=CharF+supprimes-(TextSel.Left div LPBAse.TailleChar)-1 ;
           TextSel.SelLength:=Length(Find.FindText) ;
           CharF:=CharF+Supprimes+Length(Find.FindText) ;
           CharFind:=CharF ;
           LigneFind:=LigneF ;
           PageFind:=PageF ;
           Break ;
           End ;
        End ;
     End ;
   End else PGIBox('Recherche terminée',Caption) ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPPreView.AnnuleSelect ;
Begin
if TextSel<>nil then
   Begin
   TextSel.SelStart:=0 ;
   TextSel.SelLength:=0 ;
   TextSel:=Nil ;
   End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.FondsMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
begin
if Fonds.VertScrollBar.IsScrollBarVisible then Begin Fonds.VertScrollBar.Position:=Fonds.VertScrollBar.Position-WheelDelta ; Handled:=TRUE ; End else
if Fonds.HorzScrollBar.IsScrollBarVisible then Begin Fonds.HorzScrollBar.Position:=Fonds.VertScrollBar.Position-WheelDelta ; Handled:=TRUE ; End ;
end;
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.FindClose(Sender: TObject);
begin
AnnuleSelect ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TFLPPreView.RecupereValeursImpr ( Lg : TLPChamp ) : String ;
Var Champ : String ;
Begin
Champ:=Lg.Libelle ;
Result:='' ;
if Copy(Champ,1,1)='[' then delete(champ,1,1) ;
if Copy(Champ,length(Champ),1)=']' then delete(champ,length(Champ),1) ;
//XMG 30/03/04 début
if (IsSpecialChampTexte(Champ)) and (not IsImpControl) and (not IsExport) then
   Begin //Variables speciales de l'impresion
{$IFNDEF EXPORTASCII}
   if Champ='TEX_PARTIEL' then Result:=FPrinter.Sqce[spPartialCut,True] else             //Coupe partialement le ticket
   if Champ='TEX_TOTAL'   then Result:=FPrinter.Sqce[spTotalCut,True] else               //Coupe totalement le ticket
   if Champ='TEX_OUVRE'   then OuvreTiroir else                                          //Ouvre le tiroir.
   if Champ='TEX_DEBSLIP' then                                                           //Demarre impresion en facturette
      Begin
      if PGIAsk(MC_MsgErrDefaut(1067),LPMessageCap)=mrYes then
         Begin
         Result:=FPrinter.Sqce[spActivateSlip,True] ;
         InFacturette:=TRUE ;
         End else FProgress.ForceCancel ;
      End else
   if Champ='TEX_FINSLIP' then                                                           //Fin impresion en Facturette
      Begin
      Result:=FPrinter.Sqce[spDesactivateSlip,True] ;
      InFacturette:=FALSE ;
      End else
   if Champ='TEX_DEBVAL'  then                                                           //Demare impresion validation
      Begin
      if PGIAsk(MC_MsgErrDefaut(1067),LPMessageCap)=mrYes then
         Begin
         Result:=FPrinter.Sqce[spActivateValidation,True] ;
         InValidation:=TRUE ;
         End else FProgress.ForceCancel ;
      End else
   if Champ='TEX_FINVAL'  then                                                           //Fin impresion validation
      Begin
      Result:=FPrinter.Sqce[spDesactivateValidation,True] ;
      InValidation:=FALSE ;
      End else
   if Champ='TEX_DEBCHQ'  then                                                           //Demare impresion chéque
      Begin
      if PGIAsk(MC_MsgErrDefaut(1067),LPMessageCap)=mrYes then
         Begin
         Result:=FPrinter.Sqce[spDebutCheque,True] ;
         InCheque:=TRUE ;
         End else FProgress.ForceCancel ;
      End else
   if Champ='TEX_FINCHQ'  then                                                           //Fin impresion chéque
      Begin
      Result:=FPrinter.Sqce[spFinCheque,True] ;
      InCheque:=FALSE ;
      End else
   if Champ='TEX_IMPBMP'  then Result:=FPrinter.Sqce[spPrintBmp,True] else               //Imprime l'image chargée dans l'imprimante
   if Champ='TEX_DEBLNSP' then Result:=FPrinter.Sqce[spActivateInterligne,True] else     //Augmente l'interligne
   if Champ='TEX_FINLNSP' then Result:=FPrinter.Sqce[spDesactivateInterligne,True] else  //Revient à l'interligne standard
   ;
{$ENDIF EXPORTASCII}
//XMG 30/03/04 fin
   End ;
Lg.Value:=Result ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.BCloseClick(Sender: TObject);
begin
ModalResult:=mrOk ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.FormClose(Sender: TObject; var Action: TCloseAction);
begin
BClose.Click ;   //Pour assurer la bonne sortie de la fenêtre
end;
////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPPreView.FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
begin
if Key=VK_RETURN then Key:=VK_TAB ;
if Key=VK_VALIDE then Begin BImprimer.Click ; Key:=0 ; End else
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPImprime.Create ( AOwner : TComponent ) ;
Begin
Inherited Create(AOwner) ;
FProgress:=TFLPProgress.Create(Self) ;
FForm:=TForm.Create(Self) ;
Pixel:=TLPPixel.Create(FForm) ;
LPBase:=TLPBase.Create(FForm) ;
LPBase.Parent:=FForm ;
Fillchar(LPParam,Sizeof(TLPParam),#0) ;
QPrin:=nil ;
BandeCourant:=Nil ;
LPParam:=TLPParam.Create(Self) ;
inPreView:=FALSE ;
FPreView:=nil ;
LignePage:=-1 ;
MaxLignePage:=-1 ;
TicketCoupe:=FALSE ;
InFacturette:=FALSE ;
InValidation:=FALSE ;
InCheque:=FALSE ;
UneTOB:=Nil ;
TobRupt:=TOB.Create('RUPTURE',nil,-1) ;
FSQLBased:=FALSE ;
IdxTOB:=-1 ;
Totaliser:=TOB.Create('TOTAUX_S',Nil,-1) ;
FRuptures:=Copy(FRuptures,0,0) ;
FNbRuptures:=-1 ;
isImpControl:=FALSE ;
FWhereLocal:='' ;
ChampsRequete:=TStringLIst.Create ;
ChampsAliases:=TStringList.Create ; //XMG 07/04/04
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
destructor TLPImprime.Destroy ;
Begin
//XMG 04/07/07 début
VideStringListe(ChampsAliases) ;
ChampsAliases.Free ;
//XMG 07/04/04 fin
VideStringliste(ChampsRequete) ;
ChampsRequete.Free ;
TOBRupt.Free ;
Totaliser.ClearDetail ;
Totaliser.free ;
If QPrin<>nil then Begin Ferme(QPrin) ; Qprin:=nil ; End ; //QPrin.Free ;
if FPreView<>nil then FPreview.Free ;
FForm.Free ;
FProgress.Free ;
Inherited destroy ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.VerifParam ( Var Err : TMC_Err) : Boolean ;
Begin
err.code:=0 ; Err.libelle:='' ;
//Debut verifications sur le paramètrage de l'état (??)
//fin verifications ...
MaxLignePage:=LpBase.Height div Lpbase.tailleligne ;
FSQLBased:=(Trim(LPParam.SQL)<>'') ;
Result:=(Err.Code=0) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.GetLastError       : TImprimErreurs ;
Begin
Result:=ieUnknown ;
if Not Assigned(FPreview) then exit ;
Result:=FPreview.LastError ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.ChargeChampsrequete( Champ : String ) : Variant ;
var trouve : Boolean ;
Begin
Result:=#0 ; Champ:=Uppercase(Trim(Champ)) ;
if Copy(Champ,1,9)='SQL_WHERE' then else
if Copy(Champ,1,3)='SO_'       then else
if IsSpecialChampTexte(Champ)  then else
if Copy(Champ,1,4)='SYS_'      then else
   Begin
   Trouve:=FALSE ;
   if (Assigned(UneTOB)) then
      Begin
      if (UneTOB.FieldExists(Champ)) then Trouve:=TRUE else
         if (UneTob.Detail.count>0) and (UneTOB.Detail[0].FieldExists(Champ)) then Trouve:=TRUE ;
      end ;
   if (Not Trouve) then AjouteChampRequete(Champ) ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
//XMG 25/06/03 début
Function TLPImprime.GereGFormule ( Formule : String ; Var BienPasse : Boolean ) : String ;
var LookFields : TFLoadFormule ;
Begin
ErrorFormule:='' ;
LookFields:=PrendChamp ; if IsVerifSQL then LookFields:=ChargeChampsRequete ;
Result:=vstring(GFormule(Formule,LookFields,nil,1)) ;
BienPasse:=(ErrorFormule='') ;
if (not BienPasse) and (V_PGI.Debug) then Debug('-->Erreur en Formule -->'+Formule+#13+'-->'+ErrorFormule+'<--') ;
End ;
//XMG 25/06/03 fin
//////////////////////////////////////////////////////////////////////////////////
Procedure TLPImprime.RechercheChampsRequete ( Parent : TWinControl ) ;
Var i             : Integer ;
    Formule,Bidon : String ; //XMG 25/06/03
    ErreurF       : Boolean ; //XMG 25/06/03
Begin
If assigned(parent) then
   Begin
   For i:=0 to Parent.ControlCount-1 do
     Begin
     if Parent.Controls[i] is TLPBande then
        Begin
        if (trim(TLPBande(Parent.Controls[i]).SQL)<>'') and (TLPBande(Parent.Controls[i]).TesterSQL) then
           Begin
           Formule:=uppercase(trim(TLPBande(Parent.Controls[i]).SQL)) ;
           //XMG 25/06/03 début
           Formule:=VString(GereGFormule(Formule,ErreurF)) ;
           //Formule:=vstring(GFormule(Formule,ChargeChampsRequete,nil,1)) ;
           //XMG 25/06/03 fin
           if Pos('@',Formule)>0 then TraiteFonction(Formule) ; //XMG 25/06/03
           TraiteFormule('',Formule,Bidon) ; //XMG 25/06/03
           End ;
        RechercheChampsRequete(TWinControl(Parent.controls[i])) ;
        End else
     if Parent.Controls[i] is TLPChamp then
        Begin
        Formule:=Uppercase(trim(TLPChamp(Parent.Controls[i]).Libelle)) ;
        TraiteFormule('',Formule,Bidon) ; //XMG 25/06/03
        End else
     if Parent.Controls[i] is TLPImage then ; //rien à faire l'image est charge lors du paramètrage.
     End ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TLPImprime.AjouteChampRequete ( ch : String ) ;
var CS  : TCS ;
    i   : Integer ;
    nom : String;
begin
if not inString(ch,['*','']) then
   Begin

   Nom:=Ch ;
   i:=pos(' AS ',Nom) ;
   if i>0 then delete(Nom,1,i+3) ;
   if ChampsRequete.indexof(Nom)<0 then
      Begin
      CS:=TCS.Create ;
      CS.Nom:='Formule' ;
      CS.Valeur:=Ch ;
      ChampsRequete.AddObject(Nom,CS)
      End ;
   End ;
end ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TLPImprime.ParcourtRequete ( St : String ) ;
var f  : integer ;
    ch : String ;
Begin
St:=Uppercase(Trim(St)) ;
if trim(St)<>'' then
 Begin
 if copy(st,length(St),1)<>',' then St:=St+',' ;
 repeat
   f:=pos(',',st) ;
   if f>0 then
      Begin
      Ch:=trim(Copy(St,1,f-1)) ;
      delete (St,1,f) ; st:=trim(St) ;
      ChargeChampsRequete(ch) ; 
      End ;
 until (f=0) or (trim(St)='') ;
 End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
//XMG 07/04/04 début
Procedure TLPImprime.GardeLesAliases ( NomChamp : String ) ;
var i       : integer ;
    CS      : TCS ;
    Aliases : String ;
Begin
i:=pos(' AS ',NomChamp) ;
if i>0 then
   Begin
   Aliases:=trim(Copy(NomChamp,i+4,length(NomChamp))) ;
   delete(NomChamp,i,length(Nomchamp)) ;
   if ChampsAliases.indexof(NomChamp)<0 then
      Begin
      CS:=TCS.Create ;
      CS.Nom:='Formule' ;
      CS.Valeur:=Aliases;
      ChampsAliases.AddObject(NomChamp,CS)
      End ;
   End ;
End ;
///////////////////////////////////////////////////////////////////////////////////////////////////////
//XMG 07/04/04 fin
Function TLPImprime.verifiechamps(SQL : String ) : String ;
var i,niv                : integer ;
    st,pref,table,unpref : string ;
    Tree                 : TTreeView ;
    Node                 : TTreenode ;
    okdistinct           : Boolean ; 
Begin
isVerifSQL:=TRUE ;
result:=uppercase(trim(SQL)) ;
VideStringListe(ChampsAliases) ; //XMG 07/04/04
if V_PGI.Driver=dbMsAccess then
   Begin
   if copy(result,1,7)='SELECT ' then
      begin
      OkDistinct:=copy(Result,8,9)='DISTINCT ' ;   
      Tree:=TTreeView.Create(nil) ;
      try
        Tree.Parent:=FForm ;
        if RecupSQLParts(Result,Tree.Items) then
           Begin
           Node:=CherchePart(Tree.Items,'SELECT') ;
           St:='' ;
           if Assigned(Node) then
              Begin
              Node:=Node.GetFirstChild ;
              while Assigned(Node) do
                Begin
                if Trim(St)<>'' then St:=St+', ' ;
                St:=St+Node.Text ;
                Node:=Node.getNextSibling ;
                End ;
              End ;
           pref:='' ;
           Node:=CherchePart(Tree.Items,'FROM') ;
           if Assigned(Node) then
              Begin
              Node:=Node.GetFirstChild ;
              niv:=Node.Level ;
              while Assigned(Node) do
                Begin
                table:=Node.Text ;
                while copy(Table,1,1)='(' do delete(Table,1,1) ; //XMG 23/06/03
                if (pos('JOIN',Table)<=0) and (copy(table+' ',1,3)<>'ON ') then
                   Begin
                   unPref:=tabletoprefixe(Table) ;
                   if trim(unpref)<>'' then Pref:=pref+Unpref+';' ;
                   End ;
                Node:=Node.getNext ;
                if (assigned(Node)) and (Node.Level<Niv) then Node:=nil ;
                End ;
              End ;
           End ;
       finally
         VideTree(Tree) ;
         Tree.Items.Clear ;
         Tree.free ;
       End ;
      i:=pos('FROM',Result) ;
      Delete(Result,1,i-1) ;
      VideStringListe(ChampsRequete) ;
      ParcourtRequete(St) ;
      RechercheChampsRequete(LPBase) ;
      St:='' ;
      for i:=0 to ChampsRequete.Count-1 do
        if pos(';'+ExtractPrefixe(TCS(ChampsRequete.Objects[i]).Valeur)+';',';'+Pref)>0 then
           begin
           if trim(St)<>'' then St:=St+', ' ;
           St:=St+TCS(ChampsRequete.Objects[i]).Valeur ;
           GardeLesAliases(TCS(ChampsRequete.Objects[i]).Valeur) ;
           End ;
      if OkDistinct then St:='distinct '+St ;
      St:='select '+St ;
      Result:=St+' '+Result ;
      VideStringListe(ChampsRequete) ;
      end ;
   //XMG 07/04/04 début
   end else
   Begin
   Tree:=TTreeView.Create(nil) ;
   try
     Tree.Parent:=FForm ;
     if RecupSQLParts(Result,Tree.Items) then
        Begin
        Node:=CherchePart(Tree.Items,'SELECT') ;
        St:='' ;
        if Assigned(Node) then
           Begin
           Node:=Node.GetFirstChild ;
           while Assigned(Node) do
             Begin
             GardeLesAliases(Node.Text) ;
             Node:=Node.getNextSibling ;
             End ;
           End ;
        End ;
    finally
      VideTree(Tree) ;
      Tree.Items.Clear ;
      Tree.free ;
    End ;

   //XMG 07/04/04 fin
   End ;
IsVerifSQL:=FALSE ;
End  ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImprime.ForceCancel ;
Begin
If not InPreview then
   Fpreview.CalcelImpression ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImprime.InitPreView ;
Begin
FPreView:=TFLPPreView.Create(Self) ;
FPreView.PrinterPort:=FPort ;
FPreView.PrinterParams:=FParams ;
FPreView.Imprimante:=FImpr ;
FpreView.InitImp:=FInitImp ;
FpreView.InPreview:=InPreview ;
FPreView.InitView(LPBase) ;
FPreView.IsImpControl:=IsImpControl ;
Fpreview.InTestMateriel:=InTestMateriel ;
FPreView.IsExport:=(LPParam.TypeEtat=TypeEtatExportAscii) ; //XMG 02/04/04 'F') ; //XMG 30/03/04
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImprime.RajouteWhere ( Where : String ) ;
var Tree : TTreeView ;
    St   : String ;
    Node : TTreeNode ;
    Q1   : TQuery ;
Begin
FWhereLocal:='' ;
if Copy(Uppercase(Trim(Where)),1,7)='SELECT ' then
   Begin
   try
     where:=VerifieChamps(Where) ;
     Q1:=OpenSQL(Where,TRUE) ;
     Ferme(Q1) ;
     LPParam.SQL:=Where ;
     FSQLBased:=TRUE ;
    except
     LPParam.SQL:='' ;
    End ;
   End else
   if FSQLBased then
      Begin
      Tree:=TTreeView.Create(nil) ;
      try
        Tree.Parent:=FForm ;
        if RecupSQLParts(LPParam.SQL,Tree.Items) then
           Begin
           Node:=CherchePart(Tree.Items,'WHERE') ;
           if Node<>nil then
              Begin
              Node:=Node.GetFirstChild ;
              St:=Node.Text ;
              if St<>'' then St:=St+' AND ' ;
              Node.Text:=St+Where ;
              End else
              Begin
              Node:=InsertPart(Tree.Items,nil,'WHERE') ;
              InsertPart(Tree.Items,Node,Where) ;
              End ;
           LPParam.SQL:=RecupSQLFromParts(Tree.Items) ;
           FWhereLocal:=Where ;
           End ;
       finally
         VideTree(Tree) ;
         Tree.Items.Clear ;
         Tree.free ;
       End ;
      End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.ChercheChamp(Champ : String ; Var Evalue : Boolean ) : String ; //XMG 25/06/03
   ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
   //XMG 25/06/03 début
   Function TrouveChamp( Champ : String ; Bande : TLPBande ; Var R : String ; var Evalue : Boolean ) : Boolean ;
   var i    : integer ;
       Ctrl : TLPChamp ;
   Begin
   Result:=FALSE ; Evalue:=FALSE ;
   if not assigned(Bande) then exit ;
   Champ:=Uppercase(Trim(Champ)) ;
   for i:=0 to Bande.controlcount-1 do
     if Bande.Controls[i] is TLPChamp then
        Begin
        Ctrl:=TLPChamp(Bande.Controls[i]) ;
        if Ctrl.Nom=Champ  then
           Begin
           if Ctrl.Evalue then Begin Evalue:=TRUE ; R:=trim(Ctrl.GetFormatedValue{Value}) ; End ; //XMG 23/06/03
           Result:=TRUE ;
           Break ;
           End ;
        End ;
   End ;
   //XMG 25/06/03 fin
   ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Var Bande  : TLPBande ;
    i      : TLPTypeBandes ;
    Trouve : Boolean ;
    numDet : Integer ;
Begin
Result:='' ; i:=low(TLPTypeBandes) ; Trouve:=FALSE ;
NumDet:=0 ;
While (Not Trouve) and (i<=High(TLPTypeBandes)) do
  Begin
  if i in [lbSubEntete,lbSubDetail,lbSubPied,lbenteteRupt,lbPiedRupt] then inc(NumDet) ;
  Bande:=LPBase.ChoixBande(i,Numdet) ;
  Trouve:=TrouveChamp(Champ,Bande,Result,Evalue) ;
  if (not (i in [lbSubEntete,lbSubDetail,lbsubPied,lbenteteRupt,lbPiedRupt])) or (NumDet>=LPBase.NumMaxSubdetail(i)) then
     Begin
     inc(i) ;
     if NumDet>0 then NumDet:=0 ;
     End ;
  End ;
End ;
//XMG 07/04/04 début
///////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.RenvoieNomchampRequete(Champ : String ) : String ;
var idx : Integer ;
Begin
Result:=Champ ;
idx:=ChampsAliases.Indexof(Champ) ;
if Idx>-1 then
   Result:=TCS(ChampsAliases.objects[idx]).Valeur ;
End ;
//XMG 07/04/04 fin
//////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.PrendChamp ( Champ : String ) : Variant ;
Var F      : TField ;
    Trouve : Boolean ;
{$IFDEF EAGLCLIENT}
    IndexTob : integer ;
    State: TDataSetState;
{$ENDIF}                    
Begin
Result:=#0 ; Champ:=Uppercase(Trim(Champ)) ; F:=Nil ; Trouve:=FALSE ;
if (BandeCourant.TypeBande=lbPiedRupt) then Trouve:=(TobRupt.FieldExists(Champ)) else
   Begin
   if (BandeCourant.typeBande in [lbSubEntete,lbSubDetail,lbSubPied]) and (QsbDet<>nil) then
      Begin
      F:=QSbDet.FindField(RenvoieNomchampRequete(Champ)) ; //XMG 07/04/04
      Trouve:=Assigned(F) ;
      End ;
   if (not Trouve) and (FSQLBased) then
      Begin
{$IFDEF EAGLCLIENT}
      if Assigned(QPrin) then
         begin
         IndexTob:=Qprin.CurrentFilleIndex;
         State:=QPrin.State;
         if QPrin.EOF then QPrin.Open ;
         end else
         begin
         IndexTob:=-1 ;
         State:=dsInactive ;
         end ;
{$ENDIF}
      F:=QPrin.FindField(RenvoieNomchampRequete(Champ)) ; //XMG 07/04/04
{$IFDEF EAGLCLIENT}
      if Assigned(QPrin) then
         begin
         Qprin.Seek(IndexTob);
         QPrin.State:=State ;
         end;
{$ENDIF}
      Trouve:=Assigned(F) ;
      End ;
   End ;
if Trouve then
   Begin
   if assigned(F) then Result:=F.AsVariant else Result:=TOBRupt.GetValue(Champ) ;
   if ChamptoType(Champ)='BLOB' then Result:=RtfToString(Result) ;   //XMG 23/06/03 souci avec les RichEdits...
   End else
   Begin
//XMG 25/06/03 début
   if Copy(Champ,1,9)='SQL_WHERE' then
      Begin
      Trouve:=TRUE ;
      Result:=FWhereLocal ;
      End else
   if Copy(Champ,1,3)='SO_' then
      Begin
      Trouve:=TRUE ;
      Result:=GetParamSoc(Champ) ;
      End else
   //Les variables Imprimante (TEX_) sont evaluées au moment de l'impresion
   if IsSpecialChampTexte(Champ) then
      Begin
      Result:=Champ ;
      Trouve:=TRUE ;
      TicketCoupe:=(TicketCoupe) or (pos(';'+Champ+';',';TEX_PARTIEL;TEX_TOTAL;TEX_DEBSLIP;TEX_FINSLIP;')>0) ;
      if (not InFacturette) and (Champ='TEX_DEBSLIP') then InFacturette:=TRUE else
      if (InFacturette)     and (Champ='TEX_FINSLIP') then InFacturette:=FALSE ;
      if (not InValidation) and (Champ='TEX_DEBVAL')  then InValidation:=TRUE else
      if (InValidation)     and (Champ='TEX_FINVAL')  then InValidation:=FALSE ;
      if (not InCheque)     and (Champ='TEX_DEBCHQ')  then InCheque:=TRUE else
      if (InCheque)         and (Champ='TEX_FINCHQ')  then InCheque:=FALSE ;
      end else
   if Copy(Champ,1,4)='SYS_' then
      Begin//Variables Generales
      if Champ='SYS_NATURE' then
         Begin
         Trouve:=TRUE ;
         Result:= LPParam.NatEtat ;
         End else
      if Champ='SYS_MODELE' then
         Begin
         Trouve:=TRUE ;
         Result:= LPParam.CodeEtat  ;
         End else
      if Champ='SYS_DATEENTREE' then
         Begin
         Trouve:=TRUE ;
         Result:= V_PGI.DateEntree  ;
         End else
      if Champ='SYS_DATEVERSION' then
         Begin
         Trouve:=TRUE ;
         Result:= V_PGI.DateVersion  ;
         End else
      if Champ='SYS_PROGVERSION' then
         Begin
         Trouve:=TRUE ;
         Result:= V_PGI.NumVersion  ;
         End else
      if Champ='SYS_TITRE' then
         Begin
         Trouve:=TRUE ;
         Result:= LPParam.Libelle  ;
         End else
      if Champ='SYS_COPYRIGHT' then
         Begin
         Trouve:=TRUE ;
         Result:= Copyright  ;
         End else
      if Champ='SYS_HALLEY' then
         Begin
         Trouve:=TRUE ;
         Result:= TitreHalley  ;
         End else
      if Champ='SYS_DATE' then
         Begin
         Trouve:=TRUE ;
         Result:= Date  ;
         End else
      if (Champ='SYS_TIME') or (Champ='SYS_HEURE') then
         Begin
         Trouve:=TRUE ;
         Result:= Now  ;
         End else
      if (Champ='SYS_UTILISATEUR') or (Champ='SYS_USER') then
         Begin
         Trouve:=TRUE ;
         Result:= V_PGI.User  ;
         End else
      if (Champ='SYS_NOMUTILISATEUR') or (Champ='SYS_USERNAME') then
         Begin
         Trouve:=TRUE ;
         Result:= V_PGI.UserName  ;
         End else
      if (Champ='SYS_CODESOCIETE') or (Champ='SYS_COMPANYCODE') then
         Begin
         Trouve:=TRUE ;
         Result:= V_PGI.CodeSociete  ;
         End else
      if (Champ='SYS_NOMSOCIETE') or (Champ='SYS_COMPANYNAME') then
         Begin
         Trouve:=TRUE ;
         Result:= V_PGI.NomSociete  ;
         End else
         ;
//XMG 25/06/03 fin
      End else
      Begin
      Trouve:=FALSE ;
      if (Assigned(UneTOB)) then
         Begin
         if (UneTOB.FieldExists(Champ)) then Begin Trouve:=TRUE ; Result:=UneTOB.GetValue(Champ) ; End else
            if (IdxTOB<UneTob.Detail.count) and (UneTOB.Detail[IdxTOB].FieldExists(Champ)) then Begin Trouve:=TRUE ; Result:=UneTOB.Detail[IdxTOB].GetValue(Champ) ; End ;
         end ;
      if Not Trouve then Result:=ChercheChamp(Champ,Trouve) ; //Champs du même état //on croise la notion trouve avec evalue.....
      End ;
   End ;
if Not Trouve then ErrorFormule:=ErrorFormule+Champ+';' ; //XMG 25/06/03
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.NouvellePage : Boolean ;
Begin
Result:=FpreView.NewPage ;
LignePage:=-1+ord(Result) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.SautLignes ( NLignes : Integer ) : Boolean ;
Begin
result:=FPreView.Addlignes(NLignes) ;
inc(LignePage,NLignes*ord(Result)) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.InsertChampTEX ( QueChamp : String ) : Boolean ;
Var Ctrl : TLPChamp ;
Begin
Result:=FALSE ;
QueChamp:=Uppercase(Trim(QueChamp)) ;
if IsSpecialChampTexte(QueChamp) then
   Begin
   Ctrl:=TLPChamp.Create(LPBase) ;
   Ctrl.Parent:=LPBase ;
   Ctrl.libelle:='['+QueChamp+']' ;
   Ctrl.Ligne:=LignePage ;
   Ctrl.PrinterVisible:=FALSE ;
   Ctrl.colonne:=1 ;
   Ctrl.NbrChars:=10 ;
   Result:=(FPreView.AddChamp(-2 ,Ctrl)) ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.WritelnLP ( St : String ) : Boolean ;
Var Ctrl : TLPChamp ;
Begin
Ctrl:=TLPChamp.Create(LPBase) ;
Ctrl.Parent:=LPBase ;
Ctrl.Value:=TraduireMemoire(St) ;
Ctrl.Ligne:=LignePage ;
Ctrl.colonne:=1 ;
Ctrl.NbrChars:=length(St) ;
Result:=(FPreView.AddChamp(LignePage+1,Ctrl)) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPIMprime.ChargeparamsQ ( SQLOri : String ) : String ; //XMG 25/06/03
Var Val                  : Variant ;
    Champ,Prefixe,Typ    : String ;
    OldDeb,DebChp,FinChp : integer;
    St,ValChp            : string;
    ErreurF              : Boolean ; //XMG 25/06/03
Begin
Result:=SQLOri ;
if trim(Result)<>'' then
   Begin
   Result:=FindEtReplace(FindEtReplace(Result,#13,' ',TRUE),#10,'',TRUE) ;
   DebChp:=Pos('ORDER BY',UpperCase(Result));
   if DebChp>0 then Result:=Copy(Result,1,DebChp-1)+' '+Copy(Result,DebChp,Length(Result)); //???
   OldDeb:=0 ;
   repeat
     DebChp:=Pos('=:',Copy(Result,OldDeb+1,Length(Result))) ;
     if DebChp>0 then
        begin
        St:=Copy(Result,OldDeb+DebChp+2,Length(Result));
        FinChp:=Pos(' ',St);
        if FinChp=0 then FinChp:=Length(St)+1;
        {j:=Pos(#$D,St); if j=0 then j:=FinChp;
        FinChp:=Min(FinChp,j);}
        Champ:=Copy(St,1,FinChp-1);
        if isVerifSQL then Begin ChargeChampsRequete(Champ) ; OldDeb:=OldDeb+DebChp+1 ; End else //XMG 25/06/03
           Begin
           //XMG 25/06/03 début
           Val:=GereGFormule('['+Champ+']',ErreurF) ;  //**
           //Val:=GFormule('['+Champ+']',PrendChamp,nil,1) ;
           //XMG 25/06/03 fin
           Prefixe:=ExtractPrefixe(Champ) ;
           if Prefixe='' then
              Begin
              case vartype(Val) of
                 varSmallint : Typ:='INTEGER' ;
                 varInteger  : Typ:='SMALLINT' ;
                 varSingle,
                 varDouble,
                 varCurrency : Typ:='DOUBLE' ;
                 varDate     : Typ:='DATE' ;
                 else          Typ:='CHAR' ;
                 End ;
              End else Typ:=Champtotype(Champ) ;
           if (Typ='INTEGER') or (Typ='SMALLINT') then ValChp:=IntToStr(Val) else
           if (Typ='DOUBLE') or (Typ='RATE')      then ValChp:=strfpoint(Val) else
           if Typ='DATE'                          then ValChp:='"'+USDateTime(VDate(vDouble(Val)))+'"'
                                                  else ValChp:='"'+Vstring(Val)+'"' ;
           Result:=Copy(Result,1,OldDeb+DebChp)+ValChp+Copy(St,FinChp,Length(St));
           End ;
        End ;
   until DebChp<=0 ;
   End ;
end ;
/////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.TraiteFonction ( var Formule : String ) : Boolean ; //XMG 25/06/03
var j,i,pp,pl : integer ;
    Sf,sp,st  : String ;
    //kk        : TFLoadFormule ; //XMG 25/06/03
Begin
//XMG 25/06/03 début
Result:=pos('@',Formule)<=0 ;
if not result then
   repeat
     i:=pos('@',Formule) ;
     if i>0 then
        Begin
        St:=copy(Formule,i+1,length(Formule)) ;
        j:=pos('(',St) ; if j<=0 then j:=Length(St) ;
        Sf:=Copy(St,1,j-1) ;
        Sp:=Copy(St,j,length(St)) ;
        pp:=0 ; Pl:=0 ;
        for j:=1 to length(Sp) do
           Begin
           inc(pp,ord(sp[j]='(')) ;
           dec(pp,ord(sp[j]=')')) ;
           if pp=0 then
              Begin
              pl:=j ;
              break ;
              End ;
           End ;
        if pl>0 then
           Begin
           Sp:=Copy(sp,2,pl-2)+';;;;;;;;;;' ;
           Sp:=VString(GereGFormule(sp,Result)) ; //**
           {kk:=PrendChamp ; if IsRecherche then kk:=ChargeChampsRequete ;
           Sp:=vstring(GFormule(Sp,kk,nil,1)) ;}
           if Result then
              Begin
              St:='' ;
              if (Not IsVerifSQL) and (assigned(ProcCalcEdt)) then St:=vstring(ProcCalcEdt(sf,sp)) ; //XMG 25/06/03
              delete(Formule,i,pl+1+length(sf)) ;
              Insert(St,Formule,i) ;
              End ;
           End ;
        End ;
   until (not Result) or (i<=0) ;
//XMG 25/06/03 fin
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.TraiteFormule ( LaMasque,Formule : String ; Var Resultat : String ) : Boolean ; //XMG 25/06/03
//XMG 25/06/03 début
Var Q1         : TQuery ;
    UnTot      : TOB ;
    ini        : integer ;
    //kk         : TFLoadFormule ;
Begin
Resultat:=Formule ;
Formule:=Uppercase(trim(Formule)) ;
if pos('@',Formule)>0 then
   begin
   Result:=TraiteFonction(Formule) ;
   if pos('@',Resultat)=1 then Resultat:=Formule ;
   end ;
if ('SELECT'=Copy(Formule,1,6)) then
   Begin
   Resultat:='' ;
   Formule:=VString(GereGFormule(Formule,Result)) ; //**
   //kk:=PrendChamp ; if IsRecherche then kk:=ChargeChampsRequete ;
   //Formule:=vstring(GFormule(Formule,kk,nil,1)) ;
   if Result then
      Begin
      Formule:=findetreplace(formule,''#0'','''',TRUE) ;
      Formule:=ChargeParamsQ(Formule) ;
      if (Not IsVerifSQL) then
         try
           Q1:=OpenSQL(Formule,TRUE) ;
           if not Q1.Eof then Resultat:=Q1.Fields[0].asString ;
         Finally
           if assigned(Q1) then Ferme(Q1) ;
         End ;
      End ;
   End else
if Absol=Copy(Formule,1,Length(Absol)) then
   Begin
   Formule:=Copy(Formule,length(Absol)+1,Length(Formule)-Length(Absol)-1) ;
   Result:=TraiteFormule(LaMasque,Formule,Resultat) ;
   Resultat:=floattostr(Abs(vDouble(Resultat))) ;
   End else
if Somme=Copy(Formule,1,Length(Somme)) then
   Begin
   Result:=TRUE ;
   if Not IsVerifSQL then
      Begin
      Formule:=Copy(Formule,length(Somme)+1,Length(Formule)-Length(Somme)-1) ;
      Formule:=ReadtokenSt(Formule) ;
      UnTot:=Totaliser.FindFirst(['BANDE','CHAMP'],[BandeCourant.Titre,Formule],FALSE) ;
      if Assigned(UnTot) then
         Begin
         Resultat:=formatfloat('###0.00',Vdouble(UnTot.GetValue('RESULTAT'))) ;
         if VString(UnTot.GetValue('RAZ'))='X' then UnTot.PutValue('RESULTAT',Zero) ;
         End ;
      End ;
   End else
if (Length(Formule) > 0) and (Formule[1] in ['{','(','[']) then
   Begin
   if Formule[1]='(' then Formule[1]:='{' ;
   if Formule[length(Formule)]=')' then Formule[length(Formule)]:='}' ;
   if Not IsVerifSQL then
      Begin
      if LaMasque<>'' then
         Begin
         ini:=1+ord(Formule[1]<>'[') ;
         Formule:='{"'+LaMasque+'"'+Copy(Formule,ini,length(Formule)) ;
         if ini=1 then Formule:=Formule+'}' ;
         End ;
      end ;
   Resultat:=VString(GereGFormule(Formule,Result)) ;
   //kk:=PrendChamp ; if IsRecherche then kk:=ChargeChampsRequete ;
   //Resultat:=vstring(GFormule(Formule,kk,nil,1)) ;
   End else
   Begin //gestion des textes fixes
   if (LPParam.Langue<>V_PGI.LanguePrinc) then Resultat:=TraduireMemoire(Resultat) ;
   Result:=TRUE ;
   End ;
//XMG 25/06/03 fin
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImprime.ChargeVariables ;
Var i,j             : Integer ;
    Ctrl            : TLPChamp ;
    Formule,Nom,Val : String ; //XMG 25/06/03
    UnTot           : TOB ;
    TousE           : Boolean ;
Begin
BandeCourant.InitialiseChamps(IsEof) ;
if (BandeCourant.TypeBande=lbpied) and (not isEof) then exit ;
j:=0 ; tousE:=TRUE ;
repeat
   j:=j*ord(not TousE)+ord(not TousE) ;
   For i:=0 to BandeCourant.ControlCount-1 do
      if (TControl(BandeCourant.Controls[i]) is TLPChamp) and (not TLPChamp(BandeCourant.Controls[i]).Evalue) then
         Begin
         Ctrl:=TLPChamp(BandeCourant.Controls[i]) ;
         Formule:=Ctrl.Libelle ;
         if (length(Formule)>0) and ((Formule[1]<>'#') or (Copy(Formule,1,2)='##')) then
            Begin
            if Copy(Formule,1,2)='##' then Delete(Formule,1,1) ;
            if TraiteFormule(Ctrl.GetMaskFormule,Formule,Val) then Ctrl.Value:=Val ; //XMG 25/06/03
            End ;
         if Ctrl.Value<>'' then
            Begin
            //Ctrl.FormatResult ;
            if IsNumeric(Ctrl.Value) then
               Begin
               Nom:=Uppercase(Trim(Ctrl.Nom)) ;
               UnTot:=Totaliser.FindFirst(['CHAMP'],[Nom],FALSE) ;
               While Assigned(UnTot) do
                  Begin
                  UnTot.PutValue('RESULTAT',VDouble(UnTot.GetValue('RESULTAT'))+Valeur(Ctrl.Value)) ;
                  UnTot:=Totaliser.FindNext(['CHAMP'],[Nom],FALSE) ;
                  End ;
               End ;
            End ; //if Ctrl.Value....
         End ;  //if TControl(BandeCourant.Controls....
   TousE:=BandeCourant.TousEvalues ;
until (TousE) or (j>=2) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.TestCondition ( Bande : TLPBande ) : Boolean ;
var Val : String ;
Begin
If trim(Bande.Condition)='' then Begin Result:=TRUE ; Exit ; End ;
//XMG 25/06/03 début
TraiteFormule('###0',Bande.Condition,Val) ;
Result:=valeuri(Val)<>0 ;  //=0 FALSE <>0 TRUE
//XMG 25/06/03 fin
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.TraiteBande ( Bande : TLPBande ; Var Err : TMC_Err ) : Boolean ;
Var jj,Ligne : Integer ;
Begin
Result:=FALSE ;
if Bande=nil then Begin Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1028) ; Exit ; End ;
if Bande<>BandeCourant then BandeCourant:=Bande ;
ChargeVariables ;
if (TestCondition(BandeCourant)) and (BandeCourant.Visible) then
   Begin
   For jj:=0 to BandeCourant.ControlCount-1 do
       Begin
       Ligne:=-1 ;
       if TObject(BandeCourant.Controls[jj]) is TLPChamp then Ligne:=TLPChamp(BandeCourant.Controls[jj]).Ligne else
       if TObject(BandeCourant.Controls[jj]) is TLPImage then Ligne:=TLPImage(BandeCourant.Controls[jj]).Ligne ;
       if Ligne=-1 then break ;
       Result:=(FPreView.AddChamp(LignePage+Ligne,TControl(BandeCourant.Controls[jj]))) ;
       if Not Result then Break ;
       End ;
   inc(LignePage,Bande.Lignes*ord(Result)) ;
   End ;
Result:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.TraiteSubBande ( NombreSubBnd : Integer ; TypeSubBnd : TLPTypeBandes ; Var Err : TMC_Err ) : Boolean ;
Var i,j     : Integer ;
    sbBnd   : TLPBande ;
    SQL     : String ;
    ErreurF : Boolean ; //XMG 25/06/03
Begin
Result:=FALSE ; j:=0 ;
for i:=1 to NombreSubBnd do
    Begin
    SbBnd:=LpBase.choixBande(TypeSubBnd,i) ; QsbDet:=NIl ;
    if SbBnd<>nil then
       Begin
       inc(j) ;
       If (trim(sbBnd.SQL)<>'') then
          Begin
          SQL:=uppercase(trim(sbBnd.SQL)) ;
          //XMG 25/06/03 début
          SQL:=VString(GereGFormule(SQL,ErreurF)) ; //**
          //SQL:=vstring(GFormule(SQL,PrendChamp,nil,1)) ;
          //XMG 25/06/03 fin
          if (Pos('@',SQL)<=0) or (TraiteFonction(SQL)) then
             Begin
             QsbDet:=OpenSQL(ChargeParamsQ(SQL),TRUE) ; //XMG 25/06/03
             while (Not QsbDet.Eof) and ((result) or (j=1)) and (Not FProgress.Canceled) do
               Begin
               Result:=(TraiteBande(SbBnd,Err)) and ((Result) or (j=1)) ;
               QsbDet.Next ;
               End ;
             Result:=(Result) or ((QsbDet.eof) and (QsbDet.Bof)) ;
             Ferme(QsbDet) ; QSbDet:=Nil ;
             End ;
          End else Result:=(TraiteBande(SbBnd,Err)) and ((Result) or (j=1)) ;
       End ; //if SbBnd<>nil...
    End ; //For i:=...
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImprime.RechercheTotaux ;
Var i,j    : Integer ;
    Bnd    : TLPBande ;
    Ch     : TLPChamp ;
    St,Raz : String ;
    UnTot  : TOB ;
Begin
for i:=0 to LPBase.ControlCount-1 do
   if LPBase.Controls[i] is TLPBande then
      Begin
      Bnd:=TLPBande(LPBase.Controls[i]) ;
      For j:=0 to Bnd.ControlCount-1 do
          if Bnd.Controls[j] is TLPChamp then
             Begin
             Ch:=TLPChamp(Bnd.Controls[j]) ;
             St:=Uppercase(Trim(Ch.Libelle)) ;
             if Copy(St,1,Length(Somme))=Somme then
                Begin
                Delete(St,1,length(Somme)) ; Delete(St,length(St),1) ; St:=St+';' ;
                Raz:='X' ;
                if NbCarInString(St,';')=2 then Raz:=gtfs(St+';',';',2) ;
                St:=ReadTokenSt(St) ;
                UnTot:=Totaliser.FindFirst(['BANDE','CHAMP'],[Bnd.Titre,St],FALSE) ;
                if Not assigned(UnTot) then
                   Begin
                   UnTot:=TOB.Create('UN_TOTAL',Totaliser,-1) ;
                   with UnTot do
                     Begin
                     AddChampSup('BANDE',FALSE)    ; PutValue('BANDE',Bnd.Titre) ;
                     AddChampSup('CHAMP',FALSE)    ; PutValue('CHAMP',St) ;
                     AddChampSup('RESULTAT',FALSE) ; PutValue('RESULTAT',Zero) ;
                     AddChampSup('RAZ',FALSE)      ; PutValue('RAZ',TrueFalseSt(pos(';'+Raz+';',';X;TRUE;VRAI;OUI;')>0)) ;
                     End ;
                   End ;
                End ;
             End ;
      End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImprime.ChargeRuptures ;
Var Rupt : String ;
    i    : Integer ;
Begin
FRuptures:=Copy(FRuptures,0,0) ;
Rupt:=FindEtReplace(LPBase.Ruptures,' ;','',TRUE) ;
if (Length(Rupt)>1) and (Copy(Rupt,length(Rupt),1)<>';') then Rupt:=Rupt+';' ;
FNbRuptures:=minintValue([LP_NbMaxRuptures,nbCarInString(Rupt,';')]) ;
SetLength(FRuptures,FNbRuptures) ;
For i:=0 to FNbRuptures-1 do
  Begin
  FRuptures[i].Champ:=gtfs(Rupt,';',i+1) ;
  FRuptures[i].Valeur:=Unassigned ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.TesteRuptures  : Integer ;
Var Val     : Variant ;
    i       : Integer ;
    ErreurF : Boolean ; //XMG 25/06/03
Begin
Result:=-1 ;
for i:=High(FRuptures) downto low(FRuptures) do
  Begin
  //XMG 25/06/03 début
  Val:=GereGFormule('['+FRuptures[i].Champ+']',ErreurF) ; //**
  //Val:=GFormule('['+FRuptures[i].Champ+']',PrendChamp,nil,1) ;
  //XMG 25/06/03 fin
  if (VarisEmpty(FRuptures[i].Valeur)) or (Val<>FRuptures[i].Valeur) then
     Begin
     Result:=i+1 ;
     FRuptures[i].Valeur:=Val ;
     End ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.TraiteDesRupt(DesRupt : Integer ; Tb : TLPTypeBandes ; Var Err : TMC_Err ; Final : Boolean = FALSE ) : Boolean ;
var Bnd : TLPBande ;
    Ok  : Boolean ;
    n   : Integer ;
Begin
Ok:=TRUE ;
if tb=lbPiedRupt then n:=1
   else n:=DesRupt ;
while (Ok) and (desrupt<=FNbRuptures) do
  Begin
  Bnd:=LPBase.ChoixBande(tb,n) ;
  if assigned(Bnd) then Ok:=traiteBande(Bnd,Err) ;
  Inc(DesRupt) ; inc(n) ;
  End ;
Result:=ok ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.Imprime (Var Err : TMC_Err ) : Boolean ;
Var BndTop,BndPied,Bande               : TLPBande ;
    Ltop,LBottom,LDetail,Lpied,desRupt : Integer ;
    NSbDtl                             : Array[TLPTypeBandes] of integer ;
Begin
Result:=FALSE ;
Fillchar(Err,sizeof(TMC_Err),#0) ;
//Calcul du marge haut
Pixel.AsPixel:=LpBase.MargeTop+1 ;
LTop:=Pixel.asChar[TRUE,LPTailleDef] ; //-1 ;
//Calcul du marge bas
Pixel.AsPixel:=LpBase.MargeBottom+1 ;
LBottom:=Pixel.asChar[TRUE,LPTailleDef] ;
//Recherche des bandes
BndPied:=LPBase.ChoixBande(lbPied) ;
BndTop:=LPBase.ChoixBande(lbEntete) ;
Bande:=LpBase.ChoixBande(lbDetail) ;
//Taille en lignes de la bande detail (+ bandes subdetail)
LDetail:=LPBase.Tailleensemblebande(lbDetail) ;
//Combien subdetailles
Fillchar(NsbDtl,sizeof(NSBDtl),#0) ;
NSbDtl[lbEntete]:=LPBase.NumMaxSubdetail(lbSubEntete) ;
NSbDtl[lbDetail]:=LPBase.NumMaxSubdetail(lbSubDetail) ;
NSbDtl[lbPied]:=LPBase.NumMaxSubdetail(lbSubPied) ;
//Calcul du nombre de lignes total du pied
lPied:=LBottom ;
if BndPied<>nil then inc(LPied,LPBase.tailleensembleBande(lbPied)) ;
//Recherche des champ à totaliser
Recherchetotaux ;
//Gestion des Ruptures
ChargeRuptures ;        
//Impresion du document
if FSQLBased then
  Begin
  QPrin.Open ;
  End ;
initPreView ;
TicketCoupe:=(not LPBase.Rouleau) ;
InFacturette:=FALSE ;
InValidation:=FALSE ;
InCheque:=FALSE ;
idxTOB:=0 ;
repeat
  if FSQLBased then isEof:=(Qprin.Eof) else
    if Assigned(UneTOB) then IsEof:=((UneTob.Detail.Count=0) and (IdxTOB>0)) or ((UneTob.Detail.Count>0) and (IdxTOB>UneTOB.Detail.Count-1)) else
       Begin
       IsEOF:=TRUE ;
       FProgress.ForceCancel ;
       End ;
  if ((IsEof) or ((LignePage>(MaxLignePage-LPied-lDetail)) and (not LPBase.Rouleau)) or (LignePage=-1)) then
    Begin
    if ((LignePage>(MaxLignePage-LPied-lDetail)) or (IsEof)) and (not FProgress.Canceled) then
       Begin  //impresion du pied de la page
       if (Result) and (isEof) and (FNbRuptures>0) then result:=TraiteDesRupt(1,lbPiedRupt,err,TRUE) ;
       if (Result) and (LignePage<(MaxLignePage-LPied)) and (not LPBase.Rouleau) then Result:=SautLignes(MaxlignePage-LPied-LignePage) ;
       if (Result) and (BndPied<>nil) then Result:=TraiteBande(BndPied,Err) ;
       if (Result) and (NSbDtl[LbPied]>0) then Result:=TraiteSubBande(NsbDtl[lbPied],lbSubPied,Err) ;
       if Result then Result:=SautLignes(LBottom) ;
       End ;
    if (LignePage<>-99) and (Not IsEof) and (not FProgress.Canceled) then
       Begin
       if ((Result) or (LignePage=-1)) then Result:=Nouvellepage ;
       if (Result) then Result:=SautLignes(LTop) ;
       if (Result) and (BndTop<>nil) then Result:=TraiteBande(BndTop,Err) ;
       if (Result) and (NSbDtl[LbEntete]>0) then Result:=TraiteSubBande(NsbDtl[lbEntete],lbSubEntete,Err) ;
       End ;
    End ;
 if Not IsEof then
    Begin
    FProgress.Setvalue(FProgress.Progress+1) ;
    //Teste ruptures
    if (FnbRuptures>0) and (Result) then
       Begin
       DesRupt:=TesteRuptures ;
       //Pied Rupt
       if (IdxTOB>0) and (DesRupt>0) then Result:=TraiteDesRupt(DesRupt,lbPiedRupt,err) ;
       //Entete Rupt
       if (Result) and (DesRupt>0) then Result:=TraiteDesRupt(DesRupt,lbEnteteRupt,err)
       End ;
    if (Result) and (Bande<>nil) and (Not FProgress.Canceled) then Result:=TraiteBande(Bande,Err) ;
    if (Result) and (NSbDtl[LbDetail]>0) then Result:=TraiteSubBande(NsbDtl[lbDetail],lbSubDetail,Err) ;
    if FnbRuptures>0 then
       Begin
       if FSQLBased then TobRupt.SelectDB('',QPrin,FALSE)
          else TOBRupt.Dupliquer(UneTob.Detail[IdxTOB],FALSE,TRUE) ;
       End ;
    if FSQLBased then QPrin.Next ;
    inc(IdxTOB) ;
    End ; //If not IsEof...
until (Not Result) or (Fprogress.Canceled) or (IsEof) ;
if Fprogress.Canceled then
   Begin
   sautlignes(2) ;
   WriteLnLP(MC_MsgErrDefaut(1031)) ;
   sautlignes(2) ;
   Result:=TRUE ;
   End ;
FProgress.Hide ;
if FSQLBased then QPrin.Close ;
if Not TicketCoupe then InsertChampTEX('TEX_TOTAL') ;
if InFacturette then InsertChampTEX('TEX_FINSLIP') ;
if InValidation then InsertChampTEX('TEX_FINVAL') ;
if InCheque     then InsertChampTEX('TEX_FINCHQ') ;
if Result then
   Begin
   if (InPreview) then Result:=(FPreView.ShowModal=mrOk)
      else Result:=FPreView.Imprime(TRUE,FALSE,Err) ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImprime.InitLP (TypeEtat,NatEtat,CodeEtat,Where,Imprimante,Port,Params : String ; Initialise,Preview,SQLBased : Boolean ; LaTOB : TOB ; Var Err : TMC_Err ; AIsImpControl : Boolean = FALSE ) : Boolean ;
var OkModele,ChangementVersion : Boolean ;
Begin
FProgress.SetValue(0) ;
FProgress.Libelle:=LPParam.Libelle ;
FProgress.Show ;
Result:=FALSE ; Fillchar(Err,sizeof(TMC_Err),#0) ;
InPreView:=PreView ;
IsImpControl:=AIsImpControl ;
Fport:=Port ;
FParams:=Params ;
Fimpr:=Imprimante ;
FInitImp:=Initialise ;
LPParam.TypeEtat:=TypeEtat ;
LPParam.NatEtat:=NatEtat ;
LPParam.CodeEtat:=CodeEtat ;
LPParam.Langue:=V_PGI.LanguePerso ;
OkModele:=FALSE ;
if not assigned(SauvgardeLP) then SauvgardeLP:=TLPSauvgarde.create(nil)
   else OkModele:=LpParam.isSameModele(SauvgardeLP.ParamLP) ;
if OkModele then
   Begin
   LPBase.recopiecontenu:=TRUE ;
   LPBase.assign(SauvgardeLP.BaseLP) ;
   LPBase.recopiecontenu:=FALSE ;
   LpParam.assign(SauvgardeLP.ParamLP) ;
   End else
   Begin
   if not ChargeModeleLP(LPParam,FALSE,LpBase,nil,nil,Err,ChangementVersion) then Exit ;
   LPParam.SQL:=VerifieChamps(LPParam.SQL) ;
   SauvgardeLP.BaseLP.recopiecontenu:=TRUE ;
   SauvgardeLP.BaseLP.assign(LPBase) ;
   SauvgardeLP.BaseLP.recopiecontenu:=FALSE ;
   SauvgardeLP.ParamLP.assign(LPParam) ;
   End ;
if not VerifParam(Err) then exit ;
FSQLBased:=(FSQLBased) and (SQLBased) ;
UneTOB:=LaTOB ;
if trim(Where)<>'' then RajouteWhere(Where)  ;
if FSQLBased then
   Begin
   if trim(LPParam.Sql)='' then
      Begin
      Err.Code:=-1 ;
      Err.Libelle:=MC_MsgErrDefaut(1029) ;
      End else
      Begin
      QPrin:=OpenSQL(LPParam.SQL,True) ;
      if not QPrin.Eof Then FProgress.SetMax(RecordsCount(QPrin)) else
         Begin
         Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1030) ;
         End ;
      QPrin.Close ;
      End ;
   End else
   if Assigned(UneTOB) then
      Begin
      If UneTOB.Detail.Count>0 then FProgress.SetMax(UneTOB.Detail.Count) else FPRogress.SetMax(1) ;
      End else
      Begin
      Err.Code:=-1 ;
      Err.Libelle:=MC_MsgErrDefaut(1030) ;
      End ;
Result:=(Err.Code=0);
if not Result then FProgress.Hide ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function ImprimeLP (TypeEtat,NatEtat,CodeEtat,Where,Imprimante,Port,Params : String ; Initialise,PreView,SQLBased : Boolean ; UneTOB : TOB ; Var Err : TMC_Err ; IsImpControl : Boolean = FALSE ; IsTestMateriel : Boolean = FALSE  ) : Boolean ;
Var X : TLPImprime ;
Begin
X:=TLPImprime.Create(application) ;
try
  X.InTestMateriel:=IsTestMateriel ;
  result:=(X.InitLP(TypeEtat,NatEtat,CodeEtat,Where,Imprimante,Port,Params,Initialise,Preview,SQLBased,UneTOB,Err,IsImpControl)) and
          (X.Imprime(Err)) ;
 finally
  X.Free ;
 end ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//XMG 30/03/04 début
Function ExporteLP(TypeEtat,NatEtat,CodeEtat,Where,NomFichier : String ; PreView,SQLBased : Boolean ; UneTOB : TOB ; Var Err : TMC_Err ) : Boolean ;
Var X : TLPImprime ;
Begin
X:=TLPImprime.Create(application) ;
try
  X.InTestMateriel:=FALSE ;
  result:=(X.InitLP(TypeEtat,NatEtat,CodeEtat,Where,'EXP','FTX',NomFichier,FALSE,Preview,SQLBased,UneTOB,Err,FALSE)) and
          (X.Imprime(Err)) ;
 finally
  X.Free ;
 end ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//XMG 30/03/04 fin
Function Init_ImprLP (Imprimante,Port,params : String ; parle : boolean ) : Boolean ;
Var X   : TLPPrinter;
    Err : TMC_Err ;
Begin
err.code:=0 ; err.Libelle:='' ;
X:=TLPPrinter.Create(Application,Imprimante) ;
try
   result:=(X.ChargePortetParams(Port,params,err)) and (X.InitImprimante(Err)) ;
  finally
   X.Free ;
  end ;
if (Parle) and (Not Result) and (Err.Code<>0) then PGIBox(Err.Libelle,LPMessageCap) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

initialization

Somme:=TraduireMemoire(Somme) ;
Absol:=TraduireMemoire(Absol) ;


end.

