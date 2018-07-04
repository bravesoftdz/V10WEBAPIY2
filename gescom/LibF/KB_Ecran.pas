unit kb_Ecran;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, (*Dialogs, *)HPanel,HTb97, Hent1, math,
     HCtrls, (*Lookup, (*StdCtrls, *)UTOB, ExtCtrls, buttons, HSysmenu, ImgList, HMsgBox, HStatus, ed_tools, Hgauge,
     JPeg,
{$IFDEF EAGLCLIENT}   //MR debut 06/11/02
{$ELSE}
     Db, DbTables,
{$ENDIF}              //MR fin 06/11/02

{$IFDEF PGIS5}
     MC_Lib, EntGC, Ent1, FOUtil, FODefi ;    //NA 17/11/03
{$ELSE}
     Ut, uterreur2(*, utpv *); //XMG 07/01/04
{$ENDIF}
Const CENbrBtnWidthDef    =  12 ;
      CENbrBtnHeightDef   =   5 ;

type
  TKbEcranSupport = class(TDataModule)
    ImagesBtn: TImageList;
  end;

  TOnBtnEcran  = Procedure ( Concept,Code,Extra : String ; Qte,Prix : Double ) of object ;
  TOnSaisieClc = Procedure ( Val : Double ) of object ;
  TOnActionClc = Procedure ( Val : String ) of object ;
  TParamBouton = Procedure ( UnBouton : TOB ) of object ;
  TOnChangeClc = Procedure ( AValue : Variant ) of object ;  //NA 17/11/03

  TParamsvisuels = class (TObject)
    Nom    : String ;
    r      : TRect ;
    FncRsz : integer ; //Fonction special lors le resize
    End ;

const KBt1 = VK_ESCAPE ;
      KBt2 = VK_ESCAPE ;
      KBTx = VK_ESCAPE ;
type
  TKB_Legend = class (TComponent)
   private
    FShift  : TShiftState ;
    Fkey    : string ;
    FActive : integer ;
    time    : TTimer ;
    procedure VideTampon ;
    procedure istime ( sender : TObject) ;
   public
    constructor create(AOwner : TComponent) ; override ;
    destructor destroy ; override ;

    Function gerekey ( var akey : word ) : string ;
   end ;


  TLCDVal = Class
    Private
     FVal : Variant ;
     FAff : THPanel ;
     FOnSetVal : TOnChangeClc ;  //NA 17/11/03
    Protected
     Procedure SetVal ( AValue : Variant ) ;
     Function  GetString  : String ;
     Function  GetFloat   : Double ;
     Function  GetInteger : Integer ;
    Public
     Constructor Create ; virtual ;

     Procedure Clear ;
     Procedure Ajoute ( AValue : Variant ) ;
     procedure Back ;
     Procedure Signe ;

     Property Afficheur : THPanel read FAff      write FAff ;
     property Valeur    : Variant read FVal      write SetVal ;
     Property AsString  : String  read GetString ;
     Property AsFloat   : Double  read GetFloat  ;
     property AsInteger : Integer read GetInteger ;

     property OnSetVal : TOnChangeClc read FOnSetVal write FOnSetVal ;  //NA 17/11/03
    End ;

type
  TClavierEcran = Class ;  //Forward

  TPosClc = (pcRight,pcCenter,pcLeft) ;

  TCalculatriceEcran = Class(THPanel)
    Private
     FWOrigin     : Integer ;
     FHOrigin     : Integer ;
     FValLCD      : TLCDVal ;
     FPnlAff      : THPanel ;
     FAff         : THpanel ;
     FLCDVisible  : Boolean ;
     FVisuels     : TStringlist ;
     FPosClc      : TPosClc ;

     FOnActionClc : TOnActionClc ;
     FOnSaisieClc : TOnSaisieClc  ;
     Procedure CreateBoutons ;
     Procedure BtnsClc( Sender : TObject ) ;
     Procedure SwitchBtnsCalculette ;
     Procedure GereResize (Visu : TParamsVisuels ; Ctrl : TControl ) ;
     Procedure ResizeCalculatrice ( Sender : TObject ) ;
     procedure CMTextChanged(var Message: TMessage); message CM_TEXTCHANGED;
    Protected
     Procedure SetLCDVisible ( AValue : Boolean ) ;
     Procedure SetPosCLc ( AValue : TPosClc ) ;
     procedure SetEnabled (Value : Boolean ) ; Override ; //XMG 03/10/01
    Public
     Constructor Create (AOwner : TComponent ) ; override ;
     Destructor  Destroy ; Override ;

     Procedure SetBounds (ALeft, ATop, AWidth, AHeight: Integer); override ;
     Procedure BtnsCalculetteUP ;

     Property LCD        : TLCDVal read FValLCD ;
    Published
     Property LCDVisible  : Boolean      read FLCDVisible  write SetLCDVisible default TRUE ;
     Property Position    : TPosClc      read FPosClc      Write SetPosClc     default pcLeft ;
     property OnSaisieClc : TOnSaisieClc read FOnSaisieClc write FOnSaisieClc ;
     property OnActionClc : TOnActionClc read FOnActionClc write FOnActionClc ;
    End ;


  TClavierEcran = class (THPanel)
    Private
     FGAuge            : TEnhancedGauge ;
     FImages           : TList ;
     KLng              : TKB_Legend ;
     FBtnsStc          : TStringList ;
     FPageCourante     : Integer ;
     FMaxPages         : Integer ;
     FSupport          : TKBEcranSupport ;
     FVisuels          : TStringlist ;
     FCalculette       : TCalculatriceEcran ;
     FCaisse           : String ;
     FParametrage      : Boolean ;
     FPageMenu         : Boolean ;
     FBoutons          : TOB ;
     FParamBouton      : TParamBouton ;
     FLanceBouton      : TOnBtnEcran ;
     FNbrBtnWidth      : integer ;
     FNbrBtnHeight     : Integer ;

     FMargeTopBtn      : Integer ;
     FMargeBottomBtn   : Integer ;
     FMargeLeftBtn     : Integer ;
     FMargeRightBtn    : Integer ;

     FSepHorzBtn       : Integer ;
     FSepVertBtn       : Integer ;

     FHeightClavier    : Integer ;
     FWidthClavier     : Integer ;

     FHeightBtn        : Integer ;
     FWidthBtn         : Integer ;
     FNbrMaxBoutons    : Integer ;

     FInuse            : Boolean ;
     FChargeAut        : Boolean ;  //XMG 02/10/01
     ComboPage         : THValComboBox ; //XMG 05/10/01
     ConceptMenu       : String ;  //XMG 05/10/01
     BtnMenu           : TOB ;  //XMG 05/10/01
     PagePrinc         : Boolean ; //NA 06/11/02
     Procedure CreateUnButton( i : Integer ) ;
     Procedure MetForceBouton( Btn : TToolBarButton97 ; UnBtn : TOB )  ;
     function  CalculeCleBouton ( UnBtn : TOB ) : integer ;  //XMG 05/10/01
     Function  ChercheBtnparcle(Cle : Integer ) : Tob ; //XMG 05/10/01
     Procedure LieeBoutons( Btn : TToolBarButton97 ; UnBtn : TOB )  ;
     procedure LanceRequeteFSF(Code,Extra : String ; BtnAppel : TOB) ;  //XMG 20/07/01
     Procedure ClickBouton ( Sender : TObject ) ;
     Function  ChercheQte ( Qte : Double ) : Double ;
     Procedure UnBoutonArticle ( Prop : TOB ; ChargeCaption : Boolean ) ;  //XMG 08/10/01
     Procedure ChargeImage( Prop : TOB ) ;
     Procedure DessineImage( Btn : TToolBarButton97 ; Prop : TOB ) ;
     Procedure ChargePageCourante ;
     procedure ChargeMaxPages ;
     Procedure CreeComboChargePage ;
     Procedure ChargePageMenu( Concept : String ; BtnAppel : TOB ) ;
     Procedure ChargePageMenuInterne( Concept : String ; BtnAppel : TOB ) ; //NA 06/11/02
     Procedure ChargePageMenuArt( Select : String ; BtnAppel : TOB ) ;
{$IFDEF CHR}
     Procedure ChargePageMenuRes( Select : String ; BtnAppel : TOB ) ;  //NA 17/11/03
{$ENDIF}
     Procedure CalculeNbrBoutons ;
     Procedure CalculetailleClavier ;
     Function  CalculeLeftBouton(Col : integer) : integer ;
     Function  CalculeTopBouton(Lgn : integer) : integer ;
     Procedure VideBtnsStc ;
     Function  DeplaceBoutons( OPage,NPage,ONum,NNum : Integer ; Var Err : TPGIerrS1) : Integer ; //XMG 07/01/04
     Function  BoutonSurAutresPagesInterne ( Page,Num : Integer ) : Boolean ;  //XMG 03/10/01

     Property HeightBtn     : Integer  read FHeightBtn ;
     Property WidthBtn      : Integer  read FWidthBtn ;
     Property NbrMaxBoutons : integer  read FNbrMaxBoutons ;
     Property HeightClavier : Integer  read FHeightClavier ;
     Property WidthClavier  : Integer  read FWidthClavier ;

    Protected
     Function  GetImages : TimageList ;
     Procedure SetCaisse       (AValue : String ) ;
     Procedure SetParametrage  (AValue : Boolean ) ;
     Procedure SetPageCourante (AValue : Integer ) ;
     Procedure CreeLesboutonsPave ;
     Procedure BloqueResize ( Nom : String) ;
     Procedure GereResize (Visu : TParamsVisuels ; Ctrl : TControl ) ;
     Function  VerifiesupBtns(Newval : integer ; isWidth : Boolean ) : Boolean ;
     Procedure SetNbrBtnWidth (AValue : Integer ) ;
     Procedure SetNbrBtnHeight (AValue : Integer ) ;

     Procedure SetMargeTopBtn (AValue : Integer ) ;
     Procedure SetMargeBottomBtn (AValue : Integer ) ;
     Procedure SetMargeLeftBtn (AValue : Integer ) ;
     Procedure SetMargeRightBtn (AValue : Integer ) ;

     Procedure SetSepHorzBtn (AValue : Integer ) ;
     Procedure SetSepVertBtn (AValue : Integer ) ;

     Procedure CalculeWidthBouton ;
     Procedure CalculeHeightBouton ;
     Procedure SetOnSaisieClc ( AValue : TOnSaisieClc) ;
     Procedure SetOnActionClc ( AValue : TOnActionClc) ;
     Function  GetOnSaisieClc : TOnSaisieClc ;
     Function  GetOnActionClc : TOnActionClc ;
     Function  PlaceWidthClavier (SiClc : Boolean = TRUE ) : integer ;
     Function  PlaceHeightClavier : integer ;
     Procedure SetClcVisible ( AValue : Boolean ) ;
     Function  GetClcVisible : Boolean ;
     Procedure SetClcPos ( AValue : TPosClc) ;
     Function  GetClcPos : TPosClc ;
     Procedure SetOnChangeClc ( AValue : TOnChangeClc) ;  //NA 17/11/03
     Function  GetOnChangeClc : TOnChangeClc ;  //NA 17/11/03
     Function  ComposeNumeroBtn ( Lin,Col : Integer ) : Integer ;  //XMG 03/10/01
     Function  DecomposenumeroBtn( NumBtn : Integer ; IsColonne : Boolean ) : Integer ;  //XMG 03/10/01
     Procedure CreerChampsSupp ( UnBtn : TOB ) ;  //XMG 03/10/01
     Function  AjouteBtnTOB ( UnBtn : TOB ; replace,ChargeCaption : boolean ) : Boolean ;  //XMG 05/10/01
    Public
     Constructor Create ( AOwner : TComponent ; Indice : integer = 0) ; reintroduce; overload; virtual; //NA 17/11/03
     Destructor  Destroy ; Override ;

     procedure GereKeyPress( var key : char ) ;
     Function ChercheBoutonKey ( Touche : String ) : TOB ;  //XMG 29/11/01
     Procedure GereKey(var Key: Word; Shift: TShiftState);
     Procedure ChargeClavier (Lacaisse : String ) ;
     Function  ChargeFromQ ( Q : TQuery ) : Boolean ;
     Procedure VidePave ;  //XMG 02/10/01
     Procedure EnregClavier ;
     Function  isModified : Boolean ;
     Procedure UpdateBouton ( UnBouton : TOB ) ;
     Procedure ResizeClavier ( Sender : TObject ) ;
     Function  BoutonSurAutresPages ( Btn : TOB) : Boolean ;
     Procedure AlleraPage ( BtnName : string ; Page : integer ) ;  //XMG 15/04/02
     Procedure RecopieClavier( Lacaisse : String ) ;
     Function  ExisteBouton ( Page,Ligne,Colonne : integer ) : Boolean ;
     Function  MoveBouton (OldPage,OldNumero,NewPage,NewLigne,NewCol : integer ; var Err : TPGIErrS1 ) : Boolean ; //XMG 07/01/04
     Function  BtnLibreEnPage ( NbrPage : Integer ) : Integer ;  //XMG 03/10/01
     Function  Btnsenpage ( NbrPage : Integer) : Integer ;  //XMG 03/10/01
     Function CreateBtnStream ( Parms : String ; replace,ChargeCaption : Boolean ) : TOB ;
     Procedure LanceChargePageMenu( Concept, TblTT, Plus : String  ; Couleur : TColor ; Image : Integer ) ;       //NA 13/03/02
     Function  ChercheValeurCalculette ( Qte : Double ; Raz : boolean ) : Double ;   //NA 13/01/03

     Property Caisse            : String      read FCaisse         write SetCaisse ;
     Property Parametrage       : Boolean     read FParametrage    write SetParametrage ;
     Property PageCourante      : Integer     read FPageCourante   write SetPageCourante ;
     Property MaxPages          : Integer     read FMaxPages     ;
     Property Images            : TImageList  read getImages     ;

     Property NbrBtnWidth       : integer     read FNbrBtnWidth    write SetNbrBtnWidth ;
     Property NbrBtnHeight      : Integer     read FNbrBtnHeight   write SetNbrBtnHeight ;

     Property MargeTopBtn       : Integer     read FMargeTopBtn    write SetMargeTopBtn ;
     Property MargeBottomBtn    : Integer     read FMargeBottomBtn write SetMargeBottomBtn ;
     Property MargeLeftBtn      : Integer     read FMargeLeftBtn   write SetMargeLeftBtn ;
     Property MargeRightBtn     : Integer     read FMargeRightBtn  write SetMargeRightBtn ;

     Property SepHorzBtn        : integer     read FSepHorzBtn     write SetSepHorzBtn ;
     Property SepVertBtn        : integer     read FSepVertBtn     write SetSepVertBtn ;

     Property ClcVisible        : Boolean     read GetClcVisible   Write SetClcVisible ;
     Property ClcPosition       : TPosClc     Read GetClcPos       Write SetClcPos ;

     Property LanceBouton      : TOnBtnEcran  read FLanceBouton    write FLanceBouton ;
     Property ParamBouton      : TParamBouton read FParamBouton    write FParamBouton ;
     Property LanceCalculette  : TOnSaisieClc read GetOnSaisieClc  write SetOnSaisieClc ;
     Property BoutonCalculette : TOnActionClc read GetOnActionClc  write SetOnActionClc ;
     Property OnChangeCalculette : TOnChangeClc read GetOnChangeClc  write SetOnChangeClc ; //NA 17/11/03
     Property Calculette      : TCalculatriceEcran     read FCalculette   write FCalculette ; //NA 17/11/03

     Property Gauge            : TEnhancedGauge read FGauge        write FGauge ;
     Property ChargeAut        : Boolean        read FChargeAut    Write FChargeAut ;  //XMG 02/10/01
    End ;

type TTypePave = (tpEncaissement, tpFamilles) ;
     TSetTypepave = Set of TTypePave ;
Function CreepaveDynamique ( pCE : TClavierEcran ; Caisse,Params : String ; TypePave : TSetTypePave ; OkGauge : Boolean ; Var Err :TPGIErrS1  ) : Boolean ; //XMG 07/01/04 02/10/01
Procedure LanceCreationAutomatiquePaves (Msg : String ) ;
function ConfirmeCreationPave ( Msg : String ; ForceUneCaisse : Boolean ) : Boolean ;
Function CE_Charge ( Caisse : String ; pCE_S : TOB ) : Boolean ;
Function CE_Existe (CE_CAISSE : String ) : Boolean ;
Function CE_Controle ( pCE_S : TOB ; Var Err : TPGIErrS1 ) : Boolean ; //XMG 07/01/04
Function CE_Efface (CE_CAISSE : String ) : Boolean ;
function CE_Sauve  (pCE_S : Tob ; Var Err : TPGIErrS1 ) : boolean ; //XMG 07/01/04

Procedure AppliquePolice( TextePolice : String ; Font : TFont ; Var Alignment : TAlignment )  ;
Function KBMakeSQL ( Concept, Code : String ; var CLe1, Libelle : String ) : String ;
Procedure DonneTablette ( Concept, Caisse : String ; Var tablette,plus : String ) ;  //NA 13/03/02
Function IsKBIDCaptionStd(var Caption: string; delKbId: boolean = True): boolean; //NA 06/01/04
Function GetKBIDCaptionStd : string ; //NA 06/01/04

Const ConstVide       = 'VIDE;' ;  //XMG 09/10/01
      ConstToucheVide = '-;-;-;;'  ; //XMG 29/11/01

implementation

{$R *.DFM}

uses ParamSoc      //XMG 01/04/03
{$IFDEF PGIS5}
{$ELSE}
    , S1Util
{$ENDIF}
    ;             //XMG 01/04/03
Const MargeTopBtnDef    =   6 ;
      MargeLeftBtnDef   = MargeTopBtnDef ;
      MargeBottomBtnDef = MargeTopBtnDef ;
      MargeRightBtnDef  = MargeTopBtnDef ;
      HeightBtnDef      =  65 ;
      WidthBtnDef       =  65 ;
      SepVertBtnDef     =   0 ;
      SepHorzBtnDef     =   0 ;

      TailleWStatics    =   3 ;  //En boutons en boutons pavé
      TailleHStatics    =   2 ;  //En boutons en boutons pavé
      HeightBtnStc      =  65 ;
      WidthBtnStc       =  65 ;
      SepVertBtnStc     =   1 ;
      SepHorzBtnStc     =   1 ;
      MargetopBtnStc    =   1 ;
      MargeLeftBtnStc   =   1 ;
      MargeBottomBtnStc = MargeTopBtnStc ;
      MargeRightBtnStc  = MargeLeftBtnStc ;
      WidthStatics      = MargeleftBtnStc+(WidthBtnStc+SepHorzBtnStc)*TailleWStatics-SepHorzBtnStc+MargeRightBtnStc ;
      HeightStatics     = MargeTopBtnStc+(HeightBtnStc+SepVertBtnStc)*TailleHStatics-SepVertBtnStc+MargeBottomBtnStc ;

      MargeTopClc       = MargeTopBtnDef ;
      MargeLeftClc      = MargeLeftBtnDef ;
      MargeBottomClc    = MargeBottomBtnDef ;
      MargeRightClc     = MargeRightBtnDef ;

      HeightBtnClc      =  60 ;
      WidthBtnClc       =  60 ;
      SepVertBtnClc     =   0 ;
      SepHorzBtnClc     =   0 ;

      MargetopBtnClc    =   9 ;
      MargeLeftBtnClc   = MargeTopBtnClc ;
      MargeBottomBtnClc = MargeTopBtnClc ;
      MargeRightBtnClc  = MargeTopBtnClc ;
      Pos1erBtnClc      = HeightBtnClc+MargeTopBtnClc ;

      NbrBtnClcHeight   =   5 ;
      NbrBtnClcWidth    =   4 ;


      WidthPnlLCD       = (WidthBtnClc+SepHorzBtnClc)*NbrBtnClcWidth-SepHorzBtnClc ;
      HeightPnlLCD      =  HeightBtnClc+SepVertBtnClc ;

      WidthLCD          = WidthPnlLCD-MargeLeftBtnClc-MargeRightBtnClc ;
      HeightLCD         = HeightPnlLCD-MargeTopBtnClc-MargeBottomBtnClc ;
      NbrCharsLCD       =  10 ;

      WidthCalculetteDef   = MargeleftBtnClc+(WidthBtnClc+SepHorzBtnClc)*NbrBtnClcWidth-SepHorzBtnClc+MargeRightBtnClc ;
      HeightCalculetteDef  = MargeTopBtnClc+(HeightBtnClc+SepVertBtnClc)*NbrBtnClcHeight-SepVertBtnClc+MargeBottomBtnClc ;

      ConstKbIdCaptionStd  = '&&&'; //NA 06/01/04
      ConstKbIdCaptionStdOLD = #255+#255 ; //NA 07/01/04

Type TProcResize = Procedure (Visu : TParamsVisuels ; Ctrl : TControl ) of Object ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function IsKBIDCaptionStd ( var Caption : string ; delKbId : boolean = True) : boolean ; //NA debut 06/01/04
Begin
Result := False;
if Copy(Caption,1,2)=ConstKbIdCaptionStdOld then //XMG 07/01/04
   Begin
   Result := True;
   if delKbId then Delete(Caption,1,2) ;
   End else
if Copy(Caption,1,Length(ConstKbIdCaptionStd))=ConstKbIdCaptionStd then
   Begin
   Result := True;
   if delKbId then Delete(Caption,1,Length(ConstKbIdCaptionStd)) ;
   End;
End;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function GetKBIDCaptionStd : string ;
Var OldStyle : boolean;
Begin
OldStyle:=GetSynRegKey('OldKBIDCaptionStd',False,True);
if OldStyle then Result:=ConstKbIdCaptionStdOld else Result:=ConstKbIdCaptionStd ; //XMG 07/01/04
End; //NA fin 06/01/04
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure GardeParamsVisuels ( Visuels : TStringlist ; Ctrl : TControl ; AFncRsz : integer = -1 ) ;
Var Visu : TParamsVisuels ;
    i    : integer ;
Begin
if (not assigned(Visuels)) or (Not Assigned(Ctrl)) or (Ctrl.Name='') then exit ;
i:=Visuels.IndexOf(Ctrl.Name) ;
if i>-1 then Visu:=TParamsVisuels(Visuels.objects[i])
   else Visu:=TParamsVisuels.Create ;
With visu do
  Begin
  Nom:=Ctrl.Name ;
  with r do
    Begin
    Top:=Ctrl.Top ;
    Left:=Ctrl.Left ;
    Bottom:=Ctrl.Height ;
    right:=Ctrl.Width ;
    End ;
  FncRsz:=AFncRsz ;
  End ;
if i<0 then
   Begin
   Visuels.BeginUpdate ;
   Visuels.Addobject(Visu.Nom,Visu) ;
   Visuels.EndUpdate ;
   End ; 
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure ResizelesClaviers ( Clavier : TWinControl ; Visuels : TStringlist ; W,H : integer ; ProcResize : TProcResize = nil  ) ;
           //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
           Procedure CalculCoefs ( Ctrl : TControl ; var Cx,Cy : Double  ; W,H : integer ) ;
           Begin
           Cx:=1 ; Cy:=1 ;
           W:=maxintvalue([1,W]) ; H:=MaxIntValue([1,H]) ; 
           if not assigned(Ctrl.parent) then exit ;
           Cx:=Ctrl.ClientWidth/W ;
           Cy:=Ctrl.ClientHeight/H ;
           End ;
           //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
           Procedure ResizelesControls(WCtrl : TWinControl ; Cx,Cy : Double ; Visuels : TStringlist ; ProcResize : TProcResize ) ;
           var i,j     : integer ;
               Visu    : TParamsVisuels ;
               Ctrl    : TControl ;
           Begin
           For i:=0 to WCtrl.Controlcount-1 do
             Begin
             Ctrl:=WCtrl.Controls[i] ;
             j:=Visuels.indexof(Ctrl.Name) ;
             if j>-1 then
                Begin
                Visu:=TParamsVisuels(Visuels.objects[j]) ;
                Ctrl.Setbounds(round(Visu.r.Left*Cx),round(Visu.r.Top*Cy),round(Visu.r.right*Cx),round(Visu.r.Bottom*Cy)) ;
                if (Assigned(ProcResize)) and (Visu.FncRsz>-1) then ProcResize(Visu,Ctrl) ;
                if (Ctrl is TWinControl) and (TWinControl(Ctrl).controlcount>0) then Resizelescontrols(TWinCOntrol(Ctrl),Cx,Cy,Visuels,ProcResize) ;
                End ;
             End ;
           End ;
           //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
var Cx,Cy : Double ;
Begin
if (not assigned(Clavier)) or (not assigned(Visuels)) then exit ;
CalculCoefs(Clavier,Cx,Cy,W,H) ;
ResizelesControls(Clavier,Cx,Cy,Visuels,ProcResize) ;
Clavier.invalidate ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//                           TKB_LEGEND
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
constructor TKB_Legend.create(AOwner : TComponent) ;
var i : Dword ;
begin
inherited create(AOwner) ;
videTampon ;
time:=TTimer.Create(Self) ;
time.enabled:=false ;
time.OnTimer:=Istime ;
systemparametersinfo(SPI_GETKEYBOARDDELAY,0,@i,0) ;
time.Interval:=maxintvalue([i,10]) ;
end ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Destructor TKB_Legend.Destroy ;
begin
time.Free ;
inherited destroy ;
end ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TKB_Legend.VideTampon ;
begin
fshift:=[] ;
fkey:='' ;
Factive:=0 ;
end ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TKB_Legend.istime ( sender : TObject) ;
begin
videtampon ;
Time.Enabled:=FALSE ;
end ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TKB_Legend.gerekey ( var akey : word ) : string ;
var ch : Char ;
begin
result:=#0 ;
if akey=0 then exit ;
{if (akey in [VK_SHIFT, VK_CONTROL, VK_MENU]) and (FActive=0) then
   begin
   case akey of
     VK_SHIFT   : include(fShift,ssShift) ;
     VK_CONTROL : include(FShift,ssCtrl) ;
     VK_MENU    : include(FShift,ssAlt) ;
     end ;
   //akey:=0 ;
   end else}
if ((Akey=KBt1) and (FActive=0)) or ((Akey=KBt2) and (FActive=1)) then
   begin
   //Time.Enabled:=TRUE ;
   inc(FActive) ;
   akey:=0 ;
   end else
if (akey=KBtx) and (FActive=2) and (trim(fkey)<>'') then
   begin
   FKey:=Uppercase(FKey) ;  Ch:=FKey[1] ; delete(FKey,1,1) ;
   result:=FKey ; fshift:=[] ;
   (*
     | Shift | Control | Alt |  Value |
     |       |         |     |    N   |
     |   O   |         |     |    S   |
     |   O   |    O    |     |    W   |
     |   O   |         |  O  |    X   |
     |   O   |    O    |  O  |    Y   |
     |       |    O    |     |    C   |
     |       |    O    |  O  |    Z   |
     |       |         |  O  |    A   |
     *)
   if ch in ['S','W','X','Y'] then include(fShift,ssShift) ;
   if ch in ['C','W','Y','Z'] then include(FShift,ssCtrl) ;
   if ch in ['A','X','Y','Z'] then include(FShift,ssAlt) ;
   REsult:=TrueFalsest(ssCtrl in FShift)+';'+TrueFalsest(ssShift in FShift)+';'+TrueFalsest(ssAlt in FShift)+';'+Result+';' ;
   videtampon ;
   akey:=0 ;
   end else
if (FActive=2) then
   begin
   ch:=keytochar(akey,[ssshift]) ;
   if trim(ch)<>'' then
      begin
      FKey:=FKey+ch ;
      akey:=0 ;
      end ;
   end else videtampon ;
end ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//                           TLCDVal
//
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLCDVal.Create ;
Begin
Inherited Create ;
Clear ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLCDVal.SetVal(Avalue : Variant ) ;
Begin
FVal:=AValue ;
if Assigned(Afficheur) then Afficheur.Caption:=AsString ;
if Assigned(FOnSetVal) then FOnSetVal(AValue) ; //NA 17/11/03
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLCDVal.GetString  : String ;
Begin
Result:=VString(Valeur) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLCDVal.GetFloat   : Double ;
Begin
result:=0 ;
if isnumeric(Valeur) then Result:=VDouble(Valeur) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLCDVal.GetInteger : Integer ;
Begin
result:=0 ;
if isnumeric(valeur) then Result:=VInteger(Valeur) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLCDVal.Clear ;
Begin
Valeur:=#0 ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLCDVal.Ajoute ( AValue : Variant ) ;
Var St : String ;
Begin
St:=Asstring ;
if ((Avalue=V_PGI.SepDecimal) and (pos(AValue,St)>0)) or (length(St)>=NbrCharsLCD) then exit ;
St:=St+AValue ;
if St=V_PGI.Sepdecimal then St:='0'+St ;
while(length(St)>1) and (Copy(St,1,1)='0') and (Copy(st,1,2)<>'0'+V_PGI.SepDecimal) do delete(St,1,1) ;
Valeur:=St ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLCDVal.Back ;
var St : String ;
Begin
St:=AsString ;
Valeur:=copy(St,1,length(St)-1) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLCDVal.Signe ;
var St : String ;
Begin
st:=AsString ;
if (length(St)>0) and (St[1]='-') then Valeur:=Copy(St,2,length(St))
   else if Length(St)<NbrCharsLCD then Valeur:='-'+St ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//                           TCalculatriceEcran
//
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TCalculatriceEcran.Create (AOwner : TComponent ) ;
Begin
Inherited Create(AOwner) ;
OnResize:=ResizeCalculatrice ;

FVisuels:=TStringList.Create ;
FPosClc:=pcLeft ; //pcNO ;

Height:=HeightCalculetteDef ;
Width:=WidthCalculetteDef ;
FWOrigin:=Width ;
FHOrigin:=Height ;
BackGroundEffect:=bdHorzIn ;
ColorStart:=clSilver ;
FOnSaisieClc:=nil ;
FOnActionClc:=nil ;


//Panel Afficheur
FPnlAff:=THPanel.Create(Self) ;
with FPnlAff do
  Begin
  Parent:=Self ;
  Name:='PnlLCD' ;
  Caption:='' ;
  BackGroundEffect:=bdHorzIn ;
  ColorStart:=clGray ;
  Height:=HeightPnlLCD ;
  Width:=WidthPnlLCD ;
  Left:=MargeLeftBtnClc ;
  Top:=MargeTopBtnClc ;
  End ;
GardeParamsVisuels(FVisuels,FPnlAff) ;
//LCD label
FAff:=THpanel.Create(Self) ;
with FAff do
  Begin
  Parent:=FPnlAff ;
  Name:='LCDCALCULETTE' ;
  Caption:='' ;
  Alignment:=taRightJustify ;
  BevelInner:=bvRaised ;
  bevelOuter:=bvLowered ;
  Color:=clBlack ;
  Width:=WidthLCD ;
  Height:=HeightLCD ;
  Font.Color:=clLime ;
  Font.Name:='Arial' ;
  Font.height:=-FAff.Height ;
  Font.Style:=[fsBold] ;
  top:=(Parent.Height-Height) div 2 ;
  Left:=(Parent.Width-Width) div 2 ;
  End ;
GardeParamsVisuels(FVisuels,FAff,1) ;

FValLCD:=TLCDVal.Create ;
FValLCD.Afficheur:=FAff ;
CreateBoutons ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Destructor  TCalculatriceEcran.Destroy ;
Begin
FValLCD.Free ;
VideStringListe(FVisuels) ;
FVisuels.Free ;
FPnlAff.Free ;
Inherited Destroy ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TCalculatriceEcran.SetLCDVisible ( AValue : Boolean ) ;
Begin
FAff.Visible:=AValue ;
FPnlAff.Visible:=AValue ;
//Redesine calculatrice //à voir
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TCalculatriceEcran.SetPosCLc ( AValue : TPosClc ) ;
Begin
if Avalue=position  then exit ;
FPosClc:=AValue ;
Left:=-1 ; //Force le reaffichage de la calculatrice
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
//XMG debut 03/10/01
procedure TCalculatriceEcran.SetEnabled (Value : Boolean )  ;
Begin
inherited SetEnabled(Value) ;
EnableControls(Self,Value) ;
End ;
//XMG fin 03/10/01
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TCalculatriceEcran.CreateBoutons ;
var i,W,H,T,L  : Integer ;
    Val,St,Cap : String ;
    Btn        : TToolBarButton97 ;
Begin
For i:=1 to 15 do
  Begin
  W:=WidthBtnClc ;
  H:=HeightBtnClc ;
  T:=MargeTopBtnClc+Pos1erBtnClc+(HeightBtnClc+SepVertBtnClc)*((i-1) div 3) ;
  L:=MargeLeftBtnClc+(WidthBtnClc+SepHorzBtnClc)*((i-1) mod 3) ;
  if i in [1..10] then
     Begin
     Val:=inttostr(i mod 10) ; St:=Val ; Cap:=st ;
     if i=10 then L:=MargeLeftBtnClc+(WidthBtnClc+SepHorzBtnClc)*1 ; //W:=(WidthBtnClc+SepHorzBtnClc)*2-SepHorzBtnClc ;
     if i<>10 then T:=MargeTopBtnClc+Pos1erBtnclc+(HeightBtnClc+SepVertBtnClc)*((9-i) div 3) ;
     End else
     Case i of
       11 : Begin  // separateur decimal
            St:='DEC' ;
            Cap:=V_PGI.SepDecimal ;
            Val:=Cap ;
            L:=MargeLeftBtnClc+(WidthBtnClc+SepHorzBtnClc)*(i mod 3) ;
            End ;
       12 : Begin  //ENTER
            St:='ü' ; //ou Ã = case à cocher
            Cap:=St ;
            Val:='ENTER' ;
            ////T:=MargeTopBtnClc+HeightBtnClc+SepVertBtnClc ;
            T:=MargeTopBtnClc+Pos1erBtnclc+(HeightBtnClc+SepVertBtnClc)*2 ;
            L:=MargeLeftBtnClc+(WidthBtnClc+SepHorzBtnClc)*3 ;
            ////H:=(HeightBtnClc+SepVertBtnClc)*3-SepVertBtnClc ;
            H:=(HeightBtnClc+SepVertBtnClc)*2-SepVertBtnClc ;
            End ;
       13 : Begin  //CLEAR
            St:='û' ;
            Cap:=St ;
            Val:='CLEAR' ;
            ////T:=MargeTopBtnClc ;
            T:=MargeTopBtnClc+Pos1erBtnclc+HeightBtnClc+SepVertBtnClc ;
            L:=MargeLeftBtnClc+(WidthBtnClc+SepHorzBtnClc)*3 ;
            End ;
       14 : Begin  //SIGNE (Negatif/Positif)
            St:='SIGNE' ;
            Cap:='±' ;
            Val:='-' ;
            T:=MargeTopBtnClc+Pos1erBtnclc+(HeightBtnClc+SepVertBtnClc)*3 ; //((i-1) div 3) ;
            L:=MargeLeftBtnClc ; //+(WidthBtnClc+SepHorzBtnClc)*((i-1) mod 3) ;
            End ;
       15 : Begin  //BACK SPACE
            St:='ï' ;
            Cap:=St ;
            Val:='BACK' ;
            T:=MargeTopBtnClc+Pos1erBtnclc ;
            L:=MargeLeftBtnClc+(WidthBtnClc+SepHorzBtnClc)*3 ;
            End ;
       End ;
  Btn:=TToolBarButton97.Create(Self) ;
  with Btn do
    Begin
    //Btn.IntensiteLumiere:=0 ; //XMG 11/06/01
    Parent:=Self ;
    Name:='CLC'+IntToStr(i) ;
    Caption:=Cap ;
    Font.Size:=18 ;
(*{$IFDEF PGIS5}
{$ELSE}*)
    S1Enabled:=False ;
(*{$ENDIF}*)
    if i=12 then //Bouton ENTER
       begin
       Font.Name:='Wingdings' ;
       Font.Size:=48 ;
       Font.Color:=clGreen ;
       end ;
    if i=13 then //Bouton CLEAR
       begin
       Font.Name:='Wingdings' ;
       Font.Size:=36 ;
       Font.Color:=clRed ;
       end ;
    if i=15 then //Bouton BACK SPACE
       begin
       Font.Name:='Wingdings' ;
       Font.Color:=clBlack ;
       end ;
    Font.Style:=[fsbold] ;
    hint:=Val ;
    Flat:=True ;
    Opaque:=FALSE ;
    DisplayMode:=dmTextOnly ;
    parentShowHint:=FALSE ;
    ShowHint:=FALSE ;
    Top:=t ;
    Left:=l ;
    Width:=W ;
    Height:=h ;
    Alignment:=taCenter ;
    WordWrap:=TRUE ;
    onClick:=BtnsClc ;
    End ;
  GardeParamsVisuels(FVisuels,Btn) ;
  End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TCalculatriceEcran.SetBounds (ALeft, ATop, AWidth, AHeight: Integer);
var CE : TClavierEcran ;
    l  : Integer ;
Begin
if assigned(Parent) then
   Begin
   Atop:=MargetopClc ;
   case position of
     pcRight  : Aleft:=MargeLeftClc ;
     pcCenter : begin
                CE:=TClavierEcran(Parent) ;
                l:=(CE.PlaceWidthClavier-CE.SepHorzBtn*CE.NbrBtnWidth+CE.SepHorzBtn) div CE.NbrBtnWidth;
                Aleft:=CE.MargeLeftBtn+(l+CE.SepHorzBtn)*trunc(CE.NbrBtnWidth / 2)+MargeLeftClc ;
                End ;
     pcLeft   : ALeft:=maxintValue([MargeLeftClc,Parent.Width-AWidth-MargeRightClc]) ;
     End ;
   End ;
inherited SetBounds(Aleft,Atop,Awidth,Aheight) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TCalculatriceEcran.BtnsCalculetteUP ;
Begin
FValLCD.Clear ;
SwitchBtnsCalculette ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TCalculatriceEcran.SwitchBtnsCalculette ;
Var CC : TComponent ;
Begin
CC:=FindComponent('CLC12') ;   // Bouton ENTER
if (CC<>Nil) and (CC is TToolBarButton97) then with TToolBarButton97(CC) do
    Begin
    if trim(FValLCD.AsString)='' then
       Begin
       Caption:='ü' ;
       Font.Name:='Wingdings' ;
       Font.Size:=48 ;
       Font.Color:=clGreen ;
       End else
       Begin
       Caption:='=' ;
       if V_PGI.Tahoma then Font.Name:='Tahoma'
                       else Font.Name:='MS Sans Serif' ;
       Font.Size:=18 ;
       Font.Color:=clBlack ;
       End ;
   End ;
CC:=FindComponent('CLC13') ;   // Bouton CLEAR
if (CC<>Nil) and (CC is TToolBarButton97) then with TToolBarButton97(CC) do
    Begin
    if trim(FValLCD.AsString)='' then
       Begin
       Caption:='û' ;
       Font.Name:='Wingdings' ;
       Font.Size:=36 ;
       Font.Color:=clRed ;
       End else
       Begin
       Caption:='C' ;
       if V_PGI.Tahoma then Font.Name:='Tahoma'
                       else Font.Name:='MS Sans Serif' ;
       Font.Size:=18 ;
       Font.Color:=clBlack ;
       End ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TCalculatriceEcran.BtnsClc( Sender : TObject ) ;
Var val : String ;
Begin
if not (sender is TToolBarButton97) then exit ;
Val:=TToolBarButton97(Sender).Hint ;
if val='ENTER' then
   Begin
   If (trim(FValLCD.AsString)<>'') and (Assigned(OnSaisieClc)) then OnSaisieClc(Valeur(FValLCD.AsString)) ;
   If (trim(FValLCD.AsString)='') and (Assigned(OnActionClc)) then onActionClc(val) ;
   BtnsCalculetteUp ;
   End else
if val='CLEAR' then
   Begin
   If (trim(FValLCD.AsString)='') and (Assigned(OnActionClc)) then OnActionClc(val) ;
   BtnsCalculetteUp
   End else
if val='BACK' then FValLCD.Back else
if val='-' then  FValLCD.Signe //Bouton Signe
   else FValLCD.Ajoute(val) ;
SwitchBtnsCalculette ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TCalculatriceEcran.GereResize (Visu : TParamsVisuels ; Ctrl : TControl ) ;
Begin
if (not assigned(Visu)) or (Not Assigned(Ctrl)) then exit ;
if Visu.FncRsz=1 then // FAff, il faut matenir une taille de police qui soit affichable par le control
   Begin
   THPanel(Ctrl).Font.height:=-THPanel(Ctrl).Height ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TCalculatriceEcran.ResizeCalculatrice( Sender : TObject ) ;
Begin
ResizelesClaviers(Self,FVisuels,FWOrigin,FHOrigin,GereResize) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TCalculatriceEcran.CMTextChanged(var Message: TMessage);
begin
Caption:='' ;
//On Bloque l'affichage de la caption
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//                           TClavierEcran
//
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TClavierEcran.Create (AOwner : TComponent; Indice : integer = 0) ; //NA 17/11/03
Begin
Inherited Create(Aowner) ;
ChargeAut:=TRUE ;//XMG 02/10/01
if Indice > 0 then Name:='KB_ECRAN_'+inttostr(Indice) //NA 17/11/03
else Name:='KB_ECRAN' ; //NA 17/11/03
Caption:='' ;
Gauge:=nil ;
KLng:=nil ; //TKB_Legend.create(Self) ;  //XMG 29/11/00Conserver
FCaisse:='' ;
FImages:=TList.Create ;
FSupport:=TKBEcranSupport.Create(Self) ;
FVisuels:=TStringlist.Create ;
FBoutons:=TOB.Create('KBECRAN_S',nil,-1) ;

FCalculette:=TCalculatriceEcran.Create(Self) ;
FCalculette.Parent:=Self ;
if Indice > 0 then FCalculette.Name:='CALCULETTE_'+inttostr(Indice) //NA 17/11/03
else FCalculette.Name:='CALCULETTE' ; //NA 17/11/03

FBtnsStc:=TStringList.Create ;

FMargeTopBtn      :=MargeTopBtnDef ;
FMargeBottomBtn   :=MargeBottomBtnDef ;
FMargeLeftBtn     :=MargeLeftBtnDef ;
FMargeRightBtn    :=MargeRightBtnDef ;

FSepHorzBtn       :=SepHorzBtnDef ;
FSepVertBtn       :=SepVertBtnDef ;

FNbrBtnWidth:=0 ; //CENbrBtnWidthDef ;
FNbrBtnHeight:=0 ; //CENbrBtnHeightDef ;

FPageCourante:=-1 ;
FmaxPages:=0 ;
BackGroundEffect:=bdFond ;
//Visible:=FALSE ;
onResize:=ResizeClavier ;

Parametrage:=FALSE ;
FInUSe:=FALSE ;
FPageMenu:=FALSE ;  //XMG 05/10/01
ComboPage:=nil ; //XMG 05/10/01
ConceptMenu:='' ; //XMG 05/10/01
BtnMenu:=nil ; //XMG 05/10/01
PagePrinc:=False ; //NA 06/11/02
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.VideBtnsStc ;
var i : integer ;
Begin
For i:=0 to FBtnsStc.Count-1 do FBtnsStc.objects[i]:=nil ;
VideStringListe(FBtnsStc) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Destructor TClavierEcran.Destroy ;
Begin
if assigned(ComboPage) then ComboPage.Free ;  //XMG 05/10/01
VideBtnsStc ;
FBtnsStc.Free ;
VideListe(FImages) ;
FImages.Free ;
VideStringListe(FVisuels) ;
FVisuels.Free ;
FBoutons.Free ;
if (PagePrinc) and (assigned(BtnMenu)) then BtnMenu.Free ; //NA 06/11/02
if assigned(KLng) then KLng.Free ;
Inherited Destroy ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure DonneTablette ( Concept, Caisse : String ; Var tablette,plus : String ) ;      //NA 13/03/02
var Extra : String ; //XMG 20/07/01
{$IFDEF PGIS5}
    QQ : TQuery ;  //NA 13/03/02
{$ENDIF}
Begin
Tablette:='' ; Plus :='' ;
Extra:='' ;
Concept:=uppercase(Trim(Concept)) ; //XMG 05/10/01
if NbCarInString(Concept,';')>1 then
   Begin
   Extra:=Concept ;
   Concept:=ReadtokenSt(Extra) ;
   End ;
{$IFDEF PGIS5}
if Concept='ART' then Tablette:='GCARTICLE'                                                                                    else   //NA 06/11/02
if Concept='COM' then if GetParamSoc('SO_GCCOMMENTAIRE') then Tablette:='GCCOMMENTAIRE' else Tablette:='GCCOMMENTAIRELIGNE'    else   //NA 06/11/02
if Concept='VEN' then     //NA debut 13/03/02
   Begin
   Tablette:='GCCOMMERCIAL' ;
   Plus := 'GCL_TYPECOMMERCIAL="VEN" AND GCL_DATESUPP>"' + USDateTime(Date) + '"';  
   if GetInfoParPiece(FOGetNatureTicket(False, False), 'GPP_FILTRECOMM') = '001' then
      Begin
      QQ:=OpenSQL('Select GPK_ETABLISSEMENT from PARCAISSE where GPK_CAISSE="'+Caisse+'"', True) ;
      if not QQ.EOF then Plus:=Plus+' and GCL_ETABLISSEMENT="'+QQ.Findfield('GPK_ETABLISSEMENT').asString+'"' ;
      Ferme(QQ) ;
      End ;  //NA fin 02/12/03
   End else
if Concept='REM' then Begin Tablette:='GCTYPEREMISE' ; Plus:='GTR_FERME<>"X" AND GTR_DATEDEBUT<="' + USDateTime(Date) + '" AND GTR_DATEFIN>="' + USDateTime(Date) + '"'; End else  //NA 10/05/04
if Concept='REG' then Begin Tablette:='GCMODEPAIE'   ; Plus:='MP_UTILFO="X"' ;                                             End else
if Concept='FON' then Tablette:='GCFCTTICKETFO'                                                                                else
if Concept='FPV' then Begin Tablette:='GCFCTPAVEFO'  ; if readtokenst(Extra)='X' then plus:=' and CO_LIBRE<>"INVISIBLE"' ; end else
if Concept='OCT' then Begin Tablette:='GCOPECAISSEFO'; plus:='' ;                                                          end else
if Concept='MODMOD' then Tablette:='GCIMPMODELEFO'                                                                             else  //NA 06/11/02
if Concept='MODNAT' then Tablette:='GCFORMATTICKETFO'                                                                          else  //NA 06/11/02
    ;                     //NA fin 13/03/02
{$ELSE}
//XMG debut 05/10/01
if Concept='ART'    then Tablette:='TTARTICLE'                                                                    else
if Concept='COM'    then Begin Tablette:='TTGLOSSAIRE'     ; Plus:='GL_GESCOM="X"' ;                          End else
if Concept='VEN'    then Tablette:='TTCOMMERCIAL'                                                                 else
if Concept='REM'    then Tablette:='TTMOTIFREMISE'                                                                else
if Concept='REG'    then Begin Tablette:='TTMODEREGLE'     ; Plus:=' and MR_SPECIFPV="X"' ;                   End else
if Concept='FON'    then Tablette:='TTFONCTIONSPV'                                                                else
if Concept='OCT'    then Begin Tablette:='TTCAISSEOPE'     ; plus:='CP_NATURE="OCT" AND CP_MODIFIABLE<>"-"' ; end else
if Concept='OFC'    then Begin Tablette:='TTCAISSEOPE'     ; plus:='CP_NATURE="OFC"' ;                        end else
if Concept='MODNAT' then Begin Tablette:='TTTYPEPIECE'     ; plus:=' AND CO_LIBRE LIKE "PV%" AND (CO_CODE<>"VTA" AND CO_CODE<>"OCA")' ; end else
if Concept='MODMOD' then Begin Tablette:='TTMODELEETAT'    ; plus:='MO_TYPE="T" AND MO_NATURE="'+readtokenst(Extra)+'"' ;  End else
if Concept='MIP'    then Tablette:='TTIMPRIMEPIECE'                                                               else
if Concept='FPV'    then Begin Tablette:='TTFONCTIONSPAVE' ; if readtokenst(Extra)='X' then plus:=' and CO_LIBRE<>"INVISIBLE"' ; end else
if Concept='FSF'    then Tablette:=trim(gtfs('TTFAMILLE;TTSOUSFAMILLE;',';',1+ord(TRUEFALSEbo(readtokenSt(Extra))))) else
//XMG fin 05/10/01
        ;
{$ENDIF}
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Décomposition du champ 'CE_POLICE'
// ----------------------------------
// +-- Nom de la police
// |      +-- Style (Bold, Italic, Underline, Strikeout)
// |      |    +-- Taille
// |      |    |  +-- Couleur
// |      |    |  |       +-- Alignement (Left, Right, Center)
// |      |    |  |       |
// Tahoma;BIUS;24;8421376;C
//
Procedure AppliquePolice( TextePolice : String ; Font : TFont ; Var Alignment : TAlignment )  ;
Var St : String ;
Begin
if Trim(TextePolice) = '' then
   BEGIn
   if V_PGI.Tahoma then Font.Name:='Tahoma'
                   else Font.Name:='MS Sans Serif' ;
   Font.Style:=[];
   Font.Size:=8 ;
   Font.color:=clBlack ;
   Alignment:=taCenter ;
   END else
   BEGIN
   Font.Name:=ReadTokenSt(TextePolice) ;
   St:=ReadTokenSt(TextePolice) ;
   Font.Style:=[];
   if Pos('B', St) > 0 then Font.Style:=Font.Style+[fsBold] ;
   if Pos('I', St) > 0 then Font.Style:=Font.Style+[fsItalic] ;
   if Pos('U', St) > 0 then Font.Style:=Font.Style+[fsUnderline] ;
   if Pos('S', St) > 0 then Font.Style:=Font.Style+[fsStrikeOut] ;
   Font.Size:=StrToInt(ReadTokenSt(TextePolice)) ;
   Font.Color:=TColor(StrToInt(ReadTokenSt(TextePolice))) ;
   St:=ReadTokenSt(TextePolice) ;
   if St = 'L' then Alignment:=taLeftJustify else
    if St = 'R' then Alignment:=taRightJustify else
      if St = 'C' then Alignment:=taCenter ;
   END ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.MetForceBouton( Btn : TToolBarButton97 ; UnBtn : TOB )  ;
Var Align : TAlignment ;
    St    : String ;
Begin
if (not assigned(Btn)) or (Not Assigned(unBtn)) then exit ;
St:=UnBtn.GetValue('CE_POLICE') ;
if Trim(St) = '' then
   BEGIN
   if Btn.Parent is THPanel then
      BEGIN
      Btn.Font.Assign(THPanel(Btn.Parent).Font) ;
      END else
      BEGIN
      if V_PGI.Tahoma then Btn.Font.Name:='Tahoma'
                      else Btn.Font.Name:='MS Sans Serif' ;
      Btn.Font.Style:=[];
      Btn.Font.Size:=8 ;
      END ;
   Btn.Font.color:=clBlack ;
   if (Btn.Opaque) and (Btn.Color=clBlack) then Btn.Font.Color:=clWhite ; //NA 11/02/04
   Align:=taCenter ;
   END else
   BEGIN
   AppliquePolice(St, Btn.Font, Align) ;
   END ;
Case Align of
     taLeftJustify  : Begin Btn.Margin:=1 ;  Btn.Layout:=blGlyphLeft ;  End ;
     taRightJustify : Begin Btn.Margin:=1 ;  Btn.Layout:=blGlyphRight ; End ;
     taCenter       : Begin Btn.Margin:=-1 ; Btn.Layout:=blGlyphTop ;   End ;
     END ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//XMG debut 05/10/01
function TCLavierEcran.CalculeCleBouton ( UnBtn : TOB ) : integer ;
Begin
Result:=((Vinteger(UnBtn.Getvalue('CE_PAGE'))+1)*10000)+Vinteger(UnBtn.Getvalue('CE_NUMERO')) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.ChercheBtnparcle(Cle : Integer ) : Tob ;
var pg,num : Integer ;
Begin
pg:=(Cle div 10000)-1 ; Num:=cle mod 10000 ;
Result:=FBOutons.FindFirst(['CE_PAGE','CE_NUMERO'],[Pg,Num],FALSE) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//XMG fin 05/10/01
Procedure TClavierEcran.LieeBoutons( Btn : TToolBarButton97 ; UnBtn : TOB )  ;
Var St,Concept,Code : String ;
    Ind             : Integer ;
Begin
if (not assigned(Btn)) or (Not Assigned(unBtn)) then exit ;
Btn.Opaque:=(VInteger(UnBtn.GetValue('CE_COULEUR')) and $80000000=0) ;
if Btn.Opaque then Btn.Color:=TColor(VInteger(UnBtn.GetValue('CE_COULEUR'))) ;
Btn.Font.color:=clBlack ;
if (Btn.Opaque) and (Btn.Color=clBlack) then Btn.Font.Color:=clWhite ; //NA 11/02/04
MetForceBouton(Btn, UnBtn) ;
St:=vString(UnBtn.GetValue('VALCAPTION')) ;
// Césure définie par l'utilisateur le sigle '@' provoque un saut de ligne mais '@@' donne '@'.
for Ind:=1 to Length(St) do if St[Ind] = '@' then
    Begin
    if (Ind < Length(St)) and (St[Ind+1] = '@') then Delete(St, Ind, 1)
                                                else St[Ind]:=#10 ;
    End ;
Btn.Caption:=St ;
Btn.Hint:=TRUEFALSESt(VInteger(UnBtn.GetValue('CE_PAGE'))=-1) ;
Btn.parentShowHint:=FALSE ;
Btn.ShowHint:=FALSE ;
DessineImage(Btn,UnBtn) ;
Btn.Visible:=TRUE ;
Btn.Tag:=CalculeCleBouton(UnBtn) ; //.GetIndex+1 ;  //XMG 05/10/01
st:=vstring(UnBtn.getvalue('CE_TEXTE')) ;
Concept:=readtokenst(St) ;
Code:=readtokenst(st) ;
if (Concept='FPV') and (FBtnsStc.Indexof(Code)<0) and (Code<>'REN') then FBtnsStc.AddObject(Code,Btn) ;  //XMG 05/10/01
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.CalculeLeftBouton(Col : integer) : integer ;
Begin
Result:=MargeLeftBtn+(WidthBtn+SepHorzBtn)*Col  ;
if (ClcVisible) and (Result+WidthBtn>FCalculette.Left-MargeLeftClc) {and (Result<FCalculette.BoundsRect.Right+MargeRightClc) }then
   inc(Result,MargeLeftClc+FCalculette.Width+MargeRightClc) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.CalculeTopBouton(Lgn : integer) : integer ;
Begin
Result:=MargeTopBtn+(HeightBtn+SepVertBtn)*Lgn  ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.CreateUnButton( i : integer ) ;
Var Lig,col : Integer ;
    Btn     : TToolBarButton97 ;
    Nom     : String ;
    Ctrl    : TControl ; //XMG 10/01/02
Begin
//XMG 10/01/02 debut
Nom:='CEBtn'+inttostr(i) ; //XMG 05/10/01
ctrl:=TControl(FindComponent(Nom)) ;
if (Assigned(Ctrl)) and (((Ctrl is TToolBarButton97)=FALSE) or (TToolBarButton97(Ctrl).Visible)) then exit ; //Si le controle existe et il n'est pas un buton ou le bouton est déjà visible on sort (boutons bloqués sur toutes les pages)
if Assigned(Ctrl) then Btn:=TToolBarButton97(Ctrl) else
   Begin
   Lig:=DecomposenumeroBtn(i,FALSE) ;
   Col:=DecomposenumeroBtn(i,TRUE) ;
   Btn:=TToolBarButton97.Create(Self) ;
   //Btn.IntensiteLumiere:=0 ; //XMG 11/06/01
   Btn.Name:=Nom ;
   Btn.Parent:=Self ;
   (*{$IFDEF PGIS5}
   {$ELSE}*)
   Btn.S1Enabled:=False ;
   (*{$ENDIF}*)
   Btn.width:=WidthBtn ;
   Btn.Height:=HeightBtn ;
   Btn.Top:=CalculeTopBouton(Lig) ;
   Btn.Left:=CalculeLeftBouton(Col) ;
   Btn.Opaque:=FALSE ;
   Btn.DisplayMode:=dmBoth ;
   Btn.Caption:='' ;
   Btn.Layout:=blGlyphTop ;
   Btn.Alignment:=taCenter ;
   Btn.WordWrap:=TRUE ;
   Btn.Visible:=FALSE ;
   Btn.OnClick:=ClickBouton ;
   if Parametrage then
      Begin
      Btn.GroupIndex:=2 ;
      Btn.AllowAllUp:=FALSE ; //TRUE ;
      End ;
   end ;
//XMG 10/01/02 fin
GardeParamsVisuels(FVisuels,Btn,1) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function KBMakeSQL ( Concept, Code : String ; var CLe1, Libelle : String ) : String ;
Var Table,where,SQL : String ;
    saut            : Boolean ;
Begin
Result:='' ;
if (pos(';'+Concept+';',';ART;COM;RET;')>0) or (Code='') then
   Begin
   saut:=FALSE ; where:='' ;
{$IFDEF PGIS5}
   table:='ARTICLE' ; Libelle:='GA_LIBELLE' ;
   if Concept='COM' then Begin Table:='CHOIXEXT' ; Libelle:='YX_LIBELLE' ; where:='YX_TYPE="GCI"' ; saut:=TRUE; End ;
{$ELSE}
   table:='ARTICLE' ; Libelle:='A_LIBELLE' ;
   if Concept='COM' then Begin Table:='GLOSSAIR' ; Libelle:='GL_TEXTE' ; where:='GL_GESCOM="X"' ; End ;
{$ENDIF}
   Cle1:=TableToCle1(Table) ; if saut then readtokenStV(Cle1) ;
   Cle1:=readtokenStV(Cle1) ;
   SQL:='SELECT * FROM '+Table+' WHERE ' ;
   if trim(where)<>'' then SQL:=SQL+Where+' AND ' ;
   Result:=SQL+Cle1+'="'+Code+'"' ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//XMG debut 08/10/01
Procedure TClavierEcran.UnBoutonArticle ( Prop : TOB ; ChargeCaption : Boolean ) ;  
Var St,CLe1,Libelle,Concept,SQL,code,Cap,Tablette,Extra,Stg,plus,plusStg : String ;  //XMG 05/10/01
    Q                                                                    : TQuery ;
    okvide                                                               : Boolean ;
Begin
if ChargeCaption then
   Begin
   st:=vstring(Prop.getvalue('CE_TEXTE')) ;
   Concept:=readtokenst(St) ;
   OkVide:=(trim(St)='') ;
   Code:=readtokenst(st) ;
   Plus:='' ;
   Cap:=VString(Prop.GetValue('CE_CAPTION')) ;
   if not okVide then
      Begin
      if pos(';'+Concept+';',';ART;COM;RET;')>0 then
         Begin
         SQL:=KBMakeSQL(Concept, Code, Cle1, Libelle) ;
         Q:=OpenSQL(SQL,TRUE) ;
         //if Copy(Cap,1,2)=#255+#255 then //NA 06/01/04
         if IsKBIDCaptionStd(Cap) then //NA 06/01/04
            Begin
            //Delete(Cap,1,2) ; //NA 06/01/04
            if not Q.Eof then
               Begin
{$IFDEF PGIS5}
               if (pos(';'+Concept+';',';ART;RET;')>0) and (Cap='BAR') then Cap:=Q.FindField('GA_CODEBARRE').AsString else
{$ELSE}
               if (pos(';'+Concept+';',';ART;RET;')>0) and (Cap='BAR') then Cap:=Q.FindField('A_EAN13').AsString else
{$ENDIF}
               if Cap='LIB' then Cap:=Q.FindField(Libelle).AsString else
               if Cap='COD' then Cap:=Q.FindField(Cle1).AsString ;
               End else Cap:='' ;
            End ;
         Ferme(Q) ;
         End else
      if Concept='MOD' then
         Begin
         //if Copy(Cap,1,2)=#255+#255 then //NA 06/01/04
         if IsKBIDCaptionStd(Cap) then //NA 06/01/04
            Begin
            //Delete(Cap,1,2) ; //NA 06/01/04
            Extra:=ReadTokenSt(St) ;
//{$IFDEF PGIS5}                                                         //NA fin 06/11/02
//            Tablette:='GCIMPMODELEFO' ;
//            Stg:='GCFORMATTICKETFO' ;
//{$ELSE}
//XMG debut 05/10/01                                                    //NA fin 06/11/02
            donnetablette('MODMOD;'+Extra+';',Caisse,Tablette,plus) ;   //NA 13/03/02
            //Tablette:='TTMODELEETAT' ;
            //plus:='MO_TYPE="T" AND MO_NATURE="'+Extra+'"' ;

            DonneTablette('MODNAT',Caisse,StG,PlusStg) ;                //NA 13/03/02
            //Stg:='TTTYPEPIECE' ;
//XMG fin 05/10/01
//{$ENDIF}                                                              //NA fin 06/11/02
            If (Cap='LIB') then
               begin
               if (trim(extra)<>'') then Cap:=Rechdom(Stg,Extra,FALSE,plusStg)+' '+Rechdom(Tablette,Code,FALSE,Plus) else Cap:='' ;  //XMG 05/10/01
               end else
            if Cap='COD' then Cap:=Extra+' '+Code ;
            End ;
         End else
//XMG debut 20/07/01
      if Concept='FSF' then
         Begin
         //if Copy(Cap,1,2)=#255+#255 then //NA 06/01/04
         if IsKBIDCaptionStd(Cap) then //NA 06/01/04
            Begin
            //Delete(Cap,1,2) ; //NA 06/01/04
            Extra:=readtokenSt(St) ;
            If (Cap='LIB') then
               begin
               DonneTablette(Concept+';'+Extra+';',Caisse,Tablette,Stg) ;           //NA 13/03/02
               if (trim(Code)<>'') then Cap:=Rechdom(Tablette,Code,FALSE) else Cap:='' ;
               end else
            if Cap='COD' then
               Begin
{$IFDEF PGIS5}
{$ELSE}
               Extra:=traduirememoire(gtfs('Famille;Sousfamille;',';',1+ord(TRUEFALSEbo(Extra)))) ;
{$ENDIF}
               Cap:=Extra+' '+Code ;
               End ;
            End ;
         End else
//XMG fin 20/07/01
         Begin
         //if Copy(Cap,1,2)=#255+#255 then //NA 06/01/04
         if IsKBIDCaptionStd(Cap) then //NA 06/01/04
            Begin
            //Delete(Cap,1,2) ; //NA 06/01/04
            DonneTablette(Concept,Caisse,Tablette,Plus) ;          //NA 13/03/02
            if Concept='REM' then Code:=ReadTokenSt(St) ;
            If (Cap='LIB') then
               begin
               //XMG debut 03/10/01
               if (trim(code)<>'') then Cap:=Rechdom(Tablette,Code,FALSE)
                  else Cap:=TTtoLibelle(Tablette) ;
               //XMG fin 03/10/01
               end else if Cap='COD' then Cap:=Code ;
            End ;
         End ;
      End ;
   Prop.PutValue('VALCAPTION',Cap) ;
   End ;
//XMG fin 08/10/01
ChargeImage(Prop) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.ChargeImage( Prop : TOB ) ;
Var St,Concept,code,st2 : String ;
    Q                   : TQuery ;
    im,im2              : TPicture ;
    numIm               : Integer ;
    TJ                  : TJpegImage ;
{$IFDEF EAGLCLIENT}
    str                 : TStringStream ;
{$ELSE}
    str                 : TmemoryStream ;
{$ENDIF}
    IsJpeg              : Boolean ;
    st3, st4, st0 : string;  // JTR - Image libre stockées
Begin
if (Not Assigned(Prop)) then exit ;
NumIm  := VInteger(Prop.GetValue('IndexImage')) ;
st     := vstring(Prop.getvalue('CE_TEXTE')) ;
Concept:= readtokenst(St) ;
Code   := readtokenst(st) ;
St0    := VString(Prop.GetValue('CE_IMAGE')) ;
St     := ReadTokenSt(St0) ;
St2    := Readtokenst(St0) ;
// JTR - Image libre stockées
St3    := Readtokenst(St0) ;
St4    := Readtokenst(St0) ;
// Fin JTR - Image libre stockées
Im     := TPicture.Create ;
try
   if (pos(';'+Concept+';',';ART;RET;')>0) and (St='**PHOTO**') and (Trim(Code)<>'') then
   Begin
{$IFDEF PGIS5}
      Q:=OpenSQL('SELECT LO_OBJET,LO_QUALIFIANTBLOB FROM LIENSOLE WHERE LO_TABLEBLOB="GA" '
                +'AND LO_IDENTIFIANT="'+Code+'" AND LO_EMPLOIBLOB="'+ VH_GC.GCPHOTOFICHE +'" '
                +'AND (LO_QUALIFIANTBLOB="PHO" OR  LO_QUALIFIANTBLOB="PHJ" OR  LO_QUALIFIANTBLOB="VIJ")', TRUE) ;
{$ELSE}
      Q:=OpenSQL('SELECT A_PHOTO FROM ARTICLE WHERE A_ARTICLE="'+Code+'"',TRUE) ;
{$ENDIF}
      if Not Q.Eof then
         Begin
{$IFDEF PGIS5}
         IsJpeg:=((Q.Findfield('LO_QUALIFIANTBLOB').AsString='PHJ') or (Q.Findfield('LO_QUALIFIANTBLOB').AsString='VIJ')) ;
{$ELSE}
         IsJpeg:=False ;
{$ENDIF}
{$IFDEF EAGLCLIENT}     //NA debut 04/12/02
         str := TStringStream.Create(Q.Fields[0].AsString) ;
{$ELSE}
         str:=TmemoryStream.Create ;
         TBlobField(Q.Fields[0]).SaveToStream(str) ;
{$ENDIF}
         str.Seek(0,0) ;
         if IsJpeg then
            Begin
            TJ:=TJpegImage.create ;
            TJ.LoadFromStream(str) ;
            Im.Assign(TJ) ;
            TJ.free ;
            ///End else Im.Assign(Q.Fields[0]) ;
            End else Im.Bitmap.LoadFromStream(str) ;   //NA fin 04/12/02
         str.Free ;  //NA 21/03/03
         End ;
      Ferme(Q) ;
   end else if (St='**LIENSOLE**') then  // JTR - Image libre stockées
   begin
{$IFDEF PGIS5}
      if st4 = '' then st4 := '0';
      Q:=OpenSQL('SELECT LO_OBJET,LO_QUALIFIANTBLOB FROM LIENSOLE WHERE '
                 +'LO_TABLEBLOB = "'+st2+'" AND '
                 +'LO_IDENTIFIANT = "'+st3+'" AND '
                 +'LO_RANGBLOB = '+ st4, True);
{$ENDIF}
      if Not Q.Eof then
      begin
{$IFDEF PGIS5}
        IsJpeg:=((Q.Findfield('LO_QUALIFIANTBLOB').AsString='PHJ') or (Q.Findfield('LO_QUALIFIANTBLOB').AsString='VIJ')) ;
{$ELSE}
        IsJpeg:=False ;
{$ENDIF}
{$IFDEF EAGLCLIENT}     //NA debut 04/12/02
        str := TStringStream.Create(Q.Fields[0].AsString) ;
{$ELSE}
        str:=TmemoryStream.Create ;
        TBlobField(Q.Fields[0]).SaveToStream(str) ;
{$ENDIF}
         str.Seek(0,0) ;
         if IsJpeg then
         begin
           TJ:=TJpegImage.create ;
           TJ.LoadFromStream(str) ;
           Im.Assign(TJ) ;
           TJ.free ;
            ///End else Im.Assign(Q.Fields[0]) ;
         end else
           Im.Bitmap.LoadFromStream(str) ;   //NA fin 04/12/02
        str.Free ;  //NA 21/03/03
      end ;
      Ferme(Q) ;
   end else if (St='**INTERNE**') then
     FSupport.ImagesBtn.GetBitmap(Valeuri(St2),im.Bitmap)
   else if (trim(St)<>'') and (FileExists(St)) then
     Im.Loadfromfile(St) ;                          

   if (assigned(im.Graphic)) and (not im.Graphic.empty) then //Im.Width>0) and (Im.Height>0) then
      Begin
      im2:=TPicture.Create ;
      im2.Assign(Im) ;
      if NumIm=0 then
         Begin
         NumIm:=FImages.Add(im2) ;
         Prop.PutValue('IndexImage',abs(NumIm)+1) ;
         End else
         Begin
         TPicture(FImages[abs(NumIm)-1]).free ;
         FImages[abs(NumIm)-1]:=Im2 ;
         Prop.PutValue('IndexImage',abs(NumIm)) ;
         End ;
      End else if Numim>0 then Prop.PutValue('IndexImage',-NumIm) ;
 finally
   Im.Free ;
 End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.DessineImage( Btn : TToolBarButton97 ; Prop : TOB ) ;
Var im,Im2          : TPicture ;
    r               : TRect ;
    CoefH,CoefW     : Double ;
    NumIm,Min       : Integer ;
Begin
if (Not Assigned(Btn)) or (Not Assigned(Prop)) then exit ;
NumIm:=VInteger(Prop.GetValue('IndexImage')) ;
if NUmIm<=0 then
   Begin
   Btn.Glyph:=nil ; exit ;
   End ;
im:=FIMages[numIm-1] ;
if (Im.Width>0) and (Im.Height>0) then
  Begin
  im2:=TPicture.create ;
  try
    r:=Btn.ClientRect ;
    CoefH:=1 ; CoefW:=1 ;
    if ((R.Bottom/1.5)<im.height) or ((R.Right/1.5)<im.Width) then
       Begin
       min:=minintValue([R.bottom,r.Right]) ;
       CoefH:=Min/im.Height ;
       CoefW:=Min/Im.Width ;
       End ;
    r.Bottom:=Round(im.Height*CoefH) ;
    r.Right:=Round(im.Width*CoefW) ;
    //R.left:=(btn.width-r.Right) div 2 ;
    Im2.bitmap.Width:=R.Right ;
    im2.bitmap.height:=R.Bottom ;
    im2.bitmap.Canvas.StretchDraw(r,im.Graphic) ;
    Btn.Glyph.Assign(im2.bitmap)
   finally
    im2.free ;
   End ;
  End else Btn.Glyph:=Nil ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TClavierEcran.ChargeMaxPages ;
var i : integer ;
Begin
FMaxPages:=-1 ; i:=FBoutons.Detail.count ;
while i>0 do
   Begin
   If (pos(';'+Vstring(FBoutons.detail[i-1].GetValue('CE_TEXTE'))+';',';'+ConstVide+';')<=0) and  //XMG 09/10/01
      (FMaxPages<Vinteger(FBoutons.Detail[i-1].GetValue('CE_PAGE'))) then FMaxPages:=Vinteger(FBoutons.Detail[i-1].GetValue('CE_PAGE')) ;
   Dec(i) ;
   End ;
if FMaxPages=-1 then FMaxPages:=0 ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
//XMG debut 02/10/01
Procedure TClavierEcran.VidePave ;
Begin
FBoutons.ClearDetail ;
ChargeMaxPages ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.CreerChampsSupp ( UnBtn : TOB ) ;
var nb : Integer ;
Begin
with UnBtn do
  Begin
  nb:=VInteger(GetValue('CE_NUMERO')) ;
  AddCHampSup('VALCAPTION',FALSE) ; PutValue('VALCAPTION',GetValue('CE_CAPTION')) ;
  AddChampSup('IndexImage',FALSE) ; PutValue('IndexImage',0) ;
  AddChampSup('Ligne',FALSE)      ; PutValue('Ligne',DecomposeNumeroBtn(nb,FALSE)) ;
  AddChampSup('Colonne',FALSE)    ; PutValue('Colonne',DecomposeNumeroBtn(nb,TRUE)) ;
  End ;
End ;
//XMG fin 02/10/01
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.ChargeFromQ ( Q : TQuery ) : Boolean ;
var i            : Integer ;
    IsLocalGauge : Boolean ;
    Temp         : TOB ;
Begin
Result:=FALSE ;
if Not assigned(Q) then exit ;
//XMG debut 19/10/01
Temp:=TOB.Create('KBECRAN_S',nil,-1) ;
Result:=Temp.LoadDetailDB('CLAVIERECRAN','','',Q,TRUE,TRUE) ;
if result then
   Begin
   VidePave ;
   FBoutons.Dupliquer(Temp,TRUE,TRUE,TRUE) ;
   FBoutons.Modifie:=TRUE ;
   ChargeMaxPages ;
   IsLocalGauge:=(Assigned(Gauge)) ;
   if not IsLocalGauge then initmove(FBoutons.Detail.count,'') Else
      Begin
      Gauge.MaxValue:=FBoutons.Detail.count ;
      Gauge.MinValue:=0 ;
      Gauge.Progress:=0 ;
      Gauge.Visible:=TRUE ;
      End ;
   For i:=0 to FBoutons.Detail.count-1 do
       Begin
       if not IsLocalGauge then MoveCur(FALSE)
          else Gauge.Progress:=Gauge.Progress+1 ;
       FBoutons.Detail[i].PutValue('CE_CAISSE',Caisse) ;
       CreerChampsSupp(FBoutons.Detail[i]) ;
       UnBoutonArticle(FBoutons.Detail[i],TRUE) ;
       End ;
   if Not IsLocalGauge then FiniMove
      else FGauge.Visible:=FALSE ;
   End ;
Temp.Free ;
//XMG fin 19/10/01
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.ChargeClavier (Lacaisse : String ) ;
Var Q   : TQuery ;
Begin
if trim(LaCaisse)='' then exit ;
SourisSablier ;
Q:=OpenSQL('select * from CLAVIERECRAN where CE_CAISSE="'+LaCaisse+'" Order by CE_PAGE, CE_NUMERO',TRUE) ;
ChargeFromQ(Q) ;
FBoutons.SetAllModifie(FALSE) ;  //XMG 19/10/01 
Ferme(Q) ;
SourisNormale ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.ChargePageCourante ;
Var i,nb         : Integer ;
    UnBtn        : TOB ;
    nm,{OldBtn,}st : String ;  //XMG 08/10/02
    Btn          : TToolBarButton97 ;
    Ok           : Boolean ;
Begin
if (trim(Caisse)='') or (PageCourante=-1) then exit ;
{OldBtn:='' ;
//if Parametrage then
   Begin
   i:=0 ;
   while (i<=ControlCount-1) and (Trim(OldBtn)='') do
       Begin
       if (Controls[i] is TToolbarButton97) and (TToolBarButton97(Controls[i]).Down) then begin OldBtn:=TToolBarButton97(Controls[i]).Name ; Break ; end ;  //XMG 09/01/02
       inc(i) ;
       End ;
   End ;}  //XMG 10/01/02
{if FPageMenu then} VideBtnsStc ; //NA 23/06/04
CreeLesBoutonsPave ;
for i:=0 to ControlCount-1 do
  if copy(Controls[i].name,1,5)='CEBtn' then
     Begin
     TToolBarButton97(Controls[i]).Enabled:=True ;  //NA 21/03/03
     nm:=Controls[i].name ;
     Nb:=Valeuri(Copy(nm,6,length(nm)-5)) ;
     UnBtn:=FBoutons.findFirst(['CE_PAGE','CE_NUMERO'],[-1,Nb],FALSE) ;
     if not assigned(UnBtn) then UnBtn:=FBoutons.findFirst(['CE_PAGE','CE_NUMERO'],[PageCourante,Nb],FALSE) ;
     if Assigned(UnBtn) then
        Begin
        LieeBoutons(TToolBarButton97(Controls[i]),UnBtn) ;
        End else
        Begin
        if Not Parametrage then Controls[i].Visible:=FALSE else
           Begin
           St:='CE_PAGE,'+inttostr(PageCourante)+',|CE_CAISSE,'+Caisse+',|CE_NUMERO,'+inttostr(Nb)+',|'
              +'CE_TEXTE,'+ConstVide+',|CE_CAPTION,'+TraduireMemoire('Libre')+',|' ;  //XMG 09/10/01
           UnBtn:=CreateBtnStream(St,TRUE,TRUE) ;   //XMG 05/10/01
           LieeBoutons(TToolBarButton97(Controls[i]),UnBtn) ;
           End ;
        End ;
     End ;
Invalidate ;
ResizeClavier(nil) ;
if not Parametrage then
   Begin
   For i:=0 to FBtnsStc.Count-1 do
     if FBtnsSTC[i]<>'GOT' then
        Begin
        Btn:=TToolBarButton97(FBtnsStc.Objects[i]) ;
        Ok:=TRUE ;
        if (PageCourante<=0)       and (pos(';'+FBtnsStc[i]+';',';FST;PRV;')>0) then ok:=FALSE else
        if (PageCourante=MaxPages) then
           Begin
           if (FBtnsStc[i]='LST') or ((Not Parametrage) and (FBtnsStc[i]='NXT')) then ok:=FALSE ;
           End ;
        Btn.Enabled:=Ok ;
        End ;
   End else
{if trim(OldBtn)<>'' then
   Begin
   Btn:=TToolBarButton97(FindComponent(OldBtn)) ;
   if assigned(Btn) then Btn.click ;
   End ;}//XMG 10/01/02
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.CreeComboChargePage ;
Begin
if not Assigned(ComboPage) then
   Begin
   ComboPage:=THValComboBox.Create(Self) ;
   ComboPage.Vide:=FALSE ; //XMG 05/10/01
   ComboPage.Visible:=False ;
   ComboPage.Parent:=self ; //Screen.ActiveForm ;  //XMG 25/01/02
   End else
   Begin
   ComboPage.Exhaustif:=exnon ; 
   ComboPage.plus:='' ;
   Combopage.Datatype:='' ;

   ComboPage.Items.BeginUpdate ;
   ComboPage.Items.Clear ;
   ComboPage.Items.EndUpdate ;

   ComboPage.Values.BeginUpdate ;
   ComboPage.Values.Clear ;
   ComboPage.Values.EndUpdate ;
   End ;
ComboPage.tag:=0 ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.ChargePageMenu( Concept : String ; BtnAppel : TOB ) ;
Var tbltt,Plus : String ;
Begin
// Création d'un Combo pour récupérer le contenu de la tablette
CreeComboChargePage ;
DonneTablette(Concept,Caisse,Tbltt,Plus) ;             //NA 13/03/02
ComboPage.DataType:=Tbltt ;
ComboPage.Plus:=Plus ;
comboPage.Exhaustif:=exNon ;
if trim(plus)<>'' then comboPage.Exhaustif:=exPlus ; //Oui ;  //XMG 05/10/01
ChargePageMenuInterne(Concept,BtnAppel) ; //XMG 05/10/01
End ;
//XMG debut 05/10/01
//////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.ChargePageMenuInterne( Concept : String ; BtnAppel : TOB ) ; //NA 06/11/02
Var Nbtn,j,lig,col,AllerA  : Integer ;  //XMG 10/01/02
    UnBtn                  : TOB ;
    nm,st,clr,im,plc       : String ;
    okRentree,db,fn        : Boolean ;
    okDb, okFn             : Boolean ;    //NA 13/03/02
    Blocage                : Boolean ;    //NA 11/02/04
    Ctrl                   : TControl ;
Begin
j:=maxintValue([0,minintvalue([ComboPage.Tag,ComboPage.Items.count])]) ; okRentree:=FALSE ;
if (j>=ComboPage.Items.count) or (not assigned(BtnAppel)) or (not Assigned(ComboPage)) or (trim(Concept)='') or (Parametrage) then exit ;
Enabled:=FALSE ;
SourisSablier ;
Repeat
  UnBtn:=FBoutons.FindFirst(['CE_PAGE'],[9999],FALSE) ;
  if (Assigned(UnBtn)) then UnBtn.Free ;
until not assigned(UnBtn) ;
creelesboutonsPave ;
db:=(j>0) ;
fn:=(j+NbrMaxBoutons-1-2*ord(ComboPage.Items.count>NbrMaxBoutons-1)<ComboPage.Items.count) ;
Clr:=inttostr(VInteger(BtnAppel.GetValue('CE_COULEUR'))) ;
Im:=VString(BtnAppel.GetValue('CE_IMAGE')) ;
plc:=VString(BtnAppel.GetValue('CE_POLICE')) ;
AllerA:=Vinteger(BtnAppel.GetValue('CE_ALLERA')) ;  //XMG 10/01/02
OkDb:=(ComboPage.items.count<=NbrMaxBoutons-1) ; OkFn:=OkDb ;  //NA debut 13/03/02
if ComboPage.items.count>NbrMaxBoutons-1 then
   for lig:=0 to NbrBtnHeight-1 do
      for col:=0 to NbrBtnWidth-1 do
         Begin
         Nbtn:=ComposeNumeroBtn(Lig,Col) ;
         Ctrl:=TControl(FindComponent('CEBtn'+inttostr(Nbtn))) ;
         if Assigned(Ctrl) then
            Begin
            UnBtn:=ChercheBtnparcle(Ctrl.Tag) ;
            if assigned(UnBtn) then
               Begin
               nm:=UnBtn.GetValue('CE_TEXTE') ;
               Blocage:=(VInteger(UnBtn.GetValue('CE_PAGE'))=-1) ; //NA debut 11/02/04
               okDb:=(OkDb) or ((pos(';FPV;PRV;',';'+nm)>0) and (Blocage)) ;
               okFn:=(OkFn) or ((pos(';FPV;NXT;',';'+nm)>0) and (Blocage)) ; //NA fin 11/02/04
               End ;
            End ;
         End ;                                                 //NA debut 13/03/02
for lig:=0 to NbrBtnHeight-1 do
    for col:=0 to NbrBtnWidth-1 do
        Begin
        Nbtn:=ComposeNumeroBtn(Lig,Col) ;
        Ctrl:=TControl(FindComponent('CEBtn'+inttostr(Nbtn))) ;
        if Assigned(Ctrl) then //copy(Ctrl.name,1,5)='CEBtn') then
           Begin
           //if TToolBarButton97(Ctrl).Tag<=FBoutons.Detail.Count then
           {if Ctrl.tag<99990000 then
              Begin}
           UnBtn:=ChercheBtnparcle(Ctrl.Tag) ;
           if assigned(UnBtn) then
              Begin
              nm:=UnBtn.GetValue('CE_TEXTE') ;
              Blocage:=(VInteger(UnBtn.GetValue('CE_PAGE'))=-1) ; //NA 11/02/04
              Ctrl.Enabled:=TRUE ;
              okrentree:=(Okrentree) or (pos(';FPV;REN;',';'+nm)>0) ;
              st:=readtokenSt(nm) ; nm:=readtokenSt(nm) ;
              if (ComboPage.items.count<=NbrMaxBoutons-1) or (not Blocage) or //NA 11/02/04
                 (St <> 'FPV') or (pos(';'+nm+';',';PRV;NXT;REN;')<=0) then Ctrl.hint:='-' else
                 Begin
                 if (St='FPV') then
                    Begin
                    if nm='NXT' then Ctrl.enabled:=fn else
                    if nm='PRV' then Ctrl.enabled:=db else
                    if nm='LST' then Ctrl.enabled:=fn else
                    if nm='FST' then Ctrl.enabled:=db ;
                    End ;
                 continue ;
                 End ;
              end ;
           if (Ctrl.hint<>'X') then
              Begin
              if (Not OkRentree) and (Not PagePrinc) then    //NA 13/03/02
                 Begin
                 St:='CE_PAGE,9999,|CE_CAISSE,'+caisse+',|CE_NUMERO,'+inttostr(nbtn)+',|'
                    +'CE_IMAGE,**INTERNE**;37;,|CE_TEXTE,FPV;REN;,|' ;
                 UnBtn:=CreateBtnStream(St,FALSE,TRUE) ;
                 LieeBoutons(TToolBarButton97(Ctrl),UnBtn) ;
                 OkRentree:=TRUE ;
                 Continue ;
                 End ;
              if Not OkDb then    //NA debut 13/03/02
                 Begin
                 St:='CE_PAGE,9999,|CE_CAISSE,'+caisse+',|CE_NUMERO,'+inttostr(nbtn)+',|'
                    +'CE_IMAGE,**INTERNE**;50;,|CE_TEXTE,FPV;PRV;,|' ;
                 UnBtn:=CreateBtnStream(St,FALSE,TRUE) ;
                 LieeBoutons(TToolBarButton97(Ctrl),UnBtn) ;
                 Ctrl.enabled:=db ;
                 OkDb:=TRUE ;
                 Continue ;
                 End ;
              if Not OkFn then
                 Begin
                 St:='CE_PAGE,9999,|CE_CAISSE,'+caisse+',|CE_NUMERO,'+inttostr(nbtn)+',|'
                    +'CE_IMAGE,**INTERNE**;49;,|CE_TEXTE,FPV;NXT;,|' ;
                 UnBtn:=CreateBtnStream(St,FALSE,TRUE) ;
                 LieeBoutons(TToolBarButton97(Ctrl),UnBtn) ;
                 Ctrl.enabled:=fn ;
                 OkFn:=TRUE ;
                 Continue ;
                 End ;            //NA fin 13/03/02
              if j < ComboPage.Items.Count then
                 Begin
                 St:='CE_PAGE,9999,|CE_CAISSE,'+caisse+',|CE_NUMERO,'+inttostr(nbtn)+',|CE_CAPTION,'+ComboPage.items[j]+',|'
                    +'CE_COULEUR,'+Clr+',|CE_POLICE,'+Plc+',|'
                    +'CE_IMAGE,'+Im+',|' ;
                 if Concept='REM' then St:=St+'CE_TEXTE,'+Concept+';0;'+ComboPage.Values[j]+',|'
                                  else St:=St+'CE_TEXTE,'+Concept+';'+ComboPage.Values[j]+',|' ;
                 if AllerA>0 then St:=St+'CE_AlLERA,'+inttostr(Allera)+',|' ;  //XMG 10/02/01
                 UnBtn:=CreateBtnStream(St,FALSE,FALSE) ;
                 LieeBoutons(TToolBarButton97(Ctrl),UnBtn) ;
                 Inc(j) ;
                 End else Ctrl.Visible:=FALSE ;
              End ;
           End ;
        End ;
Invalidate ;
ComboPage.Tag:=j ; //+ord(j<ComboPage.items.count) ;
if ConceptMenu<>Concept then ConceptMenu:=Concept ;
if BtnMenu<>BtnAppel then BtnMenu:=BtnAppel ;
FPageMenu:=(Assigned(BtnMenu)) and (trim(ConceptMenu)<>'') ; //TRUE ;
ResizeClavier(nil) ;
SourisNormale ;
Enabled:=TRUE ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.ChargePageMenuArt( Select : String ; BtnAppel : TOB ) ;
Var QQ      : TQuery ;
    Lib,Val : String ;
Begin
if (Select = '') or (UpperCase(Copy(Select, 0, 6))<>'SELECT') then exit ;
CreeComboChargePage ;
ComboPage.Items.BeginUpdate ;
ComboPage.Values.BeginUpdate ;
QQ := OpenSQL(Select, True) ;
while Not QQ.EOF do
  Begin
{$IFDEF PGIS5}
   Val:=QQ.FindField('GA_ARTICLE').AsString ;
{$ELSE}
   Val:=QQ.FindField('A_ARTICLE').AsString ;
{$ENDIF}
   Lib:=QQ.Fields[0].AsString ;
   ComboPage.Values.add(Val) ;
   ComboPage.Items.add(Lib) ;
  QQ.Next ;
  End ;
Ferme(QQ) ;
ComboPage.Items.EndUpdate ;
ComboPage.Values.EndUpdate ;
ChargePageMenuInterne('ART',BtnAppel) ;
End ;
//XMG fin 05/10/01

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
{$IFDEF CHR} //NA debut 17/11/03
Procedure TClavierEcran.ChargePageMenuRes( Select : String ; BtnAppel : TOB ) ;
Var QQ      : TQuery ;
    Lib,Val : String ;
Begin
if (Select = '') or (UpperCase(Copy(Select, 0, 6))<>'SELECT') then exit ;
CreeComboChargePage ;
ComboPage.Items.BeginUpdate ;
ComboPage.Values.BeginUpdate ;
QQ := OpenSQL(Select, True) ;
while Not QQ.EOF do
  Begin
   Val:=QQ.FindField('HDR_DOSRES').AsString ;
   Lib:=QQ.Fields[0].AsString+'@'+QQ.Fields[2].AsString+'@'+IntToStr(QQ.Fields[3].AsInteger);
   ComboPage.Values.add(Val) ;
   ComboPage.Items.add(Lib) ;
  QQ.Next ;
  End ;
Ferme(QQ) ;
ComboPage.Items.EndUpdate ;
ComboPage.Values.EndUpdate ;
ChargePageMenuInterne('RES',BtnAppel) ;
End ;
{$ENDIF} //NA fin 17/11/03

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.EnregClavier ;
Var i            : Integer ;
    IsLocalGauge : Boolean ; //XMG 30/11/01
Begin
if ismodified then //FBoutons.IsOneModifie then  //XMG 19/10/01
   Begin
   try
       _Begintrans ;
       ExecuteSQL('DELETE FROM CLAVIERECRAN WHERE CE_CAISSE="'+Caisse+'"') ;
       FBoutons.SetAllModifie(TRUE) ;
//XMG debut 30/11/01
       IsLocalGauge:=(Assigned(Gauge)) ;
       if not IsLocalGauge then initmove(FBoutons.Detail.count,'') Else
          Begin
          Gauge.MaxValue:=FBoutons.Detail.count ;
          Gauge.MinValue:=0 ;
          Gauge.Progress:=0 ;
          Gauge.Visible:=TRUE ;
          End ;
       For i:=FBoutons.Detail.Count-1 downto 0 do
           Begin
           if not IsLocalGauge then MoveCur(FALSE)
              else Gauge.Progress:=Gauge.Progress+1 ; 
           If pos(';'+Vstring(FBoutons.detail[i].GetValue('CE_TEXTE')),';'+ConstVide)<=0 then
              Begin
              FBoutons.detail[i].PutValue('CE_CAISSE',Caisse) ;
              FBoutons.detail[i].InsertorUpdateDB(FALSE) ;
              End ;
           End ; 
        if Not IsLocalGauge then FiniMove
           else FGauge.Visible:=FALSE ;
//XMG fin 30/11/01
        _CommitTrans ;
        FBoutons.setallmodifie(FALSE) ;
     Except
       _RollBack ;
       Raise ;
     End ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.isModified : Boolean ;
Begin
Result:=(FBoutons.Modifie) or (FBoutons.IsOneModifie) ; //XMG 19/10/01
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.GetImages : TimageList ;
Begin
Result:=nil ;
if Assigned(FSupport) then Result:=FSupport.ImagesBtn ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetCaisse (AValue : String ) ;
Begin
FCaisse:=AValue ;
if ChargeAut then ChargeClavier(Caisse) ; //XMG 02/10/01
ChargeMaxPages ;
PageCourante:=0 ;
FPageMenu:=FALSE ; 
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetParametrage (AValue : Boolean) ;
Begin
FParametrage:=AValue ;
if assigned(FCalculette) then FCalculette.enabled:=not Parametrage ; 
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetPageCourante (AValue : Integer ) ;
Begin
AValue:=maxintvalue([AValue,0]) ;
if (Not Parametrage) or (AValue>MaxPages+1) then AValue:=minintvalue([AValue,MaxPages]) ;
if (AValue=PageCourante) and (not (FPageMenu or Parametrage)) then exit ;  //XMG 05/10/01
//if (FBoutons.IsOneModifie) and (PGIAsk(MsgErrDefaut(710),MsgErrDefaut(768))=mrYes) then EnregClavier ;
FPageCourante:=AValue ;
if Parametrage then ChargeMaxPages ;
Application.processmessages ;
ChargePageCourante ;
FPageMenu:= FALSE ;  //XMG 10/01/02
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TClavierEcran.GereKeyPress( var key : char ) ;
begin
if FInuse then
   begin
   key:=#0 ;
   FinUse:=FALSE ;
   End ;
end ;
//////////////////////////////////////////////////////////////////////////////////////////////////
//XMG debut 29/11/01
Function TClavierEcran.ChercheBoutonKey ( Touche : String ) : TOB ;
Begin
Result:=nil ;
if Touche<>ConstToucheVide then Result:=FBoutons.FindFirst(['CE_TOUCHE'],[Touche],FALSE) ;
End ;
//XMG fin 29/11/01
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.GereKey(var Key: Word; Shift: TShiftState);
Var UnBouton : TOB ;
    St       : String ;
    oldk     : Word ;
    Btn      : TToolBarButton97 ;
Begin
FInUse:=FALSE ;
OldK:=Key ;
//if assigned(KLng) then st:=KLng.GereKey(Key) ;  //XMG 29/11/00 Conserver comme reference
FInUse:=(OldK<>Key) and (OldK<>0) ;
if key<>0 then
   Begin
   st:='' ;
//XMG debut 29/11/01
   if key in [ord('A')..ord('Z'),ord('0')..ord('9'),ord('$'),ord('*'),ord('%'),ord('>'),ord('<')] then st:=chr(Key)+';'
      else St:=inttostr(Key)+';' ;
   //if (length(st)=1) and (ord(st[1]) in [ord('a')..ord('z'),ord('A')..ord('Z'),ord('0')..ord('9')]) then st:=uppercase(st[1])+';' else st:='' ;
//XMG fin 29/11/01
   st:=TrueFalsest(ssCtrl in Shift)+';'+TrueFalsest(ssShift in Shift)+';'+TrueFalsest(ssAlt in Shift)+';'+st ;
   End ;
if trim(st)<>'' then
   begin
   UnBouton:=ChercheBoutonKey(St) ; //XMG 29/11/01
   if assigned(Unbouton) then
      begin
      FInUse:=TRUE ;
      Key:=0 ;
      Btn:=TToolBarButton97.Create(Self) ;
      try
         //Btn.IntensiteLumiere:=0 ; //XMG 11/06/01
         Btn.Tag:=CalculeCleBouton(UnBouton) ; //XMG 05/10/01 UnBouton.GetIndex+1 ;
         ClickBouton(Btn) ;
        Finally
         Btn.Free ;
        End ;
      end ;
   End ;
End ;
//XMG debut 20/07/01
/////////////////////////////////////////////////////////////////////////////////////////
procedure TClavierEcran.LanceRequeteFSF(Code,Extra : String ; BtnAppel : TOB) ;
var SQL,Ch,Pref : String ;
    SF          : Boolean ;
Begin
SF:=TRUEFALSEbo(ReadTokenSt(Extra)) ;
CH:='' ;
{$IFDEF PGIS5}
Ch:=gtfs('GA_FAMILLENIV1;GA_FAMILLENIV2;',';',1+ord(SF)) ;        //NA 17/11/03
{$ELSE}
Ch:=gtfs('A_FAMILLE;A_SOUSFAMILLE;',';',1+ord(SF)) ;
{$ENDIF}
if trim(CH)<>'' then
   Begin
   Pref:=extractprefixe(Ch) ;
   SQL:='select '+Pref+'_LIBELLE, '+pref+'_ARTICLE from '+prefixetotable(Pref)+' where '+ch+'="'+code+'"' ;
   ChargePageMenuArt(SQL, BtnAppel) ;
   End ;
End ;
//XMG fin 20/07/01
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function KBParamToDouble(Data: string): double; //NA debut 03/03/04
var
  OldSepDecimal, OldSepMillier: char;
begin
  // séparateurs French pour la fonction Valeur()
  OldSepDecimal := V_PGI.SepDecimal;
  OldSepMillier := V_PGI.SepMillier;
  V_PGI.SepDecimal := ',';
  V_PGI.SepMillier := ' ';
  try
    Result := Valeur(Data);
  finally
    V_PGI.SepDecimal := OldSepDecimal;
    V_PGI.SepMillier := OldSepMillier;
  end;
end;  //NA fin 03/03/04
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.ClickBouton ( Sender : TObject ) ;
Var bt                 : TToolBarButton97 ;
    UnBouton           : TOB ;
    Qte,Prix           : Double ;
    Allera             : Integer ; //XMG 08/10/01
    Concept,Code,Extra : String ;
    OkCalcul           : Boolean ;
Begin
if not (Sender is TToolBarButton97) then exit ;
Bt:=TToolBarButton97(Sender) ;
UnBouton:=ChercheBtnparcle(bt.Tag) ; //XMG 05/10/01  FBoutons.Detail[bt.tag-1] ;
if Parametrage then
  Begin
  if Assigned(ParamBouton) then ParamBouton(UnBouton) ;
  End else if assigned(UnBouton) then
  Begin
  Allera:=VInteger(UnBouton.GetValue('CE_ALLERA')) ;  //XMG 08/10/01
  Extra:=vString(UnBouton.GetValue('CE_TEXTE')) ;
  Concept:=ReadTokenSt(Extra) ;
  Qte:=0 ; Prix:=0 ; OkCalcul:=FALSE ;
  if Concept+';'=ConstVide then  //XMG 09/10/01
     Begin
     //Rien à faire
     end else
  if pos(';'+Concept+';',';ART;RET;')>0 then
     Begin
     Code:=ReadTokenSt(Extra) ;
     Qte:=KBParamToDouble(ReadTokenSt(Extra)) ; //NA 03/03/04
{$IFNDEF PGIS5}
     if Qte=0 then Qte:=1 ;
{$ELSE}                                                             //NA debut 23/10/02
     if (FCalculette.LCD.asstring<>'') and (Qte=0) then Qte:=1 ;    //NA fin 23/10/02
{$ENDIF}
     Qte:=ChercheQte(Qte) ; OkCalcul:=TRUE ;
     Prix:=KBParamToDouble(ReadtokenSt(Extra)) ; //NA 03/03/04
     if Code='' then ChargePageMenuArt(ReadTokenSt(Extra), UnBouton) ;
     End else
  if Concept='COM' then                                             //NA debut 06/11/02
     Begin
     Code:=ReadTokenSt(Extra) ;
     if Code='' then Code:='TOUS' ;
     End else                                                       //NA fin 06/11/02
  if Concept='REM' then
     Begin
     Prix:=KBParamToDouble(ReadtokenSt(Extra)); //NA 03/03/04
{$IFDEF PGIS5}                                                       //NA debut 23/10/02
     if (FCalculette.LCD.asstring<>'') and (Prix=0) then Prix:=1 ;
     Prix:=ChercheQte(Prix) ; OkCalcul:=TRUE ;
{$ENDIF}                                                            //NA fin 23/10/02
     Code:=ReadTokenSt(Extra) ;
     if (Code='') and (Prix=0) then Code:='TOUS' ;
     End else
  if Concept='VEN' then
     Begin
     Code:=ReadTokenSt(Extra) ;
     if Code='' then Code:='TOUS' ;
     End else
  if Concept='FON' then
     Begin
     Code:=ReadTokenSt(Extra) ;
     if (Code='061') or (Code='062') or (Code='063') then
        Begin
        Qte:=ChercheQte(ord(FCalculette.LCD.asstring<>'')) ; OkCalcul:=TRUE ;
        End ;
     if Code='' then Code:='TOUS' ;
     End else
  if Concept='REG' then
     Begin
     Code:=ReadTokenSt(Extra) ;
     Prix:=ChercheQte(ord(FCalculette.LCD.asstring<>'')) ; OkCalcul:=TRUE ;
     if (Code='') and (Prix=0) then Code:='TOUS' ;
     End else
  if Concept='OCT' then
     Begin
     Code:=ReadTokenSt(Extra) ;
     Prix:=KBParamToDouble(ReadTokenSt(Extra)) ; //NA 03/03/04
{$IFDEF PGIS5}                                                       //NA debut 23/10/02
     if (FCalculette.LCD.asstring<>'') and (Prix=0) then Prix:=1 ;
     Prix:=ChercheQte(Prix) ; OkCalcul:=TRUE ;
     if (Code='') and (Prix=0) then Code:='TOUS' ;                  //NA 06/11/02
{$ENDIF}                                                            //NA fin 23/10/02
     end else
  if Concept='OFC' then
     Begin
     Code:=ReadTokenSt(Extra) ;
     Prix:=KBParamToDouble(ReadTokenSt(Extra)) ; //NA 03/03/04
     end else
  if Concept='MOD' then
     Begin
     Code:=ReadTokenSt(Extra) ;
     end else
  if Concept='MIP' then
     Begin
     Code:=ReadTokenSt(Extra) ;
     end else
  if Concept='FPV' then
     Begin
     Code:=ReadTokenSt(Extra) ;
     Allera:=Valeuri(readtokenSt(Extra)) ; //XMG 15/04/02
     end else
{$IFDEF CHR} //NA debut 17/11/03
  if Concept='RES' then
     Begin
     Code:=ReadTokenSt(Extra) ;
     if Code='' then ChargePageMenuRes(ReadTokenSt(Extra), UnBouton) ;
     End else
{$ENDIF} //NA fin 17/11/03
  if Concept='FSF' then
     Begin
     Code:=ReadTokenSt(Extra) ;
     if Code<>'' then LanceRequeteFSF(Code,Extra,UnBouton) ;
//XMG 10/01/02 debut
     Concept:='' ;
     End ;
  if trim(Concept)<>'' then
     Begin
     if OkCalcul then FCalculette.BtnsCalculetteUP ;
     if Code = TraduireMemoire('TOUS') then ChargePageMenu(Concept, UnBouton) else
     if Concept='FPV' then Allerapage(Code,Allera) else  //XMG 15/04/02
        Begin
        if Assigned(LanceBouton) then
           Begin
           LanceBouton(Concept,Code,Extra,Qte,Prix) ;
           if (Allera>0) and ((Code<>'') or (not InString(Concept,['ART','RES']))) then //NA 17/11/03
              PageCourante:=VInteger(UnBouton.GetValue('CE_ALLERA'))-1 ;  //XMG 08/10/01
           End ;
        End ;
     End ;
//XMG 10/01/02 fin
  End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.ChercheQte ( Qte : Double ) : Double ;
Var Mul : Double ;
Begin
mul:=FCalculette.LCD.AsFloat ;
if Mul=0 then Mul:=1 ;
Result:=Qte*Mul ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.ChercheValeurCalculette ( Qte : Double ; Raz : boolean ) : Double ;  //NA debut 13/01/03
Begin
Result:=ChercheQte(Qte) ;
if Raz then FCalculette.BtnsCalculetteUP ;
End ;                                                                                       //NA fin 13/01/03
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.UpdateBouton ( UnBouton : TOB ) ;
Var Btn : TToolBarButton97 ;
Begin
If (not assigned(UnBouton)) or (UnBouton.NomTable<>'CLAVIERECRAN') then exit ;
Btn:=TToolBarButton97(FindComponent('CEBtn'+vString(UnBouton.GetValue('CE_NUMERO')))) ;
if assigned(Btn) then
   Begin
   //FBoutons.Detail[Btn.Tag].Assign(UnBouton) ;
   UnBoutonArticle(UnBouton,TRUE) ; //FBoutons.Detail[Btn.Tag-1]) ; //XMG 08/10/01
   LieeBoutons(Btn,unBouton) ; //FBoutons.Detail[Btn.Tag-1]) ;  //XMG 05/10/01
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.BloqueResize ( Nom : String) ;
var Trad : THSystemMenu ;
    i  : Integer ;
Begin
for i:=0 to Owner.ComponentCount-1 do
  if Owner.components[i] is THsystemmenu then
     Begin
     Trad:=THsystemMenu(Owner.Components[i]) ;
     if Trad.LockedCtrls.indexof(Nom)=-1 then
        Begin
        Trad.LockedCtrls.add(Nom) ;
        Break ;
        End ;
     End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.CreeLesboutonsPave ;
Var i,j,Lig,Col : integer ;
    Ok          : Boolean ;
    Ctrl        : TControl ;
Begin
//i:=0 ;  //XMG 08/10/02
FVisuels.BeginUpdate ;
//XMG 10/01/02 debut
//While ControlCount-i>0 do
for i:=0 to ControlCount-1 do
  Begin
  Ctrl:=Controls[i] ; //ControlCount-1-i] ;
  ok:=FALSE ;
  if (Ctrl is TToolBarButton97) then
     Begin
     j:=FVisuels.IndexOf(Ctrl.Name) ;
     if j>-1 then
        Begin
        TObject(FVisuels.objects[j]).Free ;
        FVisuels.delete(j) ;
        End ;
     ok:=(Ctrl.Hint<>'X')  ;
     End ;
  if ok then Ctrl.Visible:=FALSE  //Free
//     else inc(i) ;
  End ;
//XMG 10/01/02 fin
for i:=0 to ControlCount-1 do
  if Controls[i] is TToolBarButton97 then
     Begin
     Ctrl:=Controls[i] ;
     j:=Valeuri(Copy(Ctrl.Name,6,length(Ctrl.Name)-5)) ;
     Lig:=DecomposenumeroBtn(j,FALSE) ;  //XMG 05/10/01
     Col:=DecomposenumeroBtn(j,TRUE) ;  //XMG 05/10/01
     Ctrl.SetBounds(CalculeLeftBouton(Col),CalculeTopBouton(Lig),WidthBtn,HeightBtn) ;
     GardeParamsVisuels(FVisuels,Ctrl,1) ;
     End ;
FVisuels.EndUpdate ;
if ClcVisible then FCalculette.Left:=-1 ;
//for i:=0 to NbrMaxBoutons-1 do CreateUnButton(i) ;
if NbrMaxBoutons>0 then
   for lig:=0 to NbrBtnHeight-1 do
       for col:=0 to NbrBtnWidth-1 do
           Begin
           i:=ComposeNumeroBtn(Lig,Col) ;
           CreateUnButton(i) ;
           End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.GereResize (Visu : TParamsVisuels ; Ctrl : TControl ) ;
Begin
if (not assigned(Visu)) or (Not Assigned(Ctrl)) then exit ;
case Visu.FncRsz of
 1 : if (Ctrl is TToolBarButton97) and (Ctrl.Visible) and (Ctrl.Tag>0) then DessineImage(TToolBarButton97(Ctrl),ChercheBtnparcle(Ctrl.Tag)) ; //FBoutons.Detail[Ctrl.Tag-1]) ;  //XMG 05/10/01
 End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.ResizeClavier ( Sender : TObject ) ;
Begin
if (assigned(Parent)) then
   Begin
   if Parent.Tag>0 then Height:=parent.Tag else Height:=Parent.Height ;
   ResizelesClaviers(Self,FVisuels,WidthClavier,HeightClavier,GereResize) ;
   if ClcVisible then
      Begin
      FCalculette.Height:=ClientHeight ;
      FCalculette.Width:=ClientWidth div 4 ;
      End ;
   invalidate ;
   End ; 
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.BoutonSurAutresPages ( Btn : TOB) : Boolean ;
//XMG debut 03/10/01
Begin
Result:=BoutonSurAutresPagesInterne(VInteger(Btn.GetValue('CE_PAGE')),VInteger(Btn.GetValue('CE_NUMERO'))) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.BoutonSurAutresPagesInterne ( Page,Num : Integer ) : Boolean ;
Var Wrk : TOB ;
    i,j : Integer ;
Begin
i:=-1 ; j:=0 ;
repeat
  inc(i) ;
  if i=0 then Wrk:=FBoutons.FindFirst(['CE_NUMERO'],[Num],FALSE)
     else Wrk:=FBoutons.FindNext(['CE_NUMERO'],[Num],FALSE) ;
  inc(j,ord((Assigned(Wrk)) and (VInteger(wrk.GetValue('CE_PAGE'))<>page) and (VString(wrk.GetValue('CE_TEXTE'))<>ConstVide))) ;  
until (Not Assigned(Wrk)) ;
result:=(j>0) ;
End ;
//XMG fin 03/10/01
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.RecopieClavier( Lacaisse : String ) ;
          /////////////////////////////////////////////////////////////////////////////////////////
          Procedure ChargeParams ( Caisse : String ) ;
          Var Q  : TQuery ;
              St : String ;
          Begin
          St:='' ;
{$IFDEF PGIS5}                         //NA debut 13/03/02
          Q:=OpenSQL('Select GPK_PARAMSCE from PARCAISSE where GPK_CAISSE="'+Caisse+'"',TRUE) ;
{$ELSE}
          Q:=OpenSQL('Select PC_PARAMSCE from PCAISSE where PC_CAISSE="'+Caisse+'"',TRUE) ;
{$ENDIF}                              //NA fin 13/03/02
          if Not Q.Eof then St:=Q.Fields[0].AsString ;
          Ferme(Q) ;
          if Trim(St)<>'' then
             Begin
             if (trim(St)='') or (NbCarInString(St,';')<>3) then St:=inttostr(CENbrBtnWidthDef)+';'+inttostr(CENbrBtnHeightDef)+';'+inttostr(ord(pcLeft))+';' ;
             NbrBtnWidth:=valeuri(readtokenst(St))  ;
             NbrBtnHeight:=Valeuri(ReadTokenst(St)) ;
             ClcPosition:=tPosClc(Valeuri(readtokenst(St))) ;
             End ;
          End ;
          /////////////////////////////////////////////////////////////////////////////////////////
var i : integer ;
Begin
if trim(Lacaisse)='' then exit ;
ChargeParams(LaCaisse) ;
ChargeClavier(Lacaisse) ;
for i:=0 to FBoutons.Detail.Count-1 do FBoutons.Detail[i].PutValue('CE_CAISSE',Caisse) ;
FBoutons.Modifie:=TRUE ;           //NA 13/03/02
ChargePageCourante ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//XMG debut 05/10/01
Procedure TClavierEcran.AlleraPage(BtnName : String ; Page : Integer ) ; //XMG 15/04/02
var pg,BtnsLbr,pgMax : integer ;
Begin
CalculeWidthBouton ; CalculeHeightBouton ;
BtnsLbr:=0 ; PgMax:=0 ;
if FpageMenu then
   Begin
   BtnsLbr:=NbrMaxBoutons-3 ; //les boutons fixes sur une page menu  (
   PgMax:=(ComboPage.Items.count div BtnsLbr)+ ord((ComboPage.Items.count mod BtnsLbr)<>0) ;
   End ;
If BtnName='REN'   then
   Begin
   if FPageMenu then
      Begin
      PageCourante:=PageCourante ;
      FPageMenu:= FALSE ;
      BtnMenu:=nil ;
      ConceptMenu:='' ;
      End ;
   end else
If BtnName='PRV'   then
   begin
   if FPageMenu then
      Begin
      Combopage.tag:=Combopage.tag-(BtnsLbr*2) ;
      ChargePageMenuInterne(ConceptMenu,BtnMenu) ;
      End else PageCourante:=PageCourante-1 ;
   end else
If BtnName='NXT'   then
   begin
   if FPageMenu then ChargePageMenuInterne(ConceptMenu,BtnMenu)
      else PageCourante:=PageCourante+1 ;
   end else
If BtnName='FST'   then
   Begin
   if FPageMenu then
      Begin
      ComboPage.tag:=0 ;
      ChargePageMenuInterne(ConceptMenu,BtnMenu) ;
      End else PageCourante:=0 ;
   End else
If BtnName='LST'   then
   Begin
   if FPageMenu then
      Begin
      Combopage.tag:=(BtnsLbr*PgMax)-BtnsLbr ;
      ChargePageMenuInterne(ConceptMenu,BtnMenu) ;
      end else PageCourante:=MaxPages ;
   End else
If BtnName='GOT'   then
   Begin
//XMG 15/04/02 début
   pg:=Page ;
   if Pg<=0 then
      Begin
      Pg:=FCalculette.LCD.AsInteger ;
      FCalculette.BtnsCalculetteUP ;
      End ;
   dec(pg,ord(pg>0)) ;
//XMG 15/04/02 fin
   if FPageMenu then
      Begin
      ComboPage.Tag:=BtnsLbr*minintValue([pg,PgMax-1]) ;
      ChargePageMenuInterne(ConceptMenu,BtnMenu) ;
      End else PageCourante:=Pg ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.VerifiesupBtns(Newval : integer ; isWidth : Boolean ) : Boolean ;
var i,vl,j : integer ;
    Btn    : TOB ;
    ok     : Boolean ;
    ch     : String ;
    ctrl   : TToolBarButton97 ;
Begin
result:=TRUE ;
if isWidth then Begin vl:=NbrBtnWidth  ; ch:='colonne' ; end    //vl=Colonne  avl=Ligne
           else Begin vl:=NbrBtnHeight ; ch:='Ligne'   ; end ;  //vl=Ligne    avl=Colonne
if vl<=NewVal then exit ;
OK:=FALSE ;
for i:=newval to vl-1 do
  Begin
  j:=0 ;
  repeat
     inc(j) ;
     if j=1 then btn:=FBoutons.FindFirst([ch],[i],FALSE)
        else btn:=FBoutons.FindNext([ch],[i],FALSE) ;
     ok:=(not assigned(Btn)) or (VString(Btn.GetValue('CE_TEXTE'))<>ConstVide) ;
  until (Ok) ;
  if Ok then break ;
  End ;
result:=ok ;
if (not Ok) or (PGIAsk('La nouvelle taille choisie oblige à supprimer des boutons.'+#13+' Désirez-vous continuer ?',''{parent.caption})=mryes) then    //NA 29/07/2002
   Begin
   for i:=newval to vl-1 do
     repeat
        if Result then btn:=FBoutons.FindFirst([ch],[i],FALSE)
           else btn:=FBoutons.FindFirst([ch,'CE_TEXTE'],[i,ConstVide],FALSE)  ;
        ok:=assigned(Btn) ;
        if Ok then
           Begin
           Ctrl:=TToolBarButton97(FindComponent('CEBtn'+vString(Btn.GetValue('CE_NUMERO')))) ;
           if assigned(Ctrl) then
               Ctrl.Free ;
           Btn.Free ;
           End ;
     until not Ok ;
   result:=TRUE ;
   End ;
End ;
//XMG fin 05/10/01
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetNbrBtnWidth(AValue : Integer ) ;
Begin
if (AValue=NbrBtnWidth) or (Not VerifieSupBtns(AValue,TRUE)) then exit ;
FNbrBtnWidth:=AValue ;
CalculeNbrBoutons ;
CalculeWidthBouton ;
ChargePageCourante ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetNbrBtnHeight(AValue : Integer ) ;
Begin
if (AValue=NbrBtnHeight) or (Not VerifieSupBtns(AValue,FALSE)) then exit ;
FNbrBtnHeight:=AValue ;
CalculeNbrBoutons ;
CalculeHeightBouton ;
ChargePageCourante ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.CalculeNbrBoutons ;
Begin
FNbrMaxBoutons:=NbrBtnWidth*NbrBtnHeight ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.PlaceWidthClavier(SiClc : Boolean = TRUE ) : integer ;
Begin
Result:=ClientWidth-MargeLeftBtn-MargeRightBtn ;
if (SiClc) and (ClcVisible) then Result:=Result-MargeLeftClc-FCalculette.Width-MargeRightClc ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.PlaceHeightClavier : integer ;
Begin
Result:=ClientHeight-MargeTopBtn-MargeBottomBtn ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.CalculetailleClavier ;
Begin
FWidthClavier:=MargeLeftBtn+(WidthBtn+SepHorzBtn)*NbrBtnWidth-SepHorzBtn+MargeRightBtn ;
if ClcVisible then FWidthClavier:=WidthClavier+MargeLeftClc+FCalculette.Width+MargeRightClc ;
FHeightClavier:=MargetopBtn+(HeightBtn+SepVertBtn)*NbrBtnHeight-SepVertBtn+MargeBottomBtn ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetMargeTopBtn (AValue : Integer ) ;
Begin
if AValue=MargeTopBtn then exit ;
FMargeTopBtn:=AValue ;
CalculeHeightBouton ;
ChargePageCourante ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetMargeBottomBtn (AValue : Integer ) ;
Begin
if AValue=MargeBottomBtn then exit ;
FMargeBottomBtn:=AValue ;
CalculeHeightBouton ;
ChargePageCourante ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetMargeLeftBtn (AValue : Integer ) ;
Begin
if AValue=MargeLeftBtn then exit ;
FMargeLeftBtn:=AValue ;
CalculeWidthBouton ;
ChargePageCourante ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetMargeRightBtn (AValue : Integer ) ;
Begin
if AValue=MargeRightBtn then exit ;
FMargeRightBtn:=AValue ;
CalculeWidthBouton ;
ChargePageCourante ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetSepHorzBtn (AValue : Integer ) ;
Begin
if AValue=SepHorzBtn then exit ;
FSepHorzBtn:=AValue ;
CalculeWidthBouton ;
ChargePageCourante ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetSepVertBtn (AValue : Integer ) ;
Begin
if AValue=SepVertBtn then exit ;
FSepVertBtn:=AValue ;
CalculeHeightBouton ;
ChargePageCourante ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.CalculeWidthBouton ;
Begin
if Not Assigned(Parent) then exit ;
if NbrBtnWidth=0 then begin NbrBtnWidth:=1 ; exit ; end ;
FWidthBtn:=(PlaceWidthClavier-SepHorzBtn*NbrBtnWidth+SepHorzBtn) div NbrBtnWidth;
CalculetailleClavier ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.CalculeHeightBouton ;
Begin
if Not Assigned(Parent) then exit ;
if NbrBtnHeight=0 then begin NbrBtnHeight:=1 ; exit ; end ;
FHeightBtn:=(PlaceHeightClavier-SepVertBtn*NbrBtnHeight+SepVertBtn) div NbrBtnHeight ;
CalculetailleClavier ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetOnSaisieClc ( AValue : TOnSaisieClc) ;
Begin
FCalculette.OnSaisieClc:=AValue ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetOnActionClc ( AValue : TOnActionClc) ;
Begin
FCalculette.OnActionClc:=AValue ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.GetOnSaisieClc : TOnSaisieClc ;
Begin
Result:=FCalculette.OnSaisieClc ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.GetOnActionClc : TOnActionClc ;
Begin
Result:=FCalculette.OnActionClc ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetOnChangeClc ( AValue : TOnChangeClc) ; //NA debut 17/11/03
Begin
FCalculette.LCD.OnSetVal:=AValue ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.GetOnChangeClc : TOnChangeClc ;
Begin
Result:=FCalculette.LCD.OnSetVal ;
End ; //NA fin 17/11/03
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetClcVisible ( AValue : Boolean ) ;
Begin
If ClcVisible=AValue then exit ;
FCalculette.Visible:=AValue ;
CalculeWidthBouton ;
ChargePageCourante ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TClavierEcran.GetClcVisible : Boolean ;
Begin
Result:=FCalculette.Visible ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.SetClcPos( AValue : TPosClc) ;
Begin
If ClcPosition=AValue then exit ;
FCalculette.Position:=AValue ;
if ClcVisible then
   Begin
   CalculeWidthBouton ;
   ChargePageCourante ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TClavierEcran.GetClcPos : TPosClc ;
Begin
Result:=FCalculette.Position ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TCLavierEcran.ExisteBouton ( Page,Ligne,Colonne : integer ) : Boolean ;
Var NNum : integer ;
    Btn  : TOB ;
Begin
Result:=FALSE ;
if (not parametrage) then exit ;
NNum:=ComposeNumeroBtn(Ligne,Colonne) ; //XMG 03/10/01
if page=-1 then Btn:=FBoutons.FindFirst(['CE_NUMERO'],[NNum],FALSE) else
   Begin
   Btn:=FBoutons.FindFirst(['CE_PAGE','CE_NUMERO'],[-1,NNum],FALSE) ;
   if Not Assigned(Btn) then Btn:=FBoutons.FindFirst(['CE_PAGE','CE_NUMERO'],[page,NNum],FALSE) ;
   End ;
Result:=(Assigned(Btn)) and (VString(Btn.GetValue('CE_TEXTE'))<>ConstVide) ;  //XMG 09/10/01
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.DeplaceBoutons( OPage,NPage,ONum,NNum : Integer ; Var Err : TPGIErrS1 ) : Integer ; //XMG 07/01/04
Var wpage,l,c,APage,i : integer ;
    Wrk,Btn           : TOB ;
    isBtnFixe         : Boolean ;
Begin
Result:=0 ;
//XMG debut 03/10/01
l:=decomposeNumeroBtn(ONum,FALSE) ;
c:=DecomposeNumeroBtn(ONum,TRUE) ;
//XMG fin 03/10/01
Btn:=FBoutons.FindFirst(['CE_PAGE','CE_NUMERO'],[oPage,ONum],FALSE) ;
isBtnFixe:=(Npage=-1) ;
if not isBtnFixe then
   Begin
   Wrk:=FBoutons.FindFirst(['CE_NUMERO'],[NNum],FALSE) ;
   IsBtnFixe:=(Assigned(Wrk)) and (VInteger(Wrk.GetValue('CE_PAGE'))=-1) ;
   End ;
i:=-1 ;
Repeat
  inc(i) ;
  if (IsBtnFixe) then
     Begin
     if i=0 then Wrk:=FBoutons.FindFirst(['CE_NUMERO'],[NNum],FALSE)
        else Wrk:=FBoutons.FindNext(['CE_NUMERO'],[NNum],FALSE) ;
     end else
     Begin
     if i=0 then Wrk:=FBoutons.FindFirst(['CE_PAGE','CE_NUMERO'],[NPage,NNum],FALSE)
        else Wrk:=FBoutons.FindNext(['CE_PAGE','CE_NUMERO'],[NPage,NNum],FALSE) ;
     End ;
  if assigned(Wrk) then
     Begin
     WPage:=OPage ;
     APage:=VInteger(Wrk.GetValue('CE_PAGE')) ;
     if (Apage=-1) and (OPage<>-1) then
        Begin
        Result:=DeplaceBoutons(APage,APage,NNum,ONum,Err) ;
        Err.Code:=789 ; Err.Libelle:='' ;
        Wrk:=nil ;
        Btn:=nil ;
        End else
        Begin
        inc(result) ;
        if WPage=-1 then
           Begin
           WPage:=VInteger(Wrk.GetValue('CE_PAGE')) ;
           if (Err.Code=0) and (VString(Wrk.GetValue('CE_TEXTE'))<>ConstVide) then Err.Code:=790 ;  //XMG 09/10/01
           End ;
        if WPage<>VInteger(Wrk.GetValue('CE_PAGE')) then Wrk.PutValue('CE_PAGE',WPage) ;
        if L<>VInteger(Wrk.GetValue('Ligne')) then Wrk.PutValue('Ligne',l) ;
        if C<>VInteger(Wrk.GetValue('Colonne')) then Wrk.PutValue('Colonne',c) ;
        if oNum<>VInteger(Wrk.GetValue('CE_NUMERO')) then Wrk.PutValue('CE_NUMERO',ONum) ;
        End ;
     End ;
until (not assigned(Wrk)) ;
if Assigned(Btn) then
   Begin
   l:=DecomposenumeroBtn(NNum,FALSE) ; //XMG 05/10/01
   c:=DecomposenumeroBtn(NNum,TRUE) ;  //XMG 05/10/01 
   if NPage<>VInteger(Btn.GetValue('CE_PAGE'))     then Btn.PutValue('CE_PAGE',NPage) ;
   if l<>VInteger(Btn.GetValue('Ligne'))           then Btn.PutValue('Ligne',l) ;
   if c<>VInteger(Btn.GetValue('Colonne'))         then Btn.PutValue('Colonne',c) ;
   if NNum<>VInteger(Btn.GetValue('CE_NUMERO'))    then Btn.PutValue('CE_NUMERO',NNum) ;
   End ;
Err.COde:=Err.Code*ord(Result>1) ;
if (Err.Code<>0) and (trim(Err.Libelle)='') then Err.Libelle:=MsgErrDefautS1(Err.Code) ; //XMG 07/01/04
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.MoveBouton (OldPage,OldNumero,NewPage,NewLigne,NewCol : integer ; var Err : TPGIErrS1 ) : Boolean ; //XMG 07/01/04
Var NewNumero : Integer ;
    Btn       : TToolBarButton97 ;
Begin
Result:=FALSE ;
err.code:=0 ; err.Libelle:='' ;
if not parametrage then exit ;
NewNumero:=ComposeNumeroBtn(NewLigne,NewCol) ; //XMG 03/10/01
if (OldPage<>NewPage) or (OldNumero<>NewNumero) then DeplaceBoutons(OldPage,NewPage,OldNumero,NewNumero,Err) ;
VideBtnsStc ;
if NewPage>-1 then PageCourante:=NewPage
   else ChargePageCourante ;
if oldPage=NewPage then
   Begin
   Btn:=TToolBarButton97(FindComponent('CEBtn'+inttostr(NewNumero))) ;
   if Assigned(Btn) then Btn.Click ;
   End ;
Result:=(Err.Code=0) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
//XMG debut 02/10/01
Function TClavierEcran.AjouteBtnTOB ( UnBtn : TOB ; replace,ChargeCaption : boolean ) : Boolean ;
var oldBtn : TOB ;
    Nb,Pg  : integer ;
Begin
Result:=(not Assigned(UnBtn)) or (UnBtn.NomTable<>'CLAVIERECRAN') ;
if Result then exit ;
Nb:=Vinteger(UnBtn.GetValue('CE_NUMERO')) ;
Pg:=VInteger(UnBtn.GetValue('CE_PAGE')) ;
if replace then
   Begin
   if pg=-1 then OldBtn:=FBoutons.findFirst(['CE_NUMERO'],[Nb],FALSE) else
      Begin
      OldBtn:=FBoutons.findFirst(['CE_PAGE','CE_NUMERO'],[-1,Nb],FALSE) ;
      if not assigned(OldBtn) then OldBtn:=FBoutons.findFirst(['CE_PAGE','CE_NUMERO'],[Pg,Nb],FALSE) ;
      End ; 
   if Assigned(OldBtn) then OldBtn.Free ;
   End ;
UnBtn.ChangeParent(FBoutons,-1) ;
CreerChampsSupp(UnBtn) ;
UnBoutonArticle(UnBtn,ChargeCaption) ;
if (pg<>9999)and (Pg>MaxPages) then ChargeMaxPages ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.ComposeNumeroBtn ( Lin,Col : Integer ) : Integer ;
Begin
Result:=(Lin+1)*100 +(Col+1) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.DecomposenumeroBtn( NumBtn : Integer ; IsColonne : Boolean ) : Integer ;
Begin
if IsColonne then Result:=NumBtn mod 100-1
   else Result:=NumBtn div 100-1 ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.Btnsenpage ( NbrPage : Integer) : Integer ;
Begin
result:=VInteger(FBoutons.SommeSimple('CE_NUMERO',['CE_PAGE'],[NbrPage],TRUE)) ;
Result:=Result-VInteger(FBoutons.SommeSimple('CE_NUMERO',['CE_PAGE','CE_TEXTE'],[NbrPage,ConstVIDE],TRUE)) ;
if NbrPage>-1 then Result:=Result+VInteger(FBoutons.SommeSimple('CE_NUMERO',['CE_PAGE'],[-1],TRUE))
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.BtnLibreEnPage ( NbrPage : Integer ) : Integer ;
var IsLibre      : Boolean ;
    Lin,Col,coef : integer ;
Begin
Result:=-1 ;
if NbrMaxBoutons-BtnsEnPage(NbrPage)<=0 then exit ;
lin:=NbrBtnHeight-1 ; Col:=0 ; Coef:=-1 ;
//On commence par cherche un place libre sur la page indiquée
repeat
  isLibre:=not ExisteBouton(NbrPage,lin,col) ;
  if Not IsLibre then
     Begin  //on parcourt la page de Bas/Gauche à Haut/Droit
     inc(lin,coef) ;
     if (Lin<0) or (lin>=NbrBtnHeight) then
        Begin
        //lin:=NbrBtnHeight-1 ;
        coef:=-Coef ;
        inc(Lin,Coef) ;
        inc(Col) ;
        End ;
     if col>=NbrBtnWidth then Lin:=-1 ;
     End ;
until (IsLibre) or (Lin<0) or (Col>=NbrBtnWidth) ;

if IsLibre then Result:=ComposeNumeroBtn(Lin,Col) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function TClavierEcran.CreateBtnStream ( Parms : String ; replace,ChargeCaption : Boolean ) : TOB ;
var i,Nb        : Integer ;
    UnP,Nom,Val : String ;
    Ok          : Boolean ;
Begin
result:=nil ;
if Trim(Parms)='' then exit ;
Result:=TOB.Create('CLAVIERECRAN',nil,-1) ;
With Result do
  Begin
  PutValue('CE_PAGE',-2)  ;
  PutValue('CE_ALLERA',-1)  ;
  PutValue('CE_POLICE','') ;
  PutValue('CE_TOUCHE','-;-;-;;') ;
  //PutValue('CE_CAPTION',#255+#255+'LIB') ; //NA 06/01/04
  PutValue('CE_CAPTION',GetKBIDCaptionStd+'LIB') ; //NA 06/01/04

  PutValue('CE_COULEUR',integer($80000000)) ;
  End ;
nb:=nbCarInString(Parms,'|') ;
for i:=1 to nb do
  Begin
  UnP:=gtfs(Parms,'|',i) ;
  Nom:=gtfs(UnP,',',1) ;
  Val:=gtfs(Unp,',',2) ;
  //if (Result.FieldExists(Nom)) then
  Result.PutValue(Nom,Val) ;
  End ;
ok:=(Result.IsFieldModified('CE_PAGE'))  and (Result.IsFieldModified('CE_NUMERO')) and (Result.IsFieldModified('CE_CAISSE')) and
    (Result.IsFieldModified('CE_TEXTE')) and (Result.IsFieldModified('CE_CAPTION')) ;
if Ok then AjouteBtnTOB(Result,replace,ChargeCaption) else
   Begin
   Result.Free ;
   Result:=nil ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TClavierEcran.LanceChargePageMenu( Concept, TblTT, Plus : String ; Couleur : TColor ; Image : Integer ) ;  //NA debut 13/03/02
var //{i, nBtn}  : Integer ;  //XMG 08/10/02
    St       : String ;
    BtnAppel : TOB ;
BEGIN
// Création du bouton de référence de la page menu
BtnAppel:=TOB.Create('', Nil, -1);
BtnAppel.AddChampSupValeur('CE_COULEUR', Couleur) ;
if Image>=0 then St := '**INTERNE**;'+IntToStr(Image)+';' else St := '' ;
BtnAppel.AddChampSupValeur('CE_IMAGE', St) ;
BtnAppel.AddChampSupValeur('CE_POLICE', '') ;
BtnAppel.AddChampSupValeur('CE_ALLERA', 0) ;
// Création d'un Combo pour récupérer le contenu de la tablette
CreeComboChargePage ;
if (Concept<>'') and (Tbltt='') then                   //NA debut 03/07/02
   BEGIN
   DonneTablette(Concept,Caisse,Tbltt,St) ;
   if Plus='' then Plus:=St ;
   END ;                                               //NA fin 03/07/02
ComboPage.DataType:=Tbltt ;
ComboPage.Plus:=Plus ;
comboPage.Exhaustif:=exNon ;
if trim(plus)<>'' then comboPage.Exhaustif:=exPlus ;
if (PagePrinc) and (assigned(BtnMenu)) then BtnMenu.Free ; //NA debut06/11/02
BtnMenu:=BtnAppel ;
PagePrinc:=True ;
ChargePageMenuInterne(Concept,BtnAppel) ;
//BtnAppel.Free ;                                          //NA fin 06/11/02
END ;                                                                               //NA fin 13/03/02
//////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////
type TOptionsBtns = (obResumer) ;
     TSetOptionsBtns = Set of TOptionsBtns ;

Function CalculePage ( PageDef : Integer ; Var NBtns : integer ; pCE : TClavierEcran ; optionsBtns : TSetoptionsBtns = [] ) : Integer ;
var WNBtns,BtnFixes,mxbtns,i,BtnsPg,pgPlace,BtnsLbr : Integer ;
Begin
Result:=PageDef ;
BtnFixes:=pCE.BtnsEnpage(-1) ;
WNBtns:=NBtns ;
if Result=-1 then
   Begin
   mxbtns:=0 ; PgPlace:=-1 ;
   for i:=0 to pCE.MaxPages do
       Begin
       BtnsPg:=pCE.Btnsenpage(i) ;
       if MxBtns<BtnsPg then MxBtns:=BtnsPg ;
       if (BtnsPg+WNBtns<=pCE.NbrMaxBoutons) and (PgPlace=-1) then PgPlace:=i ;
       End ;
   if (MxBtns+WNBtns>pCE.NbrMaxBoutons) or ((BtnFixes+WNBtns)*2>pCE.NbrMaxBoutons) then
      Begin
      if PgPlace>-1 then Result:=PgPlace
         else Result:=CalculePage(0,NBtns,pCE,OptionsBtns) ;
      End ;
   End else
   Begin
   repeat
     BtnsPg:=pCE.BtnsEnPage(Result) ;
     BtnsLbr:=pCE.NbrMaxBoutons-BtnsPg ;
     if (BtnsLbr<WNBtns) and (BtnsLbr>0) and ((obResumer in OptionsBtns)=FALSE) then
        WNBtns:=BtnsLbr ;
     inc(Result,ord((BtnsPg+WNBtns>pCE.NbrMaxBoutons) or (BtnsLbr<=0))) ;
   until BtnsPg+WNBtns<=pCE.NbrMaxBoutons ;
   End ;
if (Result<>PageDef) and (obResumer in OptionsBtns) then
   Begin
   BtnsPg:=1 ;
   OptionsBtns:=OptionsBtns-[obResumer] ;
   PgPlace:=CalculePage(PageDef,BtnsPg,pCE,OptionsBtns) ;
   if PgPlace<Result then
      Begin
      NBtns:=1 ;
      Result:=PgPlace ;
      End ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function RecupereSelect ( Nat, Caisse : String ; InCount : Boolean = FALSE ) : String ;      //NA 13/03/02
var Ch : String ;
Begin
Result:='' ;
{$IFDEF PGIS5} //NA debut 13/03/02
if Nat='REG' then begin result:='MODEPAIE where MP_UTILFO="X"' ; if Not InCount then Result:=Result+' order by MP_MODEPAIE' ; End else
if Nat='VEN' then //NA debut 02/12/03
   begin
   if GetInfoParPiece(FOGetNatureTicket(False, False), 'GPP_FILTRECOMM') = '001' then
      Begin
      result:='COMMERCIAL,PARCAISSE where GCL_TYPECOMMERCIAL="VEN" and GCL_ETABLISSEMENT=GPK_ETABLISSEMENT and GPK_CAISSE="'+Caisse+'"';
      End else
      Begin
      result:='COMMERCIAL where GCL_TYPECOMMERCIAL="VEN"';
      End;
   result:=result+' and GCL_DATESUPP>"' + USDateTime(Date) + '"';
   if Not InCount then Result:=Result+' order by GCL_COMMERCIAL' ;
   End ; //NA fin 02/12/03
if InString(Nat,['OCT','OFC']) then begin result:='ARTICLE where GA_TYPEARTICLE="FI" AND GA_FERME="-"' ; if Not InCount then Result:=Result+' order by GA_CODEARTICLE ' ; End ; //NA 17/11/03
if Nat='FSF' then begin result:='CHOIXCOD where CC_TYPE="FN1"' ; if Not InCount then Result:=Result+' order by CC_CODE' ; End ; //NA 17/11/03
{$ELSE}       //NA fin 13/03/02
if Nat='REG' then begin result:='MODEREGLE where MR_SPECIFPV="X" and MR_PVFERME="-"' ; if Not InCount then Result:=Result+' order by MR_MODEREGLE' ; End else  //XMG 10/04/02
if Nat='VEN' then begin result:='COMMERC' ; if Not InCount then Result:=Result+' order by CM_COMMERCIAL' ; End ;
if InString(Nat,['OCT','OFC']) then begin result:='CAISSEOPE where CP_MODIFIABLE="X" and CP_NATURE="'+Nat+'"' ; if Not InCount then Result:=Result+' order by CP_CODEOPE ' ; End ;
if Nat='FSF' then begin result:='CHOIXCOD where CC_TYPE="FAM"' ; if Not InCount then Result:=Result+' order by CC_CODE' ; End ;
{$ENDIF}
if trim(Result)<>'' then
   Begin
   Ch:='*' ; if InCount then ch:='count('+ch+')' ;
   Result:='select '+ch+' from '+Result ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function CalculeNBtns ( Nature_S, Caisse_S : String ) : Integer ;      //NA 13/03/02
         ////////////////////////////////////////////////////////////////////////////////////////////////////
         Function CalculeParSQL ( Nat : String ) : Integer ;
         var Q : TQUery ;
         Begin
         Q:=OpenSQL(RecupereSelect(Nat,Caisse_S,TRUE),TRUE) ;          //NA 13/03/02
         Result:=Q.Fields[0].AsInteger ;
         Ferme(Q) ;
         End ;
         //////////////////////////////////////////////////////////////////////////////////////////////////
var i,nb : integer ;
    Nat  : String ;
begin
Result:=0 ;
if sright(Nature_s,1)<>';' then Nature_s:=Nature_S+';' ;
nb:=NbCarInString(Nature_S,';') ;
for i:=1 to nb do
  Begin
  Nat:=gtfs(Nature_s,';',i) ;
  if Nat='FPV' then Result:=Result+2 else
  if Nat='FON' then Result:=Result+4 else
  if Nat='VEN' then Result:=Result+CalculeParSQL(Nat) else
  if Nat='REG' then Result:=Result+CalculeparSQL(Nat) else
  if InString(Nat,['OCT','OFC']) then Result:=Result+CalculeparSQL(Nat) else
  if Nat='FSF' then Result:=Result+CalculeparSQL(Nat) else
  End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Procedure BtnsPave (pCE : TClavierEcran ; Caisse : String ) ;
//Boutons de gestion de pave (Pg. Avant, pg arriere
var i,nBtn : Integer ;
    St     : String ;
Begin
for i:=1 to 2 do
    Begin
    nbtn:=pCE.BtnLibreEnPage(-1) ;
    if nBtn>-1 then
       Begin
       St:='CE_CAISSE,'+Caisse+',|CE_PAGE,-1,|CE_NUMERO,'+inttostr(nbtn)+',|CE_TEXTE,FPV;'+gtfs('NXT;PRV;',';',i)+',|'
          +'CE_IMAGE,**INTERNE**;'+inttostr(48+i)+',|' ;
       pCE.CreateBtnStream(St,TRUE,TRUE) ;
       End ;
    End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Procedure BtnsFonction(pCE : TClavierEcran ; Caisse : String ) ;
//Boutons de gestion de pièces (nouveau ticket, nouvelle op. caisse, enreg. pièce attente, reprendre pièce attente
var i,nBtn,im,pg,Nbtns,Cnt,nb : Integer ;
    St                        : String ;
Begin
NBtns:=CalculeNbtns('FON', Caisse) ;      //NA 13/03/02
pg:=Calculepage(-1,NBtns,pCE) ;  cnt:=0 ;
for i:=1 to NBtns do
    Begin
    repeat
       nbtn:=pCE.BtnLibreEnPage(Pg) ;
       if NBtn=-1 then
          BEgin
          Nb:=NBtns-Cnt ;
          pg:=CalculePage(Pg,Nb,pCE) ;
          ENd ;
    until nBtn>-1 ;
    im:=49-i*ord(i<=2)-44*ord(i=3)-43*ord(i=4) ;
{$IFDEF PGIS5} //NA debut 17/11/03
    St:='CE_CAISSE,'+Caisse+',|CE_PAGE,'+inttostr(pg)+',|CE_NUMERO,'+inttostr(nbtn)+',|CE_TEXTE,FON;'+gtfs('112;114;251;558;',';',i)+',|'
       +'CE_IMAGE,**INTERNE**;'+inttostr(im)+',|' ;
{$ELSE}
    St:='CE_CAISSE,'+Caisse+',|CE_PAGE,'+inttostr(pg)+',|CE_NUMERO,'+inttostr(nbtn)+',|CE_TEXTE,FON;'+gtfs('005;007;015;016;',';',i)+',|'
       +'CE_IMAGE,**INTERNE**;'+inttostr(im)+',|' ;
{$ENDIF} //NA fin 17/11/03
    pCE.CreateBtnStream(St,TRUE,TRUE) ;
    inc(Cnt) ;
    End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Procedure BtnsRegle(pCE : TClavierEcran ; Caisse : String ) ;
//boutons des modes de réglement
var Q                    : TQuery ;
    nBtn,im,NBtns,pg,Cnt : Integer ;
    St                   : String ;
Begin
Q:=OpenSQL(RecupereSelect('REG', Caisse),TRUE) ;    //NA 13/03/02
{$IFDEF EAGLCLIENT}                                 //MR debut 23/10/02
NBTns:=Q.RecordCount ;
{$ELSE}
NBTns:=QCount(Q) ;
{$ENDIF}                                            //MR fin 23/10/02
pg:=CalculePage(-1,NBtns,pCE) ; cnt:=0 ;
while Not Q.Eof do
    Begin
    repeat
       nbtn:=pCE.BtnLibreEnPage(Pg) ;
       if NBtn=-1 then
          Begin
          NBtns:=NBtns-Cnt ;
          pg:=CalculePage(Pg,NBtns,pCE) ;
          End ;
    until nBtn>-1 ;
{$IFDEF PGIS5} //NA debut 13/03/02
    im:=26+ord(Q.FindField('MP_DEVISEFO').AsString<>V_PGI.DevisePivot)    // 1) Especes 2) si devise
          +(19*ord(Q.FindField('MP_DEVISEFO').AsString=FODonneCodeEuro)+ord(VH^.TenueEuro)) ;    // Si Euro
    if Q.FindField('MP_TYPEMODEPAIE').AsString=TYPEPAIECHEQUE   then im:=29 else    // Cheque
    if Q.FindField('MP_TYPEMODEPAIE').AsString=TYPEPAIECHQDIFF  then im:=29 else   // Cheque differe
    if Q.FindField('MP_TYPEMODEPAIE').AsString=TYPEPAIECB       then im:=42 else    // Carte blue
    if Q.FindField('MP_TYPEMODEPAIE').AsString=TYPEPAIEBONACHAT then im:=12 else ;  // Bon d'achat
    if Q.FindField('MP_TYPEMODEPAIE').AsString=TYPEPAIEARRHES   then im:=38 else ;  // Arrhes
    if Q.FindField('MP_TYPEMODEPAIE').AsString=TYPEPAIEAVOIR    then im:=8  else ;  // Avoirs
    if Q.FindField('MP_TYPEMODEPAIE').AsString=TYPEPAIERESTEDU  then im:=41 ;       // Reste dû

    St:='CE_CAISSE,'+Caisse+',|CE_PAGE,'+inttostr(pg)+',|CE_NUMERO,'+inttostr(nbtn)+',|CE_TEXTE,REG;'+Q.findField('MP_MODEPAIE').AsString+',|'
       +'CE_IMAGE,**INTERNE**;'+inttostr(im)+',|' ; //NA 17/11/03
{$ELSE}       //NA fin 13/03/02
    im:=26+ord(Q.FindField('MR_DEVISE').AsString<>V_PGI.DevisePivot)    // 1) Especes 2) si devise
          +(19*ord(Q.FindField('MR_DEVISE').AsString=VS1^.CodeEuro)+ord(V_PGI.TenueEuro)) ;    // Si Euro //XMG 23/02/04
    if Q.FindField('MR_TYPEREGLE').AsString='CHQ' then im:=29 else    // Cheque
    if Q.FindField('MR_TYPEREGLE').AsString='CB'  then im:=42 else    // Carte blue
    if Q.FindField('MR_TYPEREGLE').AsString='CCD' then im:=12 else ;  // Bon d'achat
    if Q.FindField('MR_TYPEREGLE').AsString='ARH' then im:=38 else ;  // Arrhes
    if Q.FindField('MR_TYPEREGLE').AsString='AVR' then im:=8  else ;  // Avoirs
    if Q.FindField('MR_TYPEREGLE').AsString='RTD' then im:=41 ;       // Reste dû

    St:='CE_CAISSE,'+Caisse+',|CE_PAGE,'+inttostr(pg)+',|CE_NUMERO,'+inttostr(nbtn)+',|CE_TEXTE,REG;'+Q.findField('MR_MODEREGLE').AsString+',|'
       +'CE_IMAGE,**INTERNE**;'+inttostr(im)+',|' ;
{$ENDIF} //NA 17/11/03
    pCE.CreateBtnStream(St,TRUE,TRUE) ;
    Q.next ;
    inc(Cnt) ;
    End ;
Ferme(Q) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Procedure BtnsVendeur(pCE : TClavierEcran ; Caisse : String ) ;
//Boutons des vendeurs (autant comme vendeur ou seuelement un bouton)
var Q                    : TQuery ;
    nBtn,pg,NBtns,Cnt,Nb : Integer ;
    St,Code              : String ;
Begin
NBtns:=CalculeNBtns('VEN', Caisse) ;                //NA 13/03/02
pg:=Calculepage(0,NBtns,pCE,[obResumer]) ;
Q:=OpenSQL(RecupereSelect('VEN', Caisse),TRUE) ;    //NA 13/03/02
{$IFDEF EAGLCLIENT}                                 //MR debut 23/10/02
if NBTns=Q.RecordCount then
{$ELSE}
if Nbtns<QCount(Q) then
{$ENDIF}                                            //MR fin 23/10/02
   Begin
   NBtns:=1 ;
   Ferme(q) ;
   Q:=nil ;
   End ;
cnt:=0 ;
while ((NBtns=1) and (Cnt<1)) or ((Assigned(Q)) and (Not Q.Eof)) do
    Begin
    repeat
       nbtn:=pCE.BtnLibreEnPage(Pg) ;
       if NBtn=-1 then
          Begin
          NB:=NBtns-Cnt ;
          pg:=CalculePage(Pg,NB,pCE) ;
          End ;
    until nBtn>-1 ;
{$IFDEF PGIS5} //NA debut 13/03/02
    Code:='' ; if assigned(Q) then Code:=Q.FindField('GCL_COMMERCIAL').AsString ;
{$ELSE}       //NA fin 13/03/02
    Code:='' ; if assigned(Q) then Code:=Q.FindField('CM_COMMERCIAL').AsString ;
{$ENDIF}
    St:='CE_CAISSE,'+Caisse+',|CE_PAGE,'+inttostr(pg)+',|CE_NUMERO,'+inttostr(nbtn)+',|CE_TEXTE,VEN;'+Code+';,|'
       +'CE_IMAGE,**INTERNE**;24,|' ;
    pCE.CreateBtnStream(St,TRUE,TRUE) ;
    if Assigned(Q) then Q.next ;
    inc(Cnt) ;
    End ;
if Assigned(Q) then ferme(Q) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
procedure BtnsOpeCaisse(pCE : TClavierEcran ; Caisse,Nature : String) ;
//Boutons des op. caisse et op. financières
var Q                    : TQuery ;
    im,nBtn,pg,NBtns,Cnt : Integer ;
    St                   : String ;
Begin
NBtns:=CalculeNBtns(Nature, Caisse) ;                                 //NA 13/03/02
pg:=Calculepage(0,NBtns,pCE) ;
Q:=OpenSQL(RecupereSelect(Nature, Caisse),TRUE) ; im:=9 ; Cnt:=0 ;    //NA 13/03/02
while (Not Q.Eof) do
    Begin
    repeat
       nbtn:=pCE.BtnLibreEnPage(Pg) ;
       if NBtn=-1 then
          Begin
          NBtns:=NBtns-Cnt ;
          pg:=CalculePage(Pg,NBtns,pCE) ;
          End ;
    until nBtn>-1 ;
{$IFDEF PGIS5} //NA debut 17/11/03
    St:='CE_CAISSE,'+Caisse+',|CE_PAGE,'+inttostr(pg)+',|CE_NUMERO,'+inttostr(nbtn)+',|CE_TEXTE,'+nature+';'+Q.FindField('GA_CODEARTICLE').AsString+';,|'
       +'CE_IMAGE,**INTERNE**;'+inttostr(im)+',|' ;
{$ELSE}
    St:='CE_CAISSE,'+Caisse+',|CE_PAGE,'+inttostr(pg)+',|CE_NUMERO,'+inttostr(nbtn)+',|CE_TEXTE,'+nature+';'+Q.FindField('CP_CODEOPE').AsString+';,|'
       +'CE_IMAGE,**INTERNE**;'+inttostr(im)+',|' ;
{$ENDIF} //NA fin 17/11/03
    pCE.CreateBtnStream(St,TRUE,TRUE) ;
    Q.next ;
    //inc(im) ;
    inc(Cnt) ;
    End ;
ferme(Q) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
procedure BtnsFamille(pCE : TClavierEcran ; Caisse : String) ;
//Boutons des familles
var Q                    : TQuery ;
    im,nBtn,pg,NBtns,Cnt : Integer ;
    St                   : String ;
Begin
NBtns:=CalculeNBtns('FSF', Caisse) ;                                  //NA 13/03/02
pg:=Calculepage(0,NBtns,pCE) ;
Q:=OpenSQL(RecupereSelect('FSF', Caisse),TRUE) ; im:=18 ; Cnt:=0 ;    //NA 13/03/02
while (Not Q.Eof) do
    Begin
    repeat
       nbtn:=pCE.BtnLibreEnPage(Pg) ;
       if NBtn=-1 then
          Begin
          NBtns:=NBtns-Cnt ;
          pg:=CalculePage(Pg,NBtns,pCE) ;
          End ;
    until nBtn>-1 ;
    St:='CE_CAISSE,'+Caisse+',|CE_PAGE,'+inttostr(pg)+',|CE_NUMERO,'+inttostr(nbtn)+',|CE_TEXTE,FSF;'+Q.FindField('CC_CODE').AsString+';-;,|'
       +'CE_IMAGE,**INTERNE**;'+inttostr(im)+',|' ;
    pCE.CreateBtnStream(St,TRUE,TRUE) ;
    Q.next ;
    //inc(im) ;
    inc(Cnt) ;
    End ;
ferme(Q) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function CreepaveDynamique ( pCE : TClavierEcran ; Caisse,Params : String ; TypePave : TSetTypePave ; OkGauge : Boolean ; Var Err :TPGIErrS1  ) : Boolean ; //XMG 07/01/04   02/10/01
         //////////////////////////////////////////////////////////////////////////////////////////////////
         function NaturespartypePave ( pCE : TClavierEcran ; typepave : TSetTypePave ) : String ;
         Begin
         result:='' ;
{$IFDEF PGIS5} //NA debut 17/11/03
         if tpEncaissement in TypePave then Result:=Result+'FON;VEN;REG;OCT;' ;
{$ELSE}
         if tpEncaissement in TypePave then Result:=Result+'FON;VEN;REG;OCT;OFC;' ;
         if tpFamilles in TypePave then Result:=Result+'FSF;' ;
{$ENDIF} //NA fin 17/11/03
         if CalculeNBtns(Result, Caisse)>=pCE.NbrMaxBoutons then Result:='FPV;'+Result ;    //NA 13/03/02
         End ;
         //////////////////////////////////////////////////////////////////////////////////////////////////
var CreerCE        : Boolean ;
    nb,i           : integer ;
    Nat,NatureBtns : String ;
    p              : TPosClc ;
Begin
Err.Code:=0 ; Err.Libelle:='' ;
result:=FALSE ;
if (trim(Params)='') or (NbCarInString(Params,';')<>3) then Params:=inttostr(CENbrBtnWidthDef)+';'+inttostr(CENbrBtnHeightDef)+';'+inttostr(ord(pcLeft))+';' ;
CreerCE:=not Assigned(pCE) ;
try
   if CreerCE then
      Begin
      pCE:=TClavierEcran.create(Application) ;
      pCE.Visible:=FALSE ;
      pCE.Parent:=Screen.ActiveCustomForm ;
      pCE.Parametrage:=TRUE ;
      pCE.ChargeAut:=FALSE ;
      pCE.Caisse:=Caisse ;
      End ;
   NatureBtns:=NaturesParTypePave(pCE,TypePAve) ;
   nb:=NbCarInString(NatureBtns,';') ;
   if (Valeuri(gtfs(Params,';',1))*Valeuri(gtfs(Params,';',2))<3) and (pos(';FPV;',';'+NatureBtns)>0) then
      Begin
      Err.Code:=-1 ;
      Err.Libelle:='Le nombre maximum de boutons est inferieur à 3 et il n''est pas possible de créer un pavé.' ;
      End else
      Begin
      pCE.VidePave ;
      i:=Valeuri(gtfs(Params,';',1))          ; if pCE.NbrBtnWidth<>i  then pCE.NbrBtnWidth:=i ;
      i:=Valeuri(gtfs(Params,';',2))          ; if pCE.NbrBtnHeight<>i then pCE.NbrBtnHeight:=i ;
      p:=TPosClc(valeuri(gtfs(Params,';',3))) ; if pCE.ClcPosition<>p  then pCE.ClcPosition:=p ;
      if OkGauge then initmove(nb,'') ;
      for i:=1 to nb do
        Begin
        if OkGauge then movecur(FALSE) ;
        Nat:=gtfs(NatureBtns,';',i) ;
        if Nat='FPV' then BtnsPave(pCE,Caisse)     else
        if Nat='FON' then BtnsFonction(pCE,Caisse) else
        if Nat='FSF' then BtnsFamille(pCE,Caisse) else
        if Nat='OCT' then BtnsOpeCaisse(pCE,Caisse,'OCT') else
        if Nat='OFC' then BtnsOpeCaisse(pCE,Caisse,'OFC') else
        if Nat='VEN' then BtnsVendeur(pCE,Caisse)  else
        if Nat='REG' then BtnsRegle(pCE,Caisse) ;
        End ;
      if OkGauge then finimove ;
      if creerCE then pCE.EnregClavier ;
      Result:=TRUE ;
      End ;
  Finally
   if creerCE then pCE.Free ;
  End ;
end ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Procedure LanceCreationAutomatiquePaves (Msg : String ) ;
var Q   : TQuery ;
    Err : TPGIErrS1 ; //XMG 07/01/04
    tpv : TSetTypepave ;
    OkG : Boolean ;
Begin
if ConfirmeCreationPave(Msg,FALSE) then
   Begin
   tpv:=[tpEncaissement] ;
{$IFDEF PGIS5} //NA debut 13/03/02
   Q:=Opensql('select GPK_CAISSE, GPK_PARAMSCE from PARCAISSE',TRUE) ;
{$ELSE}       //NA fin 13/03/02
   if VString(GetParamSoc('SO_TYPEPAVEDEFAUT'))<>'000' then tpv:=tpv+[tpFamilles] ;
   Q:=Opensql('select PC_CAISSE, PC_PARAMSCE from PCAISSE',TRUE) ;
{$ENDIF}
{$IFDEF EAGLCLIENT}                         //MR debut 23/10/02
   Okg:=(Q.RecordCount>1) ;
   if OkG then InitMove(Q.RecordCount,'') ;
{$ELSE}
   Okg:=QCount(Q)>1 ;
   if OkG then InitMove(QCount(Q),'') ;
{$ENDIF}                                    //MR fin 23/10/02
   while not q.eof do
     Begin
     if OkG then MoveCur(FALSE) ;
{$IFDEF PGIS5} //NA debut 13/03/02
     if not CreepaveDynamique(nil,Q.findField('GPK_CAISSE').AsString,Q.findField('GPK_PARAMSCE').AsString,tpv,not OkG, Err) then PGIBox(Err.Libelle,'') ;
{$ELSE}
     if not CreepaveDynamique(nil,Q.findField('PC_CAISSE').AsString,Q.findField('PC_PARAMSCE').AsString,tpv,not OkG, Err) then PGIBox(Err.Libelle,'') ;
{$ENDIF}      //NA fin 13/03/02
     Q.Next ;
     End ;
   if OkG then Finimove ;
   Ferme(Q) ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
function ConfirmeCreationPave ( Msg : String ; ForceUneCaisse : Boolean ) : Boolean ;
Begin
if trim(Msg)<>'' then Msg:=Msg+#13+' ' ;
Msg:=Msg+MsgErrDefautS1(756) ; //XMG 07/01/04
{$IFDEF PGIS5} //NA debut 13/03/02
if ForceUneCaisse then Msg:=Msg+'de la caisse ?' else Msg:=Msg+'des caisses ?' ;
{$ELSE}       //NA fin 13/03/02
if (VS1^.TypeProduit<>S1Pro)  or (ForceUneCaisse) then Msg:=Msg+'de la caisse?'
                                                  else Msg:=Msg+'des caisses?' ;
{$ENDIF}
Result:=(PGIAsk(Msg,'')=mrYes) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function CE_Charge ( Caisse : String ; pCE_S : TOB ) : Boolean ;
var Q : TQuery ;
Begin
Result:=FALSE ;
if (Assigned(pCE_S)) and (uppercase(trim(pCE_S.NomTable))='CLAVIERECRAN_S') then
  Begin
  q:=opensql('select * from CLAVIERECRAN where CE_CAISSE="'+Caisse+'" order by CE_PAGE, CE_NUMERO', true);
  result:=pCE_S.LoadDetailDB('CLAVIERECRAN','','',q, FALSE) ;
  Ferme(q) ;
  End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Function CE_Existe (CE_CAISSE : String ) : Boolean ;
Var Q  : TQuery ;
    St : String ;
Begin
Result:=FALSE ;
CE_CAISSE:=Uppercase(Trim(CE_CAISSE)) ;
if trim(CE_CAISSE)<>'' then
   Begin
   St:='select count(CE_CAISSE) from CLAVIERECRAN where CE_CAISSE="'+CE_CAISSE+'"' ;
   Q:=OpenSQL(St,TRUE) ;
   Result:=(Q.Fields[0].AsInteger>0) ;
   Ferme(Q) ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function CE_Controle ( pCE_S : TOB ; Var Err : TPGIErrS1 ) : Boolean ; //XMG 07/01/04
var i         : Integer ;
    CE_CAISSE : String ;
    CE_       : TOB ;
Begin
Err.Code:=0 ; Err.Libelle:='' ;
CE_CAISSE:='' ;
if pCE_S.Detail.Count>0 then CE_CAISSE:=VString(pCE_S.Detail[0].GetValue('CE_CAISSE')) ;
if Trim(CE_CAISSE)='' then
   Begin
   Err.Code:=3 ;
   Err.Libelle:='Le code de la caisse n''est pas renseigné.' ; //XMG 07/01/04
   End ;
if not presence('PCAISSE','PC_CAISSE',CE_CAISSE) then
   Begin
   Err.Code:=3 ;
   Err.Libelle:='La caisse n''existe pas.' ;
   End ;
if Err.code=0 then
   Begin
   for i:=0 to pCE_S.detail.count-1 do
     Begin
     CE_:=pCE_S.Detail[i] ;
     CE_.PutValue('CE_CAISSE',CE_CAISSE) ;
     End ;
   End ;
Result:=(Err.COde=0) ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Function CE_Efface (CE_CAISSE : String ) : Boolean ;
var St : String ;
Begin
Result:=FALSE ;
CE_CAISSE:=Uppercase(Trim(CE_CAISSE)) ;
if trim(CE_CAISSE)<>'' then
   Begin
   St:='delete from CLAVIERECRAN where CE_CAISSE="'+CE_CAISSE+'"' ;
   Result:=(ExecuteSQL(St)>=0) ;
   End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
function  CE_Sauve  (pCE_S : Tob ; Var Err : TPGIErrS1 ) : boolean ; //XMG 07/01/04
Begin
Err.Code:=0 ; Err.Libelle:='' ;
if assigned(pCE_S) and (CE_Controle(pCE_S,err)) then
   if not pCE_S.InsertOrUpdateDB then Err.Code:=1 ;
if (Err.COde>0) and (Trim(Err.Libelle)='') then Err.Libelle:=MsgErrDefautS1(Err.COde) ; //XMG 07/01/04
Result:=(Err.Code=0) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
//XMG fin 02/10/01
end.

