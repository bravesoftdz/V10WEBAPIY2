unit LP_Base;
interface

uses Classes, Windows, Forms, ExtCtrls, Controls, StdCtrls, Messages, Graphics, Hgauge, HCtrls ;

(*
Cete unité contient les controls créés expresement pour l'impresion en mode texte des tickets

La principale fonction de ces controls est que le positionement et fait toujours en carrant à la ligne et la colonne.

Quelques remarques à faire:
   * Les tailles prises par defaut sont 6 lignes par pouce (6 LPI) et 10 caractères par pouce ( 10 CPI)
   * Les propiétés de la police choissi: taille (Size), sousligne (Underlined), italique (Italic), gras (Bold) sont depandentes
     de l'imprimante choissi, du fait que ne toutes les imprimantes ont de possibilitées pareilles
*)
Var LP_GetParamSoc : Function ( Nom : String ; Fromdisk : boolean = FALSE) : Variant = nil;
const
      LPMessageCap  = 'Impression en mode texte' ;  //Constante pour la caption des messages qui sont affichés
      LPNomFont     = 'Courier new' ;              //Nom de la police utilissée
      LPTailleDef   = 18 ;                         //Taille par defaut (la taille est exprime pour à afficher sur l'écran)
      LPTaille05    = LPTailleDef+4 ;              //Taille la plus grande (sur quelques imprimantes correspond à 5CPI)
      LPTaille10    = LPTailleDef ;                //Taille par defaut
      LPTaille12    = LPTailleDef-4 ;              //Taille plus petite (sur quelques imprimantes correspond à 12 CPI)
      LPTaille17    = LPTailleDef-6 ;              //Taille la plus petite (sur quelques imprimantes correspond à 17 CPI)

Type  UnSequenceLP       = record                  //Estructure pour à stocker et gérer l'utilisation des sequences de control de l'imprimante (Sequences échap)
                         Sqce     : Array[boolean] of String ;
                         Groupe,
                         Valeur   : Integer ;
                         InUse    : Boolean ;
                         Extra    : String ;
                         End ;
      //Enum. de toutes les sequences qui sont gérées (peut-être qu'elles ne soient pas implementées sur toutes les imprimantes)
      TLP_SqcesPossibles    =
       {Sequences de Control}    (spCR,spLF,spInitialise,spSautPage,spStatus,spSensorPaperOut,spReStart,spTotalCut,spPartialCut,spActivate,spDesactivate,spActivateSlip,spDesactivateSlip,spActivateValidation,spDesactivateValidation,spDebutCheque,spFinCheque,
       {Sequences "d'affichage"}  sp17Cpi,sp12Cpi,sp10Cpi,sp05Cpi,spUnderline,spBold,spItalic,spBitmap,spPrintBmp,spActivateInterligne,spDesactivateInterligne,spCodeBarre) ;
      //Set des Sequences possibles
      TLP_setSqcesPossibles = Set of TLP_SqcesPossibles ;
      //Definition type du tableau qui stockera les sequences de control
      TabSequenceLP         = Array[TLP_SqcesPossibles] of UnSequenceLP ;
      //Codes des erreurs ou avis qui seront gérés
      TImprimErreurs        = (ieOk,iePowerOff,ieCover,ieErrorPaper,ieWarningPaper,ieOffLine,ieError,ieWarning,ieRecovery,ieUnknown) ;


Const //Constantes diverses pour la gestion des sequences de controls
      MinLP_SqceUser     = sp17Cpi ;
      MaxLP_SqceUser     = High(TLP_SqcesPossibles) ;
      MinLP_SqceCtrl     = Low(TLP_SqcesPossibles) ;
      MaxLP_SqceCtrl     = pred(MinLP_SqceUser)  ;
      LPStCBarre          = 'CodeBarre' ;
      //Nombre maximum d'erreurs avant d'echue
      NbErrorsMessage    = 100 ;
      //Nombre de avis permis
      NbWarnings         = 5 ;
      //Nombre d'essaies avant de passe à ieError
      NbErreursEnsemble  = 3 ;
      //Nombre de recuperations d'erreus
      NbRecoverys        = 10 ;
(*
Object TLParam.- Paramètres generiques de l'état, correspond à la table MODELES.
                 Malgré avoir tous les champs de cette table, on ne les utilisse pas tous
*)
type  TLPParam = Class (TComponent)
        private
         FTypeEtat,
         FNatEtat,
         FCodeEtat,
         FLangue,
         FLibelle,
         FSuivant,
         FUser,
         FDiagText,
         FSQL       : String ;
         FPersonnel,
         FDefaut,
         FModele,
         FProtect,
         FProtectSQL,
         FForceDT,
         FMenu,
         FExporter,
         FNouvelle : Boolean ;
        protected
         Procedure AssignTo ( Dest : TPersistent ) ; override ;
        Public
         Constructor Create ( Aowner : TComponent) ; Override ;

         Function isSameModele (AParamLP : TLPParam) : Boolean ;

         property Nouvelle   : Boolean read FNouvelle write FNouvelle ;
        Published
         property TypeEtat   : String  read FTypeEtat   write FTypeEtat   stored TRUE ;
         property NatEtat    : String  read FNatEtat    write FNatEtat    stored TRUE ;
         property CodeEtat   : String  read FCodeEtat   write FCodeEtat   stored TRUE ;
         property Langue     : String  read FLangue     write FLangue     stored TRUE ;
         property Libelle    : String  read FLibelle    write FLibelle    stored TRUE ;
         property Suivant    : String  read FSuivant    write FSuivant    stored TRUE ;
         property User       : String  read FUser       write FUser       stored TRUE ;
         property DiagText   : String  read FDiagText   write FDiagText   stored TRUE ;
         property SQL        : String  read FSQL        write FSQL        stored TRUE ;
         property Personnel  : Boolean read FPersonnel  write FPersonnel  stored TRUE ;
         property Defaut     : Boolean read FDefaut     write FDefaut     stored TRUE ;
         property Modele     : Boolean read FModele     write FModele     stored TRUE ;
         property Protect    : Boolean read FProtect    write FProtect    stored TRUE ;
         property ProtectSQL : Boolean read FProtectSQL write FPRotectSQL stored TRUE ;
         property ForceDT    : Boolean read FForceDT    write FForceDT    stored TRUE ;
         property Menu       : Boolean read FMenu       write FMenu       stored TRUE ;
         property Exporter   : Boolean read FExporter   write FExporter   stored TRUE ;
         End ;

  TLpUnits = (luNone,luInchs, luCM, luChar, luPixel ) ;  //Unités de messure possibles
(*
Object TLPixel.- Cet object est utilissé de maniere interne par tous les objects definis ci-dessous, pour à faire
                 la conversion parmi les differentes unités de messure
*)
  TLPPixel = Class (TComponent)
    Private
     FPixel     : integer ;
     Px         : TPoint ;
     Procedure RecalculPx ;
     Function  GetFormOwner : TForm ;
    protected
     procedure SetPixel (value : Integer)                                                        ;
     procedure SetChar  (Vertical : Boolean ; Taille : Integer ; value : Integer )               ;
     procedure SetInch  (value : Double)                                                         ;
     procedure SetCm    (value : Double)                                                         ;
     Function  GetChar  (Vertical : Boolean ; Taille : Integer = LPTailleDef) : Integer          ;
     Function  GetInch                                                        : Double           ;
     Function  GetCm                                                          : Double           ;
     Function  IsLastPixelofChar(Vertical : Boolean )                         : Boolean          ; 
     Function  LastPixelofChar(Vertical : Boolean ; NPixel : Integer = -1)    : Integer          ; 
    public
     Constructor Create (AOwner : TComponent ) ; override ;
     Function    ConvertPixel(Pixel : integer = -1 ;Vertical: Boolean = FALSE ; Taille : Integer = LPTailleDef ) : integer ;
     procedure   assign( Source : TPersistent ) ; override ;
     Function    TailleChar(Vertical : Boolean ; Taille : Integer = LPTailleDef ) : Integer ;
     property    AsPixel                                       : Integer    read FPixel    write SetPixel  ;
     property    AsChar[Vertical : Boolean ; Taille : Integer] : Integer    read getchar   write setchar ;
     property    AsInch                                        : Double     read getInch   write SetInch ;
     property    AsCm                                          : Double     read getCm     write SetCm ;
  End ;
(*
Object TLPPoint.- Cet object est utilissé de maniere interne par tous les objects definis ci-dessous, pour à stocker des
                  informations relatives à la position
*)
  TLPPoint = Class (TComponent)
    Private
     FX,FY  : TLPPixel ;
    protected
    public
     constructor create ( Aowner : TComponent) ; Override ;
     Destructor  Destroy ; Override ;
     Procedure   Point (AX,AY : Integer ) ;
     procedure   assign( Source : TPersistent ) ; override ;
    published
     Property X : TLPPixel read FX write FX ;
     Property Y : TLPPixel read FY write FY ;
  end ;
(*
Object TLPRect.- Cet object est utilissé de maniere interne par tous les objects definis ci-dessous, pour à stocker des
                 informations relatives à rectangles
*)

  TLPRect = Class (TComponent)
    Private
     FLeftTop,
     FRightBottom : TLPPoint ;
    Protected
     Procedure SetPixel ( Index : Integer ; Value : TLPPixel) ;
     procedure SetRect ( Value : TRect ) ;
     Function  GetPixel ( Index : Integer) : TLPPixel ;
     Function  GetRect : TRect ;
    public
     Constructor Create ( AOwner : TComponent) ; override ;
     Destructor Destroy ; Override ;
     property   PixelRect   : TRect            Read GetRect        write SetRect ;
     property   LeftTop     : TLPPoint         read FLeftTop       write FLeftTop ;
     property   RightBottom : TLPPoint         read FRightBottom   write FRightBottom ;
    Published
     Property   Left        : TLPPixel index 0 read GetPixel       write SetPixel ;
     Property   Top         : TLPPixel index 1 read GetPixel       write SetPixel ;
     Property   Right       : TLPPixel index 2 read GetPixel       write SetPixel ;
     Property   Bottom      : TLPPixel index 3 read GetPixel       write SetPixel ;
  end;

const margeGTitre  = 2 ;
      MargeDTitre  = 15 ;
      margeGFleche = 2 ;
      MargeBFleche = 3 ;
      LongFleche   = 10 ;
      HautFleche   = 3 ;
      HautPFleche  = 3 ;
      DemiBaseF    = 3 ;  //Base Fleche=2*DemibaseFleche
      TailleOnglet= 11 ; // Taille de l'onglet des bandes
(*
Object TLPBande.- Cet object permet de definir les differentes types de bande supportées et il est le contenant des
                  champs (Objects du type TLPChamp et TLPImage) qui sont imprimés
*)
const NbrMaxSubBnds = 12 ; //Nombre maximum de subbandes possible.
Type
  TLPTypeBandes = (lbEntete,lbSubentete,lbDetail,lbSubdetail,lbPied,lbSubPied,lbEnteteRupt,lbPiedRupt) ; //Type de bande supportées
  TLPBande = Class (TPanel)
    private
     FTitre      : String ;
     PIxel       : TLPPixel ;
     FTypeBande  : TLPTypeBandes ;
     Tw,
     FNbrDetail  : Integer ;
     FModified,
     FBandeVisible,
     FSbEnsemble,
     OkDrag      : Boolean ;
     FCondition,
     FSQL        : String ;
     FInRealigne : Boolean ;
     FTesteSQL   : Boolean ;
     isAssign    : Boolean ;
     FVersionControl : Double ;
     Function  GetVersion  : Double ;
    protected
     Procedure AssignTo    ( Dest : TPersistent )     ; override ;
     Procedure SetTitre ( Value : String ) ;
     Procedure SetTypeBande ( Value : TLPTypeBandes ) ;
     Function  GetUnits : TLPUnits ;
     Function  GetMarge (Index : Integer) : Integer ;
     Function  GetLimite (Index : Integer ) : Integer ;
     procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
     procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
     Function HauterMinimum : Integer ;
     Procedure RefaireBandes (Combien : Integer ) ;
     Function GetModified : Boolean ;
     procedure SetModified (Value : Boolean ) ;
     Function  GetLignes   : Integer ;
     Function  GetColonnes : Integer ;
     Procedure SetNbrDetail ( Value : Integer ) ;
     Procedure SetBandeVisible ( Value : Boolean ) ;
     Procedure SetSQL ( Value : String ) ;
     Procedure SetCondition( Value : String ) ;
     function ChangeSbEnsemble : Boolean ;
     function GetDesigning : Boolean ;
     procedure Loaded ; override ; 
    public
     Constructor Create ( AOwner : TComponent) ; Override ;
     Destructor Destroy ; Override ;
     procedure Paint ; override ;
     Procedure SetBounds(ALeft, ATop, AWidth, AHeight: Integer); override ;
     Procedure RealigneLIgne( ALigne : integer ) ;
     Function TousEvalues : Boolean ;
     Procedure InitialiseChamps (IsEof : Boolean ) ;
     property MargeLeft     : Integer index 0 read GetMarge ;
     property MargeTop      : Integer index 1 read GetMarge ;
     property MargeRight    : Integer index 2 read GetMarge ;
     property MargeBottom   : Integer index 3 read GetMarge ;
     Property LimitLeft     : Integer index 0 Read GetLimite ;
     Property LimitTop      : Integer index 1 Read GetLimite ;
     Property LimitRight    : Integer index 2 Read GetLimite ;
     Property LimitBottom   : Integer index 3 Read GetLimite ;
     Property Units         : TLPUnits        read GetUnits ;
     Property Modified      : Boolean         Read GetModified  write SetModified ;
     property Lignes        : Integer         read GetLignes ;
     Property Colonnes      : Integer         Read GetColonnes ;
     Property InRealigne    : Boolean         Read FInRealigne ;
     Property Designing     : Boolean         Read GetDesigning ;
    Published
     Property titre          : String          read FTitre        write setTitre        stored TRUE ;
     Property TypeBande      : TLPTypeBandes   read FTypeBande    write SetTypeBande    stored TRUE ;
     Property NbrDetail      : Integer         Read FNbrDetail    Write SetNbrDetail    stored TRUE ;
     Property BandeVisible   : Boolean         Read FBandeVisible Write SetBandeVisible stored TRUE ;
     Property Condition      : String          Read FCondition    Write SetCondition    stored TRUE ;
     Property SQL            : String          Read FSQL          Write SetSQL          stored TRUE ;
     Property SbDtlEnsemble  : Boolean         Read FSbEnsemble   Write FSbEnsemble     stored ChangeSbEnsemble ; //Default FALSE ;
     property Color                                                                     stored TRUE ;
     property Top                                                                       stored TRUE ;
     property Left                                                                      stored TRUE ;
     Property Width                                                                     stored TRUE ;
     Property Height                                                                    stored TRUE ;
     Property Name                                                                      stored TRUE ;
     Property TesterSQL      : Boolean         Read FtesteSQL     Write FTesteSQL       default TRUE ;
     Property VersionControl : Double          read GetVersion    Write FVersionCOntrol Stored TRUE ;
    End ;
(*
Object TLPBase.- C'est la "page" qui contient les bandes, ces propiétés principales sont la taille (longeur et largeur),
                 Les marges (Haut, bas, droite et gauche), etc...
*)
Const LP_NbMaxRuptures = 6 ;

Type
  TLPBase = class(TCustomControl)
    private
     FRouleau  : Boolean ;
     FPixel    : TLPPixel ;
     FUnits    : TLPUnits ;
     FMarges   : TRect ;
     FEnPouces : Boolean ;
     FModified : Boolean ;
     FRupt     : String ;
     FChRupt   : String ;
     FInRealigne : Boolean ;
     FDesigning  : Boolean ;
     FVersionControl : Double ;
     Procedure VideFonds ;
     Function  GetVersion  : Double ;
    protected
     Procedure Paint                                ; Override ;
     Procedure AssignTo      (Dest : TPersistent )  ; override ;
     procedure SetUnits      (Value : TLpUnits) ;
     procedure SetEnPouces   (Value : Boolean ) ;
     Procedure SetMarge      (Index,Value : Integer ) ;
     Function  GetMarge      (Index : Integer)        : Integer ;
     Function  GetLimite     (Index : Integer )       : Integer ;
     Function  LimitRect                              : TRect ;
     Function  GetModified                            : Boolean ;
     procedure SetModified   (Value : Boolean ) ;
     Function  GetTailleChar (Index : Integer)        : Integer ;
     Function  StoreChRupt                             : Boolean ;
     Procedure SetChRupt     (AChRupt : String ) ;
     Function  StoreRupt                               : Boolean ;
     procedure Loaded ; override ;
    public
     RecopieContenu : Boolean ;
     Constructor Create (Aowner : TComponent ) ; Override ;
     destructor  Destroy ; override ;
     Procedure   SetBounds(ALeft, ATop, AWidth, AHeight: Integer); override ;
     Function    ChoixBande( QueBande : TLPTypeBandes ; Numdetail : Integer = -1 ) : TLPBande ;
     function    NumMaxSubdetail (TypeSub : TLPTypeBandes ) : Integer ;
     Function    Tailleensemblebande (TypeSub : TLPTypeBandes) : Integer ;
     procedure   ScaleControls(M, D: Integer);
     Procedure   ReAligneLigne ( ALIgne : Integer ) ;
     procedure   RecalculeHeight ;
     property Pixel         : TLPPixel Read FPixel ;
     Property Modified      : Boolean         Read GetModified  write SetModified ;
     Property LimitLeft     : Integer index 0 Read GetLimite ;
     Property LimitTop      : Integer index 1 Read GetLimite ;
     Property LimitRight    : Integer index 2 Read GetLimite ;
     Property LimitBottom   : Integer index 3 Read GetLimite ;
     property tailleligne   : integer index 1 read GetTailleChar ;
     Property TailleChar    : Integer Index 0 read getTailleChar ;
     Property InRealigne    : Boolean         Read FInRealigne ;
    published
     Property Units         : TLpUnits         Read FUnits     write SetUnits        stored TRUE ;
     Property EnPouces      : Boolean          Read FEnPouces  write SetEnPouces     stored TRUE ;
     Property MargeLeft     : Integer  index 0 Read GetMarge   Write SetMarge        stored TRUE ;
     Property MargeTop      : Integer  index 1 Read GetMarge   Write SetMarge        stored TRUE ;
     Property MargeRight    : Integer  index 2 Read GetMarge   Write SetMarge        stored TRUE ;
     Property MargeBottom   : Integer  index 3 Read GetMarge   Write SetMarge        stored TRUE ;
     Property Rouleau       : Boolean          Read FRouleau   Write FRouleau        Default TRUE ;
     property Color                                                                  stored TRUE ;
     property Top                                                                    stored TRUE ;
     property Left                                                                   stored TRUE ;
     Property Width                                                                  stored TRUE ;
     Property Height                                                                 stored TRUE ;
     Property Name                                                                   stored TRUE ;
     Property Enabled                                                                            ;
     property onResize                                                                           ;
     Property ChRupture      : String          Read FChRupt    Write SetChRupt       stored StoreChRupt ;
     Property Ruptures       : String          Read FRupt      Write FRupt           stored StoreRupt ;
     Property Designing      : Boolean         Read FDesigning write FDesigning      default FALSE ;
     Property VersionControl : Double          read GetVersion Write FVersionCOntrol Stored TRUE ;
  end;

  //Valeurs d'utilisation interne dans TLPChamp..
  TLPTailleChar   = (lt17cpi,lt12cpi,lt10cpi,lt05cpi) ;   //... pour la taille de la police
  TLPQuiChange    = (qcaucun,qcWidth,qcHeight,qcLongeur,qcLargeur,qcTaille,qcLigne,qcColonne,qcRealigne) ; //... pendant le resize du champ
  TLPSiValeurZero = (szImprimer,szSupprimeligne,szLigneenBlanc) ; //... pour l'action à prendre si la valeurs est zero
(*
Object TLPChamp.- C'est l'object qui supporte les informations à imprimer, ces propiétés permetent paramètre: son largeur,
                  les différentes propiétés de la police, la source des données, la visibilité pendant l'impression,
                  son positionament dedans la bande, etc...
*)
  TLPChamp = Class (TEdit)
     private
      Pixel           : TLPPixel ;
      FPrinterVisible : Boolean ;
      FAlignTexte     : TAlignment ;
      FLibelle        : String ;
      FMask           : String ;
      FNom            : String ;
      FValue          : String ;
      FTaille         : TLpTailleChar ;
      FNbrChars       : Integer ;
      FNbrLignes      : Integer ;
      FModified       : Boolean ;
      FTextBloque     : Boolean ;
      QuiChange       : TLPQuiChange ;
      ChangeZoom      : Boolean ;
      FSiZero         : TLPSiValeurZero ;
      Fligne          : Integer ;
      FColonne        : Integer ;
      FEvalue         : Boolean ;
      FVersionControl : Double ;
      LibCombo        : String ;
      //Obsolete //XMG ne pas supprimer (compatibilité en arrière....
      FTipe           : String ;
      Procedure WMPaint(var Message: TWMPaint); message WM_PAINT ;
      procedure CMCtl3DChanged(var Message: TMessage); message CM_CTL3DCHANGED;
      procedure WMSetFocus(var Message: TWMSetFocus); message WM_SETFOCUS;
      Function JustifText( Texte : String ) : String ;
      procedure SetPrinterVisible   ( Value   : Boolean)         ;
      Procedure SetLibelle          ( Value   : String )         ;
      Procedure SetMask             ( Value   : String )         ;
      Procedure SetNom              ( Value   : String )         ;
      Procedure SetFAlignTexte      ( Value   : TAlignment)      ;
      procedure SetTaille           ( Value   : TLPTailleChar)   ;
      procedure SetLigne            ( Value   : Integer)         ;
      procedure SetColonne          ( Value   : Integer)         ;
      procedure SetNbrChars         ( Value   : Integer)         ;
      procedure SetNbrLignes        ( Value   : Integer)         ;
      Procedure SetValue            ( AValue  : String )         ;
      Function  FormatResult        (AValue   : String )         : String ;
      Function  GetLigne                                         : Integer ;
      Function  GetColonne                                       : Integer ;
      Function  GetInVisibleParZero                              : Boolean ;
      Function  GetMask                                          : String ;
      Function  IsDesigning                                      : Boolean ;
      Function  GetVersion                                       : Double ;
      Procedure GereCombos                                       ; //XMG 16/09/03
     protected
      procedure PaintWindow ( DC : HDC)                         ; override ;
      Procedure AssignTo ( Dest : TPersistent )                 ; override ;
      procedure MouseMove ( Shift : TShiftState; X,Y : Integer) ; Override ;
      procedure ChangeScale ( M, D: Integer)                    ; override ;
      Procedure SetParent ( AParent : TWinCOntrol)              ; Override ;
      procedure Loaded ; Override ;
     public
      Constructor Create ( AOwner : TComponent) ; Override ;
      Destructor Destroy ; Override ;
      Procedure SetBounds(ALeft, ATop, AWidth, AHeight: Integer); override ;
      Procedure Initialise ( IsEOF : Boolean ) ;
      Function  Police : TLP_SetSqcesPossibles ;
      Function  GetNatureMasque  : String ;
      Function  GetMasqueliteral : String ;
      Function  GetMaskFormule   : String ;
      Function  GetFormatedValue : String ;
      Function  GetJustifiedValue : String ; //XMG 25/06/03
      Function  IsCodeBarre      : Boolean ;
      Property Value            : String   read FValue                Write SetValue ;
      Property Modified         : Boolean  read FModified             Write FModified ;
      Property TextBloque       : Boolean  read FTextBloque           Write FTextBloque ;
      Property InVisibleParZero : Boolean  read GetInvisibleParZero ;
      Property Evalue           : Boolean  read FEvalue               Write FEvalue ;
     published
      property PrinterVisible : Boolean          read FPrinterVisible write SetPrinterVisible stored TRUE ;
      Property Libelle        : String           read FLibelle        write SetLibelle        stored TRUE ;
      Property AlignTexte     : TAlignment       read FAlignTexte     write SetFAlignTexte    stored TRUE ;
      Property Mask           : String           read GetMask         write SetMask           stored TRUE ;
      property Nom            : String           read FNom            write SetNom            stored TRUE ;
      property NbrChars       : Integer          read FNbrChars       write SetNbrChars       stored TRUE ;
      property NbrLignes      : Integer          read FNbrLignes      write SetNbrLignes      stored TRUE ;
      property Ligne          : Integer          read GetLigne        write SetLigne          Stored TRUE ;
      property Colonne        : Integer          read GetColonne      write SetColonne        Stored TRUE ;
      property Taille         : TLPTailleChar    read FTaille         write SetTaille         stored TRUE ;
      Property Name                                                                           stored TRUE ;
      Property SiZero         : TLPSiValeurZero  read FSiZero         write FSiZero           Default szImprimer ;
      Property VersionControl : Double           read GetVersion      Write FVersionCOntrol   Stored TRUE ;
      //Obsolete //XMG ne pas supprimer (compatibilité en arrière....
      property Tipe           : String           read FTipe           write FTipe           ;
      End ;

Const TailleBoule = 5 ;  //Constante interne de TLPContourne qui donne la taille de la petite Boule qui permet modifier
                         //la taille des champs à la souris

Type
  //Valeurs utilissés de maiere interne par TLPContourne pour à ...
  TLPDirection = (laNone,laPosition,laN,laNE,laE,laSE,laS,laSW,laW,laNW) ;  //... Connaitre la direction dans laquelle on modifie la taille
  TLPFixe      = (lfLeft,lfTop,lfWidth,lffixedwidth,lfHeight,lffixedHeight) ; //... connaitre quels directions sont interdites
  TLPSetFixe   = Set of TLPFIxe ; // Definition de l'ensemble de direcions interdites
(*
Object TLPContourne.- Cet object permet de modifier, à la souris et au clavier, la taille et la position du champs qui est en train d'être paramètre
*)
  TLPContourne = Class(TCustomControl)
    Private
     FControl       : TControl ;
     ControlFont    : TFont ;
     Pixel          : TLPPixel ;
     FTailleBloquee : TLPSetFixe ;
     IsAttaching,
     OldTabStop,
     FDrag,
     FIsDblClick,
     PossibleDrag,
     FMouseDown     : Boolean ;
     OldTabOrder    : Integer ;
     MoveRect       : TRect ;
     OffDrag        : TPoint ;
     Direction      : TLPDirection ;
     FCtrlDblClick  : TNotifyEvent ;
     procedure WMEraseBkgnd(var Message: TWMEraseBkgnd); message WM_ERASEBKGND;
     procedure WMGetDLGCode(var Message: TMessage); message WM_GETDLGCODE;
     Procedure updateRect (X,Y,Nw,Nh : integer)  ;
     Procedure CalculeNewPosition(x,y : INteger ; Var Nx,Ny,Nw,Nh : Integer ) ;
    Protected
     Procedure Paint ; Override ;
     Function  GetUnits : TLPUnits ;
     procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
     procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
     procedure KeyDown(var key: Word; Shift: TShiftState); override;
     Procedure SetTailleBLoquee ( Value : TLPSetFixe ) ;
     Function  GetIsAttached : Boolean ;
     Function  FindBandeDrag(P : TPoint ) : TControl ;
     Procedure CreateParams(var Params: TCreateParams); override;
    Public
     Constructor Create ( AOwner : TComponent) ; override;
     Destructor Destroy ; Override ;
     Procedure SetBounds(ALeft, ATop, AWidth, AHeight: Integer); override ;
     Procedure AttachControl(AControl : TControl ; TailleFixe : TLPSetFIxe = []) ;
     Procedure DeAttachControl ;
     Procedure DessineMove(Nx,Ny : Integer ) ;
     Property Control       : TControl read FControl ;
     property isMouseDown   : Boolean read FMouseDown ;
     property Units         : TLPUnits read GetUnits ;
     Property TailleBloquee : TLPSetFIxe read FTailleBloquee write SetTailleBloquee ;
     Property IsAttached    : Boolean read GetIsAttached ;
    Published
     Property OnCtrlDblCLick : TNotifyEvent read FCtrlDblClick write FCtrlDblCLick ;
     Property OnMouseWheel ;
    End ;

(*
Object TLPSqcesLigne.- Avec cet object on peut connaître, au moment de l'impresion, quelles sont les sequences de control activées
                       pour toutes les colonnes possibles
*)
Type
  TLPSqcesLIgne = Class
     Private
      FStP     : array of TLP_SetSqcesPossibles ;
      FValeurs : Array of String ;
      FLarg    : Integer ;
     Protected
      Procedure SetFLarg     (AValue : Integer ) ;
      Procedure SetFStp     ( Index : INteger ; AValue : TLP_SetSqcesPossibles ) ;
      Function  GetFStp     ( Index : INteger ) : TLP_SetSqcesPossibles  ;
      Procedure SetFValeurs ( Index : INteger ; AValue : String ) ;
      Function  GetFValeurs ( Index : INteger ) : String ;
     Public
      Constructor Create ;
      Procedure Clear ;
      Procedure Assign ( Source : TObject ) ;
      Property Largeur       : Integer                          read FLarg       write SetFLarg ;
      Property Sqce[Index    : Integer] : TLP_SetSqcesPossibles read GetFStp     write SetFStP ; default ;
      Property Valeurs[Index : Integer] : String                read GetFValeurs write SetFValeurs ;
     End ;
(*
Object TFLPProgress.- C'est une petite fenêtre qui affiche la prograssion d'un procesus
*)
type
  TFLPProgress = Class (TComponent)
    Private
     FForm     : TForm ;
     FProgress : TEnhancedGauge ;
     FLabel    : THLabel ;
     BCancel   : TButton ;
     FCanceled : Boolean ;
     Procedure BCancelClick(Sender : TObject) ;
    protected
     Function  GetProgress : Integer ;
     Procedure SetLibelle ( NewLibelle : String ) ;
     Function  GetLibelle  : String ;
    public
     Constructor Create (AOwner : TComponent) ; Override ;
     Destructor  Destroy ; Override ;
     Procedure   SetMax ( MaxValue : Integer) ;
     procedure   SetValue ( NewValue : Integer ) ;
     Procedure   Show ;
     Procedure   Hide ;
     Procedure   ForceCancel ;
     Property    Libelle : String read GetLibelle write setlibelle ;
     Property    Progress : Integer Read GetProgress ;
     Property    Canceled : Boolean Read FCanceled ;
    End ;

(*
Object TLPImage.- Cet object est très pareil à TLPChamp, mais il ne gére que des images à imprimer
*)
Type
  TLPImage = Class (TImage)
   Private
     Pixel           : TLPPixel ;
     FPrinterVisible : Boolean ;
     FNom            : String ;
     FFichier        : String ;
     FModified       : Boolean ;
     FLargeur        : Integer ;
     FLongeur        : Integer ;
     BMP             : TBitMap ;
     QuiChange       : Set of TLPQuiChange ;
     PictureBW       : Boolean ;
     FLigne          : integer ;
     FColonne        : Integer ;
     FVersionControl : Double  ;
     Procedure SetNom       ( Value : String ) ;
     Procedure SetFichier   ( Value : String ) ;
     procedure SetLargeur   ( Value   : Integer)         ;
     procedure SetLongeur   ( Value   : Integer)         ;
     Function  GetLigne     : Integer                    ;
     Procedure RealigneImage ( LigneIni : integer ) ;
     Function  GetColonne   : Integer                    ;
     procedure SetLigne     ( Value   : Integer)         ;
     procedure SetColonne   ( Value   : Integer)         ;
     Procedure ChangePicture ( Sender : TObject ) ;
     Function  GetVersion : Double ;
   Protected
     Procedure AssignTo     ( Dest : TPersistent )     ; override ;
     Procedure SetParent ( AParent : TWinCOntrol) ; Override ;
     procedure Loaded ; Override ;
   Public
     Constructor Create (AOwner : TComponent) ; override ;
     Destructor  Destroy ; Override ;
     Procedure   SetBounds(ALeft, ATop, AWidth, AHeight: Integer); override ;
     Procedure   ImprimeImage ;
     Function    ImagetoSt ( Ligne : Integer ) : String ;
     Function    Police : TLP_SetSqcesPossibles ;

     Property    Modified       : Boolean  read FModified           Write FModified ;
   Published
     Property    PrinterVisible : Boolean  read FPrinterVisible write FprinterVisible default TRUE ;
     Property    Nom            : String   read FNom            Write SetNom          stored  TRUE ;
     Property    Fichier        : String   read FFichier        Write SetFichier      stored  TRUE ;
     Property    Largeur        : Integer  read FLargeur        Write SetLargeur      Stored  TRUE ;
     Property    Longeur        : Integer  read FLongeur        Write SetLongeur      Stored  TRUE ;
     property    Ligne          : Integer  read GetLigne        write SetLigne        stored  TRUE ;
     property    Colonne        : Integer  read GetColonne      write SetColonne      stored  TRUE ;
     Property    VersionControl : Double   read GetVersion      Write FVersionControl Stored  TRUE ;
   end ;

Function IsSpecialChamp(Ctrl : TControl ) : Boolean ;
Function IsSpecialChampTexte ( St : String ) : Boolean ;

Type TLPSauvgarde = class(TComponent)
     Private
      FForm    : TForm ;
      FParamLP : TLPParam ;
      FBaseLP  : TLPBase ;
     Public
      Constructor Create ( AOwner : TComponent ) ; override ;
      Destructor  Destroy ; Override ;

      property ParamLP : TLPParam read FParamLP write FParamLP ;
      Property BaseLP  : TLPBase  read FBaseLP  write FBaseLP ;
      End ;

Type TLPTypeZone = ( tzChamp, tzImage) ;
Type TLP_ZoneImp = Class(TObject)
     private
      FType       : TLPTypeZone ;
      FLigne,
      FNbrLignes,
      FColonne,
      FLong,
      FTaille    : Longint ;
      FPolice    : TLP_SetSqcesPossibles ;
      FDonnees,
      FExtra     : String ;
     public
      Property TypeZone  : TLPTypeZone           Read FType      Write FType ;
      Property Ligne     : Longint               Read FLigne     Write Fligne ;
      Property NbrLignes : Longint               Read FNbrLignes Write FNbrlignes ;
      Property Colonne   : Longint               Read FColonne   Write FColonne  ;
      Property Long      : Longint               Read FLong      Write FLong ;   //Caractères ocupées en la ligne...
      Property Taille    : Longint               Read FTaille    Write FTaille ; //longeur des données
      Property Police    : TLP_SetSqcesPossibles Read FPolice    Write FPolice ;
      Property Donnees   : String                Read FDonnees   Write FDonnees ;
      Property Extra     : String                Read FExtra     Write FExtra ;
     End ;

Procedure LP_ChargeMask(CB : THValComboBox ; Ch : TLPChamp ) ;

Procedure LibereSauvgardeLP ;
Var SauvgardeLP : TLPSauvgarde = nil ;

implementation

{$R *.RES}

uses Hent1, Sysutils, Math, Ed_Tools, Dialogs, MC_Erreur, MC_Lib ;

Var CbarreIm : TBitmap = nil ;
//////////////////////////////////////////////////////////////////////////////////
Procedure ChargeCBarreImage ;
Begin
if not assigned(CBarreIm) then
   Begin
   CBarreIm:=TBitmap.Create ;
   CBarreIm.LoadFromResourcename(HINSTANCE,'CODEBAR') ;
   cbarreim.FreeImage ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure LibereCBarreIm ;
Begin
if assigned(CBarreIm) then CBarreIm.Free ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//Estructure des masques utilissables
type TMaskPossibles        = record
                              Nat,
                              Value,
                              Libelle : String ;
                              End ;
      //Nombre de masques utilisables
Const NbMaxMask          = 39  ;
      //Nombre de masques utilisables
      LPNatMaskUser       = 'USR' ;
      LPNatCBarre         = 'CBARRE' ;
      LPChampMaskNat      = 1 ;
      LPChampMaskFmt      = 2 ;
      //Tableau constante des masques utilisables
      MaskPossibles      : Array[1..NbMaxMask] of TMaskPossibles=( (Nat: 'COMBO'     ; Value: 'CVL' ; Libelle: 'Valeur'),
                                                                   (Nat: 'COMBO'     ; Value: 'CLB' ; libelle: 'Libellé'),
                                                                   (Nat: 'COMBO'     ; Value: 'CAB' ; libelle: 'Abrégé'),

                                                                   (Nat: 'DATE'      ; Value: 'DCR' ; Libelle: 'dd/mm/yyyy'),
                                                                   (Nat: 'DATE'      ; Value: 'DLN' ; Libelle: 'dd mmm yyyy'),
                                                                   (Nat: 'DATE'      ; Value: 'DHC' ; Libelle: 'dd/mm/yyyy hh:mm.ss'),
                                                                   (Nat: 'DATE'      ; Value: 'DHL' ; Libelle: 'dd mmm yyyy hh:mm:ss'),
                                                                   (Nat: 'DATE'      ; Value: 'HCR' ; Libelle: 'hh:mm:ss'),

                                                                   (Nat: 'NUM'       ; Value: 'ISZ' ; Libelle: '#,##0'),
                                                                   (Nat: 'NUM'       ; Value: 'ISB' ; Libelle: '#,###'),
                                                                   (Nat: 'NUM'       ; Value: 'INZ' ; Libelle: '###0'),
                                                                   (Nat: 'NUM'       ; Value: 'INB' ; Libelle: '####'),
                                                                   (Nat: 'NUM'       ; Value: '2SZ' ; Libelle: '#,##0.00'),
                                                                   (Nat: 'NUM'       ; Value: '2SB' ; Libelle: '#,###.##'),
                                                                   (Nat: 'NUM'       ; Value: '2S*' ; Libelle: '**#,##0.00**'),
                                                                   (Nat: 'NUM'       ; Value: '2NZ' ; Libelle: '###0.00'),
                                                                   (Nat: 'NUM'       ; Value: '2NB' ; Libelle: '####.##'),
                                                                   (Nat: 'NUM'       ; Value: '2N*' ; Libelle: '**###0.00**'),
                                                                   (Nat: 'NUM'       ; Value: 'LET' ; Libelle: 'Montant en lettres'),

                                                                   (Nat: 'NUM'       ; Value: 'EISZ' ; Libelle: '€#,##0'),
                                                                   (Nat: 'NUM'       ; Value: 'EISB' ; Libelle: '€#,###'),
                                                                   (Nat: 'NUM'       ; Value: 'EINZ' ; Libelle: '€###0'),
                                                                   (Nat: 'NUM'       ; Value: 'EINB' ; Libelle: '€####'),
                                                                   (Nat: 'NUM'       ; Value: 'E2SZ' ; Libelle: '€#,##0.00'),
                                                                   (Nat: 'NUM'       ; Value: 'E2SB' ; Libelle: '€#,###.##'),
                                                                   (Nat: 'NUM'       ; Value: 'E2S*' ; Libelle: '€**#,##0.00**'),
                                                                   (Nat: 'NUM'       ; Value: 'E2NZ' ; Libelle: '€###0.00'),
                                                                   (Nat: 'NUM'       ; Value: 'E2NB' ; Libelle: '€####.##'),
                                                                   (Nat: 'NUM'       ; Value: 'E2N*' ; Libelle: '€**###0.00**'),
                                                                   (Nat: 'NUM'       ; Value: 'LEE' ; Libelle: 'Montant en lettres (Euro)'),

                                                                   (Nat: LPNatCBarre ; Value: 'CBR1' ; Libelle: LPStCBarre+'(UPC-A)'),
                                                                   (Nat: LPNatCBarre ; Value: 'CBR2' ; Libelle: LPStCBarre+'(UPC-E)'),
                                                                   (Nat: LPNatCBarre ; Value: 'CBR3' ; Libelle: LPStCBarre+'(EAN13)'),
                                                                   (Nat: LPNatCBarre ; Value: 'CBR4' ; Libelle: LPStCBarre+'(EAN8)'),
                                                                   (Nat: LPNatCBarre ; Value: 'CBR5' ; Libelle: LPStCBarre+'(CODE39)'),
                                                                   (Nat: LPNatCBarre ; Value: 'CBR6' ; Libelle: LPStCBarre+'(ITF)'),
                                                                   (Nat: LPNatCBarre ; Value: 'CBR7' ; Libelle: LPStCBarre+'(CODABAR)'),
                                                                   (Nat: LPNatCBarre ; Value: 'CBR8' ; Libelle: LPStCBarre+'(CODE93)'),
                                                                   (Nat: LPNatCBarre ; Value: 'CBR9' ; Libelle: LPStCBarre+'(CODE128)')

                                                                   ) ;

//////////////////////////////////////////////////////////////////////////////////
Function LP_ConstruitMasque(Masque : String ) : String ;
var i   : Integer ;
Begin
Result:='' ;
if NbCarInString(masque,';')=2 then Result:=Masque else
if masque<>'' then
   For i:=1 to NbMaxMask do
       if (Masque=MaskPossibles[i].Libelle) then
          Begin
          Result:=MaskPOssibles[i].Nat+';'+MaskPossibles[i].Libelle+';' ;
          Break ;
          End ;
if Result='' then
   Result:=LPNatMaskUser+';'+Masque+';' ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function LP_RecupereNaturemasque( LaMasque : String ) : String ;
Begin
Result:=gtfs(LaMasque,';',LPChampMaskNat) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function LP_TrouveNatureUserMask ( Masque : String) : String ;
Begin
result:='' ;
if (pos('*',Masque)+pos('#',Masque)+pos('0',Masque)+pos('€',Masque)>0)                               then Result:='NUM' else
if pos('d',Masque)+pos('m',Masque)+pos('y',Masque)+pos('h',Masque)+pos('m',Masque)+pos('s',Masque)>0 then Result:='DATE' else
if stricomp(LPStCBarre,PChar(Copy(Masque,1,length(LPStCBarre))))=0                                   then Result:=LPNatCBarre else 
   ;
End ;
/////////////////////////////////////////////////////////////////////////////////
Function LP_RecupereFormatmasque( LaMasque : String ) : String ;
var ii,Dec : integer ;
    nat    : String ;
Begin
Result:='' ;
Nat:=gtfs(LaMasque,';',LPChampMaskNat) ;
//if not instring(Nat,['DATE','COMBO','CBARRE']) then
   begin
   Result:=gtfs(LaMasque,';',LPChampMaskFmt) ;
   Dec:=0 ;
   ii:=pos('0.QQ',Result) ;
   if ii>0 then Dec:=V_PGI.OkdecQ else
      Begin
      ii:=pos('0.VV',Result) ;
      if ii>0 then Dec:=V_PGI.OkdecV else
         Begin
         ii:=pos('0.PP',Result) ; if ii>0 then Dec:=V_PGI.OkdecP ;
         End ;
      End ;
   if ii>0 then
      Begin
      Delete(Result,ii,4) ;
      Insert('0.'+stringofchar('0',dec),Result,ii) ;
      End ;
   End ;
End ;
/////////////////////////////////////////////////////////////////////////////////
Procedure LP_ChargeMask(CB : THValComboBox ; Ch : TLPChamp ) ;
Var St : String ;
    i  : Integer ;
Begin
CB.Items.BeginUpdate ;
CB.Values.BeginUpdate ;
CB.Items.Clear ;
CB.Values.Clear ;
For i:=1 to NbMaxMask do
       Begin
       CB.Items.add(TraduireMemoire(MaskPossibles[i].Libelle)) ;
       CB.Values.add(MaskPossibles[i].Nat) ;
       End ;
St:=Ch.GetMasqueliteral; ;
if (LP_RecupereNatureMasque(St)=LPNatMaskUser) or
   (CB.Items.Indexof(gtfs(St,';',LPChampMaskFmt))<0) then //on ne peut pas utiliser LP_RecupereFormatMasque à cause des .VV, .QQ et .PP
   Begin
   CB.Items.add(gtfs(St,';',LPChampMaskFmt)) ;
   CB.Values.add(LPNatMaskUser) ;
   end ;
CB.Items.EndUpdate ;
CB.Values.EndUpdate ;
CB.itemindex:=CB.Items.Indexof(gtfs(St,';',LPChampMaskFmt)) ;
End ;
/////////////////////////////////////////////////////////////////////////////////
Function RecupOldMask ( Code : String) : String ; //On ne l'utilise que pour la MAJVER du controle
Var i   : Integer ;
Begin
Result:='' ;
if Code='' then Code:='CVL' ; //Valeur
For i:=1 to NbMaxMask do
    if (Code=MaskPossibles[i].Value) then
       Begin
       Result:=MaskPOssibles[i].Nat+';'+MaskPossibles[i].Libelle+';' ;
       Break ;
       End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////
Function  IsMemeStToken ( StToken1,stToken2 : String ) : Boolean ;
Var St,NonTrouves : String ;
    i             : Integer ;
Begin
NonTrouves:='' ;
While StToken1<>'' do
  Begin
  St:=ReadtokenPipe(StToken1,';') ;
  i:=Pos(';'+St+';',';'+StToken2) ;
  if i>0 then Delete(StToken2,i,length(St)+1) else NonTrouves:=NonTrouves+St+';' ;
  End ;
Result:=(StToken2='') and (NonTrouves='') ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function IsSpecialChampTexte ( St : String ) : Boolean ;
BEgin
if copy(St,1,1)='[' then delete(St,1,1) ;
Result:=(Copy(St,1,4)='TEX_') ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function IsSpecialChamp(Ctrl : TControl ) : Boolean ;
//Les champs qui demarrent par "TEX_" sont des champs speciaux à l'impresion texte, P.Ex.- Coupe le tickect, etc...
Var Champ : String ;
Begin
Result:=FALSE ;
if not (Ctrl is TLPChamp) then exit ;
Champ:=TLPChamp(Ctrl).Libelle ;
Result:=IsSpecialChampTexte(Champ) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function IsValidPixel ( P : TLPPoint ; Units : TLPUnits ;  Vertical : Boolean ; Var N : integer) : Boolean ;
//Teste si le Point passe dans P est multiple de l'unité choissi et renvoie, aussi, la valeur à affiche
//Cette fonction est utilisse seulement pendant le paramètrage de l'état
Var M : Integer ;
    C : Double ;
Begin
Result:=False ;
N:=0 ;
if P.x.aspixel=0 then exit ;
case Units of
   luInchs : Begin
             C:=(P.X.AsInch - Int(P.X.AsInch)) ;
             result:=(C=0) or (C=0.5) ;
             if C=0 then N:=round(P.X.AsInch) ;
             End ;
   luCM    : Begin
             Result:=((P.X.AsCm - Int(P.X.AsCm))>=0) and ((P.X.AsCm - Int(P.X.AsCm))<=0.02) ;
             N:=round(P.X.AsCm) ;
             End ;
   luPixel : Begin
             Result:=(P.x.aspixel mod 32) = 0 ;
             N:=P.X.AsPixel ;
             End ;
   luChar  : Begin
             if vertical then M:=3 else M:=10 ;
             Result:=(P.X.IsLastPixelofChar(Vertical)) and ((P.X.AsChar[Vertical,LPTailleDef]+1) mod M=0) ;
             N:=P.X.AsChar[Vertical,LPTailleDef]+1 ;
             End ;
   End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
function TriLaLigne(Item1, Item2: Pointer): Integer;
         /////////////////////////////////////////////////////////////////////////////////////////
         Function TrouveColonne ( Item : Pointer ) : Integer ;
         Begin
         Result:=0 ;
         if TObject(Item) is TLPChamp then Result:=TLPChamp(Item).Colonne else
         if TObject(Item) is TLPImage then Result:=TLPImage(Item).Colonne ;
         End ;
         /////////////////////////////////////////////////////////////////////////////////////////
Var Val1,Val2 : integer ;
Begin
Val1:=TrouveColonne(Item1) ;
Val2:=TrouveColonne(Item2) ;
Result:=Val1-Val2 ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Procedure CherchechampsLigne ( Prnt : TWincontrol ; ALigne : integer ; pLaLigne : TList ) ;
Var i,Larg : Integer ;
    Ctrl   : TControl ;
Begin
if (Not Assigned(pLaligne)) or (not Assigned(Prnt)) then exit ;
For i:=0 to Prnt.ControlCount-1 do
  Begin
  Ctrl:=Prnt.Controls[i] ;
  if (Ctrl is TLPChamp) then
     Begin
     Larg:=TLPChamp(Ctrl).Ligne+TLPChamp(Ctrl).NbrLignes-1 ;
     if (ALigne>=TLPChamp(Ctrl).Ligne) and (Aligne<=Larg) and (pLaLigne.IndexOf(Ctrl)<0) then pLaLigne.Add(Ctrl) ;
     End else
  if (Ctrl is TLPImage) then
     Begin
     Larg:=TLPImage(Ctrl).Ligne+TLPImage(Ctrl).Largeur-1 ;
     if (ALigne>=TLPImage(Ctrl).Ligne) and (ALIgne<=Larg) and (pLaLigne.IndexOf(Ctrl)<0) then pLaLigne.Add(Ctrl) ;
     End ;
  End ;
pLaLigne.Sort(TriLaLigne) ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Procedure GereRealigneLigne( ALigne : Integer ; Prnt : TWinControl  ) ;
Var LaLigne    : TList ;
    i,P,W,diff : Integer ;
    LPBase     : TWinControl ;
Begin
if (not Assigned(Prnt)) then exit ;
LPBase:=Prnt ;
while (Assigned(LPBase)) and (not (LPBase is TLPBase)) do LPBase:=LPbase.Parent ;
if Not Assigned(LPBase) then exit ;
LaLigne:=TList.Create ;
ChercheChampsLigne(Prnt,ALigne,LaLigne) ;
Diff:=0 ;
For i:=0 to LaLigne.Count-1 do
    Begin
    if TObject(LaLigne[i]) is TLPChamp then
       Begin
       TLPBase(LPBase).Pixel.AsChar[FALSE,LPTailleDef]:=TLPChamp(LaLigne[i]).Colonne-1 ;
       W:=TLPBase(LPBase).Pixel.AsPixel ;
       TLPChamp(LaLigne[i]).Left:=W-diff  ;
       TLPChamp(LalIgne[i]).SendToBack ;
       P:=TLPChamp(LaLigne[i]).Left+TLPChamp(LaLigne[i]).Width ;
       TLPBase(LPBase).Pixel.AsChar[FALSE,LPtailleDef]:=TLPChamp(Laligne[i]).Colonne+TLPChamp(Laligne[i]).NbrChars-1 ;
       Diff:=(TLPBase(LPBase).Pixel.AsPixel-P) ;
       End else
    if TObject(LaLigne[i]) is TLPImage then
       Begin
       if TLPImage(LaLigne[i]).Ligne=ALigne then
          Begin
          TLPBase(LPBase).Pixel.AsChar[FALSE,LPTailleDef]:=TLPImage(LaLigne[i]).Colonne-1 ;
          W:=TLPBase(LPBase).Pixel.AsPixel ;
          TLPImage(LaLigne[i]).Left:=W-diff  ;
          End ;
       TLPImage(LalIgne[i]).SendToBack ;
       P:=TLPImage(LaLigne[i]).Left+TLPImage(LaLigne[i]).Width ;
       TLPBase(LPBase).Pixel.AsChar[FALSE,LPtailleDef]:=TLPImage(Laligne[i]).Colonne+TLPImage(Laligne[i]).Longeur-1 ;
       Diff:=(TLPBase(LPBase).Pixel.AsPixel-P) ;
       End ;
    End ;
LaLigne.Free ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Procedure LanceRealigneLigne ( Prnt : TWinControl ; ALIgne : Integer) ;
Begin
if Assigned(Prnt) then
   begin
   if Prnt is TLPBande then TLPBande(Prnt).RealigneLigne(ALigne) else
   if Prnt is TLPBase  then TLPBase(Prnt).RealigneLigne(ALigne) else
   End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Function ConvertenPixel ( pPixel : TLPPixel ; Val,Taille : integer ; Vertical : Boolean ) : Integer ;
Begin
Result:=0 ;
if not assigned(pPixel) then exit ;
pPixel.AsChar[Vertical,Taille]:=Val-1 ;
Result:=pPixel.AsPixel ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Function ConvertenChar ( pPixel : TLPPixel ; Val, taille : integer ; Vertical : Boolean ) : Integer ;
Begin
Result:=0 ;
if not assigned(pPixel) then exit ;
pPixel.asPixel:=Val ;
Result:=pPixel.AsChar[Vertical,Taille]+1 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPParam.Create ( Aowner : TComponent) ;
//Creation de TLPParam
Begin
Inherited Create ( Aowner) ;
Name:='PP'+formatdatetime('yyyymmddhhmmss',now) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPparam.AssignTo ( Dest : TPersistent ) ;
Begin
If Dest is TLPParam then
   With TLPParam(Dest) do
      Begin
      TypeEtat   :=Self.TypeEtat ;
      NatEtat    :=Self.NatEtat ;
      CodeEtat   :=Self.CodeEtat ;
      Langue     :=Self.Langue ;
      Libelle    :=Self.Libelle ;
      Suivant    :=Self.Suivant ;
      User       :=Self.User ;
      DiagText   :=Self.DiagText ;
      SQL        :=Self.SQL ;
      Personnel  :=Self.Personnel ;
      Defaut     :=Self.Defaut ;
      Modele     :=Self.Modele ;
      Protect    :=Self.Protect ;
      ProtectSQL :=Self.ProtectSQL ;
      ForceDT    :=Self.ForceDT ;
      Menu       :=Self.Menu ;
      Exporter   :=Self.Exporter ;
      End else Inherited AssignTo(Dest) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPParam.isSameModele (AParamLP : TLPParam) : Boolean ;
Begin
Result:=FALSE ;
if Self.TypeEtat <>AparamLP.TypeEtat then exit ;
if Self.NatEtat  <>AparamLP.NatEtat  then exit ;
if Self.CodeEtat <>AparamLP.CodeEtat then exit ;
if Self.Langue   <>AparamLP.Langue   then exit ;
Result:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPPixel.Create (AOwner : TComponent) ;
//Createion de TLPPixel
Begin
Inherited Create(AOwner) ;
FPixel:=0 ;
RecalculPx ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPPixel.assign( Source : TPersistent ) ;
//Recuper les valeurs depuis un autre TLPPixel
Begin
if Source is TLPPixel then AsPixel:=TLPPixel(Source).AsPixel
   else inherited Assign(Source) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPixel.GetFormOwner : TForm ;
//Recuper la form propièter
Var F : TComponent;
Begin
F:=Owner ;
while (F<>nil) and (not (F is TForm))  do F:=F.Owner ;
Result:=TForm(F) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPixel.RecalculPx ;
//recalculr les coef dependantes de la form
Var F : TForm ;
Begin
Px.X:=0 ;
Px.Y:=0 ;
F:=GetFormOwner ;
if F<>nil then
   Begin
   Px.X:=(F.pixelsperinch) ;
   Px.Y:=(F.pixelsperinch+16) ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPixel.SetPixel (value : Integer ) ;
//Stocke la valeur en Pixels
Begin
if (Value=AsPixel) then exit ;
FPixel:=Value ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPixel.SetInch (value : Double ) ;
//Stocke la valeur en pouces
Var P : Integer ;
    F : TForm ;
Begin
F:=GetFormOwner ;
P:=round(Value * F.pixelsperinch) ;
if P<>AsPixel then AsPixel:=P ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPixel.GetInch : Double ;
//lit la valeur en pouces
Var F : TForm ;
Begin
result:=AsPixel ;
F:=GetFormOwner ;
if assigned(F) then Result:=int(AsPixel/F.pixelsperinch*100)/100 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPixel.TailleChar(Vertical : Boolean ; Taille : Integer = LPTailleDef ) : Integer ;
//Calcule la taille en caractères
Var P : TPoint ;
Begin
if (PX.Y=0) or (PX.X=0) then RecalculPX ;
if vertical then P:=Point(7,PX.Y)
   else P:=Point(round((LPTailleDef div 2)*(LPTailleDef/Taille)),PX.X) ;
if (P.X<>0) and (P.Y<>0) then Result:=P.Y div P.X
   else Result:=1 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPixel.SetChar  (Vertical : Boolean ; Taille : Integer ; value : Integer ) ;
//Stocke la taille en caracteres
Var P,C : Integer ;
Begin
C:=TailleChar(Vertical,Taille) ;
P:=Value * C ;
if P<>AsPixel then AsPixel:=P ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPixel.GetChar (Vertical : Boolean ; Taille : INteger = LPTailleDef) : Integer ;
//lit la valeur en caracteres
Var C : Integer ;
Begin
C:=TailleChar(Vertical,Taille) ;
Result:=(AsPixel div C) ; 
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPixel.SetCm  (value : Double ) ;
//Stocke la valeur en centimètres
Var XX : Double ;
    P  : Integer ;
Begin
xx:=Px.X/2.54 ;
P:=Round(Value*xx) ;
if P<>AsPixel then AsPixel:=P ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPixel.GetCm   : Double ;
//Lit la valeur en centimètres
Var XX : double ;
Begin
xx:=Px.X/2.54 ;
Result:=int(AsPixel/xx*100)/100 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPixel.LastPixelofChar(Vertical : Boolean ; NPixel : Integer = -1 ) : Integer ;
//Renvoie le pixel plus bas ou droit d'un caractere
Var Old : Integer ;
Begin
result:=0 ;
Old:=AsPixel ;
if Npixel>-1 then AsPixel:=NPixel ;
if AsPixel>0 then
   Begin
   AsChar[Vertical,LPTailleDef]:=AsChar[Vertical,LPTailleDef]+1 ;
   Result:=AsPixel ;
   End ;
AsPixel:=Old ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPixel.IsLastPixelofChar(Vertical : Boolean) : Boolean ;
//Teste si c'est le dernier pixel d'un char
Var Old,R1,R2 : Integer ;
Begin
Old:=AsPixel ;
R1:=AsChar[Vertical,LPTailleDef] ;
AsPixel:=AsPixel+1 ;
R2:=AsChar[Vertical,LPTailleDef] ;
AsPixel:=Old ;
Result:=(R1<>R2) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPixel.ConvertPixel(Pixel : integer = -1 ;Vertical: Boolean = FALSE ; Taille : Integer = LPTailleDef ) : integer ;
//Arrondi le pixel au premier de la colonne/ligne
Var old : Integer ;
Begin
Old:=AsPixel ;
if Pixel>-1 then Aspixel:=Pixel ;
AsChar[Vertical,Taille]:=AsChar[Vertical,Taille] ;
Result:=AsPixel ;
AsPixel:=Old ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
constructor TLPPoint.create ( Aowner : TComponent) ;
Begin
inherited create(AOwner) ;
FX:=TLPPixel.Create(Self) ;
FY:=TLPPixel.Create(Self) ;
point(0,0) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Destructor TLPPoint.Destroy ;
Begin
FX.Free ;
FY.Free ;
Inherited Destroy ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPPoint.assign( Source : TPersistent ) ;
Begin
if Source is TLPPoint then Begin X.aspixel:=TLPPoint(Source).X.Aspixel ; Y.AsPixel:=TLPPoint(Source).Y.AsPixel ; End
   else inherited Assign(Source) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPPoint.Point ( AX,AY : Integer ) ;
Begin
X.Aspixel:=AX ;
Y.AsPixel:=AY ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPRect.Create ( AOwner : TComponent) ;
Begin
Inherited Create(AOwner) ;
FLeftTop:=TLPPoint.Create(Self) ;
FRightBottom:=TLPPoint.Create(Self) ;
SetRect(rect(10,10,10,10)) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Destructor TLPRect.Destroy ;
Begin
FLeftTop.Free ;
FRightBottom.Free ;
inherited Destroy ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPRect.GetPixel ( Index : Integer) : TLPPixel ;
Begin
Result:=nil ;
Case index of
  0 : Result:=LeftTop.X ;
  1 : Result:=LeftTop.Y ;
  2 : Result:=RightBottom.X ;
  3 : Result:=RightBottom.Y ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPRect.SetPixel ( Index : Integer ; Value : TLPPixel ) ;
Begin
Case index of
  0 : LeftTop.X.assign(Value) ;
  1 : LeftTop.Y.assign(Value) ;
  2 : RightBottom.X.Assign(Value) ;
  3 : RightBottom.Y.Assign(Value) ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPRect.SetRect ( Value : TRect ) ;
Begin
LeftTop.X.AsPixel:=Value.Left ;
LeftTop.Y.AsPixel:=Value.Top ;
RightBottom.X.AsPixel:=Value.Right ;
RightBottom.Y.AsPixel:=Value.Bottom ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPRect.GetRect : TRect ;
Begin
Result.Left:=LeftTop.X.AsPixel ;
Result.Top:=LeftTop.Y.AsPixel ;
Result.Right:=RightBottom.X.AsPixel ;
Result.Bottom:=RightBottom.Y.AsPixel ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPBase.Create ( AOwner : TComponent) ;
Begin
Inherited Create(AOwner) ;
ControlStyle:=ControlStyle+[csAcceptsControls] ;
Height:=400 ;
Width:=450 ;
FMarges:=Rect(0,0,0,0) ;
FEnPouces:=FALSE ;
Funits:=luInchs ;
FPixel:=TLPPixel.Create(Self) ;
FRouleau:=TRUE;
RecopieContenu:=False ;
Designing:=FALSE ;
FVersionControl:=-1 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Destructor TLPBase.Destroy ;
Begin
inherited destroy ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBase.SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
Begin
Modified:=(Modified) or (AWidth<>Width) or (AHeight<>Height) ;
if Assigned(parent) then
   Begin
   if aleft<>Left then ALeft:=maxintValue([8,(parent.ClientWidth-AWidth-8) div 2]) ;
   if atop<>ATop then ATop:=maxintValue([8,(parent.ClientHeight-AHeight-8) div 2]) ;
   //XMG 15/04/04 début
   if Parent is TScrollingWinControl then
      Begin
      TScrollingWincontrol(Parent).HorzScrollBar.Range:=(TScrollingWincontrol(Parent).HorzScrollBar.scrollpos+aleft)*2+AWidth ; // 16+AWidth ;
      TScrollingWincontrol(Parent).HorzScrollBar.Position:=minintvalue([TScrollingWincontrol(Parent).HorzScrollBar.scrollpos,TScrollingWincontrol(Parent).HorzScrollBar.Range]) ;
      TScrollingWincontrol(Parent).VertScrollBar.range:=(TScrollingWincontrol(Parent).VertScrollBar.scrollpos+aTop)*2+AHeight ; // 16+Aheight ;
      TScrollingWincontrol(Parent).VertScrollBar.Position:=minintvalue([TScrollingWincontrol(Parent).VertScrollBar.scrollpos,TScrollingWincontrol(Parent).VertScrollBar.Range]) ;
      End ;
   //XMG 15/04/04 fin
   End ;
inherited setbounds(Aleft,Atop,Awidth,Aheight) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Const LPBaseVersionCourante = 2 ;
procedure TLPBase.Loaded ;
var ver : Double ;
Begin
Ver:=VersionControl ;
if Ver<LPBaseVersionCourante then
   Begin
   //Les adaptations ici
   Fmodified:=V_PGI.SAV ; //Si on est en train de parametrer des états en SAV on forcera l'enregistrement du nouveau format
                          // (si client on n'enregistrea les modif que si on modifie d'ailleurs)
   FVersionControl:=LPBaseVersionCourante ;
   End ;
Inherited Loaded ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Function TLPBase.GetVersion : Double ;
Begin
Result:=FVersionControl ;
if (Not (csLoading in ComponentState)) and (Result<>LPBaseVersionCourante) then
   Result:=LPBaseVersionCourante ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBase.VideFonds ;
Var OBrush : TBrush ;
Begin
OBrush:=TBrush.Create ;
OBrush.Assign(Canvas.Brush) ;
Canvas.Brush.Color:=Color ;
Canvas.Brush.Style:=bsSolid ;
Canvas.Fillrect(Canvas.ClipRect) ;
Canvas.Brush.assign(OBrush) ;
OBrush.Free ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBase.SetUnits ( Value : TLPUnits ) ;
Begin
If Value=Units then exit ;
FUnits:=Value ;
Modified:=TRUE ;
VideFonds ;
Refresh ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBase.SetEnPouces( Value : Boolean) ;
Begin
If Value=EnPouces then exit ;
FEnPouces:=Value ;
Modified:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPBase.ScaleControls(M, D: Integer);
Begin
Inherited ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBase.ReAligneLigne ( ALIgne : Integer ) ;
Begin
if (InRealigne) or (Aligne<0) then exit ;
FInRealigne:=TRUE ;
GereRealigneLigne(ALIgne,Self) ;
FInRealigne:=FALSE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBase.SetMarge (Index,Value : Integer ) ;
Begin
Case index of
  0 : if (FMarges.Left<>Value)   and (Value<Width-FMarges.Right)   then
         Begin
         if Value>0 then FMarges.Left:=Value
            else FMarges.Left:=0 ;
         Modified:=TRUE ;
         End ;
  1 : if (FMarges.Top<>Value)    and (Value<Height-FMarges.Bottom) then
         Begin
         if Value>0 then FMarges.Top:=value
            else FMarges.Top:=0 ;
         Modified:=TRUE ;
         End ;
  2 : if (FMarges.Right<>Value)  and (Width-Value>FMarges.left)    then
         Begin
         if Value>0 then FMarges.Right:=value 
            else FMarges.Right:=0 ;
         Modified:=TRUE ;
         End ;
  3 : if (FMarges.Bottom<>Value) and (Height-Value>FMarges.Top)    then
         Begin
         if Value>0 then FMarges.Bottom:=value 
            else FMarges.Bottom:=0 ;
         Modified:=TRUE ;
         End ;
  End ;
Refresh ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLpBase.GetMarge (Index : Integer) : Integer ;
Begin
Case index of
  0 : Result:=FMarges.Left ;
  1 : Result:=FMarges.Top ;
  2 : Result:=FMarges.Right ;
  else  Result:=FMarges.Bottom ;
  End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
procedure TLPBase.SetModified (Value : Boolean ) ;
Var i : Integer ;
Begin
FModified:=Value ;
if Not Value then
   For i:=0 to ControlCount-1 do
     if Controls[i] is TLPBande then TLPBande(Controls[i]).modified:=FALSE else
        if Controls[i] is TLPChamp then TLPChamp(Controls[i]).modified:=FALSE else
        if Controls[i] is TLPImage then TLPImage(Controls[i]).modified:=FALSE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPBase.GetModified : Boolean ;
Var i : Integer ;
Begin
Result:=FModified ;
For i:=0 to ControlCount-1 do
    if Controls[i] is TLPBande then Result:=(Result) or (TLPBande(Controls[i]).Modified) else
    if Controls[i] is TLPChamp then Result:=(Result) or (TLPChamp(Controls[i]).Modified) else
    if Controls[i] is TLPImage then Result:=(Result) or (TLPImage(Controls[i]).Modified) else  ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPBase.GetTailleChar (Index : Integer) : Integer ;
Begin
Result:=Pixel.TailleChar((index=1),LPTailleDef) ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBase.SetChRupt ( AChRupt : String) ;
Begin
AChRupt:=Uppercase(Trim(AChRupt)) ;
if AChRupt=ChRupture then exit ;
if (not IsMemeStToken(AChRupt,ChRupture)) and (trim(ChRupture)<>'') and (copy(Ruptures,1,1)<>';') then Ruptures:='' ;
FChRupt:=AChRupt ; 
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Function TLPBase.StoreChRupt : Boolean ;
Begin
Result:=Trim(ChRupture)<>'' ;
End ;
/////////////////////////////////////////////////////////////////////////////////
Function TLPBase.StoreRupt : Boolean ;
Begin
Result:=Trim(Ruptures)<>'' ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPBase.RecalculeHeight ;
Var BaseH : integer ;
    j     : TLPTypeBandes ;
Begin
if (Rouleau) then
   Begin
   BaseH:=0 ;
   for j:=low(TLPTypeBandes) to High(TLPTypeBandes) do
       if j in [lbEntete,lbDetail,lbPied,lbEnteteRupt,lbPiedRupt] then BaseH:=BaseH+Tailleensemblebande(j) ;
   Pixel.AsChar[TRUE,LPTailleDef]:=BaseH+10 ;  //5 lignes vierges
   BaseH:=PIxel.AsPixel ;
   if Height<>BaseH then Height:=BaseH ;  
   End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBase.Paint ;
Var i,N   : Integer ;
    P : TLPPoint ;
Begin
inherited paint ;

if Units=luNone then exit ;
P:=TLPPoint.Create(self) ;
Canvas.Font.name:='Arial' ;
Canvas.Font.Size:=6 ;
Canvas.Font.Color:=clGray ;

Canvas.Pen.Width:=1 ;

Canvas.Pen.Style:=psDot ;
Canvas.Pen.Color:=clSilver ;
if Units<>luNone then
   Begin
   For i:=0 to Height do
     Begin
     P.X.AsPixel:=i ;
     if IsValidPixel(P,Units,TRUE,N) then
        Begin
        Canvas.Moveto(0,i) ;
        Canvas.LineTo(width,i) ;
        if N<>0 then
           Canvas.TextOut(2,2+P.X.AsPixel,Inttostr(N)) ;
        End ;
     End ;
   For i:=0 to Width do
     Begin
     P.X.AsPixel:=i ;
     if IsValidPixel(P,units,FALSE,N) then
        Begin
        Canvas.Moveto(i,0) ;
        Canvas.LineTo(i,height) ;
        if N<>0 then Canvas.TextOut(2+P.X.AsPixel,2,Inttostr(N)) ;
        End ;
     End ;
   End ;
P.free ;
Canvas.Pen.Color:=clblack ;
Canvas.Pen.Style:=psSolid ;
Canvas.Brush.Style:=bsClear ;
Canvas.Rectangle(MargeLeft-1,MargeTop-1,Width-margeright+1,Height-Margebottom+1) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPBase.LimitRect : TRect ;
Begin
Result:=Rect(0,0,0,0) ;
//Left
Pixel.aspixel:=MargeLeft ;
Result.Left:=Pixel.AsPixel ;
//Top
Pixel.aspixel:=MargeTop ;
Result.Top:=PIxel.AsPixel ;
//Right
Pixel.AsPixel:=Width-MargeRight ;
Pixel.asChar[FALSE,LPTailleDef]:=MaxIntValue([1,Pixel.AsChar[FALSE,LPTailleDef]]) ;
Result.Right:=Pixel.AsPixel ;
//Bottom
Pixel.AsPixel:=Height-MargeBottom ;
Pixel.asChar[TRUE,LPTailleDef]:=MaxIntValue([1,Pixel.AsChar[TRUE,LPTailleDef]]) ;
Result.Bottom:=Pixel.AsPixel ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPBase.GetLimite (Index : Integer ) : Integer ;
Var R : TRect ;
Begin
Result:=0 ;
R:=LimitRect ;
Case index of
  0 : Result:=R.Left ;
  1 : Result:=R.Top ;
  2 : Result:=R.right ;
  3 : Result:=R.bottom ;
  ENd ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPBase.ChoixBande( QueBande : TLPTypeBandes ; Numdetail : Integer = -1 ) : TLPBande ;
var i : integer ;
Begin
Result:=nil ;
For i:=0 to ControlCount-1 do
    if (Controls[i] is TLPBande) and
       (TLPBande(Controls[i]).TypeBande=QueBande) and
       ((not (Quebande in [lbSubDetail,lbSubEntete,lbSubPied,lbenteteRupt,lbPiedRupt])) or (TLPBande(Controls[i]).NbrDetail=NumDetail)) then
       Begin
       Result:=TLPBande(Controls[i]) ;
       Break ;
       End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
function TLPBase.NumMaxSubdetail (TypeSub : TLPTypeBandes ) : Integer ;
Var i : Integer ;
Begin
Result:=0 ;
For i:=0 to ControlCount-1 do
    if (Controls[i] is TLPBande) and (TLPBande(Controls[i]).TypeBande=TypeSub) then Result:=MaxIntValue([Result,TLPBande(Controls[i]).NbrDetail]) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPBase.Tailleensemblebande (TypeSub : TLPTypeBandes) : Integer ;
Var i,j : Integer ;
    Bnd : TLPBande ;
Begin
result:=0 ;
Bnd:=ChoixBande(TypeSub) ;
if (Bnd<>nil) then
   Begin
   Result:=Bnd.Lignes ;
   if Bnd.SbDtlEnsemble then
      Begin
      TypeSub:=succ(TypeSub) ;
      j:=NumMaxSubdetail(TypeSub) ;
      For i:=1 to j do
        Begin
        Bnd:=ChoixBande(TypeSub,i) ;
        if Bnd<>nil then inc(result,Bnd.lignes) ;
        End ;
      End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBase.AssignTo ( Dest : TPersistent ) ;
var i    : Integer ;
    Ctrl : TControl ;
    Cl   : TControlClass ;
Begin
If Dest is TLPBase then
   Begin
   With TLPBase(Dest) do
      Begin
      Left        :=Self.Left ;
      Top         :=Self.Top ;
      Width       :=Self.Width ;
      Height      :=Self.Height ;
      Units       :=Self.Units ;
      EnPouces    :=Self.EnPouces ;
      Color       :=Self.Color ;
      MargeLeft   :=Self.MargeLeft ;
      MargeTop    :=Self.MargeTop ;
      MargeRight  :=Self.MargeRight ;
      MargeBottom :=Self.MargeBottom ;
      enabled     :=Self.Enabled ;
      onResize    :=Self.onResize ;
      Rouleau     :=Self.Rouleau ;
      Pixel.assign(self.pixel) ;
      End ;
      if TLPBase(Dest).RecopieContenu then
         Begin
         While TLPBase(Dest).ControlCount>0 do TLPBase(Dest).Controls[TLPBase(Dest).ControlCount-1].free ;
         TLPBase(Dest).Ruptures:=Self.Ruptures ;
         For i:=0 to Self.Controlcount-1 do
           Begin
           cl:=TControlClass(Self.controls[i].ClassType) ;
           Ctrl:=cl.Create(TLpBase(Dest).owner) ;
           Ctrl.Parent:=TLPBase(Dest) ;
           Ctrl.Assign(Self.Controls[i]) ;
           End ;
         End ;
   End else Inherited AssignTo(Dest) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPContourne.Create ( AOwner: TComponent) ;
Begin
Inherited Create(AOwner) ;
FControl:=Nil ;
ControlFont:=nil ;
FMouseDown:=FALSE ;
PossibleDrag:=FALSE ;
Direction:=laNone ;
OffDrag:=Point(0,0) ;
MoveRect:=rect(0,0,0,0) ;
OldTabStop:=FALSE ;
Pixel:=TLPPixel.create(Self) ;
IsAttaching:=FALSE ;
FTailleBloquee:=[];
FDrag:=FALSE ;
FIsDblClick:=FALSE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Destructor TLPContourne.Destroy ;
Begin
DeAttachControl ;
Pixel.Free ;
if Assigned(ControlFont) then ControlFont.Free ;
Inherited Destroy ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPContourne.CreateParams(var Params: TCreateParams);
begin
inherited CreateParams(Params);
Params.ExStyle := Params.ExStyle + WS_EX_TRANSPARENT;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TLPContourne.WMEraseBkgnd(var Message: TWMEraseBkgnd);
begin
Message.Result := 1 ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPContourne.WMGetDLGCode(var Message: TMessage);
begin
Message.Result := DLGC_WANTARROWS;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPContourne.SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
Begin
if (Aleft<>0) or (Atop<>0) or (AWidth<>0) or (AHeight<>0) then
   Begin
   if (Not isAttaching) then
      Begin
      Aleft:=Aleft+(TailleBoule-2) ;
      Atop:=Atop+(TailleBoule-2) ;
      AWidth:=AWidth-(TailleBoule) ;
      AHeight:=AHeight-(tailleBoule) ;
      End ;

   if Assigned(Control) then Control.SetBounds(ALeft,ATop,Awidth,Aheight) ;

   Aleft:=Control.left-(TailleBoule-2) ;
   Atop:=Control.top-(TailleBoule-2) ;
   AWidth:=Control.Width+(TailleBoule) ;
   AHeight:=Control.Height+(tailleBoule) ;
   End ;

inherited SetBounds(Aleft,Atop,Awidth,Aheight) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPContourne.Paint ;
var i,j,x,y : Integer ;
Begin
if (not Assigned(Control)) then exit ;

inherited paint ;

with Canvas do
  Begin
  Pen.color:=clblack ;
  Pen.Style:=psSolid ;
  Pen.Mode:=pmcopy ;
  Brush.color:=Pen.color ;
  brush.style:=bsSolid ;
  for i:=1 to 3 do
    Begin
    X:=0 ;
    Case i of
      2 : X:=width div 2-(TailleBoule-3)  ;
      3 : X:=width-TailleBoule ;
      End ;
    for j:=1 to 3 do
       Begin
       y:=0 ;
       case j of
         2 : y:=Height div 2-(Tailleboule-3) ;
         3 : Y:=Height-TailleBoule ;
         End ;
       if (lfWidth  in TailleBLoquee) and (i<>2) then Y:=-1 ;
       if (lfheight in TailleBLoquee) and (j<>2) then Y:=-1 ;
       if (Y>=0) and ((i<>2) or ((i=2) and (j<>2))) then FillRect(rect(X,Y,X+TailleBoule,Y+Tailleboule)) ;
       End ;
    End ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPContourne.GetIsAttached : Boolean ;
Begin
Result:=assigned(Control) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPContourne.AttachControl(AControl : TControl ; TailleFixe : TLPSetFIxe = []) ;
Begin
IsAttaching:=TRUE ;
if IsAttached then DeAttachControl ;
FControl:=AControl ;
Parent:=Control.Parent ;
Hint:=Control.Hint ; ShowHint:=(trim(Hint)<>'') ;
TailleBloquee:=TailleFixe ;
SetBounds(Control.Left,Control.Top,Control.Width,Control.Height) ;
if Control is TWinControl then
   Begin
   OldTabStop:=TWincontrol(Control).TabStop ;
   TWincontrol(Control).TabStop:=FALSE ;
   OldTabOrder:=TWinControl(Control).TabOrder ;
   TabStop:=TRUE ;
   TabOrder:=OldTabOrder ;
   End ;
OnCtrlDblClick:=nil ;
if Assigned(COntrolFont) then ControlFont.Free ;
ControlFont:=TFont.Create ;
ControlFont.Name:=LPNomFont ;
ControlFont.Height:=LPTailleDef ;
if Control is TLPChamp then
   Begin
   OnCtrlDblClick:=TLPChamp(Control).OnDblClick ;
   ControlFont.Assign(TLPChamp(Control).Font) ;
   End else if Control is TLPImage then OnCtrlDblClick:=TLPImage(Control).OnDblClick ;
IsAttaching:=FALSE ;
BringToFront ;
if CanFocus then SetFocus ;
FDrag:=FALSE ;
Application.processMessages ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPContourne.DeAttachControl ;
Begin
TabStop:=FALSE ;
TabOrder:=-1 ;
if Assigned(Control) then
   Begin
   if Control is TWinControl then
      Begin
      TWincontrol(Control).TabStop:=OldTabStop ;
      TWinControl(Control).TabOrder:=OldTabOrder ;
      End ;
   FControl:=Nil ;
   OnCtrlDblClick:=nil ;
   SetBounds(0,0,0,0) ;
   Refresh ;
   End ;
FDrag:=FALSE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPContourne.SetTailleBLoquee ( Value : TLPSetFixe ) ;
Begin
if Value=TailleBloquee then exit ;
FTailleBLoquee:=Value
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPContourne.GetUnits : TLPUnits ;
Begin
Result:=luPixel ;
if (Parent is TLPBande) and (not (TLPBande(Parent).Units in [LuNone,LuPixel])) then Result:=TLPBande(Parent).Units ;
if (Parent is TLPBase)  and (not (TLPBase(Parent).Units  in [LuNone,LuPixel])) then Result:=TLPBase(Parent).Units ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPContourne.updateRect (X,Y,Nw,Nh : integer)  ;
Var T,L,H,W : Integer ;
Begin
if ((lfWidth  in TailleBloquee) and (Direction in [laNW,laW,laSW,laSE,laE,laNE])) or
   ((lfHeight in TailleBloquee) and (Direction in [laNW,laN,laSW,laSE,laS,laNE]))    then Exit ;
T:=Top ; L:=Left ; H:=Height  ; W:=Width ;
Case Direction of
  laN  : if t>0 then
            Begin
            T:=T+Y ;
            H:=H-Nh ;
            End ;
  laNE : Begin
         T:=T+Y ;
         H:=H-Nh ;
         W:=W+Nw ;
         End ;
  laE  : W:=W+Nw ;
  laSE : Begin
         H:=H+Nh ;
         W:=W+Nw ;
         End ;
  laS  : H:=H+Nh ;
  laSW : Begin
         H:=H+Nh ;
         L:=L+X ;
         W:=W-Nw ;
         End ;
  laW  : Begin
         L:=L+X ;
         W:=W-Nw ;
         End ;
  laNW : Begin
         T:=T+Y ;
         H:=H-Nh ;
         L:=L+X ;
         W:=W-Nw ;
         End ;
  else //laPosition
         Begin
         L:=L+X ;
         T:=T+Y ;
         End ;
  end ;
Setbounds(L,T,W,H) ;
Refresh ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPContourne.CalculeNewPosition(x,y : INteger ; Var Nx,Ny,Nw,Nh : Integer ) ;
Var Cx,Cy,Ox,Oy,ow,oh : Integer ;
Begin
Ox:=1 ; Oy:=1 ; Nh:=0 ; Nw:=0 ;
if (Parent is TLPBande) or (Parent is TLPBase) then
   Begin
   Ox:=pixel.taillechar(FALSE) ;
   Oy:=pixel.tailleChar(TRUE) ;
   End ;

Cx:=0 ; Ow:=0 ;
Cy:=0 ; Oh:=0 ;
if Direction=laPosition then
   Begin
   Cx:=OffDrag.X ;
   Cy:=OffDrag.Y ;
   End else
   Begin
   if Direction in [laSW,laS,laSE] then Cy:=Height ;
   if Direction in [laSE,laE,laNE] then Cx:=Width ;
   Ow:=pixel.taillechar(FALSE,ControlFont.Height) ;
   Oh:=pixel.taillechar(TRUE,ControlFont.Height) ;
   End ;

Nx:=((X-Cx) div Ox)*Ox ;
Ny:=((Y-Cy) div Oy)*Oy ;
if direction<>laPosition then
   Begin
   Nw:=((X-Cx) div Ow)*Ow ;
   Nh:=((Y-Cy) div Oh)*Oh ;
   if (lfWidth  in TailleBloquee) then Begin Nx:=0 ; Nw:=0 ; end ;
   if (lfHeight in TailleBloquee) then Begin Ny:=0 ; Nh:=0 ; end ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPContourne.MouseMove(Shift: TShiftState; X, Y: Integer);
Var Y1,Y2,Y3,X1,X2,Nx,Ny,Nw,Nh : Integer ;
    NewCur                     : TCursor ;
Begin
if Not Assigned(Control) then exit ;
if IsMouseDown then
   Begin
   CalculeNewPosition(x,y,Nx,Ny,Nw,Nh) ;
   if ((Abs(Nx)>=1) or (abs(Ny)>=1)) then
      Begin
      FDrag:=TRUE ;
      if Direction=laPosition then DessineMove(Nx,Ny)
         else UpdateRect(Nx,Ny,Nw,Nh) ;
      End
   End else
   Begin
   NewCur:=crdefault ;
   Direction:=laNone ;
   Y1:=Height div 2 - (TailleBoule-3) ;
   Y2:=Height div 2 + (TailleBoule-3) ;
   Y3:=Height -  tailleBoule ;
   X1:=width div 2  - (tailleboule-3) ;
   X2:=Width div 2  + (tailleboule-3) ;
   if (X>=0) and (X<=TailleBoule) then
      Begin
      if (Y>=0)  and (Y<=TailleBoule) then Begin Direction:=laNW ; NewCur:=crSizeNWSE ; End else
      if (Y>=Y1) and (Y<=Y2)          then Begin Direction:=laW  ; NewCur:=crSizeWE ;   End else
      if (Y>=Y3) and (Y<=Height)      then Begin Direction:=laSW ; NewCur:=crSizeNESW ; End ;
      End else
   if (X>=X1) and (X<=X2) then
      Begin
      if ((Y>=0) and (Y<=TailleBoule)) or
         ((Y>=Height-TailleBoule) and (Y<=Height)) then
         Begin
         if ((Y>=0) and (Y<=TailleBoule)) then Direction:=laN else Direction:=laS ;
         NewCur:=crSizeNS ;
         End ;
      End else
   if (X>width-tailleboule) and (X<=width) then
      Begin
      if (Y>=0)  and (Y<=TailleBoule) then Begin Direction:=laNE ; NewCur:=crSizeNESW ; End else
      if (Y>=Y1) and (Y<=Y2)          then Begin Direction:=laE  ; NewCur:=crSizeWE ;   End else
      if (Y>=Y3) and (Y<=Height)      then Begin Direction:=laSE ; NewCur:=crSizeNWSE ; End ;
      End ;
   if ((lfWidth  in TailleBloquee) and (Direction in [laNW,laW,laSW,laSE,laE,laNE])) or
      ((lfHeight in TailleBloquee) and (Direction in [laNW,laN,laSW,laSE,laS,laNE]))    then NewCur:=crDefault ;
   Cursor:=NewCur ;
   End ;
PossibleDrag:=(Cursor<>crDefault) ;
inherited mousemove(Shift,x,y) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPContourne.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Begin
if (not Assigned(Control)) or (IsAttaching) then exit ;
if (Button=mbLeft) then
   Begin
   FIsDblClick:=(ssDouble in Shift) ;
   if not FIsDblClick then
      Begin
      OffDrag:=Point(X,Y) ;
      if not (PossibleDrag) then
         Begin
         Moverect:=BoundsRect ;
         DessineMove(0,0) ;
         Direction:=laPosition ;
         End ;
      FMouseDown:=TRUE ;
      End else if assigned(OnCtrlDblClick) then OnCtrlDblClick(Control) ;
   End ;
inherited MouseDown(Button,Shift,x,y) ;
application.processmessages ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPContourne.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Var Nx,Ny,Nw,Nh : Integer ;
    Reattache   : Boolean ;
    Ctrl        : TControl ;
Begin
if Not Assigned(Control) then exit ;
Reattache:=FALSE ;
if (IsMouseDown) then
   Begin
   if (not FIsDblClick) then
      Begin
      if Direction=laPosition then
         Begin
         CalculeNewPosition(x,y,Nx,Ny,Nw,Nh) ;
         DessineMove(Nx,Ny) ;
         Ctrl:=FindBandeDrag(parent.Clienttoscreen(point(MoveRect.left+TailleBoule,Moverect.Top+TailleBoule))) ;
         if Ctrl=Self then Ctrl:=Parent ;
         if (Ctrl is TLPChamp) or (Ctrl is TLPImage) then Ctrl:=Ctrl.Parent ;
         if (Ctrl<>nil) and (Ctrl<>parent) then
            Begin
            Control.Parent:=TWinControl(Ctrl) ;
            Reattache:=TRUE ;
            Moverect.top:=Moverect.top+parent.top-Control.Parent.top ;
            End ;
         SetBounds(Moverect.Left,moverect.top,Width,Height) ;
         End else reattache:=(Control is TLPImage) ;
      if FDrag then recreateWnd ;
      Refresh ;
      End ;
   FDrag:=FALSE ;
   MoveRect:=rect(0,0,0,0) ;
   OffDrag:=point(0,0) ;
   Direction:=laNone ;
   possibledrag:=FALSE ;
   FMouseDown:=FALSE ;
   FIsDblClick:=FALSE ;
   End ;
inherited MouseDown(Button,Shift,x,y) ;
if Reattache then AttachControl(Control,TailleBloquee) ;
application.processmessages ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPContourne.FindBandeDrag(P : TPoint ) : TControl ;
Begin
Result:=FindDragTarget(P,FALSE) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPContourne.DessineMove(Nx,Ny : Integer ) ;
          ////////////////////////////////////////////////////////////////////////////////////////////////////////
          Procedure DessineRectangle ( P2 : TWinControl ) ;
          Var DC : HDC ;
              P  : TPoint ;
          Begin
          DC:=GetDC(P2.Handle) ;
          p:=MoveRect.TopLeft ;
          p:=P2.ScreenToClient(Parent.clienttoscreen(p)) ;
          drawfocusrect(DC,rect(P.X+(TailleBoule-3),P.Y+(TailleBoule-3),P.X+Width-(TailleBoule-2),P.Y+Height-(TailleBoule-2))) ;
          ReleaseDC(P2.Handle,DC) ;
          End ;
          ////////////////////////////////////////////////////////////////////////////////////////////////////////
Var C1,C2 : TControl ;
    NewR  : TRect ;
Begin
NewR:=boundsrect ;
NewR.Left:=NewR.Left+Nx ;
NewR.Top:=NewR.top+Ny ;

C1:=FindBandeDrag(Parent.Clienttoscreen(point(MoveRect.left+TailleBoule,Moverect.Top+TailleBoule))) ;
C2:=FindBandeDrag(parent.Clienttoscreen(point(NewR.left+TailleBoule,NewR.Top+TailleBoule))) ;

If (C1=Self) or (C1=Control) then C1:=parent ;
if C1 is TLPChamp then C1:=C1.Parent ;
If (C2=Self) or (C2=Control) then C2:=parent ;
if C2 is TLPChamp then C2:=C2.Parent ;

if (C2=nil) or (not (C2 is TLPBande)) then exit ;

DessineRectangle(TWinControl(C1)) ;

MoveRect:=NewR ;

DessineRectangle(TWinControl(C2)) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPContourne.KeyDown(var key: Word; Shift: TShiftState);
Var OldAdd      : TLPDirection ;
    X,Y,Nw,Nh,t : Integer ;
Begin
OldAdd:=Direction ;
Direction:=lanone ;
X:=0 ; Y:=0 ; Nw:=0 ; Nh:=0 ;
Case key of
  VK_UP   :  Begin
             if (Shift=[ssShift]) or (Shift=[ssCtrl]) then
                Begin
                if Shift=[ssShift] then
                   Begin
                   Direction:=laS ;
                   t:=ControlFont.Height ;
                   end else
                   Begin
                   Direction:=laPosition ;
                   t:=LPTailleDef ;
                   End ;
                If (Parent is TLPBande) or (Parent is TLPBase) then
                   Begin
                   Y:=-Pixel.TailleChar(TRUE,t) ;
                   if Direction<>laPosition then Nh:=Y ;
                   End else Y:=-1 ;
                End ;
             End ;
  VK_DOWN :  Begin
             if (Shift=[ssShift]) or (Shift=[ssCtrl]) then
                Begin
                if Shift=[ssShift] then
                   Begin
                   Direction:=laS ;
                   t:=ControlFont.Height ;
                   end else
                   Begin
                   Direction:=laPosition ;
                   t:=LPTailleDef ;
                   End ;
                If (Parent is TLPBande) or (Parent is TLPBase) then
                   BEgin
                   Y:=Pixel.TailleChar(TRUE,t) ;
                   if Direction<>laPosition then Nh:=Y ;
                   end else Y:=1 ;
                End ;
             End ;
  VK_LEFT :  Begin
             if (Shift=[ssShift]) or (Shift=[ssCtrl]) then
                Begin
                if Shift=[ssShift] then
                   Begin
                   Direction:=laE ;
                   t:=ControlFont.Height ;
                   end else
                   Begin
                   Direction:=laPosition ;
                   t:=LPTailleDef ;
                   End ;
                If (Parent is TLPBande) or (Parent is TLPBase) then
                   Begin
                   X:=-Pixel.TailleChar(FALSE,t) ;
                   if Direction<>laPosition then Nw:=X ;
                   End else X:=-1 ;
                End ;
             End ;
  VK_RIGHT : Begin
             if (Shift=[ssShift]) or (Shift=[ssCtrl]) then
                Begin
                if Shift=[ssShift] then
                   Begin
                   Direction:=laE ;
                   t:=ControlFont.Height ;
                   end else
                   Begin
                   Direction:=laPosition ;
                   t:=LPTailleDef ;
                   End ;
                If (Parent is TLPBande) or (Parent is TLPBase) then
                   Begin
                   X:=Pixel.TailleChar(FALSE,t) ;
                   if Direction<>laPosition then Nw:=X ;
                   end else X:=1 ;
                End ;
             End ;
  End ;
if Direction<>lanone then
   Begin
   If Direction=laPosition then
      Begin
      PossibleDrag:=TRUE ;
      MoveRect:=BoundsRect ;
      End ;
   UpdateRect(X,Y,Nw,Nh) ;
   If Direction=laPosition then
      Begin
      PossibleDrag:=FALSE ;
      MoveRect:=rect(0,0,0,0) ;
      End ;
   Direction:=OldAdd ;
   recreateWnd ;
   End ;
inherited KeyDown(Key,Shift) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPBande.Create ( AOwner : TComponent) ;
Begin
Inherited Create(AOwner) ;
controlstyle:=ControlStyle+[csOpaque] ;
Autosize:=FALSE ;
BevelInner:=bvNone ;
BevelOuter:=bvNone ;
Caption:='' ;
Pixel:=TLPPixel.Create(Self) ;
FTitre:='' ;
OkDrag:=FALSE ;
FNbrDetail:=-1 ;
FCondition:='' ;
FSQL:='' ;
TW:=0 ;
FBandeVisible:=TRUE ;
FSbEnsemble:=FALSE ;
FInRealigne:=FALSE ;
FTesteSQL:=TRUE ;
isAssign:=FALSE ;
FVersionControl:=-1 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Destructor TLPBande.Destroy ;
Begin
Inherited Destroy ;
End ;
//////////////////////////////////////////////////////////////////////////////////
procedure TLPBande.SetModified (Value : Boolean ) ;
Var i : Integer ;
Begin
FModified:=Value ;
if Not Value then
   For i:=0 to ControlCount-1 do
     if Controls[i] is TLPChamp then TLPChamp(Controls[i]).modified:=FALSE else
        if Controls[i] is TLPImage then TLPImage(Controls[i]).modified:=FALSE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPBande.GetModified : Boolean ;
Var i : Integer ;
Begin
Result:=FModified ;
For i:=0 to ControlCount-1 do
    if Controls[i] is TLPChamp then Result:=(Result) or (TLPChamp(Controls[i]).Modified) else
    if Controls[i] is TLPImage then Result:=(Result) or (TLPImage(Controls[i]).Modified) else ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
Var ChangeHeight : Boolean ;
Begin
ChangeHeight:=FALSE ;
if (Parent<>nil) then
  Begin
  Aleft:=0 ;
  AWidth:=Parent.Width ;
  End ;
if (Parent is TLPBase) then
  Begin
  Aleft:=TLPBase(Parent).Margeleft ;
  Atop:=maxintvalue([TLPBase(Parent).MargeBottom-height,maxintValue([Atop,TLPBase(Parent).Margetop+1])]) ; Atop:=pixel.convertPixel(Atop,TRUE) ;
  AWidth:=TLPBase(Parent).Width-TLPBase(Parent).margeRight-Aleft  ;
  ChangeHeight:=(Aheight<>Height)and (TLPBase(Parent).Rouleau) ;
  End ;
modified:=(Modified) or ((Atop<>Top) or (ALeft<>Left) or (AWidth<>Width) or (Aheight<>Height)) ;
Inherited SetBounds(Aleft,Atop,AWidth,AHeight) ;
if ChangeHeight then TLPBase(Parent).RecalculeHeight ;
refresh ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.ReAligneLigne ( ALIgne : Integer ) ;
Begin
if (InRealigne) or (Aligne<0) then exit ;
FInRealigne:=TRUE ;
GereRealigneLigne(ALIgne,Self) ;
FInRealigne:=FALSE ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPBande.TousEvalues : Boolean ;
var i : integer ;
Begin
Result:=TRUE ;
For i:=0 to ControlCount-1 do
  if Controls[i] is TLPChamp then
     Begin
     Result:=(Result) and (TLPChamp(Controls[i]).Evalue) ;
     if not result then break ;
     End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.InitialiseChamps (IsEof : Boolean ) ;
Var ii   : Longint ;
Begin
For ii:=0 to ControlCount-1 do
  if Controls[ii] is TLPChamp then TLPChamp(Controls[ii]).Initialise(IsEof) ; ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPBande.GetLignes   : Integer ;
Begin
Pixel.asPixel:=Height ;
Result:=Pixel.AsChar[TRUE,LPTailleDef] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPBande.GetColonnes : Integer ;
Begin
Pixel.asPixel:=Width ;
Result:=Pixel.AsChar[FALSE,LPTailleDef] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.SetTitre( Value : String ) ;
Begin
if Value=Titre then exit ;
FTitre :=Value ;
Modified:=TRUE ;
Refresh ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.SetBandeVisible ( Value : Boolean ) ;
Begin
if Value=BandeVisible then exit ;
FBandeVisible:=Value ;
Modified:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.SetSQL ( Value : String ) ;
Begin
if Value=SQL then exit ;
FSQL:=Value ;
Modified:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
function TLPBande.ChangeSbEnsemble : Boolean ;
Begin
Result:=FALSE ;
if TypeBande in [lbEntete,lbDetail,lbPied] then
   Begin
   SbDtlEnsemble:=(TLPBase(Parent).NumMaxSubdetail(succ(TypeBande))>0) ;
   Result:=SbDtlEnSemble ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Const LPBandeVersionCourante = 2 ;
procedure TLPBande.Loaded ;
var ver : Double ;
Begin
Ver:=VersionControl ;
if Ver<LPBandeVersionCourante then
   Begin
   //Les adaptations ici
   Fmodified:=V_PGI.SAV ; //Si on est en train de parametrer des états en SAV on forcera l'enregistrement du nouveau format
                          // (si client on n'enregistrea les modif que si on modifie d'ailleurs)
   FVersionControl:=LPBandeVersionCourante ;
   End ;
Inherited Loaded ;
End ;
//////////////////////////////////////////////////////////////////////////////////
function TLPBande.GetDesigning : Boolean ;
Begin
Result:=FALSE ;
if (Assigned(Parent)) and (Parent is TLPBase) then Result:=TLPBase(Parent).Designing ; 
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.SetCondition( Value : String ) ;
Begin
if Value=Condition then exit ;
FCondition:=Value ;
Modified:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.SetTypeBande ( Value : TLPTypeBandes ) ;
Begin
if Value=TypeBande then exit ;
if (not isAssign) and (Value in [lbSubDetail,lbSubEntete,lbSubPied]) and (Parent is TLPBase) and (TLPBase(Parent).ChoixBande(pred(Value))=nil) then exit ;
FTypeBande:=Value ;
if (TypeBande in [lbSubDetail,lbSubEntete,lbSubPied,lbEnteteRupt,lbPiedRupt]) and (not isAssign) then NbrDetail:=maxint ;
Modified:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.SetNbrDetail ( Value : Integer ) ;
Var i : Integer ;
Begin
if (not (TypeBande in [lbSubEntete,lbSubdetail,lbSubPied,lbEnteteRupt,lbPiedRupt])) or (NbrDetail=Value) then Exit ;
if (Parent is TLPBase) and (not isAssign) then
   Begin
   Value:=minIntValue([Value,TLPBase(Parent).NumMaxSubdetail(TypeBande)+1]) ;
   For i:=0 to Parent.ControlCount-1 do
     if (Parent.Controls[i] is TLPBande) and (TLPBande(Parent.Controls[i]).TypeBande=TypeBande) and
        (TLPBande(Parent.Controls[i])<>Self) then
        if ((NbrDetail=-1) or (TLPBande(Parent.Controls[i]).NbrDetail<NbrDetail)) and
           (TLPBande(Parent.Controls[i]).NbrDetail>=Value)    Then
           Begin
           TLPBande(Parent.Controls[i]).NbrDetail:=TLPBande(Parent.Controls[i]).NbrDetail+1 ;
           End ;
   End ;
FNbrDetail:=Value ;
Modified:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPBande.Paint ;
Var I,N,ML,MT : Integer ;   
    P         : TLPPoint ;
Begin
inherited paint ;
P:=TLPPoint.Create(self) ;
Canvas.Font.name:='Arial' ;
Canvas.Font.Size:=6 ;
Canvas.Font.Color:=clGray ;

Canvas.Pen.Width:=1 ;

Canvas.Brush.Color:=clWhite ;
Canvas.Brush.Style:=bsSolid ;
Canvas.Fillrect(Canvas.ClipRect) ;

Canvas.Pen.Style:=psDot ;
Canvas.Pen.Color:=clSilver ;
ML:=0 ; MT:=0 ;
If Parent is TLPBase then
   Begin
   ML:=TLPBase(Parent).Margeleft ;
   MT:=TLPBase(Parent).MargeTop ;
   End ;
if Units<>luNone then
   Begin
   For i:=0 to Height do
     Begin
     P.X.AsPixel:=i+top ;
     if IsValidPixel(p,Units,TRUE,N) then
        Begin
        Canvas.Moveto(0,i) ;
        Canvas.LineTo(width,i) ;
        if (ML=0) and (N<>0) then Canvas.TextOut(2,2+i,Inttostr(N)) ;
        End ;
     End ;
   For i:=0 to Width do
     Begin
     P.X.AsPixel:=i+Left ;
     if IsValidPixel(P,Units,FALSE,N) then
        Begin
        Canvas.Moveto(i,0) ;
        Canvas.LineTo(i,height) ;
//        if (MT=Top) and (top=0) then begin
          Canvas.TextOut(2+i,2,Inttostr(N));
          Canvas.TextOut(2+i,height-10,Inttostr(N)); // VL
//        end;
        End ;
     End ;
   End ;
P.free ;
With canvas do
  Begin
  //Initialise
  Pen.color:=clblue ;
  Pen.width:=1 ;
  Pen.style:=psSolid ;
  Pen.mode:=pmCopy ;
  //Ligne en bas
  Moveto(0,Height-1) ;
  lineto(Width,Height-1) ;
  //prepare l'affichage du titre
  Font.Color:=clwhite ;
  Font.Size:=7 ;
  tw:=textWidth(Titre) ;
  //Onglet
  Brush.color:=Pen.color ;
  Brush.style:=bssolid ;
  fillrect(rect(0,Height-Tailleonglet,tw+MargeGTitre+MargeDTitre,height-1)) ;
  //Affiche le titre
  TextOut(MargeGTitre,Height-Tailleonglet,Titre) ;
  //Dessine la flèche
  Pen.Color:=clWhite ;
  Brush.Color:=Pen.color ;
  Brush.Style:=bssolid ;
  Moveto(Tw+MargeGTitre+margeGFleche,Height-MargeBFleche) ;
  Lineto(tw+MargeGTitre+LongFleche,Height-MargeBFleche) ;
  Lineto(tw+MargeGTitre+LongFleche,Height-MargeBFleche-HautFleche) ;
  polygon([point(tw+MargeGTitre+LongFleche-DemiBaseF,Height-MargeBFleche-HautFleche),point(tw+MargeGTitre+LongFleche,Height-MargeBFleche-HautFleche-HautPFleche),point(tw+MargeGTitre+LongFleche+DemiBaseF,Height-MargeBFleche-HautFleche)]) ;
  End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPBande.GetVersion : Double ;
Begin
Result:=FVersionControl ;
if (Not (csLoading in ComponentState)) and (Result<>LPBandeVersionCourante) then
   Result:=LPBandeVersionCourante ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.AssignTo ( Dest : TPersistent ) ;
var Cl   : TControlClass ;
    Ctrl : TControl ;
    i    : Integer ;
Begin
If Dest is TLPBande then
   Begin
   With TLPBande(Dest) do
      Begin
      isAssign      :=TRUE ; 
      Left          :=Self.Left ;
      Top           :=Self.Top ;
      Width         :=Self.Width ;
      Height        :=Self.Height ;
      Titre         :=Self.Titre ;
      TypeBande     :=Self.TypeBande ;
      Color         :=Self.Color ;
      titre         :=Self.Titre ;
      NbrDetail     :=Self.NbrDetail ;
      BandeVisible  :=Self.BandeVisible ;
      Condition     :=Self.Condition ;
      SQL           :=Self.SQL ;
      SbDtlEnsemble :=Self.SbDtlEnsemble ;
      End ;
      if (TLPBande(Dest).Parent is TLPBase) and
         (TLPBase(TLPBande(Dest).parent).RecopieContenu) then
         Begin
         While TLPBande(Dest).ControlCount>0 do TLPBande(Dest).Controls[TLPBande(Dest).ControlCount-1].free ;
         For i:=0 to Self.ControlCOunt-1 do
             Begin
             cl:=TControlClass(Self.controls[i].ClassType) ;
             Ctrl:=cl.Create(TLpBande(Dest).owner) ;
             Ctrl.Parent:=TLPBande(Dest) ;
             Ctrl.Assign(Self.Controls[i]) ;
             End ;
         End ;
      isAssign:=FALSE ;
   End else Inherited AssignTo(Dest) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPBande.GetUnits : TLPUnits ;
Begin
Result:=luNone ;
if (Parent is TLPBase) and (TLPBase(Parent).Units<>LuNone) then
   Result:=TLPBase(Parent).Units ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLpBande.GetMarge (Index : Integer) : Integer ;
Begin
Result:=0 ;
if Parent is TLPBase then
  Begin
  Case index of
    0 : Result:=left ;
    1 : Result:=Top ;
    2 : Result:=Pixel.ConvertPixel(Width) ;
    else  Result:=Pixel.ConvertPixel(Height,TRUE) ; 
    End ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPBande.GetLimite (Index : Integer ) : Integer ;
Var R : TRect ;
Begin
Result:=0 ;
R:=ClientRect ;
Case index of
  0 : Result:=R.Left ;
  1 : Result:=R.Top ;
  2 : Result:=R.right ;
  3 : Result:=R.bottom ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer) ;
var ww : Integer ;
Begin
inherited Mousedown(Button,Shift,X,Y) ;
ww:=tw+MargeGTitre+MargeDTitre ; 
OkDrag:=(((X>=0) and (X<=ww)) and ((Y>(Height-TailleOnglet)) and (Y<=Height)) and (Button=mbLeft)) ;
if OkDrag then screen.Cursor:=crSizeNS ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPBande.HauterMinimum : Integer ;
Var i : Integer ;
Begin
Result:=Pixel.TailleChar(TRUE,LPTailleDef) ;
For i:=0 to ControlCount-1 do
  if (Not (Controls[i] is TLPContourne)) and ((Controls[i].Top+Controls[i].Height)>Result) then Result:=(Controls[i].Top+Controls[i].Height) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.RefaireBandes (Combien : Integer ) ;
Var Bnd : TLPBande ;
    i   : Integer ;
Begin
For i:=0 to Parent.ControlCount-1 do
  if (Parent.Controls[i] is TLPBande) and (TLPBande(Parent.controls[i]).top>=Top+Height)  then
     Begin
     Bnd:=TLPBande(Parent.controls[i]) ;
     Bnd.Top:=Bnd.Top+Combien ;
     End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBande.MouseMove(Shift: TShiftState; X, Y: Integer) ;
Var NPix : Integer ;
Begin
inherited mouseMove(Shift,X,Y) ;
if OkDrag then
   Begin
   NPix:=pixel.ConvertPixel(Height-(Height-y),TRUE) ;
   if (Pixel.ConvertPixel(Height,TRUE)<>NPix) and (NPix>=HauterMinimum) then
       Begin
       ReFaireBandes(NPix-Height) ;
       Height:=NPix ;
       End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPBAnde.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer) ;
Begin
if (OkDrag) and (Button=mbLeft) then OkDrag:=FALSE ;
If Not OkDrag then Screen.Cursor:=crDefault ;
inherited MouseUp(Button,Shift,X,Y) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPChamp.Create (AOwner : TComponent) ;
Begin
Inherited Create(Aowner) ;
Pixel:=TLPPixel.Create(Self) ;
FColonne:=-1 ; FLigne:=-1 ;
QuiChange:=qcAucun ;
Ctl3D:=FALSE ;
BorderStyle:=bsSingle ;
CharCase:=ecNormal ;
Font.Name:=LPNomFont ;
Font.Height:=LPTailleDef ;
PrinterVisible:=TRUE ;
Mask:='' ;
NbrChars:=10 ;
AutoSize:=FALSE ;
Taille:=lt10Cpi ;
setbounds(-1,-1,Pixel.AsChar[FALSE,LPTailleDef]*NbrChars,Pixel.asChar[TRUE,LPTailleDef]) ;
ChangeZoom:=FALSE ;
SiZero:=szImprimer ;
Evalue:=FALSE ;
FVersionControl:=-1 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Destructor TLPChamp.Destroy ;
Begin
Pixel.Free ;
Inherited Destroy ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPChamp.ChangeScale(M, D: Integer) ;
Begin
ChangeZoom:=TRUE ;
Inherited ChangeScale(M,D) ;
ChangeZoom:=FALSE ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Const LPChampVersionCourante = 2.00 ;
procedure TLPChamp.Loaded ;
var ver : Double ;
Begin
Ver:=VersionControl ;
if Ver<LPChampVersionCourante then
   Begin
   if Ver<0 then
      Begin //avant d'instauration de control de version....
      //Adaptation des Masques
      if pos(';',FMask)<=0 then FMask:=RecupOldMask(FMask) ; //XMG 11/06/03
      End ;
   Fmodified:=V_PGI.SAV ; //Si on est en train de parametrer des états en SAV on forcera l'enregistrement du nouveau format
                          // (si client on n'enregistrea les modif que si on modifie d'ailleurs)
   FVersionControl:=LPChampVersionCourante ;
   End ;

Inherited Loaded ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TLPChamp.SetParent ( AParent : TWinCOntrol) ;
Var OParent : TWinControl ;
Begin
if (Assigned(AParent)) and (not (AParent is TLPBande)) and (not (AParent is TLPBase)) then Raise EInvalidop.CreateFmt('Le parent doit être une LPBande',[]) ;
OParent:=Parent ;
inherited ;
if assigned(oParent) then LanceRealigneligne(OParent,Ligne) ;
if assigned(Parent)  then LanceRealigneligne(Parent,Ligne) ;
End ;
//XMG 16/09/03 Début
//////////////////////////////////////////////////////////////////////////////////
Procedure TLPChamp.GereCombos ;
var cb,TT,Fmt : String ;
    IsAbrege  : Boolean ;
Begin
LibCombo:=Value ;
if getNatureMasque='COMBO' then
   Begin
   fmt:=Mask ;
   if Fmt<>TraduireMemoire(gtfs(Recupoldmask('CVL'),';',LPChampMaskFmt)) then
      Begin
      Cb:=Libelle ;
      if Cb[1]='[' then Delete(Cb,1,1) ;
      if Cb[length(Cb)]=']' then Delete(Cb,Length(Cb),1) ;
      TT:=Get_Join(Cb) ;
      if TT<>'' then
         Begin
         isAbrege:=(fmt=TraduireMemoire(gtfs(Recupoldmask('CAB'),';',LPChampMaskFmt))) ;
         LibCombo:=rechDom(TT,Value,IsAbrege) ;
         if (LibCombo='') and (IsAbrege) then LibCombo:=rechDom(TT,Value,FALSE) ;
         End ;
      End ;
   End ;
end ;
//XMG 16/09/03 fin
//////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPChamp.SetValue ( AValue : String ) ;
Begin
FValue:=AValue ; //FindEtReplace(FindEtReplace(AValue,#13,'',TRUE),#10,'',TRUE) ; //XMG 23/06/03
GereCombos ; //XMG 16/09/03
Evalue:=TRUE ;
End ;
//////////////////////////////////////////////////////////////////////////////////
//XMG 25/06/03 début
Function TLPChamp.GetFormatedValue : String ;
Begin
Result:=FormatResult(Value) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPChamp.GetJustifiedValue : String ;
Begin
Result:=Copy(GetFormatedValue,1,NbrChars*NbrLignes) ;
//XMG 25/06/03 fin
if not IsCodeBarre then
   Begin
   if NbrLignes>1 then Result:=MC_WrapText(Result,NbrChars,#13) ; //XMG 25/07/03
   Result:=Justiftext(Result) ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function  TLPChamp.IsCodeBarre : Boolean ;
Begin
Result:=GetNatureMasque=LPNatCBarre ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPChamp.GetInvisibleParZero : Boolean ;
Begin
Result:=(SiZero<>szImprimer)  ;
if result then
   Begin
   If (IsNumeric(Value)) then Result:=(Valeur(Value)=0)
      else Result:=Trim(Value)='' ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPChamp.GetMask : String ;
begin
result:=FMask ;
if not (csWriting in componentstate) then result:=LP_RecupereFormatMasque(Result) ;
end ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPChamp.GetNatureMasque  : String ;
Begin
result:=LP_RecupereNatureMasque(Fmask) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPChamp.GetMasqueliteral : String ;
Begin
result:=Fmask ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPChamp.GetMaskFormule   : String ;
Var NAt : String ;
Begin
result:=Mask ;
Nat:=GetNaturemasque ;
if Nat=LPNatMaskUser then Nat:=LP_TrouveNatureUserMask(Result) ; 
if instring(Nat,['DATE','COMBO',LPNatCBarre]) then Result:='' ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPChamp.SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
type TValeursChanges = (vcTop,vcLeft,vcWidth,vcHeight) ;
Var LL,LT,LR,LB,Diff : Integer ;
    ValChange        : Set of TValeursChanges ;
    ABounds          : TRect ; //XMG 15/04/04
Begin
ABounds:=Rect(Aleft,Atop,AWidth,AHeight) ; //XMG 15/04/04
diff:=0 ;
if (QuiChange=qcAucun) and (Assigned(Parent)) then
   Begin
   if ((Parent is TLPBase) and (TLPBase(Parent).InRealigne))   or
      ((Parent is TLPBande) and (TLPBande(Parent).InRealigne))   then QuiChange:=qcRealigne ; //XMG 25/06/03
   End ;
ValChange:=[] ;
if Not ChangeZoom then
   Begin
   if (Parent is TLPBase) or (Parent is TLPBande) then
     Begin
     if Parent is TLPBase then
        Begin
        Ll:=TLPBase(Parent).LimitLeft ;
        Lt:=TLPBase(Parent).LimitTop ;
        Lr:=TLPBase(Parent).LimitRight ;
        Lb:=TLPBase(Parent).Limitbottom ;
        End else
        Begin
        Ll:=TLPBande(Parent).LimitLeft ;
        Lt:=TLPBande(Parent).LimitTop ;
        Lr:=TLPBande(Parent).LimitRight ;
        Lb:=TLPBande(Parent).Limitbottom ;
        End ;
//XMG 15/04/04 début
     if (not (csLoading in ComponentState)) and (isSpecialChampTexte(Libelle)) then
        Begin
        AHeight:=Pixel.TailleChar(TRUE) ;// Height ;
        AWidth:=Pixel.TailleChar(FALSE) ; //Width ;
        end else
        Begin
        AWidth:=MinintValue([LR-Aleft{+1},AWidth]) ;
        AWidth:=MaxIntValue([Pixel.TailleChar(FALSE,Font.Height),AWidth]) ;

        AHeight:=MinintValue([LB-ATop,AHeight]) ;
        AHeight:=MaxIntValue([Pixel.TailleChar(TRUE,Font.Height),AHeight]) ;

        Atop:=Pixel.Convertpixel(Atop,TRUE,LPTailleDef) ;
        AWidth:=Pixel.Convertpixel(AWidth,FALSE,Font.Height) ;
        AHeight:=Pixel.Convertpixel(AHeight,TRUE,Font.Height) ;
        End ;

     if ALeft<>Left then
        Begin
        ALeft:=maxintValue([Aleft,LL]) ;
        if Aleft+AWidth>=LR{-Pixel.TailleChar(FALSE,Font.Height)} then ALeft:=LR-AWidth ; //-Pixel.TailleChar(FALSE,Font.Height) ; 
        End ;
     if ATop<>Top then
        Begin
        ATop:=maxintValue([Atop,LT]) ;
        if Atop+AHeight>LB-Pixel.TailleChar(TRUE,Font.Height) then Atop:=LB-Pixel.TailleChar(TRUE,Font.Height) ;
        End ;
     End ;

   if (Left=-1) and (ALeft<0) then Aleft:=Left ;
   if (Top=-1) and (ATop<0) then ATop:=Top ;
   if (Atop<>Top) or (ATop<>ABounds.Top) then  ValChange:=ValChange+[vcTop] ;
   if (Aleft<>Left) or (ALeft<>ABounds.Left) then Begin ValChange:=ValChange+[vcLeft] ; Diff:=Left-ALeft ; End ;
   if (Width<>AWidth) or (AWidth<>ABounds.Right) then Begin ValChange:=ValChange+[vcWidth] ; if QuiChange=qcAucun then QuiChange:=qcWidth ; End ;
   if (AHeight<>Height) or (AWidth<>ABounds.Bottom) then ValChange:=ValChange+[vcHeight] ;
//XMG 15/04/04 fin
   modified:=(Modified) or (ValChange<>[]) ;
   End ;
Inherited SetBounds(Aleft,Atop,AWidth,AHeight) ;
if (Assigned(Parent)) and (not (QuiChange in [qcLargeur,qcLongeur,qcTaille,qcLigne,qcColonne,qcRealigne])) then
   Begin
   if vcLeft   in ValChange then Colonne:=Colonne-(ConvertenChar(Pixel,Diff,LPTailleDef,FALSE)-1) ;
   if vcTop    in ValChange then
      Begin
      Colonne:=ConvertenChar(Pixel,Left,LPTailleDef,FALSE) ;
      Ligne:=ConvertenChar(pixel,Top,LPTailleDef,TRUE) ;
      End ;
   if vcWidth  in ValChange then NbrChars:=ConvertenChar(Pixel,Width,Font.Height,FALSE)-1 ;
   if vcHeight in ValChange then NbrLignes:=ConvertenChar(Pixel,Height,Font.Height,TRUE)-1 ;
   End ;
if (Not ChangeZoom) and ((vcWidth in ValChange) or (vcHeight in ValChange)) then
   Begin
   if (Pixel<>nil) then
      Begin
      if QuiChange=qcWidth  then NbrChars:=AWidth div Pixel.TailleChar(FALSE,Font.Height) ;
      if QuiChange=qcHeight then NbrLignes:=AHeight div Pixel.TailleChar(TRUE,Font.Height) ;
      End ;
   End ;
if QuiChange in [qcHeight,qcWidth,qcRealigne] then QuiChange:=qcAucun ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPChamp.AssignTo ( Dest : TPersistent ) ;
Begin
If Dest is TLPChamp then
   With TLPChamp(Dest) do
      Begin
      Font.Assign(Self.Font) ;
      AlignTexte     :=Self.AlignTexte ;
      Ligne          :=Self.Ligne ;
      Colonne        :=Self.Colonne ;
      Caption        :=Self.Caption ;
      Libelle        :=Self.Libelle ;
      PrinterVisible :=Self.PrinterVisible ;
      Nom            :=Self.Nom ;
      Mask           :=Self.GetmasqueLiteral ;
      taille         :=Self.Taille ;
      NbrChars       :=Self.NbrChars ;
      NbrLignes      :=Self.NbrLignes ;
      Value          :=Self.Value ;
      Visible        :=Self.Visible ;
      SiZero         :=Self.SiZero ;
      tag            :=Self.tag ;
      End else Inherited AssignTo(Dest) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPChamp.SetPrinterVisible ( Value : Boolean) ;
Begin
if Value=PrinterVisible then Exit ;
FPrinterVisible:=Value ;
Modified:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPChamp.SetFAlignTexte ( Value : TAlignment) ;
Begin
FAlignTexte:=Value ;
Modified:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPChamp.SetLibelle ( Value : String ) ;
Begin
FLibelle:=Value ;
Modified:=TRUE ;
Hint:=copy('@',1,ord(copy(Hint,1,1)='@'))+Libelle ; ShowHint:=(trim(Hint)<>'') and IsDesigning ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TLPChamp.SetMask ( Value : String ) ;
Begin
if not (csLoading in ComponentState) then //XMG 11/06/03
   Begin
   Value:=LP_ConstruitMasque(Value) ;
   if Value=Mask then Exit ;
   End ; 
FMask:=Value ;
Modified:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPChamp.SetNom ( Value : String ) ;
Begin
if Value=Nom then Exit ;
FNom:=Value ;
Modified:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPChamp.SetTaille ( Value : TLPTailleChar) ;
Var XX : Integer ;
Begin
if QuiChange=qcAucun then QuiChange:=qcTaille ;
FTaille:=Value ;
case Taille of
  lt17Cpi : Font.Height:=LPTaille17 ;
  lt12Cpi : Font.Height:=LPTaille12 ;
  lt10Cpi : Font.Height:=LPTaille10 ;
  else      Font.Height:=LPTaille05 ;
  End ;
XX:=Pixel.TailleChar(FALSE,Font.Height)*NbrChars ;
if XX<>Width then Width:=XX ;
if Assigned(Parent) then LanceRealigneligne(Parent,Ligne) ;
Modified:=TRUE ;
if QuiChange=qcTaille then QuiChange:=qcAucun ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPChamp.SetLigne ( Value : Integer) ;
var Oligne : integer ;
Begin
if Value<1 then Value:=1 ; 
if Value=Ligne then exit ;
if QuiChange=qcAucun then QuiChange:=qcLigne ;
OLigne:=Ligne ;
FLigne:=Value ;
modified:=TRUE ;
if assigned(parent) then
   Begin
   if (OLigne<>Ligne) and (OLigne>-1) then LanceRealigneligne(Parent,OLigne) ;
   if QuiChange=qcLigne then Top:=ConvertEnPixel(Pixel,Ligne,LPTailleDef,true) ;
   if (OLigne<>Ligne) then LanceRealigneligne(Parent,Ligne) ;
   End ;
if QuiChange=qcLigne then QuiChange:=qcAucun ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPChamp.SetColonne ( Value : Integer) ;
Begin
if Value<1 then Value:=1 ;
if Value=colonne then exit ;
if QuiChange=qcAucun then QuiChange:=qcColonne ;
FColonne:=Value ;
modified:=TRUE ;
if (QuiChange=qcColonne) and (assigned(parent)) then LanceRealigneligne(Parent,Ligne) ;
if QuiChange=qcColonne then QuiChange:=qcAucun ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPChamp.SetNbrChars( Value : Integer) ;
Var XX : Integer ;
Begin
if Value<1 then Value:=1 ;
if Value=NbrChars then exit ;
if QuiChange=qcAucun then QuiChange:=qcLongeur ;
FNbrChars:=Value ;
if (QuiChange=qcLongeur) then
   Begin
   XX:=Pixel.TailleChar(FALSE,Font.Height)*NbrChars ;
   if (XX<>Width) then Width:=XX ;
   End ;
if Assigned(Parent) then LanceRealigneligne(Parent,Ligne) ;
if QuiChange=qcLongeur then QuiChange:=qcAucun ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPChamp.SetNbrLignes( Value : Integer) ;
Var ii,XX : Integer ;
Begin
if Value<1 then Value:=1 ;
if Value=NbrLignes then exit ;
if QuiChange=qcAucun then QuiChange:=qcLargeur ;
FNbrLignes:=Value ;
if (QuiChange=qcLargeur) then
   Begin
   XX:=Pixel.TailleChar(TRUE,Font.Height)*NbrLignes ;
   if (XX<>Height) then Height:=XX ;
   End ;
if Assigned(Parent) then
   Begin
   For ii:=Ligne to Ligne+NbrLignes-1 do
       LanceRealigneligne(Parent,ii) ;
   End ;
if QuiChange=qcLargeur then QuiChange:=qcAucun ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPChamp.GetLigne : Integer ;
Begin
Result:=FLIgne ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPChamp.GetColonne : Integer ;
Begin
Result:=FColonne ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPChamp.WMPaint(var Message: TWMPaint);
Begin
ControlState:=ControlState+[csCustomPaint] ;
inherited ;
ControlState:=ControlState-[csCustomPaint] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPChamp.PaintWindow(DC: HDC);
const WordWraps : array[Boolean] of Word = (0, DT_WORDBREAK);
      Alignment : array[TAlignment] of Word = (DT_LEFT, DT_RIGHT, DT_CENTER) ; //XMG 25/06/036
Var St       : String ;
    Rct,Rct2 : TRect ;
    SFont    : HFont ;
    Coef     : integer ;
    SBrush   : HBrush ;
Begin
Rct:=ClientRect ;
if IsCodebarre then
   Begin
   ChargeCBarreImage ;
   stretchblt(DC,rct.left,rct.top,rct.right,rct.Bottom,CBarreim.Canvas.Handle,0,0,cBarreIm.Width,CBarreIm.Height,SRCCOPY) ;
   if IsDesigning then
      Begin
      St:=Mask ;
      St:=copy(St,length(LPStCBarre)+1,length(St)) ;
      End else St:=Value ;
   Rct2:=Rct ;
   DrawText(dc,PChar(St),Length(St),Rct2,DT_CALCRECT or DT_SINGLELINE or DT_EXPANDTABS or DT_NOPREFIX) ;

   Font.Height:=Rct2.Bottom -2 ;

   DrawText(dc,PChar(St),Length(St),Rct2,DT_CALCRECT or DT_SINGLELINE or DT_EXPANDTABS or DT_NOPREFIX) ;

   Coef:=(Rct.Bottom-Rct2.Bottom) div 3 ;
   Rct.Top:=Coef ; Rct.Bottom:=Rct.Bottom-Coef ;
   Coef:=(Rct.Right-Rct2.Right) div 2 ;
   Rct.Left:=coef ; Rct.Right:=Rct.Right-Coef ;
   SFont:=SelectObject(DC,Font.Handle) ;
   SBrush:=SelectObject(DC,Brush.Handle) ;
   DrawText(dc,PChar(St),Length(St),Rct,DT_SINGLELINE or DT_EXPANDTABS or DT_NOPREFIX) ;
   SelectObject(DC,SFont) ;
   SelectObject(DC,SBrush) ;
   end else
   Begin
   if IsDesigning then St:=Libelle else St:=GetJustifiedValue ; //XMG 25/06/03
   St:=FindEtReplace(St,#13,' ',TRUE) ;
   if BorderStyle<>bsNone then
      Begin
      with rct do
        Rectangle(DC,Left,top,Right,Bottom) ;
      inflaterect(rct,-1,-1) ;
      End ;
   SFont:=SelectObject(DC,Font.Handle) ;
   SBrush:=SelectObject(DC,Brush.Handle) ;
   DrawText(dc,PChar(St),Length(St),Rct,DT_EXPANDTABS or wordwraps[NbrLignes>1] or
                                        Alignment[AlignTexte] or DT_NOPREFIX) ; //XMG 25/06/03
   SelectObject(DC,SFont) ;
   SelectObject(DC,SBrush) ;
   End ;
if not printerVisible then
   with TCanvas.Create do
     try
        Handle:=DC ;
        Brush.Color:=clblack ;
        Pen.Color   := Brush.Color ;
        Pen.Mode    := pmNot ;
        Pen.Style   := psDot;
        Pen.Width   := 1 ;
        Rectangle(0,0,width,height) ;
       finally
        Free ;
       End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPChamp.CMCtl3DChanged(var Message: TMessage);
begin
if Ctl3D then Ctl3D:=FALSE ; //On anule systematiquement l'effet 3D
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPChamp.WMSetFocus(var Message: TWMSetFocus);
Begin
Message.Result:=0 ; // On annule systematiquement le focus
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPChamp.MouseMove (Shift : TShiftState; X,Y : Integer) ;
Begin
Cursor:=crArrow ; //on dessine systematiquetent la feche
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TrouveMarque ( Posibles,Value : String) : String ;
Var ii : integer ;
Begin
Result:='' ;
For ii:=1 to Length(Posibles) do
    if pos(Copy(posibles,ii,1),Value)<=0 then Begin Result:=Copy(Posibles,ii,1) ; Break ; End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPChamp.FormatResult (AValue : String ) : String ;
Var Fmt,Nat,CBArre,Marque : String ;
    Nm                    : Double ;
Begin
Fmt:=Mask ;
Nat:=getNatureMasque ;
If Nat=LPNatMaskUser then nat:=LP_TrouveNatureUserMask(fmt) ;
if Nat='DATE' then AValue:=FormatDatetime(Fmt,Valeur(AValue)) else
if Nat='NUM' then
   Begin
   Nm:=Valeur(AValue) ;
   if pos(TraduireMemoire(gtfs(Recupoldmask('LET'),';',LPChampMaskFmt)),fmt)>0 then
      Begin
      AValue:=ChiffreLettre(Nm,trim(Copy(' €',1+ord(pos(traduireMemoire('(Euro)'),Fmt)>0),1)))+'---' ;
      End else
      Begin
      if Copy(Fmt,1,1)='€' then
         Begin
         Delete(Fmt,1,1) ;
         if (Assigned(ConvertEuro)) and (not V_PGI.TenueEuro) then Nm:=ConvertEuro(Nm)
         End ;
      if (Nm<>0) and (pos('#.##',fmt)>0) then
         Fmt:=gtfs(RecupOldMask('2SZ'),';',LPChampMaskFmt) ; //???
      AValue:=FormatFloat(Fmt,Nm) ;
      End ;
   End else
If Nat='COMBO' then AValue:=LibCombo else //XMG 16/09/03
If Nat=LPNatCBarre then
   Begin
   CBarre:=copy(Fmt,length(LPStCBarre)+1,length(Fmt)) ;
   if stricomp('(UPC-A)',pchar(CBarre))=0   then AValue:=sright(stringofchar('0',12)+AValue,12) else
   if stricomp('(UPC-E)',pchar(CBarre))=0   then AValue:='0'+sright(stringofchar('0',10)+AValue,10)else
   if stricomp('(EAN13)',pchar(CBarre))=0   then AValue:=sright(stringofchar('0',13)+AValue,13) else
   if stricomp('(EAN8)',pchar(CBarre))=0    then AValue:=sright(stringofchar('0',8)+AValue,8) else
   if stricomp('(CODE39)',pchar(CBarre))=0  then AValue:='*'+FindEtReplace(AValue,'*',' ',TRUE)+'*' else
   if stricomp('(ITF)',pchar(CBarre))=0     then AValue:=Sright('0'+AValue,length(AValue)+ord(length(Avalue) mod 2=1)) else
   if stricomp('(CODABAR)',pchar(CBarre))=0 then
      Begin
      Marque:=TrouveMarque('ABCD',AValue) ;
      AValue:=MArque+AValue+Marque ;
      End else
   if stricomp('(CODE93)',pchar(CBarre))=0  then else
   if stricomp('(CODE128)',pchar(CBarre))=0 then
      Begin
      AValue:='{B'+findetreplace2(AValue,'{','{{',TRUE)
      End  else
      ;
   End ;
Result:=AValue ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPChamp.GetVersion : Double ;
Begin
Result:=FVersionControl ;
if (Not (csLoading in ComponentState)) and (Result<>LPChampVersionCourante) then
   Result:=LPChampVersionCourante ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Procedure TLPChamp.Initialise ( IsEof : Boolean ) ;
Begin
if (TLPBande(Parent).TypeBande=lbpied) and (not isEof) then Value:=' **** ' else Value:='' ;
Evalue:=FALSE ;
End ;
/////////////////////////////////////////////////////////////////////////////////
Function TLPChamp.Police : TLP_SetSqcesPossibles ;
Begin
Result:=[] ;
if not IsSpecialChamp(Self) then
if IsCodeBarre then include(Result,spCodeBarre) else //les codesbarres sont excluients
   Begin
   Case Taille of
     lt05cpi : include(Result,sp05Cpi) ;
     lt10cpi : Include(Result,sp10Cpi) ;
     lt12cpi : include(Result,sp12Cpi) ;
     lt17cpi : include(Result,sp17Cpi) ;
     else Include(Result,sp10Cpi) ;
     End ;
   if fsbold      in Font.style then include(Result,spbold) ;
   if fsItalic    in Font.style then include(Result,spItalic) ;
   if fsUnderline in Font.style then include(Result,spUnderline) ;
   End ;
End ;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPChamp.IsDesigning : Boolean ;
Begin
Result:=FALSE ;
if (Assigned(Parent)) then
   if (Parent is TLPBase) then Result:=TLPBase(Parent).Designing else
   if (Parent is TLPBande) then Result:=TLPBande(Parent).Designing ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPChamp.JustifText( Texte : String ) : String ;
Var Combien,jj     : Integer ;
    Blancs,St,UnSt : String ;
Begin
Result:='' ;
for jj:=1 to NbrLignes do
   Begin
   if NbrLignes<=1 then St:=texte
                   else St:=gtfs(Texte,#13,jj) ;
   Combien:=(NbrChars-Length(St)) ;
   if Combien>=0 then
      Begin
      Blancs:=Format_String('',Combien) ;
      Case AlignTexte of
        taLeftJustify  : Result:=Result+St+Blancs ;
//XMG 01/07/03 début
        taCenter       : Begin
                         UnSt:=copy(Blancs,1,combien div 2)
                              +Copy(St,maxintValue([1,Length(St)-NbrChars]),NbrChars+1) ;
                         UnSt:=UnSt+Copy(Blancs,1,nbrChars-length(UnSt)) ;
                         Result:=Result+UnSt ;
                         End ;
//XMG 01/07/03 fin
        taRightJustify : Result:=Result+Blancs+Copy(St,maxintValue([1,length(St)-NbrChars+1]),NbrChars+1) ;
        End ;
      End ;
   if NbrLignes>1 then Result:=Result+#13 ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
constructor TLPSqcesLigne.Create ;
Begin
Inherited Create ;
FLarg:=-1 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPSqcesLigne.SetFLarg (AValue : Integer ) ;
Begin
If AValue=FLarg then exit ;
try
  SetLength(FStp,AValue) ;
  SetLength(FValeurs,AValue) ;
  FLarg:=AValue ;
 Except
  On EOutofMemory do ShowMessage('Il n''est plus possible de modifier la taille.') ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPSqcesLigne.SetFSTP ( Index : Integer ; AValue :TLP_SetSqcesPossibles ) ;
Begin
If (Index>Largeur) or (FStp[Index]=AValue) then Exit ;
FStp[Index]:=AValue ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPSqcesLigne.SetFValeurs( Index : Integer ; AValue : String ) ;
Begin
If (Index>Largeur) then exit ; //or ((VarType(FValeurs[index])<>VarEmpty) and (Vartype(FValeurs[Index])=vartype(AValue))) then Exit ;
FValeurs[Index]:=AValue ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPSqcesLigne.GetFStp ( Index : INteger ) : TLP_SetSqcesPossibles ;
Begin
Result:=[] ;
If Index>Largeur then Exit ;
Result:=FStp[Index] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPSqcesLigne.GetFValeurs( Index : INteger ) : String ;
Begin
Result:=#0 ;
If Index>Largeur then Exit ;
Result:=FValeurs[Index] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPSqcesLIgne.Clear ;
Var i : Integer ;
Begin
For i:=0 to Largeur-1 do
    Begin
    Sqce[i]:=[] ;
    Valeurs[i]:=#0 ;
    End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPSqcesLigne.Assign ( Source : TObject ) ;
var i : Integer ;
Begin
If Source is TLPSqcesLigne then
   Begin
   Clear ;
   Largeur:=TLPSqcesLigne(Source).Largeur ;
   For i:=0 to Largeur-1 do
       Begin
       Sqce[i]:=TLPSqcesLIgne(Source).Sqce[i] ;
       Valeurs[i]:=TLPSqcesLIgne(Source).Valeurs[i] ;
       End ;
   End else Raise EConvertError.CreateFmt('On ne peut pas assigne ce type d''object (%s)',[Source.ClassName]) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TFLPProgress.Create ( AOwner : TComponent ) ;
Begin
Inherited Create(AOwner) ;
FForm:=TForm.Create(Application) ;
FCanceled:=FALSE ;
With FForm do
  Begin
  Position:=PoScreenCenter;
  Height:=140;
  Width:=400;
  BorderIcons:=[];
  BorderStyle:=bsDialog;
  Caption:=NomHalley ;
  End ;

FProgress:=TEnhancedGauge.Create(FForm) ;
With FProgress do
  Begin
  Parent:=FForm ;
  Left:=8 ;
  Width:=376 ;
  Top:=49 ;
  Height:=16 ;
  ForeColor:=clActiveCaption ;
  End ;

FLabel:=THLabel.Create(FForm) ;
With FLabel do
  Begin
  Parent:=FForm ;
  Left:=8 ;
  Top:=24 ;
  End ;

BCancel:=TButton.Create(FForm) ;
With BCancel do
  Begin
  Parent:=FForm ;
  Top:=81 ;
  Left:=156 ;
  Height:=25 ;
  Width:=80 ;
  Caption:=TraduireMemoire(MC_MsgErrDefaut(1040)) ;
  OnCLick:=BCancelClick ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Destructor TFLPProgress.Destroy ;
Begin
BCancel.Free ;
Flabel.Free ;
FProgress.Free ; 
FForm.Free ;
Inherited Destroy ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPProgress.BCancelCLick(Sender : TObject ) ;
Begin
FCanceled:=TRUE ;
FForm.invalidate ;
Application.ProcessMessages ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPProgress.GetProgress : Integer ;
Begin
Result:=FProgress.Progress ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPProgress.SetMax ( MaxValue : Integer) ;
Begin
FProgress.MaxValue:=MaxValue ;
FForm.invalidate ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPProgress.SetValue ( NewValue : Integer ) ;
Begin
FProgress.Progress:=NewValue ;
FForm.invalidate ;
Application.ProcessMessages ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPProgress.SetLibelle ( NewLibelle : String ) ;
Begin
FLabel.Caption:=NewLibelle ;
FForm.invalidate ;
Application.ProcessMessages ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPProgress.GetLibelle  : String ;
Begin
Result:=FLabel.Caption ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPProgress.Show ;
Begin
FCanceled:=FALSE ;
FForm.Show ;
FForm.invalidate ;
Application.processMessages ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLpProgress.Hide ;
Begin
FForm.Hide ;
FForm.invalidate ;
FCanceled:=FALSE ; 
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPProgress.ForceCancel ;
Begin
FCanceled:=TRUE ;
Application.processmessages ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPImage.Create (AOwner : TComponent) ;
Begin
Inherited Create(Aowner) ;
Pixel:=TLPPixel.Create(Self) ;
PrinterVisible:=TRUE ;
setbounds(-1,-1,Pixel.AsChar[FALSE,LPTailleDef]*10,Pixel.asChar[TRUE,LPTailleDef]*10) ;
QuiChange:=[] ;
Modified:=FALSE ;
Fichier:='' ;
Stretch:=TRUE ;
BMP:=nil ;
PictureBW:=FALSE ;
Picture.onChange:=ChangePicture ;
FVersionControl:=-1 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Destructor TLPImage.Destroy ;
Begin
if Assigned(BMP) then BMP.Free ;
Pixel.Free ;
Inherited Destroy ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImage.SetParent ( AParent : TWinCOntrol) ;
Var OParent : TWinControl ;
Begin
if (Assigned(AParent)) and (not (AParent is TLPBande)) and (not (AParent is TLPBase)) then Raise EInvalidop.CreateFmt('Le parent doit être une LPBande',[]) ;
OParent:=Parent ;
inherited ;
if assigned(oParent) then LanceRealigneligne(OParent,Ligne) ;
if assigned(Parent)  then LanceRealigneligne(Parent,Ligne) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImage.ChangePicture ( Sender : TObject ) ;
Begin
PictureBW:=FALSE ;
Stretch:=TRUE ;
invalidate ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImage.SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
Var LL,LT,LR,LB : Integer ;
Begin
if QuiChange=[] then QuiChange:=[qcWidth,qcHeight] ;
if (Parent is TLPBande) then
  Begin
  Ll:=TLPBande(Parent).LimitLeft ;
  Lt:=TLPBande(Parent).LimitTop ;
  Lr:=TLPBande(Parent).LimitRight ;
  Lb:=TLPBande(Parent).Limitbottom ;
   if ALeft<>Left then
      Begin
      ALeft:=maxintValue([Aleft,LL]) ;
      ALeft:=minIntValue([Aleft+AWidth,LR])-AWidth ;
      End ;
   if ATop<>Top then
      Begin
      ATop:=maxintValue([Atop,LT]) ;
      ATop:=minintValue([Atop+AHeight,LB])-AHeight ;
      End ;
  AWidth:=MaxIntValue([Pixel.TailleChar(FALSE,LPTailleDef),AWidth]) ;
  AWidth:=MinintValue([LR-Aleft,AWidth]) ;

  AHeight:=MaxIntValue([Pixel.TailleChar(TRUE,LPTailleDef),AHeight]) ;
  AHeight:=MinintValue([LB-ATop,AHeight]) ;

  Atop:=Pixel.Convertpixel(Atop,TRUE,LPTailleDef) ;
  AWidth:=Pixel.Convertpixel(AWidth,FALSE,LPTailleDef) ;
  AHeight:=Pixel.Convertpixel(AHeight,TRUE,LPTailleDef) ;
  End ;
if (Left=-1) and (ALeft<0) then Aleft:=Left ;
if (Top=-1) and (ATop<0) then ATop:=Top ;
Inherited SetBounds(Aleft,Atop,AWidth,AHeight) ;
if Assigned(Pixel) then
   Begin
   if qcWidth in QuiChange  then Longeur:=AWidth div Pixel.TailleChar(FALSE,LPTailleDef) ;
   if qcHeight in QuiChange then Largeur:=AHeight div Pixel.TailleChar(TRUE,LPTailleDef) ;
   End ;
if QuiChange=[qcWidth,qcHeight] then QuiChange:=[] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImage.AssignTo ( Dest : TPersistent ) ;
Begin
If Dest is TLPImage then
   With TLPImage(Dest) do
      Begin
      Ligne          :=Self.Ligne ;
      Colonne        :=Self.Colonne ;
      Fichier        :=Self.Fichier ;
      PrinterVisible :=Self.PrinterVisible ;
      Nom            :=Self.Nom ;
      Longeur        :=Self.Longeur ;
      Largeur        :=Self.Largeur ;
      Height         :=Self.Height ;
      Width          :=Self.Width ;
      Visible        :=Self.Visible ;
      Picture.Assign(Self.Picture) ;
      End else Inherited AssignTo(Dest) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImage.SetNom ( Value : String ) ;
Begin
if Value=Nom then Exit ;
FNom:=Value ;
Modified:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImage.SetFichier ( Value : String ) ;
Var Pc  : TPicture ;
    Nom : String ;
Begin
Value:=Uppercase(Trim(Value)) ;
Nom:=Value ;
if (Nom='**LOGO**') and (Assigned(LP_GetParamSoc)) then Nom:=trim(LP_GetParamSoc('SO_LOGOGC')) ;
if ((Value=Fichier) and (Value<>'**LOGO**')) or ((not (csLoading in componentState)) and (Not FileExists(Nom))) then exit ;
FFichier:=Value ;
Hint:=copy('@',1,ord(copy(Hint,1,1)='@'))+Fichier ; ShowHint:=(trim(Hint)<>'') ;
if ((not (csLoading in componentState)) and (FileExists(Nom))) then
   Begin
   Pc:=TPicture.Create ;
   try
     pc.LoadFromFile(Nom) ;
     Picture.Assign(Pc) ;
    finally
     Pc.Free ;
    End ;
   End ; 
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPImage.SetLargeur ( Value   : Integer)         ;
Var XX : Integer ;
Begin
if Value<1 then Value:=1 ; 
if Value=Largeur then exit ;
if QuiChange=[] then QuiChange:=[qcLargeur] ;
FLargeur:=Value ;
if (qcLargeur in QuiChange) then
   Begin
   XX:=Pixel.TailleChar(TRUE,LPTailleDef)*Largeur ;
   if (XX<>Height) then Height:=XX ;
   End ;
RealigneImage(Ligne) ;
if QuiChange=[qcLargeur] then QuiChange:=[] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPImage.SetLongeur( Value   : Integer)         ;
Var XX : Integer ;
Begin
if Value<1 then Value:=1 ;
if Value=Longeur then exit ;
if QuiChange=[] then QuiChange:=[qcLongeur] ;
FLongeur:=Value ;
if (qcLongeur in QuiChange) then
   Begin
   XX:=Pixel.TailleChar(FALSE,LPTailleDef)*Longeur ;
   if (XX<>Width) then Width:=XX ;
   End ;
RealigneImage(Ligne) ;
if QuiChange=[qcLongeur] then QuiChange:=[] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImage.GetLigne : Integer ;
Begin
Result:=FLIgne ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPImage.GetColonne : Integer ;
Begin
Result:=FColonne ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImage.RealigneImage ( LigneIni : integer ) ;
Var i : Integer ;
Begin
if Not assigned(Parent) then exit ;
for i:=ligneIni to LigneIni+Largeur do LanceRealigneligne(Parent,i) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPImage.SetLigne ( Value : Integer) ;
var OLigne : Integer ;
Begin
if Value<1 then Value:=1 ; 
if Value=Ligne then exit ;
if QuiChange=[] then QuiChange:=[qcLigne] ;
OLigne:=Ligne ;
FLigne:=Value ;
modified:=TRUE ;
if assigned(parent) then
   Begin
   if (OLigne<>Ligne) and (OLigne>-1) then RealigneImage(OLigne) ;
   if QuiChange=[qcLigne] then Top:=ConvertEnPixel(Pixel,Ligne,LPtailleDef,true) ;
   if OLigne<>Ligne then RealigneImage(Ligne) ;
   End ;
if QuiChange=[qcLigne] then QuiChange:=[] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TLPImage.SetColonne ( Value : Integer) ;
Begin
if Value<1 then Value:=1 ; 
if Value=colonne then exit ;
if QuiChange=[] then QuiChange:=[qcColonne] ;
FColonne:=Value ;
modified:=TRUE ;
if (QuiChange=[qcColonne]) and (assigned(parent)) then LanceRealigneligne(Parent,Ligne) ;
if QuiChange=[qcColonne] then QuiChange:=[] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPImage.ImprimeImage ;
var x,y     : Integer ;
    Im2,im3 : TBitmap ;
begin
if PictureBW then exit ;
im3:=Nil ;
Stretch:=FALSE ;
Im2:=TBitMap.Create ;
try
    Im3:=TBitMap.Create ;
    Im2.Width:=width ;
    im2.height:=height ;
    im2.Canvas.StretchDraw(rect(0,0,im2.width,im2.height),Picture.Graphic) ;
    im3.height:=Im2.height ;
    Im3.width:=im2.width ;
     for x:=0 to Im2.Height do
        Begin
        for y:=0 to Im2.Width-1 do
           Begin
           im3.canvas.pixels[y,x]:=clwhite*ord(im2.canvas.pixels[y,x]<>clblack) ;
           End ;
        End ;
    picture.assign(im3) ;
    PictureBW:=TRUE ;
    Refresh ;
  finally
    Im2.Free ;
    if im3<>nil then im3.free ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPImage.ImagetoSt ( Ligne : Integer ) : String ;
Var x,Y,l,h,Val,pix,Coef : Integer ;
Begin
if Not Assigned(BMP) then
   Begin
   if Not PictureBw then ImprimeImage ;
   BMP:=TBitMap.Create ;
   BMP.Assign(Picture.Bitmap) ;
   End ;
Result:='' ;
Coef:=2 ;
Ligne:=Largeur-Ligne-1 ;
l:=8*ligne*Coef ;
h:=8*(Ligne+1)*Coef-1 ;
for y:=0 to BMP.Width-1 do
   Begin
   val:=0 ;
   For x:=h downto l do //to h do
      if x mod 2 =0 then
         Begin
         if x<=BMP.Height then pix:=BMP.canvas.pixels[y,x] else Pix:=clWhite ;
         val:=val shl 1+ord(Pix=clblack) ;
         End ;
   Result:=Result+chr(Val) ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPImage.Police : TLP_SetSqcesPossibles ;
Begin
Result:=[spBitMap] ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Const LPImageVersionCourante = 2 ;
procedure TLPImage.Loaded ;
var ver : Double ;
Begin
Ver:=VersionControl ;
if Ver<LPImageVersionCourante then
   Begin
   //Les adaptations ici
   Fmodified:=V_PGI.SAV ; //Si on est en train de parametrer des états en SAV on forcera l'enregistrement du nouveau format
                          // (si client on n'enregistrea les modif que si on modifie d'ailleurs)
   FVersionControl:=LPImageVersionCourante ;
   End ;
Inherited Loaded ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
Function TLPImage.GetVersion : Double ;
Begin
Result:=FVersionControl ;
if (Not (csLoading in ComponentState)) and (Result<>LPImageVersionCourante) then
   Result:=LPImageVersionCourante ;
End ;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPSauvgarde.Create ( AOwner : TComponent ) ;
Begin
Inherited create (AOwner) ;
FForm:=TForm.Create(Self) ;
ParamLP:=TLPparam.Create(Self) ;
BaseLP:=TLPBase.Create(FForm) ;
BaseLP.Parent:=FForm ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Destructor  TLPSauvGarde.Destroy ;
Begin
if assigned(BaseLP) then Begin BaseLP.Free ; BaseLP:=nil ; End ;
if assigned(ParamLP) then Begin ParamLP.Free ; ParamLP:=nil ; End ;
if assigned(FForm) then Begin FForm.Free ; FForm:=nil ; End ;
inherited destroy ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure LibereSauvgardeLP ;
Begin
if assigned(SauvgardeLP) then
   Begin
   SauvgardeLP.free ;
   SauvgardeLP:=Nil ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
initialization
   RegisterClasses([TLPChamp,TLPIMage]) ;
   {//YCP 20/05/03} SauvgardeLP:=nil ;
   {//YCP 20/05/03} LP_GetParamSoc:=nil ;
   {//YCP 20/05/03} CBarreIm:=nil ;

finalization
   LibereSauvgardeLP ;
   LibereCBarreIm ;
end.

