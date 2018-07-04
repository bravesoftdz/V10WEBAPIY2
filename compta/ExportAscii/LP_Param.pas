unit LP_Param;

interface

uses Windows, Dialogs, ExtCtrls, Menus, Db, DBTables, StdCtrls, Buttons,
     ColMemo, Hctrls, ComCtrls, Grids, HPanel, Spin, HTB97, Controls, 
     LP_Base, Classes, Forms, HSysMenu, HMenu97, LP_Util, Graphics, Mask ;

Procedure Param_LPTexte ( Inside : THpanel ; TypEtat,NatEtat,CodEtat : String ; CreerEtat : Boolean ) ;
//XMG 02/04/04 30/03/04 début
//TypEtat = 'T' .- Modèle Texte (Voir MC_lib.TypeEtatTexte)
//        = 'F' .- Fichier Exportation (voir MC_lib.TypeEtatExportAscii)
//XMG 02/04/04 30/03/04 fin
type
  TFLPDessin = class(TForm)
    HMTrad: THSystemMenu;
    TopDock: TDock97;
    Q1: TQuery;
    PFichier: TToolbar97;
    MenuMain: TMainMenu;
    Fichier1: TMenuItem;
    BNouveau: TToolbarButton97;
    BLire: TToolbarButton97;
    bEnreg: TToolbarButton97;
    BEnregSous: TToolbarButton97;
    BottomDock: TDock97;
    LeftDock: TDock97;
    RightDock: TDock97;
    Timer: TTimer;
    SbFonds: TScrollBox;
    LPBase: TLPBase;
    EtatPanel: TToolWindow97;
    HPanel2: THPanel;
    BValiderEtat: TToolbarButton97;
    BCancelEtat: TToolbarButton97;
    CtrPanel: TToolWindow97;
    TCtrNom: THLabel;
    TCtrX: THLabel;
    TCtrY: THLabel;
    TCtrCaption: THLabel;
    TCtrLong: THLabel;
    TCtrMask: THLabel;
    CtrNom: TEdit;
    HPanel1: THPanel;
    BValiderPnl: TToolbarButton97;
    BCancelPnl: TToolbarButton97;
    CtrX: TSpinEdit;
    CtrY: TSpinEdit;
    CtrLong: TSpinEdit;
    CtrVisible: TCheckBox;
    CtrMask: THValComboBox;
    CtrAlign: TRadioGroup;
    TCtrlBande: THLabel;
    CtrlBande: THValComboBox;
    PBottom: TPanel;
    bImport: TToolbarButton97;
    bExport: TToolbarButton97;
    dlgIMport: TOpenDialog;
    DlgExport: TSaveDialog;
    Nouveau1: TMenuItem;
    Recuperer1: TMenuItem;
    Enregistre1: TMenuItem;
    Enregistresous1: TMenuItem;
    N1: TMenuItem;
    Importermodle1: TMenuItem;
    Exportermodele1: TMenuItem;
    N2: TMenuItem;
    Quitter1: TMenuItem;
    Propits1: TMenuItem;
    Sourcedesdonns1: TMenuItem;
    Champs1: TMenuItem;
    Insererunchamp1: TMenuItem;
    Propits2: TMenuItem;
    Supprimer1: TMenuItem;
    Aide1: TMenuItem;
    N3: TMenuItem;
    BndPanel: TToolWindow97;
    HPanel3: THPanel;
    BValideBnd: TToolbarButton97;
    BCancelBnd: TToolbarButton97;
    BndVisible: TCheckBox;
    BndCondition: TEdit;
    BBndSQL: TToolbarButton97;
    TBndCOndition: THLabel;
    TBBndSQL: THLabel;
    CtrSiZero: THValComboBox;
    TCtrSiZero: THLabel;
    TWImage: TToolWindow97;
    TImaNom: THLabel;
    ImaNom: TEdit;
    TImaX: THLabel;
    ImaX: TSpinEdit;
    TImaY: THLabel;
    ImaY: TSpinEdit;
    tImaLong: THLabel;
    ImaLong: TSpinEdit;
    ImaVisible: TCheckBox;
    ImaLarg: TSpinEdit;
    TImaLarg: THLabel;
    tImaContenu: THLabel;
    ImaContenu: THCritMaskEdit;
    HPanel4: THPanel;
    BValiderIma: TToolbarButton97;
    BCancelIma: TToolbarButton97;
    TImaBande: THLabel;
    ImaBAnde: THValComboBox;
    ImaLogo: TCheckBox;
    pgGeneral: TPageControl;
    shPapier: TTabSheet;
    GrPapier: TGroupBox;
    TLargeur: THLabel;
    FLongeur: THLabel;
    Largeur: THNumEdit;
    FRouleau: TCheckBox;
    Longeur: THNumEdit;
    tailleEn: TRadioGroup;
    tsMarges: TTabSheet;
    GrMarges: TGroupBox;
    TMargeH: THLabel;
    TMargeG: THLabel;
    TMargeD: THLabel;
    TMargeB: THLabel;
    MargeH: TSpinEdit;
    MargeG: TSpinEdit;
    MargeD: TSpinEdit;
    MargeB: TSpinEdit;
    shDiverses: TTabSheet;
    GrBandes: TGroupBox;
    ChkEntete: TCheckBox;
    ChkDetail: TCheckBox;
    ChkPIed: TCheckBox;
    CntDetail: TSpinEdit;
    CntEntete: TSpinEdit;
    CntPied: TSpinEdit;
    GrMessures: TRadioGroup;
    TsRupt: TTabSheet;
    Bevel2: TBevel;
    TFRupt1: THLabel;
    TFRupt2: THLabel;
    TFRupt3: THLabel;
    TFRupt4: THLabel;
    TFRupt5: THLabel;
    TFRupt6: THLabel;
    FRupt1: THValComboBox;
    FRupt2: THValComboBox;
    FRupt3: THValComboBox;
    FRupt4: THValComboBox;
    FRupt5: THValComboBox;
    FRupt6: THValComboBox;
    shSAV: TTabSheet;
    Bevel1: TBevel;
    TFTypeEtat: THLabel;
    TFNatEtat: THLabel;
    TFCodeEtat: THLabel;
    TFLibelle: THLabel;
    FTypeEtat: THCritMaskEdit;
    FNatEtat: THCritMaskEdit;
    FCOdeEtat: THCritMaskEdit;
    FLibelle: THCritMaskEdit;
    FMOdele: TCheckBox;
    BndTesteSQL: TCheckBox;
    TWListe: TToolWindow97;
    HPanel5: THPanel;
    BValideLst: TToolbarButton97;
    BCancelLst: TToolbarButton97;
    FindChamp: TFindDialog;
    GLst: THGrid;
    BRechLst: TToolbarButton97;
    BImprimeLst: TToolbarButton97;
    BPropCtlLst: TToolbarButton97;
    BPropietes: TToolbarButton97;
    BSQL: TToolbarButton97;
    bPreview: TToolbarButton97;
    bPrint: TToolbarButton97;
    BInsereChamp: TToolbarButton97;
    BInsereImage: TToolbarButton97;
    BChprop: TToolbarButton97;
    bList: TToolbarButton97;
    BSuprim: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    CtrCaption: TColorMemo;
    BElipsisiCaption: TBitBtn;
    TCtrlLarg: THLabel;
    CtrLarg: TSpinEdit;
    pnlPolice: TPanel;
    CtrBold: TCheckBox;
    CtrItalic: TCheckBox;
    CtrUnderlined: TCheckBox;
    CtrSize: TRadioGroup;
    ImExportFichier: TImage;
    Procedure ControlDblCLick(Sender : TObject ) ;
    procedure BCancelPnlClick(Sender: TObject);
    procedure BValiderPnlClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CtrCaptionElipsisClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure BSQLClick(Sender: TObject);
    procedure BNouveauClick(Sender: TObject);
    procedure BSuprimClick(Sender: TObject);
    Function  ChercheBandeParent : TLPBande ;
    procedure BInsereChampClick(Sender: TObject);
    procedure bEnregClick(Sender: TObject);
    procedure BEnregSousClick(Sender: TObject);
    procedure BLireClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BPropietesClick(Sender: TObject);
    procedure BValiderEtatClick(Sender: TObject);
    procedure BCancelEtatClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure RPoucesClick(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure bImportClick(Sender: TObject);
    procedure bExportClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BValideBndClick(Sender: TObject);
    procedure BBndSQLClick(Sender: TObject);
    procedure BCancelBndClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure bPreviewClick(Sender: TObject);
    procedure FRouleauClick(Sender: TObject);
    procedure BInsereImageClick(Sender: TObject);
    procedure BCancelImaClick(Sender: TObject);
    procedure BValiderImaClick(Sender: TObject);
    procedure ImaLogoClick(Sender: TObject);
    procedure ChkEnteteClick(Sender: TObject);
    procedure FormMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure CtrlMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bListClick(Sender: TObject);
    procedure BRechLstClick(Sender: TObject);
    procedure FindChampFind(Sender: TObject);
    procedure BImprimeLstClick(Sender: TObject);
    procedure BCancelLstClick(Sender: TObject);
    procedure BValideLstClick(Sender: TObject);
    procedure BPropCtlLstClick(Sender: TObject);
  private
    Menu97        : TMenu97 ;
    Pixel         : TLPPixel ;
    LPParam       : TLPParam ;
    Cont          : TLPContourne ;
    FChamps       : TFLPChamps ;
    FCtrl         : TControl ;
    LesBnds       : TList ;
    BandeCourant  : TLPBande ;
    LibNature     : String ;
    FindFirst     : Boolean ;
    ChangementVersion : Boolean ;
    BtnGros       : Boolean ; 
    procedure AdaptationsParTypeEtat ; //XMG 30/03/04
    Function  OnSauve : Boolean ;
    Function  CHercheNUmeroChamps : Integer ;
    Procedure PropCtr ( Ctrl : TControl ) ;
    Procedure PropietesChamp(Ctrl : TLPChamp) ;
    Procedure PropietesImage(Ctrl : TLPImage) ;
    Function  MajProp (Ctrl : TControl) : Boolean ;
    Function  MajPropietesImage ( Ctrl : TLPImage) : Boolean ;
    Function  MajPropietesChamp ( Ctrl : TLPChamp) : Boolean ;
    Procedure ChangeControl ( Sender : TObject) ;
    Function NomExist ( NomUser,NameCtrl : String ) : Boolean ;
    Procedure ResizeBtns ;
    Procedure UpdateContext ;
    Procedure Enregistrer ;
    Procedure Chargemodele ;
    Procedure AsigneEventsChamp(Ctrl : TControl ) ;
    Procedure SupprimeControl(Ctrl : TControl) ;
    Function ChargeTables (SQL : String ) : String ;
    Procedure ChargeChampsSurTree ;
    Function OkSQL ( St : String ) : Boolean ;
    Procedure MajChampSpecial( Activer : Boolean ) ;
    Procedure MAJChampsPolice (Activer : Boolean ) ;
    Function  Enregistremodeles : boolean ;
    Function  PrendNom (OkEnregistrer : Boolean ) : Boolean ;
    procedure EnableCnt ( tpBande : TLPTypeBandes ; Chk : TCheckBox ; Cnt : TSpinEdit ) ;
    Procedure NetoyerBandes(Nouveau : Boolean ) ;
    Procedure ChargeComboBande (Combo : THValCOmboBox) ;
    Procedure ChargeBandes ;
    Procedure CreateBande(TBande : TLPTypeBandes ; NumDetail : Integer ; Titre : String ) ;
    procedure BandeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    Function  SaisieSQL ( var SQL : String ; Tester : Boolean ) : Boolean ;
    Function  UnifierTables(Table1,Table2 : String ) : String ;
    Procedure AjouteTablesSubDetail ( SQL : String ) ;
    Function  DecomposeSQL( SQL,Sentence : String ; Var resultat : String ) : Boolean ;
    Procedure ChercheRuptures (SQL : String ) ;
    Function  FindControl(NomChamp : String) : TControl ;
  public
    FTables,
    TypEtat,
    NatEtat,
    CodETat   : String ;
    CreerEtat : Boolean ;
  end;


implementation
//XMG 30/03/04 Début
//Modifications à l'interface....
//  TWListe.Caption Changé de "Liste de champs de l'état" à "Liste des champs"
//  Nouveau panel "PnlPolice" dans CtrPanel qui regroupe les contrôles CtrBold, CtrItalic, CtrUnderlined et CtrSize...
//  Bproprietes.hint:='Propriétés du modèle'
//  BPreview.Hint:='Aperçu du modèle
//  imExportFichier du type TImage, sert pour à stocker le glyph du bouton BImprimer lors le paramétrage des Export ASCII
//XMG 30/03/04 fin
uses Uiutil, Hent1, Math, Sysutils, {graphics,} MC_Lib, HMsgBox, LP_Impri, MC_Erreur, //XMG 30/03/04 déplacée à l'utre Uses....
{$IFNDEF EXPORTASCII}
     MC_Admin,
{$ENDIF EXPORTASCII} //XMG 30/03/04
     SQLParts, Ed_Tools, Ed_Sql, Ed_nm, LP_View, PrintDBG ;

{$R *.DFM}
Procedure Param_LPTexte ( Inside : THpanel ; TypEtat,NatEtat,COdEtat : String ; CreerEtat : Boolean ) ;
var X : TFLPDessin ;
Begin
X:=TFLPDessin.Create(Application) ;
X.TypEtat:=TypEtat ;
X.NatEtat:=NatEtat ;
X.CodEtat:=CodEtat ;
X.CreerEtat:=CreerEtat ;
if Inside=nil then
   BEGIN
   try
      X.ShowModal ;
    Finally
      X.free ;
    End ;
   end else
   Begin
   initInside(X,Inside) ;
   X.Show ;
   End ;
SourisNormale ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function Compare ( item1,item2 : Pointer ) : integer ;
         /////////////////////////////////////////////////////////////////////////////////////////
         Function ValeurBande ( tb : TLPTypeBandes ) : Integer ;
         Begin
         Result:=ord(tb) ;
         inc(result,ord(tb in [lbdetail..lbsubpied])+ord(tb in [lbpied..lbSubpied])-(4*ord(tb=lbEnteteRupt))-(2*ord(tb=lbPiedRupt))) ;
         End ;
         /////////////////////////////////////////////////////////////////////////////////////////
Var Bnd1,Bnd2 : TLPBande ;
    Tb1,Tb2   : Integer ;
Begin
Result:=0 ;
Bnd1:=nil ; if TObject(Item1).classtype=TLPBande then Bnd1:=TLPBande(Item1) ;
Bnd2:=nil ; if TObject(Item2).classtype=TLPBande then Bnd2:=TLPBande(Item2) ;
If (Bnd1=nil) or (Bnd2=nil) then exit ;
Tb1:=ValeurBande(Bnd1.TypeBande) ;
Tb2:=ValeurBande(Bnd2.TypeBande) ;
Result:=Tb1-Tb2 ;
if (Result=0) and (Bnd1.Typebande in [lbSubEntete,lbSubDetail,lbSUbPied,lbENteteRupt,lbPiedRupt]) then Result:=Bnd1.NbrDetail-Bnd2.NbrDetail ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.FormResize(Sender: TObject);
begin
sbFonds.VertScrollBar.Position:=0 ;
sbFonds.HorzScrollBar.Position:=0 ;
LpBase.Left:=maxintValue([8,(LpBase.parent.ClientWidth-LPBase.Width-8) div 2]) ;
LpBase.Top:=maxintValue([8,(LpBase.parent.ClientHeight-LPBase.Height-8) div 2]) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.NomExist ( NomUser,NameCtrl : String ) : Boolean ;
var Nom  : String ;
    Ctrl : TControl ;
    i    : Integer ;
Begin
Result:=FALSE ;
for i:=0 to Componentcount-1 do
    if (Components[i] is TLPChamp) or
       (Components[i] is TLPImage)   then
       Begin
       Ctrl:=TControl(Components[i]) ;
       if (Ctrl is TLPChamp) then Nom:=TLpChamp(Ctrl).nom else
       if (Ctrl is TLPImage) then Nom:=TLpImage(Ctrl).Nom ;
       if (NomUser=Nom) and (NameCtrl<>Ctrl.Name) then Begin Result:=TRUE ; Break ; End ;
       End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.ChercheNumeroChamps : Integer ;
Var i,j,n : Integer ;
    Nom   : String ;
Begin
Result:=0 ;
for i:=0 to LPBase.controlcount-1 do
  if LPBase.Controls[i] is TWincontrol then
     for j:=0 to TWincontrol(LPBase.Controls[i]).ControlCount-1 do
         if (TWincontrol(LPBase.Controls[i]).Controls[j] is TLPChamp) or
            (TWincontrol(LPBase.Controls[i]).Controls[j] is TLPImage)     then
            Begin
            if (TWincontrol(LPBase.Controls[i]).Controls[j] is TLPChamp) then Nom:=TLPChamp(TWinControl(LPBase.controls[i]).controls[j]).nom
               else Nom:=TLPImage(TWinControl(LPBase.controls[i]).controls[j]).nom ;
            if Copy(Nom,1,1)='#' then Delete(Nom,1,1) ;
            if IsNumeric(Nom) then
               Begin
               n:=Valeuri(Nom) ;
               result:=maxintValue([Result,n]) ;
               End ;
            End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.ChangeControl ( Sender : TObject) ;
Var Ctrl       : TControl ;
    Fixe       : TLPSetFixe ; //XMG 30/03/04
Begin
Ctrl:=Nil ;
If FCtrl<>nil then
   Begin
   if (Assigned(Cont)) and (Cont.isAttached) then Cont.DeattachControl ;
   FCtrl.invalidate ;
   End ;
if (Sender is TLPChamp) or ((Sender is TLPImage) and (LPParam.TypeEtat<>TypeEtatExportAscII)) then //XMG 02/04/04 'F' 30/03/04
   Begin
   Ctrl:=TControl(Sender) ;
   if Ctrl.visible then
      Begin
      if not assigned(Cont) then Cont:=TLPContourne.Create(Self) ;
      Cont.onMouseWheel:=FormMouseWheel ;
      //XMG 30/03/04 début
      Fixe:=[] ;
      if (LPParam.TypeEtat=TypeEtatExportAscii) and (Sender is TLPChamp) then Fixe:=[lfHeight] ; //XMG 02/04/04 'F'
      Cont.AttachControl(Ctrl,Fixe) ;
      //XMG 30/03/04 fin
      End ;
   End ;
FCtrl:=Ctrl ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Procedure TFLPDessin.ControlDblCLick(Sender : TObject ) ;
Var Ctrl : TControl ;
Begin
if ((FCtrl=nil) and (Not (Sender is TLPChamp)) and (Not (Sender is TLPImage)) and (Not (Sender is TLPContourne))) then exit ;
EnableControls(LPBASE,FALSE)  ;
if (Sender is TLPContourne) then Ctrl:=TLPContourne(Sender).Control else
if (Sender is TLPChamp) or (Sender is TLPImage) then Ctrl:=TControl(Sender) else Ctrl:=FCtrl ;
try
  PropCtr(Ctrl) ;
 except
  EnableControls(LPBASE,TRUE)  ;
  Raise ;
 End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.PropCtr ( Ctrl : TControl ) ;
Begin
ChangeControl(Ctrl) ;
if Ctrl is TLPChamp then
   Begin
   PropietesChamp(TLPChamp(Ctrl)) ;
   CtrPanel.Visible:=TRUE ;
   CtrNom.SetFocus ;
   End else
if Ctrl is TLPImage then
   Begin
   PropietesImage(TLPImage(Ctrl)) ;
   TWImage.Visible:=TRUE ;
   ImaNom.SetFocus ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.PropietesChamp(Ctrl : TLPChamp) ;
Var i  : integer ;
    St : String ;
Begin
CtrNom.Text           :=Ctrl.Nom ;
CtrX.Value            :=Ctrl.Ligne ;
CtrY.Value            :=Ctrl.Colonne ;
CtrLong.Value         :=Ctrl.NbrChars ;
CtrLarg.Value         :=Ctrl.NbrLignes ;
CtrVisible.checked    :=Ctrl.PrinterVisible ;
CtrCaption.Text       :=CtrL.Libelle ;
LP_ChargeMask(CtrMask,Ctrl) ;
St:=inttostr(ord(TLPBande(Ctrl.Parent).TypeBande)) ;
if TLPBande(Ctrl.Parent).TypeBande in [lbSubEntete,lbSubDetail,lbSubPied,lbEnteteRupt,lbPiedRupt] then St:=St+inttostr(TLPBande(Ctrl.Parent).NbrDetail) ;
CtrlBande.Value       :=St ;
CtrBold.Checked       :=(fsBold in Ctrl.Font.Style) ;
CtrItalic.Checked     :=(fsItalic in Ctrl.Font.Style) ;
CtrUnderLIned.Checked :=(fsUnderLine in Ctrl.Font.Style) ;
case Ctrl.AlignTexte of
  taLeftJustify   : CtrAlign.ItemIndex:=0 ;
  taCenter        : CtrAlign.ItemIndex:=1 ;
  taRightJustify  : CtrAlign.ItemIndex:=2 ;
  end ;
CtrSize.itemindex:=ord(Ctrl.Taille) ;
CtrSiZero.Value:=inttostr(ord(Ctrl.SiZero)) ;
St:=CtrPanel.Caption ;
i:=pos(':',St) ;
if i>0 then St:=Copy(CtrPanel.Caption,1,i)+' '+CtrNom.Text ; ;
CtrPanel.Caption:=St ;
MajChampSpecial(isSpecialChamp(Ctrl)) ;
MajChampsPolice(isSpecialChamp(Ctrl) or Ctrl.IsCodeBarre) ;
CentreTW(Self,CtrPanel) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.PropietesImage(Ctrl : TLPImage) ;
Var st : String ;
    i  : Integer ;
Begin
ImaNom.Text           :=Ctrl.Nom ;
ImaX.Value            :=Ctrl.Ligne ;
ImaY.Value            :=Ctrl.Colonne ;
ImaLong.Value         :=Ctrl.Longeur ;
ImaLarg.Value         :=Ctrl.Largeur ;
ImaVisible.checked    :=Ctrl.PrinterVisible ;
ImaLogo.State         :=cbUnchecked ;
St:=Ctrl.Fichier ;
if St='**LOGO**' then
   Begin
   St:='' ;
   ImaLogo.State:=CbChecked ;
   End ;
ImaContenu.Text       :=St ;
ImaLogoClick(nil) ;
St:=inttostr(ord(TLPBande(Ctrl.Parent).TypeBande)) ;
if TLPBande(Ctrl.Parent).TypeBande in [lbSubEntete,lbSubDetail,lbSubPied,lbEnteteRupt,lbPiedRupt] then St:=St+inttostr(TLPBande(Ctrl.Parent).NbrDetail) ;  
ImaBande.Value        :=St ;
St:=TWImage.Caption ;
i:=pos(':',St) ;
if i>0 then St:=Copy(TWImage.Caption,1,i)+' '+ImaNom.Text ; ;
TWImage.Caption:=St ;
CentreTW(Self,TWImage) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BCancelPnlClick(Sender: TObject);
begin
CtrPanel.Visible:=FALSE ;
EnableControls(LPBASE,TRUE)  ;
if Assigned(FCtrl) then
   Begin
   if (FCtrl is TLPChamp) and (Not FCtrl.Visible) then supprimeControl(FCtrl)
      else ChangeControl(FCtrl) ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.MajProp (Ctrl : TControl ) : Boolean ;
Begin
Result:=FALSE ;
if Ctrl=nil then Exit ;
if Ctrl is TLPChamp then Result:=MajPropietesChamp(TLPChamp(Ctrl)) else
if Ctrl is TLPImage then Result:=MajPropietesImage(TLPImage(Ctrl)) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.MajPropietesImage ( Ctrl : TLPImage) : Boolean ;
Var idx : Integer ;
    St  : String ;
Begin
REsult:=FALSE ;
Idx:=ImaBande.values.indexof(ImaBande.Value) ;
Ctrl.Parent:=TLPBande(ImaBande.Values.Objects[idx]) ;
if NomExist(ImaNom.Text,Ctrl.Name) then Begin PGIBox(MC_MsgErrDefaut(1060),Caption) ; Exit ; End ;
Ctrl.Nom            :=ImaNom.Text ;
Ctrl.Ligne          :=imaX.Value ;
Ctrl.Colonne        :=imaY.Value ;
Ctrl.Longeur        :=ImaLong.Value ;
Ctrl.Largeur        :=ImaLarg.Value ;
Ctrl.PrinterVisible :=ImaVisible.checked ;
St:=Trim(ImaContenu.Text) ;
if ImaLogo.State=cbChecked then St:='**LOGO**' ;
Ctrl.Fichier        :=St ;
Ctrl.visible:=TRUE ;
Result:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.MajPropietesChamp( Ctrl : TLPChamp) : Boolean ;
Var idx : Integer ;
Begin
REsult:=FALSE ;
Idx:=CtrlBande.values.indexof(CtrlBande.Value) ;
Ctrl.Parent:=TLPBande(CtrlBande.Values.Objects[idx]) ;
if NomExist(CtrNom.Text,Ctrl.Name) then Begin PGIBox(MC_MsgErrDefaut(1060),Caption) ; Exit ; End ;
Ctrl.Nom            :=CtrNom.Text ;
Ctrl.Ligne          :=CtrX.Value ;
Ctrl.Colonne        :=CtrY.Value ;
Ctrl.NbrChars       :=CtrLong.Value ;
Ctrl.NbrLignes      :=CtrLarg.Value ;
Ctrl.PrinterVisible :=CtrVisible.checked ;
Ctrl.Libelle        :=Trim(CtrCaption.Text) ;
Ctrl.Mask           :=CtrMask.Text ;
Ctrl.Font.Style:=[] ;
if CtrBold.Checked       then Ctrl.Font.Style:=Ctrl.Font.Style+[fsBold] ;
if CtrItalic.Checked     then Ctrl.Font.Style:=Ctrl.Font.Style+[fsItalic] ;
if CtrUnderLined.Checked then Ctrl.Font.Style:=Ctrl.Font.Style+[fsUnderLine] ;
Case CtrAlign.ItemIndex of
  0 : Ctrl.AlignTexte:=taLeftJustify ;
  1 : Ctrl.AlignTexte:=taCenter ;
  2 : Ctrl.AlignTexte:=taRightJustify ;
  End ;
Ctrl.SiZero:=TLPSiValeurZero(valeuri(CtrSiZero.value)) ;
Ctrl.Taille:=TLPTailleChar(CtrSize.ItemIndex) ;
Ctrl.visible:=TRUE ;
MajChampSpecial(isSpecialChamp(Ctrl)) ;
MajChampsPolice(isSpecialChamp(Ctrl) or Ctrl.IsCodeBarre) ;
Result:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TFLPDessin.BValiderPnlClick(Sender: TObject);
Var Ok : Boolean ;
begin
Ok:=MajProp(FCtrl) ;
EnableControls(LPBASE,TRUE)  ;
if Ok then
   Begin
   CtrPanel.Visible:=FALSE ;
   ChangeControl(FCtrl) ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.FormCreate(Sender: TObject);
begin
BandeCourant:=Nil ;
Menu97:=Nil ;
LPParam:=TLPParam.create(Self) ;
Cont:=TLPContourne.Create(Self) ;
FChamps:=TFLPChamps.Create(Self) ;
FChamps.Visible:=FALSE ;
Pixel:=TLPPixel.Create(Self) ;
LesBnds:=TList.Create ;
CntEntete.MaxValue:=NbrMaxSubBnds ;
CntDetail.MaxValue:=NbrMaxSubBnds ;
CntPied.MaxValue:=NbrMaxSubBnds ;
BtnGros:=FALSE ; 
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.MajChampSpecial( Activer : Boolean ) ;
Begin
CtrLong.Enabled       :=Not Activer ; if Activer and (CtrLong.Value<>1) then CtrLong.value:=1 ;
CtrLarg.Enabled       :=Not Activer ; if Activer and (CtrLarg.Value<>1)  then CtrLarg.value:=1 ;
CtrMask.Enabled       :=Not Activer ; if Activer then Begin CtrMask.ItemIndex:=-1 ; CtrMask.Values.Clear ; CtrMask.Items.Clear ; End ;
CtrVisible.Enabled    :=not Activer ; If (Activer) and (CtrVisible.Checked) then CtrVisible.Checked:=FALSE ;
CtrSiZero.Enabled     :=not Activer ; If (Activer) and (CtrSiZero.ItemIndex<>0) then CtrSiZero.ItemIndex:=0 ;
end;
//////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.MAJChampsPolice (Activer : Boolean ) ;
Begin
CtrBold.Enabled       :=Not Activer ; If (Activer) and (CtrBold.Checked) then CtrBold.Checked:=FALSE ;
CtrItalic.Enabled     :=Not Activer ; If (Activer) and (CtrItalic.Checked) then CtrItalic.Checked:=FALSE ;
CtrUnderLIned.Enabled :=Not Activer ; If (Activer) and (CtrUnderLined.Checked) then CtrUnderLined.Checked:=FALSE ;
CtrAlign.Enabled      :=Not Activer ; If (Activer) and (CtrAlign.ItemIndex<>0) then CtrAlign.ItemIndex:=0 ;
CtrSize.Enabled       :=Not Activer ; If (Activer) and (CtrSize.itemindex<>2) then CtrSize.itemindex:=2 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.CtrCaptionElipsisClick(Sender: TObject);
begin
if FChamps.TreeCH.items.count<=0 then ChargeChampsSurTree ;
FChamps.Champ.Text:=CtrCaption.Text ;
if Fchamps.Showmodal=mrOk then CtrCaption.Text:='['+FChamps.Champ.Text+']' ;
FChamps.Champ.Text:='' ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.FormClose(Sender: TObject; var Action: TCloseAction);
begin
If isInside(Self) then Action:=caFree ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Procedure TFLPDessin.Enregistrer ;
Begin
V_PGI.IOError:=OeSaisie ;
With LPParam do
   Begin
   Personnel  :=FALSE ;
   Defaut     :=FALSE ;
   Protect    :=FALSE ;
   Suivant    :='' ;
   ProtectSQL :=FALSE ;
   User       :=V_PGI.User ;
   DiagText   :='' ;
   ForceDT    :=FALSE ;
   Menu       :=FALSE ;
   Exporter   :=FALSE ;
   End ;
if EnregLP(LPParam,LpBase) then V_PGI.IOError:=oeOk ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.UpdateContext ;
Begin
{$IFNDEF EXPORTASCII}
BPrint.Enabled:=((LPParam.TypeEtat=TypeEtatExportAscii) or (UsesTPV(V_MC.GetDispositif(mcPrinter),mcPrinter))) and (LPParam.CodeEtat<>'') ;
{$ELSE}
BPrint.Enabled:=(LPParam.CodeEtat<>'') ;
{$ENDIF EXPORTASCII} //XMG 02/04/04
BNouveau.Enabled:=CreerEtat ;
BImport.Enabled:=BNouveau.Enabled ;
BEnreg.enabled:=(not LPParam.Modele) or (V_PGI.SAV) or (V_PGI.SUperviseur) ;
BEnregSous.Enabled:=(BEnreg.Enabled) and (BNouveau.Enabled) ;
//------ MAJ MENU ------------
Nouveau1.Enabled:=BNouveau.Enabled ;
ImporterModle1.Enabled:=BImport.Enabled ;
Enregistre1.Enabled:=BEnreg.Enabled ;
EnregistreSous1.Enabled:=BEnregSous.Enabled ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//XMG 30/03/04 début
procedure TFLPDessin.AdaptationsParTypeEtat ;
Begin
{if TypEtat = TypeEtatTexte then    //Modèle Texte  //XMG 02/04/04
  Begin
  End else}
if TypEtat = TypeEtatExportAscII then  //Fichier Exportation  //XMG 02/04/04 'F'
  Begin
  //--->Général<---
  Caption:=TraduireMemoire('Paramétrage d''une exportation texte:') ;
  bPrint.Hint:=traduirememoire('Génération fichier') ;
  bprint.Glyph.Assign(imExportFichier.Picture) ;
  //--->Images<----
  BInsereImage.Visible:=FALSE ;
  //--->Propriétés Control<---
  PnlPolice.Visible:=FALSE ;
  CtrPanel.Height:=CtrPanel.Height-PnlPolice.Height ;
  CtrAlign.top:=PnlPolice.Top ;
  CtrLarg.Visible:=FALSE ;
  TCtrlLarg.Visible:=FALSE ;
  //--->Propriétés "état"<---
  EtatPanel.Caption:=TraduireMemoire('Propriétés de l''exportation') ;
  // Papier
  shPapier.Caption:=TraduireMemoire('Général') ;
  FRouleau.Checked:=TRUE ;
  FRouleau.Visible:=FALSE ;
  //XMG 15/04/04 début
  Largeur.Masks.PositiveMask:='#,##0' ;
  TailleEn.Visible:=FALSE ;
  //XMG 15/04/04 fin

  // Marges
  tsMarges.TabVisible:=FALSE ;
  End ;
End ;
//XMG 30/03/04 fin
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.ResizeBtns ;
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
End ;
///////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.FormShow(Sender: TObject);
begin
if V_PGI.ZoomForm then ResizeBtns ;
TopDock.AllowDrag:=not IsInside(Self) ;
LeftDock.AllowDrag:=not IsInside(Self) ;
RightDock.AllowDrag:=not IsInside(Self) ;
BottomDock.AllowDrag:=not IsInside(Self) ;
Timer.Enabled:=not IsInside(Self) ;
PBottom.Visible:=Not IsInside(Self) ;
{if Timer.Enabled  then
   Begin
   Menu97:=TMenu97.Create(Self) ;
   with Menu97 do
     Begin
     Parent:=Self ;
     MainMenu:=MenuMain ;
     Caption:=TraduireMemoire(MC_MsgErrDefaut(1036)) ;
     DockedTo:=TopDock ;
     DockRow:=0 ;
     DockPos:=0 ;
     Rebuild ;
     End ;
   End ;  } //XMG non pour l'instant
AdaptationsParTypeEtat ; //XMG 30/03/04
LibNature:=RechDom('YYNATUREETAT',NatEtat,FALSE) ; //XMG 28/08/01
initCaption(Self,LibNature,'') ;
LPParam.TypeEtat :=TypEtat ;
LPParam.NatEtat  :=NatEtat ;
LPParam.CodeEtat :=CodEtat ;
ChargeModele ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.Chargemodele ;
Var Err      : TMC_Err ;
Begin
FillChar(Err,sizeof(TMC_Err),#0) ;
LPParam.Langue   :=V_PGI.LanguePerso ;
if (Not ChargeModeleLP(LPParam,CreerEtat,LpBase,AsigneEventsChamp,nil,Err,ChangementVersion)) and (Err.Code<>0) then PGIBox(Err.Libelle,Caption) ;
InitCaption(Self,LibNature,LPParam.Libelle) ; //XMG 20/04/04
ChargeChampssurTree ;
ChercheRuptures(LPParam.SQL) ;
ChargeBandes ;
ChargeComboBande(CtrlBande) ;
ChargeComboBande(ImaBande) ;
FormResize(nil) ;
if (LPParam.Modele) and (PGIAsk(MC_MsgErrDefaut(1072),Caption)=mrYes) then
   Begin
   LPParam.Modele:=FALSE ;
   LPParam.CodeEtat:='' ;
   LPParam.Nouvelle:=TRUE ;
   End ;
LpBase.Modified:=LPParam.Nouvelle ;
BPreview.Enabled:=(LPParam.CodeEtat<>'') ;
LpBase.SetFocus ;
UpdateContext ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.DecomposeSQL( SQL,Sentence : String ; Var resultat : String ) : Boolean ;
Var TN       : TTreeView ;
    N1,N2,N3 : TTreeNode ;
    St       : String ;
Begin
Result:=FALSE ;
Sentence:=Uppercase(Trim(Sentence)) ;
resultat:='' ;
TN:=TTreeview.Create(self) ;
try
   TN.Parent:=Self ;
   TN.Visible:=FALSE ;
   Resultat:='' ;
   if RecupSQLParts(SQL,TN.items) then
      Begin
      N1:=CherchePart(TN.Items,Sentence) ;
      if assigned(N1) then
        Begin
        Result:=TRUE ;
        N2:=N1.getFirstChild ;
        while assigned(N2) do
          Begin
          if (not N2.HasChildren) or (N2.Index=0) then
             Begin
             St:=N2.Text+';' ;
             while (Length(St)>1) and (St[1]='(') do Delete(St,1,1) ;
             Resultat:=Resultat+St ;
             End ;
          if N2.HasChildren then
             Begin
             N3:=N2.GetFirstChild ;
             if Assigned(N3) then Resultat:=Resultat+N3.Text+';' ;
             End ;
          N2:=N1.GetNextChild(N2) ;
          End ;
        End ;
      End ;
  Finally
    Videtree(TN) ;
    TN.Items.Clear ;
    TN.Free ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.ChargeTables (SQL : String ) : String ;
Var St,St2,St3 : String ;
    QQ         : TQuery ;
Begin
Result:='' ;
DecomposeSQL(SQL,'FROM',Result) ;
St:=Result ;
Result:='' ;
while St<>'' do
  Begin
  St2:=ReadTokenSt(St) ;
  if TableToPrefixe(St2)='' then
     Begin
     QQ:=OpenSQL('select DV_SQL from DEVUES where DV_NOMVUE="'+St2+'"', True) ;
     if not QQ.EOF then
        Begin
        DecomposeSQL(QQ.FindField('DV_SQL').AsString,'FROM',St3) ;
        if St3<>'' then Result:=Result+St3 ;
        End ;
     Ferme(QQ) ;
     End else Result:=Result+St2+';' ;
  End ;
//XMG 02/04/04 début
Result:=Result+'PARAMSOC;' ;
if LPParam.TypeEtat=TypeEtatTexte then Result:=Result+'LPRINTER;' ;
//XMG 02/04/04 fin
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.ChercheRuptures (SQL : String ) ;
Var ChRupture,Ch,v : String ;
    it,va          : TStrings ;
    i              : Integer ;
    CB             : THValComboBox ;
Begin
if not DecomposeSQL(SQL,'GROUP BY',ChRupture) then DecomposeSQL(SQL,'ORDER BY',ChRupture) ;
if trim(ChRupture)<>'' then
   Begin
   if Copy(ChRupture,length(ChRupture),1)<>';' then ChRupture:=ChRupture+';' ;
   LPBase.ChRupture:=ChRupture ;
   It:=TStringList.Create ;
   Va:=TStringList.Create ;
   try
     It.Add(Traduirememoire(MC_MsgErrDefaut(1073))) ; va.Add(' ') ;
     For i:=1 to NbCarInString(ChRupture,';') do
       Begin
       Ch:=gtfs(ChRupture,';',i) ;
       if it.IndexOf(Ch)<0 then
          Begin
          It.Add(Ch) ;
          Va.Add(Ch) ;
          End ;
       End ;
     For i:=1 to LP_NbMaxRuptures do
       Begin
       CB:=THValComboBox(FindComponent('FRupt'+inttostr(i))) ;
       if Assigned(CB) then
          Begin
          v:=CB.Value ;
          CB.Items.Assign(It) ;
          CB.Values.assign(Va) ;
          CB.Value:=V ;
          End ;
       End ;
    finally
     It.Free ;
     va.Free ;
    End ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLpDessin.ChargeChampsSurTree ;
Var OldTables : String ;
    i         : Integer ;
    j         : TLPTypeBandes ;
Begin
OldTables:=FTables ;
FTables:=ChargeTables(LPParam.SQL) ;
for j:=low(TLPTypeBandes) to High(TLPTypeBandes) do
    if j in [lbSubEntete,lbSubDetail,lbSubPied,lbEnteteRupt,lbPiedRupt] then  
       For i:=1 to LPBase.numMaxSubDetail(j) do
           if (Assigned(LpBase.choixBande(j,i))) and (trim(LpBase.choixBande(j,i).SQL)<>'') then AjouteTablesSubDetail(LpBase.choixBande(j,i).SQL) ;
if OldTables=FTables then exit ;
VideTree(FChamps.TreeCH) ;
FChamps.TreeCH.Items.Clear ;
ChargeChampsTree(FTables,FChamps.TreeCH) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.AsigneEventsChamp(Ctrl : TControl ) ;
Begin
if (Not(Ctrl is TLPChamp)) and (Not(Ctrl is TLPImage)) then exit ;
if Ctrl is TLPChamp then
   with Ctrl as TLPChamp do
     Begin
     OnClick:=ChangeControl ;
     OnDblCLick:=ControlDblCLick ;
     onenter:=ChangeControl ;
     OnMouseDown:=CtrlMouseDown ;
     End else
if Ctrl is TLPImage then
   with Ctrl as TLPImage do
     Begin
     OnClick:=ChangeControl ;
     OnDblCLick:=ControlDblCLick ;
     onenter:=CHANGECONTROL ;
     OnMouseDown:=CtrlMouseDown ;
     End else ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BSQLClick(Sender: TObject);
Var SQL : String ;
Begin
SQL:=LPParam.SQL ;
if SaisieSQL(SQL,TRUE) then  
   Begin
   LPParam.SQL:=SQL ;
   ChargeChampsSurTree ;
   ChercheRuptures(SQL) ;  
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.SaisieSQL ( var SQL : String ; Tester : Boolean ) : Boolean ;  
Var xx    : TReqSQL ;
begin
Result:=FALSE ;
XX:=TReqSQL.Create(Application) ;
try
  {$IFNDEF AGL570d}
  XX.Q2.SQL.clear ;
  XX.Q2.SQL.Add(SQL) ;
  {$ELSE}
  XX.Q2.clear ;
  XX.Q2.Add(SQL) ;
  {$ENDIF}
  if XX.ShowModal=mrOk then
     Begin
     {$IFNDEF AGL570d}
     if (trim(XX.Q2.SQL.Text)='') or (not Tester) or (OKSQL(XX.Q2.SQL.Text))  then
     {$ELSE}
     if (trim(XX.Q2.Text)='') or (not Tester) or (OKSQL(XX.Q2.Text))  then
     {$ENDIF}
        Begin
        LpBase.Modified:=TRUE ;
        Result:=TRUE ;
        {$IFNDEF AGL570d}
        SQL:=XX.Q2.SQL.Text ;
        {$ELSE}
        SQL:=XX.Q2.Text ;
        {$ENDIF}
        End ;
     End ;
 finally
  XX.Free
 End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.OkSQL ( St : String ) : Boolean ;
Var Err : TMC_Err ;
    Q2 : TQuery ;
Begin
err.code:=0 ; err.Libelle:='' ;
Result:=FALSE ;
try
   Q2:=PrepareSQL(St,TRUE) ;
   Result:=TRUE ;
   Ferme(Q2) ;
 except
    on E:EDatabaseError do Begin Err.Code:=-1 ; Err.Libelle:=E.Message ; End ;
    on E:Exception do Begin Err.Code:=-1 ; Err.Libelle:=E.Message ; End ;
 End ;
if Not Result then
   Begin
   Err.Libelle:=Err.Libelle+MC_MsgErrDefaut(1001) ;
   PGIBox(Err.Libelle,Caption) ;
   End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BNouveauClick(Sender: TObject);
Begin
NetoyerBandes(TRUE) ;
UpdateContext ; 
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.NetoyerBandes(Nouveau : Boolean ) ;
begin
if assigned(FCtrl) then ChangeControl(nil) ;
If Assigned(Cont) then Begin Cont.free ; Cont:=Nil ; End ;
While LPBase.ControlCount>0 do LPBase.Controls[LpBase.ControlCount-1].free ;
LPBase.ChRupture:='' ;
if Nouveau then
  Begin
  LPParam.CodeEtat:='' ;
  LPParam.Libelle:='' ;
  LPParam.Nouvelle:=TRUE ;
  ChargeModele ;
  End ;
ChangementVersion:=TRUE ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BSuprimClick(Sender: TObject);
begin
If Assigned(Fctrl) then SupprimeControl(FCtrl) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.SupprimeControl(Ctrl : TControl) ;
Begin
if (assigned(FCtrl)) and (Ctrl=FCtrl) then ChangeControl(nil) ;
Ctrl.Free ;
LpBase.Modified:=TRUE ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.ChercheBandeParent : TLPBande ;
Var i : Integer ;
Begin
i:=0 ; Result:=nil ;
while (i<LPBase.ControlCount) and (Result=Nil) do
  Begin
  if LPBase.Controls[i] is TLPBande then Result:=TLPBande(LPBase.Controls[i]) ;
  inc(i) ;
  End ;
if not assigned(Result) then PGIBox(MC_MsgErrDefaut(1059),caption) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BInsereChampClick(Sender: TObject);
Var p     : TPoint ;
    Num   : Integer ;
    Bande : TLPBande ;
    Ctrl  : TLPChamp ;
begin
p:=Point(1,1) ;
Bande:=ChercheBandeParent ;
if Assigned(Bande) then
   Begin
   Num:=ChercheNumeroChamps+1 ;
   Ctrl:=TLpChamp.Create(Self) ;
   Ctrl.Visible:=FALSE ;
   Ctrl.Parent:=Bande ;
   Ctrl.Nom:='#'+inttostr(Num) ;
   Ctrl.Libelle:=Ctrl.Nom ;
   Ctrl.Top:=P.Y ;
   Ctrl.Left:=P.X ;
   Ctrl.NbrChars:=10 ;
   Ctrl.TabStop:=TRUE ;
   Ctrl.TabOrder:=Ctrl.Parent.ControlCount-1 ;
   AsigneEventsChamp(Ctrl) ;
   ControlDBlCLick(Ctrl) ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TFLPDessin.bEnregClick(Sender: TObject);
Begin
Enregistremodeles ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.PrendNom (OkEnregistrer : Boolean ) : Boolean ;
Begin
With TNomModele.Create(Self) do
  Try
      {$IFNDEF AGL570d}
      TipNature.Caption:=LPParam.TypeEtat+LPParam.NatEtat ;
      QModele.SQL.Clear ;
      QModele.SQL.Add('Select MO_CODE, MO_LIBELLE, MO_MODELE from MODELES WHERE MO_TYPE="'+LPParam.TypeEtat+'" and MO_NATURE="'+LPParam.NatEtat+'"') ;
      {$ELSE}
      LTip.Caption:=LPParam.TypeEtat ;
      LNat.Caption:=LPParam.NatEtat ;
      QModeleSQL:='Select MO_CODE, MO_LIBELLE, MO_MODELE from MODELES WHERE MO_TYPE="'+LPParam.TypeEtat+'" and MO_NATURE="'+LPParam.NatEtat+'"' ;
      {$ENDIF}
      ModeleActif.caption:=LPParam.CodeEtat ;
      ModeleActif.Hint:=LibNature ;
      Creation.Checked:=OkEnregistrer ;
      NModele.Text:=LPParam.CodeEtat ;
      NLibelle.Text:=LPParam.Libelle ;
      Result:=(showModal=mrOk) ;
      if Result then
          Begin
          LPParam.CodeEtat:=NModele.Text ;
          if OkEnregistrer then LPParam.Libelle:=NLibelle.Text ;
          End ;
    Finally
      Free ;
    End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Function TFLPDessin.Enregistremodeles : Boolean ;
Var Err : TMC_Err ;
begin
Err.Code:=0 ; Err.Libelle:='' ;
if ChangementVersion and (not LPParam.Nouvelle) then
   Err.Code:=1083*ord(PGIAsk(MC_MsgErrDefaut(1082),Caption)<>mrYes) ;
if Err.Code=0 then
   Begin
   if LPBase.Modified or LPParam.Nouvelle Or ChangementVersion then
      Begin
      if (trim(LPParam.CodeEtat)='') or (LPParam.Nouvelle) then
         if not PrendNom(TRUE) Then Err.Code:=-1 ;
      if Err.Code=0 then
         Begin
         if transactions(Enregistrer,3)<>oeOK then Err.Code:=1 ;
         LPBase.Modified:=FALSE ;
         UpdateCOntext ;
         End ;
      End ;
   End ;
if Err.code>0 then
   Begin
   if trim(Err.Libelle)='' then Err.Libelle:=MC_MsgErrDefaut(Err.Code) ;
   PGIBox(Err.Libelle,Caption) ;
   End ;
Result:=(Err.Code=0) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BEnregSousClick(Sender: TObject);
begin
LPParam.Nouvelle:=TRUE ;
LPParam.CodeEtat:='' ;
LPParam.Libelle:='' ;
Enregistremodeles ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BLireClick(Sender: TObject);
begin
if (onsauve) and (PrendNom (FALSE)) then
   Begin
   NetoyerBandes(FALSE) ;
   ChargeModele ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
CanClose:=OnSAuve ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.EnableCnt ( tpBande : TLPTypeBandes ; Chk : TCheckBox ; Cnt : TSpinEdit ) ;
Begin
Cnt.Enabled:=Chk.Checked ;
Cnt.Value:=1 ;
if Cnt.Enabled then Cnt.Value:=LPBase.NumMaxSubDetail(tpBande)+1 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BPropietesClick(Sender: TObject);
var i    : integer ;
    rupt : String ;
    CB   : THValComboBox ;
begin
//XMG 15/04/04 début
if TypEtat = TypeEtatExportAscII then
   Begin
   pixel.AsPixel:=LPBase.Width ;
   Largeur.Value:=Pixel.AsChar[FALSE,LPTailleDef] ;
   End else
   Begin
   TailleEn.ItemIndex:=Ord(not LPBase.EnPouces) ; RPoucesclick(Nil) ;

   FRouleau.Checked:=LPBase.Rouleau ; FRouleauClick(nil) ;

   if TailleEn.ItemIndex=0 then
      Begin
      Pixel.AsPixel:=LpBase.Width ;
      Largeur.Value:=Pixel.AsInch ;
      Pixel.AsPixel:=LpBase.Height ;
      Longeur.Value:=Pixel.AsInch ;  //XMG cette ligne manquée (?)
      end else
      Begin
      Pixel.AsPixel:=LpBase.Width ;
      Largeur.Value:=Pixel.AsCM ;
      Pixel.AsPixel:=LpBase.Height ;
      Longeur.Value:=Pixel.AsCM ;
      end ;
   End ;
//XMG 15/04/04 fin
Pixel.AsPixel:=LPBase.MargeLeft ;
MargeG.Value:=Pixel.AsChar[FALSE,LPTailleDef] ;

Pixel.AsPixel:=LpBase.MargeTop ;
MargeH.Value:=Pixel.asChar[TRUE,LPTailleDef] ;

Pixel.AsPixel:=LpBase.MargeRight ;
MargeD.Value:=Pixel.AsChar[FALSE,LPTailleDef] ;

Pixel.AsPixel:=LPBase.MargeBottom ;
MargeB.Value:=Pixel.AsChar[TRUE,LPTailleDef] ;

GRMessures.itemIndex:=ord(LpBase.Units) ;
ChkEntete.Checked:=(LPBase.ChoixBande(lbEntete)<>nil) ;
EnableCnt(lbSubEntete,ChkEntete,cntEntete) ;

ChkDetail.Checked:=(LPBase.ChoixBande(lbDetail)<>nil) ;
EnableCnt(lbSubDetail,ChkDetail,cntDetail) ;

ChkPied.Checked:=(LPBase.ChoixBande(lbPied)<>nil) ;
EnableCnt(lbSubPied,ChkPied,cntPied) ;

CentreTW(Self,EtatPanel) ;
pgGeneral.ActivePage:=shPapier ;
shSAV.TabVisible:=(V_PGI.SAV) or (V_PGI.Superviseur) ;
FTypeEtat.Text:=LPParam.TypeEtat ;
FNatEtat.Text:=LPParam.NatEtat ;
FCodeEtat.Text:=LPParam.CodeEtat ;
FLibelle.Text:=LPParam.Libelle ;
FModele.Checked:=LPParam.Modele ;

Rupt:=LPBase.Ruptures ;
if (Trim(Rupt)<>'') and (Copy(Rupt,length(Rupt),1)<>';') then Rupt:=Rupt+';' ;
i:=0 ;
while (i<=LP_NbMaxRuptures) do
  Begin
  inc(i) ;
  CB:=THValComboBox(FindComponent('FRupt'+inttostr(i))) ;
  if assigned(CB) then
     Begin
     CB.ItemIndex:=0 ;
     if Trim(Rupt)<>'' then
        Begin
        CB.Value:=ReadtokenSt(Rupt) ;
        End ;
     End ;
  End ;
TsRupt.TabVisible:=Trim(LPBase.ChRupture)<>'' ;

EnableControls(LPBASE,FALSE)  ;
EtatPanel.Visible:=Not EtatPanel.Visible ;
Largeur.SetFocus ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.CreateBande(TBande : TLPTypeBandes ; NumDetail : Integer ; Titre : String ) ;
Var Bnd : TLPBande ;
Begin
Bnd:=TLPBande.Create(Self) ;
Bnd.Parent:=LpBase ;
If NumDetail<0 then NumDetail:=LPBase.NumMaxSubdetail(TBande) ;  
Bnd.Name:='Bande'+inttostr(ord(TBande))+inttostr(NumDetail) ;
Bnd.height:=Pixel.TailleChar(TRUE,LPTailleDef)*5 {Lignes} ;
Bnd.TypeBande:=TBande ;
if TBande in [lbSubEntete,lbSubDetail,lbSubPIed,lbEnteteRupt,lbPiedRupt] then Bnd.NbrDetail:=NumDetail ;  
Bnd.Titre:=Titre ;
Bnd.OnMouseDown:=BandeMouseDown ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BValiderEtatClick(Sender: TObject);
Var j,tb            : TLPTypeBandes ;
    TopBnd,ND,Cnt,i : Integer ;
    Ok              : Boolean ;
    Titre,Rupt,St   : String ;
    CB              : THValComboBox ;
    Bnd             : TLPBande ;
begin
NextControl(Self,TRUE) ;
//XMG  début
//XMG 15/04/04 début
if TypEtat = TypeEtatExportAscII then
   Begin
   Pixel.AsChar[FALSE,LPTailleDef]:= minintvalue([1000,round(Largeur.Value)]) ;
   LPBase.Width:=pixel.AsPixel ;
   End else
   Begin
   LPBase.EnPouces:=(TailleEn.ItemIndex=0) ;
   LpBase.Rouleau:=(FRouleau.state=cbChecked) ;
   if LpBase.EnPouces then
      Begin
      Pixel.AsInch:=Largeur.Value ;
      LpBase.Width:=Pixel.AsPixel ;
      Pixel.AsInch:=Longeur.Value ;
      LpBase.Height:=Pixel.AsPixel ;
      end else
      Begin
      Pixel.AsCM:=Largeur.Value ;
      LpBase.Width:=Pixel.AsPixel ;
      Pixel.AsCM:=Longeur.Value ;
      LpBase.Height:=Pixel.AsPixel ;
      end ;
   End ;
//XMG 15/04/04 fin
Pixel.AsChar[FALSE,LPTailleDef]:=MargeG.Value ;
LPBase.MargeLeft:=Pixel.AsPixel ;

Pixel.AsChar[TRUE,LPTailleDef]:=MargeH.Value ;
LpBase.MargeTop:=Pixel.AsPixel ;

Pixel.AsChar[FALSE,LPTailleDef]:=MargeD.Value ;
LpBase.MargeRight:=Pixel.AsPixel ;

Pixel.AsChar[TRUE,LPTailleDef]:=MargeB.Value ;
LPBase.MargeBottom:=Pixel.AsPixel ; 

LpBase.Units:=TLPUnits(GRMessures.itemIndex) ;

Rupt:='' ;
For i:=1 to LP_NbMaxRuptures do
  Begin
  CB:=THvalCombobox(FindComponent('FRupt'+inttostr(i))) ;
  if assigned(CB) and (CB.Value<>'') then Rupt:=Rupt+CB.Value+';' ;   
  End ;
LPBase.Ruptures:=Rupt ;
Rupt:=FindetReplace(Rupt,' ;','',TRUE) ;

for j:=low(TLPTypeBandes) to High(TLPTypeBandes) do
   if not (j in [lbPiedRupt]) then
      Begin                               
      Ok:=FALSE ;
      Cnt:=-1 ;
      Case j of
        lbEntete    : Begin Ok:=ChkEntete.Checked ; Titre:=TraduireMemoire(MC_MsgErrDefaut(1033)) ; End ;
        lbSubEntete : Begin Ok:=CntEntete.Value>1 ; Titre:=TraduireMemoire(MC_MsgErrDefaut(1070)) ; Cnt:=CntEntete.Value ; End ;
        lbDetail    : Begin Ok:=ChkDetail.Checked ; Titre:=TraduireMemoire(MC_MsgErrDefaut(1034)) ; End ;
        lbSubDetail : Begin Ok:=CntDetail.Value>1 ; Titre:=TraduireMemoire(MC_MsgErrDefaut(1037)) ; Cnt:=CntDetail.Value ; End ;
        lbPied      : Begin Ok:=ChkPied.Checked   ; Titre:=TraduireMemoire(MC_MsgErrDefaut(1035)) ; End ;
        lbSubPied   : Begin Ok:=CntPied.Value>1   ; Titre:=TraduireMemoire(MC_MsgErrDefaut(1071)) ; Cnt:=CntPied.Value ; End ;   
        End ;
      if not (j in [lbSubEntete,lbSubDetail,lbSubPied]) then
         Begin
         if (j in [lbDetail,lbPied,lbEnteteRupt]) then
            Begin
            tb:=lbEnteteRupt ;
            if j>=lbPied then tb:=lbPiedRupt ;
            Cnt:=NbCarInString(Rupt,';') ;
            while LPBase.NumMaxSubdetail(tb)<Cnt do
              Begin
              Nd:=Cnt-(Cnt-1-LPBase.NumMaxSubDetail(tb)) ;
              i:=Nd ;
              if Tb=lbPiedRupt then i:=cnt-(nd-1) ;
              St:=gtfs(Rupt,';',i) ;
              if trim(St)<>'' then
                 Begin
                 St:=TraduireMemoire(MC_MsgErrDefaut(1074+ord(tb=lbPiedRupt)))+St ;
                 CreateBande(tb,nd,St) ;
                 End ;
              End ;
            while (LPBase.NumMaxSubdetail(tb)>Cnt) do
               Begin
               LpBase.ChoixBande(tb,LPBase.NumMaxSubDetail(tb)).Free ;
               End ;
            for nd:=1 to LPBase.NumMaxSubDetail(tb) do
                Begin
                i:=Nd ;
                if Tb=lbPiedRupt then i:=cnt-(nd-1) ;
                St:=TraduireMemoire(MC_MsgErrDefaut(1074+ord(tb=lbPiedRupt)))+gtfs(Rupt,';',i) ;
                Bnd:=LPBase.ChoixBande(tb,nd) ;
                if Bnd.Titre<>St then Bnd.Titre:=st ;
                End ;
            End ;
         if j<>lbEnteteRupt then
            Begin
            if (LpBase.ChoixBande(j)=nil) and (Ok) then CreateBande(j,-1,Titre) ;
            if (LpBase.ChoixBande(j)<>nil) and (Not Ok) then LPBase.ChoixBande(j).Free ;
            End ;
         end else
         Begin
         while LPBase.NumMaxSubDetail(j)<Cnt-1 do
           Begin
           Nd:=Cnt-(Cnt-1-LPBase.NumMaxSubDetail(j)) ;
           CreateBande(j,Nd,Titre+inttostr(Nd+1)) ;
           End ;
         while (LPBase.NumMaxSubdetail(j)>Cnt-1) do
            Begin
            LpBase.ChoixBande(j,LPBase.NumMaxSubDetail(j)).Free ;
            End ;
         End ;
      End ;
ChargeBandes ;
TopBnd:=0 ;
for Nd:=0 to LesBnds.Count-1 do
  if TObject(LesBnds[Nd]).Classtype=TLPBande then
    Begin
    TLPBande(LesBnds[Nd]).Top:=TopBnd ;
    TopBnd:=TLPBande(LesBnds[ND]).Top+TLPBande(LesBnds[ND]).Height ;
    End ;
ChargeComboBande(CtrlBande) ;
ChargeComboBande(ImaBande) ;
With LPParam do
  Begin
  TypeEtat:=FTypeEtat.Text ;
  NatEtat:=FNatEtat.Text ;
  CodeEtat:=FCodeEtat.Text ;
  Libelle:=FLibelle.Text ;
  Modele:=FModele.Checked ;
  End ;
UpdateContext ;
EtatPanel.Visible:=FALSE ;
LPBase.Modified:=TRUE ;
EnableControls(LPBASE,TRUE)  ;
ChangeControl(FCtrl) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.ChargeBandes ;
Var i : Integer ;
Begin
LesBnds.clear ; LesBnds.Capacity:=LesBnds.Count ;
for i:=0 to LpBase.Controlcount-1 do
  if LpBase.Controls[i] is TLPBande then
     Begin
     LesBnds.add(TLPBande(LPBase.Controls[i])) ;
     TLPBande(LPBase.Controls[i]).OnMouseDown:=BandeMouseDown ;
     end ;
LesBnds.Sort(Compare) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.ChargeComboBande (Combo : THValCOmboBox) ;
Var j  : Integer ;
    St : String ;
Begin //Maj des bandes qu'on peut choissir
Combo.items.BeginUpdate ; Combo.Values.BeginUpdate ;
COmbo.Items.Clear ; Combo.Values.clear ;
for j:=0 to LesBnds.Count-1 do
  if LesBnds.items[j]<>nil then
     Begin
     St:=inttostr(ord(TLPBande(LesBnds[j]).TypeBande)) ;
     if TLPBande(LesBnds[j]).TypeBande in [lbSUbEntete,lbSubdetail,lbSubPied,lbEnteteRupt,lbPiedRupt] then St:=St+inttostr(TLPBande(LesBnds[j]).NbrDetail) ;  
     Combo.Values.addobject(St,LesBnds[j]) ;
     Combo.Items.add(TLPBande(LesBnds[j]).Titre) ;
     End ;
Combo.items.EndUpdate ; Combo.Values.EndUpdate ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BCancelEtatClick(Sender: TObject);
begin
EtatPanel.Visible:=FALSE ;
EnableControls(LPBASE,TRUE)  ;
ChangeControl(FCtrl) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

procedure TFLPDessin.RPoucesClick(Sender: TObject);
begin
if TailleEn.itemindex=0 then
   Begin
   Pixel.AsCM:=Largeur.Value ;
   Largeur.Value:=Pixel.AsInch ;
   if Frouleau.state=cbUnChecked then
      Begin
      Pixel.AsCM:=Longeur.Value ;
      Longeur.Value:=Pixel.AsInch ;
      end ;
   end else
   Begin
   Pixel.AsInch:=Largeur.Value ;
   Largeur.Value:=Pixel.AsCM;
   if Frouleau.state=cbUnChecked then
      Begin
      Pixel.AsInch:=Longeur.Value ;
      Longeur.Value:=Pixel.AsCM ;
      end ;
   end ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.TimerTimer(Sender: TObject);
begin
Timer.Enabled:=FALSE ;
WindowState:=wsMaximized ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.bImportClick(Sender: TObject);
Var err : TMC_Err ;
begin
If DlgImport.execute then
   Begin
   NetoyerBandes(FALSE) ;
   if LP_ImporteModeleInterne(DlgIMport.Filename,LPParam,LPBase,AsigneEventsChamp,nil,Err,ChangementVersion) then
      Begin
      ChargeBandes ;
      ChargeComboBande(CtrlBande) ;
      ChargeComboBande(ImaBande) ;
      End else PGIBox(Err.Libelle,Caption) ;
   UpdateContext ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.bExportClick(Sender: TObject);
Var Err : TMC_Err ;
begin
if PGIAsk(MC_MsgErrDefaut(1084),Caption)=mrYes then
   Begin
   DlgExport.Filename := LPParam.NatEtat+'-'+LPParam.CodeEtat+'-'+LPParam.Libelle+'.'+DlgExport.DefaultExt ;
   If DlgExport.execute then
      if Not LP_ExporteModeleInterne(DlgExport.Filename,LPParam,LPBase,Err) then PGIBox(Err.Libelle,Caption) ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BFermeClick(Sender: TObject);
begin
If Not IsInside(Self) then Modalresult:=mrCancel ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.FormDestroy(Sender: TObject);
begin
LesBnds.free ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BandeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Var Stg : String ;
    Ind : Integer ;
begin
if (ssCtrl in Shift) and (Button=mbLeft) then
   Begin
   CentreTW(Self,BndPanel) ;
   BandeCourant:=TLPBande(Sender) ;
   BndVisible.Checked:=BandeCourant.BandeVisible ;
   BndCondition.Text:=BandeCourant.Condition ;
   TBBndSQL.Visible:=(BandeCourant.typeBande in [lbSubEntete,lbSubDetail,lbSubPied]) ;
   BBndSQL.Visible:=TBBndSQL.Visible ;
   BndTesteSQL.visible:=TBBndSQL.Visible ;
   BndTesteSQL.checked:=BandeCourant.TesterSQL ; 
   BndPanel.Visible:=TRUE ;
   Stg:=BndPanel.Caption ;
   Ind:=pos(':',Stg) ;
   if Ind>0 then Stg:=Copy(BndPanel.Caption,1,Ind)+' '+BandeCourant.titre ;
   BndPanel.Caption:=Stg ;
   EnableControls(LPBASE,FALSE)  ;
   BndVisible.Setfocus ;
   Abort ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BValideBndClick(Sender: TObject);
begin
BandeCourant.BandeVisible:=BndVisible.Checked ;
BandeCourant.Condition:=BndCondition.Text ;
BandeCourant.TesterSQL:=BndTesteSQL.Checked ; 
BndPanel.Visible:=FALSE ;
EnableControls(LPBASE,TRUE)  ;
ChangeControl(FCtrl) ; 
BandeCourant:=nil ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.UnifierTables(Table1,Table2 : String ) : String ;
Var St : String ;
Begin
Result:=Table1 ;
While Table2<>'' do
  Begin
  St:=ReadTokenSt(Table2) ;
  if Pos(St+';',Result)<=0 then Result:=Result+St+';' ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFLPDessin.AjouteTablesSubDetail ( SQL : String ) ;
Var Tables : String ;
Begin
Tables:=ChargeTables(SQL) ;
FTables:=UnifierTables(FTables,Tables) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BBndSQLClick(Sender: TObject);
Var SQL : String ;
begin
if BandeCourant=nil then Exit ;
SQL:=BandeCourant.SQL ;
if SaisieSQL(SQL,BndTesteSQL.Checked) then  
   Begin
   BandeCourant.SQL:=SQL ;
   ChargeChampsSurTree ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BCancelBndClick(Sender: TObject);
begin
Bndpanel.Visible:=FALSE ;
EnableControls(LPBASE,TRUE)  ;
BandeCourant:=nil ;
ChangeControl(FCtrl) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if (Key=VK_RETURN) then
   Begin
   if not isparent(LpBase,Screen.Activecontrol) then
      Begin
      Key:=0 ;
      NextControl(TForm(GetParentForm(screen.activecontrol)),TRUE) ;
      End else
   if Assigned(FCtrl) then
      Begin
      Key:=0 ;
      ControlDblClick(FCtrl) ;
      End ;
   End else
if Key=VK_VALIDE then
   Begin
   if BndPanel.Visible  then Begin BValideBnd.Click   ; Key:=0 ; End else
   if EtatPanel.Visible then Begin BValiderEtat.Click ; Key:=0 ; End else
   if TWImage.Visible   then Begin BValiderIma.Click  ; Key:=0 ; End else
   if CtrPanel.Visible  then Begin BValiderPnl.Click  ; Key:=0 ; End ;
   End else
if Key=VK_ESCAPE then
   Begin
   if BndPanel.Visible  then Begin BCancelBnd.Click  ; Key:=0 ; End else
   if EtatPanel.Visible then Begin BCancelEtat.Click ; Key:=0 ; End else
   if TWImage.Visible   then Begin BCancelIma.Click  ; Key:=0 ; End else
   if CtrPanel.Visible  then Begin BCancelPnl.Click  ; Key:=0 ; End else
   BFerme.Click ;
   End else
if Key=VK_DELETE then
   Begin
   if (Sender is TFLPDessin) and (Assigned(FCtrl)) then
      Begin
      Key:=0 ;
      SupprimeControl(FCtrl) ;
      End ;
   End ;
application.processmessages ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.OnSauve : Boolean ;
var resp : integer ;
Begin
Result:=FALSE ;
Resp:=mrNo ;
if (LPBase.Modified) then Resp:=PGIAskCancel(MC_MsgErrDefaut(1076),Caption) ;
Case Resp of
  mrYes    : Result:=EnregistreModeles ;
  mrno     : result:=TRUE ;
  mrCancel : ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.bPreviewClick(Sender: TObject);
Var Err       : TMC_Err ;
    Apercu,ok : Boolean ; //XMG 30/03/04
begin
if not OnSauve then exit ;
err.code:=0 ; err.Libelle:='' ;
Apercu:=(Sender=bPreview) ;
if Trim(LPParam.SQL)='' then
   Begin
   Err.Code:=-1 ;
   Err.Libelle:=MC_MsgErrDefaut(1030) ;
   End ;
LibereSauvGardeLP ;
//XMG 30/03/04 début
if (Err.COde=0) then
  begin
  {$IFNDEF EXPORTASCII}
  if (LPParam.TypeEtat=TypeEtatTexte) then  //XMG 02/04/04 'T'
     ok:=((ImprimeLP(LPParam.TypeEtat,LPParam.NatEtat,LPParam.CodeEtat,'',V_MC.GetDispositif(mcPrinter),V_MC.GetPort(mcPrinter),V_MC.GetParams(mcPrinter),FALSE,Apercu,TRUE,nil,Err)) and (Err.COde=0))
  else
  {$ENDIF EXPORTASCII}
  //if (LPParam.TypeEtat=TypeEtatExportAscII) then //XMG 02/04/04
     ok:=((ExporteLP(LPParam.TypeEtat,LPParam.NatEtat,LPParam.CodeEtat,'','',Apercu,TRUE,nil,Err)) and (Err.COde=0)) ;
  if Not Ok then PGIBox(Err.Libelle,Caption) ;
  End ;
//XMG 30/03/04 fin
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.FRouleauClick(Sender: TObject);
begin
Longeur.Visible:=(FRouleau.state=cbUnChecked) ; FLongeur.Visible:=Longeur.Visible ;
Pixel.AsPixel:=1056 ;
if TailleEn.Itemindex=0 then Longeur.value:=Pixel.AsInch
   else Longeur.Value:=Pixel.AsCM ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BInsereImageClick(Sender: TObject);
Var p     : TPoint ;
    Num   : Integer ;
    Bande : TLPBande ;
    Ctrl  : TLPImage ;
begin
p:=Point(1,1) ;
Bande:=ChercheBandeParent ;
if Assigned(Bande) then
   Begin
   Num:=ChercheNumeroChamps+1 ;
   Ctrl:=TLPImage.Create(Self) ;
   Ctrl.Visible:=FALSE ;
   Ctrl.Parent:=Bande ;
   Ctrl.Nom:='#'+inttostr(Num) ;
   Ctrl.Top:=P.Y ;
   Ctrl.Left:=P.X ;
   AsigneEventsChamp(Ctrl) ;
   ControlDBlCLick(Ctrl) ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BCancelImaClick(Sender: TObject);
begin
TWImage.Visible:=FALSE ;
EnableControls(LPBASE,TRUE)  ;
if Assigned(FCtrl) then
   Begin
   if (FCtrl is TLPImage) and (Not FCtrl.Visible) then supprimeControl(FCtrl)
      else ChangeControl(FCtrl) ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BValiderImaClick(Sender: TObject);
Var Ok : Boolean ;
begin
Ok:=MajProp(FCtrl) ;
EnableControls(LPBASE,TRUE)  ;
if Ok then
   Begin
   TWImage.Visible:=FALSE ;
   ChangeControl(FCtrl) ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.ImaLogoClick(Sender: TObject);
begin
ImaContenu.Enabled:=ImaLogo.State=cbUnchecked ;
if not ImaContenu.Enabled then ImaContenu.Text:='' ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.ChkEnteteClick(Sender: TObject);
Var Chk : TCheckBox ;
    Cnt : TSpinEdit ;
begin
if not (sender is TCheckBox) then exit ;
Chk:=TCheckBox(Sender) ;
Cnt:=TSpinEdit(FindComponent('Cnt'+copy(Chk.name,4,length(Chk.Name)))) ;
if assigned(Cnt) then
   Begin
   Cnt.Enabled:=Chk.State=cbChecked ;
   if not Cnt.enabled then Cnt.Value:=1 ;
   End ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.FormMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
begin
if sbFonds.VertScrollBar.IsScrollBarVisible then Begin sbFonds.VertScrollBar.Position:=sbFonds.VertScrollBar.Position-WheelDelta ; Handled:=TRUE ; End else
if sbFonds.HorzScrollBar.IsScrollBarVisible then Begin sbFonds.HorzScrollBar.Position:=sbFonds.VertScrollBar.Position-WheelDelta ; Handled:=TRUE ; End ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.CtrlMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
if (assigned(FCtrl)) and ([ssdouble,ssleft]=Shift) then controldblClick(FCtrl) ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.bListClick(Sender: TObject);      
Var i,j,n : Integer ;
Begin
n:=0 ;
for i:=0 to LPBase.controlcount-1 do
  if LPBase.Controls[i] is TWincontrol then
     for j:=0 to TWincontrol(LPBase.Controls[i]).ControlCount-1 do
         if (TWincontrol(LPBase.Controls[i]).Controls[j] is TLPChamp) or
            (TWincontrol(LPBase.Controls[i]).Controls[j] is TLPImage) then
            Begin
            Inc(n) ;
            GLst.RowCount := n+GLst.FixedRows ;
            if (TWincontrol(LPBase.Controls[i]).Controls[j] is TLPChamp) then
               Begin
               GLst.Cells[0,n]:=TLPChamp(TWinControl(LPBase.controls[i]).controls[j]).Nom ;
               GLst.Cells[1,n]:=TLPChamp(TWinControl(LPBase.controls[i]).controls[j]).Libelle ;
               GLst.Cells[2,n]:=TLPChamp(TWinControl(LPBase.controls[i]).controls[j]).Mask ;
               End else
               Begin
               GLst.Cells[0,n]:=TLPImage(TWinControl(LPBase.controls[i]).controls[j]).Nom ;
               GLst.Cells[1,n]:=TLPImage(TWinControl(LPBase.controls[i]).controls[j]).Fichier ;
               GLst.Cells[2,n]:=Traduirememoire('Image') ;
               End ;
            End ;
// Affichage de la liste
EnableControls(LPBASE,FALSE)  ;
TWListe.Visible:=TRUE ;
CentreTW(Self, TWListe) ;
if GLst.CanFocus then GLst.SetFocus ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BRechLstClick(Sender: TObject);
Begin
FindFirst:=True ; FindChamp.Execute ;
End;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BImprimeLstClick(Sender: TObject);
Begin
PrintDBGrid(GLst, Nil, Caption, '');
End;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BValideLstClick(Sender: TObject);
Var Ctrl : TControl ;
Begin
TWListe.Visible:=FALSE ;
Ctrl:=FindControl(GLst.Cells[0,Glst.Row]) ;
if Ctrl<>Nil then ChangeControl(Ctrl) ;
GLst.RowCount:=GLst.FixedRows+1 ;
GLst.VidePile(False) ;
EnableControls(LPBASE,TRUE) ;
End;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BCancelLstClick(Sender: TObject);
Begin
TWListe.Visible:=FALSE ;
GLst.RowCount:=GLst.FixedRows+1 ;
GLst.VidePile(False) ;
EnableControls(LPBASE,TRUE) ;
End;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.BPropCtlLstClick(Sender: TObject);
Var Ctrl : TControl ;
Begin
Ctrl:=FindControl(GLst.Cells[0,Glst.Row]) ;
if Ctrl<>Nil then
   Begin
   PropCtr(Ctrl) ;
   EnableControls(LPBASE,FALSE) ;
   End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TFLPDessin.FindChampFind(Sender: TObject);
Begin
Rechercher(GLst, FindChamp, FindFirst) ;
End;
/////////////////////////////////////////////////////////////////////////////////////////
Function TFLPDessin.FindControl(NomChamp : String) : TControl ;
Var Nom : String ;
    i,j : Integer ;
Begin
result:=Nil ;
if NomChamp='' then Exit ;
for i:=0 to LPBase.controlcount-1 do
  if LPBase.Controls[i] is TWincontrol then
     for j:=0 to TWincontrol(LPBase.Controls[i]).ControlCount-1 do
         if (TWincontrol(LPBase.Controls[i]).Controls[j] is TLPChamp) or
            (TWincontrol(LPBase.Controls[i]).Controls[j] is TLPImage) then
            Begin
            if (TWincontrol(LPBase.Controls[i]).Controls[j] is TLPChamp) then Nom:=TLPChamp(TWinControl(LPBase.controls[i]).controls[j]).Nom
               else Nom:=TLPImage(TWinControl(LPBase.controls[i]).controls[j]).Nom ;
            if Nom=NomChamp then
               Begin
               Result:=TWinControl(LPBase.Controls[i]).Controls[j] ;
               Exit ;
               End ;
            End ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////

end.

