unit Scenario;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Hctrls, Buttons, ExtCtrls, ComCtrls, DBCtrls,
  hmsgbox, Mask, Grids, DBGrids, HDB, SaisUtil, HCompte, HEnt1, Spin, Ent1,
  HSysMenu,ParamDBG,  MajTable,
  AtChComp,LibChpLi,
  Hqry, HPanel, UiUtil,UtilSais,
  HTB97, Choix, ULibEcriture,ULibWindows,UTOB,ed_tools, ADODB; // pour le TQRProgressForm ;

type
  TFScenario = class(TForm)
    SSuiviCPTA: TDataSource;
    TSuiviCPTA: THTable;
    MsgBox: THMsgBox;
    Panel2: TPanel;
    GR: THDBGrid;
    DBNav: TDBNavigator;
    HMTrad: THSystemMenu;
    Pages: TPageControl;
    TS1: TTabSheet;
    TS3: TTabSheet;
    FListe: THGrid;
    Racines: TGroupBox;
    LE6: TCheckBox;
    LE3: TCheckBox;
    LE7: TCheckBox;
    LE9: TCheckBox;
    LE8: TCheckBox;
    LE4: TCheckBox;
    LE5: TCheckBox;
    Libres: TGroupBox;
    LT1: TCheckBox;
    LT2: TCheckBox;
    LT3: TCheckBox;
    LT4: TCheckBox;
    LT5: TCheckBox;
    LT6: TCheckBox;
    LT7: TCheckBox;
    LT8: TCheckBox;
    LT9: TCheckBox;
    LT10: TCheckBox;
    Scenario: TGroupBox;
    SC_OUVREECHE: TDBCheckBox;
    SC_OUVREANAL: TDBCheckBox;
    SC_LETTRAGESAISIE: TDBCheckBox;
    SC_VALIDE: TDBCheckBox;
    SC_ALERTEDEV: THDBValComboBox;
    TSC_NumAxe: THLabel;
    CONTROLETVA: THLabel;
    SC_CONTROLETVA: THDBValComboBox;
    SC_NUMAXE: THDBValComboBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    LC1: TCheckBox;
    LC2: TCheckBox;
    LC3: TCheckBox;
    LC4: TCheckBox;
    LM1: TCheckBox;
    LM2: TCheckBox;
    LM3: TCheckBox;
    LM4: TCheckBox;
    LB1: TCheckBox;
    LB2: TCheckBox;
    LD1: TCheckBox;
    TS2: TTabSheet;
    Entete: TGroupBox;
    SC_LIBELLE: TDBCheckBox;
    SC_REFINTERNE: TDBCheckBox;
    SC_REFLIBRE: TDBCheckBox;
    SC_REFEXTERNE: TDBCheckBox;
    SC_AFFAIRE: TDBCheckBox;
    SC_DATEREFEXTERNE: TDBCheckBox;
    LibreEnt: TGroupBox;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    ET1: TCheckBox;
    ET2: TCheckBox;
    ET3: TCheckBox;
    ET4: TCheckBox;
    ET5: TCheckBox;
    ET6: TCheckBox;
    ET7: TCheckBox;
    ET8: TCheckBox;
    ET9: TCheckBox;
    ET10: TCheckBox;
    ET11: TCheckBox;
    ET12: TCheckBox;
    ET13: TCheckBox;
    ET14: TCheckBox;
    ET15: TCheckBox;
    ET16: TCheckBox;
    ET17: TCheckBox;
    ET18: TCheckBox;
    ET19: TCheckBox;
    ET20: TCheckBox;
    ET21: TCheckBox;
    SC_CONTROLEBUDGET: TDBCheckBox;
    TSuiviCPTASC_JOURNAL: TStringField;
    TSuiviCPTASC_NATUREPIECE: TStringField;
    TSuiviCPTASC_OUVREECHE: TStringField;
    TSuiviCPTASC_OUVREANAL: TStringField;
    TSuiviCPTASC_NUMAXE: TStringField;
    TSuiviCPTASC_LETTRAGESAISIE: TStringField;
    TSuiviCPTASC_CONTROLETVA: TStringField;
    TSuiviCPTASC_QUALIFPIECE: TStringField;
    TSuiviCPTASC_REFINTERNE: TStringField;
    TSuiviCPTASC_LIBELLE: TStringField;
    TSuiviCPTASC_REFAUTO: TStringField;
    TSuiviCPTASC_REFEXTERNE: TStringField;
    TSuiviCPTASC_DATEREFEXTERNE: TStringField;
    TSuiviCPTASC_REFLIBRE: TStringField;
    TSuiviCPTASC_AFFAIRE: TStringField;
    TSuiviCPTASC_RADICAL1: TStringField;
    TSuiviCPTASC_RADICAL2: TStringField;
    TSuiviCPTASC_RADICAL3: TStringField;
    TSuiviCPTASC_RADICAL4: TStringField;
    TSuiviCPTASC_RADICAL5: TStringField;
    TSuiviCPTASC_RADICAL6: TStringField;
    TSuiviCPTASC_RADICAL7: TStringField;
    TSuiviCPTASC_RADICAL8: TStringField;
    TSuiviCPTASC_RADICAL9: TStringField;
    TSuiviCPTASC_RADICAL10: TStringField;
    TSuiviCPTASC_COMPLEMENTS1: TStringField;
    TSuiviCPTASC_COMPLEMENTS2: TStringField;
    TSuiviCPTASC_COMPLEMENTS3: TStringField;
    TSuiviCPTASC_COMPLEMENTS4: TStringField;
    TSuiviCPTASC_COMPLEMENTS5: TStringField;
    TSuiviCPTASC_COMPLEMENTS6: TStringField;
    TSuiviCPTASC_COMPLEMENTS7: TStringField;
    TSuiviCPTASC_COMPLEMENTS8: TStringField;
    TSuiviCPTASC_COMPLEMENTS9: TStringField;
    TSuiviCPTASC_COMPLEMENTS10: TStringField;
    TSuiviCPTASC_SOCIETE: TStringField;
    TSuiviCPTASC_USERGRP: TStringField;
    TSuiviCPTASC_RIB: TStringField;
    TSuiviCPTASC_TIERSPAYEUR: TStringField;
    TSuiviCPTASC_ETABLISSEMENT: TStringField;
    TSuiviCPTASC_VALIDE: TStringField;
    TSuiviCPTASC_COMPLIBRE1: TStringField;
    TSuiviCPTASC_COMPLIBRE2: TStringField;
    TSuiviCPTASC_COMPLIBRE3: TStringField;
    TSuiviCPTASC_COMPLIBRE4: TStringField;
    TSuiviCPTASC_COMPLIBRE5: TStringField;
    TSuiviCPTASC_COMPLIBRE6: TStringField;
    TSuiviCPTASC_COMPLIBRE7: TStringField;
    TSuiviCPTASC_COMPLIBRE8: TStringField;
    TSuiviCPTASC_COMPLIBRE9: TStringField;
    TSuiviCPTASC_COMPLIBRE10: TStringField;
    TSuiviCPTASC_CONTROLEBUDGET: TStringField;
    TSuiviCPTASC_LIBREBOOL1: TStringField;
    TSuiviCPTASC_LIBREBOOL2: TStringField;
    TSuiviCPTASC_LIBREBOOL3: TStringField;
    TSuiviCPTASC_LIBREBOOL4: TStringField;
    TSuiviCPTASC_LIBREBOOL5: TStringField;
    TSuiviCPTASC_DOCUMENT: TStringField;
    TSuiviCPTASC_LIBREENTETE: TStringField;
    HSC_RIB: THLabel;
    SC_RIB: THDBValComboBox;
    H_DOCUMENT: THLabel;
    SC_DOCUMENT: THDBValComboBox;
    TSuiviCPTASC_ATTRCOMP: TMemoField;
    FAutoSave: TCheckBox;
    Dock: TDock97;
    HPB: TToolWindow97;
    BFerme: TToolbarButton97;
    BDelete: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BNouveau: TToolbarButton97;
    BInfoComp: TToolbarButton97;
    BCopier: TToolbarButton97;
    TSuiviCPTASC_ALERTEDEV: TStringField;
    TSC_ALERTEDEV: THLabel;
    SC_DATEOBLIGEE: TDBCheckBox;
    TSuiviCPTASC_DATEOBLIGEE: TStringField;
    PageControl1: TPageControl;
    TS4: TTabSheet;
    TS5: TTabSheet;
    Cle: TPanel;
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    TSC_QUALIFPIECE: THLabel;
    TSC_ETABLISSEMENT: THLabel;
    JOURNAL: THDBValComboBox;
    NATUREPIECE: THDBValComboBox;
    GROUPE: THDBValComboBox;
    SC_QUALIFPIECE: THDBValComboBox;
    SC_ETABLISSEMENT: THDBValComboBox;
    JOURNALMULTI: THMultiValComboBox;
    NATUREPIECEMULTI: THMultiValComboBox;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    BGenerer: TToolbarButton97;
    SC_CONTROLEQTE: TDBCheckBox;
    TSuiviCPTASC_CONTROLEQTE: TStringField;
    Procedure FormShow(Sender: TObject);
    Procedure JOURNALChange(Sender: TObject);
    Procedure BDeleteClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure FListeRowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure BImprimerClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure SSuiviCPTADataChange(Sender: TObject; Field: TField);
    procedure GRKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BNouveauClick(Sender: TObject);
    procedure SSuiviCPTAStateChange(Sender: TObject);
    Procedure OkModifRacines (Sender : TObject) ;
    procedure ForceChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure SSuiviCPTAUpdateData(Sender: TObject);
    procedure SC_QUALIFPIECEChange(Sender: TObject);
    procedure BInfoCompClick(Sender: TObject);
    procedure BCopierClick(Sender: TObject);
    procedure BGenererClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure OnClickQte( Sender : TObject ) ;
  private
    FClosing : boolean;  //SG6 FQ 15009 30/11/2004
    Complement : Array[1..10,1..10] of Char ;
    CompLibre  : Array[1..10,1..40] of Char ;
    Radicaux   : Array[1..10] of String[10] ;
    LibreEntete: Array[1..40] of Char ;
    Modifier,GeneCharge,CbFlaguer : Boolean ;
    FBoBordereau : boolean ; // tru si le journal est de type bordereau, les deux premier onglet sont desactive
    Procedure GoEnable(St : String) ;
    procedure EnableCeQuiFaut ;
    procedure AttribLesHint ;
    procedure Disa ( CT,CL : TCheckBox ; VV : Boolean ) ;
    Procedure RAZRacines ;
    Procedure InitTableaux ;
    Procedure InitValeurs ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure NewEnreg ;
    Procedure Chargecompl ;
    Procedure SauveCompl ;
    Procedure MajCompl(ou : Longint) ;
    Function  OnSauve : Boolean ;
    Function  EnregOk : Boolean ;
    Function  EnregVide : Boolean ;
    procedure ChargeEnreg ;
    procedure VerifDisa ( Tout : boolean ) ;
    Procedure GenereListe(Li : HTStrings) ;
    Procedure ChercheAcces(Li : HTStrings) ;
    Function  AjouteAcces(C : TCustomCheckBox ; MemoSt,StTemp1,Radic : String ; QuelTab : Byte) : String ;
    function EnregOkMulti: Boolean;
    function  EstSaisieQte : Boolean ;
    procedure GereAccesCtrlQte ;
  public
    GeneJal,GeneNat : String3 ;
    Mode : integer ;
  end;

Procedure ParamScenario ( Jal,Nat : String3 ) ;
Procedure ValeurAttributComp(Li : HTStrings ; Radical,Champ : String ; Var Oblg,Modif,Valdef,TexteLibre : String ) ;
Procedure MultiScenario ;

implementation

{$R *.DFM}

uses PrintDBG,UtilPGI , TntStdCtrls, TntWideStrings;

Procedure ParamScenario ( Jal,Nat : String3 ) ;
Var X  : TFScenario ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture'],True,'nrSaisieCreat') then Exit ;
X:=TFScenario.Create(Application) ;
X.GeneJal:=Jal ; X.GeneNat:=Nat ; X.Mode:=0 ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.ShowModal ;
    Finally
     X.Free ;
     _Bloqueur('nrSaisieCreat',False) ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

Procedure MultiScenario ;
Var X  : TFScenario ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture'],True,'nrSaisieCreat') then Exit ;
X:=TFScenario.Create(Application) ;
X.Mode:=1 ; PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.ShowModal ;
    Finally
     X.Free ;
     _Bloqueur('nrSaisieCreat',False) ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

procedure TFScenario.Disa ( CT,CL : TCheckBox ; VV : Boolean ) ;
BEGIN
if Not VV then
   BEGIN
   CT.Checked:=False ; CT.Enabled:=False ;
   CL.Checked:=False ; CL.Enabled:=False ;
   END ;
END ;

procedure TFScenario.AttribLesHint ;
Var S     : String ;
    VV    : Boolean ;
    i     : integer ;
    CT,CL : TCheckBox ;
BEGIN
(*{$IFNDEF CCS3}*)
VV:=False ;
{Textes libres}
for i:=1 to 10 do
    BEGIN
    CT:=TCheckBox(FindComponent('ET'+IntToStr(i))) ; CL:=TCheckBox(FindComponent('LT'+IntToStr(i))) ;
    S:='E_LIBRETEXTE'+IntToStr(i-1) ;
    if PersoChamp('E',S,VV) then BEGIN CT.Hint:=S ; CL.Hint:=S ; Disa(CT,Cl,VV) ; END ;
    END ;
{Tables libres}
for i:=1 to 4 do
    BEGIN
    CT:=TCheckBox(FindComponent('ET'+IntToStr(i+10))) ; CL:=TCheckBox(FindComponent('LC'+IntToStr(i))) ;
    S:='E_TABLE'+IntToStr(i-1) ;
    if PersoChamp('E',S,VV) then BEGIN CT.Hint:=S ; CL.Hint:=S ; Disa(CT,Cl,VV) ; END ;
    END ;
{Montants libres}
for i:=1 to 4 do
    BEGIN
    CT:=TCheckBox(FindComponent('ET'+IntToStr(i+14))) ; CL:=TCheckBox(FindComponent('LM'+IntToStr(i))) ;
    S:='E_LIBREMONTANT'+IntToStr(i-1) ;
    if PersoChamp('E',S,VV) then BEGIN CT.Hint:=S ; CL.Hint:=S ; Disa(CT,Cl,VV) ; END ;
    END ;
{Choixs O/N libres}
for i:=1 to 2 do
    BEGIN
    CT:=TCheckBox(FindComponent('ET'+IntToStr(i+18))) ; CL:=TCheckBox(FindComponent('LB'+IntToStr(i))) ;
    S:='E_LIBREBOOL'+IntToStr(i-1) ;
    if PersoChamp('E',S,VV) then BEGIN CT.Hint:=S ; CL.Hint:=S ; Disa(CT,Cl,VV) ; END ;
    END ;
{Dates libre}
CT:=TCheckBox(FindComponent('ET21')) ; CL:=TCheckBox(FindComponent('LD1')) ;
S:='E_LIBREDATE' ; if PersoChamp('E',S,VV) then BEGIN CT.Hint:=S ; CL.Hint:=S ; Disa(CT,Cl,VV) ; END ;
(*{$ENDIF}*)
END ;

Procedure TFScenario.GoEnable(St : String) ;
Var C : TComponent ;
    CC : TControl ;
BEGIN
C:=Self.FindComponent(St) ;
If C<>NIL Then BEGIN CC:=TControl(C) ; CC.Enabled:=FALSE ; END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 11/09/2002
Modifié le ... : 11/10/2002
Description .. : - 11/09/2002 - gestion de l'affichage du bouton de
Suite ........ : generation multiple
Suite ........ : - 08/10/2002 - amelioration du look de le fenetre
Suite ........ : - 11/10/2002 - pour les journaux de type bordereau les 2
Suite ........ : premier onglet ne sont pas accessibles
Mots clefs ... :
*****************************************************************}
procedure TFScenario.EnableCeQuiFaut ;
Var i : Integer ;
BEGIN
For i:=1 To 21 Do GoEnable('ET'+IntToStr(i)) ;
ET11.Enabled:=TRUE ;
If not EstSerie(S3) Then ET12.Enabled:=TRUE ;
GoEnable('Label7') ; GoEnable('Label9') ;
GoEnable('Label10') ; GoEnable('Label11') ;
For i:=1 To 10 Do GoEnable('LT'+IntToStr(i)) ;
For i:=1 To 4 Do GoEnable('LC'+IntToStr(i)) ;
For i:=1 To 4 Do GoEnable('LM'+IntToStr(i)) ;
For i:=1 To 2 Do GoEnable('LB'+IntToStr(i)) ;
GoEnable('LD1') ;
GoEnable('Label1') ; GoEnable('Label3') ;
GoEnable('Label5') ; GoEnable('Label4') ;
LC1.Enabled:=TRUE ;
If not EstSerie(S3) Then LC2.Enabled:=TRUE ;
If EstSerie(S3) Then BEGIN GoEnable('LE7') ; GoEnable('LE8') ; END ;
if Mode=0 then Pages.ActivePageIndex:=0 else Pages.ActivePageIndex:=2 ;
TS1.Enabled:=Mode=0 ; Scenario.Enabled:=Mode=0 ;
TS2.Enabled:=Mode=0 ;
GR.Visible:=Mode=0 ;
if (not Scenario.Enabled) or FBoBordereau then
 begin
  CDisableControl(TCustomControl(Scenario)) ; CDisableControl(TCustomControl(Entete)) ;
 end;

if not GR.Visible then
 begin
  self.Width:=self.Width-GR.Width ;
  GR.Width:=0;
  GR.align:=alNone ;
  BFerme.Left:=HPB.Width-35 ;
  BGenerer.Left:=BFerme.Left-35 ;
 end; // if
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 31/07/2002
Modifié le ... : 17/09/2002
Description .. : - 31/07/2002 - gestion de l'affichage des panels
Suite ........ : - pour la generation multiple on se place sur un nouvelle
Suite ........ : enregistrement
Suite ........ : 
Suite ........ : - 17/09/2002 - alignement du bouton fermer à droite en cas 
Suite ........ : de multiecheance
Mots clefs ... : 
*****************************************************************}
procedure TFScenario.FormShow(Sender: TObject);
Var i : Byte ;
Begin
Journal.Style := csDropDownList ;
SC_ETABLISSEMENT.Style := csDropDownList;
ChangeSizeMemo(TSuiviCPTASC_ATTRCOMP) ;
Pages.ActivePage:=Pages.Pages[0] ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; GeneCharge:=True ; CbFlaguer:=False ;
// VL 071103 FQ 12792
if ((EstSerie(S3)) or (EstSerie(S5))) then
   BEGIN
   EnableCeQuiFaut ;
   if EstSerie(S3) then
      BEGIN
      SC_QUALIFPIECE.Visible:=False   ; TSC_QUALIFPIECE.Visible:=False ;
      SC_ETABLISSEMENT.Visible:=False ; TSC_ETABLISSEMENT.Visible:=False ;
      SC_CONTROLEBUDGET.Visible:=False ;
      SC_NUMAXE.Visible:=False ; TSC_NUMAXE.Visible:=False ;
      SC_DOCUMENT.Visible:=False ; H_DOCUMENT.Visible:=False ;
      END ;
   END else
   BEGIN
   AttribLesHint ;
   END ;
BGenerer.visible:=Mode=1;
// FIN VL FQ 12792
TSuiviCPTA.Open ;
For i:=1 to 10 do
  Begin
  FListe.Colaligns[0]:=TaCenter ;
  FListe.Cells[0,i]:=IntToStr(i) ;
  End ;
TS5.TabVisible:=false ;TS4.TabVisible:=false ;
if Mode=0 then PageControl1.ActivePage:=TS4 else PageControl1.ActivePage:=TS5 ;
BAnnuler.Visible:=Mode=0 ; BNouveau.Visible:=BAnnuler.Visible ; BDelete.Visible:=BAnnuler.Visible ;
BInfoComp.Visible:=BAnnuler.Visible ;BCopier.Visible:=BAnnuler.Visible ; BImprimer.Visible:=BAnnuler.Visible ;
BValider.Visible:=BAnnuler.Visible ; BAide.Visible:=BAnnuler.Visible ;BFirst.Visible:=BAnnuler.Visible ;
BPrev.Visible:=BAnnuler.Visible ; BNext.Visible:=BAnnuler.Visible ; BLast.Visible:=BAnnuler.Visible ;
// LG plus la generation multiple on se place sur un nouvelle enregistrement
if EnregVide or (Mode=1) then BEGIN Bouge(nbInsert) ; Modifier:=False ; Exit ; END ;
if (GeneJal<>'') and (GeneNat<>'') then TSuiviCpta.Findkey([V_PGI.Groupe,GeneJal,GeneNat]) ;
GeneCharge:=False ;
FClosing := False; //SG6 30/11/2004 FQ 15009
//SG6 11.03.05
if VH^.AnaCroisaxe then
begin
  TSC_NumAxe.Visible := False;
  SC_NumAxe.Visible := False;
end;


end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 11/10/2002
Modifié le ... :   /  /    
Description .. : - 11/10/2002 - en mode bordereau les deux premier onglet 
Suite ........ : ne sont pas accessible
Mots clefs ... : 
*****************************************************************}
procedure TFScenario.JOURNALChange(Sender: TObject);
  Var  Q : TQuery ;
begin
// Maj du combo des natures de pieces
  // FQ 20134 BVE 02.05.07
  if (JOURNAL.Field.AsString <> '') and (JOURNAL.Value = '' )then
     JOURNAL.ItemIndex := JOURNAL.Values.IndexOf(trim(JOURNAL.Field.AsString));
  // END FQ 20134
Q:=OpenSQL('Select J_NATUREJAL,J_MODESAISIE from JOURNAL Where J_JOURNAL="'+JOURNAL.Value + '"',True) ;
If Not Q.EOF then
  BEGIN
  case CaseNatJal(Q.Findfield('J_NatureJal').AsString) of
    tzJVente       : NaturePiece.DataType:='ttNatPieceVente'  ;
    tzJAchat       : NaturePiece.DataType:='ttNatPieceAchat'  ;
    tzJBanque      : NaturePiece.DataType:='ttNatPieceBanque' ;
    tzJEcartChange : NaturePiece.DataType:='ttNatPieceEcartChange' ;
    tzJOD          : NaturePiece.DataType:='ttNaturePiece'    ;
    END ;
  END ;
NATUREPIECEMULTI.DataType:='ttNaturePiece'  ; FBoBordereau:=(Q.Findfield('J_MODESAISIE').AsString='BOR') or (Q.Findfield('J_MODESAISIE').AsString='LIB') ;
Ferme(Q) ;
if FBoBordereau then
 begin
  CDisableControl(TCustomControl(Scenario)) ; CDisableControl(TCustomControl(Entete)) ;
  Pages.ActivePageIndex:=2 ;
 end
  else
   begin
    CEnableControl(TCustomControl(Scenario)) ; CEnableControl(TCustomControl(Entete)) ;
   end;
   // FQ 20134 BVE 02.05.07
   NaturePiece.ReLoad;
   NaturePiece.ItemIndex := NaturePiece.Values.IndexOf(Trim(NaturePiece.Field.AsString));
   // END FQ 20134
end ;


Function TFScenario.EnregOk : Boolean ;
var UserGrp : String ;
BEGIN
Result:=False ;
Modifier:=True ;
if TSuiviCpta.State in [dsEdit,dsInsert]=False then Exit ;
if (JOURNAL.Value='') then BEGIN Msgbox.Execute(10,'','') ; Exit ; END ;
if (NATUREPIECE.Value='') then BEGIN Msgbox.Execute(11,'','') ; Exit ; END ;
if (GROUPE.ItemIndex=-1) then BEGIN Msgbox.Execute(9,'','') ; Exit ; END ;
{$IFNDEF CCS3}
if (SC_ETABLISSEMENT.ItemIndex=-1) then BEGIN Msgbox.Execute(15,'','') ; Exit ; END ;
if (SC_QUALIFPIECE.ItemIndex=-1) then BEGIN Msgbox.Execute(16,'','') ; Exit ; END ;
{$ENDIF}
if ((TSuiviCPTASC_ETABLISSEMENT.AsString='') or (SC_ETABLISSEMENT.Value='')) then TSuiviCPTASC_ETABLISSEMENT.AsString:='...' ;
if ((TSuiviCPTASC_QUALIFPIECE.AsString='') or (SC_QUALIFPIECE.Value='')) then TSuiviCPTASC_QUALIFPIECE.AsString:='...' ;
if ((TSuiviCPTASC_USERGRP.AsString='') or (GROUPE.Value='')) then TSuiviCPTASC_USERGRP.AsString:='...' ;
if GROUPE.ItemIndex=-1 then UserGrp:='' else UserGrp:=GROUPE.Value ;
if TSuiviCpta.State=dsinsert then
   BEGIN
   if PresenceComplexe('SUIVCPTA',
                       ['SC_ETABLISSEMENT','SC_USERGRP','SC_JOURNAL','SC_NATUREPIECE','SC_QUALIFPIECE'],
                       ['=','=','=','=','='],
                       [TSuiviCPTASC_ETABLISSEMENT.AsString,TSuiviCPTASC_USERGRP.AsString,JOURNAL.Value,NaturePiece.Value,TSuiviCPTASC_QUALIFPIECE.AsString],
                       ['S','S','S','S','S']) then BEGIN Msgbox.Execute(12,'','') ; Exit ; END ;
   END ;
Modifier:=False ;
Result:=True ;
END ;

Procedure TFScenario.NewEnreg ;
BEGIN
Cle.Enabled:=True ;
Scenario.Enabled:=True ; Entete.Enabled:=True ; Racines.Enabled:=True ; Libres.Enabled:=True ;
InitValeurs ; InitTableaux ;
RAZRacines ; Caption:=MsgBox.Mess[14] ;
END ;

Procedure TFScenario.InitValeurs ;
var i : Integer ;
    Lib : String ;
BEGIN
InitNew(TSuiviCpta) ;
TSuiviCptaSC_OUVREECHE.AsString:='X' ;
TSuiviCptaSC_OUVREANAL.AsString:='X' ;
TSuiviCptaSC_CONTROLEQTE.AsString:='-' ;
TSuiviCptaSC_CONTROLETVA.AsString:='RIE' ;
TSuiviCptaSC_LETTRAGESAISIE.AsString:='-' ;
TSuiviCptaSC_NUMAXE.AsString:='A1' ;
TSuiviCptaSC_ALERTEDEV.AsString:='JOU' ;
TSuiviCptaSC_RIB.AsString:='...' ;
// Index
TSuiviCPTASC_ETABLISSEMENT.AsString:='...' ;
TSuiviCPTASC_QUALIFPIECE.AsString:='...' ;
TSuiviCPTASC_USERGRP.AsString:='...' ;
// Entete
TSuiviCptaSC_DATEREFEXTERNE.AsString:= '-' ;
TSuiviCptaSC_REFLIBRE.AsString:= '-' ;
TSuiviCptaSC_AFFAIRE.AsString:= '-' ;
TSuiviCptaSC_REFINTERNE.AsString:='-' ;
TSuiviCptaSC_REFEXTERNE.AsString:='-' ;
TSuiviCptaSC_REFAUTO.AsString:='' ;
// Radicaux / compléments / Zones libre ligne et entetes
Lib:='' ;
for i:=1 to 40 do Lib:=Lib+'-' ;
TSuiviCpta.FindField('SC_LIBREENTETE').AsString:=Lib ;
For i := 1 to 10 do
  BEGIN
  TSuiviCpta.FindField('SC_COMPLEMENTS'+IntToStr(i)).AsString:='----------' ;
  TSuiviCpta.FindField('SC_RADICAL'+IntToStr(i)).AsString:='' ;
  FListe.Cells[1,i]:='' ;
  TSuiviCpta.FindField('SC_COMPLIBRE'+IntToStr(i)).AsString:=Lib ;
  END;
END ;

procedure TFScenario.BValiderClick(Sender: TObject);
begin
Bouge(NbPost) ;
end;

procedure TFScenario.BAnnulerClick(Sender: TObject);
begin
Bouge(nbCancel) ;
end;

procedure TFScenario.BDeleteClick(Sender: TObject);
begin
Bouge(NbDelete) ;
end;

procedure TFScenario.FormClose(Sender: TObject; var Action: TCloseAction);
BEGIN
TSuiviCpta.Close ;
if Parent is THPanel then
   BEGIN
   _Bloqueur('nrSaisieCreat',False) ;
   Action:=caFree ;
   END ;
FClosing := True;  //SG6 30/11/2004 FQ 15009
END;

Function TFScenario.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
Result:=FALSE  ;
Case Button of
  NbFirst,nbPrior,nbNext,nbLast,
  NbInsert : if not OnSauve then exit ;
  NbPost : if Not EnregOk then Exit else BEGIN MajCompl(Fliste.Row) ; SauveCompl ; END ;
  NbDelete : if Msgbox.Execute(2,'','')<>mrYes then Exit ;
  End ;
if not TransacNav(DBNav.BtnClick,Button,10) then BEGIN MessageAlerte(Msgbox.Mess[5]) ; Exit ; END ;
Case Button of
  NbFirst,nbPrior,nbNext,nbLast : ;//InitTableaux ;
  NbInsert : NewEnreg ;
  NbPost   : ;
  NbCancel, NbDelete : if EnregVide then Bouge(nbInsert) ;
  END ;
Result:=True ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 01/08/2002
Modifié le ... :   /  /    
Description .. : - 01/08/2002 - on genaration de scenario cette fonction 
Suite ........ : renvoie toujours vrai. 
Mots clefs ... :
*****************************************************************}
Function TFScenario.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=true ; if Mode=1 then exit ;
if TSuiviCpta.State <> dsBrowse then nextprevcontrol(Self);          //SG6 30/11/2004 FQ 15009
Result:=False  ; Modifier:=True ;
If TSuiviCpta.Modified or FAutoSave.Checked then
   BEGIN
   If FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,'','') ;
   END else Rep:=321 ;
   Case Rep of
     mrYes    : if not Bouge(nbPost) then Exit ;
     mrNo     : Bouge(nbCancel) ;
     mrCancel : Exit ;
   END ;
Modifier:=False ;
Result:=True  ;
end ;

Function TFScenario.EnregVide : Boolean ;
BEGIN
Result:=(TSuiviCpta.BOF and TSuiviCpta.EOF) ;
END ;

Procedure TFScenario.RAZRacines ;
  var i : integer ;
      C: TControl ;
BEGIN
For i:=1 to 10 do FListe.Cells[1,i]:='' ;
For i:=0 to Racines.ControlCount-1 do
  BEGIN ;
  C:=Racines.Controls[i] ;
  If C is TCheckBox Then TCheckBox(C).Checked:=False ;
  END ;
For i:=0 to Libres.ControlCount-1 do
  BEGIN ;
  C:=Libres.Controls[i] ;
  If C is TCheckBox Then TCheckBox(C).Checked:=False ;
  END ;
For i:=0 to LibreEnt.ControlCount-1 do
  BEGIN ;
  C:=LibreEnt.Controls[i] ;
  If C is TCheckBox Then TCheckBox(C).Checked:=False ;
  END ;
END ;

Procedure TFScenario.InitTableaux ;
var i,j :  integer ;
    Compl : String[10] ;
    Lib : String[40] ;
BEGIN
FillChar(Radicaux,SizeOf(Radicaux),#0) ;
FillChar(Complement,SizeOf(Complement),'-') ;
FillChar(CompLibre,SizeOf(CompLibre),'-') ;
FillChar(LibreEntete,SizeOf(LibreEntete),'-') ;
if EnregVide then Exit ;
For i:=1 to 10 do
  BEGIN
  Radicaux[i]:=Copy(TSuiviCpta.FindField('SC_RADICAL'+IntToStr(i)).AsString,1,10) ;
  Fliste.Cells[1,i]:=Radicaux[i] ;
  Compl:=TSuiviCpta.FindField('SC_COMPLEMENTS'+IntToStr(i)).AsString ;
  Lib:=TSuiviCpta.FindField('SC_COMPLIBRE'+IntToStr(i)).AsString ;
  For j:=1 to 10 do Complement[i,j]:=Compl[j] ;
  For j:=1 to 40 do CompLibre[i,j]:=Lib[j] ;
  END ;
Lib:=TSuiviCpta.FindField('SC_LIBREENTETE').AsString ;
For j:=1 to 40 do LibreEntete[j]:=Lib[j] ;
ChargeCompl ;
END ;

Procedure TFScenario.ChargeCompl ;
  var i :  integer ;
      C : TControl ;
BEGIN
For i:=0 to Racines.ControlCount-1 do
  BEGIN ;
  C:=Racines.Controls[i] ;
  If C is TCheckBox Then
    BEGIN
    If C.Tag>0 Then TCheckBox(C).Checked:=(Complement[Fliste.Row,C.Tag]='X') ;
    END ;
  END ;
For i:=0 to Libres.ControlCount-1 do
  BEGIN ;
  C:=Libres.Controls[i] ;
  If C is TCheckBox Then
    BEGIN
    If C.Tag>0 Then TCheckBox(C).Checked:=(CompLibre[Fliste.Row,C.Tag]='X') ;
    END ;
  END ;
For i:=0 to LibreEnt.ControlCount-1 do
  BEGIN ;
  C:=LibreEnt.Controls[i] ;
  If C is TCheckBox Then
    BEGIN
    If C.Tag>0 Then TCheckBox(C).Checked:=(LibreEntete[C.Tag]='X') ;
    END ;
  END ;
END ;

Procedure TFScenario.MajCompl (ou : Longint) ;
  var i : Integer ;
      C : TControl ;
      Ch : Char ;
BEGIN
Radicaux[ou]:=Fliste.Cells[1,ou] ;
For i:=0 to Racines.ControlCount-1 do
    BEGIN
    C:=Racines.Controls[i] ;
    If C is TCheckBox Then
       BEGIN
       If C.Tag>0 Then
          BEGIN
          If TCheckBox(C).Checked then Ch:='X' else Ch:='-' ;
          Complement[ou,C.Tag]:=Ch ;
          END ;
       END ;
    END ;
For i:=0 to Libres.ControlCount-1 do
    BEGIN
    C:=Libres.Controls[i] ;
    If C is TCheckBox Then
       BEGIN
       If C.Tag>0 Then
          BEGIN
          If TCheckBox(C).Checked then Ch:='X' else Ch:='-' ;
          CompLibre[ou,C.Tag]:=Ch ;
          END ;
       END ;
    END ;
For i:=0 to LibreEnt.ControlCount-1 do
    BEGIN
    C:=LibreEnt.Controls[i] ;
    If C is TCheckBox Then
       BEGIN
       If C.Tag>0 Then
          BEGIN
          If TCheckBox(C).Checked then Ch:='X' else Ch:='-' ;
          LibreEntete[C.Tag]:=Ch ;
          END ;
       END ;
    END ;

  // Gestion accès contrôle des qtes
  GereAccesCtrlQte;

END ;

Procedure TFScenario.SauveCompl ;
  var i,j : Integer ;
      Compl : String[10] ;
      Lib   : String[40] ;
BEGIN
//TSuiviCpta.Edit ;
For i:=1 to 10 do
    BEGIN
    Compl:='--' ; Lib:='' ;
    TSuiviCpta.FindField('SC_RADICAL'+IntToStr(i)).AsString:=Radicaux[i] ;
    For j:=3 to 10 do Compl:=Compl+Complement[i,j] ;
    TSuiviCpta.FindField('SC_COMPLEMENTS'+IntToStr(i)).AsString:=Compl ;
    For j:=1 to 40 do Lib:=Lib+CompLibre[i,j] ;
    TSuiviCpta.FindField('SC_COMPLIBRE'+IntToStr(i)).AsString:=Lib ;
    END ;
Lib:='' ;
for i:=1 to 40 do
    Lib:=Lib+LibreEntete[i] ;
TSuiviCpta.FindField('SC_LIBREENTETE').AsString:=Lib ;
//TSuiviCpta.Post ;
END ;

procedure TFScenario.FListeRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
BEGIN ChargeCompl ; END;

procedure TFScenario.FListeRowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
BEGIN
Radicaux[ou]:=FListe.Cells[1,Ou] ;
MajCompl(ou); ForceChange(Nil) ;
END ;

procedure TFScenario.BImprimerClick(Sender: TObject);
begin
PrintDBGrid(Nil,Nil,Caption,'PRT_SUIVCPTA') ;
end;

procedure TFScenario.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFScenario.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFScenario.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFScenario.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFScenario.VerifDisa ( Tout : boolean ) ;
BEGIN
if ((SC_QUALIFPIECE.Value<>'N') and (SC_QUALIFPIECE.Value<>'')) then
   BEGIN
   SC_VALIDE.Enabled:=False ; if Tout then SC_VALIDE.Checked:=False ;
   SC_LETTRAGESAISIE.Enabled:=False ; if Tout then SC_LETTRAGESAISIE.Checked:=False ;
   END else
   BEGIN
   SC_VALIDE.Enabled:=True ; SC_LETTRAGESAISIE.Enabled:=True ;
   END ;
if ((SC_RIB.Value='') and (Tout)) then SC_RIB.Value:='...' ;
END ;

procedure TFScenario.SSuiviCPTADataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean ;
begin
BNouveau.Enabled:=(Not(TSuiviCpta.State in [dsEdit,dsInsert])) ;
BDelete.Enabled:=(Not(TSuiviCpta.State in [dsEdit,dsInsert])) ;
if EnregVide then BDelete.Enabled:=False ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TSuiviCpta.BOF ;
   DnEnable := Enabled and not TSuiviCpta.EOF ;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable ;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable ;
   if TSuiviCpta.State=dsBrowse then
      BEGIN
      ChargeEnreg ;
      if Mode=0 then VerifDisa(False) ;
      END else
      BEGIN
      if Groupe.ItemIndex=-1 then Groupe.ItemIndex:=0 ;
      if SC_ETABLISSEMENT.ItemIndex=-1 then SC_ETABLISSEMENT.ItemIndex:=0 ;
      if SC_QUALIFPIECE.ItemIndex=-1 then SC_QUALIFPIECE.ItemIndex:=0 ;
      if Mode=0 then VerifDisa(True) ;
      END ;
   END ;
end;

procedure TFScenario.GRKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if (ssCtrl in Shift) and (Key=VK_DELETE) and (TSuiviCpta.Bof and TSuiviCpta.Eof) then BEGIN Key:=0 ; Exit ; END ;
if (ssCtrl in Shift) and (Key=VK_DELETE) then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
end;

procedure TFScenario.ChargeEnreg ;
begin
GeneCharge:=True ; CbFlaguer:=False ;
If (TSuiviCpta.State<>dsInsert) then Cle.Enabled:=False ;
if GROUPE.ItemIndex=-1 then GROUPE.ItemIndex:=0  ;
if SC_ETABLISSEMENT.ItemIndex=-1 then SC_ETABLISSEMENT.ItemIndex:=0  ;
if SC_QUALIFPIECE.ItemIndex=-1 then SC_QUALIFPIECE.ItemIndex:=0  ;
InitTableaux ; JOURNALChange(Nil) ;
Caption:=MsgBox.Mess[14]+SC_ETABLISSEMENT.Value+' '+GROUPE.Value+' '+JOURNAL.Text+' '+NATUREPIECE.Text+' '+SC_QUALIFPIECE.Text ;
GeneCharge:=False ;
UpdateCaption(Self) ;
GereAccesCtrlQte ;
end;

procedure TFScenario.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if Mode=0 then
begin
  FClosing := TRUE;   //SG6 30/11/2004 FQ 15009
  CanClose := OnSauve;
  FClosing := FALSE;  //SG6 30/11/2004 FQ 15009
end;

end;

procedure TFScenario.BNouveauClick(Sender: TObject);
begin Bouge(NbInsert) ; end;

procedure TFScenario.SSuiviCPTAStateChange(Sender: TObject);
begin Modifier:=True ; end ;

procedure TFScenario.OkModifRacines (Sender : TObject) ;
var i,j : integer ;
    Compl,Lib : string ;
    OkModif : boolean ;
BEGIN
MajCompl(Fliste.Row) ;
OkModif:=False ;
for i:=1 to 10 do
  BEGIN
  if Radicaux[i]<>Copy(TSuiviCpta.FindField('SC_RADICAL'+IntToStr(i)).AsString,1,10) then BEGIN OkModif:=True ;Break ; END ;
  Compl:=TSuiviCpta.FindField('SC_COMPLEMENTS'+IntToStr(i)).AsString ;
  For j:= 1 to 10 do
    BEGIN
    if Compl[j]<>Complement[i,j] then BEGIN OkModif:=True ; Break ; END ;
    END ;
  if OkModif then Break ;
  // Zones libres lignes
  Lib:=TSuiviCpta.FindField('SC_COMPLIBRE'+IntToStr(i)).AsString ;
  For j:=1 to 40 do
    BEGIN
    if Lib[j]<>CompLibre[i,j] then BEGIN OkModif:=True ; Break ; END ;
    END ;
  if OkModif then Break ;
  END ;
// Entete libre
Lib:=TSuiviCpta.FindField('SC_LIBREENTETE').AsString ;
for i:=1 to 40 do
  BEGIN
  if Lib[i]<>LibreEntete[i] then BEGIN OkModif:=True ; Break ; END ;
  END ;
if OkModif then if TSuiviCpta.State=dsBrowse then BEGIN TSuiviCpta.Edit ; TSuiviCptaSC_Societe.AsString:=V_PGI.CodeSociete ; END ;
END ;

procedure TFScenario.ForceChange(Sender: TObject);
Var Old : String ;
begin
if GeneCharge then Exit ;
Old:=SC_CONTROLETVA.Value ; SC_CONTROLETVA.Value:='' ; SC_CONTROLETVA.Value:=Old ;
OnClickQte(Sender);
end;

procedure TFScenario.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFScenario.SSuiviCPTAUpdateData(Sender: TObject);
begin
if Not TSuiviCPTA.Modified then Exit ;
if Modifier then
   BEGIN
   Modifier:=False ;
   if Not OnSauve then
     if TSuiviCPTA.Modified then SysUtils.Abort ;
   END ;
end;

procedure TFScenario.SC_QUALIFPIECEChange(Sender: TObject);
begin
VerifDisa(False) ;
end;

Procedure TFScenario.GenereListe(Li : HTStrings) ;
Var i,j : Integer ;
BEGIN
for i:=0 to Entete.ControlCount-1 do
    if Entete.Controls[i] is TCustomCheckBox then Li.Add('SC_LIBREENTETE;'+Entete.Controls[i].Name+';-;X;;;') ;
for i:=1 to 21 do Li.Add('SC_LIBREENTETE;ET'+IntToStr(i)+';-;X;;;') ;
for i:=1 to 10 do
  BEGIN
  for j:=1 to 7 do Li.Add('SC_RADICAL'+IntToStr(i)+';LE'+IntToStr(j+2)+';-;X;;;') ;
  for j:=1 to 21 do Li.Add('SC_RADICAL'+IntToStr(i)+';ZLL'+IntToStr(j)+';-;X;;;') ;
  END ;
END ;

Function TFScenario.AjouteAcces(C : TCustomCheckBox ; MemoSt,StTemp1,Radic : String ; QuelTab : Byte) : String ;
BEGIN
if C.Enabled then
   BEGIN
   MemoSt:=MemoSt+'X;' ;
   if Radic='' then
      BEGIN
      if TCheckBox(C).Checked then BEGIN MemoSt:=MemoSt+'X;' ; CbFlaguer:=True ; END
                              else MemoSt:=MemoSt+'-;' ;
      END else
      BEGIN
      Case QuelTab of
           1 : BEGIN
               if Complement[StrToInt(Radic),TCheckBox(C).Tag]='X' then
                  BEGIN
                  MemoSt:=MemoSt+'X;' ; CbFlaguer:=True ;
                  END else MemoSt:=MemoSt+'-;' ;
               END ;
           2 : BEGIN
               if CompLibre[StrToInt(Radic),TCheckBox(C).Tag]='X' then
                  BEGIN
                  MemoSt:=MemoSt+'X;' ; CbFlaguer:=True ;
                  END else MemoSt:=MemoSt+'-;' ;
               END ;
           End ;
      END ;
   END else MemoSt:=MemoSt+'-;-;' ;
StTemp1:=StTemp1+':'+C.Hint+';' ; MemoSt:=StTemp1+MemoSt ;
Result:=MemoSt ;
END ;

Procedure TFScenario.ChercheAcces(Li : HTStrings) ;
Var i,j : Integer ;
    St,StTemp,StTemp1,St1,MemoSt,Radic : String ;
    C : TCustomCheckBox ;
    C1 : TControl ;
BEGIN
for i:=0 to Li.Count-1 do
  BEGIN
  St:=Li.Strings[i] ; StTemp1:=ReadTokenSt(St) ; MemoSt:=St ; StTemp:=ReadTokenSt(St) ;
  if Pos('SC_',StTemp)=1 then
     BEGIN
     C:=TCustomCheckBox(FindComponent(StTemp)) ;
     if C<>Nil then Li.Strings[i]:=AjouteAcces(C,MemoSt,StTemp1,'',0) ;
     END else
     if Pos('ET',StTemp)=1 then
        BEGIN
        C:=TCustomCheckBox(FindComponent(StTemp)) ;
        if C<>Nil then Li.Strings[i]:=AjouteAcces(C,MemoSt,StTemp1,'',0) ;
        END else
        if (Pos('LE',StTemp)=1) and (Pos('SC_',StTemp)<=0) then
           BEGIN
           St1:='' ;
           for j:=1 to Length(StTemp) do if StTemp[j] in ['0'..'9'] then St1:=St1+StTemp[j] ;
           for j:=0 to Racines.ControlCount-1 do
              BEGIN
              C1:=Racines.Controls[j] ;
              if C1 is TCheckBox then
                 BEGIN
                 if TCheckBox(C1).Tag=StrToInt(St1) then
                    BEGIN
                    if Pos('SC_RADICAL',StTemp1)=1 then Radic:=Copy(StTemp1,11,2)
                                                   else Radic:='' ;
                    Li.Strings[i]:=AjouteAcces(TCusTomCheckBox(C1),MemoSt,StTemp1,Radic,1) ;
                    Break ;
                    END ;
                 END ;
              END ;
           END else
           if Pos('ZLL',StTemp)=1 then
              BEGIN
              St1:='' ;
              for j:=1 to Length(StTemp) do if StTemp[j] in ['0'..'9'] then St1:=St1+StTemp[j] ;
              for j:=0 to Libres.ControlCount-1 do
                  BEGIN
                  C1:=Libres.Controls[j] ;
                  if C1 is TCheckBox then
                     BEGIN
                     if TCheckBox(C1).Tag=StrToInt(St1) then
                        BEGIN
                        if Pos('SC_RADICAL',StTemp1)=1 then Radic:=Copy(StTemp1,11,2)
                                                       else Radic:='' ;
                        Li.Strings[i]:=AjouteAcces(TCusTomCheckBox(C1),MemoSt,StTemp1,Radic,2) ;
                        Break ;
                        END ;
                     END ;
                  END ;
              END  ;
  END ;
END ;

procedure TFScenario.BInfoCompClick(Sender: TObject);
Var InfoComp : HTStrings ;
begin
InfoComp:=HTStringList.Create ;
InfoComp.Assign(TMemoField(TSuiviCPTASC_ATTRCOMP)) ;
if InfoComp.Count<=0 then GenereListe(InfoComp) ;
ChercheAcces(InfoComp) ;
if CbFlaguer then
   BEGIN
   AttributComplementaires(InfoComp) ;
   if TSuiviCPTA.State=dsBrowse then TSuiviCPTA.Edit ;
   TSuiviCPTASC_ATTRCOMP.Assign(InfoComp) ;
   END else MsgBox.Execute(17,'','') ;
InfoComp.Clear ; InfoComp.Free ;
end;

Procedure ValeurAttributComp(Li : HTStrings ; Radical,Champ : String ; Var Oblg,Modif,Valdef,TexteLibre : String ) ;
Var i,j : Integer ;
    St : String ;
BEGIN
Oblg:='-' ; Modif:='X' ; Valdef:='' ; TexteLibre:='' ;
if (Champ='') or (Li=Nil) or (Li.Count<=0) then Exit ;
for i:=0 to Li.Count-1 do
   BEGIN
   St:=Li.Strings[i] ;
   if Pos(Radical,St)<=0 then Continue ;
   if Pos(Champ,St)>0 then
      BEGIN
      for j:=1 to 2 do ReadTokenSt(St) ;
      Oblg:=ReadTokenSt(St) ; Modif:=ReadTokenSt(St) ; ValDef:=ReadTokenSt(St) ;
      TexteLibre:=ReadTokenSt(St) ;
      Break ;
      END ;
   END ;
END ;

procedure TFScenario.BCopierClick(Sender: TObject);
Var Q : TQuery ;
    Num,i : integer ;
    TF  : TField ;
    Nam,CodeN,Lib : String ;
begin
// Numéroter les scénar car par de code unique
Num:=0 ; Lib:='' ;
Q:=OpenSQL('Select SC_REFAUTO FROM SUIVCPTA',False) ;
While Not Q.EOF do
   BEGIN
   Inc(Num) ;
   Q.Edit ; Q.Fields[0].AsString:=IntToStr(Num) ; Q.Post ;
   Q.Next ;
   END ;
Ferme(Q) ;
// Choisir le scénario
CodeN:=Choisir(MsgBox.Mess[18],'SUIVCPTA','SC_JOURNAL || " \ " || SC_NATUREPIECE || " \ " || SC_ETABLISSEMENT','SC_REFAUTO','','') ;
if CodeN='' then Exit ;
if TSuiviCpta.State=dsBrowse then TSuiviCpta.Edit ;
Q:=OpenSQL('Select * from SUIVCPTA Where SC_REFAUTO="'+CodeN+'"',True) ;
if Not Q.EOF then
   BEGIN
   for i:=0 to Q.Fields.Count-1 do
       BEGIN
       Nam:=Q.Fields[i].FieldName ;
       if ((Nam='SC_JOURNAL') or (Nam='SC_NATUREPIECE') or (Nam='SC_ETABLISSEMENT') or (Nam='SC_QUALIFPIECE') or (Nam='SC_USERGRP')) then Continue ;
       TF:=TField(FindComponent('TSuiviCpta'+Nam)) ; if TF=Nil then Continue ;
       if Nam<>'SC_ATTRCOMP' then TF.AsVariant:=Q.Fields[i].AsVariant else
          BEGIN
          if Not TMemoField(Q.Fields[i]).isNull then TMemoField(TF).Assign(TMemoField(Q.Fields[i])) ; 
          END ;
       END ;
   InitTableaux ; 
(*
   Lib:=Q.FindField('SC_LIBREENTETE').AsString ;
   for i:=1 to 10 do
       BEGIN
       TSuiviCpta.FindField('SC_COMPLEMENTS'+IntToStr(i)).AsString:='----------' ;
       TSuiviCpta.FindField('SC_RADICAL'+IntToStr(i)).AsString:='' ;
       FListe.Cells[1,i]:='' ;
       TSuiviCpta.FindField('SC_COMPLIBRE'+IntToStr(i)).AsString:=Lib ;
       END;
*)
   END ;
Ferme(Q) ;
end;

procedure TFScenario.BGenererClick(Sender: TObject);
begin
 EnregOkMulti ;
end;

function TFScenario.EnregOkMulti : Boolean ;
var
 lStJournal,lStNature,lStJal,lStNat : string ;
 lTOB : TOB ;
 i : integer ;
 lJournal : TZListJournal ;
begin
 result:=false ; lJournal:=TZListJournal.Create ;
 if (JOURNALMULTI.Text='') then BEGIN Msgbox.Execute(10,'','') ; Exit ; END ;
 if (NATUREPIECEMULTI.Text='') then BEGIN Msgbox.Execute(11,'','') ; Exit ; END ;
 lTOB:=TOB.create('SUIVCPTA',nil,-1) ; lTOB.AddChampSup('_ETAT',false) ;
 if JOURNALMULTI.Tous then
 begin JOURNALMULTI.text:='' ; for i:=0 to JOURNALMULTI.Values.Count-1 do JOURNALMULTI.text:=JOURNALMULTI.text+JOURNALMULTI.Values[i]+';' ; end ;
 if NATUREPIECEMULTI.Tous then
 begin NATUREPIECEMULTI.text:='' ; for i:=0 to NATUREPIECEMULTI.Values.Count-1 do NATUREPIECEMULTI.text:=NATUREPIECEMULTI.text+NATUREPIECEMULTI.Values[i]+';' ; end ;
 InitMoveProgressForm(self,'Génération en cours...','Génération en cours',10,true,false) ;
 try
 if TSuiviCPTA.State in [dsBrowse] then TSuiviCPTA.Insert ; MajCompl(Fliste.Row) ; SauveCompl ; TOBM(lTOB).ChargeMvt(TSuiviCPTA) ;
 lTOB.PutValue('SC_ATTRCOMP',TSuiviCPTASC_ATTRCOMP.AsString) ;
 lTOB.PutValue('SC_USERGRP','...') ; lTOB.PutValue('SC_ETABLISSEMENT','...') ; lTOB.PutValue('SC_QUALIFPIECE','...') ;
 lStJournal:=JOURNALMULTI.text ;
 lStJal:=readtokenst(lStJournal) ; i:=0 ;
 while lStJal<>'' do
  begin
   lStNature:=NATUREPIECEMULTI.Text ;
   lStNat:=readtokenst(lStNature) ;
   while lStNat<>'' do
    begin
     Inc(i) ;
     if i=10 then i:=1 ;
     lJournal.Load([lStJal]) ;
     if NATUREJALNATPIECEOK (lJournal.GetValue('J_NATUREJAL'),lStNat) then
      BEGIN
			 MoveCurProgressForm('journal : ' + lStJal + ' nature de pièce :' + lStNat) ;
       lTOB.PutValue('SC_JOURNAL',UpperCase(lStJal)) ; lTOB.PutValue('SC_NATUREPIECE', UpperCase(lStNat)) ;
       lTOB.SetAllModifie(true) ; lTOB.ChargeCle1 ; lTOB.DeleteDB(false); lTOB.InsertDB(nil) ;
      END;
     lStNat:=readtokenst(lStNature) ;
    end ; // while
   lstJal:=readtokenst(lStJournal) ;
  end; //
 finally
	FiniMoveProgressForm ;
  FreeAndNil(lTOB) ; FreeAndNil(lJournal) ;
 end;
end;


procedure TFScenario.BFermeClick(Sender: TObject);
begin
  Close;  //SG6 30/11/2004 FQ 15009
  if FClosing and IsInside(Self) then THPanel(parent).CloseInside;
end;

function TFScenario.EstSaisieQte: Boolean;
var i : integer ;
begin
  // Test pour contrôle qtés... décoché si accessible
  result := false ;
  for i := 1 to 10 do
    if ( Complement[i, 7]<>'-' ) or ( Complement[i, 8]<>'-' ) then
      begin
      result := true ;
      exit ;
      end ;
end;

procedure TFScenario.GereAccesCtrlQte;
begin
  if EstSaisieQte then
    SC_CONTROLEQTE.Enabled := True
    else begin
         SC_CONTROLEQTE.Enabled := False ;
         if ( TSuiviCpta.State <> dsBrowse ) and  ( TSuiviCpta.FindField('SC_CONTROLEQTE').AsString = 'X' ) then
           TSuiviCpta.FindField('SC_CONTROLEQTE').AsString := '-' ;
         end ;
end;

procedure TFScenario.OnClickQte(Sender: TObject);
begin
  if not Assigned(sender) then Exit ;
  if not ( ( TCheckBox(Sender).Name = 'LE7' ) or ( TCheckBox(Sender).Name = 'LE8' ) )  then Exit ;
  if TCheckBox(Sender).Checked then
    SC_CONTROLEQTE.Enabled := True ;

end;

END.
