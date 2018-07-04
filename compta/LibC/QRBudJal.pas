{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par PLANBUDJAL_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit QRBudJal;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HQuickRP, StdCtrls, Buttons, ExtCtrls, Mask, Hctrls, ComCtrls, hmsgbox, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Ent1, HCompte, HQry, DBCtrls, Spin, HDB, HEnt1, ParamDat,
  UObjFiltres {SG6 04/01/05 Gestion Filtres V6 QF 15145}, UtilEDT,MajTable,
  Menus, HSysMenu, HRichEdt, HRichOLE, HTB97, HPanel, UiUtil ;

procedure PlanBudJal(Axe : String ;UnCpte : String ; DuBouton : Boolean) ;

type
  TFQRBudJal = class(TForm)
    Pages: TPageControl;
    CritStandards: TTabSheet;
    FLibelle: TEdit;
    OptionsEditions: TTabSheet;
    TypeEtat: TRadioGroup;
    Apercu: TCheckBox;
    OptionsType: TGroupBox;
    CheckBlocnote: TCheckBox;
    TriPar: TRadioGroup;
    Panel: TPanel;
    EnteteEtat: TQRBand;
    TTitre: TQRSysData;
    TRFerme: TQRLabel;
    RFerme: TQRLabel;
    TRDatemodification: TQRLabel;
    RDatemodification1: TQRLabel;
    QRLabel1: TQRLabel;
    RCompte1: TQRLabel;
    QRLabel3: TQRLabel;
    RCompte2: TQRLabel;
    TRLibelle: TQRLabel;
    RLibelle: TQRLabel;
    QRLabel5: TQRLabel;
    RDatemodification2: TQRLabel;
    BDliste: TQRBand;
    LBJ_BUDJAL: TQRDBText;
    BEntetePage: TQRBand;
    RapTitre: TQRSysData;
    RapTitreTrait: TQRShape;
    Pieddepage: TQRBand;
    QRSysData2: TQRSysData;
    RCopyright: TQRLabel;
    QRSysData3: TQRSysData;
    DetailFiche: TQRBand;
    QRShape2: TQRShape;
    FTitreBudget: TQRLabel;
    FTitreCarac: TQRLabel;
    QBudJal: TQuery;
    SBudJal: TDataSource;
    QRP: TQuickReport;
    MsgRien: THMsgBox;
    QRDetailLink1: TQRDetailLink;
    Label9: TLabel;
    FFerme: TCheckBox;
    TLegende: TQRLabel;
    LRFerme: TQRLabel;
    RSociete: TQRLabel;
    RUtilisateur: TQRLabel;
    HLabel3: THLabel;
    HLabel6: THLabel;
    RNumversion: TQRLabel;
    HMTrad: THSystemMenu;
    TFJoker: THLabel;
    TFCompte1: THLabel;
    TFCompte2: TLabel;
    FJoker: TEdit;
    FCompte1: THCpteEdit;
    FCompte2: THCpteEdit;
    LBJ_LIBELLE: TQRDBText;
    LBJ_FERME: TQRDBText;
    FBJ_BUDJAL: TQRDBText;
    TFBJ_BUDJAL: TQRLabel;
    LBJ_DATEMODIFICATION: TQRDBText;
    TFBJ_LIBELLE: TQRLabel;
    FBJ_LIBELLE: TQRDBText;
    FBJ_ABREGE: TQRDBText;
    TFBJ_ABREGE: TQRLabel;
    TFBJ_AXE: TQRLabel;
    FBJ_AXE: TQRDBText;
    Label1: TLabel;
    FAxe: THValComboBox;
    TRAxe: TQRLabel;
    RAxe: TQRLabel;
    FPlusieurs: TCheckBox;
    FCouleur: TCheckBox;
    BComplement: TQRBand;
    QRShape1: TQRShape;
    TFBS_COMPTERUB: TQRLabel;
    QRLabel2: TQRLabel;
    FBJ_BUDGENES: TQRRichEdit;
    FBJ_BUDSECTS: TQRRichEdit;
    FTitreChantier: TQRLabel;
    HBJ_BUDGENES: THDBRichEditOLE;
    HBJ_BUDSECTS: THDBRichEditOLE;
    BInformations: TQRBand;
    FCadreInfo: TQRShape;
    QRLabel8: TQRLabel;
    QRShape6: TQRShape;
    TFBJ_DATEMODIF: TQRLabel;
    FBJ_DATECREATION: TQRDBText;
    TFBJ_DATECREATION: TQRLabel;
    FBJ_DATEMODIF: TQRDBText;
    TFBJ_DATEOUVERTURE: TQRLabel;
    FBJ_DATEOUVERTURE: TQRDBText;
    TFBJ_DATEFERMETURE: TQRLabel;
    FBJ_DATEFERMETURE: TQRDBText;
    FTitreDates: TQRLabel;
    TFBJ_COMPTEURNORMAL: TQRLabel;
    FBJ_COMPTEURNORMAL: TQRDBText;
    TFBJ_COMPTEURSIMUL: TQRLabel;
    FBJ_COMPTEURSIMUL: TQRDBText;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRLabel9: TQRLabel;
    TFBJ_GENEATTENTE: TQRLabel;
    FBJ_GENEATTENTE: TQRDBText;
    TFBJ_SECTATTENTE: TQRLabel;
    FBJ_SECTATTENTE: TQRDBText;
    QRShape5: TQRShape;
    QRShape7: TQRShape;
    QRLabel12: TQRLabel;
    TFBJ_EXODEB: TQRLabel;
    FBJ_EXODEB: TQRDBText;
    TFBJ_EXOFIN: TQRLabel;
    FBJ_EXOFIN: TQRDBText;
    TFBJ_PERDEB: TQRLabel;
    FBJ_PERDEB: TQRDBText;
    TFBJ_PERFIN: TQRLabel;
    FBJ_PERFIN: TQRDBText;
    QBudJalBJ_AXE: TStringField;
    QBudJalBJ_BUDJAL: TStringField;
    QBudJalBJ_LIBELLE: TStringField;
    QBudJalBJ_ABREGE: TStringField;
    QBudJalBJ_FERME: TStringField;
    QBudJalBJ_DATECREATION: TDateTimeField;
    QBudJalBJ_DATEMODIF: TDateTimeField;
    QBudJalBJ_DATEOUVERTURE: TDateTimeField;
    QBudJalBJ_DATEFERMETURE: TDateTimeField;
    QBudJalBJ_BLOCNOTE: TMemoField;
    QBudJalBJ_UTILISATEUR: TStringField;
    QBudJalBJ_COMPTEURNORMAL: TStringField;
    QBudJalBJ_COMPTEURSIMUL: TStringField;
    QBudJalBJ_GENEATTENTE: TStringField;
    QBudJalBJ_SECTATTENTE: TStringField;
    QBudJalBJ_EXODEB: TStringField;
    QBudJalBJ_EXOFIN: TStringField;
    QBudJalBJ_PERDEB: TDateTimeField;
    QBudJalBJ_PERFIN: TDateTimeField;
    QRShape8: TQRShape;
    QRShape17: TQRShape;
    LBJ_COMPTEURNORMAL: TQRDBText;
    LBJ_COMPTEURSIMUL: TQRDBText;
    LBJ_GENEATTENTE: TQRDBText;
    LBJ_SECTATTENTE: TQRDBText;
    BDCol: TQRBand;
    Hor0: TQRShape;
    QRShape9: TQRShape;
    TLBJ_BUDJAL: TQRLabel;
    TLBJ_LIBELLE: TQRLabel;
    TLBJ_DATEMODIFICATION: TQRLabel;
    TLBJ_COMPTEURNORMAL: TQRLabel;
    TLBJ_COMPTEURSIMUL: TQRLabel;
    TLBJ_GENEATTENTE: TQRLabel;
    TLBJ_SECTATTENTE: TQRLabel;
    Colonnes: TCheckBox;
    QRShape14: TQRShape;
    QRShape13: TQRShape;
    QRLabel37: TQRLabel;
    QRShape12: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QBudJalBJ_BUDGENES: TMemoField;
    QBudJalBJ_BUDSECTS: TMemoField;
    FDateModif1: THCritMaskEdit;
    FDateModif2: THCritMaskEdit;
    TBS_SENS: TQRLabel;
    FBJ_SENS: TQRDBText;
    QBudJalBJ_SENS: TStringField;
    QBudJalBJ_NATJAL: TStringField;
    QRLabel4: TQRLabel;
    FBJ_NATJAL: TQRDBText;
    HBlocNote: THDBRichEditOLE;
    FSens: THValComboBox;
    Label2: TLabel;
    RSens: TQRLabel;
    TRSens: TQRLabel;
    DetailblocNote: TQRBand;
    TLBJ_BLOCNOTE: TQRLabel;
    FBJ_BLOCNOTE: TQRRichEdit;
    LBJ_BLOCNOTE: TQRRichEdit;
    BDFoot: TQRBand;
    Overlay: TQRBand;
    Col8: TQRLigne;
    Col0: TQRLigne;
    Col3: TQRLigne;
    Col1: TQRLigne;
    Col2: TQRLigne;
    Col4: TQRLigne;
    Col5: TQRLigne;
    Col6: TQRLigne;
    Col7: TQRLigne;
    Hor2: TQRLigne;
    Hor1: TQRLigne;
    BDSummary: TQRBand;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Dock971: TDock97;
    PanelFiltre: TToolWindow97;
    HPB: TToolWindow97;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BEntetePageAfterPrint(BandPrinted: Boolean);
    procedure EnteteEtatAfterPrint(BandPrinted: Boolean);
    procedure BDlisteBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DetailFicheBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure LDetailblocNoteBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TypeEtatClick(Sender: TObject);
    procedure DetailblocNoteBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FDateModif1KeyPress(Sender: TObject; var Key: Char);
    procedure FCompte1Change(Sender: TObject);
    procedure ColonnesClick(Sender: TObject);
    procedure CheckBlocnoteClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    //SG6 04/01/05 Gestion Filtres V6 FQ 15145
    ObjFiltre : TObjFiltre;
    QuiMappel : Boolean ;
    LeCpte, LAxe : String ;
    Liste, Fiche : Boolean ;
    procedure ActiveDetail ;
    procedure ConstruitSQL ;
    procedure RenseigneCritere ;
    procedure InitDivers ;
  public

  end;


implementation

{$R *.DFM}

procedure PlanBudJal(Axe : String ; UnCpte : String ; DuBouton : Boolean) ;
var FBudJal : TFQRBudJal ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
FBudJal:=TFQRBudJal.Create(Application) ;
FBudJal.QuiMappel:=DuBouton ;
if PP=Nil then
   BEGIN
   try
    if Not DuBouton then FBudJal.ShowModal
                    else BEGIN FBudJal.LAxe:=Axe ; FBudJal.LeCpte:=UnCpte ; FBudJal.BValiderClick(Nil) ; END ;
    finally
    FBudJal.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END Else
   BEGIN
   InitInside(FBudJal,PP) ;
   FBudJal.Show ;
   END ;
END ;


procedure TFQRBudJal.BValiderClick(Sender: TObject);
begin
SourisSablier ;
EnableControls(Self,False) ;
if QuiMappel then
   BEGIN FCompte1.Text:=LeCpte ; FCompte2.Text:=LeCpte ; TypeEtat.ItemIndex:=1 ; END ;
RedimLargeur(Self, QRP,FCouleur.checked, TTitre);
ConstruitSQL ;
if NOT QuiMappel then RenseigneCritere ;
InitDivers ;
ChangeSizeMemo(QBudJalBJ_BLOCNOTE) ;
ChangeSizeLongChar(QBudjalBJ_BUDGENES) ;
ChangeSizeLongChar(QBudjalBJ_BUDSECTS) ;
if QuiMappel then QRP.QRPrinter.SynCouleur:=V_PGI.QRCouleur else QRP.QRPrinter.SynCouleur:=FCouleur.Checked ;
ActiveDetail ; QBudJal.Close ; ChangeSQL(QBudJal) ; //QBudJal.Prepare ;
PrepareSQLODBC(QBudJal) ;
QBudJal.Open ;
if QuiMappel then BEGIN EnteteEtat.Enabled:=False ; QRP.ReportTitle:=MsgRien.Mess[1] ; END ;
if QBudJal.Eof then BEGIN MsgRien.Execute(0,'','') END
               else BEGIN if Apercu.Checked then QRP.Preview else QRP.Print ; END ;
QBudJal.Close ;
EnableControls(Self,True) ;
SourisNormale ;
end;


procedure TFQRBudJal.ActiveDetail ;
Var i : Byte ;
BEGIN
Liste:=(TypeEtat.ItemIndex=0) ; Fiche:=(TypeEtat.ItemIndex=1) ;
BDListe.Enabled:=Liste; DetailFiche.Enabled:=Fiche;
BDCol.Enabled:=Liste; //Overlay.Enabled:=Liste ;
Col0.Visible:=Liste ; Col8.Visible:=Liste ;
DetailFiche.ForceNewPage:=not FPlusieurs.Checked ;
BComplement.Enabled:=Fiche ;
BInformations.Enabled:=Fiche ;
{ En ce qui Concerne Le Bloc-Note }
DetailblocNote.Enabled:=((CheckBlocnote.Checked and Liste)or Fiche)  ;
TLBJ_BLOCNOTE.Visible:=Liste ; LBJ_BLOCNOTE.Visible:=Liste ;
FBJ_BLOCNOTE.Visible:=Fiche ;
DetailblocNote.Frame.DrawLeft:=Fiche ; DetailblocNote.Frame.DrawRight:=Fiche ;
DetailblocNote.Frame.DrawBottom:=Fiche ;
{ Fin Des Carac. du Bloc-Note }

{ Labels en Liste }
TLBJ_BUDJAL.Visible:=Liste ; LBJ_BUDJAL.Visible:=Liste ;
TLBJ_LIBELLE.Visible:=Liste ; LBJ_LIBELLE.Visible:=Liste ;
TLBJ_COMPTEURNORMAL.Visible:=Liste   ; LBJ_COMPTEURNORMAL.Visible:=Liste ;
TLBJ_COMPTEURSIMUL.Visible:=Liste    ; LBJ_COMPTEURSIMUL.Visible:=Liste ;
TLBJ_DATEMODIFICATION.Visible:=Liste ; LBJ_DATEMODIFICATION.Visible:=Liste ;
TLBJ_GENEATTENTE.Visible:=Liste      ; LBJ_GENEATTENTE.Visible:=Liste ;
TLBJ_SECTATTENTE.Visible:=Liste      ; LBJ_SECTATTENTE.Visible:=Liste ;
{ Legendes en Liste }
TLegende.Visible:=Liste ; LRFerme.Visible:=Liste ;
For i:=1 to 7 do TQRLigne(FindComponent('Col'+IntToStr(i))).Visible:=Colonnes.Checked ;
END ;

procedure TFQRBudJal.ConstruitSQL ;
Var  Traduction : string ;
BEGIN
QBudJal.SQL.Clear ;
QBudJal.SQL.Add('select BJ_AXE, BJ_BUDJAL, BJ_LIBELLE, BJ_ABREGE, BJ_FERME, ' );
QBudJal.SQL.Add('BJ_DATECREATION, BJ_DATEMODIF, BJ_DATEOUVERTURE, BJ_DATEFERMETURE, ' );
QBudJal.SQL.Add('BJ_BLOCNOTE, BJ_UTILISATEUR, BJ_COMPTEURNORMAL, BJ_COMPTEURSIMUL, ' );
QBudJal.SQL.Add('BJ_GENEATTENTE, BJ_SECTATTENTE, BJ_EXODEB, BJ_EXOFIN, BJ_PERDEB, BJ_PERFIN, ' );
QBudJal.SQL.Add('BJ_BUDGENES, BJ_BUDSECTS, BJ_SENS, BJ_NATJAL ');
QBudJal.SQL.Add('from BUDJAL ' );
if NOT QuiMappel then
   BEGIN
    QBudJal.SQL.Add('where BJ_AXE="'+FAxe.Value+'" ' );
    if FJoker.Visible then
       BEGIN
       QBudJal.SQL.Add(' And BJ_BUDJAL like "'+TraduitJoker(FJoker.Text)+'"') ;
       END Else
       BEGIN
       if FCompte1.Text<>'' then QBudJal.SQL.Add(' And BJ_BUDJAL>="'+FCompte1.Text+'" ') ;
       if FCompte2.Text<>'' then QBudJal.SQL.Add(' And BJ_BUDJAL<="'+FCompte2.Text+'" ') ;
       END ;
    if FSens.ItemIndex>0 then QBudJal.SQL.Add('and BJ_SENS="'+FSens.Value+'" ' );
    if FDateModif1.text<>StDate1900 then QBudJal.SQL.Add('and BJ_DATEMODIF>="'+USDATE(FDateModif1)+'" and BJ_DATEMODIF<="'+USDATE(FDateModif2)+'" ');
    if FLibelle.Text<>'' then
      BEGIN
      Traduction:=TraduitJoker(FLibelle.text) ;
      QBudJal.SQL.Add('and UPPER(BJ_LIBELLE) like "'+Traduction+'" ' );
      END ;
    if FFerme.State=cbChecked then QBudJal.SQL.Add('and BJ_FERME="X" ' );
    if FFerme.State=cbUnChecked then QBudJal.SQL.Add('and BJ_FERME="-" ' );
    if TriPar.ItemIndex=0 then QBudJal.SQL.Add('Order by BJ_AXE, BJ_BUDJAL ')
                          else QBudJal.SQL.Add('Order by BJ_LIBELLE') ;
   END else
   BEGIN
   QBudJal.SQL.Add('where BJ_AXE="'+LAxe+'" ' );
   QBudJal.SQL.Add('and BJ_BUDJAL>="'+FCompte1.Text+'"' );
   QBudJal.SQL.Add('and BJ_BUDJAL<="'+FCompte2.Text+'" ' );
   END ;
END ;

procedure TFQRBudJal.RenseigneCritere ;
BEGIN
RAxe.Caption:=FAxe.Text ; RSens.Caption:=FSens.Text ;
RCompte1.Caption:=FCompte1.text ; Rcompte2.Caption:=FCompte2.Text ;
CaseACocher(FFerme,RFerme) ;
RLibelle.Caption:=FLibelle.Text ;
RDatemodification1.Caption:=FDateModif1.Text ;
RDatemodification2.Caption:=FDateModif2.Text ;
END ;

procedure TFQRBudJal.InitDivers ;
BEGIN
RSociete.Caption:=Copy(RSociete.Caption,1,Pos(':',RSociete.Caption))
                       +' '+V_PGI.CodeSociete+' '+V_PGI.NomSociete ;
RUtilisateur.Caption:=Copy(RUtilisateur.Caption,1,Pos(':',RUtilisateur.Caption))
                       +' '+V_PGI.User+' '+V_PGI.UserName ;
RNumversion.Caption:=MsgRien.Mess[2]+' '+V_PGI.NumVersion+' '+MsgRien.Mess[3]+' '+DateToStr(V_PGI.DateVersion);
RCopyright.Caption := Copyright + ' - ' + TitreHalley ;
END ;

procedure TFQRBudJal.FormShow(Sender: TObject);
begin
FDateModif1.Text:=StDate1900 ; FDateModif2.Text:=StDate2099 ;
QRP.ReportTitle:=Caption ;
Pages.ActivePage:=CritStandards ;
FAxe.ItemIndex:=0 ; FSens.ItemIndex:=0 ;
FCouleur.Checked:=V_PGI.QRCouleur ;
//Sg6 04/01/05 Gestion Filtres V6 FQ 15145
ObjFiltre.Charger;
end;

procedure TFQRBudJal.BEntetePageAfterPrint(BandPrinted: Boolean);
begin
RapTitre.Visible:=True ; RapTitreTrait.Visible:=True ;
end;

procedure TFQRBudJal.EnteteEtatAfterPrint(BandPrinted: Boolean);
begin
RapTitre.Visible:=False ; RapTitreTrait.Visible:=False ;
end;

procedure TFQRBudJal.BDlisteBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
if LBJ_FERME.Caption='X' then LBJ_FERME.Caption:=MsgRien.Mess[4] else LBJ_FERME.Caption:='';
end;

procedure TFQRBudJal.DetailFicheBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
FTitreBudget.Caption:=Copy(FTitreBudget.Caption,1,Pos(':',FTitreBudget.Caption))
                       +' '+FBJ_BUDJAL.Caption+' '+FBJ_LIBELLE.Caption ;
//FBJ_SENS.Caption:=FSens.Text ;
//FBG_SIGNE.Caption:=RechDom(ttRubSigne,QBudgetBG_SIGNE.AsString,False) ;
FBJ_BLOCNOTE.Visible:=True ;
if (FSens.ItemIndex<=0) or not liste then FBJ_SENS.Caption:=RechDom('ttSens',QBudJalBJ_SENS.AsString,False) else FBJ_SENS.Caption:=RSens.Caption ;
if (FAxe.ItemIndex<=0) or not liste then FBJ_AXE.Caption:=RechDom('ttAxe',QBudJalBJ_AXE.AsString,False) else FBJ_AXE.Caption:=FAxe.Text ;
FBJ_NATJAL.Caption:=RechDom('ttNatJalBud',QBudJalBJ_NATJAL.AsString,False) ;
FBJ_EXODEB.Caption:=RechDom('ttExercice',QBudJalBJ_EXODEB.AsString,False) ;
FBJ_EXOFIN.Caption:=RechDom('ttExercice',QBudJalBJ_EXOFIN.AsString,False) ;
end;


procedure TFQRBudJal.LDetailblocNoteBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
PrintBand:=Not (QBudJalBJ_BLOCNOTE.isNull) ;
end;

procedure TFQRBudJal.TypeEtatClick(Sender: TObject);
begin
If TypeEtat.itemIndex=1 then BEGIN CheckBlocnote.Checked:=TRUE ; CheckBlocnote.Enabled:=FALSE ; FPlusieurs.Checked:=False ; FPlusieurs.Enabled:=True ; END
                        else BEGIN CheckBlocnote.Checked:=False ;CheckBlocnote.Enabled:=TRUE ; FPlusieurs.Checked:=False ; FPlusieurs.Enabled:=False ; END ;
Colonnes.Checked:= FALSE ; Colonnes.Enabled:= Not(TypeEtat.ItemIndex=1) ;
end;

procedure TFQRBudJal.DetailblocNoteBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
DetailBlocNote.Frame.DrawTop:=False ;
if Liste then PrintBand:=Not (QBudJalBJ_BLOCNOTE.isNull) ;
end;

procedure TFQRBudJal.FDateModif1KeyPress(Sender: TObject; var Key: Char);
begin
ParamDate(Self,Sender,Key) ;
end;

procedure TFQRBudJal.FCompte1Change(Sender: TObject);
var AvecJoker : boolean ;
begin
AvecJoker:=Joker(FCompte1, FCompte2, FJoker) ;
TFCompte2.Visible:=Not AvecJoker ;
TFCompte1.Visible:=Not AvecJoker ;
TFJoker.Visible:=AvecJoker ;
end;
procedure TFQRBudJal.ColonnesClick(Sender: TObject);
begin
if CheckBLOCNote.Checked then Colonnes.Checked :=FALSE ;
end;

procedure TFQRBudJal.CheckBlocnoteClick(Sender: TObject);
begin
if CheckBlocNote.Checked then Colonnes.Checked:=FALSE ;
end;

procedure TFQRBudJal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//SG6 04/01/05 Gestion Filtres V6 FQ 15145
ObjFiltre.Free;
if Parent is THPanel then Action:=caFree;
end;

procedure TFQRBudJal.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
  //SG6 04/01/2005 Gestion Filtres V6 FQ 15145
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, 'QRBUDJAL');

  HorzScrollBar.Range:=0 ; HorzScrollBar.Visible:=FALSE ;
VertScrollBar.Range:=0 ; VertScrollBar.Visible:=FALSE ;
Pages.Top:=0 ; Pages.Left:=0 ;
ClientHeight:=Pages.Top+Pages.Height+HPB.Height+PanelFiltre.Height-1 ;
ClientWidth:=Pages.Left+Pages.Width ;
if V_PGI.OutLook then Pages.Align:=alClient ;
end;

procedure TFQRBudJal.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFQRBudJal.BFermeClick(Sender: TObject);
begin
  //SG6 04/01/05 Vide le panel;
  Close;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

end.
