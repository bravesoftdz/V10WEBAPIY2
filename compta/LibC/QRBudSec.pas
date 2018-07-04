{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par PLANBUDSEC_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit QRBudSec;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HQuickRP, StdCtrls, Buttons, ExtCtrls, Mask, Hctrls, ComCtrls, hmsgbox, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Ent1, HCompte, HQry, DBCtrls, Spin, HDB, HEnt1, ParamDat,
  UObjFiltres {SG6 04/01/05 Gestion Filtres V6 FQ 15145},  UtilEDT,MajTable, Menus, HSysMenu, HRichEdt, HRichOLE, HTB97, HPanel, UiUtil ;

procedure PlanBudSec(Axe : String ;UnCpte : String ; DuBouton : Boolean) ;

type
  TFQRBudSec = class(TForm)
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
    TRSens: TQRLabel;
    TRDatemodification: TQRLabel;
    RSens: TQRLabel;
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
    LBS_BUDSECT: TQRDBText;
    BEntetePage: TQRBand;
    RapTitre: TQRSysData;
    RapTitreTrait: TQRShape;
    Pieddepage: TQRBand;
    QRSysData2: TQRSysData;
    RCopyright: TQRLabel;
    QRSysData3: TQRSysData;
    DetailFiche: TQRBand;
    QRShape2: TQRShape;
    TFBS_SENS: TQRLabel;
    FTitreBudget: TQRLabel;
    FTitreCarac: TQRLabel;
    QBudSec: TQuery;
    SBudSec: TDataSource;
    QRP: TQuickReport;
    MsgRien: THMsgBox;
    QRDetailLink1: TQRDetailLink;
    Label9: TLabel;
    Label7: TLabel;
    FSens: THValComboBox;
    FFerme: TCheckBox;
    TLegende: TQRLabel;
    LRFerme: TQRLabel;
    LRSens: TQRLabel;
    RSociete: TQRLabel;
    RUtilisateur: TQRLabel;
    TLRSens: TQRLabel;
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
    LBS_LIBELLE: TQRDBText;
    LBS_FERME: TQRDBText;
    FBS_BUDSECT: TQRDBText;
    TFBS_BUDSECT: TQRLabel;
    LBS_SENS: TQRDBText;
    LBS_DATEMODIFICATION: TQRDBText;
    TFBS_LIBELLE: TQRLabel;
    FBS_LIBELLE: TQRDBText;
    TFBS_HT: TQRLabel;
    FBS_HT: TQRDBText;
    FBS_SENS: TQRDBText;
    TFBS_SIGNE: TQRLabel;
    FBS_SIGNE: TQRDBText;
    TFBS_ATTENTE: TQRLabel;
    FBS_ATTENTE: TQRDBText;
    LBS_HT: TQRDBText;
    FBS_ABREGE: TQRDBText;
    TFBS_ABREGE: TQRLabel;
    FBS_REPORTDISPO: TQRDBText;
    TFBS_REPORTDISPO: TQRLabel;
    TFBS_AXE: TQRLabel;
    FBS_AXE: TQRDBText;
    QBudSecBS_AXE: TStringField;
    QBudSecBS_BUDSECT: TStringField;
    QBudSecBS_LIBELLE: TStringField;
    QBudSecBS_ABREGE: TStringField;
    QBudSecBS_SENS: TStringField;
    QBudSecBS_FERME: TStringField;
    QBudSecBS_DATECREATION: TDateTimeField;
    QBudSecBS_DATEMODIF: TDateTimeField;
    QBudSecBS_DATEOUVERTURE: TDateTimeField;
    QBudSecBS_DATEFERMETURE: TDateTimeField;
    QBudSecBS_REPORTDISPO: TStringField;
    QBudSecBS_BLOCNOTE: TMemoField;
    QBudSecBS_HT: TStringField;
    QBudSecBS_UTILISATEUR: TStringField;
    QBudSecBS_SIGNE: TStringField;
    QBudSecBS_ATTENTE: TStringField;
    QBudSecBS_SECTIONRUB: TStringField;
    QBudSecBS_EXCLURUB: TStringField;
    Label1: TLabel;
    FAxe: THValComboBox;
    TRAxe: TQRLabel;
    RAxe: TQRLabel;
    Rien: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel24: TQRLabel;
    LBS_ATTENTE: TQRDBText;
    FPlusieurs: TCheckBox;
    FCouleur: TCheckBox;
    BComplement: TQRBand;
    QRShape1: TQRShape;
    TFBS_COMPTERUB: TQRLabel;
    QRLabel2: TQRLabel;
    FBS_SECTIONRUB: TQRRichEdit;
    FBS_EXCLURUB: TQRRichEdit;
    FTitreChantier: TQRLabel;
    HBS_SECTIONRUB: THDBRichEditOLE;
    HBS_EXCLURUB: THDBRichEditOLE;
    BInformations: TQRBand;
    FCadreInfo: TQRShape;
    QRLabel8: TQRLabel;
    QRShape6: TQRShape;
    TFBS_DATEMODIF: TQRLabel;
    FBS_DATECREATION: TQRDBText;
    TFBS_DATECREATION: TQRLabel;
    FBS_DATEMODIF: TQRDBText;
    TFBS_DATEOUVERTURE: TQRLabel;
    FBS_DATEOUVERTURE: TQRDBText;
    TFBS_DATEFERMETURE: TQRLabel;
    FBS_DATEFERMETURE: TQRDBText;
    FTitreDates: TQRLabel;
    DetailblocNote: TQRBand;
    FBS_BLOCNOTE: TQRRichEdit;
    TLBS_BLOCNOTE: TQRLabel;
    LBS_BLOCNOTE: TQRRichEdit;
    HBlocNote: THDBRichEditOLE;
    QRShape8: TQRShape;
    Overlay: TQRBand;
    Col8: TQRLigne;
    Col0: TQRLigne;
    Col1: TQRLigne;
    Col2: TQRLigne;
    Col3: TQRLigne;
    Col4: TQRLigne;
    Col5: TQRLigne;
    Col6: TQRLigne;
    Hor2: TQRLigne;
    Hor1: TQRLigne;
    BDCol: TQRBand;
    Hor0: TQRShape;
    TLBS_BUDSECT: TQRLabel;
    TLBS_LIBELLE: TQRLabel;
    TLBS_DATEMODIFICATION: TQRLabel;
    TLBS_SENS: TQRLabel;
    TLBS_HT: TQRLabel;
    TLBS_ATTENTE: TQRLabel;
    QRShape17: TQRShape;
    QRShape3: TQRShape;
    Colonnes: TCheckBox;
    QRShape13: TQRShape;
    QRShape14: TQRShape;
    QRLabel37: TQRLabel;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape7: TQRShape;
    FDateModif1: THCritMaskEdit;
    FDateModif2: THCritMaskEdit;
    BLibres: TQRBand;
    QRShape9: TQRShape;
    FBS_TABLE0: TQRDBText;
    TFBS_TABLE0: TQRLabel;
    QRLabel6: TQRLabel;
    QBudSecBS_TABLE0: TStringField;
    QBudSecBS_TABLE1: TStringField;
    QBudSecBS_TABLE2: TStringField;
    QBudSecBS_TABLE3: TStringField;
    QBudSecBS_TABLE4: TStringField;
    QBudSecBS_TABLE5: TStringField;
    QBudSecBS_TABLE6: TStringField;
    QBudSecBS_TABLE7: TStringField;
    QBudSecBS_TABLE8: TStringField;
    QBudSecBS_TABLE9: TStringField;
    TFBS_TABLE1: TQRLabel;
    FBS_TABLE1: TQRDBText;
    TFBS_TABLE2: TQRLabel;
    FBS_TABLE2: TQRDBText;
    TFBS_TABLE3: TQRLabel;
    FBS_TABLE3: TQRDBText;
    TFBS_TABLE4: TQRLabel;
    FBS_TABLE4: TQRDBText;
    TFBS_TABLE5: TQRLabel;
    FBS_TABLE5: TQRDBText;
    TFBS_TABLE6: TQRLabel;
    FBS_TABLE6: TQRDBText;
    TFBS_TABLE7: TQRLabel;
    FBS_TABLE7: TQRDBText;
    TFBS_TABLE8: TQRLabel;
    FBS_TABLE8: TQRDBText;
    TFBS_TABLE9: TQRLabel;
    FBS_TABLE9: TQRDBText;
    BDFoot: TQRBand;
    BDSummary: TQRBand;
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
    procedure FAxeChange(Sender: TObject);
    procedure ColonnesClick(Sender: TObject);
    procedure CheckBlocnoteClick(Sender: TObject);
    procedure BLibresBeforePrint(var PrintBand: Boolean; var Quoi: String);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    //SG6 04/01/2005 Gestion Filtres V6 FQ 15145
    ObjFiltre : TObjFiltre;
    QuiMappel : Boolean ;
    LeCpte, LAxe : String ;
    Liste, Fiche : Boolean ;
    procedure CASEACOCHERDB(FTex : TQRDBText);
    procedure ActiveDetail ;
    procedure ConstruitSQL ;
    procedure RenseigneCritere ;
    procedure InitDivers ;
  public

  end;


implementation

{$R *.DFM}

procedure PlanBudSec(Axe : String ; UnCpte : String ; DuBouton : Boolean) ;
var FBudSec : TFQRBudSec ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
FBudSec:=TFQRBudSec.Create(Application) ;
FBudSec.QuiMappel:=DuBouton ;
if PP=Nil then
   BEGIN
   try
    if Not DuBouton then FBudSec.ShowModal
                    else BEGIN FBudSec.LAxe:=Axe ; FBudSec.LeCpte:=UnCpte ; FBudSec.BValiderClick(Nil) ; END ;
    finally
    FBudSec.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END Else
   BEGIN
   InitInside(FBudSec,PP) ;
   FBudSec.Show ;
   END ;
END ;


procedure TFQRBudSec.BValiderClick(Sender: TObject);
begin
SourisSablier ;
EnableControls(Self,False) ;
if QuiMappel then
   BEGIN FCompte1.Text:=LeCpte ; FCompte2.Text:=LeCpte ; TypeEtat.ItemIndex:=1 ; END ;
RedimLargeur(Self, QRP,FCouleur.checked, TTitre);
ConstruitSQL ;
if NOT QuiMappel then RenseigneCritere ;
InitDivers ;
ChangeSizeMemo(QBudSecBS_BLOCNOTE) ;
if QuiMappel then QRP.QRPrinter.SynCouleur:=V_PGI.QRCouleur else QRP.QRPrinter.SynCouleur:=FCouleur.Checked ;
ActiveDetail ; QBudsec.Close ; ChangeSQL(QBudsec) ; //QBudsec.Prepare ;
PrepareSQLODBC(QBudSec) ;
QBudsec.Open ;
if QuiMappel then BEGIN EnteteEtat.Enabled:=False ; QRP.ReportTitle:=MsgRien.Mess[1] ; END ;
if QBudsec.Eof then BEGIN MsgRien.Execute(0,'','') END
               else BEGIN if Apercu.Checked then QRP.Preview else QRP.Print ; END ;
QBudsec.Close ;
EnableControls(Self,True) ;
SourisNormale ;
end;


procedure TFQRBudSec.ActiveDetail ;
Var i : Byte ;
BEGIN
Liste:=(TypeEtat.ItemIndex=0) ; Fiche:=(TypeEtat.ItemIndex=1) ;
BDListe.Enabled:=Liste; DetailFiche.Enabled:=Fiche;
BDCol.Enabled:=Liste; BLibres.Enabled:=Fiche ;
Col0.Visible:=Liste ; Col8.Visible:=Liste ;
DetailFiche.ForceNewPage:=not FPlusieurs.Checked ;
BComplement.Enabled:=Fiche ; BInformations.Enabled:=Fiche ;
{ En ce qui Concerne Le Bloc-Note }
DetailblocNote.Enabled:=((CheckBlocnote.Checked and Liste)or Fiche)  ;
TLBS_BLOCNOTE.Visible:=Liste ; LBS_BLOCNOTE.Visible:=Liste ;
FBS_BLOCNOTE.Visible:=Fiche ;
DetailblocNote.Frame.DrawLeft:=Fiche ; DetailblocNote.Frame.DrawRight:=Fiche ;
DetailblocNote.Frame.DrawBottom:=Fiche ;
{ Fin Des Carac. du Bloc-Note }

{ Labels en Liste }
TLBS_BUDSECT.Visible:=Liste ; LBS_BUDSECT.Visible:=Liste ;
TLBS_LIBELLE.Visible:=Liste ; LBS_LIBELLE.Visible:=Liste ;
//TLBS_SIGNE.Visible:=Liste   ; LBS_SIGNE.Visible:=Liste ;
TLBS_SENS.Visible:=Liste    ; LBS_SENS.Visible:=Liste ;
TLBS_DATEMODIFICATION.Visible:=Liste ; LBS_DATEMODIFICATION.Visible:=Liste ;
TLBS_HT.Visible:=Liste      ; LBS_HT.Visible:=Liste ;
TLBS_ATTENTE.Visible:=Liste ; LBS_ATTENTE.Visible:=Liste ;
{ Legendes en Liste }
TLegende.Visible:=Liste ; LRFerme.Visible:=Liste ;
TLRSens.Visible:=Liste ; LRSens.visible:=Liste ;
For i:=1 to 6 do TQRLigne(FindComponent('Col'+IntToStr(i))).Visible:=Colonnes.Checked ;
END ;

procedure TFQRBudSec.ConstruitSQL ;
Var  Traduction : string ;
BEGIN
QBudsec.SQL.Clear ;
QBudsec.SQL.Add('Select BS_AXE, BS_BUDSECT, BS_LIBELLE, BS_ABREGE, BS_SENS, BS_FERME, ' );
QBudsec.SQL.Add('BS_DATECREATION, BS_DATEMODIF, BS_DATEOUVERTURE, BS_DATEFERMETURE, ' );
QBudsec.SQL.Add('BS_REPORTDISPO, BS_BLOCNOTE, BS_HT, BS_UTILISATEUR, BS_SIGNE, ' );
QBudsec.SQL.Add('BS_ATTENTE, BS_SECTIONRUB, BS_EXCLURUB, ' );
QBudsec.SQL.Add('BS_TABLE0, BS_TABLE1, BS_TABLE2, BS_TABLE3, BS_TABLE4, ' );
QBudsec.SQL.Add('BS_TABLE5, BS_TABLE6, BS_TABLE7, BS_TABLE8, BS_TABLE9 ' );
QBudsec.SQL.Add('from BUDSECT ' );
if NOT QuiMappel then
   BEGIN
    QBudsec.SQL.Add('where BS_AXE="'+FAxe.Value+'" ' );
    if FJoker.Visible then
       BEGIN
       QBudsec.SQL.Add(' And BS_BUDSECT like "'+TraduitJoker(FJoker.Text)+'"') ;
       END Else
       BEGIN
       if FCompte1.Text<>'' then QBudsec.SQL.Add(' And BS_BUDSECT>="'+FCompte1.Text+'" ') ;
       if FCompte2.Text<>'' then QBudsec.SQL.Add(' And BS_BUDSECT<="'+FCompte2.Text+'" ') ;
       END ;
    if FSens.ItemIndex<>0 then QBudsec.SQL.Add('and BS_SENS="'+FSens.Value+'" ' );
    if FDateModif1.text<>StDate1900 then QBudsec.SQL.Add('and BS_DATEMODIF>="'+USDATE(FDateModif1)+'" ');
    if FDateModif2.text<>StDate1900 then QBudsec.SQL.Add(' and BS_DATEMODIF<="'+USDATE(FDateModif2)+'" ');
    if FLibelle.Text<>'' then
      BEGIN
      Traduction:=TraduitJoker(FLibelle.text) ;
      QBudsec.SQL.Add('and UPPER(BS_LIBELLE) like "'+Traduction+'" ' );
      END ;
    if FFerme.State=cbChecked then QBudsec.SQL.Add('and BS_FERME="X" ' );
    if FFerme.State=cbUnChecked then QBudsec.SQL.Add('and BS_FERME="-" ' );
    if TriPar.ItemIndex=0 then QBudsec.SQL.Add('Order by BS_AXE, BS_BUDSECT ')
                         else QBudsec.SQL.Add('Order by BS_LIBELLE') ;
   END else
   BEGIN
    QBudsec.SQL.Add('where BS_AXE="'+LAxe+'" ' );
    QBudsec.SQL.Add('and BS_BUDSECT>="'+FCompte1.Text+'"' );
    QBudsec.SQL.Add('and BS_BUDSECT<="'+FCompte2.Text+'" ' );
   END ;
END ;

procedure TFQRBudSec.RenseigneCritere ;
BEGIN
RAxe.Caption:=FAxe.Text ;
RCompte1.Caption:=FCompte1.text ; Rcompte2.Caption:=FCompte2.Text ;
CaseACocher(FFerme,RFerme) ;
RLibelle.Caption:=FLibelle.Text ; RSens.Caption:=FSens.Text ;
RDatemodification1.Caption:=FDateModif1.Text ;
RDatemodification2.Caption:=FDateModif2.Text ;
END ;

procedure TFQRBudSec.InitDivers ;
BEGIN
RSociete.Caption:=Copy(RSociete.Caption,1,Pos(':',RSociete.Caption))
                       +' '+V_PGI.CodeSociete+' '+V_PGI.NomSociete ;
RUtilisateur.Caption:=Copy(RUtilisateur.Caption,1,Pos(':',RUtilisateur.Caption))
                       +' '+V_PGI.User+' '+V_PGI.UserName ;
RNumversion.Caption:=MsgRien.Mess[2]+' '+V_PGI.NumVersion+' '+MsgRien.Mess[3]+' '+DateToStr(V_PGI.DateVersion);
RCopyright.Caption := Copyright + ' - ' + TitreHalley ;
END ;

procedure TFQRBudSec.FormShow(Sender: TObject);
begin
FDateModif1.Text:=StDate1900 ; FDateModif2.Text:=StDate2099 ;
QRP.ReportTitle:=Caption ;
Pages.ActivePage:=CritStandards ;
FSens.ItemIndex:=0 ; FAxe.ItemIndex:=0 ;
FCouleur.Checked:=V_PGI.QRCouleur ;
//SG6 04/01/05 Gestion Filtre V6 FQ 15145
ObjFiltre.Charger;
end;

procedure TFQRBudSec.BEntetePageAfterPrint(BandPrinted: Boolean);
begin
RapTitre.Visible:=True ; RapTitreTrait.Visible:=True ;
end;

procedure TFQRBudSec.EnteteEtatAfterPrint(BandPrinted: Boolean);
begin
RapTitre.Visible:=False ; RapTitreTrait.Visible:=False ;
end;

procedure TFQRBudSec.CASEACOCHERDB(FTex : TQRDBText);
BEGIN
FTex.Font.Name:='WingDings' ; FTex.Font.Size:=10 ;
//if FTex.Caption='X' then FTex.Caption:='þ' else FTex.Caption:='o' ;

if FTex.Caption='-' then BEGIN FTex.Caption:='o' ; FTex.Font.Color:=clBlack ; END;
if FTex.Caption='X' then BEGIN FTex.Caption:='þ' ; FTex.Font.Color:=clBlack ; END;
//if FBox.State=cbGrayed then BEGIN RBox.Caption:='n' ; RBox.Font.Color:=clGray ; END;
END ;

procedure TFQRBudSec.BDlisteBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
if LBS_FERME.Caption='X' then LBS_FERME.Caption:=MsgRien.Mess[4] else LBS_FERME.Caption:='';
//If QBudsecBS_SIGNE.AsString<>'' then LBS_SIGNE.Caption:=RechDom(ttRubSigne,QBudsecBS_SIGNE.AsString,False) ;
CASEACOCHERDB(LBS_HT) ;
CASEACOCHERDB(LBS_ATTENTE) ;
end;

procedure TFQRBudSec.DetailFicheBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
FTitreBudget.Caption:=Copy(FTitreBudget.Caption,1,Pos(':',FTitreBudget.Caption))
                       +' '+FBS_BUDSECT.Caption+' '+FBS_LIBELLE.Caption ;
//CASEACOCHERDB(FBS_REPORTDISPO) ;
CASEACOCHERDB(FBS_HT) ;
CASEACOCHERDB(FBS_ATTENTE) ;
//FBS_SENS.Caption:=FSens.Text ;
//FBS_AXE.Caption:=FAxe.Text ;
//If QBudsecBS_SIGNE.AsString<>'' then FBS_SIGNE.Caption:=RechDom(ttRubSigne,QBudsecBS_SIGNE.AsString,False) ;
FBS_BLOCNOTE.Visible:=True ;
if (FSens.ItemIndex<=0) or not liste then FBS_SENS.Caption:=RechDom('ttSens',QBudsecBS_SENS.AsString,False) else FBS_SENS.Caption:=RSens.Caption ;
if (FAxe.ItemIndex<=0) or not liste then FBS_AXE.Caption:=RechDom('ttAxe',QBudsecBS_AXE.AsString,False) else FBS_AXE.Caption:=RAxe.Caption ;
end;


procedure TFQRBudSec.LDetailblocNoteBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
PrintBand:=Not (QBudsecBS_BLOCNOTE.isNull) ;
end;

procedure TFQRBudSec.TypeEtatClick(Sender: TObject);
begin
If TypeEtat.itemIndex=1 then BEGIN CheckBlocnote.Checked:=TRUE ; CheckBlocnote.Enabled:=FALSE ; FPlusieurs.Checked:=False ; FPlusieurs.Enabled:=True ; END
                        else BEGIN CheckBlocnote.Checked:=False ;CheckBlocnote.Enabled:=TRUE ; FPlusieurs.Checked:=False ; FPlusieurs.Enabled:=False ; END ;
Colonnes.Checked:= FALSE ; Colonnes.Enabled:= Not(TypeEtat.ItemIndex=1) ;
end;

procedure TFQRBudSec.DetailblocNoteBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
if Liste then PrintBand:=Not (QBudsecBS_BLOCNOTE.isNull) ;
end;

procedure TFQRBudSec.FDateModif1KeyPress(Sender: TObject; var Key: Char);
begin
ParamDate(Self,Sender,Key) ;
end;

procedure TFQRBudSec.FCompte1Change(Sender: TObject);
var AvecJoker : boolean ;
begin
AvecJoker:=Joker(FCompte1, FCompte2, FJoker) ;
TFCompte2.Visible:=Not AvecJoker ;
TFCompte1.Visible:=Not AvecJoker ;
TFJoker.Visible:=AvecJoker ;
end;

procedure TFQRBudSec.FAxeChange(Sender: TObject);
begin
If FAxe.Value='' then Exit ;
Case Char(FAxe.Value[2]) of
 '1' : BEGIN FCompte1.ZoomTable:=tzBudSec1 ; FCompte2.ZoomTable:=FCompte1.ZoomTable ; END ;
 '2' : BEGIN FCompte1.ZoomTable:=tzBudSec2 ; FCompte2.ZoomTable:=FCompte1.ZoomTable ; END ;
 '3' : BEGIN FCompte1.ZoomTable:=tzBudSec3 ; FCompte2.ZoomTable:=FCompte1.ZoomTable ; END ;
 '4' : BEGIN FCompte1.ZoomTable:=tzBudSec4 ; FCompte2.ZoomTable:=FCompte1.ZoomTable ; END ;
 '5' : BEGIN FCompte1.ZoomTable:=tzBudSec5 ; FCompte2.ZoomTable:=FCompte1.ZoomTable ; END ;
 End ;
end;

procedure TFQRBudSec.ColonnesClick(Sender: TObject);
begin
if CheckBLOCNote.Checked then Colonnes.Checked :=FALSE ;
end;

procedure TFQRBudSec.CheckBlocnoteClick(Sender: TObject);
begin
if CheckBlocNote.Checked then Colonnes.Checked:=FALSE ;
end;

procedure TFQRBudSec.BLibresBeforePrint(var PrintBand: Boolean; var Quoi: String);
Var LesLib : HTStringList ;
    i : Integer ;
    Cap,St : String ;
    Trouver : Boolean ;
BEGIN
LesLib:=HTStringList.Create ; GetLibelleTableLibre('D',LesLib) ; Trouver:=False ;
for i:=0 to LesLib.Count-1 do
    BEGIN
    St:=LesLib.Strings[i] ;
    Cap:=ReadTokenSt(St) ;
    TQRLabel(FindComponent('TFBS_TABLE'+IntToStr(i))).Caption:=Cap ;
    if St='X' then Trouver:=True ;
    END ;
PrintBand:=Trouver ;
LesLib.Free ;
end;

procedure TFQRBudSec.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//SG6 04/01/05 Gestiopn Filtres V6 FQ 15145
ObjFiltre.Charger;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFQRBudSec.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
  //SG6 04/01/2005 Gestion Filtres V6 FQ 15145
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, 'QRBUDSEC');

HorzScrollBar.Range:=0 ; HorzScrollBar.Visible:=FALSE ;
VertScrollBar.Range:=0 ; VertScrollBar.Visible:=FALSE ;
Pages.Top:=0 ; Pages.Left:=0 ;
ClientHeight:=Pages.Top+Pages.Height+HPB.Height+PanelFiltre.Height-1 ;
ClientWidth:=Pages.Left+Pages.Width ;
if V_PGI.OutLook then Pages.Align:=alClient ;
end;

procedure TFQRBudSec.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFQRBudSec.BFermeClick(Sender: TObject);
begin
//SG6 04/01/05 Vide le panel
  Close;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

end.
