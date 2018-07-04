{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par PLANBUDGET_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit QRBudGen;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HQuickRP, StdCtrls, Buttons, ExtCtrls, Mask, Hctrls, ComCtrls, hmsgbox, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Ent1, HCompte, HQry, DBCtrls, Spin, HDB, HEnt1, ParamDat,
  UObjFiltres {SG6 04/01/2005 Gestion Filtres V6}, UtilEDT,MajTable, Menus, HSysMenu, HRichEdt, HRichOLE, HTB97, HPanel, UiUtil ;

procedure PlanBudget(UnCpte : String ; DuBouton : Boolean) ;

type
  TFQRBudGen = class(TForm)
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
    BDListe: TQRBand;
    LBG_BUDGENE: TQRDBText;
    BEntetePage: TQRBand;
    RapTitre: TQRSysData;
    RapTitreTrait: TQRShape;
    Pieddepage: TQRBand;
    QRSysData2: TQRSysData;
    RCopyright: TQRLabel;
    QRSysData3: TQRSysData;
    DetailFiche: TQRBand;
    QRShape2: TQRShape;
    TBS_SENS: TQRLabel;
    FTitreBudget: TQRLabel;
    FTitreCarac: TQRLabel;
    QBudget: TQuery;
    SBudget: TDataSource;
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
    QBudgetBG_BUDGENE: TStringField;
    QBudgetBG_LIBELLE: TStringField;
    QBudgetBG_ABREGE: TStringField;
    QBudgetBG_SENS: TStringField;
    QBudgetBG_FERME: TStringField;
    QBudgetBG_DATECREATION: TDateTimeField;
    QBudgetBG_DATEMODIF: TDateTimeField;
    QBudgetBG_DATEOUVERTURE: TDateTimeField;
    QBudgetBG_DATEFERMETURE: TDateTimeField;
    QBudgetBG_REPORTDISPO: TStringField;
    QBudgetBG_BLOCNOTE: TMemoField;
    QBudgetBG_HT: TStringField;
    QBudgetBG_UTILISATEUR: TStringField;
    LBG_LIBELLE: TQRDBText;
    LBG_FERME: TQRDBText;
    FBG_BUDGENE: TQRDBText;
    TFBG_BUDGENE: TQRLabel;
    LBG_SENS: TQRDBText;
    LBG_DATEMODIFICATION: TQRDBText;
    TFBG_LIBELLE: TQRLabel;
    FBG_LIBELLE: TQRDBText;
    TFBG_HT: TQRLabel;
    FBG_HT: TQRDBText;
    FBG_SENS: TQRDBText;
    TFBG_SIGNE: TQRLabel;
    FBG_SIGNE: TQRDBText;
    QBudgetBG_SIGNE: TStringField;
    QBudgetBG_ATTENTE: TStringField;
    FBG_ATTENTE: TQRDBText;
    LBG_HT: TQRDBText;
    TLBG_HT: TQRLabel;
    QBudgetBG_COMPTERUB: TStringField;
    QBudgetBG_EXCLURUB: TStringField;
    LBG_ATTENTE: TQRDBText;
    FCouleur: TCheckBox;
    FPlusieurs: TCheckBox;
    BComplement: TQRBand;
    QRShape7: TQRShape;
    FBG_EXCLURUB: TQRRichEdit;
    QRLabel2: TQRLabel;
    FTitreOptionImp: TQRLabel;
    FBG_COMPTERUB: TQRRichEdit;
    FTitreChantier: TQRLabel;
    HBG_COMPTERUB: THDBRichEditOLE;
    HBG_EXCLURUB: THDBRichEditOLE;
    BInformations: TQRBand;
    QRShape6: TQRShape;
    QRLabel8: TQRLabel;
    FTitreDates: TQRLabel;
    TFBG_DATECREATION: TQRLabel;
    FBG_DATECREATION: TQRDBText;
    TFBG_DATEMODIFICATION: TQRLabel;
    FBG_DATEMODIFICATION: TQRDBText;
    TFBG_DATEOUVERTURE: TQRLabel;
    FBG_DATEOUVERTURE: TQRDBText;
    TFBG_DATEFERMETURE: TQRLabel;
    FBG_DATEFERMETURE: TQRDBText;
    FCadreInfo: TQRShape;
    BDCol: TQRBand;
    Hor0: TQRShape;
    TLBG_BUDGENE: TQRLabel;
    TLBG_LIBELLE: TQRLabel;
    TLBG_DATEMODIFICATION: TQRLabel;
    TLBG_SENS: TQRLabel;
    TLBS_HT: TQRLabel;
    TLBG_ATTENTE: TQRLabel;
    QRShape17: TQRShape;
    QRShape1: TQRShape;
    Colonnes: TCheckBox;
    QRShape13: TQRShape;
    QRLabel37: TQRLabel;
    QRShape14: TQRShape;
    QRShape12: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    FDateModif1: THCritMaskEdit;
    FDateModif2: THCritMaskEdit;
    QBudgetBG_TABLE0: TStringField;
    QBudgetBG_TABLE1: TStringField;
    QBudgetBG_TABLE2: TStringField;
    QBudgetBG_TABLE3: TStringField;
    QBudgetBG_TABLE4: TStringField;
    QBudgetBG_TABLE5: TStringField;
    QBudgetBG_TABLE6: TStringField;
    QBudgetBG_TABLE7: TStringField;
    QBudgetBG_TABLE8: TStringField;
    QBudgetBG_TABLE9: TStringField;
    BDFoot: TQRBand;
    QRShape4: TQRShape;
    FBG_RUB: TQRDBText;
    QRLabel4: TQRLabel;
    QRLabel7: TQRLabel;
    QBudgetBG_RUB: TStringField;
    BDSummary: TQRBand;
    DetailblocNote: TQRBand;
    TLBG_BLOCNOTE: TQRLabel;
    FBG_BLOCNOTE: TQRRichEdit;
    LBG_BLOCNOTE: TQRRichEdit;
    HBlocNote: THDBRichEditOLE;
    BLibres: TQRBand;
    QRShape9: TQRShape;
    FBG_TABLE0: TQRDBText;
    TFBG_TABLE0: TQRLabel;
    QRLabel6: TQRLabel;
    TFBG_TABLE1: TQRLabel;
    FBG_TABLE1: TQRDBText;
    TFBG_TABLE2: TQRLabel;
    FBG_TABLE2: TQRDBText;
    TFBG_TABLE3: TQRLabel;
    FBG_TABLE3: TQRDBText;
    TFBG_TABLE4: TQRLabel;
    FBG_TABLE4: TQRDBText;
    TFBG_TABLE5: TQRLabel;
    FBG_TABLE5: TQRDBText;
    TFBG_TABLE6: TQRLabel;
    FBG_TABLE6: TQRDBText;
    TFBG_TABLE7: TQRLabel;
    FBG_TABLE7: TQRDBText;
    TFBG_TABLE8: TQRLabel;
    FBG_TABLE8: TQRDBText;
    TFBG_TABLE9: TQRLabel;
    FBG_TABLE9: TQRDBText;
    Overlay: TQRBand;
    Col0: TQRLigne;
    Col8: TQRLigne;
    Col1: TQRLigne;
    Col2: TQRLigne;
    Col3: TQRLigne;
    Col4: TQRLigne;
    Col5: TQRLigne;
    Col6: TQRLigne;
    Hor1: TQRLigne;
    Hor2: TQRLigne;
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
//    procedure BSaveFiltreClick(Sender: TObject);
    procedure BEntetePageAfterPrint(BandPrinted: Boolean);
    procedure EnteteEtatAfterPrint(BandPrinted: Boolean);
    procedure BDListeBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DetailFicheBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure LDetailblocNoteBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TypeEtatClick(Sender: TObject);
    procedure DetailblocNoteBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FDateModif1KeyPress(Sender: TObject; var Key: Char);
  //  procedure FFiltresChange(Sender: TObject);
//    procedure BDelFiltreClick(Sender: TObject);
//    procedure BCreerFiltreClick(Sender: TObject);
//    procedure BRenFiltreClick(Sender: TObject);
//    procedure BNouvRechClick(Sender: TObject);
    procedure FCompte1Change(Sender: TObject);
    procedure ColonnesClick(Sender: TObject);
    procedure CheckBlocnoteClick(Sender: TObject);
  //  procedure POPFPopup(Sender: TObject);
    procedure BLibresBeforePrint(var PrintBand: Boolean; var Quoi: String);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    //SG6 04/01/2005 Gestion Filtres V6 FQ 15145
    ObjFiltre : TObjFiltre;
    QuiMappel : Boolean ;
    LeCpte : String ;
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

procedure PlanBudget(UnCpte : String ; DuBouton : Boolean) ;
var FBudget : TFQRBudgen ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
FBudget:=TFQRBudgen.Create(Application) ;
FBudget.QuiMappel:=DuBouton ;
if PP=Nil then
   BEGIN
   try
    if Not DuBouton then FBudget.ShowModal
                    else BEGIN FBudget.LeCpte:=UnCpte  ; FBudget.BValiderClick(Nil) ; END ;
    finally
    FBudget.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END Else
   BEGIN
   InitInside(FBudget,PP) ;
   FBudget.Show ;
   END ;
END ;


procedure TFQRBudGen.BValiderClick(Sender: TObject);
begin
SourisSablier ;
EnableControls(Self,False) ;
if QuiMappel then
   BEGIN FCompte1.Text:=LeCpte ; FCompte2.Text:=LeCpte ; TypeEtat.ItemIndex:=1 ; END ;
RedimLargeur(Self, QRP,FCouleur.checked, TTitre);
ConstruitSQL ;
if NOT QuiMappel then RenseigneCritere ;
InitDivers ;
ChangeSizeMemo(QBudgetBG_BLOCNOTE) ;
if QuiMappel then QRP.QRPrinter.SynCouleur:=V_PGI.QRCouleur else QRP.QRPrinter.SynCouleur:=FCouleur.Checked ;
ActiveDetail ; QBudget.Close ; ChangeSQL(QBudget) ; //QBudget.Prepare ;
PrepareSQLODBC(QBudget) ;
QBudget.Open ;
if QuiMappel then BEGIN EnteteEtat.Enabled:=False ; QRP.ReportTitle:=MsgRien.Mess[1] ; END ;
if QBudget.Eof then BEGIN MsgRien.Execute(0,'','') END
               else BEGIN if Apercu.Checked then QRP.Preview else QRP.Print ; END ;
QBudget.Close ;
EnableControls(Self,True) ;
SourisNormale ;
end;


procedure TFQRBudGen.ActiveDetail ;
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
TLBG_BLOCNOTE.Visible:=Liste ; LBG_BLOCNOTE.Visible:=Liste ;
//TFBG_BLOCNOTE.Visible:=Fiche ;
FBG_BLOCNOTE.Visible:=Fiche ;

DetailblocNote.Frame.DrawLeft:=Fiche ; DetailblocNote.Frame.DrawRight:=Fiche ;
DetailblocNote.Frame.DrawBottom:=Fiche ;

//TraitBlocnote1.Visible:=Fiche ; TraitBlocnote2.Visible:=Fiche ;
{ Fin Des Carac. du Bloc-Note }

{ Labels en Liste }
TLBG_BUDGENE.Visible:=Liste ; LBG_BUDGENE.Visible:=Liste ;
TLBG_LIBELLE.Visible:=Liste ; LBG_LIBELLE.Visible:=Liste ;
//TLBG_SIGNE.Visible:=Liste   ; LBG_SIGNE.Visible:=Liste ;
TLBG_SENS.Visible:=Liste    ; LBG_SENS.Visible:=Liste ;
TLBG_DATEMODIFICATION.Visible:=Liste ; LBG_DATEMODIFICATION.Visible:=Liste ;
TLBG_HT.Visible:=Liste      ; LBG_HT.Visible:=Liste ;
TLBG_ATTENTE.Visible:=Liste ; LBG_ATTENTE.Visible:=Liste ;
{ Legendes en Liste }
TLegende.Visible:=Liste ; LRFerme.Visible:=Liste ;
TLRSens.Visible:=Liste ; LRSens.visible:=Liste ;

{ Colonnages }
For i:=1 to 6 do TQRLigne(FindComponent('Col'+IntToStr(i))).Visible:=Colonnes.Checked ;
END ;

procedure TFQRBudGen.ConstruitSQL ;
Var  Traduction : string ;
BEGIN
QBudget.SQL.Clear ;
QBudget.SQL.Add('Select BG_BUDGENE, BG_LIBELLE, BG_ABREGE, BG_SENS, BG_FERME, ' );
QBudget.SQL.Add('BG_DATECREATION, BG_DATEMODIF, BG_DATEOUVERTURE, BG_DATEFERMETURE, ' );
QBudget.SQL.Add('BG_REPORTDISPO, BG_BLOCNOTE, BG_HT, BG_UTILISATEUR, BG_SIGNE, ' );
QBudget.SQL.Add('BG_ATTENTE, BG_COMPTERUB, BG_EXCLURUB, BG_RUB, ' );
QBudget.SQL.Add('BG_TABLE0, BG_TABLE1, BG_TABLE2, BG_TABLE3, BG_TABLE4, ' );
QBudget.SQL.Add('BG_TABLE5, BG_TABLE6, BG_TABLE7, BG_TABLE8, BG_TABLE9 ' );
QBudget.SQL.Add('from BUDGENE ' );
if NOT QuiMappel then
   BEGIN
    QBudget.SQL.Add('where BG_BUDGENE <> "'+w_w+'" ' );
    if FJoker.Visible then
       BEGIN
       QBudget.SQL.Add(' And BG_BUDGENE like "'+TraduitJoker(FJoker.Text)+'"') ;
       END Else
       BEGIN
       if FCompte1.Text<>'' then QBudget.SQL.Add(' And BG_BUDGENE>="'+FCompte1.Text+'" ') ;
       if FCompte2.Text<>'' then QBudget.SQL.Add(' And BG_BUDGENE<="'+FCompte2.Text+'" ') ;
       END ;
    if FSens.ItemIndex<>0 then QBudget.SQL.Add('and BG_SENS="'+FSens.Value+'" ' );
    if FDateModif1.text<>StDate1900 then QBudget.SQL.Add(' and BG_DATEMODIF>="'+USDATE(FDateModif1)+'" ');
    if FDateModif2.text<>StDate1900 then QBudget.SQL.Add(' and BG_DATEMODIF<="'+USDATE(FDateModif2)+'" ');
    if FLibelle.Text<>'' then
      BEGIN
      Traduction:=TraduitJoker(FLibelle.text) ;
      QBudget.SQL.Add('and UPPER(BG_LIBELLE) like "'+Traduction+'" ' );
      END ;
    if FFerme.State=cbChecked then QBudget.SQL.Add('and BG_FERME="X" ' );
    if FFerme.State=cbUnChecked then QBudget.SQL.Add('and BG_FERME="-" ' );
    if TriPar.ItemIndex=0 then QBudget.SQL.Add('Order by BG_BUDGENE ')
                         else QBudget.SQL.Add('Order by BG_LIBELLE ') ;
   END else
   BEGIN
    QBudget.SQL.Add('where ' );
    QBudget.SQL.Add('BG_BUDGENE>="'+FCompte1.Text+'"' );
    QBudget.SQL.Add('and BG_BUDGENE<="'+FCompte2.Text+'" ' );
   END ;
END ;

procedure TFQRBudGen.RenseigneCritere ;
BEGIN
RCompte1.Caption:=FCompte1.text ; Rcompte2.Caption:=FCompte2.Text ;
CaseACocher(FFerme,RFerme) ;
RLibelle.Caption:=FLibelle.Text ; RSens.Caption:=FSens.Text ;
RDatemodification1.Caption:=FDateModif1.Text ;
RDatemodification2.Caption:=FDateModif2.Text ;
END ;

procedure TFQRBudGen.InitDivers ;
BEGIN
RSociete.Caption:=Copy(RSociete.Caption,1,Pos(':',RSociete.Caption))
                       +' '+V_PGI.CodeSociete+' '+V_PGI.NomSociete ;
RUtilisateur.Caption:=Copy(RUtilisateur.Caption,1,Pos(':',RUtilisateur.Caption))
                       +' '+V_PGI.User+' '+V_PGI.UserName ;
RNumversion.Caption:=MsgRien.Mess[2]+' '+V_PGI.NumVersion+' '+MsgRien.Mess[3]+' '+DateToStr(V_PGI.DateVersion);
RCopyright.Caption := Copyright + ' - ' + TitreHalley ;
END ;

procedure TFQRBudGen.FormShow(Sender: TObject);
begin
FDateModif1.Text:=StDate1900 ; FDateModif2.Text:=StDate2099 ;
QRP.ReportTitle:=Caption ;
Pages.ActivePage:=CritStandards ;
FSens.ItemIndex:=0 ;
FCouleur.Checked:=V_PGI.QRCouleur ;

//SG6 04/01/2005 Gestion filtres V6 FQ 15145
ObjFiltre.Charger;
end;

procedure TFQRBudGen.BEntetePageAfterPrint(BandPrinted: Boolean);
begin
RapTitre.Visible:=True ; RapTitreTrait.Visible:=True ;
end;

procedure TFQRBudGen.EnteteEtatAfterPrint(BandPrinted: Boolean);
begin
RapTitre.Visible:=False ; RapTitreTrait.Visible:=False ;
end;

procedure TFQRBudGen.CASEACOCHERDB(FTex : TQRDBText);
BEGIN
FTex.Font.Name:='WingDings' ; FTex.Font.Size:=10 ;
//if FTex.Caption='X' then FTex.Caption:='þ' else FTex.Caption:='o' ;

if FTex.Caption='-' then BEGIN FTex.Caption:='o' ; FTex.Font.Color:=clBlack ; END;
if FTex.Caption='X' then BEGIN FTex.Caption:='þ' ; FTex.Font.Color:=clBlack ; END;
//if FBox.State=cbGrayed then BEGIN RBox.Caption:='n' ; RBox.Font.Color:=clGray ; END;
END ;

procedure TFQRBudGen.BDListeBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
if LBG_FERME.Caption='X' then LBG_FERME.Caption:=MsgRien.Mess[4] else LBG_FERME.Caption:='';
//LBG_SIGNE.Caption:=RechDom(ttRubSigne,QBudgetBG_SIGNE.AsString,False) ;
CASEACOCHERDB(LBG_HT) ;
CASEACOCHERDB(LBG_ATTENTE) ;
end;

procedure TFQRBudGen.DetailFicheBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
FTitreBudget.Caption:=Copy(FTitreBudget.Caption,1,Pos(':',FTitreBudget.Caption))
                       +' '+FBG_BUDGENE.Caption+' '+FBG_LIBELLE.Caption ;
//CASEACOCHERDB(FBG_REPORTDISPO) ;
CASEACOCHERDB(FBG_HT) ;
CASEACOCHERDB(FBG_ATTENTE) ;
FBG_SENS.Caption:=FSens.Text ;
//FBG_SIGNE.Caption:=RechDom(ttRubSigne,QBudgetBG_SIGNE.AsString,False) ;
FBG_BLOCNOTE.Visible:=True ;
if FSens.ItemIndex<=0 then FBG_SENS.Caption:=RechDom('ttSens',QBudgetBG_SENS.AsString,False) else FBG_SENS.Caption:=RSens.Caption ;
end;


procedure TFQRBudGen.LDetailblocNoteBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
PrintBand:=Not (QBudgetBG_BLOCNOTE.isNull) ;
end;

procedure TFQRBudGen.TypeEtatClick(Sender: TObject);
begin
If TypeEtat.itemIndex=1 then BEGIN CheckBlocnote.Checked:=TRUE ; CheckBlocnote.Enabled:=FALSE ; FPlusieurs.Checked:=False ; FPlusieurs.Enabled:=True ; END
                        else BEGIN CheckBlocnote.Checked:=False ;CheckBlocnote.Enabled:=TRUE ; FPlusieurs.Checked:=False ; FPlusieurs.Enabled:=False ; END ;
Colonnes.Checked:= FALSE ; Colonnes.Enabled:= Not(TypeEtat.ItemIndex=1) ;
end;

procedure TFQRBudGen.DetailblocNoteBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
DetailBlocNote.Frame.DrawTop:=False ;
if Liste then PrintBand:=Not (QBudgetBG_BLOCNOTE.isNull) ;
end;

procedure TFQRBudGen.FDateModif1KeyPress(Sender: TObject; var Key: Char);
begin
ParamDate(Self,Sender,Key) ;
end;

procedure TFQRBudGen.FCompte1Change(Sender: TObject);
var AvecJoker : boolean ;
begin
AvecJoker:=Joker(FCompte1, FCompte2, FJoker) ;
TFCompte2.Visible:=Not AvecJoker ;
TFCompte1.Visible:=Not AvecJoker ;
TFJoker.Visible:=AvecJoker ;
end;
procedure TFQRBudGen.ColonnesClick(Sender: TObject);
begin
if CheckBLOCNote.Checked then Colonnes.Checked :=FALSE ;
end;

procedure TFQRBudGen.CheckBlocnoteClick(Sender: TObject);
begin
if CheckBlocNote.Checked then Colonnes.Checked:=FALSE ;
end;

procedure TFQRBudGen.BLibresBeforePrint(var PrintBand: Boolean; var Quoi: String);
Var LesLib : HTStringList ;
    i : Integer ;
    Cap,St : String ;
    Trouver : Boolean ;
BEGIN
LesLib:=HTStringList.Create ; GetLibelleTableLibre('BG',LesLib) ; Trouver:=False ;
for i:=0 to LesLib.Count-1 do
    BEGIN
    St:=LesLib.Strings[i] ;
    Cap:=ReadTokenSt(St) ;
    TQRLabel(FindComponent('TFBG_TABLE'+IntToStr(i))).Caption:=Cap ;
    if St='X' then Trouver:=True ;
    END ;
PrintBand:=Trouver ;
LesLib.Free ;
end;

procedure TFQRBudGen.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFQRBudGen.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//SG6 04/01/2005 Gestion filtre V6 FQ 15145
ObjFiltre.Free;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFQRBudGen.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
  //SG6 04/01/2005 Gestion Filtres V6 FQ 15145
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, 'QRBUDGEN');

HorzScrollBar.Range:=0 ; HorzScrollBar.Visible:=FALSE ;
VertScrollBar.Range:=0 ; VertScrollBar.Visible:=FALSE ;
Pages.Top:=0 ; Pages.Left:=0 ;
ClientHeight:=Pages.Top+Pages.Height+HPB.Height+PanelFiltre.Height-1 ;
ClientWidth:=Pages.Left+Pages.Width ;
if V_PGI.OutLook then Pages.Align:=alClient ;
end;

procedure TFQRBudGen.BFermeClick(Sender: TObject);
begin
  //SG6 04/01/05 Vide le panel
  Close;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

end.
