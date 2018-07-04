unit PointageNew;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DB, DBTables, Grids, DBGrids, ExtCtrls, ComCtrls, StdCtrls, Buttons, ParamDat,
  Hctrls, Hcompte, Mask, Hqry, DBCtrls, hmsgbox, Ent1,  HEnt1, LettUtil,
  SaisUtil, SaisComm, Choix, Menus, HDebug, HDB, Filtre,
{$IFNDEF CCS3}
  eSaisieTr,QRPointg,
{$ENDIF}
  HSysMenu,CritEdt,UtilEdt, HTB97, HPanel
  ,UtilSais //CUpdateCumulsPointeMS
  ,UiUtil ;

Procedure PointageCompteNew (Ok : Boolean) ;
Procedure ZoomPointageCompteNew (CptZoom : String ; Ok : Boolean) ;

Type TTotCptInitiaux = Array[TSorteMontant,0..2] of Record { 0 : NonPointe , 1 : Pointé , 2 : Total }
                                                    Deb,Cre : Double ;
                                                    End ;

type
  TFPointageNew = class(TForm)
    Pages: TPageControl;
    PComplement: TTabSheet;
    PStandard: TTabSheet;
    SGene: TDataSource;
    QGene: TQuery;
    QGeneG_GENERAL: TStringField;
    QGeneG_LIBELLE: TStringField;
    FPied: TPanel;
    Messages: THMsgBox;
    FDetail: TPanel;
    QDevise: TQuery;
    SDevise: TDataSource;
    FListe: THGrid;
    HPB: TToolWindow97;
    Label2: TLabel;
    FSession: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    FNbP: THNumEdit;
    FDebitPA: THNumEdit;
    FCreditPA: THNumEdit;
    FSoldeA: THNumEdit;
    Label13: TLabel;
    FindDialog: TFindDialog;
    QGeneBQ_DEVISE: TStringField;
    QGeneG_TOTDEBP: TFloatField;
    QGeneG_TOTCREP: TFloatField;
    FEntete: TPanel;
    Label8: TLabel;
    Label16: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label23: TLabel;
    FAvantPointage: TLabel;
    FDebitDPIniA: THNumEdit;
    FCreditDPIniA: THNumEdit;
    FDebitPIniA: THNumEdit;
    FCreditPIniA: THNumEdit;
    FSoldePIniA: THNumEdit;
    FSoldeIniA: THNumEdit;
    Label1: TLabel;
    E_NATUREPIECE: TLabel;
    E_DATEECHEANCE: TLabel;
    Label5: TLabel;
    Label17: TLabel;
    E_NUMEROPIECE: TLabel;
    E_MODEPAIE: TLabel;
    Label6: TLabel;
    E_LIBELLE: TLabel;
    Label26: TLabel;
    E_REFEXTERNE: TLabel;
    Label29: TLabel;
    E_REFLIBRE: TLabel;
    Label7: TLabel;
    E_JOURNAL: TLabel;
    QGeneG_DEBNONPOINTE: TFloatField;
    QGeneG_CREDNONPOINTE: TFloatField;
    POPS: TPopupMenu;
    TGen: THLabel;
    FGene: THCpteEdit;
    DBNav: TDBNavigator;
    TDateArrete: THLabel;
    FDateCptaFin: THCritMaskEdit;
    Bevel1: TBevel;
    Label3: TLabel;
    FMontant1: TEdit;
    FMontant2: TEdit;
    Label4: TLabel;
    HLabel10: THLabel;
    FDateEche1: THCritMaskEdit;
    HLabel7: THLabel;
    FDateEche2: THCritMaskEdit;
    Label14: TLabel;
    FModePaie: THValComboBox;
    Bevel2: TBevel;
    POPZ: TPopupMenu;
    BAgrandir: TToolbarButton97;
    BReduire: TToolbarButton97;
    BChercher: TToolbarButton97;
    BComplement: TToolbarButton97;
    BInfoPointage: TToolbarButton97;
    BLanceSaisie: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    BAppliquer: TToolbarButton97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    Bevel3: TBevel;
    QGeneG_NATUREGENE: TStringField;
    FAffichDev: TLabel;
    BZoomPiece: TToolbarButton97;
    BZoom: TToolbarButton97;
    BEtatRappro: TToolbarButton97;
    BEtatPointage: TToolbarButton97;
    BEnleve: TToolbarButton97;
    BRajoute: TToolbarButton97;
    BPointe: TToolbarButton97;
    HMTrad: THSystemMenu;
    HLabel12: THLabel;
    HLabel15: THLabel;
    FDPDepoint: THNumEdit;
    FRestePointeCur: THNumEdit;
    HLabel1: THLabel;
    FDateCptaDeb: THCritMaskEdit;
    BParam: TToolbarButton97;
    HLabel4: THLabel;
    FDateCpta1: THCritMaskEdit;
    HLabel5: THLabel;
    FDateCpta2: THCritMaskEdit;
    FDevise: TDBEdit;
    FNatureGene: TDBEdit;
    Dock: TDock97;
    PParams: TTabSheet;
    Bevel4: TBevel;
    TRefPointage: THLabel;
    FRefPointage: TEdit;
    BRefP: TToolbarButton97;
    HLabel3: THLabel;
    FDebitPointagePivot: TEdit;
    HLabel13: THLabel;
    FCreditPointagePivot: TEdit;
    HLabel14: THLabel;
    HLabel2: THLabel;
    FDatePointage: THCritMaskEdit;
    HDiv: THMsgBox;
    FDebitPointageEuro: TEdit;
    HLabel6: THLabel;
    FCreditPointageEuro: TEdit;
    HLabel8: THLabel;
    HLabelDevise: THLabel;
    iEuroEuro: TImage;
    QGeneG_TOTDEBPTP: TFloatField;
    QGeneG_TOTCREPTP: TFloatField;
    QGeneG_TOTDEBPTD: TFloatField;
    QGeneG_TOTCREPTD: TFloatField;
    TFChoixDevise: THLabel;
    FChoixDevise: THValComboBox;
    ISigneEuro: TImage;
    ISigneEuro1: TImage;
    QGeneJ_JOURNAL: TStringField;
    QGeneJ_LIBELLE: TStringField;
    fJournal: THCpteEdit;
    procedure BAbandonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BAppliquerClick(Sender: TObject);
    procedure BComplementClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FDebitPointagePivotExit(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BPointeClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BLanceSaisieClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FindDialogFind(Sender: TObject);
    procedure SGeneDataChange(Sender: TObject; Field: TField);
    procedure FListeDblClick(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure FDateEche1KeyPress(Sender: TObject; var Key: Char);
    procedure FRefPointageExit(Sender: TObject);
    procedure BInfoPointageClick(Sender: TObject);
    procedure FListeSelectCell(Sender: TObject; Col, Row: Longint; var CanSelect: Boolean);
    procedure BZoomPieceClick(Sender: TObject);
    procedure FGeneEnter(Sender: TObject);
    procedure FGeneExit(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure BRefPClick(Sender: TObject);
    procedure BEtatRapproClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BEnleveClick(Sender: TObject);
    procedure BRajouteClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BZoomClick(Sender: TObject);
    procedure BEtatPointageClick(Sender: TObject);
    procedure DBNavClick(Sender: TObject; Button: TNavigateBtn);
    procedure FListeMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure FRefPointageKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure BAideClick(Sender: TObject);
    procedure BParamClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
//    procedure CAfficheOpposeClick(Sender: TObject);
    procedure FChoixDeviseChange(Sender: TObject);
    procedure FGeneChange(Sender: TObject);
    procedure FListeEnter(Sender: TObject);
    procedure FChoixDeviseEnter(Sender: TObject);
  private    { Déclarations privées }
    OkPointe, FirstFind    : Boolean ;
    DebitCpBanquePivot,CreditCpBanquePivot : Double ;
    DebitCpBanqueEuro,CreditCpBanqueEuro : Double ; {Monnaie Opposée}
    XDebitP,XCreditP,XDebitD,XCreditD : double ; { Totaux pointés sur la référence en cours }
    DPDebitP,DPCreditP,DPDebitD,DPCreditD,DPDebitE,DPCreditE : double ;{ Totaux dépointés en dépointage }
    DecimP,DecimE          : integer ;
    Symbole                : string ;
    CRefP,CDateP      : integer;
    GX,GY                  : integer ;
    Modif, ListeVide       : boolean;
    LesTotauxCpt             : TTotCptInitiaux ;
    { Totaux pour MAJ des comptes }
    MAJCredPivot,MAJDebPivot,{MAJCredEuro,MAJDebEuro,}MAJCredDevise,MAJDebDevise : double;
    DateRef                : TDateTime;
    TuPeuxSupprimer,OkRef  : boolean ;
    Crit                   : TCritEdt ;
    CptZoom                : String ;
    WMinX,WMinY : Integer ;
    OkMajTotPDev : Boolean ; //False si devise selectionnée différente de la devise du compte
    CptEnPivot : Boolean ; // True si compte de banque tenu en pivot
    CptBqe : Boolean ;
    RestitutionEnDevise : Boolean ; // True si la devise de restitution est ni le pivot ni le fongible
    BAppliquerAFaire    : Boolean ;
    BValiderEnCours     : Boolean ;
    OldGen              : String ;
    OkZoom              : Boolean ;
    gbJournal           : Boolean ;
    gszJournal          : String;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    procedure PointageFichier ;
    procedure CalculeTotaux ;
    procedure ChercheEcritures ;
    procedure PointeLes(R :integer) ;
    procedure DepointeLes(R :integer) ;
    procedure EnvoiePointage(R :integer) ;
    procedure EnvoieDepointage(R :integer) ;
//    procedure RecupDevise(var Decim :integer; var Symbole:string) ;
//    procedure CalculRestePointe;
    procedure CalculDepointe(R:integer; Plus:boolean);
    procedure RemplitGrid(Q:TQuery);
    Function  AttribCol(Champ: string) : integer;
    procedure RemplitCellules(Q :TQuery) ;
    procedure SauvePointage;
    procedure CalculPointes(RefP:String);
    procedure RenseigneDetail (Row: LongInt);
    procedure ZoomPiece(R : LongInt) ;
    procedure MajTotPointeCpt;
    procedure EnregRefPointage;
    procedure RepositionGene;
    Function  BalanceSaisie(var M : RMVT ) : Boolean ;
    procedure RecupereSaisie( M : RMVT );
    function  WhereEcr(OkSaisie: boolean) : string ;
    procedure SauveReseau;
//    function  WhereLastModif: string;
    procedure MajSessionMemoire ;
    procedure ZoomCompte;
    procedure ClickEtat(Lequel : string);
    procedure ClickValide ;
    procedure PointeDepointe ;
    procedure CachePointes(OkOk : boolean);
    procedure GrandPetitListe(OkOk : boolean);
    procedure ClickRecherche ;
    procedure ClickInfoPointage (OkOk : boolean) ;
    procedure ClickSaisie ;
    procedure ClickChercher ;
    function  ChercheJal(Gene:string): string;
    procedure CloseFen ;
    procedure AfficheTotauxBasEcran ;
    procedure AfficheTotauxPointeHautDEcran ;
    procedure AfficheTotalDepointe ;
    procedure ChargeInfosCompte ;
    Function  InitPourDevise : Boolean ;
    Procedure RetoucheTitreListe;
    procedure MajTotalPourUnPointage(OBM : TOBM) ;
  public
  end;

implementation

uses ParamDBG, Saisie, QRRappro, PrintDBG,
    CPGeneraux_TOM,
{$IFNDEF IMP}
  SaisBor,
{$ENDIF}
{$IFNDEF CCS3}
{$ENDIF}
  Extraibq,UtilPGI;

{
Règles de calcul des totaux avant pointage :

Total non pointés :
Somme(Debit) non pointés depuis le debut de l'encours hors a-nouveaux
}

{$R *.DFM}

Procedure PointageCompteNew (Ok : Boolean) ;
var FPointage: TFPointageNew;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture'],False,'nrPointage') then Exit ;
FPointage:=TFPointageNew.Create(Application) ;
FPointage.OkPointe:=Ok;
FPointage.gbJournal:=VH^.PointageJal ;
FPointage.OkZoom:=FALSE ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FPointage.ShowModal ;
    Finally
     FPointage.Free ;
     _Bloqueur('nrPointage',False) ;
    end ;
   Screen.Cursor:=crDefault ;
   END else
   BEGIN
   InitInside(FPointage,PP) ;
   FPointage.Show ;
   END ;
END ;

Procedure ZoomPointageCompteNew (CptZoom : String ; Ok : Boolean) ;
var FPointage: TFPointageNew;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture'],False,'nrPointage') then Exit ;
FPointage:=TFPointageNew.Create(Application) ;
FPointage.OkPointe:=Ok;
FPointage.gbJournal:=VH^.PointageJal ;
FPointage.OkZoom:=TRUE ;
If Trim(CptZoom)='' Then FPointage.OkZoom:=FALSE ;
FPointage.CptZoom:=CptZoom ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FPointage.ShowModal ;
    Finally
     FPointage.Free ;
     _Bloqueur('nrPointage',False) ;
    end ;
   Screen.Cursor:=crDefault ;
   END else
   BEGIN
   InitInside(FPointage,PP) ;
   FPointage.Show ;
   END ;
END ;


procedure TFPointageNew.BAgrandirClick(Sender: TObject);
begin
GrandPetitListe(true);
end;

procedure TFPointageNew.BReduireClick(Sender: TObject);
begin
GrandPetitListe(false);
end;

procedure TFPointageNew.FDateEche1KeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFPointageNew.BChercherClick(Sender: TObject);
begin
ClickRecherche ;
end;

// Affiche les montants déjà pointés
procedure TFPointageNew.ChargeInfosCompte ;
BEGIN
	MAJDebPivot:=QGene.FindField('G_TOTDEBPTP').AsFloat; MAJCredPivot:=QGene.FindField('G_TOTCREPTP').AsFloat ;
	MAJDebDevise:=QGene.FindField('G_TOTDEBPTD').AsFloat; MAJCredDevise:=QGene.FindField('G_TOTCREPTD').AsFloat ;
END ;

procedure TFPointageNew.BAbandonClick(Sender: TObject);
begin CloseFen ; end;

procedure TFPointageNew.FormClose(Sender: TObject; var Action: TCloseAction);
begin
QGene.Close;
FListe.VidePile(True) ;
PurgePopup(POPS) ; PurgePopup(POPZ) ;
if Parent is THPanel then
   BEGIN
   _Bloqueur('nrPointage',False) ;
   Action:=caFree ;
   END ;
end;

procedure CalculLesTotaux(Cpt : String ; Var TotInit : TTotCptInitiaux ; LaDevise : String) ;
Var Q : TQuery ;
    St : String ;
    StSup : String ;
BEGIN
Fillchar(TotInit,SizeOf(TotInit),#0) ;
St:='Select G_TOTDEBPTP,G_TOTCREPTP,G_TOTDEBPTD,G_TOTCREPTD FROM GENERAUX WHERE G_GENERAL="'+Cpt+'" ' ;
Q:=OpenSQL(St,TRUE) ;
// Total des mouvements pointés
If Not Q.Eof Then
  BEGIN
  TotInit[tsmPivot,1].Deb:=Q.Fields[0].AsFloat ;
  TotInit[tsmPivot,1].Cre:=Q.Fields[1].AsFloat ;
  TotInit[tsmDevise,1].Deb:=Q.Fields[2].AsFloat ;
  TotInit[tsmDevise,1].Cre:=Q.Fields[3].AsFloat ;
  TotInit[tsmEuro,1].Deb:=0;
  TotInit[tsmEuro,1].Cre:=0;
  END ;
Ferme(Q) ;
(*
If VH^.Suivant.Code<>'' Then StSup:='(E_EXERCICE="'+VH^.EnCours.Code+'" OR E_EXERCICE="'+VH^.Suivant.Code+'") '
                        Else StSup:='E_EXERCICE="'+VH^.EnCours.Code+'" ' ;
*)
If VH^.Suivant.Code<>'' Then StSup:='(E_EXERCICE="'+VH^.EnCours.Code+'" OR (E_EXERCICE="'+VH^.Suivant.Code+'" AND E_ECRANOUVEAU="N")) '
                        Else StSup:='E_EXERCICE="'+VH^.EnCours.Code+'" ' ;
St:='SELECT SUM(E_DEBIT),SUM(E_CREDIT),SUM(E_DEBITDEV),SUM(E_CREDITDEV) '+
    'FROM ECRITURE WHERE E_GENERAL="'+Cpt+'" '+
    'AND E_QUALIFPIECE="N" AND '+StSup ;
If LaDevise<>'' Then St:=St+' AND E_DEVISE="'+LaDevise+'" ' ;
Q:=OpenSQL(St,TRUE) ;

// Total de tous les mouvements
If Not Q.Eof Then
  BEGIN
  TotInit[tsmPivot,2].Deb:=Q.Fields[0].AsFloat ;
  TotInit[tsmPivot,2].Cre:=Q.Fields[1].AsFloat ;
  TotInit[tsmDevise,2].Deb:=Q.Fields[2].AsFloat ;
  TotInit[tsmDevise,2].Cre:=Q.Fields[3].AsFloat ;
  TotInit[tsmEuro,2].Deb:=0;
  TotInit[tsmEuro,2].Cre:=0;
  END ;
Ferme(Q) ;
// Total des mouvements non pointés
TotInit[tsmPivot,0].Deb:=TotInit[tsmPivot,2].Deb-TotInit[tsmPivot,1].Deb ;
TotInit[tsmPivot,0].Cre:=TotInit[tsmPivot,2].Cre-TotInit[tsmPivot,1].Cre ;
TotInit[tsmDevise,0].Deb:=TotInit[tsmDevise,2].Deb-TotInit[tsmDevise,1].Deb ;
TotInit[tsmDevise,0].Cre:=TotInit[tsmDevise,2].Cre-TotInit[tsmDevise,1].Cre ;
TotInit[tsmEuro,0].Deb:=TotInit[tsmEuro,2].Deb-TotInit[tsmEuro,1].Deb ;
TotInit[tsmEuro,0].Cre:=TotInit[tsmEuro,2].Cre-TotInit[tsmEuro,1].Cre ;
END ;

procedure TFPointageNew.BComplementClick(Sender: TObject);
begin
ClickInfoPointage(false) ;
end;

procedure TFPointageNew.FormShow(Sender: TObject);
var Q   : TQuery;
    SQL,StGene,StConf : string;
    i   : integer;
begin
BAppliquerAFaire:=FALSE ; BValiderEnCours:=FALSE ; OldGen:='' ;
FDateEche1.Text:=StDate1900 ; FDateEche2.Text:=StDate2099 ;
Pages.ActivePage:=PStandard ;
DebitCpBanquePivot:=0 ; CreditCpBanquePivot:=0 ;
DebitCpBanqueEuro:=0 ; CreditCpBanqueEuro:=0 ;
BParam.Visible:=TRUE ;
If V_PGI.Synap Then BEGIN If OkPointe Then BEGIN BRefP.Tag:=0 ; END ; END ;
DecimP:= V_PGI.OkDecV ; DecimE:=V_PGI.OkDecE ; Symbole:=V_PGI.SymbolePivot ;
FModePaie.ItemIndex:=0; FListe.ListeParam:='CPPOINTAGE' ;
FListe.TypeSais:=tsPointage; Modif:=false; DBNav.Enabled:=Not Modif; BValider.Enabled:=Modif; FChoixDevise.Enabled:=Not Modif ;
CRefP:=-1; CDateP:=-1; ListeVide:=True ;
DPDebitP:=0 ; DPCreditP:=0 ; DPDebitD:=0 ; DPCreditD:=0 ; DPDebitE:=0 ; DPCreditE:=0 ;
BLanceSaisie.Enabled:=OkPointe;
TuPeuxSupprimer:=False ;
HLabel15.Visible:=False ; FRestePointeCur.Visible:=False ;
HLabel12.Visible:=Not OkPointe ; FDPDepoint.Visible:=Not OkPointe ;
if OkPointe then
   BEGIN
	if gbJournal then Caption:=Messages.Mess.Strings[28]
							 else Caption:=Messages.Mess.Strings[3];
   HelpContext:=7604000 ;
   for i:=0 to PStandard.ControlCount-1 do
      BEGIN
      PStandard.Controls[i].Visible:=(PStandard.Controls[i].Tag<=1);
      PStandard.Controls[i].Enabled:=(PStandard.Controls[i].Tag<=1);
      END ;
   FDatePointage.Color:=clWindow; FDatePointage.EditMask:=DateMask ;
   FDatePointage.Text:=DateToStr(V_PGI.DateEntree);
   FDateCptaDeb.Text:=DateToStr(VH^.Encours.Deb) ;
   FDateCptaFin.Text:=DateToStr(V_PGI.DateEntree) ;
   END else
   BEGIN
   if gbJournal then Caption:=Messages.Mess.Strings[29]
								else Caption:=Messages.Mess.Strings[4];
   HelpContext:=7607000 ;
   for i:=0 to PStandard.ControlCount-1 do PStandard.Controls[i].Visible:=(PStandard.Controls[i].Tag<>1);
   FDatePointage.Color:=clBtnFace; FDatePointage.EditMask:='';
   FDatePointage.Text:=''; FDatePointage.Enabled:=False ;
   FDebitPointagePivot.Enabled:=FALSE ; FCreditPointagePivot.Enabled:=FALSE ;
   FDebitPointageEuro.Enabled:=FALSE ; FCreditPointageEuro.Enabled:=FALSE ;
   FDateCptaDeb.Text:=StDate1900 ;
   FDateCptaFin.Text:=StDate2099 ;
//   CAfficheOppose.Left:=TGen.Left ;
   END;
if gbJournal then begin
  FGene.Visible := False; fJournal.Visible := True;
  TGen.Caption := 'Journal';
end
else fJournal.Visible := False;
FEntete.Visible:=OkPointe  ; FPied.Visible:= OkPointe ;
QGene.Close ; QGene.SQL.Clear ;
StConf:=SQLConf('GENERAUX') ; if StConf<>'' then StConf:=' AND '+StConf ;
if gbJournal then
	StGene:='Select G_GENERAL, G_LIBELLE, G_NATUREGENE, G_TOTDEBP, G_TOTCREP, G_DEBNONPOINTE, G_CREDNONPOINTE, '
			   +'BQ_DEVISE, G_TOTDEBPTP, G_TOTCREPTP, G_TOTDEBPTD, G_TOTCREPTD, J_JOURNAL, J_LIBELLE '
         +'From GENERAUX Left join BANQUECP on G_GENERAL=BQ_GENERAL left join Journal on G_GENERAL=J_CONTREPARTIE '
         +'Where G_POINTABLE="X" AND J_NATUREJAL="BQE"'+StConf+' Order by G_GENERAL, J_JOURNAL'
else begin
	StGene:='Select G_GENERAL, G_LIBELLE, G_NATUREGENE, G_TOTDEBP, G_TOTCREP, G_DEBNONPOINTE, G_CREDNONPOINTE, '
			   +'BQ_DEVISE, G_TOTDEBPTP, G_TOTCREPTP, G_TOTDEBPTD, G_TOTCREPTD, J_JOURNAL, J_LIBELLE '
//			   +'BQ_DEVISE, G_TOTDEBPTP, G_TOTCREPTP, G_TOTDEBPTD, G_TOTCREPTD '
         +'From GENERAUX Left join BANQUECP on G_GENERAL=BQ_GENERAL left join Journal on G_GENERAL=J_CONTREPARTIE '
//         +'From GENERAUX Left join BANQUECP on G_GENERAL=BQ_GENERAL '
      	 +'Where G_POINTABLE="X"'+StConf+' Order by G_GENERAL' ;
end;
QGene.SQL.Add(StGene) ; // Compte pointable
ChangeSQL(QGene) ; QGene.Open ; SGeneDataChange(Nil,Nil) ;
ChargeInfosCompte ;
// pour ouvrir sur liste vide
SQL:='Select * from ECRITURE Where ';
SQL:=SQL+'E_JOURNAL="'+W_W+'" and E_EXERCICE="'+W_W+'" and E_DATECOMPTABLE="'+UsDateTime(IDate1900)+'" AND E_NUMEROPIECE=000'; //Lhe le 21/05/97 n°1522
Q:=OpenSQL(SQL,true);
RemplitGrid(Q);
Ferme(Q); ChangeSQL(QDevise) ; //QDevise.Prepare ;
PrepareSQLODBC(QDevise) ;
UpdateCaption(Self) ;
if VH^.TenueEuro then
   BEGIN
//   CAfficheOppose.Caption:=HDiv.Mess[0]+' '+RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False) ;
   FListe.Cells[4,0]:='F' ;
   FListe.Cells[7,0]:=HDiv.Mess[1]+' '+RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False) ;
   FListe.Cells[8,0]:=HDiv.Mess[2]+' '+RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False) ;
   END ;
If OkZoom Then
  BEGIN
  FGene.Text:=CptZoom ; RepositionGene ; ChargeInfosCompte ; FRefPointage.Text:='' ;
  END ;
{$IFDEF CCS3}
BLanceSaisie.Visible:=False ;
{$ENDIF}
end;

procedure TFPointageNew.EnvoiePointage(R:integer) ;
BEGIN
if FListe.Cells[CRefP,R]<>'' then
   BEGIN
   if FListe.Cells[CRefP,R]=Trim(FRefPointage.Text) then
      BEGIN
      DepointeLes(R);
//      CalculRestePointe;
      END else if Messages.Execute(1,'','')= mrYes then DepointeLes(R);
   END else
   BEGIN
   PointeLes(R) ; //CalculRestePointe;
   END ;
Modif:=true; DBNav.Enabled:=Not Modif; BValider.Enabled:=Modif; FChoixDevise.Enabled:=Not Modif ;
END ;

procedure TFPointageNew.EnvoieDepointage(R:integer) ;
BEGIN
if FListe.Cells[CRefP,R]<>''then
   BEGIN
   DepointeLes(R); CalculDepointe(R,true);
   END else
   BEGIN
   PointeLes(R); CalculDepointe(R,false);
   END;
Modif:=true; DBNav.Enabled:=Not Modif; BValider.Enabled:=Modif; FChoixDevise.Enabled:=Not Modif ;
END ;

procedure TFPointageNew.AfficheTotauxBasEcran ;
BEGIN
{ Totaux Pointés sur la référence (bas d'écran) }
{if CAfficheOppose.Checked then
   BEGIN
   FCreditPA.Value:=XCreditE ;
   FDebitPA.Value:=XDebitE ;
   END else}
   if RestitutionEnDevise And OkMajTotPDev then BEGIN
   FCreditPA.Value:=XCreditD ; FDebitPA.Value:=XDebitD ;
   END else
   BEGIN
   FCreditPA.Value:=XCreditP ; FDebitPA.Value:=XDebitP ;
   END;
AfficheLeSolde(FSoldeA,FDebitPA.Value,FCreditPA.Value) ;
END ;

procedure TFPointageNew.AfficheTotauxPointeHautDEcran ;
BEGIN
{ Totaux pointés sur le compte (haut d'écran) }
(*GG
if FAffichDev.Visible then
   BEGIN
   FCreditPIniA.Value:=Valeur(FTotCredPDevise.Text)+QCreditD ;
   FDebitPIniA.Value:=Valeur(FTotDebPDevise.Text)+QDebitD ;
   END else if CAfficheOppose.Checked then
   BEGIN
   FCreditPIniA.Value:=Valeur(FTotCredPEuro.Text)+QCreditE ;
   FDebitPIniA.Value:=Valeur(FTotDebPEuro.Text)+QDebitE ;
   END else
   BEGIN
   FCreditPIniA.Value:=Valeur(FTotCredPPivot.Text)+QCreditP ;
   FDebitPIniA.Value:=Valeur(FTotDebPPivot.Text)+QDebitP ;
   END ;
*)
{if CAfficheOppose.Checked then
   BEGIN
   FCreditPIniA.Value:=LesTotauxCpt[tsmEuro,1].Cre ;
   FDebitPIniA.Value:=LesTotauxCpt[tsmEuro,1].Deb ;
   END else} if RestitutionEnDevise And OkMajTotPDev then
   BEGIN
   FCreditPIniA.Value:=LesTotauxCpt[tsmDevise,1].Cre ;
   FDebitPIniA.Value:=LesTotauxCpt[tsmDevise,1].Deb ;
   END else
   BEGIN
   FCreditPIniA.Value:=LesTotauxCpt[tsmPivot,1].Cre ;
   FDebitPIniA.Value:=LesTotauxCpt[tsmPivot,1].Deb ;
   END ;
AfficheLeSolde(FSoldePIniA,FDebitPIniA.Value,FCreditPIniA.Value) ;
AfficheLeSolde(FSoldeIniA,FDebitPIniA.Value+FDebitDPIniA.Value,FCreditPIniA.Value+FCreditDPIniA.Value) ;
END ;

procedure TFPointageNew.AfficheTotalDepointe ;
BEGIN
{ Affichage total dépointé en dépointage }
//if CAfficheOppose.Checked then AfficheLeSolde(FDPDepoint,DPDebitE,DPCreditE) else
   if RestitutionEnDevise And OkMajTotPDev then AfficheLeSolde(FDPDepoint,DPDebitD,DPCreditD) else
      AfficheLeSolde(FDPDepoint,DPDebitP,DPCreditP) ;
END ;

procedure TFPointageNew.MajTotalPourUnPointage(OBM : TOBM) ;
Var DP,CP,DD,CD : double;
BEGIN
If OBM=NIL Then Exit ;
DD:=OBM.GetMvt('E_DEBITDEV')  ; CD:=OBM.GetMvt('E_CREDITDEV') ;
DP:=OBM.GetMvt('E_DEBIT')     ; CP:=OBM.GetMvt('E_CREDIT') ;
XDebitP:=XDebitP+DP ; XCreditP:=XCreditP+CP ;
If OBM.GetMvt('E_DEVISE')=FDevise.Text Then
  BEGIN
  XDebitD:=XDebitD+DD ; XCreditD:=XCreditD+CD ;
  END ;
FNbP.Value:=FNbP.Value+1 ;
AfficheTotauxBasEcran ;
MAJCredPivot:=MAJCredPivot+CP   ; MAJDebPivot:=MAJDebPivot+DP ;
If OBM.GetMvt('E_DEVISE')=FDevise.Text Then
  BEGIN
  MAJCredDevise:=MAJCredDevise+CD ; MAJDebDevise:=MAJDebDevise+DD ;
  END ;
END ;

procedure TFPointageNew.PointeLes(R:integer) ;
var OBM : TOBM;
BEGIN
if OkPointe then FListe.Cells[FListe.ColCount-1,R]:='*' else FListe.Cells[FListe.ColCount-1,R]:='';
FListe.Cells[CRefP,R]:=Trim(FRefPointage.Text);
if CDateP>=0 then FListe.Cells[CDateP,R]:=FDatePointage.Text ;
OBM:=GetO(FListe,R);
OBM.PutMvt('E_REFPOINTAGE',Trim(FRefPointage.Text)); OBM.PutMvt('E_DATEPOINTAGE',StrToDate(FDatePointage.Text));
MajTotalPourUnPointage(OBM) ;
(*
DD:=OBM.GetMvt('E_DEBITDEV')  ; CD:=OBM.GetMvt('E_CREDITDEV') ;
DE:=OBM.GetMvt('E_DEBITEURO') ; CE:=OBM.GetMvt('E_CREDITEURO') ;
DP:=OBM.GetMvt('E_DEBIT')     ; CP:=OBM.GetMvt('E_CREDIT') ;
XDebitP:=XDebitP+DP ; XCreditP:=XCreditP+CP ;
If OBM.GetMvt('E_DEVISE')=FDevise.Text Then
  BEGIN
  XDebitD:=XDebitD+DD ; XCreditD:=XCreditD+CD ;
  END ;
XDebitE:=XDebitE+DE ; XCreditE:=XCreditE+CE ;
FNbP.Value:=FNbP.Value+1 ;
AfficheTotauxBasEcran ;
MAJCredPivot:=MAJCredPivot+CP   ; MAJDebPivot:=MAJDebPivot+DP ;
If OBM.GetMvt('E_DEVISE')=FDevise.Text Then
  BEGIN
  MAJCredDevise:=MAJCredDevise+CD ; MAJDebDevise:=MAJDebDevise+DD ;
  END ;
MAJCredEuro:=MAJCredEuro+CE     ; MAJDebEuro:=MAJDebEuro+DE ;
*)
END ;

procedure TFPointageNew.DepointeLes(R:integer) ;
var OBM : TOBM;
    DP,CP,DD,CD : double;
BEGIN
if OkPointe then FListe.Cells[FListe.ColCount-1,R]:='' else FListe.Cells[FListe.ColCount-1,R]:='*';
FListe.Cells[CRefP,R]:='' ; FListe.Cells[CDateP,R]:=StDate1900 ;
OBM:=GetO(FListe,R); OBM.PutMvt('E_REFPOINTAGE','');
OBM.PutMvt('E_DATEPOINTAGE',IDate1900);
DD:=OBM.GetMvt('E_DEBITDEV')  ; CD:=OBM.GetMvt('E_CREDITDEV') ;
DP:=OBM.GetMvt('E_DEBIT')     ; CP:=OBM.GetMvt('E_CREDIT') ;
XDebitP:=XDebitP-DP ; XCreditP:=XCreditP-CP ;
If OBM.GetMvt('E_DEVISE')=FDevise.Text Then
  BEGIN
  XDebitD:=XDebitD-DD ; XCreditD:=XCreditD-CD ;
  END ;
FNbP.Value:=FNbP.Value-1 ;
AfficheTotauxBasEcran ;
// mise à jour des totaux non pointés sur N-1
(*
ExoEcr:=OBM.GetMvt('E_EXERCICE');
if ((ExoEcr<>VH^.Suivant.Code) and (ExoEcr<>VH^.Encours.Code)) then
   BEGIN
   MAJCredPivot:=MAJCredPivot+CP ; MAJDebPivot:=MAJDebPivot+DP ;
   MAJCredEuro:=MAJCredEuro+CE   ; MAJDebEuro:=MAJDebEuro+DE ;
   END;
*)
MAJCredPivot:=MAJCredPivot-CP   ; MAJDebPivot:=MAJDebPivot-DP ;
If OBM.GetMvt('E_DEVISE')=FDevise.Text Then
  BEGIN
  MAJCredDevise:=MAJCredDevise-CD ; MAJDebDevise:=MAJDebDevise-DD ;
  END ;
END ;

procedure TFPointageNew.BAppliquerClick(Sender: TObject);
Var QC : tQuery ;
    OkOk : Boolean ;
begin
OkOk:=TRUE ;
QC:=OpenSQL('SELECT G_GENERAL, G_POINTABLE FROM GENERAUX WHERE G_GENERAL="'+FGene.Text+'" ',TRUE) ;
If Not QC.Eof Then
  BEGIN
  If QC.FindField('G_POINTABLE').AsString<>'X' Then
    BEGIN
    OkOk:=FALSE ;
    If OkPointe Then Messages.Execute(26,'','') Else Messages.Execute(27,'','') ;
    END ;
  END Else
  BEGIN
  OkOk:=FALSE ;
  If OkPointe Then Messages.Execute(24,'','') Else Messages.Execute(25,'','') ;
  END ;
Ferme(QC) ;
If Not OkOk Then Exit ;
ClickChercher ;
end;

procedure TFPointageNew.ChercheEcritures ;
var SQL : string ;
    i           : integer;
    C           : THNumEdit;
    Q           : TQuery;
BEGIN
FListe.VidePile(True) ;

if gbJournal then SQL :='Select * From ECRITURE Where E_NUMEROPIECE in(Select E_NUMEROPIECE '
		 				 else SQL :='Select * ';
SQL:=SQL+'From ECRITURE Where '+ WhereEcr(false);
if gbJournal then begin
	if OkPointe then
    SQL := SQL+') and E_GENERAL<>"'+FGene.Text+'" and E_REFPOINTAGE="" and E_JOURNAL="'+gszJournal+'"'
  else
  	SQL := SQL+') and E_GENERAL<>"'+FGene.Text+'" and E_JOURNAL="'+gszJournal+'"  and UPPER(E_REFPOINTAGE)="'+Trim(FRefPointage.Text)+'" and E_DATEPOINTAGE="'+USDATE(FDatePointage)+'"'
  end
else begin
    SQL := SQL+' Order by E_GENERAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE';
end;
Q:=OpenSQL(SQL,True) ; ListeVide:=(Q.EOF) ;
RemplitGrid(Q) ;
Ferme(Q) ;
CalculeTotaux ;
for i:=0 to ComponentCount-1 do if Components[i] is THNumEdit then
    BEGIN
    C:=THNumEdit(Components[i]);
    if C.Tag<>-1 then
       BEGIN
       {if CAfficheOppose.Checked then ChangeMask(C,DecimE,'') else} ChangeMask(C,DecimP,Symbole) ;
       END ;
    END;
END ;

// Construit la clause WHERE en fonction des paramètres saisis
function TFPointageNew.WhereEcr(OkSaisie: boolean): string ;
var StV8,StWhere,StConf : string;
BEGIN
WhereEcr:='';
if gbJournal then StWhere:=' E_GENERAL<>"'+FGene.Text+'"'
             else StWhere:=' E_GENERAL="'+FGene.Text+'"';
if OkPointe then StWhere:=StWhere+' and E_EXERCICE>="'+QuelExo(FDateCptaDeb.Text) + '" and E_EXERCICE<="'+QuelExo(FDateCptaFin.Text)+'"' // and E_DATECOMPTABLE>="'+USDATE(FDateCptaDeb)+'" and E_DATECOMPTABLE<="'+USDATE(FDateCptaFin)+'"'
            else StWhere:=StWhere+' and E_EXERCICE<="'+QuelExo(FDatePointage.Text)+'" ';
/// test à améliorer
//if ((FDateEche1.Text<>Encode date'01/01/1990') or (FDateEche2.Text<>'31/12/2099')) then
StWhere:=StWhere+' and E_DATEECHEANCE>="'+USDATE(FDateEche1)+'" and E_DATEECHEANCE<="'+USDATE(FDateEche2)+'"';
StWhere:=StWhere+' and E_DATECOMPTABLE>="'+USDATE(FDateCptaDeb)+'" and E_DATECOMPTABLE<="'+USDATE(FDateCptaFin)+'"';
StWhere:=StWhere+' and E_QUALIFPIECE="N" and E_ECRANOUVEAU<>"OAN" and E_ECRANOUVEAU<>"C" and E_TRESOLETTRE<>"X" '; //écritures normales et <> solde d'à nouveau
StV8:=LWhereV8 ; if StV8<>'' then StWhere:=StWhere+' AND '+StV8 ;
(*
If RestitutionEnDevise Then
  BEGIN
  If (FChoixDevise.ItemIndex>0) And OkMajTotPDev Then StWhere:=StWhere+' AND E_DEVISE="'+FChoixDevise.Value+'" ' ;
  END ;
*)
If FChoixDevise.ItemIndex>0 Then BEGIN StWhere:=StWhere+' AND E_DEVISE="'+FChoixDevise.Value+'" ' ; END ;
if OkSaisie then
   BEGIN
   StWhere:=StWhere+ ' and E_REFPOINTAGE<>""';
   END else
   BEGIN
   if OkPointe then StWhere:=StWhere+ ' and E_REFPOINTAGE=""'
               else StWhere:=StWhere+ ' and E_REFPOINTAGE<>""';
   END;
if (RestitutionEnDevise) And OkMajTotPDev then
   BEGIN
   if FMontant1.Text<>'' then  StWhere:=StWhere+ ' and (E_DEBITDEV+E_CREDITDEV)>="'+FMontant1.Text+'"' ;
   if FMontant2.Text<>'' then  StWhere:=StWhere+ ' and (E_DEBITDEV+E_CREDITDEV)<="'+FMontant2.Text+'"' ;
   END else
   BEGIN
   if FMontant1.Text<>'' then  StWhere:=StWhere+ ' and (E_DEBIT+E_CREDIT)>="'+FMontant1.Text+'"' ;
   if FMontant2.Text<>'' then  StWhere:=StWhere+ ' and (E_DEBIT+E_CREDIT)<="'+FMontant2.Text+'"' ;
   END;
if FModePaie.ItemIndex<>0 then StWhere:=StWhere+ ' and E_MODEPAIE="'+FModePaie.Value+'"';
if Not OkPointe then StWhere:=StWhere+' and UPPER(E_REFPOINTAGE)="'+Trim(FRefPointage.Text)+'" and E_DATEPOINTAGE="'+USDATE(FDatePointage)+'"';
StConf:=SQLConf('ECRITURE') ; if StConf<>'' then StWhere:=StWhere+' AND '+StConf ;
WhereEcr:=StWhere;
END;

procedure TFPointageNew.CalculeTotaux ;
var i           : integer;
    OBM         : TOBM;
    DD,CD,DP,CP : Double ;
    LaDevise : String ;
BEGIN
// calcul et affectations des cumuls non pointés avant pointage
LaDevise:='' ; If Not CptEnPivot Then LaDevise:=FDevise.Text ;
CalculLesTotaux(FGene.Text,LesTotauxCpt,LaDevise) ;
{if CAfficheOppose.Checked then
   BEGIN
   FDebitDPIniA.Value:=LesTotauxCpt[tsmEuro,0].Deb ;
   FCreditDPIniA.Value:=LesTotauxCpt[tsmEuro,0].Cre ;
   END else} if RestitutionEnDevise And OkMajTotPDev then
   BEGIN
   FDebitDPIniA.Value:=LesTotauxCpt[tsmDevise,0].Deb ;
   FCreditDPIniA.Value:=LesTotauxCpt[tsmDevise,0].Cre ;
   END else
   BEGIN
   FDebitDPIniA.Value:=LesTotauxCpt[tsmPivot,0].Deb ;
   FCreditDPIniA.Value:=LesTotauxCpt[tsmPivot,0].Cre ;
   END ;
AfficheTotauxPointeHautDEcran ;
// calcul du solde du compte avant pointage
// calcul des cumuls pointés avec référence en cours
if Not OkPointe then
   BEGIN
   For i:=1 to FListe.RowCount-1 do
       BEGIN
       OBM:=GetO(FListe,i); if OBM=NIL then Break;
       FNbP.Value:=FNbP.Value+1 ;
       DD:=OBM.GetMvt('E_DEBITDEV')  ; CD:=OBM.GetMvt('E_CREDITDEV') ;
       DP:=OBM.GetMvt('E_DEBIT')     ; CP:=OBM.GetMvt('E_CREDIT') ;
       XDebitP:=XDebitP+DP ; XCreditP:=XCreditP+CP ;
       If OBM.GetMvt('E_DEVISE')=FDevise.Text Then
         BEGIN
         XDebitD:=XDebitD+DD ; XCreditD:=XCreditD+CD ;
         END ;
       END ;
   AfficheTotauxBasEcran ;
   END else
   BEGIN
   CalculPointes(Trim(FRefPointage.Text)) ;
   END ;
//CalculRestePointe ;
END ;

procedure TFPointageNew.FDebitPointagePivotExit(Sender: TObject);
begin
DebitCpBanquePivot:=Arrondi(Valeur(FDebitPointagePivot.Text),DecimP) ;
CreditCpBanquePivot:=Arrondi(Valeur(FCreditPointagePivot.Text),DecimP) ;
DebitCpBanqueEuro:=Arrondi(Valeur(FDebitPointageEuro.Text),DecimE) ;
CreditCpBanqueEuro:=Arrondi(Valeur(FCreditPointageEuro.Text),DecimE) ;
FDebitPointagePivot.Text:=StrfMontant(DebitCpBanquePivot,15,DecimP,'',TRUE) ;
FCreditPointagePivot.Text:=StrfMontant(CreditCpBanquePivot,15,DecimP,'',TRUE) ;
FDebitPointageEuro.Text:=StrfMontant(DebitCpBanqueEuro,15,DecimE,'',TRUE) ;
FCreditPointageEuro.Text:=StrfMontant(CreditCpBanqueEuro,15,DecimE,'',TRUE) ;
if ((TEdit(Sender).Name='FDebitPointagePivot') or (TEdit(Sender).Name='FDebitPointageEuro')) then
   BEGIN
   FCreditPointagePivot.Enabled:=(DebitCpBanquePivot=0);
   FCreditPointageEuro.Enabled:=(DebitCpBanqueEuro=0);
   END ;
if ((TEdit(Sender).Name='FCreditPointagePivot') or (TEdit(Sender).Name='FCreditPointageEuro')) then
   BEGIN
   FDebitPointagePivot.Enabled:=(CreditCpBanquePivot=0) ;
   FDebitPointageEuro.Enabled:=(CreditCpBanqueEuro=0) ;
   END ;
//CalculRestePointe ;
end;

procedure TFPointageNew.BValiderClick(Sender: TObject);
begin ClickValide; end;

procedure TFPointageNew.RepositionGene;
var St : string;
BEGIN
if gbJournal then begin
	St:=gszJournal;
	if QGene.Active then QGene.Locate('J_JOURNAL',St,[loCaseInsensitive]) ;
	gszJournal:=St;
end
else begin
	St:=FGene.Text;
	if QGene.Active then QGene.Locate('G_GENERAL',St,[loCaseInsensitive]) ;
	FGene.Text:=St;
end;
END;

procedure TFPointageNew.MajTotPointeCpt;
var SQL : String;
BEGIN
  if not EstTablePartagee( 'GENERAUX' ) then
    begin
    SQL:='UPDATE GENERAUX SET G_TOTDEBPTP='+StrFPoint(MAJDebPivot)+', G_TOTCREPTP='+StrFPoint(MAJCredPivot)+', ';
    SQL:=SQL+'G_TOTDEBPTD='+StrFPoint(MAJDebDevise)+', G_TOTCREPTD='+StrFPoint(MAJCredDevise)+' ';
    SQL:=SQL+'WHERE G_GENERAL="'+FGene.Text+'"';
    if ExecuteSQL(SQL)<=0
      then V_PGI.IoError := oePointage ;
    end
  else // MAJ table cumuls si besoin
    if not CUpdateCumulsPointeMS( FGene.Text, MAJDebPivot, MAJCredPivot, MAJDebDevise, MAJCredDevise )
      then V_PGI.IoError := oePointage ;


END;

procedure TFPointageNew.SauvePointage;
var i         : integer;
    OBM       : TOBM;
    St, SQL   : string;
    M         : RMVT ;
    DD        : TDateTime ;
BEGIN
for i:=1 to FListe.RowCount-1 do
   BEGIN
   OBM:=GetO(FListe,i); if OBM=Nil then Break ;
   if OkPointe then BEGIN if OBM.GetMvt('E_REFPOINTAGE')='' then Continue ;END
               else BEGIN if OBM.GetMvt('E_REFPOINTAGE')<>'' then Continue ;END;
   St:=OBM.StPourUpdate ; if St='' then Continue ;
   if Round(OBM.GetMvt('OLDDEBIT'))<>0 then Continue ;
   St:=St+', E_DATEMODIF="'+UsTime(DateRef)+'"';
   M:=OBMToIdent(OBM,True) ; DD:=OBM.GetMvt('E_DATEMODIF') ;
   SQL:='UPDATE ECRITURE SET '+St+' Where '+WhereEcriture(tsPointage,M,True)+' AND E_DATEMODIF="'+UsTime(DD)+'"' ;
   if ExecuteSQL(SQL)<=0 then BEGIN V_PGI.IOError:=oePointage ; Break ; END ;
   END ;
if V_PGI.IOError=oeOK then EnregRefPointage;
END;

procedure TFPointageNew.BPointeClick(Sender: TObject);
begin PointeDepointe ; end;

{procedure TFPointage.RecupDevise(var Decim :integer; var Symbole:string) ;
BEGIN
//if ((FDevise.Text=V_PGI.DevisePivot) OR (Not FAffichDev.Visible)) then
if ((FDevise.Text=V_PGI.DevisePivot) OR (Not RestitutionEnDevise) Or (RestitutionEnDevise And (Not OkMajTotPDev))) then   BEGIN
   Decim:=V_PGI.OkDecV; Symbole:=V_PGI.SymbolePivot;
   END else
   BEGIN
   QDevise.Close; QDevise.ParamByName('Dev').AsString:=FDevise.Text; QDevise.Open;
   if Not QDevise.EOF then
      BEGIN
      Decim:=QDevise.Findfield('D_DECIMALE').AsInteger;
      Symbole:=QDevise.Findfield('D_SYMBOLE').AsString;
      END ;
   END ;
END ;}

procedure TFPointageNew.BLanceSaisieClick(Sender: TObject);
begin
ClickSaisie ;
end;

Function TFPointageNew.BalanceSaisie(var M:RMVT) : Boolean ;
var DEV    : RDEVISE ;
    DateCp : TDateTime;
    JAL    : string ;
BEGIN
Result:=FALSE ;
{$IFNDEF CCS3}
DateCp:=StrToDate(FDatePointage.Text);
FillChar(DEV,Sizeof(DEV),#0) ;
DEV.Code:=FDevise.Text ; GetInfosDevise(DEV) ;
DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,DateCp) ;
JAL:=ChercheJal(FGene.Text);
If Jal='' Then BEGIN Messages.Execute(20,'','') ; Exit ; END ;
Fillchar(M,Sizeof(M),#0) ;
M.Simul:='N' ; M.Valide:=False ;
M.Etabl:=VH^.ETABLISDEFAUT ; M.ANouveau:=FALSE; M.Treso:=TRUE ; M.MajDirecte:=TRUE ;
M.DateC:=DateCp ; M.Jal:=JAL ;
M.Exo:=QuelExoDT(M.DateC) ; M.Nature:='OD' ;
M.CodeD:=DEV.Code ; M.TauxD:=DEV.Taux ; M.DateTaux:=DEV.DateTaux ;
PrepareSaisTresAgio(M) ;
Result:=TRUE ;
{$ENDIF}
END;

procedure TFPointageNew.RecupereSaisie(M:RMVT);
var SQL : string;
    Q   : TQuery;
    OBM : TOBM;
BEGIN
// pointe et maj totaux pointés en pied d'écran
SQL:='Select E_DEBIT, E_CREDIT, E_REFPOINTAGE, E_DATEPOINTAGE From ECRITURE';
SQL:=SQL+' Where '+ WhereEcriture(tsPointage,M,False);
SQL:=SQL+'AND E_GENERAL="'+FGene.Text+'"' ;
Q:=OpenSQL(SQL,false);
Q.UpdateMode:=upWhereChanged ;
if Q.EOF then BEGIN Ferme(Q) ; Exit; END ;
BeginTrans;
While Not Q.EOF do
   BEGIN
   Q.Edit;
   Q.FindField('E_DATEPOINTAGE').AsDateTime:=StrToDate(FDatePointage.Text);
   Q.FindField('E_REFPOINTAGE').AsString:=Trim(FRefPointage.Text);
   Q.Post;
   Q.Next ;
   END;
CommitTrans;
Ferme(Q);
//CalculPointes(Trim(FRefPointage.Text));
// maj liste
SQL:='Select * From ECRITURE';
SQL:=SQL+' Where '+ WhereEcriture(tsPointage,M,False);
SQL:=SQL+'AND '+WhereEcr(true) ;
Q:=OpenSQL(SQL,true);
if Not Q.EOF then
   BEGIN
   FListe.SetFocus ;
   if Not ListeVide then FListe.RowCount:=FListe.RowCount+1 ;
   FListe.Row:=Fliste.RowCount-1;
   OBM:=TOBM.Create(EcrGen,'',True) ;
   OBM.ChargeMvt(Q) ; OBM.PutMvt('OLDDEBIT',1) ;
   RemplitCellules(Q);
   FListe.Objects[0,FListe.RowCount-1]:=OBM ;
   ListeVide:=False ;
   MajTotalPourUnPointage(OBM) ;
   END ;
Ferme(Q) ;
Modif:=True ; BValider.Enabled:=Modif ;
END;


{procedure TFPointage.CalculRestePointe;
var D1,C1,D2,C2 :double;
BEGIN
Exit ; { Affichage sur un champ invisible !!! (FRestePointeCur) }
{*** FDebitPointage et FCreditPointage = solde bancaire donc sens bancaire   ***}
{*** FCreditP et FDebitP               = solde comptable donc sens comptable ***}
{*** FRestePointe                      = sens comptable                      ***}
{if CAfficheOppose.Checked then
   BEGIN
   if DebitCpBanqueEuro-CreditCpBanqueEuro<0
      then BEGIN C1:=0 ; D1:=Abs(DebitCpBanqueEuro-CreditCpBanqueEuro); END
      else BEGIN C1:=DebitCpBanqueEuro-CreditCpBanqueEuro ; D1:=0 ; END;
   if CreditCpBanqueEuro-DebitCpBanqueEuro<0
      then BEGIN D2:=0 ; C2:=Abs(CreditCpBanqueEuro-DebitCpBanqueEuro); END
      else BEGIN D2:=CreditCpBanqueEuro-DebitCpBanqueEuro; C2:=0 ; END;
   END else
   BEGIN
   if DebitCpBanquePivot-CreditCpBanquePivot<0
      then BEGIN C1:=0 ; D1:=Abs(DebitCpBanquePivot-CreditCpBanquePivot); END
      else BEGIN C1:=DebitCpBanquePivot-CreditCpBanquePivot ; D1:=0 ; END;
   if CreditCpBanquePivot-DebitCpBanquePivot<0
      then BEGIN D2:=0 ; C2:=Abs(CreditCpBanquePivot-DebitCpBanquePivot); END
      else BEGIN D2:=CreditCpBanquePivot-DebitCpBanquePivot; C2:=0 ; END;
   END ;
AfficheLeSolde(FRestePointeCur,D1+D2,C1+C2);
END;}

procedure TFPointageNew.CalculDepointe(R:integer; Plus:boolean);
var
  OBM : TOBM;
BEGIN
(*
if FDPDepoint.Debit then BEGIN TDS:=FDPDepoint.Value ; TCS:=0 ; END
                    else BEGIN TDS:= 0; TCS:=FDPDepoint.Value ; END ;
*)
OBM:=GetO(FListe,R) ;
if Plus then
   BEGIN
   DPDebitD:=DPDebitD+OBM.GetMvt('E_DEBITDEV')  ; DPCreditD:=DPCreditD+OBM.GetMvt('E_CREDITDEV') ;
   DPDebitP:=DPDebitP+OBM.GetMvt('E_DEBIT')     ; DPCreditP:=DPCreditP+OBM.GetMvt('E_CREDIT') ;
   END else
   BEGIN
   DPDebitD:=DPDebitD-OBM.GetMvt('E_DEBITDEV')  ; DPCreditD:=DPCreditD-OBM.GetMvt('E_CREDITDEV') ;
   DPDebitP:=DPDebitP-OBM.GetMvt('E_DEBIT')     ; DPCreditP:=DPCreditP-OBM.GetMvt('E_CREDIT') ;
   END ;
AfficheTotalDepointe ;
END;


procedure TFPointageNew.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var Rep:integer ;
begin
if Modif then
   BEGIN
   Rep:= Messages.Execute(2,'','');
   case Rep of
        mrYes    : SauveReseau;
        mrNo     : ;
        mrCancel : CanClose:=false;
   end ;
   END ;

end;

procedure TFPointageNew.FindDialogFind(Sender: TObject);
begin
Rechercher(FListe,FindDialog, FirstFind);
end;


Procedure TFPointageNew.RetoucheTitreListe;
Var StDevise,St : String ;
//    OkTotEnDevise : Boolean ;
BEGIN
StDevise:=VH^.libDevisePivot ; //OkTotEnDevise:=FALSE ;
//If RestitutionEnDevise And OkMajTotPDev Then
If FChoixDevise.ItemIndex>0 Then
  BEGIN
  {OkTotEnDevise:=TRUE ;} StDevise:=RechDom('TTDEVISETOUTES',FChoixDevise.Value,False) ;
  END ;
FListe.Cells[5,0]:=HDiv.Mess[1]+' '+StDevise ;
FListe.Cells[6,0]:=HDiv.Mess[2]+' '+StDevise ;
{If CAfficheOppose.Checked Then
  BEGIN
  If VH^.TenueEuro Then St:=' ('+VH^.libDeviseFongible+')' Else St:=' (Euro)' ;
  END Else} If OkMajTotPDev Then St:=' ('+StDevise+')' Else St:=' ('+VH^.libDevisePivot+')' ;
//  END Else If OkMajTotPDev Then St:=' ('+StDevise+')' Else St:=' ('+VH^.libDevisePivot+')' ;
FAvantPointage.Caption:=HDiv.Mess[3]+St ; FSession.Caption:=HDiv.Mess[4]+St ;
END ;


Function TFPointageNew.InitPourDevise : Boolean ;
BEGIN
Result:=TRUE ; OkMajTotPDev:=TRUE ;
If FAffichDev.Visible Then
  BEGIN
  If (FChoixDevise.ItemIndex<0) Or (FChoixDevise.Value<>FDevise.Text) Then OkMajTotPDev:=FALSE ;
  END ;
RestitutionEnDevise:=(FChoixDevise.ItemIndex>0) And (FchoixDevise.Value<>V_PGI.DevisePivot) And (FchoixDevise.Value<>V_PGI.DeviseFongible) ;
RetoucheTitreListe ;
END ;

procedure TFPointageNew.SGeneDataChange(Sender: TObject; Field: TField);
Var Q : TQuery ;
    St : String ;
begin
If BValiderEnCours Then Exit ; If Modif Then Exit ;
OkMajTotPDev:=TRUE ;
if Field=Nil then
   BEGIN
   if gbJournal then
   begin
   	InitCaption(Self,QGeneJ_LIBELLE.AsString,QGeneG_GENERAL.AsString);
   	gszJournal := QGene.FindField('J_JOURNAL').AsString;
//    TGen.Caption := QGeneJ_LIBELLE.AsString;
    fJournal.Text := QGeneJ_JOURNAL.AsString;
   end
   else InitCaption(Self,QGeneG_GENERAL.AsString,QGeneG_LIBELLE.AsString);
   FGene.Text:= QGeneG_General.AsString ;

   CptBqe:=FNatureGene.Text='BQE' ;
   FAffichDev.Visible:=TRUE ; If Not CptBqe Then FAffichDev.Visible:=FALSE ;
   CptEnPivot:=TRUE ;
   If ((FDevise.Text<>V_PGI.DevisePivot) And (FDevise.Text<>V_PGI.DeviseFongible) And (FNatureGene.Text='BQE')) Then CptEnPivot:=FALSE ;
(*
   if FAffichDev.Visible then
      BEGIN
      FAffichDev.Caption:=Messages.Mess.Strings[17]+' '+FDevise.Text ;
      CAfficheOppose.Enabled:=False ; CAfficheOppose.Checked:=False ;
      TFChoixDevise.visible:=TRUE ; FChoixDevise.visible:=TRUE ;
      FChoixDevise.Value:=FDevise.Text ;
      END else
      BEGIN
      CAfficheOppose.Enabled:=True ;
      TFChoixDevise.visible:=FALSE ; FChoixDevise.visible:=FALSE ;
      FChoixDevise.Itemindex:=-1 ;
      END ;
*)
  If CptBqe Then
    BEGIN
    TFChoixDevise.visible:=TRUE ; FChoixDevise.visible:=TRUE ;
    If ((OldGen<>FGene.Text) and (gbJournal=False)) or ((OldGen<>gszJournal) and (gbJournal=True)) Then
      BEGIN
      FChoixDevise.Value:=FDevise.Text ;
      St:=FDevise.Text ;
      Q:=Opensql('SELECT D_LIBELLE FROM DEVISE WHERE D_DEVISE="'+FDevise.Text+'" ',TRUE) ;
      If Not Q.Eof Then St:=Q.Fields[0].AsString ;
      Ferme(Q) ;
      FAffichDev.Caption:=Messages.Mess.Strings[17]+' '+St ;
      FChoixDevise.Value:=FDevise.Text ;
      If Not CptEnPivot Then FChoixDevise.Value:=FDevise.Text Else FChoixDevise.Itemindex:=0 ;
      END ;
    END Else
    BEGIN
    TFChoixDevise.visible:=FALSE ; FChoixDevise.visible:=FALSE ;
    FChoixDevise.Itemindex:=0 ;
    END;
   END;
InitPourDevise ;
ColorOpposeEuro(FListe,False,RestitutionEnDevise) ;//CAfficheOppose.Checked,RestitutionEnDevise) ;
ISigneEuro1.Visible:=ISigneEuro.Visible ;
if ((FDevise.Text=V_PGI.DevisePivot) and (VH^.TenueEuro))
   then HlabelDevise.Caption:='('+RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False)+')'
   else HlabelDevise.Caption:='('+RechDom('TTDEVISETOUTES',FDevise.Text,False)+')' ;
if gbJournal then OldGen:=gszJournal else OldGen:=FGene.Text;
end;

procedure TFPointageNew.RemplitGrid(Q : TQuery);
var   OBM : TOBM;
      i   : integer;
Const Max = 10000 ;
BEGIN
i:=0 ;
CRefP:=AttribCol('E_REFPOINTAGE'); CDateP:=AttribCol('E_DATEPOINTAGE');
if CRefP<0 then
   BEGIN
   if OkPointe then Messages.Execute(9,'','') else Messages.Execute(10,'','');
   exit;
   END;
if Q.EOF then exit;
While ((Not Q.EOF) and (i<=Max)) do
   BEGIN
   OBM:=TOBM.Create(EcrGen,'',False) ;
   OBM.ChargeMvt(Q) ;
   OBM.PutMvt('OLDDEBIT',0) ;
   RemplitCellules(Q);
   FListe.Objects[0,FListe.RowCount-1]:=TObject(OBM) ;
   FListe.RowCount:=FListe.RowCount+1 ;
   inc(i) ;
   Q.Next ;
   END;
FListe.RowCount:=FListe.RowCount-1 ;
if i>Max then  Messages.Execute(16,'','');
END;


Function TFPointageNew.AttribCol(Champ: string) : integer;
var St,St1 :string;
    i : integer;
BEGIN
St:=FListe.Titres[0]; i:=0; Result:=0 ;
while St<>'' do
   BEGIN
   St1:= ReadTokenSt(St);
   if St1=Champ then BEGIN Result:=i; exit; END else inc(i);
   END;

END;

procedure TFPointageNew.RemplitCellules(Q :TQuery);
Var St,St1,StCell,StP : String ;
    C,LaDec       : integer ;
    T             : TField ;
BEGIN
St:=FListe.Titres[0] ; C:=0 ;
While St<>'' do
   BEGIN
   St1:=ReadTokenSt(St) ;
   if St1<>'' then
      BEGIN
      if St1='E_DEBIT' then BEGIN if RestitutionEnDevise then BEGIN (*If OkMajTotPDev Then*) St1:=St1+'DEV'; END ; END else
        if St1='E_CREDIT' then BEGIN if RestitutionEnDevise then BEGIN (*If OkMajTotPDev Then*) St1:=St1+'DEV'; END ; END else
        if St1='E_DEBITDEV' then BEGIN if (Not RestitutionEnDevise) (*Or (Not OkMajTotPDev)*) then St1:='E_DEBIT' ;END else
        if St1='E_CREDITDEV' then BEGIN if (Not RestitutionEnDevise) (*Or (Not OkMajTotPDev)*) then St1:='E_CREDIT' ; END ;
      T:=Q.FindField(St1) ;
      if T<>Nil then
        BEGIN
        LaDec:=DecimP ;
//        if ((St1='E_DEBITEURO') or (St1='E_CREDITEURO')) then LaDec:=DecimE ;
        if T.AsString <> '0' then begin
	        if ((FListe.ColFormats[C]<>'') and (T.DataType=ftFloat)) then StCell:=StrfMontant(T.AsFloat,0,LaDec,'',True)
  	                                                               else StCell:=T.AsString ;
          end
        else
        	StCell:='';
        if St1='E_REFINTERNE' then
           BEGIN
           StP:=GetRefParamDB(Q,VH^.CPRefPointage) ;
           if StP<>'ERROR' then StCell:=StP ;
           END ;
        FListe.Cells[C,FListe.RowCount-1]:=StCell ;
        END ;
      END ;
   inc(C) ;
   END;
END ;

procedure TFPointageNew.FListeDblClick(Sender: TObject);
begin PointeDepointe ; end;

procedure TFPointageNew.CalculPointes ( RefP : String ) ;
var StSQL : string ;
    Q     : TQuery ;
BEGIN
// calcule les cumuls pointés avec la ref RefP
if Not OkPointe then exit;
if RefP='' then
   BEGIN
   FNbP.Value:=0;
   XDebitP:=0 ; XCreditP:=0 ; XDebitD:=0 ; XCreditD:=0 ; 
   END else
   BEGIN
   StSQL:='Select SUM(E_CREDITDEV), SUM(E_DEBITDEV), SUM(E_CREDIT), SUM(E_DEBIT), '
         +'Count(E_GENERAL) From ECRITURE';
   StSQL:=StSQL+' Where E_GENERAL="'+FGene.Text+'" and E_REFPOINTAGE="'+RefP+'"';
   StSQL:=StSQL+' And E_DATEPOINTAGE="'+USDATE(FDatePointage)+'" Group by E_GENERAL';
   Q:=OpenSQL(StSQL,True) ;
   if Not Q.EOF then
      BEGIN
      (*
      XDebitD:=Q.Fields[0].AsFloat ; XCreditD:=Q.Fields[1].AsFloat ;
      XDebitP:=Q.Fields[2].AsFloat ; XCreditP:=Q.Fields[3].AsFloat ;
      XDebitE:=Q.Fields[4].AsFloat ; XCreditE:=Q.Fields[5].AsFloat ;
      *)
      XCreditD:=Q.Fields[0].AsFloat;
      XDebitD :=Q.Fields[1].AsFloat;
      XCreditP:=Q.Fields[2].AsFloat ;
      XDebitP:=Q.Fields[3].AsFloat ;
      FNbP.Value:=Q.Fields[4].AsInteger ;
      END ;
   Ferme(Q) ;
   END;
AfficheTotauxBasEcran ;
END;

procedure TFPointageNew.FRefPointageExit(Sender: TObject);
var SQL,St : String ;
    Q      : TQuery ;
begin
St:=FRefPointage.Text ;
While Pos('''',St)>0 do Delete(St,Pos('''',St),1) ;
While Pos('"',St)>0 do Delete(St,Pos('"',St),1) ;
While Pos('°',St)>0 do Delete(St,Pos('°',St),1) ;
FRefPointage.Text:=St ; If FDatePointage.Text='' Then FDatePointage.Text:=DateToStr(IDate1900) ;
if Not OkPointe then
   BEGIN
   SQL:='Select EE_DATEPOINTAGE From EEXBQ ' ;
   SQL:=SQL+' Where EE_GENERAL="'+FGene.Text+'" and Upper(EE_REFPOINTAGE)="'+FRefPointage.Text+'" and EE_DATEPOINTAGE="'+USDATE(FDatePointage)+'"';
   Q:=OpenSQL(SQL, true);
   If Q.Eof Then
     BEGIN
     Ferme(Q) ;
     SQL:='Select EE_DATEPOINTAGE From EEXBQ ' ;
     SQL:=SQL+' Where EE_GENERAL="'+FGene.Text+'" and Upper(EE_REFPOINTAGE)="'+FRefPointage.Text+'" ';
     Q:=OpenSQL(SQL, true);
     END ;
   if Not Q.EOF then
      BEGIN
      FDatePointage.Text:=Q.FindField('EE_DATEPOINTAGE').AsString ;
      Ferme(Q) ; OkRef:=True ;
      CalculPointes(Trim(FRefPointage.Text));
      END else
      BEGIN
      FDatePointage.Text:=DateToStr(IDate1900) ;
      Ferme(Q);
      OkRef:=False ;
      END ;
   END ;
end;

procedure TFPointageNew.BInfoPointageClick(Sender: TObject);
begin
ClickInfoPointage(true) ;
end;

procedure TFPointageNew.RenseigneDetail(Row: LongInt);
var OBM : TOBM;
BEGIN
if Not FDetail.Visible then Exit ;
OBM:=GetO(FListe,Row);
if OBM=NIL then
   BEGIN
   E_JOURNAL.Caption:=''; E_NUMEROPIECE.Caption:='';
   E_NATUREPIECE.Caption:=''; E_LIBELLE.Caption:='';
   E_REFEXTERNE.Caption:=''; E_MODEPAIE.Caption:='';
   E_DATEECHEANCE.Caption:=''; E_REFLIBRE.Caption:='';
   END else
   BEGIN
   E_JOURNAL.Caption:=OBM.GetMvt('E_JOURNAL');
   E_NATUREPIECE.Caption:=RechDom('ttNaturePiece',OBM.GetMvt('E_NATUREPIECE'), false);
   E_NUMEROPIECE.Caption:=OBM.GetMvt('E_NUMEROPIECE');
   E_LIBELLE.Caption:=OBM.GetMvt('E_LIBELLE');
   E_REFEXTERNE.Caption:=OBM.GetMvt('E_REFEXTERNE');
   E_MODEPAIE.Caption:=RechDom('ttModePaie',OBM.GetMvt('E_MODEPAIE'),false);
   E_DATEECHEANCE.Caption:=OBM.GetMvt('E_DATEECHEANCE');
   E_REFLIBRE.Caption:=OBM.GetMvt('E_REFLIBRE');
   END;
END;

procedure TFPointageNew.FListeSelectCell(Sender: TObject; Col, Row: Longint; var CanSelect: Boolean);
begin
//RenseigneDetail(Row); Mis sur RowEnter pour accélérer le traitement LHE le 21/05/1997
end;

procedure TFPointageNew.ZoomPiece(R : LongInt) ;
Var O : TOBM ;
    M : RMVT ;
{$IFNDEF IMP}
    P  : RParFolio ;
{$ENDIF}
BEGIN
{$IFNDEF IMP}
if Not BZoomPiece.Enabled then Exit ;
O:=GetO(FListe,R);if O=Nil then Exit ;
M:=OBMToIdent(O,False) ;
if ((M.ModeSaisieJal<>'-') and (M.ModeSaisieJal<>'')) then
   BEGIN
   FillChar(P, Sizeof(P), #0) ;
   P.ParPeriode:=DateToStr(DebutDeMois(O.GetMvt('E_DATECOMPTABLE'))) ;
   P.ParCodeJal:=O.GetMvt('E_JOURNAL') ;
   P.ParNumFolio:=IntToStr(O.GetMvt('E_NUMEROPIECE')) ;
   P.ParNumLigne:=O.GetMvt('E_NUMLIGNE') ;
   ChargeSaisieFolio(P, taConsult) ;
   END else
   BEGIN
   LanceSaisie(Nil,taConsult,M) ;
   END ;
{$ENDIF}
END ;

procedure TFPointageNew.BZoomPieceClick(Sender: TObject);
begin ZoomPiece(FListe.Row); end;

procedure TFPointageNew.EnregRefPointage;
var Q     : TQuery;
    SQL   : String;
    Trouv : boolean ;
BEGIN
SQL:='Select EE_GENERAL From EEXBQ ';
SQL:=SQL+' Where EE_GENERAL="'+FGene.Text+'" and EE_DATEPOINTAGE="'+USDATE(FDatePointage)+'"';
SQL:=SQL+' and EE_REFPOINTAGE="'+Trim(FRefPointage.Text)+'"';
Q:=OpenSQL(SQL,True) ; Trouv:=Not Q.EOF ; Ferme(Q);
if Trouv then
   BEGIN
   Exit ; 
   (*
   If V_PGI.Synap And OkPointe Then
      BEGIN
      SQL:='UPDATE EEXBQ SET EE_NEWSOLDECRE='+StrFPoint(CreditCpBanque)+', EE_NEWSOLDEDEB="'+StrFPoint(DebitCpBanque)+'"' ;
      SQL:=SQL+' Where EE_GENERAL="'+FGene.Text+'" and EE_DATEPOINTAGE="'+USDATE(FDatePointage)+'"';
      SQL:=SQL+' and EE_REFPOINTAGE="'+Trim(FRefPointage.Text)+'"';
      ExecuteSQL(SQL) ; Exit ;
      END Else exit;
   *)
   END ;
SQL:='INSERT INTO EEXBQ (EE_GENERAL, EE_REFPOINTAGE,' ;
SQL:=SQL+' EE_NEWSOLDECRE, EE_NEWSOLDEDEB, EE_NEWSOLDECREEURO, EE_NEWSOLDEDEBEURO, EE_NUMERO, EE_DATEPOINTAGE)' ;
SQL:=SQL+' VALUES ("'+FGene.Text+'", "'+Trim(FRefPointage.Text)+'", ' ;
SQL:=SQL+StrFPoint(CreditCpBanquePivot)+', '+StrFPoint(DebitCpBanquePivot)+', ' ;
SQL:=SQL+StrFPoint(CreditCpBanqueEuro)+', '+StrFPoint(DebitCpBanqueEuro)+', 1, "'+USDate(FDatePointage)+'")' ;
if ExecuteSQL(SQL)<=0 then V_PGI.IoError:=oePointage ;
END;

procedure TFPointageNew.FGeneEnter(Sender: TObject);
begin
if Not Modif then exit;
if Messages.Execute(5,'','')=mrYes then
   BEGIN
   SauveReseau;
   Modif:=false; DBNav.Enabled:=Not Modif; BValider.Enabled:=Modif; FChoixDevise.Enabled:=Not Modif ;
   END else FListe.SetFocus;
end;

procedure TFPointageNew.FGeneExit(Sender: TObject);
var
	Q : TQuery;
begin
if gbJournal then begin
  Q := OpenSQL('SELECT J_CONTREPARTIE FROM JOURNAL WHERE J_JOURNAL="'+FJournal.Text+'"',True);
  if (not Q.EOF) then FGene.Text := Q.Fields[0].AsString;
  gszJournal := fJournal.Text;
  Ferme(Q);
end;
GChercheCompte(FGene,FicheGene) ;
if FGene.ExisteH>0 then RepositionGene ;
end;

procedure TFPointageNew.POPSPopup(Sender: TObject);
begin InitPopup(Self); end;

procedure TFPointageNew.BRefPClick(Sender: TObject);
var St,Ref  : string;
    Q   : TQuery;
    StDate : String ;
    DateP : TDateTime ;
begin
// A VOIR : Rajouter un champ JOURNAL dans la table EEXBQ pour faire une recherche sur le journal et non sur le compte
if gbJournal then St:=Choisir(Messages.Mess[13],'EEXBQ','EE_REFPOINTAGE','EE_DATEPOINTAGE','EE_GENERAL="'+FGene.Text+'"','EE_DATEPOINTAGE DESC',TRUE)
             else St:=Choisir(Messages.Mess[13],'EEXBQ','EE_REFPOINTAGE','EE_DATEPOINTAGE','EE_GENERAL="'+FGene.Text+'"','EE_DATEPOINTAGE DESC',TRUE);
if St='' then Exit ;
Ref:=ReadTokenSt(St) ;
StDate:=ReadTokenSt(St) ; DateP:=StrToDate(StDate) ;
Q:=OpenSQL('select * from EEXBQ where EE_GENERAL="'+FGene.Text+'" and EE_REFPOINTAGE="'+Ref+'" and EE_DATEPOINTAGE="'+USDATETIME(DateP)+'"', true) ;
if Q.EOF then BEGIN Ferme(Q) ; Exit ; END ;
DebitCpBanquePivot:=Q.FindField('EE_NEWSOLDEDEB').AsFloat ;
CreditCpBanquePivot:=Q.FindField('EE_NEWSOLDECRE').AsFloat ;
FDebitPointagePivot.Text:=StrfMontant(DebitCpBanquePivot,15,DecimP,'',TRUE) ;
FCreditPointagePivot.Text:=StrfMontant(CreditCpBanquePivot,15,DecimP,'',TRUE) ;

DebitCpBanqueEuro:=Q.FindField('EE_NEWSOLDEDEBEURO').AsFloat ;
CreditCpBanqueEuro:=Q.FindField('EE_NEWSOLDECREEURO').AsFloat ;
FDebitPointageEuro.Text:=StrfMontant(DebitCpBanqueEuro,15,2,'',TRUE) ;
FCreditPointageEuro.Text:=StrfMontant(CreditCpBanqueEuro,15,2,'',TRUE) ;
Ferme(Q);
FDatePointage.Text:=StDate ; FRefPointage.Text:=Ref ;
if FRefPointage.CanFocus then FRefPointage.SetFocus ;
end;

procedure TFPointageNew.BEtatRapproClick(Sender: TObject);
begin  ClickEtat('RAP'); end;

procedure TFPointageNew.BitBtn1Click(Sender: TObject);
begin
PrintDBGrid(FListe,Pages,caption,'') ;
end;

procedure TFPointageNew.PointageFichier ;
BEGIN
SauvePointage ;
if V_PGI.IOError=oeOK then MajTotPointeCpt ;
END ;

procedure TFPointageNew.MajSessionMemoire ;
Var OBM : TOBM ;
    i   : integer ;
BEGIN
for i:=1 to FListe.RowCount-1 do
   BEGIN
   OBM:=GetO(FListe,i); if OBM=Nil then Break ;
   if OkPointe then BEGIN if OBM.GetMvt('E_REFPOINTAGE')='' then Continue ;END
               else BEGIN if OBM.GetMvt('E_REFPOINTAGE')<>'' then Continue ;END;
   OBM.PutMvt('OLDDEBIT',1) ;
   END ;
END ;

procedure TFPointageNew.SauveReseau;
Var ii : TIOErr ;
BEGIN
DateRef:=NowH ;
ii:=Transactions(PointageFichier,3) ;
Case ii of
   oeOK       : MajSessionMemoire ;
   oePointage : MessageAlerte(Messages.Mess[14]) ;
   oeUnknown  : MessageAlerte(Messages.Mess[15]) ;
   END ;
END;

{function TFPointage.WhereLastModif: string;
BEGIN
Result:=' E_GENERAL="'+FGene.Text+'"';
if OkPointe then Result:=Result+' and E_DATECOMPTABLE>="'+USDATE(FDateCptaDeb)+' and E_DATECOMPTABLE<="'+USDATE(FDateCptaFin)+'"';
// Result:=Result+' and E_DATEECHEANCE>="'+USDATE(FDateEche1)+'" and E_DATEECHEANCE<="'+USDATE(FDateEche2)+'"';
Result:=Result+' and E_QUALIFPIECE="N" and E_ECRANOUVEAU<>"OAN" and E_ECRANOUVEAU<>"C"';
if OkPointe then Result:=Result+' and E_REFPOINTAGE="'+Trim(FRefPointage.Text)+'"' else Result:=Result+ ' and E_REFPOINTAGE=""';
Result:=Result+' and E_DATEMODIF="'+USTime(DateRef)+'"';
END;}

procedure TFPointageNew.BMenuZoomMouseEnter(Sender: TObject);
begin PopZoom97(BMenuZoom,POPZ) ; end;

procedure TFPointageNew.BEnleveClick(Sender: TObject);
begin  CachePointes(true); end;

procedure TFPointageNew.BRajouteClick(Sender: TObject);
begin CachePointes(false) ;  end;

procedure TFPointageNew.CloseFen ;
BEGIN
if Not IsInside(Self) then Close ;
END ;

procedure TFPointageNew.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide,OKG : boolean ;
begin
OkG:=(FListe=Screen.ActiveControl) ; Vide:=(Shift=[]) ;
Case Key of
   VK_F1     : BEGIN Key:=0 ; END ;
   VK_F5     : BEGIN Key:=0 ; if Vide then ZoomCompte else if Shift=[ssShift] then ZoomPiece(FListe.Row) ; END ;
   VK_F9     : BEGIN Key:=0 ; if Vide then BAppliquerClick(Nil) ; END ;
   VK_F10    : BEGIN Key:=0 ; if Vide then ClickValide ; END ;
   VK_SPACE  : BEGIN Key:=0 ; if ((OkG) and (Vide)) then PointeDepointe END ;
   VK_RETURN : BEGIN Key:=0 ; if (Shift=[ssCtrl]) then ClickChercher ; END ;
   VK_ESCAPE : BEGIN Key:=0 ; if Vide then CloseFen ; END ;
   VK_INSERT : BEGIN Key:=0 ; if (Shift=[ssShift]) then CachePointes(false) ; END ;
   VK_DELETE : if not TuPeuxSupprimer then BEGIN Key:=0 ; if (Shift=[ssShift]) then CachePointes(true) ; END else TuPeuxSupprimer:=False ;
   67        : BEGIN Key:=0 ; if (Shift=[ssAlt]) then ClickInfoPointage(BInfoPointage.Enabled); END ;
   70        : BEGIN Key:=0 ; if (Shift=[ssCtrl]) then ClickRecherche ; END ;
   80        : BEGIN Key:=0 ; if (Shift=[ssCtrl]) then ClickEtat('POI') ; END ;
   82        : BEGIN Key:=0 ; if (Shift=[ssCtrl]) then ClickEtat('RAP') ; END ;
{$IFNDEF CCS3}
   83        : BEGIN Key:=0 ; if (Shift=[ssAlt]) then ClickSaisie ; END ;
{$ENDIF}
   END ;
end;

procedure TFPointageNew.BZoomClick(Sender: TObject);
begin  ZoomCompte;  end;

procedure TFPointageNew.ZoomCompte;
BEGIN
if Not BZoom.Enabled then Exit ;
FicheGene(Nil,'',FGene.Text,taConsult,0);
END;

procedure TFPointageNew.ClickEtat(Lequel : string);
BEGIN
if Lequel='RAP' then BEGIN if Not BEtatRappro.Enabled then Exit; END else
if Lequel='POI' then BEGIN if Not BEtatPointage.Enabled then Exit; END ;
if Modif then
   if Messages.Execute(7,'','')=mrYes then
      BEGIN
      SauveReseau; Modif:=False; BValider.Enabled:=Modif;
      END else exit;
if FGene.Text='' then exit;
if Lequel='RAP' then
   BEGIN
   Fillchar(Crit,SizeOf(Crit),#0) ;
   Crit.Date2:=StrToDate(FDatePointage.Text) ;
   Crit.NatureEtat:=neRap ;
   InitCritEdt(Crit) ;
   Crit.Cpt1:=FGene.Text ;
   Crit.Cpt2:=Crit.Cpt1 ;
   Crit.Rap.RefP:=FRefPointage.Text ;
   Crit.DeviseSelect:=FDevise.Text ;
   If Crit.DeviseSelect<>V_PGI.DevisePivot Then Crit.Monnaie:=1 ;
   EtatRapproZoom(Crit) ;
   END else
   if Lequel='POI' then
      BEGIN
      Fillchar(Crit,SizeOf(Crit),#0) ;
      Crit.NatureEtat:=nePoi ;
      InitCritEdt(Crit) ;
      Crit.Cpt1:=FGene.Text ;  Crit.Cpt2:=Crit.Cpt1 ;
      Crit.Poi.RefP1:=FRefPointage.Text ; Crit.Poi.RefP2:=Crit.Poi.RefP1 ;
      Crit.Poi.QueLesBque:=TRUE ;
      Crit.DeviseSelect:=FDevise.Text ;
      If Crit.DeviseSelect<>V_PGI.DevisePivot Then Crit.Monnaie:=1 ;
{$IFDEF CCS3}
{$ELSE}
      EtatPointageZoom(Crit) ;
{$ENDIF}
      END else
   if Lequel='JUS' then
      BEGIN
      END ;
END;

procedure TFPointageNew.ClickValide ;
BEGIN
if Not BValider.Enabled then Exit ;
if Not FListe.SynEnabled then Exit ;
GridEna(Fliste,False) ;
FListe.Enabled:=False ; SauveReseau ; FListe.Enabled:=True ;
BValiderEnCours:=TRUE ; RepositionGene ; BValiderEnCours:=FALSE ;
Modif:=false; DBNav.Enabled:=Not Modif ; BValider.Enabled:=Modif ; FChoixDevise.Enabled:=Not Modif ;
ClickChercher ;
GridEna(Fliste,True) ;
END ;

procedure TFPointageNew.PointeDepointe ;
var i     : integer ;
    GRect : TGridRect;
begin
if ListeVide then Exit ;
If ((OkPointe) and (Trim(FRefPointage.Text)='')) then
   BEGIN
   Messages.Execute(0,'','') ; Pages.ActivePage:=PStandard ;
   if FRefPointage.CanFocus then FRefPointage.SetFocus ;
   Exit ;
   END ;
GRect:=FListe.Selection ; BEtatRappro.Enabled:=True ;
For i:=GRect.Top to GRect.Bottom do
   BEGIN
   if OkPointe then EnvoiePointage(i) else EnvoieDepointage(i);
   END ;
end ;

procedure TFPointageNew.CachePointes(OkOk : boolean);
var i : integer;
BEGIN
For i:=1 to FListe.RowCount-1 do
   BEGIN
   if OkOk then BEGIN if FListe.Cells[CRefP,i]<>'' then FListe.RowHeights[i]:=0 ; END
           else BEGIN FListe.RowHeights[i]:=FListe.DefaultRowHeight ; END;
   END;
BRajoute.Enabled:= OkOk ; BEnleve.Enabled:= Not OkOk ;
END ;

procedure TFPointageNew.GrandPetitListe(OkOk : boolean);
BEGIN
ChangeListeCrit(Self,OkOk) ;
BAgrandir.visible:=Not OkOk ; BReduire.visible:=OkOk ;
END ;

procedure TFPointageNew.ClickRecherche ;
BEGIN
FirstFind:=true; FindDialog.Execute ;
END ;

procedure TFPointageNew.ClickInfoPointage (OkOk : boolean) ;
BEGIN
FDetail.Visible:=Not OkOk;
if OkPointe then FEnTete.Visible:=OkOk else FEntete.Visible:=false;
BComplement.Enabled:=OkOk ; BInfoPointage.Enabled:=Not OkOk;
BComplement.Visible:=OkOk ; BInfoPointage.Visible:=Not OkOk;
if Not OkOk then RenseigneDetail(FListe.Selection.Top);
END ;

procedure TFPointageNew.ClickSaisie ;
var M   : RMVT;
BEGIN
if Not BLanceSaisie.Enabled then Exit ;
{$IFNDEF CCS3}
if Trim(FRefPointage.Text)='' then BEGIN Messages.Execute(8,'','') ; Exit ; END ;
if _Blocage(['nrPointage'],True,'nrAucun') then Exit ;
// valider la session
if Modif then
   if Messages.Execute(7,'','')=mrYes then
      BEGIN
      SauveReseau; Modif:=False; BValider.Enabled:=Modif;
      END else exit;

// lancer saisie tréso
If BalanceSaisie(M) Then RecupereSaisie(M);
                         // mettre à jour liste (affiche les lignes saisies et pointées en fin de liste)
{$ENDIF}
END ;

procedure TFPointageNew.ClickChercher ;
var Rep  : integer;
    SQL  : String ;
    OkOk : Boolean ;
    DD   : TDateTime ;
    Q : Tquery ;
BEGIN
BAppliquerAFaire:=FALSE ;
EnableControls(Self,False) ;
if Modif then
   BEGIN
   Rep:= Messages.Execute(2,'','');
   if Rep=mrYes then SauveReseau else if Rep=mrCancel then BEGIN EnableControls(Self,True) ; Exit; END ;
   END ;
If ((Not OkPointe) and (FRefPointage.Text='')) then BEGIN Messages.Execute(19,'','') ; EnableControls(Self,True) ; Exit ; END ;
If ((Not OkPointe) and (Not OkRef)) then
   BEGIN
   Messages.Execute(18,'','') ; EnableControls(Self,True) ;
   if FRefPointage.CanFocus then FRefPointage.SetFocus ;
   Exit ;
   END ;

If OkPointe And (Trim(FRefPointage.Text)<>'') Then
  BEGIN
  SQL:='Select EE_DATEPOINTAGE From EEXBQ Where EE_GENERAL="'+FGene.Text+'" and Upper(EE_REFPOINTAGE)="'+FRefPointage.Text+'" ' ;
  OkOk:=TRUE ; Q:=OpenSQL(SQL, true); If Not Q.Eof Then OkOk:=FALSE ;
  DD:=StrToDate(FDatePointage.Text) ;
  While (Not Q.Eof) And (Not OkOk) Do BEGIN If Q.Fields[0].AsDateTime=DD Then OkOk:=TRUE ; Q.Next ; END ;
  Ferme(Q) ;
  If Not OkOk Then
    BEGIN
    Messages.Execute(22,'','') ; EnableControls(Self,True) ;
    if FRefPointage.CanFocus then FRefPointage.SetFocus ;
    Exit ;
    END ;
  END ;

If OKPointe Then
  BEGIN
  OkOk:=TRUE ;
  If (Not CptEnPivot) And ((FChoixDevise.ItemIndex=0) Or (FDevise.Text<>FChoixDevise.Value)) Then OkOk:=FALSE ;
  If Not OkOk Then
    BEGIN
    Messages.Execute(23,'','') ; EnableControls(Self,True) ;
    if FRefPointage.CanFocus then FRefPointage.SetFocus ;
    Exit ;
    END ;
  END ;

(*
MAJCredPivot:=QGene.FindField('G_DEBNONPOINTE').AsFloat  ; MAJDebPivot:=QGene.FindField('G_CREDNONPOINTE').AsFloat;
MAJCredEuro:=QGene.FindField('G_DEBNONPOINTEEU').AsFloat ; MAJDebEuro:=QGene.FindField('G_CREDNONPOINTEEU').AsFloat;
*)
//ChargeInfosCompte ;
ChercheEcritures; EnableControls(Self,True) ;
FListe.Setfocus ;
Modif:=false ; DBNav.Enabled:=Not Modif ; BValider.Enabled:=Modif ; FChoixDevise.Enabled:=Not Modif ; RenseigneDetail(FListe.Row);
END ;

procedure TFPointageNew.BEtatPointageClick(Sender: TObject);
begin ClickEtat('POI') ; end;

function TFPointageNew.ChercheJal(Gene:string): string;
var Q : TQuery ;
BEGIN
Result:='' ;
Q:=OpenSQL('Select J_JOURNAL From JOURNAL Where J_NATUREJAL="BQE" And J_CONTREPARTIE="'+Gene+'" ',true);
if Not Q.EOF then Result:=Q.Fields[0].AsString ;
Ferme(Q);
END;

procedure TFPointageNew.DBNavClick(Sender: TObject; Button: TNavigateBtn);
begin
if FGene.ExisteH>0 then
  BEGIN
  RepositionGene ;
  ChargeInfosCompte ;
  FRefPointage.Text:='' ;
  END ;
end;

procedure TFPointageNew.FListeMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Var C,R : Longint ;
begin
GX:=X ; GY:=Y ;
if ((ssCtrl in Shift) and (Button=mbLeft)) then
   BEGIN
   FListe.MouseToCell(X,Y,C,R) ;
   if ((R>0) and (R<FListe.RowCount-1)) then PointeDepointe ;
   END ;
end;

procedure TFPointageNew.FRefPointageKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin TuPeuxSupprimer:=True ; end;

procedure TFPointageNew.FormCreate(Sender: TObject);
begin WMinX:=Width ; WMinY:=Height ; end;

procedure TFPointageNew.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFPointageNew.FListeRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
if FDetail.Visible then RenseigneDetail(FListe.Row);
end;

procedure TFPointageNew.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFPointageNew.BParamClick(Sender: TObject);
begin
ExtraitBanquaire ;
end;

{procedure TFPointageNew.CAfficheOpposeClick(Sender: TObject);
begin
if CAfficheOppose.Checked then
   BEGIN
   FDebitDPIniA.Value:=LesTotauxCpt[tsmEuro,0].Deb ; FCreditDPIniA.Value:=LesTotauxCpt[tsmEuro,0].Cre ;
   END else if RestitutionEnDevise And OkMajTotPDev then
   BEGIN
   FDebitDPIniA.Value:=LesTotauxCpt[tsmDevise,0].Deb ; FCreditDPIniA.Value:=LesTotauxCpt[tsmDevise,0].Cre ;
   END else
   BEGIN
   FDebitDPIniA.Value:=LesTotauxCpt[tsmPivot,0].Deb ; FCreditDPIniA.Value:=LesTotauxCpt[tsmPivot,0].Cre ;
   END ;
AfficheTotauxBasEcran ;
AfficheTotauxPointeHautDEcran ;
//CalculRestePointe ;
AfficheTotalDepointe ;
ColorOpposeEuro(FListe,CAfficheOppose.Checked,RestitutionEnDevise) ;
ISigneEuro1.Visible:=ISigneEuro.Visible ;
end;}

procedure TFPointageNew.FChoixDeviseChange(Sender: TObject);
begin
BAppliquerAFaire:=TRUE ;
InitPourDevise ;
end;

procedure TFPointageNew.FGeneChange(Sender: TObject);
begin
BAppliquerAFaire:=TRUE ;
end;

procedure TFPointageNew.FListeEnter(Sender: TObject);
begin
If BAppliquerAFaire Then ClickChercher ;
end;

procedure TFPointageNew.FChoixDeviseEnter(Sender: TObject);
begin
if Not Modif then exit;
if Messages.Execute(5,'','')=mrYes then
   BEGIN
   SauveReseau;
   Modif:=false; DBNav.Enabled:=Not Modif; BValider.Enabled:=Modif; FChoixDevise.Enabled:=Not Modif ;
   END else FListe.SetFocus;
end;

end.
