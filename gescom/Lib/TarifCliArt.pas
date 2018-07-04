unit TarifCliArt;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, ExtCtrls, HPanel, HTB97, Menus, StdCtrls, Mask, HSysMenu,
  hmsgbox, UIUtil, Hent1, Ent1, TarifUtil, SaisUtil, UTOB,Hqry,
{$IFDEF EAGLCLIENT}
  MaineAGL,eFiche,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,Fiche,
{$ENDIF}
  ComCtrls,math, HDimension, AglInit, TarifRapide, HRichEdt, HRichOLE,
{$IFNDEF CCS3}
  TarifCond,
{$ENDIF}
  UtilArticle, AglInitGC, M3FP, EntGC, TntGrids, TntStdCtrls, TntComCtrls,
  TntExtCtrls ;

Function EntreeTarifCliArt (Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
Function EntreeTarifFouArt (Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
Function SaisieTarifCliArt (NatTiers, CodeT : string; Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
Procedure SaisieTarifCliArtKnown (NatTiers, CodeT ,Article: string; Action : TActionFiche; TarifTTC : Boolean=False); // : boolean ;


Type T_ActionValid = (tavArt, tavFam, tavTout) ;

type
  TFTarifCliArt = class(TForm)
    Dock971: TDock97;
    ToolBar972: TToolWindow97;
    PENTETE: THPanel;
    PPIED: THPanel;
    GF_TIERS: THCritMaskEdit;
    TGF_TIERS: THLabel;
    TGF_LIBELLE: THLabel;
    GF_DEVISE: THValComboBox;
    TGF_DEVISE: THLabel;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    BChercher: TToolbarButton97;
    BInfos: TToolbarButton97;
    POPZ: TPopupMenu;
    InfArticle: TMenuItem;
    InfClient: TMenuItem;
    MsgBox: THMsgBox;
    HMTrad: THSystemMenu;
    FindLigne: TFindDialog;
    TGF_REMISE: THLabel;
    TGF_CASCADEREMISE: THLabel;
    GF_REMISE: THCritMaskEdit;
    GF_CASCADEREMISE: THValComboBox;
    ENREGAUTO: TCheckBox;
    HMess: THMsgBox;
    BSaisieRapide: TToolbarButton97;
    PTITRE: THPanel;
    PFAMARTICLE: THPanel;
    G_Fam: THGrid;
    HTitre: THMsgBox;
    PTARIFART: THPanel;
    PARTICLE: THPanel;
    TGF_ARTICLE: THLabel;
    TGF_LIBELLEART: TLabel;
    GF_CODEARTICLE: THCritMaskEdit;
    PDIMENSION: THPanel;
    TGF_GRILLEDIM1: THLabel;
    TGF_GRILLEDIM2: THLabel;
    TGF_GRILLEDIM3: THLabel;
    TGF_GRILLEDIM4: THLabel;
    TGF_GRILLEDIM5: THLabel;
    GF_CODEDIM1: THCritMaskEdit;
    GF_CODEDIM2: THCritMaskEdit;
    GF_CODEDIM3: THCritMaskEdit;
    GF_CODEDIM4: THCritMaskEdit;
    GF_CODEDIM5: THCritMaskEdit;
    PQTEART: THPanel;
    PTART: THPanel;
    G_Qte: THGrid;
    G_TA: THGrid;
    CBQUANTITATIF: TCheckBox;
    CONDAPPLIC: THRichEditOLE;
    CBARTICLE: TCheckBox;
    TCONDTARF: TToolWindow97;
    G_COND: THGrid;
    FComboTIE: THValComboBox;
    FComboART: THValComboBox;
    FComboLIG: THValComboBox;
    FComboPIE: THValComboBox;
    GF_CONDAPPLIC: THRichEditOLE;
    FTable: THValComboBox;
    FOpe: THValComboBox;
    BCollerCond: TToolbarButton97;
    BCopierCond: TToolbarButton97;
    BCondAplli: TToolbarButton97;
    BVoirCond: TToolbarButton97;
    TTYPETARIF: THLabel;
    ISigneEuro: TImage;
    TGF_PRIXCON: THLabel;
    ISigneFranc: TImage;
    GF_PRIXCON: THNumEdit;
    GA_PRIXPOURQTE: THNumEdit;
    TGA_PRIXPOURQTE: THLabel;
    ImportAchat: TToolbarButton97;
    POPE: TPopupMenu;
    MnuTarifFrs: TMenuItem;
    MnuPrixNet: TMenuItem; //JT eQualité 10807
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GF_DEVISEChange(Sender: TObject);
    procedure GF_TIERSExit(Sender: TObject);
    procedure GF_TIERSElipsisClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure InfClientClick(Sender: TObject);
    procedure GF_REMISEChange(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    procedure G_QteCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_QteCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_QteEnter(Sender: TObject);
    procedure G_QteRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_QteRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
    procedure G_QteElipsisClick(Sender: TObject);
    procedure GF_CODEARTICLEExit(Sender: TObject);
    procedure GF_CODEARTICLEElipsisClick(Sender: TObject);
    procedure InfArticleClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure G_FamCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_TACellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_TAEnter(Sender: TObject);
    procedure G_FamEnter(Sender: TObject);
    procedure G_TARowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_FamRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_TARowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_FamRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_TACellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_FamCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_TAElipsisClick(Sender: TObject);
    procedure G_FamElipsisClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BSaisieRapideClick(Sender: TObject);
    procedure CBQUANTITATIFClick(Sender: TObject);
    procedure BCondAplliClick(Sender: TObject);
    procedure BCopierCondClick(Sender: TObject);
    procedure BCollerCondClick(Sender: TObject);
    procedure BVoirCondClick(Sender: TObject);
    procedure CBARTICLEClick(Sender: TObject);
    procedure TCONDTARFClose(Sender: TObject);
    procedure GF_CASCADEREMISEChange(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure ReInitSousFamille (Arow : integer;ttd : T_TableTarif);
    procedure ImportAchatClick(Sender: TObject);
    procedure ImportPrixNetClick(Sender: TObject);
  private
    { Déclarations privées }
    iTableLigne : integer ;
    FindDebut,FClosing : Boolean;
    StCellCur,LesColQtes, LesColTA, LesColFam: String ;
    ColsInter : Array of boolean ;
    DEV       : RDEVISE ;
    TarifTTC  : Boolean ;
    NatureTiers : string ;
    PrixPourQte : Double;
// Objets mémoire
    TOBTarif, TOBTarfQte, TOBTarfTA, TOBTarfFam, TOBTarifDelArt, TOBTarifDelFam, TOBTiers, TOBArt  : TOB;
// Menu
    procedure AffectMenuCondApplic (G_GRID : THGrid; ttd : T_TableTarif) ;
// Actions liées au Grid
    procedure EtudieColsListe ;
    procedure PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    procedure FormateZoneSaisie (ACol,ARow : Longint; ttd : T_TableTarif) ;
    procedure InsertLigne (ARow : Longint; ttd : T_TableTarif) ;
    procedure SupprimeLigne (ARow : Longint; ttd : T_TableTarif) ;
    procedure SupprimeTOBTarif (ARow : Longint; ttd : T_TableTarif) ;
    Function  GrilleModifie : Boolean;
    Function  SortDeLaLigne (ttd : T_TableTarif) : boolean ;
// Initialisations
    Procedure InitialiseCols ;
    procedure LoadLesTOB (tav: T_ActionValid);
    procedure ChargeTarifTout;
    procedure ChargeTarifArt;
    procedure ChargeTarifFam;
    procedure ChargeTarif (tav: T_ActionValid);
    procedure ErgoGCS3 ;
// ENTETE
    Procedure PrepareEntete ;
    Procedure InitialiseEntete ;
    Procedure AffecteEntete ;
    Procedure ChercherTiers ;
    Procedure TraiterTiers ;
    Function  QuestionTarifEnCours (tav: T_ActionValid): Integer;
// ARTICLE
    Procedure InitialiseEnteteArticle ;
    Procedure AffecteEnteteArticle ;
    Procedure AffecteDimension;
    Procedure ChercherArticle ;
    Procedure TraiterArticle ;
// LIGNES
    Procedure InitLaLigne (ARow : integer;  ttd : T_TableTarif) ;
    Function  GetTOBLigne (ARow : integer;  ttd : T_TableTarif) : TOB ;
    {$IFDEF MONCODE}
    procedure AffichePrix (TOBL : TOB ; ARow : integer; ttd : T_TableTarif) ;
    {$ENDIF}
    procedure AfficheLaLigne (ARow : integer;  ttd : T_TableTarif) ;
    procedure InitialiseGrille (tav : T_ActionValid);
    procedure InitialiseLigne (ARow : integer; ttd : T_TableTarif) ;
    Procedure DepileTOBLigne (tav : T_ActionValid);
    Procedure CreerTOBLigne (ARow : integer; ttd : T_TableTarif);
    Function  LigneVide ( ARow : integer; ttd : T_TableTarif) : Boolean;
    Procedure PreAffecteLigne (ARow : integer; ttd : T_TableTarif);
// CHAMPS LIGNES
    procedure TraiterDepot (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterFamille (ACol, ARow : integer; ttd : T_TableTarif; var Cancel: Boolean);
    procedure TraiterSousFamille (ACol, ARow : integer; ttd : T_TableTarif; var Cancel: Boolean);
    procedure TraiterLibelle (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterBorneInf (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterBorneSup (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterPrix (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterRemise (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterDateDeb (ACol, ARow : integer; ttd : T_TableTarif);
    procedure TraiterDateFin (ACol, ARow : integer; ttd : T_TableTarif);
/// PIED
    Procedure InitialisePied ;
    Procedure AffichePrixCon ;
    Procedure AffichePied (ttd : T_TableTarif) ;
// Boutons
    procedure VoirFicheArticle;
    procedure VoirFicheTiers;
// Validations
    procedure ValideTarifTout;
    procedure ValideTarifArticle;
    procedure ValideTarifFamille;
    procedure VerifLesTOB (tav : T_ActionValid);
    procedure VerifLesTOBArticle (MaxTarif : integer);
    procedure VerifLesTOBFamille (MaxTarif : integer);
// Conditions tarifaires
    procedure InitComboChamps ;
    procedure RemplitComboChamps(NomTable : String ; FCombo : THValComboBox) ;
    procedure EffaceGrid ;
    Procedure AfficheCondTarf (ARow : Longint; ttd : T_TableTarif) ;
    function  ValueToItem(CC : THValComboBox ; St : String) : String ;
    procedure GetConditions (TOBL : TOB) ;
  public
    { Déclarations publiques }
    CodeTiers, CodeArticle, CodeDevise  : string;
    FicheAncetre : Boolean;
    Action      : TActionFiche ;
    TheArticle : string;
    TiersIsClient : boolean; // JT eQualité 10807
  end;


Const SQ_Depot     : integer = 0 ;
      SQ_Lib       : integer = 0 ;
      SQ_Qinf      : integer = 0 ;
      SQ_QSup      : integer = 0 ;
      SQ_Px        : integer = 0 ;
      SQ_Rem       : integer = 0 ;
      SQ_Datedeb   : integer = 0 ;
      SQ_Datefin   : integer = 0 ;

Const ST_Depot     : integer = 0 ;
      ST_Lib       : integer = 0 ;
      ST_Px        : integer = 0 ;
      ST_Rem       : integer = 0 ;
      ST_Datedeb   : integer = 0 ;
      ST_Datefin   : integer = 0 ;

Const SF_Depot     : integer = 0 ;
      SF_Fam       : integer = 0 ;
      SF_SFam      : integer = 0 ;
      SF_LibSFam   : integer = 0 ;
      SF_Lib       : integer = 0 ;
      SF_Px        : integer = 0 ;
      SF_Rem       : integer = 0 ;
      SF_Datedeb   : integer = 0 ;
      SF_Datefin   : integer = 0 ;


implementation

uses         
   ParamSoc, 
   Tarifs
  ,CbpMCD
   ;

{$R *.DFM}

{***********A.G.L.***********************************************
Auteur  ...... : Michel Richaud
Créé le ...... : 12/05/2000
Modifié le ... : 12/05/2000
Description .. : Saisie des tarifs par Clients ou fournisseurs
Mots clefs ... : TARIF; TIERS
*****************************************************************}
Procedure AppeiTarifCliArt ( Parms : array of variant ; nb : integer ) ;
var
  sFonctionnalite : string;
  StNat, StTiers, StT, StM, stTiersOuTarifTiers : string ;
//  i_ind     : integer ;
  Action    : TActionFiche ;
  TarifTTC  : Boolean ;
  F         : TFFiche;
BEGIN
  F := TFFiche (Longint (Parms[0]));
  Action:=F.TypeAction;

  StNat               :=String(Parms[1]);
  StTiers             :=String(Parms[2]);
  StT                 :=String(Parms[3]);
  StM                 :=String(Parms[4]);
  stTiersOuTarifTiers :=String(Parms[5]);

  if StTiers='' then Action:=taConsult ;
  TarifTTC:= (StT = 'TTC');

  {i_ind:=Pos('ACTION=',StM) ;
  if i_ind>0 then
     BEGIN
     Delete(StM,1,i_ind+6) ; StM:=uppercase(ReadTokenSt(StM)) ;
     if StM='CREATION' then BEGIN Action:=taCreat ; END ;
     if StM='MODIFICATION' then BEGIN Action:=taModif ; END ;
     if StM='CONSULTATION' then BEGIN Action:=taConsult ; END ;
     END ;}

  if GetParamSoc('SO_PREFSYSTTARIF') then                                                                                                                                                                                 
  begin                                                                                                                                                                                                                   
     sFonctionnalite:='';                                                                                                                                                                                                 
     if       (stNat='FOU') then sFonctionnalite:=sTarifFournisseur                                                                                                                                                       
     else if  (sTNat='CLI') then sFonctionnalite:=sTarifClient                                                                                                                                                            
     else                        PGIError('Fonctionnalité non disponible, opération abandonnée');                                                                                                                         
     if (sFonctionnalite<>'') then                                                                                                                                                                                        
     begin                                                                                                                                                                                                                
        if (stTiersOuTarifTiers='TARIFTIERS') then                                                                                                                                                                        
           AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_TARIFTIERS='+stTiers,'','ACTION=CREATION    ;DROIT=CDAMV;MONOFICHE;YTA_FONCTIONNALITE='+sFonctionnalite+';APPEL=TIERS')        
        else                                                                                                                                                                                                              
           AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE='+sFonctionnalite+';YTA_TIERS='+stTiers     ,'','ACTION=CREATION    ;DROIT=CDAMV;MONOFICHE;YTA_FONCTIONNALITE='+sFonctionnalite+';APPEL=TIERS');
     end;
  end
  else
  begin
    if Action=tacreat then
    begin
      if StNat = 'FOU' then EntreeTarifFouArt (Action, TarifTTC)
                       else EntreeTarifCliArt (Action, TarifTTC) ;
    end
    else
    begin
      SaisieTarifCliArt (StNat, StTiers, Action, TarifTTC) ;
    end;
  end;
END ;

Function EntreeTarifCliArt (Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
BEGIN
Result := SaisieTarifCliArt ('CLI', '', Action, TarifTTC) ;
END;

Function EntreeTarifFouArt (Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
BEGIN
Result := SaisieTarifCliArt ('FOU', '', Action, TarifTTC) ;
END;

Function SaisieTarifCliArt (NatTiers, CodeT : string; Action : TActionFiche; TarifTTC : Boolean=False) : boolean ;
var FF : TFTarifCliArt ;
    PPANEL  : THPanel ;
begin
SourisSablier;
FF := TFTarifCliArt.Create(Application) ;
FF.Action:=Action ;
FF.TarifTTC:=TarifTTC ;
FF.NatureTiers:=NatTiers;
FF.CodeTiers:=CodeT ;
if NatTiers = 'CLI' then FF.HelpContext := 110000061
else FF.HelpContext := 110000206;
if CodeT<>'' then FF.FicheAncetre:=True else FF.FicheAncetre:=False ;
if FF.NatureTiers='FOU' then
    BEGIN
    FF.Caption := FF.HTitre.Mess[8];
    FF.InfClient.Caption := FF.HTitre.Mess[12];
    FF.TGF_TIERS.Caption := FF.HTitre.Mess[10];
    FF.GF_TIERS.DataType:='GCTIERSFOURN';
    END else
    BEGIN
    FF.Caption := FF.HTitre.Mess[7] ;
    FF.InfClient.Caption := FF.HTitre.Mess[11];
    FF.TGF_TIERS.Caption := FF.HTitre.Mess[9];
    ff.GF_TIERS.DataType:='GCTIERSCLI';
    END;
PPANEL := FindInsidePanel ;
if PPANEL = Nil then
   BEGIN
    try
      FF.ShowModal ;
    finally
      FF.Free ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside (FF, PPANEL) ;
   FF.Show ;
   END ;
Result := True;
END ;

procedure TFTarifCliArt.PostDrawCell(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
begin
//
end ;

{==============================================================================================}
{======================================= Initialisations ======================================}
{==============================================================================================}
Procedure TFTarifCliArt.InitialiseCols ;
BEGIN
SQ_Depot:=-1 ; SQ_Lib:=-1 ; SQ_QInf:=-1 ; SQ_QSup:=-1 ;
SQ_Px:=-1 ; SQ_Rem:=-1 ; SQ_Datedeb:=-1 ;  SQ_Datefin:=-1 ;
ST_Depot:=-1 ; ST_Lib:=-1 ; ST_Px:=-1 ;
ST_Rem:=-1 ; ST_Datedeb:=-1 ;  ST_Datefin:=-1 ;
SF_Depot:=-1 ; SF_Fam := -1; SF_SFAM := -1 ; SF_LibSFam:=-1 ;SF_Lib:=-1 ;
SF_Px:=-1 ; SF_Rem:=-1 ; SF_Datedeb:=-1 ;  SF_Datefin:=-1 ;
END ;

procedure TFTarifCliArt.LoadLesTOB (tav: T_ActionValid);
Var QQ        : TQuery ;
    i_ind     : integer;
    WhereTTC  : String ;
    StSQL     : String;
BEGIN
if (tav = tavArt) or (tav = tavTout) then
    BEGIN
    for i_ind := TOBTarfQte.Detail.Count - 1 downto 0 do
        BEGIN
        TOBTarfQte.Detail[i_ind].Free;
        END;

    for i_ind := TOBTarfTA.Detail.Count - 1 downto 0 do
        BEGIN
        TOBTarfTA.Detail[i_ind].Free;
        END;

    if TarifTTC then WhereTTC := ' AND GF_REGIMEPRIX = "TTC" '
                else WhereTTC := ' AND GF_REGIMEPRIX <> "TTC" ' ;
    // Lecture Quantitatif

    QQ := OpenSQL('SELECT * FROM TARIF WHERE '+
                  WhereTarifCli (CodeTiers, CodeArticle, CodeDevise, ttdCliQte,False)+
                  WhereTTC + ' ORDER BY GF_DEPOT, GF_BORNEINF',True,-1,'',true) ;
    TOBTarfQte.LoadDetailDB('TARIF','','',QQ,False) ;
    Ferme(QQ) ;
    QQ := OpenSQL('SELECT * FROM TARIF WHERE '+
                  WhereTarifCli (CodeTiers, CodeArticle, CodeDevise, ttdCliArt,False)+
                  WhereTTC + ' ORDER BY GF_DEPOT, GF_DATEDEBUT',True,-1,'',true) ;
    TOBTarfTA.LoadDetailDB('TARIF','','',QQ,False) ;
    Ferme(QQ) ;
    END;
if (tav = tavFam) or (tav = tavTout) then
    BEGIN
    for i_ind := TOBTarfFam.Detail.Count - 1 downto 0 do
        BEGIN
        TOBTarfFam.Detail[i_ind].Free;
        END;
    //QQ := OpenSQL('SELECT TARIF.*, CC_LIBELLE FROM TARIF LEFT JOIN CHOIXCOD ON CC_TYPE="BFT" AND CC_CODE=GF_SOUSFAMTARART WHERE '+
    StSQl := 'SELECT TARIF.*, BSF_LIBELLE FROM TARIF LEFT JOIN BTSOUSFAMILLETARIF ';
    StSQL := StSQL + '    ON BSF_FAMILLETARIF=GF_TARIFARTICLE AND BSF_SOUSFAMTARART=GF_SOUSFAMTARART ';
    StSQl := StSQL + ' WHERE ' + WhereTarifCli (CodeTiers, '', CodeDevise, ttdCliFam, False);
    StSQL := StSQL + ' ORDER BY GF_TARIFARTICLE, GF_SOUSFAMTARART, GF_ARTICLE, GF_BORNEINF';
    QQ := OpenSQL(StSQL,True,-1,'',true) ;
    TOBTarfFam.LoadDetailDB('TARIF','','',QQ,False) ;
    Ferme(QQ) ;
    END;
END ;

procedure TFTarifCliArt.ChargeTarifTout;
BEGIN
ChargeTarif (tavTout) ;
END;

procedure TFTarifCliArt.ChargeTarifArt;
BEGIN
ChargeTarif (tavArt) ;
END;

procedure TFTarifCliArt.ChargeTarifFam ;
BEGIN
ChargeTarif (tavFam) ;
END;

procedure TFTarifCliArt.ChargeTarif  (tav: T_ActionValid);
var i_ind : integer;
BEGIN
LoadLesTOB (tav);
if (tav = tavArt) or (tav = tavTout) then
    BEGIN
    for i_ind:=0 to TOBTarfQte.Detail.Count-1 do
        BEGIN
        // Affichage
        AfficheLaLigne (i_ind + 1, ttdCliQte) ;
        END ;
    for i_ind:=0 to TOBTarfTA.Detail.Count-1 do
        BEGIN
        // Affichage
        AfficheLaLigne (i_ind + 1, ttdCliArt) ;
        END ;
    END ;
if (tav = tavFam) or (tav = tavTout) then
    BEGIN
    for i_ind:=0 to TOBTarfFam.Detail.Count-1 do
        BEGIN
        // Affichage
        AfficheLaLigne (i_ind + 1, ttdCliFam) ;
        END ;
    END;

if TOBTarfQte.detail.count < 1 then G_Qte.RowCount := NbRowsInit
										  				 else G_QTE.RowCount := TOBTarfQte.detail.count +2;

if TOBTarfTa.detail.count < 1 then G_TA.RowCount := NbRowsInit
										 					else G_TA.RowCount := TOBTarfTa.detail.count +2;


if TOBTarfFam.detail.count < 1 then G_FAM.RowCount := NbRowsInit
															 else G_FAM.RowCount := TOBTarfFam.detail.count +2;
END ;

{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}

procedure TFTarifCliArt.FormCreate(Sender: TObject);
begin
G_Qte.RowCount := NbRowsInit ;
G_TA.RowCount := NbRowsInit ;
G_Fam.RowCount := NbRowsInit ;
StCellCur := '' ;
iTableLigne := PrefixeToNum('GF') ;
TOBTarif := TOB.Create ('', Nil, -1) ;
TOBTarfQte := TOB.Create ('', TOBTarif, 0) ;
TOBTarfTA := TOB.Create ('', TOBTarif, 1) ;
TOBTarfFam := TOB.Create ('', TOBTarif, 2) ;
TOBTarifDelArt := TOB.Create ('', Nil, -1) ;
TOBTarifDelFam := TOB.Create ('', Nil, -1) ;
TOBTiers := TOB.Create ('TIERS', Nil, -1) ;
TOBArt := TOB.Create ('ARTICLE', Nil, -1) ;
InitialiseCols ;
FClosing:=False ;
end;

procedure TFTarifCliArt.ErgoGCS3 ;
BEGIN
{$IFDEF CCS3}
BVoirCond.Visible:=False   ; BCondAplli.Visible:=False ;
BCopierCond.Visible:=False ; BCollerCond.Visible:=False ;
{$ENDIF}
END ;

procedure TFTarifCliArt.FormShow(Sender: TObject);
begin
if Action = taConsult then BValider.Visible:=False;
If (ctxAffaire in V_PGI.PGIContexte) then G_Qte.ListeParam:='AFTARIFQTEPRIX'
else G_Qte.ListeParam:='GCTARIFQTEPRIX' ;
G_Qte.PostDrawCell:=PostDrawCell ;
If (ctxAffaire in V_PGI.PGIContexte) then G_TA.ListeParam:='AFTARIFPRIX'
else G_TA.ListeParam:='GCTARIFPRIX' ;
G_TA.PostDrawCell:=PostDrawCell ;
If (ctxAffaire in V_PGI.PGIContexte) then
begin
{$IFDEF BTP}
	G_Fam.ListeParam:='BTTARIFCA';
{$ELSE}
	G_Fam.ListeParam:='AFTARIFCA';
{$ENDIF}
end else G_Fam.ListeParam:='GCTARIFCA' ;
G_Fam.PostDrawCell:=PostDrawCell ;
EtudieColsListe ;
HMTrad.ResizeGridColumns (G_Qte) ;
HMTrad.ResizeGridColumns (G_TA) ;
HMTrad.ResizeGridColumns (G_Fam) ;
AffecteGrid (G_Qte,Action) ;
AffecteGrid (G_TA,Action) ;
AffecteGrid (G_Fam,Action) ;
PFAMARTICLE.Visible := True;
PTARIFART.Visible := False;
PQTEART.Visible := False ;
PTART.Visible := False ;
PTITRE.Caption := HTitre.Mess[0] ;
TTYPETARIF.Caption := HTitre.Mess[4] ;
CBARTICLE.Checked := False;
ENREGAUTO.Visible := CBARTICLE.Checked;
CBQUANTITATIF.Checked := False ;
CodeDevise := V_PGI.DevisePivot;
DEV.Code := CodeDevise ;
GetInfosDevise (DEV) ;

if CodeTiers <> '' then
  PrepareEntete
else
  InitialiseEntete ;

if TarifTTC then
begin
  CBARTICLE.Checked := True;
  CBARTICLE.Enabled := False;
end;
InitComboChamps ;
ErgoGCS3 ;
end;

procedure TFTarifCliArt.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
if Action=taConsult then Exit ;
if GrilleModifie then
   BEGIN
   if MsgBox.Execute(6,Caption,'')<>mrYes then CanClose:=False ;
   END ;
end;

procedure TFTarifCliArt.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
G_Qte.VidePile(True) ;
G_TA.VidePile(True) ;
G_Fam.VidePile(True) ;
TOBTarif.Free ; TOBTarif:=Nil ;
TOBTarifDelArt.Free ; TOBTarifDelArt:=Nil ;
TOBTarifDelFam.Free ; TOBTarifDelFam:=Nil ;
TOBArt.Free ; TOBArt:=Nil ;
TOBTiers.Free ; TOBTiers:=Nil ;
if IsInside(Self) then Action:=caFree ;
FClosing:=True ;
end;

procedure TFTarifCliArt.FormKeyDown(Sender: TObject; var Key: Word;
                                    Shift: TShiftState);
var FocusGrid : Boolean;
    ttd : T_TableTarif;
    ARow : Longint;
BEGIN
FocusGrid := False;
ttd := ttdCliQte;
ARow := -1;
if(Screen.ActiveControl = G_Qte) then
    BEGIN
    FocusGrid := True;
    ttd := ttdCliQte;
    ARow := G_Qte.Row;
    END else
    if (Screen.ActiveControl = G_TA) then
        BEGIN
        FocusGrid := True;
        ttd := ttdCliArt;
        ARow := G_TA.Row;
        END else
        if (Screen.ActiveControl = G_Fam) then
            BEGIN
            FocusGrid := True;
            ttd := ttdCliFam;
            ARow := G_Fam.Row;
            END;
Case Key of
    VK_RETURN : Key:=VK_TAB ;
    VK_INSERT : BEGIN
                if (FocusGrid) and (ARow <> -1) then
                    BEGIN
                    Key := 0;
                    InsertLigne (ARow, ttd);
                    END;
                END;
    VK_DELETE : BEGIN
                if ((FocusGrid) and (Shift=[ssCtrl])) then
                    BEGIN
                    Key := 0 ;
                    SupprimeLigne (ARow, ttd) ;
                    END ;
                END;
    END;
END;

{==============================================================================================}
{============================= Manipulation liées au Menu =====================================}
{==============================================================================================}
procedure TFTarifCliArt.AffectMenuCondApplic (G_GRID : THGrid; ttd : T_TableTarif) ;
BEGIN
if (G_GRID.Row<>0) and ((LigneVide (G_GRID.Row, ttd)) or (Action = taConsult)) then
    BEGIN
    BCondAplli.Enabled := False ; BCopierCond.Enabled := False; BCollerCond.Enabled := False ;
    END else
    BEGIN
    if not G_GRID.enabled then G_GRID.enabled := true;
    G_GRID.SetFocus;
    BCondAplli.Enabled := True ; BCopierCond.Enabled := True;
    if CONDAPPLIC.Text = '' then BCollerCond.Enabled := False
                            else BCollerCond.Enabled := True ;
    END;
END;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}
procedure TFTarifCliArt.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp, i_ind : integer ;
		Mcd : IMCDServiceCOM;
    Control : string;
BEGIN
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();

G_Qte.ColWidths[0]:=0 ;
G_TA.ColWidths[0]:=0 ;
G_Fam.ColWidths[0]:=0 ;
SetLength(ColsInter,G_Qte.ColCount) ;
for i_ind:=Low(ColsInter)to High(ColsInter) do ColsInter[i_ind]:=False ;
LesCols:=G_Qte.Titres[0] ; LesColQtes:=LesCols ; icol:=0 ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    BEGIN
    ichamp:=ChampToNum(NomCol) ;
    if ichamp>=0 then
       BEGIN
       if Pos('X',Mcd.getField(NomCol).Control)>0 then ColsInter[icol]:=True ;
       if NomCol='GF_DEPOT'        then SQ_Depot:=icol else
       if NomCol='GF_LIBELLE'      then SQ_Lib:=icol else
       if NomCol='GF_BORNEINF'     then SQ_QInf:=icol else
       if NomCol='GF_BORNESUP'     then SQ_QSup:=icol else
       if NomCol='GF_PRIXUNITAIRE' then SQ_Px:=icol else
       if NomCol='GF_CALCULREMISE' then SQ_Rem:=icol else
       if NomCol='GF_DATEDEBUT'    then SQ_Datedeb:=icol else
       if NomCol='GF_DATEFIN'      then SQ_Datefin:=icol ;
       END ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;

SetLength(ColsInter,G_TA.ColCount) ;
for i_ind:=Low(ColsInter)to High(ColsInter) do ColsInter[i_ind]:=False ;
LesCols:=G_TA.Titres[0] ; LesColTA:=LesCols ; icol:=0 ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    BEGIN
    ichamp:=ChampToNum(NomCol) ;
    if ichamp>=0 then
       BEGIN
       if Pos('X',Mcd.getField(NomCol).Control)>0 then ColsInter[icol]:=True ;
       if NomCol='GF_DEPOT'        then ST_Depot:=icol else
       if NomCol='GF_LIBELLE'      then ST_Lib:=icol else
       if NomCol='GF_PRIXUNITAIRE' then ST_Px:=icol else
       if NomCol='GF_CALCULREMISE' then ST_Rem:=icol else
       if NomCol='GF_DATEDEBUT'    then ST_Datedeb:=icol else
       if NomCol='GF_DATEFIN'      then ST_Datefin:=icol ;
       END ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;

  SetLength(ColsInter,G_Fam.ColCount) ;
  for i_ind:=Low(ColsInter)to High(ColsInter) do ColsInter[i_ind]:=False ;
  LesCols:=G_Fam.Titres[0] ;
  LesColFam:=LesCols ;
  icol:=0 ;

  Repeat
  NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
  if NomCol<>'' then
    BEGIN
    ichamp:=ChampToNum(NomCol) ;
    if ichamp>=0 then
       BEGIN
       if Pos('X',Mcd.getField(NomCol).Control)>0 then ColsInter[icol]:=True ;
       if NomCol='GF_DEPOT'         then SF_Depot:=icol else
       if NomCol='GF_TARIFARTICLE'  then SF_Fam:=icol else
       if NomCol='GF_SOUSFAMTARART' then SF_SFam:=icol else
       //if NomCol='CC_LIBELLE'       then SF_LibSFam:=icol else
       if Nomcol='BSF_LIBELLE'      then SF_LibSFam:=icol else
       if NomCol='GF_LIBELLE'       then SF_Lib:=icol else
       if NomCol='GF_PRIXUNITAIRE'  then SF_Px:=icol else
       if NomCol='GF_CALCULREMISE'  then SF_Rem:=icol else
       if NomCol='GF_DATEDEBUT'     then SF_Datedeb:=icol else
       if NomCol='GF_DATEFIN'       then SF_Datefin:=icol ;
       END ;
    END ;
  Inc(icol) ;
  Until ((LesCols='') or (NomCol='')) ;

  G_Fam.FColEditables[SF_LibSFam] := False;

END ;

procedure TFTarifCliArt.FormateZoneSaisie (ACol,ARow : Longint; ttd : T_TableTarif) ;
Var St,StC : String ;
BEGIN
Case ttd of
    ttdCliQte : BEGIN
                St:=G_Qte.Cells[ACol,ARow] ; StC:=St ;
                if ACol=SQ_Depot then StC:=uppercase(Trim(St)) else
                {$IFDEF MONCODE}
                    if ACol=SQ_Px then StC:=StrF00(Valeur(St),DEV.Decimale+NbreDecimalSuppPrix(NatureTiers,TOBArt)) else
                {$ELSE}
                    if ACol=SQ_Px then StC:=StrF00(Valeur(St),V_PGI.OkDecP) else
                {$ENDIF}
                         if ((ACol=SQ_QInf) or (ACol=SQ_QSup)) then StC:=StrF00(Valeur(St),0);
                G_Qte.Cells[ACol,ARow]:=StC ;
                END;
    ttdCliArt : BEGIN
                St:=G_TA.Cells[ACol,ARow] ; StC:=St ;
                if ACol=ST_Depot then StC:=uppercase(Trim(St)) else
                {$IFDEF MONCODE}
                    if ACol=ST_Px then StC:=StrF00(Valeur(St),DEV.Decimale+NbreDecimalSuppPrix(NatureTiers,TOBArt)) ;
                {$ELSE}
                    if ACol=ST_Px then StC:=StrF00(Valeur(St),V_PGI.OkDecP);
                {$ENDIF}
                G_TA.Cells[ACol,ARow]:=StC ;
                END;
    ttdCliFam : BEGIN
                St:=G_Fam.Cells[ACol,ARow] ;
                StC:=St ;
                if (ACol=SF_Depot) or (ACol=SF_Fam) or (Acol = SF_SFAM) then StC:=uppercase(Trim(St)) else
                    if ACol=SF_Px then StC:=StrF00(Valeur(St),V_PGI.OkDecP);
                G_Fam.Cells[ACol,ARow]:=StC ;
                END;
    END;
END ;

procedure TFTarifCliArt.InsertLigne (ARow : Longint; ttd : T_TableTarif) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if LigneVide (ARow, ttd) then exit;
Case ttd of
    ttdCliQte : BEGIN
                if (ARow > TOBTarfQte.Detail.Count) then Exit;
                G_Qte.CacheEdit; G_Qte.SynEnabled := False;
                TOB.Create ('TARIF', TOBTarfQte, ARow-1) ;
                G_Qte.InsertRow (ARow); G_Qte.Row := ARow;
                InitialiseLigne (ARow, ttd) ;
                PreAffecteLigne (ARow, ttd);
                G_Qte.MontreEdit; G_Qte.SynEnabled := True;
                AffectMenuCondApplic (G_Qte, ttd);
                AfficheCondTarf (G_Qte.Row, ttd);
                END;
    ttdCliArt : BEGIN
                if (ARow > TOBTarfTA.Detail.Count) then Exit;
                G_TA.CacheEdit; G_TA.SynEnabled := False;
                TOB.Create ('TARIF', TOBTarfTA, ARow-1) ;
                G_Qte.InsertRow (ARow); G_TA.Row := ARow;
                InitialiseLigne (ARow, ttd) ;
                PreAffecteLigne (ARow, ttd);
                G_TA.MontreEdit; G_TA.SynEnabled := True;
                AffectMenuCondApplic (G_TA, ttd);
                AfficheCondTarf (G_TA.Row, ttd);
                END;
    ttdCliFam : BEGIN
                if (ARow > TOBTarfFam.Detail.Count) then Exit;
                G_Fam.CacheEdit; G_Fam.SynEnabled := False;
                TOB.Create ('TARIF', TOBTarfFam, ARow-1) ;
                G_Fam.InsertRow (ARow); G_Fam.Row := ARow;
                InitialiseLigne (ARow, ttd) ;
                PreAffecteLigne (ARow, ttd);
                G_Fam.MontreEdit; G_Fam.SynEnabled := True;
                AffectMenuCondApplic (G_Fam, ttd);
                AfficheCondTarf (G_Fam.Row, ttd);
                END;
    END;
END;

procedure TFTarifCliArt.SupprimeLigne (ARow : Longint; ttd : T_TableTarif) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
Case ttd of
    ttdCliQte : BEGIN
                if (ARow > TOBTarfQte.Detail.Count) then Exit;
                G_Qte.CacheEdit; G_Qte.SynEnabled := False;
                G_Qte.DeleteRow (ARow);
                if (ARow = TOBTarfQte.Detail.Count) then
                    CreerTOBLigne (ARow + 1, ttd);
                SupprimeTOBTarif (ARow, ttd);
                if G_Qte.RowCount < NbRowsInit then G_Qte.RowCount := NbRowsInit;
                G_Qte.MontreEdit; G_Qte.SynEnabled := True;
                AffectMenuCondApplic (G_Qte, ttd);
                AfficheCondTarf (G_Qte.Row, ttd);
                END;
    ttdCliArt : BEGIN
                if (ARow > TOBTarfTA.Detail.Count) then Exit;
                G_TA.CacheEdit; G_TA.SynEnabled := False;
                G_TA.DeleteRow (ARow);
                if (ARow = TOBTarfTA.Detail.Count) then
                    CreerTOBLigne (ARow + 1, ttd);
                SupprimeTOBTarif (ARow, ttd);
                if G_TA.RowCount < NbRowsInit then G_TA.RowCount := NbRowsInit;
                G_TA.MontreEdit; G_TA.SynEnabled := True;
                AffectMenuCondApplic (G_TA, ttd);
                AfficheCondTarf (G_TA.Row, ttd);
                END;
    ttdCliFam : BEGIN
                if (ARow > TOBTarfFam.Detail.Count) then Exit;
                G_Fam.CacheEdit; G_Fam.SynEnabled := False;
                G_Fam.DeleteRow (ARow);
                if (ARow = TOBTarfFam.Detail.Count) then
                    CreerTOBLigne (ARow + 1, ttd);
                SupprimeTOBTarif (ARow, ttd);
                if G_Fam.RowCount < NbRowsInit then G_Fam.RowCount := NbRowsInit;
                G_Fam.MontreEdit; G_Fam.SynEnabled := True;
                AffectMenuCondApplic (G_Fam, ttd);
                AfficheCondTarf (G_Fam.Row, ttd);
                END;
    END;
END;

procedure TFTarifCliArt.SupprimeTOBTarif (ARow : Longint; ttd : T_TableTarif) ;
Var i_ind: integer;
BEGIN
Case ttd of
    ttdCliQte : BEGIN
                if TOBTarfQte.Detail[ARow-1].GetValue ('GF_TARIF') <> 0 then
                    BEGIN
                    i_ind := TOBTarifDelArt.Detail.Count;
                    TOB.Create ('TARIF', TOBTarifDelArt, i_ind) ;
                    TOBTarifDelArt.Detail[i_ind].Dupliquer (TOBTarfQte.Detail[ARow-1], False, True);
                    END;
                TOBTarfQte.Detail[ARow-1].Free;
                END;
    ttdCliArt : BEGIN
                if TOBTarfTA.Detail[ARow-1].GetValue ('GF_TARIF') <> 0 then
                    BEGIN
                    i_ind := TOBTarifDelArt.Detail.Count;
                    TOB.Create ('TARIF', TOBTarifDelArt, i_ind) ;
                    TOBTarifDelArt.Detail[i_ind].Dupliquer (TOBTarfTA.Detail[ARow-1], False, True);
                    END;
                TOBTarfTA.Detail[ARow-1].Free;
                END;
    ttdCliFam : BEGIN
                if TOBTarfFam.Detail[ARow-1].GetValue ('GF_TARIF') <> 0 then
                    BEGIN
                    i_ind := TOBTarifDelFam.Detail.Count;
                    TOB.Create ('TARIF', TOBTarifDelFam, i_ind) ;
                    TOBTarifDelFam.Detail[i_ind].Dupliquer (TOBTarfFam.Detail[ARow-1], False, True);
                    END;
                TOBTarfFam.Detail[ARow-1].Free;
                END;
    END;
END;

Function TFTarifCliArt.GrilleModifie : Boolean;
BEGIN
Result:=False ;
if Action=taConsult then Exit ;
Result:=(TOBTarfQte.IsOneModifie) or (TOBTarfTA.IsOneModifie) or
        (TOBTarifDelArt.IsOneModifie) or (TOBTarfFam.IsOneModifie) or
        (TOBTarifDelFam.IsOneModifie);
END;

Function TFTarifCliArt.SortDeLaLigne (ttd : T_TableTarif) : boolean ;
Var ACol,ARow : integer ;
    Cancel : boolean ;
BEGIN
Result:=False ;
Case ttd of
    ttdCliQte : BEGIN
                ACol:=G_Qte.Col ; ARow:=G_Qte.Row ; Cancel:=False ;
                G_QteCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
                G_QteRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
                END;
    ttdCliArt : BEGIN
                ACol:=G_TA.Col ; ARow:=G_TA.Row ; Cancel:=False ;
                G_TACellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
                G_TARowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
                END;
    ttdCliFam : BEGIN
                ACol:=G_Fam.Col ; ARow:=G_Fam.Row ; Cancel:=False ;
                G_FamCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
                G_FamRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
                END;
    END;
Result:=True ;
END ;

{==============================================================================================}
{=============================== Evènements de la Grid ========================================}
{==============================================================================================}
procedure TFTarifCliArt.G_QteEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
begin
Cancel := False; Chg := False;
G_QteRowEnter (Sender, G_Qte.Row, Cancel, Chg);
ACol := G_Qte.Col ; ARow := G_Qte.Row ;
G_QteCellEnter (Sender, ACol, ARow, Cancel);
end;

procedure TFTarifCliArt.G_QteRowEnter(Sender: TObject; Ou: Integer;
                                      var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
begin
if Ou >= G_Qte.RowCount - 1 then G_Qte.RowCount := G_Qte.RowCount + 1 ;;
ARow := Min (Ou, TOBTarfQte.detail.count + 1);
if (ARow = TOBTarfQte.detail.count + 1) AND (not LigneVide (ARow - 1, ttdCliQte)) then
    BEGIN
    CreerTOBligne (ARow, ttdCliQte);
    END;
if (LigneVide (ARow, ttdCliQte)) AND (not LigneVide (ARow - 1, ttdCliQte))then
    PreAffecteLigne (ARow, ttdCliQte);
if Ou > TOBTarfQte.detail.count then
    BEGIN
    G_Qte.Row := TOBTarfQte.detail.count;
    END;
AffichePied (ttdCliQte);
AfficheCondTarf (ARow, ttdCliQte);
end;

procedure TFTarifCliArt.G_QteRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if LigneVide (Ou, ttdCliQte) Then G_Qte.Row := Min (G_Qte.Row,Ou);
end;

procedure TFTarifCliArt.G_QteCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_Qte.Col <> SQ_Depot) AND (G_Qte.Col <> SQ_Lib) AND
       (G_Qte.Cells [SQ_Lib,G_Qte.Row] = '') then G_Qte.Col := SQ_Lib;
    G_Qte.ElipsisButton := ((G_Qte.Col = SQ_Depot) or (G_Qte.col = SQ_Datedeb) or
                            (G_Qte.col = SQ_Datefin)) ;
    StCellCur := G_Qte.Cells [G_Qte.Col,G_Qte.Row] ;
    AffectMenuCondApplic (G_Qte, ttdCliQte);
    END ;
end;

procedure TFTarifCliArt.G_QteCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
FormateZoneSaisie (ACol,ARow, ttdCliQte);
if ACol = SQ_Depot then TraiterDepot (ACol, ARow, ttdCliQte) else
    if ACol = SQ_Lib then TraiterLibelle (ACol, ARow, ttdCliQte) else
        if ACol = SQ_QInf then TraiterBorneInf (ACol, ARow, ttdCliQte) else
            if ACol = SQ_QSup then TraiterBorneSup (ACol, ARow, ttdCliQte) else
                if ACol = SQ_Px then TraiterPrix (ACol, ARow, ttdCliQte) else
                    if ACol = SQ_Rem then TraiterRemise (ACol, ARow, ttdCliQte) else
                        if ACol = SQ_Datedeb then TraiterDateDeb (ACol, ARow, ttdCliQte) else
                            if ACol = SQ_Datefin then TraiterDateFin (ACol, ARow, ttdCliQte);
if Not Cancel then
    BEGIN
    END;
end;

procedure TFTarifCliArt.G_QteElipsisClick(Sender: TObject);
Var DEPOT, DATE : THCritMaskEdit;
    Coord : TRect;
begin
if G_Qte.Col = SQ_Depot then
    BEGIN
    Coord := G_Qte.CellRect (G_Qte.Col, G_Qte.Row);
    DEPOT := THCritMaskEdit.Create (Self);
    DEPOT.Parent := G_Qte;
    DEPOT.Top := Coord.Top;
    DEPOT.Left := Coord.Left;
    DEPOT.Width := 3; DEPOT.Visible := False;
    DEPOT.DataType := 'GCDEPOT';
    GetDepotRecherche (DEPOT) ;
    if DEPOT.Text <> '' then G_Qte.Cells[G_Qte.Col,G_Qte.Row]:= DEPOT.Text;
    DEPOT.Destroy;
    END ;
if (G_Qte.Col = SQ_Datedeb) or (G_Qte.Col = SQ_Datefin) then
    BEGIN
    Coord := G_Qte.CellRect (G_Qte.Col, G_Qte.Row);
    DATE := THCritMaskEdit.Create (Self);
    DATE.Parent := G_Qte;
    DATE.Top := Coord.Top;
    DATE.left := Coord.Left;
    DATE.Width := 3; DATE.Visible := False;
    DATE.OpeType:=otDate;
    GetDateRecherche (TForm(DATE.Owner), DATE) ;
    if DATE.Text <> '' then G_Qte.Cells[G_Qte.Col,G_Qte.Row]:= DATE.Text;
    DATE.Destroy;
    END;
end;

procedure TFTarifCliArt.G_TAEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
begin
Cancel := False; Chg := False;
G_TARowEnter (Sender, G_TA.Row, Cancel, Chg);
ACol := G_TA.Col ; ARow := G_TA.Row ;
G_TACellEnter (Sender, ACol, ARow, Cancel);
end;

procedure TFTarifCliArt.G_TARowEnter(Sender: TObject; Ou: Integer;
                                     var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
begin
if Ou >= G_TA.RowCount - 1 then G_TA.RowCount := G_TA.RowCount +1  ;;
ARow := Min (Ou, TOBTarfTA.detail.count + 1);
if (ARow = TOBTarfTA.detail.count + 1) AND (not LigneVide (ARow - 1, ttdCliArt)) then
    BEGIN
    CreerTOBligne (ARow, ttdCliArt);
    END;
if (LigneVide (ARow, ttdCliArt)) AND (not LigneVide (ARow - 1, ttdCliArt))then
    PreAffecteLigne (ARow, ttdCliArt);
if Ou > TOBTarfTA.detail.count then
    BEGIN
    G_TA.Row := TOBTarfTA.detail.count;
    END;
AffichePied (ttdCliArt);
AfficheCondTarf (ARow, ttdCliArt);
end;

procedure TFTarifCliArt.G_TARowExit(Sender: TObject; Ou: Integer;
                                    var Cancel: Boolean; Chg: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if LigneVide (Ou, ttdCliArt) Then G_TA.Row := Min (G_TA.Row,Ou);
end;

procedure TFTarifCliArt.G_TACellEnter(Sender: TObject; var ACol,
                                      ARow: Integer; var Cancel: Boolean);
BEGIN
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_TA.Col <> ST_Depot) AND (G_TA.Col <> ST_Lib) AND
       (G_TA.Cells [ST_Lib,G_TA.Row] = '') then G_TA.Col := ST_Lib;
    G_TA.ElipsisButton := ((G_TA.Col = ST_Depot) or (G_TA.col = ST_Datedeb) or
                            (G_TA.col = ST_Datefin)) ;
    StCellCur := G_TA.Cells [G_TA.Col,G_TA.Row] ;
    AffectMenuCondApplic (G_TA, ttdCliArt);
    END ;
end;

procedure TFTarifCliArt.G_TACellExit(Sender: TObject; var ACol,
                                     ARow: Integer; var Cancel: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
FormateZoneSaisie (ACol,ARow, ttdCliArt);
if ACol = ST_Depot then TraiterDepot (ACol, ARow, ttdCliArt) else
    if ACol = ST_Lib then TraiterLibelle (ACol, ARow, ttdCliArt) else
        if ACol = ST_Px then TraiterPrix (ACol, ARow, ttdCliArt) else
            if ACol = ST_Rem then TraiterRemise (ACol, ARow, ttdCliArt) else
                if ACol = ST_Datedeb then TraiterDateDeb (ACol, ARow, ttdCliArt) else
                    if ACol = ST_Datefin then TraiterDateFin (ACol, ARow, ttdCliArt);
if Not Cancel then
    BEGIN
    END;
end;

procedure TFTarifCliArt.G_TAElipsisClick(Sender: TObject);
Var DEPOT, DATE : THCritMaskEdit;
    Coord : TRect;
begin
if G_TA.Col = ST_Depot then
    BEGIN
    Coord := G_TA.CellRect (G_TA.Col, G_TA.Row);
    DEPOT := THCritMaskEdit.Create (Self);
    DEPOT.Parent := G_TA;
    DEPOT.Top := Coord.Top;
    DEPOT.Left := Coord.Left;
    DEPOT.Width := 3; DEPOT.Visible := False;
    DEPOT.DataType := 'GCDEPOT';
    GetDepotRecherche (DEPOT) ;
    if DEPOT.Text <> '' then G_TA.Cells[G_TA.Col,G_TA.Row]:= DEPOT.Text;
    DEPOT.Destroy;
    END ;
if (G_TA.Col = ST_Datedeb) or (G_TA.Col = ST_Datefin) then
    BEGIN
    Coord := G_TA.CellRect (G_TA.Col, G_TA.Row);
    DATE := THCritMaskEdit.Create (Self);
    DATE.Parent := G_TA;
    DATE.Top := Coord.Top;
    DATE.left := Coord.Left;
    DATE.Width := 3; DATE.Visible := False;
    DATE.OpeType:=otDate;
    GetDateRecherche (TForm(DATE.Owner), DATE) ;
    if DATE.Text <> '' then G_TA.Cells[G_TA.Col,G_TA.Row]:= DATE.Text;
    DATE.Destroy;
    END;
end;

procedure TFTarifCliArt.G_FamEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
begin
Cancel := False; Chg := False;
G_FamRowEnter (Sender, G_Fam.Row, Cancel, Chg);
ACol := G_Fam.Col ; ARow := G_Fam.Row ;
G_FamCellEnter (Sender, ACol, ARow, Cancel);
end;

procedure TFTarifCliArt.G_FamRowEnter(Sender: TObject; Ou: Integer;
                                      var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
begin
if Ou >= G_Fam.RowCount - 1 then G_Fam.RowCount := G_Fam.RowCount + 1 ;;
ARow := Min (Ou, TOBTarfFam.detail.count + 1);
if (ARow = TOBTarfFam.detail.count + 1) AND (not LigneVide (ARow - 1, ttdCliFam)) then
    BEGIN
    CreerTOBligne (ARow, ttdCliFam);
    END;
if (LigneVide (ARow, ttdCliFam)) AND (not LigneVide (ARow - 1, ttdCliFam))then
    PreAffecteLigne (ARow, ttdCliFam);
if Ou > TOBTarfFam.detail.count then
    BEGIN
    G_Fam.Row := TOBTarfFam.detail.count;
    END;
AffichePied (ttdCliFam);
AfficheCondTarf (ARow, ttdCliFam);
end;

procedure TFTarifCliArt.G_FamRowExit(Sender: TObject; Ou: Integer;
                                     var Cancel: Boolean; Chg: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if LigneVide (Ou, ttdCliFam) Then G_Fam.Row := Min (G_Fam.Row,Ou);
end;

procedure TFTarifCliArt.G_FamCellEnter(Sender: TObject; var ACol,
                                       ARow: Integer; var Cancel: Boolean);
BEGIN
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_Fam.Col <> SF_Depot) AND (G_Fam.Col <> SF_Lib) AND (G_Fam.Col <> SF_Fam) AND (G_Fam.Col <> SF_SFam) then
        BEGIN
        if (G_Fam.Cells [SF_Lib,G_Fam.Row] = '') then G_Fam.Col := SF_Lib;
        if (G_Fam.Cells [SF_Fam,G_Fam.Row] = '') then G_Fam.Col := SF_Fam;
        END;
    G_Fam.ElipsisButton := ((G_Fam.Col = SF_Depot) or (G_Fam.col = SF_Fam) or (G_Fam.col = SF_SFam) or
                            (G_Fam.col = SF_Datedeb) or (G_Fam.col = SF_Datefin)) ;
    StCellCur := G_Fam.Cells [G_Fam.Col,G_Fam.Row] ;
    AffectMenuCondApplic (G_Fam, ttdCliFam);
    END ;
end;

procedure TFTarifCliArt.G_FamCellExit(Sender: TObject; var ACol,
                                      ARow: Integer; var Cancel: Boolean);
begin

  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application

  FormateZoneSaisie (ACol,ARow, ttdCliFam);

  if ACol = SF_Depot    then TraiterDepot (ACol, ARow, ttdCliFam)              else
  if ACol = SF_Lib      then TraiterLibelle (ACol, ARow, ttdCliFam)            else
  if ACol = SF_Fam      then TraiterFamille (ACol, ARow, ttdCliFam,Cancel)     else
  if ACol = SF_SFam     then TraiterSousFamille (ACol, ARow, ttdCliFam,Cancel) else
  if ACol = SF_Px       then TraiterPrix (ACol, ARow, ttdCliFam)               else
  if ACol = SF_Rem      then TraiterRemise (ACol, ARow, ttdCliFam)             else
  if ACol = SF_Datedeb  then TraiterDateDeb (ACol, ARow, ttdCliFam)            else
  if ACol = SF_Datefin  then TraiterDateFin (ACol, ARow, ttdCliFam);

  if Not Cancel then
  BEGIN
  END;

end;

procedure TFTarifCliArt.G_FamElipsisClick(Sender: TObject);
Var DEPOT, DATE, FAM: THCritMaskEdit;
    Coord : TRect;
begin
if G_Fam.Col = SF_Depot then
    BEGIN
    Coord := G_Fam.CellRect (G_Fam.Col, G_Fam.Row);
    DEPOT := THCritMaskEdit.Create (Self);
    DEPOT.Parent := G_Fam;
    DEPOT.Top := Coord.Top;
    DEPOT.Left := Coord.Left;
    DEPOT.Width := 3; DEPOT.Visible := False;
    DEPOT.DataType := 'GCDEPOT';
    GetDepotRecherche (DEPOT) ;
    if DEPOT.Text <> '' then G_Fam.Cells[G_Fam.Col,G_Fam.Row]:= DEPOT.Text;
    DEPOT.Destroy;
    END ;
if G_Fam.Col = SF_Fam then
    BEGIN
    Coord := G_Fam.CellRect (G_Fam.Col, G_Fam.Row);
    FAM := THCritMaskEdit.Create (Self);
    FAM.Parent := G_Fam;
    FAM.Top := Coord.Top;
    FAM.Left := Coord.Left;
    FAM.Width := 3; FAM.Visible := False;
    FAM.DataType := 'GCTARIFARTICLE';
    GetFamilleRecherche (FAM) ;
    if FAM.Text <> '' then G_Fam.Cells[G_Fam.Col,G_Fam.Row]:= FAM.Text;
    FAM.Destroy;
    END ;
if G_Fam.Col = SF_SFam then
    BEGIN
    Coord := G_Fam.CellRect (G_Fam.Col, G_Fam.Row);
    FAM := THCritMaskEdit.Create (Self);
    FAM.Parent := G_Fam;
    FAM.Top := Coord.Top;
    FAM.Left := Coord.Left;
    FAM.Width := 3; FAM.Visible := False;
    FAM.DataType := 'BTSOUSFAMTARART';
    FAM.Plus := ' AND BSF_FAMILLETARIF="'+G_FAM.cells[SF_FAM,G_Fam.row]+'"';
    GetFamilleRecherche (FAM) ;
    if FAM.Text <> '' then G_Fam.Cells[G_Fam.Col,G_Fam.Row]:= FAM.Text;
    FAM.Destroy;
    END ;
if (G_Fam.Col = SF_Datedeb) or (G_Fam.Col = SF_Datefin) then
    BEGIN
    Coord := G_Fam.CellRect (G_Fam.Col, G_Fam.Row);
    DATE := THCritMaskEdit.Create (Self);
    DATE.Parent := G_Fam;
    DATE.Top := Coord.Top;
    DATE.left := Coord.Left;
    DATE.Width := 3; DATE.Visible := False;
    DATE.OpeType:=otDate;
    GetDateRecherche (TForm(DATE.Owner), DATE) ;
    if DATE.Text <> '' then G_Fam.Cells[G_Fam.Col,G_Fam.Row]:= DATE.Text;
    DATE.Destroy;
    END;
end;

{==============================================================================================}
{========================= Manipulation des LIGNES Quantitatif ================================}
{==============================================================================================}
Procedure TFTarifCliArt.InitLaLigne (ARow : integer;  ttd : T_TableTarif);
Var TOBL : TOB;
BEGIN
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL=Nil then Exit ;
TOBL.PutValue ('GF_NATUREAUXI', NatureTiers);
TOBL.PutValue ('GF_TIERS', CodeTiers);
TOBL.PutValue ('GF_BORNESUP', 999999);
TOBL.PutValue ('GF_PRIXUNITAIRE', 0);
TOBL.PutValue ('GF_REMISE', 0);
TOBL.PutValue ('GF_CALCULREMISE', '');
TOBL.PutValue ('GF_DATEDEBUT', V_PGI.DateEntree);
TOBL.PutValue ('GF_DATEFIN', IDate2099);
TOBL.PutValue ('GF_MODECREATION', 'MAN');
if GF_CASCADEREMISE.Value <> '' then
    TOBL.PutValue ('GF_CASCADEREMISE', GF_CASCADEREMISE.Value)
    else TOBL.PutValue ('GF_CASCADEREMISE', 'MIE');
TOBL.PutValue ('GF_DEVISE', GF_DEVISE.Value);
TOBL.PutValue ('GF_QUALIFPRIX', 'GRP');
TOBL.PutValue ('GF_FERME', '-');
TOBL.PutValue ('GF_SOCIETE', V_PGI.CodeSociete) ;
Case ttd of
    ttdCliQte : BEGIN
                TOBL.PutValue ('GF_BORNEINF', 1);
                TOBL.PutValue ('GF_QUANTITATIF', 'X');
                TOBL.PutValue ('GF_ARTICLE', CodeArticle);
                if TarifTTC then TOBL.PutValue ('GF_REGIMEPRIX', 'TTC')
                            else TOBL.PutValue ('GF_REGIMEPRIX', 'HT') ;
                END;
    ttdCliArt : BEGIN
                TOBL.PutValue ('GF_BORNEINF', -999999);
                TOBL.PutValue ('GF_QUANTITATIF', '-');
                TOBL.PutValue ('GF_ARTICLE', CodeArticle);
                if TarifTTC then TOBL.PutValue ('GF_REGIMEPRIX', 'TTC')
                            else TOBL.PutValue ('GF_REGIMEPRIX', 'HT') ;
                END;
    ttdCliFam : BEGIN
                TOBL.PutValue ('GF_BORNEINF', -999999);
                TOBL.PutValue ('GF_QUANTITATIF', '-');
                TOBL.PutValue ('GF_ARTICLE', '');
                TOBL.PutValue ('GF_REGIMEPRIX', 'GLO') ;
                END;
    END;
AfficheLAligne (ARow, ttd);
END;

Function TFTarifCliArt.GetTOBLigne ( ARow : integer;  ttd : T_TableTarif) : TOB ;
BEGIN
Result:=Nil ;
case ttd of
    ttdCliQte : BEGIN
                if ((ARow<=0) or (ARow>TOBTarfQte.Detail.Count)) then Exit ;
                Result:=TOBTarfQte.Detail[ARow-1] ;
                END;
    ttdCliArt : BEGIN
                if ((ARow<=0) or (ARow>TOBTarfTA.Detail.Count)) then Exit ;
                Result:=TOBTarfTA.Detail[ARow-1] ;
                END;
    ttdCliFam : BEGIN
                if ((ARow<=0) or (ARow>TOBTarfFam.Detail.Count)) then Exit ;
                Result:=TOBTarfFam.Detail[ARow-1] ;
                END;
    END;
END ;

procedure TFTarifCliArt.InitialiseGrille (tav : T_ActionValid);
BEGIN
if (tav = tavArt) or (tav = tavTout) then
    BEGIN
    G_Qte.VidePile(false) ;
    G_Qte.RowCount:= NbRowsInit ;
    G_TA.VidePile(false) ;
    G_TA.RowCount:= NbRowsInit ;
    END;
if (tav = tavFam) or (tav = tavTout) then
    BEGIN
    G_Fam.VidePile(false) ;
    G_Fam.RowCount:= NbRowsInit ;
    END;
END;

procedure TFTarifCliArt.InitialiseLigne (ARow : integer; ttd : T_TableTarif) ;
Var TOBL : TOB ;
    i_ind : integer ;
BEGIN
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL<>Nil then TOBL.InitValeurs ;
case ttd of
    ttdCliQte : BEGIN
                for i_ind := 1 to G_Qte.ColCount-1 do
                    BEGIN
                    G_Qte.Cells [i_ind, ARow]:='' ;
                    END;
                END;
    ttdCliArt : BEGIN
                for i_ind := 1 to G_TA.ColCount-1 do
                    BEGIN
                    G_TA.Cells [i_ind, ARow]:='' ;
                    END;
                END;
    ttdCliFam : BEGIN
                for i_ind := 1 to G_Fam.ColCount-1 do
                    BEGIN
                    G_Fam.Cells [i_ind, ARow]:='' ;
                    END;
                END;
    END;
END;

{$IFDEF MONCODE}
procedure TFTarifCliArt.AffichePrix (TOBL : TOB ; ARow : integer; ttd : T_TableTarif) ;
var PuFix : double;
begin
if (TOBArt.GetValue('GA_DECIMALPRIX')='X') then
  begin
  PuFix:=TOBL.GetValue('GF_PRIXUNITAIRE');
  if PrixPourQte<=0 then PrixPourQte:=1;
  PuFix:=PuFix/PrixPourQte;
Case ttd of
  ttdCliQte : G_QTE.cells[SQ_Px,Arow] := strf00 (PuFix,DEV.Decimale+NbreDecimalSuppPrix(NatureTiers,TOBArt));
  ttdCliArt : G_TA.cells[ST_Px,Arow] := strf00 (PuFix,DEV.Decimale+NbreDecimalSuppPrix(NatureTiers,TOBArt));
  END;
  end;
end;
{$ENDIF}

procedure TFTarifCliArt.AfficheLaLigne (ARow : integer;  ttd : T_TableTarif) ;
Var TOBL : TOB ;
    i_ind,i_ind2 : integer ;
    nbDec : string;
BEGIN
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL = Nil then exit;
nbDec := '0.';
Case ttd of
    ttdCliQte : BEGIN
                for i_ind := 1 to G_Qte.ColCount-1 do
                    BEGIN
                    if ((i_ind=SQ_Datedeb) or (i_ind=SQ_Datefin)) then
                        G_QTE.ColFormats [i_ind] := '';
                    if i_ind=SQ_Px then
                       begin
                      {$IFDEF MONCODE}
                       for i_ind2 := 1 to DEV.Decimale+NbreDecimalSuppPrix(NatureTiers,TOBArt) do nbDec:=nbDec + '0';
                      {$ELSE}
                       for i_ind2 := 1 to V_PGI.OkDecP do nbDec:=nbDec + '0';
                      {$ENDIF}
                       G_QTE.ColFormats [i_ind] := '# ##' + nbDec;
                       end;
                    END;
                TOBL.PutLigneGrid(G_Qte,ARow,False,False,LesColQtes) ;
                {$IFDEF MONCODE}
                AffichePrix(TOBL,ARow,ttd) ;
                {$ENDIF}
                if G_Qte.Cells [SQ_Datedeb, ARow] = '' then
                    G_Qte.Cells [SQ_Datedeb, ARow] := DateToStr (IDate1900);
                for i_ind := 1 to G_Qte.ColCount-1 do
                    BEGIN
                    FormateZoneSaisie(i_ind, ARow, ttd) ;
                    END;
                END;
    ttdCliArt : BEGIN
                for i_ind := 1 to G_TA.ColCount-1 do
                    BEGIN
                    if ((i_ind=ST_Datedeb) or (i_ind=ST_Datefin)) then
                        G_TA.ColFormats [i_ind] := '';
                    if i_ind=ST_Px then
                       begin
                      {$IFDEF MONCODE}
                       for i_ind2 := 1 to DEV.Decimale+NbreDecimalSuppPrix(NatureTiers,TOBArt) do nbDec:=nbDec + '0';
                      {$ELSE}
                       for i_ind2 := 1 to V_PGI.OkDecP do nbDec:=nbDec + '0';
                      {$ENDIF}
                       G_TA.ColFormats [i_ind] := '# ##' + nbDec;
                       end;
                    END;
                TOBL.PutLigneGrid(G_TA,ARow,False,False,LesColTA) ;
                {$IFDEF MONCODE}
                AffichePrix(TOBL,ARow,ttd) ;
                {$ENDIF}
                if G_TA.Cells [ST_Datedeb, ARow] = '' then
                    G_TA.Cells [ST_Datedeb, ARow] := DateToStr (IDate1900);
                for i_ind := 1 to G_TA.ColCount-1 do
                    BEGIN
                    FormateZoneSaisie(i_ind, ARow, ttd) ;
                    END;
                END;
    ttdCliFam : BEGIN
                for i_ind := 1 to G_Fam.ColCount-1 do
                    BEGIN
                    if ((i_ind=SF_Datedeb) or (i_ind=SF_Datefin)) then
                        G_Fam.ColFormats [i_ind] := '';
                    END;
                TOBL.PutLigneGrid(G_Fam,ARow,False,False,LesColFam) ;
                if G_Fam.Cells [SF_Datedeb, ARow] = '' then
                    G_Fam.Cells [SF_Datedeb, ARow] := DateToStr (IDate1900);
                for i_ind := 1 to G_Fam.ColCount-1 do
                    BEGIN
                    FormateZoneSaisie(i_ind, ARow, ttd) ;
                    END;
                END;
    END;
END ;

Procedure TFTarifCliArt.DepileTOBLigne (tav : T_ActionValid);
var i_ind : integer;
BEGIN
if (tav = tavArt) or (tav = tavTout) then
    BEGIN
    for i_ind := TOBTarfQte.Detail.Count - 1 Downto 0 do
        BEGIN
        TOBTarfQte.Detail[i_ind].Free ;
        END;
    for i_ind := TOBTarfTA.Detail.Count - 1 Downto 0 do
        BEGIN
        TOBTarfTA.Detail[i_ind].Free ;
        END;
    for i_ind := TOBTarifDelArt.Detail.Count - 1 Downto 0 do
        BEGIN
        TOBTarifDelArt.Detail[i_ind].Free ;
        END;
    END;
if (tav = tavFam) or (tav = tavTout) then
    BEGIN
    for i_ind := TOBTarfFam.Detail.Count - 1 Downto 0 do
        BEGIN
        TOBTarfFam.Detail[i_ind].Free ;
        END;
    for i_ind := TOBTarifDelFam.Detail.Count - 1 Downto 0 do
        BEGIN
        TOBTarifDelFam.Detail[i_ind].Free ;
        END;
    END;
END;

Procedure TFTarifCliArt.CreerTOBLigne (ARow : integer; ttd : T_TableTarif);
BEGIN
Case ttd of
    ttdCliQte : BEGIN
                if ARow <> TOBTarfQte.Detail.Count + 1 then exit;
                TOB.Create ('TARIF', TOBTarfQte, ARow-1) ;
                InitialiseLigne (ARow, ttd) ;
                END;
    ttdCliArt : BEGIN
                if ARow <> TOBTarfTA.Detail.Count + 1 then exit;
                TOB.Create ('TARIF', TOBTarfTA, ARow-1) ;
                InitialiseLigne (ARow, ttd) ;
                END;
    ttdCliFam : BEGIN
                if ARow <> TOBTarfFam.Detail.Count + 1 then exit;
                TOB.Create ('TARIF', TOBTarfFam, ARow-1) ;
                InitialiseLigne (ARow, ttd) ;
                END;
    END;
END;

Function TFTarifCliArt.LigneVide (ARow : integer; ttd : T_TableTarif) : Boolean;
BEGIN
Result := True;
Case ttd of
    ttdCliQte : BEGIN
                if G_Qte.Cells [SQ_Lib, ARow] <> '' then
                    BEGIN
                    Result := False;
                    Exit;
                    END;
                END;
    ttdCliArt : BEGIN
                if G_TA.Cells [ST_Lib, ARow] <> '' then
                    BEGIN
                    Result := False;
                    Exit;
                    END;
                END;
    ttdCliFam : BEGIN
                if (G_Fam.Cells [SF_Lib, ARow] <> '') and (G_Fam.Cells [SF_Fam, ARow] <> '') then
                    BEGIN
                    Result := False;
                    Exit;
                    END;
                END;
    END;
END;

Procedure TFTarifCliArt.PreAffecteLigne (ARow : integer; ttd : T_TableTarif);
var TOBL, TOBLPrec : TOB;
BEGIN
TOBLPrec := GetTOBLigne (ARow - 1, ttd); if TOBLPrec = nil then exit;
if TOBLPrec.GetValue ('GF_BORNESUP') < 999999 then
    BEGIN
    Case ttd of
        ttdCliQte : TOBTarfQte.Detail[ARow - 1].Dupliquer (TOBLPrec, False, True);
        ttdCliArt : TOBTarfTA.Detail[ARow - 1].Dupliquer (TOBLPrec, False, True);
        ttdCliFam : TOBTarfFam.Detail[ARow - 1].Dupliquer (TOBLPrec, False, True);
        END;
    TOBL := GetTOBLigne (ARow, ttd);
    TOBL.PutValue ('GF_BORNEINF', TOBLPrec.GetValue ('GF_BORNESUP') + 1);
    TOBL.PutValue ('GF_BORNESUP', 999999);
    TOBL.PutValue ('GF_TARIF', 0);
    AfficheLAligne (ARow, ttd);
    END;
END;


{==============================================================================================}
{===================== Manipulation des Champs LIGNES Quantitatif =============================}
{==============================================================================================}
procedure TFTarifCliArt.TraiterDepot (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdCliQte : St := G_Qte.Cells [ACol, ARow];
    ttdCliArt : St := G_TA.Cells [ACol, ARow];
    ttdCliFam : St := G_Fam.Cells [ACol, ARow];
    END ;
if ExisteDepot ('GCDEPOT', St) then
    BEGIN
    TOBL.PutValue ('GF_DEPOT', St);
    END else
    BEGIN
    // message dépôt inexistant
    MsgBox.Execute (4,Caption,'') ;
    Case ttd of
        ttdCliQte : G_Qte.Cells [ACol, ARow] := TOBL.GetValue ('GF_DEPOT');
        ttdCliArt : G_TA.Cells [ACol, ARow] := TOBL.GetValue ('GF_DEPOT');
        ttdCliFam : G_Fam.Cells [ACol, ARow] := TOBL.GetValue ('GF_DEPOT');
        END ;
    END;
END;

procedure TFTarifCliArt.TraiterFamille (ACol, ARow : integer; ttd : T_TableTarif; var Cancel: Boolean);
var TOBL : TOB;
    B_NewLine : Boolean;
    St : string;
    NewFamille : boolean;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
if TOBL.GetValue ('GF_TARIFARTICLE') = '' then B_NewLine :=True else B_NewLine := False;
Case ttd of
    ttdCliFam : St := G_Fam.Cells [ACol, ARow];
    END;
if St <> '' then
    BEGIN
    NewFamille := (TOBL.GetValue ('GF_TARIFARTICLE') <> st);
    if ExisteFamille ('GCTARIFARTICLE', St) then
        BEGIN
        TOBL.PutValue ('GF_TARIFARTICLE', St);
        //
        if NEWFamille then ReInitSousFamille (Arow,ttd);
        //
        END else
        BEGIN
        // message Famille tarif article inexistante
        MsgBox.Execute (5,Caption,'') ;
        Cancel:=True;
        Case ttd of
            ttdCliFam : G_Fam.Cells [ACol, ARow] := TOBL.GetValue ('GF_TARIFARTICLE');
          END;
        END;
    END else if Not B_Newline then
    BEGIN
    Case ttd of
         ttdCliFam : G_Fam.Cells [ACol, ARow] := TOBL.GetValue ('GF_TARIFARTICLE');
        END;
    END;
END;

procedure TFTarifCliArt.TraiterSousFamille (ACol, ARow : integer; ttd : T_TableTarif; var Cancel: Boolean);
var TOBL      : TOB;
    St        : string;
    SFamTarif : String;
    SLibTarif : String;
BEGIN

  TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;

  //if TOBL.GetValue ('GF_SOUSFAMTARART') = '' then B_NewLine :=True else B_NewLine := False;
  Case ttd of
  	ttdCliFam : St := G_Fam.Cells [ACol, ARow];
  END;

  if St <> '' then
  BEGIN
    SFamTarif := G_Fam.Cells [SF_Fam, ARow];
  	if ExisteSousFamille (SFamTarif,St) then
    BEGIN
      TOBL.PutValue ('GF_SOUSFAMTARART', St);
      //RechDom('BTSOUSFAMTARART',St,False);
      SLibTarif := RechDom('BTSOUSFAMTARART',St, False, ' AND BSF_FAMILLETARIF="' + SFamTarif +'"');
      G_Fam.Cells [SF_LibSFam, ARow] := SLibTarif;
    END
    else
    BEGIN
    // message Sous-Famille tarif article inexistante
      MsgBox.Execute (7,Caption,'') ;
      Cancel:=True;
      Case ttd of
        ttdCliFam : G_Fam.Cells [ACol, ARow] := TOBL.GetValue ('GF_SOUSFAMTARART');
      END;
    END;
(*
    if (TOBL.GetValue ('GF_LIBELLE') <> '') AND (TOBL.GetValue ('GF_TARIFARTICLE') <> '') AND
       (B_NewLine) then
        BEGIN
        InitLaLigne (ARow, ttd);
        END;
    END else if Not B_Newline then
        BEGIN
        Case ttd of
    	ttdCliFam : TOBL.PutValue ('GF_SOUSFAMTARART', St); // on autorise pour tout les sous familles
            END;
*)
  END;

END;

procedure TFTarifCliArt.TraiterLibelle (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    B_NewLine : Boolean;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
if TOBL.GetValue ('GF_LIBELLE') = '' then B_NewLine :=True else B_NewLine := False;
Case ttd of
    ttdCliQte : BEGIN
                if G_Qte.Cells [ACol, ARow] <> '' then
                    BEGIN
                    TOBL.PutValue ('GF_LIBELLE', G_Qte.Cells [ACol, ARow]);
                    if (TOBL.GetValue ('GF_LIBELLE') <> '') AND (B_NewLine) then
                        BEGIN
                        InitLaLigne (ARow, ttd);
                        END;
                    END else if Not B_Newline then
                        G_Qte.Cells [ACol, ARow] := TOBL.GetValue ('GF_LIBELLE');
                 END;
    ttdCliArt : BEGIN
                if G_TA.Cells [ACol, ARow] <> '' then
                    BEGIN
                    TOBL.PutValue ('GF_LIBELLE', G_TA.Cells [ACol, ARow]);
                    if (TOBL.GetValue ('GF_LIBELLE') <> '') AND (B_NewLine) then
                        BEGIN
                        InitLaLigne (ARow, ttd);
                        END;
                    END else if Not B_Newline then
                        G_TA.Cells [ACol, ARow] := TOBL.GetValue ('GF_LIBELLE');
                END;
    ttdCliFam : BEGIN
                if G_Fam.Cells [ACol, ARow] <> '' then
                    BEGIN
                    TOBL.PutValue ('GF_LIBELLE', G_Fam.Cells [ACol, ARow]);
                    if (TOBL.GetValue ('GF_LIBELLE') <> '') AND
                       (TOBL.GetValue ('GF_TARIFARTICLE') <> '') AND
                       (B_NewLine) then
                        BEGIN
                        InitLaLigne (ARow, ttd);
                        END;
                    END else if Not B_Newline then
                        G_Fam.Cells [ACol, ARow] := TOBL.GetValue ('GF_LIBELLE');
                END;
    END ;
END;

procedure TFTarifCliArt.TraiterBorneInf (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    f_QteInf : Extended;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
f_QteInf := 0;
Case ttd of
    ttdCliQte : f_QteInf := Valeur (G_Qte.Cells [ACol, ARow]);
    END ;
if f_QteInf > TOBL.GetValue ('GF_BORNESUP') then
    BEGIN
    TOBL.PutValue ('GF_BORNESUP', f_QteInf);
    Case ttd of
        ttdCliQte : G_Qte.Cells [SQ_QSup, ARow] := floattostr (TOBL.GetValue ('GF_BORNESUP'));
        END ;
    END;
if f_QteInf < 1 then
    BEGIN
    MsgBox.Execute (2,Caption,'') ;
    Case ttd of
        ttdCliQte : BEGIN
                    G_Qte.Cells [ACol, ARow] := floattostr (TOBL.GetValue ('GF_BORNEINF'));
                    G_Qte.Col := ACol; G_Qte.Row := ARow;
                    END;
        END ;
    END else TOBL.PutValue ('GF_BORNEINF', f_QteInf);
END;

procedure TFTarifCliArt.TraiterBorneSup (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    f_QteSup : Extended;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
f_QteSup := 999999;
Case ttd of
    ttdCliQte : f_QteSup := Valeur (G_Qte.Cells [ACol, ARow]);
    END ;
if f_QteSup < TOBL.GetValue ('GF_BORNEINF') then
    BEGIN
    MsgBox.Execute (2,Caption,'') ;
    Case ttd of
        ttdCliQte : BEGIN
                    G_Qte.Cells [ACol, ARow] := floattostr (TOBL.GetValue ('GF_BORNESUP'));
                    G_Qte.Col := ACol; G_Qte.Row := ARow;
                    END;
        END ;
    END else
    BEGIN
    TOBL.PutValue ('GF_BORNESUP', f_QteSup);
    END;
END;

procedure TFTarifCliArt.TraiterPrix (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    PuFix : double;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
PuFix := 0;
Case ttd of
    ttdCliQte : PuFix := Valeur (G_Qte.Cells [ACol, ARow]);
    ttdCliArt : PuFix := Valeur (G_TA.Cells [ACol, ARow]);
    ttdCliFam : PuFix := Valeur (G_Fam.Cells [ACol, ARow]);
    END ;
{$IFDEF MONCODE}
if TOBArt.GetValue('GA_DECIMALPRIX')='X' then
  begin
  if PrixPourQte<=0 then PrixPourQte:=1;
  PuFix:=PuFix*PrixPourQte;
  end;
{$ENDIF}
TOBL.PutValue ('GF_PRIXUNITAIRE',PuFix);
END;

procedure TFTarifCliArt.TraiterRemise(ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdCliQte : begin
                G_Qte.Cells [ACol, ARow]:=ModifFormat(G_Qte.Cells [ACol, ARow]);
                St := G_Qte.Cells [ACol, ARow];
                end;

    ttdCliArt : begin
                G_TA.Cells [ACol, ARow] := ModifFormat(G_TA.Cells [ACol, ARow]);
                St:= G_TA.Cells [ACol, ARow];
                end;

    ttdCliFam : begin
                G_Fam.Cells [ACol, ARow] := ModifFormat(G_Fam.Cells [ACol, ARow]);
                St := G_Fam.Cells [ACol, ARow];
                end;
    END ;
TOBL.PutValue ('GF_CALCULREMISE', St);
TOBL.PutValue ('GF_REMISE', RemiseResultante (St));
AffichePied (ttd);
END;


procedure TFTarifCliArt.TraiterDateDeb (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    St_Date : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdCliQte : St_Date := G_Qte.Cells [ACol, ARow] ;
    ttdCliArt : St_Date := G_TA.Cells [ACol, ARow] ;
    ttdCliFam : St_Date := G_Fam.Cells [ACol, ARow] ;
    END ;
if IsValidDate (st_Date) then
    BEGIN
    if StrToDate (St_Date) > TOBL.GetValue ('GF_DATEFIN') then
        BEGIN
            MsgBox.Execute (3,Caption,'') ;
            Case ttd of
                ttdCliQte : BEGIN
                            G_Qte.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
                            G_Qte.Col := ACol; G_Qte.Row := ARow;
                            END;
                ttdCliArt : BEGIN
                            G_TA.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
                            G_TA.Col := ACol; G_TA.Row := ARow;
                            END;
                ttdCliFam : BEGIN
                            G_Fam.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
                            G_Fam.Col := ACol; G_Fam.Row := ARow;
                            END;
                END ;
        END else
        BEGIN
        TOBL.PutValue ('GF_DATEDEBUT', StrToDate (St_Date));
        END;
    END else
    BEGIN
    if TOBL.GetValue ('GF_LIBELLE') <> '' then
        BEGIN
        Case ttd of
            ttdCliQte : G_Qte.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
            ttdCliArt : G_Fam.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
            ttdCliFam : G_Fam.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEDEBUT');
            END ;
        END;
    END;
END;

procedure TFTarifCliArt.TraiterDateFin (ACol, ARow : integer; ttd : T_TableTarif);
var TOBL : TOB;
    St_Date : string;
BEGIN
TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
Case ttd of
    ttdCliQte : St_Date := G_Qte.Cells [ACol, ARow] ;
    ttdCliArt : St_Date := G_TA.Cells [ACol, ARow] ;
    ttdCliFam : St_Date := G_Fam.Cells [ACol, ARow] ;
    END ;
if IsValidDate (st_Date) then
    BEGIN
    if StrToDate (St_Date) < TOBL.GetValue ('GF_DATEDEBUT') then
        BEGIN
            MsgBox.Execute (3,Caption,'') ;
            Case ttd of
                ttdCliQte : BEGIN
                            G_Qte.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
                            G_Qte.Col := ACol; G_Qte.Row := ARow;
                            END;
                ttdCliArt : BEGIN
                            G_TA.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
                            G_TA.Col := ACol; G_TA.Row := ARow;
                            END;
                ttdCliFam : BEGIN
                            G_Fam.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
                            G_Fam.Col := ACol; G_Fam.Row := ARow;
                            END;
               END ;
        END else
        BEGIN
        TOBL.PutValue ('GF_DATEFIN', StrToDate (St_Date));
        END;
    END else
    BEGIN
    if TOBL.GetValue ('GF_LIBELLE') <> '' then
        BEGIN
        Case ttd of
            ttdCliQte : G_Qte.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
            ttdCliArt : G_TA.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
            ttdCliFam : G_Fam.Cells [ACol, ARow] := TOBL.GetValue ('GF_DATEFIN');
            END ;
        END;
    END;
END;

{==============================================================================================}
{========================= Evenement de l'Article  ============================================}
{==============================================================================================}
procedure TFTarifCliArt.GF_CODEARTICLEExit(Sender: TObject);
BEGIN
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if GF_CODEARTICLE.Text <> TOBArt.GetValue ('GA_CODEARTICLE') then TraiterArticle ;
END;

procedure TFTarifCliArt.GF_CODEARTICLEElipsisClick(Sender: TObject);
begin
ChercherArticle;
end;

{==============================================================================================}
{========================= Manipulation de l'article ==========================================}
{==============================================================================================}
Procedure TFTarifCliArt.InitialiseEnteteArticle ;
BEGIN
TOBArt.InitValeurs ;
AffecteEnteteArticle;
CodeArticle := '';
G_Qte.Enabled := False ; G_TA.Enabled := False;
GA_PRIXPOURQTE.Value:=0; PrixPourQte:=1;
GA_PRIXPOURQTE.Visible:=False;
TGA_PRIXPOURQTE.Visible:=False;
END;

Procedure TFTarifCliArt.AffecteEnteteArticle ;
BEGIN
GF_CODEARTICLE.Text := TOBArt.GetValue ('GA_CODEARTICLE');
TGF_LIBELLEART.Caption := TOBArt.GetValue ('GA_LIBELLE');
// Dimensions
AffecteDimension;
END;

Procedure TFTarifCliArt.AffecteDimension ;
var TLIB, TLIB2 : THLabel;
    CHPS, CHPS2 : THCritMaskEdit;
    i_ind, i_Dim : integer;
    b_Dim : Boolean ;
    GrilleDim,CodeDim,LibDim : String ;
BEGIN
i_Dim := 1;
b_Dim := False;
for i_ind := 1 to MaxDimension do
    BEGIN
    TLIB := THLabel (FindComponent ('TGF_GRILLEDIM' + IntToStr (i_Dim)));
    CHPS := THCritMaskEdit (FindComponent ('GF_CODEDIM' + IntToStr (i_Dim)));
    TLIB2 := THLabel (FindComponent ('TGF_GRILLEDIM' + IntToStr (i_ind)));
    CHPS2 := THCritMaskEdit (FindComponent ('GF_CODEDIM' + IntToStr (i_ind)));
    TLIB2.Caption := '';
    CHPS2.Text := '';
    CHPS2.Visible := False;
    GrilleDim:=TOBArt.GetValue('GA_GRILLEDIM'+IntToStr(i_ind)) ;
    CodeDim:=TOBArt.GetValue('GA_CODEDIM'+IntToStr(i_ind)) ;
    if  GrilleDim<>'' then
        BEGIN
        TLIB.Caption:=RechDom('GCGRILLEDIM'+IntToStr(i_ind),GrilleDim,FALSE) ;
        LibDim:=GCGetCodeDim(GrilleDim,CodeDim,i_ind) ;
        if LibDim<>'' then
           BEGIN
           CHPS.Text:=LibDim ; CHPS.Visible:=True ; b_Dim:=True ; Inc(i_Dim) ;
           END else
           BEGIN
           TLIB.Caption:='' ; CHPS.Text:='' ; CHPS.Visible:=False ;
           END;
        END ;
    END;
if Not b_Dim then
    BEGIN
    PDIMENSION.Visible := False ;
    END else
    BEGIN
    PDIMENSION.Visible := True ;
    END;
END;

Procedure TFTarifCliArt.ChercherArticle ;
BEGIN
DispatchRechArt (GF_CODEARTICLE, 1, '',
                 'GA_CODEARTICLE=' + Trim (Copy (GF_CODEARTICLE.Text, 1, 18)), '');
if GF_CODEARTICLE.Text <> '' then
    BEGIN
    GF_CODEARTICLE.Text:=Format('%-33.33sX',[GF_CODEARTICLE.Text]);
    TOBArt.SelectDB ('"' + GF_CODEARTICLE.Text + '"', Nil) ;
    GF_CODEARTICLE.Text := TOBArt.GetValue ('GA_CODEARTICLE') ;
    TraiterArticle;
    END;
END;

Procedure TFTarifCliArt.TraiterArticle ;
Var RechArt : T_RechArt ;
    OkArt   : Boolean ;
    i_Rep: Integer;
    ioerr : TIOErr ;
    CodeArt : string;
BEGIN
OkArt:=False ;
CodeArt := CodeArticle ;
if GF_CODEARTICLE.Text <> '' then
    BEGIN
    RechArt := TrouverArticle (GF_CODEARTICLE, TOBArt);
    Case RechArt of
            traOk : OkArt:=True ;
         traAucun : BEGIN
                    // Recherche sur code via LookUp ou Recherche avancée
                    DispatchRechArt (GF_CODEARTICLE, 1, '',
                                     'GA_CODEARTICLE=' + Trim (Copy (GF_CODEARTICLE.Text, 1, 18)), '');
                    if GF_CODEARTICLE.Text <> '' then
                        BEGIN
                        GF_CODEARTICLE.Text:=Format('%-33.33sX',[GF_CODEARTICLE.Text]);
                        Okart := TOBArt.SelectDB ('"' + GF_CODEARTICLE.Text + '"', nil);
                        END;
                    END ;
        traGrille : BEGIN
                    // Forcement objet dimension avec saisie obligatoire
                    if ChoisirDimension (TOBArt.GetValue ('GA_ARTICLE'), TOBArt) then
                        BEGIN
                        Okart := True;
                        END;
                    END;
        End ; // Case
    END;
if (Okart) then
    BEGIN
    AffecteEnteteArticle;
    G_Qte.Enabled := True ; G_TA.Enabled := True ;
    if (NatureTiers='CLI') then PrixPourQte := TOBArt.GetValue('GA_PRIXPOURQTE')
                           else PrixPourQte := TOBArt.GetValue('GA_PRIXPOURQTEAC');
    GA_PRIXPOURQTE.Value:=PrixPourQte;
    if (GA_PRIXPOURQTE.Value>1) then
       begin
       {$IFDEF MONCODE}
       if (TOBArt.GetValue('GA_DECIMALPRIX')='X') then
         begin
         GA_PRIXPOURQTE.Value:=DEV.Decimale+NbreDecimalSuppPrix(NatureTiers,TOBArt);
         TGA_PRIXPOURQTE.Caption := HTitre.Mess[13] ;
         end;
       {$ENDIF}
       GA_PRIXPOURQTE.Visible:=True;
       TGA_PRIXPOURQTE.Visible:=True;
       end else
       begin
       GA_PRIXPOURQTE.Visible:=False;
       TGA_PRIXPOURQTE.Visible:=False;
       end;
    END else
    BEGIN
    InitialiseEnteteArticle ;
    TOBArt.InitValeurs;
    END;
if CodeArt <> TOBArt.GetValue ('GA_ARTICLE') then
    BEGIN
    i_Rep := QuestionTarifEnCours (tavArt);
    Case i_Rep of
        mrYes    : BEGIN
                   ioerr := Transactions (ValideTarifArticle, 2);
                   Case ioerr of
                          oeOk : ;
                       oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                        oeSaisie : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                        END ;
                   CodeArticle := TOBArt.GetValue ('GA_ARTICLE');
                   Transactions (ChargeTarifArt, 1);
                   END;
        mrNo     : BEGIN
                   TOBTarfQte.SetAllModifie (False);
                   TOBTarfTA.SetAllModifie (False);
                   TOBTarifDelArt.SetAllModifie (False);
                   InitialiseGrille (tavArt);
                   DepileTOBLigne (tavArt);
                   CodeArticle := TOBArt.GetValue ('GA_ARTICLE');
                   Transactions (ChargeTarifArt, 1);
                   END;
        mrCancel : BEGIN
                   CodeArticle:=CodeArt ;
                   if TOBArt.SelectDB ('"' + CodeArticle + '"', Nil) then
                       BEGIN
                       AffecteEnteteArticle ;
                       G_Qte.Enabled := True ; G_TA.Enabled := True ;
                       END ;
                   END;
        END ;
    END;
END;

{==============================================================================================}
{========================= Evenement de l'Entete  =============================================}
{==============================================================================================}
procedure TFTarifCliArt.GF_TIERSExit(Sender: TObject);
BEGIN

  if GF_TIERS.Text = '' then ImportAchat.Visible := False else ImportAchat.Visible := True;

  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application

  if GF_TIERS.Text <> TOBTiers.GetValue ('T_TIERS') then TraiterTiers;

END;


procedure TFTarifCliArt.GF_TIERSElipsisClick(Sender: TObject);
begin
ChercherTiers;
end;

procedure TFTarifCliArt.GF_DEVISEChange(Sender: TObject);
var i_Rep : integer;
    ioerr : TIOErr ;
begin
if CodeDevise = GF_DEVISE.Value then Exit;
i_Rep := QuestionTarifEnCours (tavTout);
Case i_Rep of
    mrYes    : BEGIN
               if GF_CODEARTICLE.Text = '' then
                    BEGIN
                    ioerr := Transactions (ValideTarifFamille, 2);
                    Case ioerr of
                       oeOk      : ;
                       oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                       oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                       END ;
                    CodeDevise := GF_DEVISE.Value;
                    DEV.Code := GF_DEVISE.Value ; GetInfosDevise (DEV) ;
                    Transactions (ChargeTarifFam, 1);
                    END else
                    BEGIN
                    ioerr := Transactions (ValideTarifTout, 2);
                    Case ioerr of
                       oeOk      : ;
                       oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                       oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                       END ;
                    CodeDevise := GF_DEVISE.Value;
                    Transactions (ChargeTarifTout, 1);
                    END;
               END;
    mrNo     : BEGIN
               TOBTarfQte.SetAllModifie (False);
               TOBTarfTA.SetAllModifie (False);
               TOBTarfFam.SetAllModifie (False);
               TOBTarifDelArt.SetAllModifie (False);
               TOBTarifDelFam.SetAllModifie (False);
               InitialiseGrille (tavTout);
               DepileTOBLigne (tavTout);
               CodeDevise := GF_DEVISE.Value;
               DEV.Code := GF_DEVISE.Value ; GetInfosDevise (DEV) ;
               if GF_CODEARTICLE.Text = '' then
                    Transactions (ChargeTarifFam, 1) else
                    Transactions (ChargeTarifTout, 1);
               END;
    mrCancel : GF_DEVISE.Value := CodeDevise;
    END ;
InitialisePied ;
end;

procedure TFTarifCliArt.CBQUANTITATIFClick(Sender: TObject);
begin
InitialisePied ;
if CBQUANTITATIF.Checked then
    BEGIN
    PQTEART.Visible := True ;
    PTART.Visible := False ;
    PTITRE.Caption := HTitre.Mess[1] ;
    AffectMenuCondApplic (G_Qte, ttdCliQte);
    END else
    BEGIN
    PQTEART.Visible := False ;
    PTART.Visible := True ;
    PTITRE.Caption := HTitre.Mess[2] ;
    AffectMenuCondApplic (G_TA, ttdCliArt);
    END;
HMTrad.ResizeGridColumns (G_Qte) ;
HMTrad.ResizeGridColumns (G_TA) ;
HMTrad.ResizeGridColumns (G_Fam) ;
if TarifTTC then TTYPETARIF.Caption := HTitre.Mess[6]
            else TTYPETARIF.Caption := HTitre.Mess[5] ;
end;

procedure TFTarifCliArt.CBARTICLEClick(Sender: TObject);
begin
InitialisePied ;
if CBARTICLE.Checked then
    BEGIN
    PFAMARTICLE.Visible := False;
    PTARIFART.Visible := True;
    if PrixPourQte>1 then
       begin
       GA_PRIXPOURQTE.Visible:=True;
       TGA_PRIXPOURQTE.Visible:=True;
       end;
    if CBQUANTITATIF.Checked then
        BEGIN
        PQTEART.Visible := True ;
        PTART.Visible := False ;
        PTITRE.Caption := HTitre.Mess[1] ;
        AffectMenuCondApplic (G_Qte, ttdCliQte);
        END else
        BEGIN
        PQTEART.Visible := False ;
        PTART.Visible := True ;
        PTITRE.Caption := HTitre.Mess[2] ;
        AffectMenuCondApplic (G_TA, ttdCliArt);
        END;
    if TarifTTC then TTYPETARIF.Caption := HTitre.Mess[6]
                else TTYPETARIF.Caption := HTitre.Mess[5] ;
    END else
    BEGIN
    PFAMARTICLE.Visible := True;
    PTARIFART.Visible := False;
    PQTEART.Visible := False ;
    PTART.Visible := False ;
    GA_PRIXPOURQTE.Visible:=False;
    TGA_PRIXPOURQTE.Visible:=False;
    PTITRE.Caption := HTitre.Mess[0] ;
    TTYPETARIF.Caption := HTitre.Mess[4] ;
    AffectMenuCondApplic (G_Fam, ttdCliFam);
    END;
HMTrad.ResizeGridColumns (G_Qte) ;
HMTrad.ResizeGridColumns (G_TA) ;
HMTrad.ResizeGridColumns (G_Fam) ;
AffichePrixCon ;
ENREGAUTO.Visible := CBARTICLE.Checked;
end;

{==============================================================================================}
{========================= Manipulation de l'entête ===========================================}
{==============================================================================================}
Procedure TFTarifCliArt.PrepareEntete ;
var QTiers : TQuery;
BEGIN
GF_DEVISE.Value := CodeDevise;
TOBArt.InitValeurs ;
TOBTiers.InitValeurs ;
QTiers:=OpenSQL('SELECT * FROM TIERS WHERE T_TIERS="'+CodeTiers+'"',true,-1,'',true);
if not QTiers.eof then TOBTiers.SelectDB('',QTiers);
Ferme(QTiers);
GF_TIERS.Text := TOBTiers.GetValue ('T_TIERS') ;
TiersIsClient := (TOBTiers.GetValue('T_NATUREAUXI') = 'CLI'); // JT eQualité 10807
if GF_TIERS.Text <> '' then
    BEGIN
    AffecteEntete;
    GF_CODEARTICLE.Enabled := True;
    G_Qte.Enabled := False ; G_TA.Enabled := False; G_Fam.Enabled := True ;
    InitialiseGrille (tavTout);
    DepileTOBLigne (tavTout);
    InitialiseEnteteArticle;
    Transactions (ChargeTarifFam, 1);
    END else
    BEGIN
    TOBTiers.InitValeurs ;
    AffecteEntete;
    G_Qte.Enabled := False ; G_TA.Enabled := False; G_Fam.Enabled := False ;
    END;
GF_TIERS.Enabled:=False ; GF_CASCADEREMISE.Enabled:=False; GF_CASCADEREMISE.Style:=csSimple; G_Fam.SetFocus ;
BCondAplli.Enabled := False ; BCopierCond.Enabled := False; BCollerCond.Enabled := False ;
CONDAPPLIC.Text := '' ;
ImportAchat.Visible := True;
InitialisePied;
END;

Procedure TFTarifCliArt.InitialiseEntete ;
BEGIN
  GF_DEVISE.Value := CodeDevise;
  TOBArt.InitValeurs ; PrixPourQte:=1;
  TOBTiers.InitValeurs ;
  AffecteEntete;
  CodeTiers := '';
  GF_TIERS.SetFocus ;
  InitialiseEnteteArticle;
  GF_CODEARTICLE.Enabled := False;
  G_Qte.Enabled := False ; G_TA.Enabled := False; G_Fam.Enabled := False ;
  BCondAplli.Enabled := False ; BCopierCond.Enabled := False; BCollerCond.Enabled := False ;
  CONDAPPLIC.Text := '' ;
  ImportAchat.Visible := False;
  InitialisePied ;
END;

Procedure TFTarifCliArt.AffecteEntete ;
BEGIN
GF_TIERS.Text := TOBTiers.GetValue ('T_TIERS');
TGF_LIBELLE.Caption := TOBTiers.GetValue ('T_LIBELLE');
END;


Procedure TFTarifCliArt.ChercherTiers ;
var QTiers : TQuery;
BEGIN
DispatchRecherche (GF_TIERS, 2, 'T_NATUREAUXI="' + NatureTiers + '"',
                   'T_TIERS=' + Trim (Copy (CodeTiers, 1, 17)), '');
  if GF_TIERS.Text <> '' then
    BEGIN
    QTiers:=OpenSQL('SELECT * FROM TIERS WHERE T_TIERS="'+GF_TIERS.Text+'"',true,-1,'',true);
    if not QTiers.eof then TOBTiers.SelectDB('',QTiers);
    Ferme(QTiers);
    GF_TIERS.Text := TOBTiers.GetValue ('T_TIERS') ;
    TraiterTiers ;
    END;

  if GF_TIERS.text = '' then ImportAchat.Visible := False else ImportAchat.Visible := true;

END;

Procedure TFTarifCliArt.TraiterTiers ;
Var RechTiers : T_RechArt ;
    OkTiers   : Boolean ;
    QTiers : TQuery;
    i_Rep: Integer;
    ioerr : TIOErr ;
BEGIN
if GF_TIERS.Text = '' then
    BEGIN
    InitialiseEntete ;
    InitialiseGrille (tavTout);
    DepileTOBLigne (tavTout);
    exit
    END;

OkTiers:=False ;
if NatureTiers = 'FOU' then RechTiers := TrouverFouSQL (GF_TIERS, TOBTiers)
                       else RechTiers := TrouverTiersSQL (GF_TIERS, TOBTiers);
Case RechTiers of
        traOk : OkTiers:=True ;
     traAucun : BEGIN
                // Recherche sur code via LookUp ou Recherche avancée
                    DispatchRecherche (GF_TIERS, 2, 'T_NATUREAUXI="' + NatureTiers + '"',
                                       'T_TIERS=' + Trim (Copy (CodeTiers, 1, 17)), '');
                    if GF_TIERS.Text <> '' then
                    BEGIN
                    QTiers:=OpenSQL('SELECT * FROM TIERS WHERE T_TIERS="'+GF_TIERS.Text+'"',true,-1,'',true);
                    if not QTiers.eof then TOBTiers.SelectDB('',QTiers);
                    Ferme(QTiers);
                    OkTiers := True;
                    END;
                END ;
     End ; // Case

if (OkTiers) then
    BEGIN
    AffecteEntete;
    GF_CODEARTICLE.Enabled := True;
    //G_Qte.Enabled := True ; G_TA.Enabled := True ;
    G_Fam.Enabled := True ;
    if CodeTiers <> TOBTiers.GetValue ('T_TIERS') then
        BEGIN
        i_Rep := QuestionTarifEnCours (tavTout);
        Case i_Rep of
            mrYes    : BEGIN
                       ioerr := Transactions (ValideTarifTout, 2);
                       Case ioerr of
                           oeOk      : ;
                           oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
                           oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
                           END ;
                       CodeTiers := TOBTiers.GetValue ('T_TIERS');
                       InitialiseEnteteArticle;
                       Transactions (ChargeTarifFam, 1);
                       END;
            mrNo     : BEGIN
                       TOBTarfQte.SetAllModifie (False);
                       TOBTarfTA.SetAllModifie (False);
                       TOBTarifDelArt.SetAllModifie (False);
                       TOBTarfFam.SetAllModifie (False);
                       TOBTarifDelFam.SetAllModifie (False);
                       InitialiseGrille (tavTout);
                       DepileTOBLigne (tavTout);
                       CodeTiers := TOBTiers.GetValue ('T_TIERS');
                       InitialiseEnteteArticle;
                       Transactions (ChargeTarifFam, 1);
                       END;
            mrCancel : BEGIN
                       QTiers:=OpenSQL('SELECT * FROM TIERS WHERE T_TIERS="'+CodeTiers+'"',true,-1,'',true);
                       if not QTiers.eof then TOBTiers.SelectDB('',QTiers);
                       Ferme(QTiers);
                       AffecteEntete;
                       END;
            END ;
        END;
    END else
    BEGIN
    InitialiseEntete ;
    END;
END;

Function TFTarifCliArt.QuestionTarifEnCours (tav: T_ActionValid): Integer;
var b_question : boolean;
BEGIN
Result := mrNo;
b_question := False;
if Action = taConsult then Exit;
if (tav = tavArt) or (tav = tavTout) then
    BEGIN
    if (TOBTarfQte.IsOneModifie) or (TOBTarfTA.IsOneModifie) or
       (TOBTarifDelArt.IsOneModifie) then
        BEGIN
        if ENREGAUTO.Checked then Result := mrYes else b_question := True;
        END;
    END;
if (tav = tavFam) or (tav = tavTout) then
    BEGIN
    if (TOBTarfFam.IsOneModifie) or (TOBTarifDelFam.IsOneModifie) then
        BEGIN
        b_question := True;
        END;
    END;
if b_question then Result := MsgBox.Execute (0, Caption, '');
END;

{==============================================================================================}
{=============================== Evenements du pied ===========================================}
{==============================================================================================}
procedure TFTarifCliArt.GF_REMISEChange(Sender: TObject);
var St : string;
begin
St := GF_REMISE.Text;
St:=StrF00(Valeur(St),ADecimP);
GF_REMISE.Text := St;
end;

procedure TFTarifCliArt.GF_CASCADEREMISEChange(Sender: TObject);
Var Row : integer;
    ttd : T_TableTarif;
    TOBL : TOB ;
begin
if PFAMARTICLE.Visible=True then
    BEGIN
    Row := G_Fam.Row;
    ttd := ttdCliFam;
    END else if PQTEART.Visible=True then
    BEGIN
    Row := G_Qte.Row;
    ttd := ttdCliQte;
    END else if PTART.Visible=True then
    BEGIN
    Row := G_TA.Row;
    ttd := ttdCliArt;
    END else exit;
if Row > 0 then
    BEGIN
    TOBL := GetTOBLigne (Row, ttd); if TOBL=nil then exit;
    TOBL.PutValue ('GF_CASCADEREMISE', GF_CASCADEREMISE.Value);
    END;
end;

{==============================================================================================}
{============================= Manipulation du pied ===========================================}
{==============================================================================================}
Procedure TFTarifCliArt.InitialisePied ;
BEGIN
GF_CASCADEREMISE.Text := '';
GF_REMISE.Text := '';
GF_PRIXCON.Text := '';
AffichePrixCon;
END;

Procedure TFTarifCliArt.AffichePrixCon ;
BEGIN
if (GF_DEVISE.Value<>V_PGI.DevisePivot) or (PFAMARTICLE.Visible) or (ForceEuroGC) then
    BEGIN
    GF_PRIXCON.Visible:=False;
    TGF_PRIXCON.Visible:=False;
    ISigneEuro.Visible:=False;
    ISigneFranc.Visible:=False;
    END else
    BEGIN
    GF_PRIXCON.Visible:=True;
    TGF_PRIXCON.Visible:=True;
    if VH^.TenueEuro then
        BEGIN
        ISigneEuro.Visible:=False;
        ISigneFranc.Visible:=True;
        END else
        BEGIN
        ISigneEuro.Visible:=True;
        ISigneFranc.Visible:=False;
        END;
    END;
END;

Procedure TFTarifCliArt.AffichePied (ttd : T_TableTarif) ;
var TOBL : TOB;
    FF   : TForm;
    ARow : Longint ;
BEGIN

  ARow := 0;
  if ttd = ttdCliQte then  ARow := G_Qte.Row;
  if ttd = ttdCliArt then  ARow := G_TA.Row;
  if ttd = ttdCliFam then  ARow := G_Fam.Row;
  TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL=Nil then Exit ;
  FF := TForm (PPIED.Owner);
  TOBL.PutEcran (FF, PPIED);

END;

{==============================================================================================}
{============================ Evenement lié aux Boutons =======================================}
{==============================================================================================}
procedure TFTarifCliArt.BChercherClick(Sender: TObject);
begin
if PTARIFART.Visible then
    BEGIN
    if PQTEART.Visible then
        if G_Qte.RowCount < 3 then Exit;
    if PTART.Visible then
        if G_TA.RowCount < 3 then Exit;
    END;
if PFAMARTICLE.Visible then
    if G_Fam.RowCount < 3 then Exit;
FindDebut:=True ; FindLigne.Execute ;
end;

procedure TFTarifCliArt.InfClientClick(Sender: TObject);
begin
if GF_TIERS.Text = '' then Exit;
VoirFicheTiers;
end;

procedure TFTarifCliArt.InfArticleClick(Sender: TObject);
begin
if GF_CODEARTICLE.Text = '' then Exit;
VoirFicheArticle;
end;


procedure TFTarifCliArt.BVoirCondClick(Sender: TObject);
begin
TCONDTARF.Visible := BVoirCond.Down ;
end;

procedure TFTarifCliArt.BCondAplliClick(Sender: TObject);
Var TOBL : TOB;
begin
{$IFNDEF CCS3}
if GF_TIERS.Text = '' then exit;
if PTARIFART.Visible then
    BEGIN
    if GF_CODEARTICLE.Text = '' then Exit;
    if PQTEART.Visible then
        BEGIN
        TOBL:=GetTOBLigne(G_Qte.Row, ttdCliQte) ; if TOBL=Nil then Exit ;
        EntreeTarifCond (Action, TOBL);
        AfficheCondTarf (G_Qte.Row, ttdCliQte);
        END;
    if PTART.Visible then
        BEGIN
        TOBL:=GetTOBLigne(G_TA.Row, ttdCliArt) ; if TOBL=Nil then Exit ;
        EntreeTarifCond (Action, TOBL);
        AfficheCondTarf (G_TA.Row, ttdCliArt);
        END;
    END;
if PFAMARTICLE.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_Fam.Row, ttdCliFam) ; if TOBL=Nil then Exit ;
    EntreeTarifCond (Action, TOBL);
    AfficheCondTarf (G_Fam.Row, ttdCliFam);
    END;
{$ENDIF}
end;

procedure TFTarifCliArt.BCopierCondClick(Sender: TObject);
Var TOBL : TOB;
begin
if GF_TIERS.Text = '' then exit;
if PTARIFART.Visible then
    BEGIN
    if GF_CODEARTICLE.Text = '' then Exit;
    if PQTEART.Visible then
        BEGIN
        TOBL:=GetTOBLigne(G_Qte.Row, ttdCliQte) ; if TOBL=Nil then Exit ;
        CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
        END;
    if PTART.Visible then
        BEGIN
        TOBL:=GetTOBLigne(G_TA.Row, ttdCliArt) ; if TOBL=Nil then Exit ;
        CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
        END;
    END;
if PFAMARTICLE.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_Fam.Row, ttdCliFam) ; if TOBL=Nil then Exit ;
    CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC') ;
    END;
end;

procedure TFTarifCliArt.BCollerCondClick(Sender: TObject);
Var TOBL : TOB;
begin
if GF_TIERS.Text = '' then exit;
if PTARIFART.Visible then
    BEGIN
    if GF_CODEARTICLE.Text = '' then Exit;
    if PQTEART.Visible then
        BEGIN
        TOBL:=GetTOBLigne(G_Qte.Row, ttdCliQte) ; if TOBL=Nil then Exit ;
        TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
        AfficheCondTarf (G_Qte.Row, ttdCliQte);
        END;
    if PTART.Visible then
        BEGIN
        TOBL:=GetTOBLigne(G_TA.Row, ttdCliArt) ; if TOBL=Nil then Exit ;
        TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
        AfficheCondTarf (G_TA.Row, ttdCliArt);
        END;
    END;
if PFAMARTICLE.Visible then
    BEGIN
    TOBL:=GetTOBLigne(G_Fam.Row, ttdCliFam) ; if TOBL=Nil then Exit ;
    TOBL.PutValue ('GF_CONDAPPLIC', CONDAPPLIC.Text) ;
    AfficheCondTarf (G_Fam.Row, ttdCliFam);
    END;
end;

{==============================================================================================}
{============================== Action lié aux Boutons ========================================}
{==============================================================================================}

procedure TFTarifCliArt.BSaisieRapideClick(Sender: TObject);
var i_ind : integer;
    TOBL : TOB;
begin
if Action = taConsult then exit;
if GF_CODEARTICLE.Text = '' then exit;
if (PTARIFART.Visible) and (PQTEART.Visible) then
    begin
    EntreeTarifRapide (taModif, TOBart, '', CodeTiers, CodeDevise, ttdCliQte, TOBTarfQte, tstCliArt);
    // TobTarfQte.PutGridDetail (G_QTE, True, True, LesColQtes); des que ca marche avec les dates !!
    for i_ind:=0 to TOBTarfQte.Detail.Count-1 do
        BEGIN
        TOBL:=TOBTarfQte.Detail[i_ind];
        TOBL.PutValue ('GF_NATUREAUXI', NatureTiers);
        AfficheLaLigne (i_ind + 1, ttdCliQte) ;
        END;
    end;
end;

procedure TFTarifCliArt.FindLigneFind(Sender: TObject);
begin
if PTARIFART.Visible then
    BEGIN
    if PQTEART.Visible then
        Rechercher (G_Qte, FindLigne, FindDebut) ;
    if PTART.Visible then
        Rechercher (G_TA, FindLigne, FindDebut) ;
    END;
if PFAMARTICLE.Visible then
    Rechercher (G_Fam, FindLigne, FindDebut) ;
end;

procedure TFTarifCliArt.VoirFicheArticle;
BEGIN
{$IFDEF BTP}
  V_PGI.dispatchTT (7,taConsult,CodeArticle,'','');
{$ELSE BTP }
  {$IFNDEF GPAO}
	  if ctxAffaire in V_PGI.PGIContexte then V_PGI.dispatchTT (7,taConsult,CodeArticle,'','')  //mcd 10/12/02
	  else AglLanceFiche ('GC', 'GCARTICLE', '', CodeArticle, 'ACTION=CONSULTATION;TARIF=N');
  {$ELSE GPAO }
    V_PGI.dispatchTT(7, taConsult, CodeArticle, 'TARIF=N', '');
  {$ENDIF GPAO }
{$ENDIF BTP }
END;

procedure TFTarifCliArt.VoirFicheTiers;
BEGIN
  // JT eQualité 10807 et 10867
  if TiersIsClient then
    AglLanceFiche ('GC', 'GCTIERS', '' , TOBTiers.GetValue('T_AUXILIAIRE'), 'ACTION=CONSULTATION;TARIF=N;T_NATUREAUXI=CLI')
    else
    AglLanceFiche ('GC', 'GCFOURNISSEUR', '' , TOBTiers.GetValue('T_AUXILIAIRE'), 'ACTION=CONSULTATION;TARIF=N;T_NATUREAUXI=FOU');
END;

{==============================================================================================}
{================================= Validation =================================================}
{==============================================================================================}
procedure TFTarifCliArt.BValiderClick(Sender: TObject);
Var ioerr : TIOErr ;
begin

  if Action = taConsult then exit;

  // validation
  ioerr := Transactions (ValideTarifTout, 2);
  Case ioerr of
    oeOk      : ;
    oeUnknown : BEGIN MessageAlerte(HMess.Mess[1]) ; END ;
    oeSaisie  : BEGIN MessageAlerte(HMess.Mess[2]) ; END ;
  END ;

  if Not FicheAncetre then InitialiseEntete;

  InitialiseGrille(tavFam);

end;

procedure TFTarifCliArt.ValideTarifTout;
begin
ValideTarifArticle ;
ValideTarifFamille ;
end;

procedure TFTarifCliArt.ValideTarifArticle;
begin
if Not SortDeLaLigne(ttdCliQte) then Exit ;
if Not SortDeLaLigne(ttdCliArt) then Exit ;
TOBTarifDelArt.DeleteDB (False);
VerifLesTOB (tavArt);
TOBTarfQte.InsertOrUpdateDB(False) ;
TOBTarfTA.InsertOrUpdateDB(False) ;
TOBTarfQte.SetAllModifie (False);
TOBTarfTA.SetAllModifie (False);
TOBTarifDelArt.SetAllModifie (False);
InitialiseGrille (tavArt);
DepileTOBLigne (tavArt);
end;

procedure TFTarifCliArt.ValideTarifFamille;
begin
if Not SortDeLaLigne(ttdCliFam) then Exit ;
TOBTarifDelFam.DeleteDB (False);
VerifLesTOB (tavFam);
TOBTarfFam.InsertOrUpdateDB(False) ;
TOBTarfFam.SetAllModifie (False);
TOBTarifDelFam.SetAllModifie (False);
InitialiseGrille (tavFam);
DepileTOBLigne (tavFam);
end;

procedure TFTarifCliArt.VerifLesTOB (tav : T_ActionValid);
var Q : TQuery ;
    MaxTarif : Longint ;
BEGIN
Q := OpenSQL ('SELECT MAX(GF_TARIF) FROM TARIF', TRUE,-1,'',true) ;
if Q.EOF then MaxTarif := 1 else MaxTarif := Q.Fields[0].AsInteger + 1 ;
Ferme(Q) ;
Case tav of
    tavFam : VerifLesTOBFamille (MaxTarif) ;
    tavArt : VerifLesTOBArticle (MaxTarif) ;
END;
END;

procedure TFTarifCliArt.VerifLesTOBArticle (MaxTarif : integer);
var i_ind : integer;
BEGIN
for i_ind := TOBTarfQte.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1, ttdCliQte) then
        BEGIN
        TOBTarfQte.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBTarfQte.Detail[i_ind].GetValue ('GF_TARIF') = 0 then
            BEGIN
            TOBTarfQte.Detail[i_ind].PutValue ('GF_TARIF', MaxTarif);
            Inc (MaxTarif);
            END;
        CalcPriorite (TOBTarfQte.Detail[i_ind]);
        END;
    END;
for i_ind := TOBTarfTA.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1, ttdCliArt) then
        BEGIN
        TOBTarfTA.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBTarfTA.Detail[i_ind].GetValue ('GF_TARIF') = 0 then
            BEGIN
            TOBTarfTA.Detail[i_ind].PutValue ('GF_TARIF', MaxTarif);
            Inc (MaxTarif);
            END;
        CalcPriorite (TOBTarfTA.Detail[i_ind]);
        END;
    END;
END;

procedure TFTarifCliArt.VerifLesTOBFamille (MaxTarif : integer);
var i_ind : integer;
BEGIN
for i_ind := TOBTarfFam.Detail.count - 1 Downto 0 do
    BEGIN
    if LigneVide (i_ind + 1, ttdCliFam) then
        BEGIN
        TOBTarfFam.Detail[i_ind].Free ;
        END else
        BEGIN
        if TOBTarfFam.Detail[i_ind].GetValue ('GF_TARIF') = 0 then
            BEGIN
            TOBTarfFam.Detail[i_ind].PutValue ('GF_TARIF', MaxTarif);
            Inc (MaxTarif);
            END;
        CalcPriorite (TOBTarfFam.Detail[i_ind]);
        END;
    END;
END;

{==============================================================================================}
{=============================== Conditions tarifaires ========================================}
{==============================================================================================}
procedure TFTarifCliArt.InitComboChamps ;
begin
SourisSablier ;
RemplitComboChamps('TIERS',FComboTIE) ;
RemplitComboChamps('ARTICLE',FComboART) ;
RemplitComboChamps('LIGNE',FComboLIG) ;
RemplitComboChamps('PIECE',FComboPIE) ;
SourisNormale ;
end ;

procedure TFTarifCliArt.RemplitComboChamps(NomTable : String ; FCombo : THValComboBox) ;
begin
ExtractFields(NomTable,'T',FCombo.Items,FCombo.Values,Nil,False);
end ;

procedure TFTarifCliArt.EffaceGrid ;
var lig, col : Integer ;
begin
for lig := 1 to G_COND.RowCount - 1 do
   begin
   for col := 0 to G_COND.ColCount - 1 do G_COND.Cells[col, lig] := '' ;
   end ;
G_COND.Row := 1 ;
end ;

Procedure TFTarifCliArt.AfficheCondTarf (ARow : Longint; ttd : T_TableTarif) ;
var TOBL : TOB;
BEGIN
TOBL:=GetTOBLigne(ARow, ttd) ; if TOBL=Nil then Exit ;
GetConditions (TOBL) ;
END;

function TFTarifCliArt.ValueToItem(CC : THValComboBox ; St : String) : String ;
var i_ind : Integer ;
begin
i_ind := CC.Values.IndexOf(St) ;
if i_ind >= 0 then result := CC.Items[i_ind] else result := '' ;
end ;

{Charge dans le grid les conditions stockées dans le TMemoField}
procedure TFTarifCliArt.GetConditions (TOBL : TOB) ;
var i_ind   : Integer ;
    st, NomTable, s1,s2 : String ;
begin
EffaceGrid ;
GF_CONDAPPLIC.Text := TOBL.GetValue ('GF_CONDAPPLIC');
for i_ind := 0 to GF_CONDAPPLIC.Lines.Count-1 do
   begin
   st := GF_CONDAPPLIC.Lines[i_ind] ;
   s1 := ReadTokenSt(st) ; // table
   NomTable := ValueToItem(FTable,s1) ;
   If NomTable='Client' then Nomtable :='Tiers'; //mcd 26/08/03 .. traduction à tort du nom de la table dans la fiche
   //ChargeComboChamps(s1) ; // Charge ponctuellement les champs de la table s1
   G_COND.Cells[0, i_ind + 1] := NomTable ;
   s1 := ReadTokenSt(st) ; // champ
   s2 := ValueToItem(THValComboBox(FindComponent('FCombo'+Copy(NomTable,1,3))),s1) ;
   G_COND.Cells[1, i_ind + 1] := s2 ;
   s1 := ReadTokenSt(st) ; // opérateur
   s2 := ValueToItem(FOpe,s1) ;
   G_COND.Cells[2, i_ind + 1] := s2 ;
   s1 := ReadTokenSt(st) ; // valeur
   G_COND.Cells[3, i_ind + 1] := s1 ;
   // Appel de la première ligne
   G_COND.Row := 1 ;
   end ;
end ;

procedure TFTarifCliArt.TCONDTARFClose(Sender: TObject);
begin
BVoirCond.Down := False ;
end;

Procedure AGLEntreeTarifCliArt ( Parms : array of variant ; nb : integer ) ;
Var Action : TActionFiche ;
    St     : String ;
    OkTTC : Boolean ;
BEGIN
Action:=StringToAction(String(Parms[1])) ;
OkTTC:=False ;
St:=String(Parms[2]) ; if St='TTC' then OkTTC:=True ;
EntreeTarifCliArt(Action,OkTTC) ;
END ;

procedure TFTarifCliArt.BAbandonClick(Sender: TObject);
begin
Close ;
if FClosing and IsInside(Self) then THPanel(parent).CloseInside ;
end;

procedure InitTarifCliArt();
begin
RegisterAglProc('AppeiTarifCliArt',True,4,AppeiTarifCliArt) ;
RegisterAglProc('EntreeTarifCliArt',True,2,AGLEntreeTarifCliArt) ;
end;

procedure TFTarifCliArt.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

Procedure SaisieTarifCliArtKnown (NatTiers, CodeT ,Article: string; Action : TActionFiche; TarifTTC : Boolean=False); // : boolean ;
var FF : TFTarifCliArt ;
    PPANEL  : THPanel ;
begin
  SourisSablier;

  FF := TFTarifCliArt.Create(Application) ;
  FF.Action:=Action ;
  FF.TarifTTC:=TarifTTC ;
  FF.NatureTiers:=NatTiers;
  FF.CodeTiers:=CodeT ;
  FF.TheArticle := Article;

  if NatTiers = 'CLI' then
    FF.HelpContext := 110000061
  else
    FF.HelpContext := 110000206;

  if CodeT <> '' then FF.FicheAncetre:=True else FF.FicheAncetre:=False ;

  if FF.NatureTiers='FOU' then
    BEGIN
    FF.Caption := FF.HTitre.Mess[8];
    FF.InfClient.Caption := FF.HTitre.Mess[12];
    FF.TGF_TIERS.Caption := FF.HTitre.Mess[10];
    FF.GF_TIERS.DataType:='GCTIERSFOURN';
  END
  else
    BEGIN
    FF.Caption := FF.HTitre.Mess[7] ;
    FF.InfClient.Caption := FF.HTitre.Mess[11];
    FF.TGF_TIERS.Caption := FF.HTitre.Mess[9];
    ff.GF_TIERS.DataType:='GCTIERSCLI';
    END;

  PPANEL := FindInsidePanel ;
  if PPANEL = Nil then
   BEGIN
    try
      FF.ShowModal ;
    finally
      FF.Free ;
    end ;
   SourisNormale ;
  END
  else
   BEGIN
   InitInside (FF, PPANEL) ;
   FF.Show ;
   END ;

END ;

procedure TFTarifCliArt.ReInitSousFamille(Arow: integer;ttd : T_TableTarif);
var TOBL : TOB;
begin
	TOBL := GetTOBLigne (ARow, ttd); if TOBL=nil then exit;
  Case ttd of
      ttdCliFam : BEGIN
      							TOBL.PutValue ('GF_SOUSFAMTARART','');
										G_Fam.Cells [SF_SFam, ARow] := '';
										G_Fam.Cells [SF_LibSFam, ARow] := '';
      						END;
  END;

end;

procedure TFTarifCliArt.ImportAchatClick(Sender: TObject);
begin

  AGLLanceFiche('BTP','BTIMPORTXLS','','','FOURNISSEUR=' + GF_TIERS.Text + ';TYPEIMPORT=TAR');

  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application

  CodeTiers := '';

  TraiterTiers;

end;

procedure TFTarifCliArt.ImportPrixNetClick(Sender: TObject);
begin

  If GF_Tiers.Text = '' then
  begin
    PGIError('Le code fournisseur ne peut pas être à blanc', 'Intégration Prix Net');
    GF_TIERS.SetFocus;
    Exit;
  end;

  AGLLanceFiche('BTP','BTIMPORTXLS','','','FOURNISSEUR=' + GF_TIERS.Text + ';TYPEIMPORT=PXNET');

  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application

  CodeTiers := '';

  TraiterTiers;

end;

Initialization
InitTarifCliArt();

end.
