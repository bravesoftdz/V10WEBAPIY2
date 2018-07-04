unit ConsEcr ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, StdCtrls, Hcompte, Mask, Hctrls, hmsgbox, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry, Grids, DBGrids, ExtCtrls, ComCtrls, Buttons, Hent1, Ent1, Saisie, SaisUtil,
  SaisComm, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil, UTOB,
  CPGeneraux_TOM,
  CPTiers_TOM,
  uTofCpMulMvt,
  CritEDT,
  UtilEDT,
  CPJUSTSOLD_QR1_TOF,
  ParamSoc,
  AGLInit,
  utofCPGLAuxi,
  uTofCPGLGene,
  HRichOLE, Lettrage,FE_Main,
{$IFDEF AMORTISSEMENT}
  AMLISTE_TOF,
  Outils,
  IMMO_TOM,
{$ENDIF}
  RappAuto, CumMens, SaisBor,
{$IFNDEF CCMP}
  MulAna, uTofPointageMul,
{$ENDIF}
  SaisComp,
  UtilSais,
  uLibEcriture,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ADODB;

procedure OperationsSurComptes ( LeGene,LeExo,ChampsTries : string; LeAux : string='' ; FromSaisie : boolean = False ) ;

type
  TFConsEcr = class(TFMul)
    HM: THMsgBox;
    TE_DATEECHEANCE: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    TE_DATEECHEANCE2: THLabel;
    E_DATEECHEANCE_: THCritMaskEdit;
    E_VALIDE: TCheckBox;
    Pzlibre: TTabSheet;
    Bevel5: TBevel;
    TE_TABLE0: TLabel;
    TE_TABLE2: TLabel;
    E_TABLE2: THCpteEdit;
    E_TABLE0: THCpteEdit;
    TE_TABLE1: TLabel;
    TE_TABLE3: TLabel;
    E_TABLE3: THCpteEdit;
    E_TABLE1: THCpteEdit;
    TE_REFINTERNE: THLabel;
    E_REFINTERNE: TEdit;
    PEcritures: TTabSheet;
    Bevel6: TBevel;
    TE_JOURNAL: THLabel;
    TE_EXERCICE: THLabel;
    TE_NATUREPIECE: THLabel;
    TE_NUMEROPIECE: THLabel;
    HLabel1: THLabel;
    TE_DATECOMPTABLE: THLabel;
    TE_DATECOMPTABLE2: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    E_NUMEROPIECE_: THCritMaskEdit;
    E_JOURNAL: THValComboBox;
    E_EXERCICE: THValComboBox;
    E_NATUREPIECE: THValComboBox;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_DATECOMPTABLE_: THCritMaskEdit;
    TE_GENERAL: THLabel;
    E_GENERAL: THCpteEdit;
    TE_AUXILIAIRE: THLabel;
    E_AUXILIAIRE: THCpteEdit;
    TGLIBELLE: THLabel;
    TTLIBELLE: THLabel;
    BActionCPT: TToolbarButton97;
    PopACPT: TPopupMenu;
    LETTRAGEM: TMenuItem;
    PointageG: TMenuItem;
    ModifGENE: TMenuItem;
    BActionEcr: TToolbarButton97;
    ModifAux: TMenuItem;
    PopAEdt: TPopupMenu;
    JUSTIFAUX: TMenuItem;
    Immos: TMenuItem;
    LettrageA: TMenuItem;
    CumulsGENE: TMenuItem;
    N1: TMenuItem;
    JustifGENE: TMenuItem;
    GLGENE: TMenuItem;
    GLAUX: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    ModifEcr: TMenuItem;
    SaisieEcr: TMenuItem;
    XX_WHERE1: TEdit;
    TCompte: TLabel;
    TLibelle: TLabel;
    TSoldes: TLabel;
    Analytiques: TMenuItem;
    N4: TMenuItem;
    CumulsAUX: TMenuItem;
    PopZ: TPopupMenu;
    N5: TMenuItem;
    XX_WHEREAN: TEdit;
    XX_WHERELET: TEdit;
    SaisieBor: TMenuItem;
    PParams: TTabSheet;
    Bevel7: TBevel;
    TLeSolde: THLabel;
    RDefil: TRadioGroup;
    RAction: TRadioGroup;
    InfosComp: TMenuItem;
    N6: TMenuItem;
    Delettre: TMenuItem;
    DelettreAuto: TMenuItem;
    HLabel9: THLabel;
    E_LETTRAGE: THCritMaskEdit;
    EtatLettrage: TComboBox;
    HLabel2: THLabel;
    TE_DEVISE: THLabel;
    E_DEVISE: THValComboBox;
    TE_ETABLISSEMENT: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    GeneUp: TToolbarButton97;
    GeneDown: TToolbarButton97;
    AuxiUp: TToolbarButton97;
    AuxiDown: TToolbarButton97;
    FicheImmo: TMenuItem;
    DepointageG: TMenuItem;
    N7: TMenuItem;
    E_CREERPAR: THCritMaskEdit;
    CAvecSimu: TCheckBox;
    XX_WHEREQUALIF: TEdit;
    XX_WHEREDET: TEdit;
    XX_WHEREDC: TEdit;
    LESOLDE: THNumEdit;
    GLAUXSITU: TMenuItem;
    N8: TMenuItem;
    ANALREVIS: TMenuItem;
    JUSTIFAUXSITU: TMenuItem;
    N10: TMenuItem;
    ICCDETAIL: TMenuItem;
    ICCCALCUL: TMenuItem;
    N9: TMenuItem;
    EmpruntCRE: TMenuItem;
    TPASFAIT: THEdit;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure E_EXERCICEChange(Sender: TObject);
    procedure E_GENERALExit(Sender: TObject);
    procedure E_AUXILIAIREExit(Sender: TObject);
    procedure E_GENERALEnter(Sender: TObject);
    procedure E_AUXILIAIREEnter(Sender: TObject);
    procedure FormCreate(Sender: TObject); override;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ModifGENEClick(Sender: TObject);
    procedure ModifAuxClick(Sender: TObject);
    procedure JUSTIFAUXClick(Sender: TObject);
    procedure LETTRAGEMClick(Sender: TObject);
    procedure ImmosClick(Sender: TObject);
    procedure LettrageAClick(Sender: TObject);
    procedure CumulsGENEClick(Sender: TObject);
    procedure JustifGENEClick(Sender: TObject);
    procedure GLGENEClick(Sender: TObject);
    procedure GLAUXClick(Sender: TObject);
    procedure ModifEcrClick(Sender: TObject);
    procedure SaisieEcrClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject); Override ;
    procedure AnalytiquesClick(Sender: TObject);
    procedure PointageGClick(Sender: TObject);
    procedure CumulsAUXClick(Sender: TObject);
    procedure PopACPTPopup(Sender: TObject);
    procedure PopAEdtPopup(Sender: TObject);
    procedure PopZPopup(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SaisieBorClick(Sender: TObject);
    procedure RActionClick(Sender: TObject);
    procedure InfosCompClick(Sender: TObject);
    procedure ROpposeClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure DelettreClick(Sender: TObject);
    procedure DelettreAutoClick(Sender: TObject);
    procedure EtatLettrageChange(Sender: TObject);
    procedure AuxiUpClick(Sender: TObject);
    procedure AuxiDownClick(Sender: TObject);
    procedure GeneUpClick(Sender: TObject);
    procedure GeneDownClick(Sender: TObject);
    procedure FicheImmoClick(Sender: TObject);
    procedure DepointageGClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure CAvecSimuClick(Sender: TObject);
    procedure GLAUXSITUClick(Sender: TObject);
    procedure ANALREVISClick(Sender: TObject);
    procedure JUSTIFAUXSITUClick(Sender: TObject);
    procedure ICCDETAILClick(Sender: TObject);
    procedure ICCCALCULClick(Sender: TObject);
    procedure E_GENERALKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure EmpruntCREClick(Sender: TObject);
  private
    OldGene,OldAuxi : String ;
    ToutAcc,GeneCharge,FocusUD : Boolean ;
    TOBGene,TOBTiers : TOB ;
    DroitEcritures : boolean ;
    Optim : Boolean ;
    procedure InitCriteres ;
    procedure AttribTitre ;
    Function  AjouteCrit : String ;
    Function  PositionneExo : TExoDate ;
    procedure EnabledMenus ;
    procedure GotoEntete ;
    procedure AfficheLeSoldeUnique ;
    procedure GeneSuivant ( Suiv : boolean ) ;
    procedure AuxSuivant ( Suiv : boolean ) ;
    procedure CacheZones ;
    Function  OkLettre : boolean ;
    Function  OkFicheImmo : boolean ;
    procedure GereInfoOrig ;
    procedure ChargeGeneral ( Force : boolean ) ;
    procedure ChargeAuxiliaire ( Force : boolean ) ;
    procedure EtudieModeLettrage ( Var R : RLETTR ) ;
    procedure EtudieANouveau ;
    Function  RecupWhereDataType : String ;
    Procedure AfficheLeSolde2 (T : THNumEdit ; TD,TC : Double ; PasFait : Boolean = TRUE) ;
    procedure RecupSolde ;
    procedure RecupLesNoms ;
    procedure UpdateLaCaption ;
    {JP 28/06/06 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc}
    procedure GereEtablissement;
    {JP 28/06/06 : FQ 16149 : on s'assure que le filtre coincide avec les restrictions utilisateurs sur l'établissement}
    procedure ControlEtab;
    {JP 28/06/06 : FQ 16149 : pour appeler ControlEtab à la fin du chargement des filtres}
    procedure AfterSelectFiltre;
  public
    LeGene,LeExo,ChampsTries,LeAux : String ;
    FromSaisie                     : boolean ;
  end;

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  ULibExercice,
  CPVersion,
  {$ENDIF MODENT1}
  uTofConsEcr;

{$R *.DFM}

procedure OperationsSurComptes ( LeGene,LeExo,ChampsTries : string; LeAux : string='' ; FromSaisie : boolean = False ) ;
var X : TFConsEcr ;
//    M : RMVT ;
    PP : THPanel ;
begin

//{$IFNDEF CCMP}
//if EstSpecif('51503') or GetParamSoc('SO_CONSULTSOLDEPROG') then
if ((ctxPCL in V_PGI.PGIContexte) and EstSpecif('51503')) or
      ((not (ctxPCL in V_PGI.PGIContexte)) and (GetParamSoc('SO_CONSULTSOLDEPROG')<>True)) then
begin
//{$ENDIF}
    X:=TFConsEcr.Create(Application) ;
    X.TypeAction:=taModif ; X.FNomFiltre:='CPOPERCPTE' ;
    X.Q.Manuel:=True ; X.Q.Liste:='CPCONSULTCPTES' ;
    X.LeGene:=LeGene ; X.LeExo:=LeExo ; X.ChampsTries:=ChampsTries ;
//    if EstSerie(S7) then X.HelpContext := 7602000;  // 10871
    X.LeAux:=LeAux ; X.FromSaisie:=FromSaisie ;
    PP:=FindInsidePanel ;
    if ((PP=Nil) or (V_PGI.ZoomOLE)) then
    BEGIN
      try
        X.ShowModal ;
      finally
        X.Free ;
      end;
      Screen.Cursor:=SyncrDefault ;
    END
    else
    BEGIN
      InitInside(X,PP) ;
      X.Show ;
    END ;
//{$IFNDEF CCMP}
  end
  else
  begin
    uTofConsEcr.OperationsSurComptes( LeGene, LeExo, ChampsTries, LeAux, FromSaisie ) ;
  end;
//{$ENDIF}
END ;

procedure TFConsEcr.AttribTitre ;
BEGIN
//HelpContext:=0 ;
UpdateCaption(Self) ;
END ;

procedure TFConsEcr.EtudieANouveau ;
Var DD : TDateTime ;
BEGIN
XX_WHEREAN.Text:='E_ECRANOUVEAU="N"' ;
DD:=StrToDate(E_DATECOMPTABLE.Text) ; if (VH^.Suivant.Deb>0) And (DD>VH^.Suivant.Deb) then Exit ;
XX_WHEREAN.Text:='E_ECRANOUVEAU="N" OR ((E_ECRANOUVEAU="H" OR E_ECRANOUVEAU="OAN") AND E_DATECOMPTABLE="'+UsDateTime(DD)+'")' ;
END ;


procedure TFConsEcr.CacheZones ;
Var i : integer ;
    VisuGene,VisuAux : boolean ;
BEGIN
if (ctxPCL in V_PGI.PGIContexte)=FALSE Then Exit ;
VisuGene:=True ; VisuAux:=True ;
if TOBGene.GetValue('G_GENERAL')<>'' then
   BEGIN
   VisuGene:=False ;
   if TOBGene.GetValue('G_COLLECTIF')='X' then
      BEGIN
      if TOBTiers.GetValue('T_AUXILIAIRE')<>'' then VisuAux:=False ;
      END else
      BEGIN
      VisuAux:=False ;
      END ;
   END else
   BEGIN
   VisuAux:=False ;
   END ;
for i:=0 to FListe.Columns.Count-1 do
    BEGIN
    if FListe.Columns[i].FieldName='E_GENERAL' then FListe.Columns[i].Visible:=VisuGene ;
    if FListe.Columns[i].FieldName='E_AUXILIAIRE' then FListe.Columns[i].Visible:=VisuAux ;
    END ;
if ((HMTrad.ActiveResize) and (HMTrad.ResizeDBGrid)) then
   BEGIN
   HMTrad.ResizeDBGridColumns(FListe) ;
   FListe.Refresh ;
   END ;
END ;

procedure TFConsEcr.RecupSolde ;
BEGIN
If Optim Then RecupLesNoms ;
END ;

procedure TFConsEcr.BChercheClick(Sender: TObject);
Var SavePlace: TBookmark;
begin

//RR gestion de la confidentialité
if EstSQLConfidentiel('GENERAUX',E_GENERAL.Text) or estsqlconfidentiel('TIERS',E_AUXILIAIRE.text) then
begin
  //PGIBox(traduirememoire('Compte confidentiel !' + #13#10 + #13#10 + ' Accés refusé'));
  Caption:='Le compte ' + E_GENERAL.TEXT + traduirememoire(' est confidentiel')  ;
  //Caption:=traduirememoire(' Accés confidentiel !!!')  ;
  UpdateCaption(Self) ;
  exit ;
end ;

EtudieANouveau ;
// VL
if Not Optim then begin AfficheLeSoldeUnique; LESOLDE.Visible:=True; TPASFAIT.Visible:=False; end;
// VL
E_GENERALExit(Nil) ; E_AUXILIAIREExit(Nil) ;
if ((TOBGene.GetValue('G_GENERAL')='') and (TOBTiers.GetValue('T_AUXILIAIRE')='')) then
   if Not GeneCharge then
   BEGIN
   HM.Execute(9,Caption,'') ;
   if E_GENERAL.CanFocus then E_GENERAL.SetFocus ;
   Exit ;
   END ;
if (Sender=Nil) and Not((Q.EOF) and (Q.BOF)) then SavePlace:=Q.GetBookmark else SavePlace:=Nil ;
  inherited;
CacheZones ;
if (Sender=Nil) and Not((Q.EOF) and (Q.BOF)) then if SavePlace<>Nil then
   BEGIN
   Try
   Q.GotoBookmark(SavePlace) ;
   Q.FreeBookmark(SavePlace) ;
   Except
   End ;
   END ;
RecupSolde ;
end;

procedure TFConsEcr.GereInfoOrig ;
BEGIN
if LeGene<>'' then
   BEGIN
   E_GENERAL.Text:=LeGene ; ChargeGeneral(False) ;
   if LeAux<>'' then E_AUXILIAIRE.Text:=LeAux else E_AUXILIAIRE.Text:='' ;
   ChargeAuxiliaire(False) ;
   XX_WHERE1.Text:='' ;
   END ;
if LeExo='0' then E_EXERCICE.Value:=VH^.Encours.Code else
 if LeExo='1' then E_EXERCICE.Value:=VH^.Suivant.Code else
  if LeExo='-1' then E_EXERCICE.Value:=VH^.Precedent.Code ;
END ;

procedure TFConsEcr.UpdateLaCaption ;
Var StC,StSolde : String ;
BEGIN
If LESOLDE.Visible Then StSolde:=LeSolde.Text Else StSolde:='' ;
StC:=Copy(Caption,1,Pos(':',Caption)) ;
if ((E_AUXILIAIRE.Text<>'') and (E_GENERAL.Text<>'')) then
   BEGIN
   StC:=StC+' '+E_GENERAL.Text+'  '+E_AUXILIAIRE.Text+'  '+TTLIBELLE.Caption+'   '+StSolde ;
   END else
   BEGIN
   if E_AUXILIAIRE.Text<>'' then StC:=StC+'  '+E_AUXILIAIRE.Text+'  '+TTLIBELLE.Caption+'   '+StSolde else
     if E_GENERAL.Text<>'' then StC:=StC+' '+E_GENERAL.Text+'  '+TGLIBELLE.Caption+'   '+StSolde
                           else StC:=StC ;
   END ;
Caption:=StC ; UpdateCaption(Self) ;
END ;

procedure TFConsEcr.RecupLesNoms ;
Var HNE : THnumEdit ;
    i : Integer ;
    St : String ;
    XD,XC : Double ;
    OkD,OkC : Boolean ;
    OkDD,OkCC : Boolean ;
BEGIN
{OkD:=FALSE ; OkC:=FALSE ;} XD:=0 ; XC:=0 ; OKDD:=FALSE ; OKCC:=FALSE ;
for i:=0 to ComponentCount-1 do
   BEGIN
   if (Components[i] is THNUMEDIT) then
     BEGIN
     HNE:=THNumEdit(Components[i]) ;
     If Copy(HNE.Name,1,5)='__QRY' Then
       BEGIN
       St:=HNE.Name ;
       OkD:=(Pos('E_DEBIT',St)>0) ;
       OkC:=(Pos('E_CREDIT',St)>0) ;
       If OkD Then BEGIN OKDD:=TRUE ; XD:=HNE.Value ; END ;
       If OkC Then BEGIN OKCC:=TRUE ; XC:=HNE.Value ; END ;
       END ;
     END ;
   END ;
If OkDD And OkCC Then AfficheLeSolde2(LeSolde,XD,XC,FALSE) ;
UpdateLaCaption ;
END ;

procedure TFConsEcr.FormShow(Sender: TObject);
Var  Simul : String ;
begin
If ctxPCL in V_PGI.PGIContexte Then Optim:=FALSE Else Optim:=Not GetParamSoc('SO_BDEFGCPTE'); //Not EstSpecif('51194') ;
//If Optim Then
//  BEGIN
//  GeneUp.Visible:=FALSE ; GeneDown.Visible:=FALSE ; AuxiUp.Visible:=FALSE ; AuxiDown.Visible:=FALSE ;
//  END ;
MakeZoomOLE(Handle) ;
GeneCharge:=True ;
DroitEcritures:=ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False) ;
if Not DroitEcritures then BEGIN RAction.ItemIndex:=0 ; RAction.Enabled:=False ; END ;
InitCriteres ; LibellesTableLibre(PzLibre,'TE_TABLE','E_TABLE','E') ;
Simul:='N' ; EtatLettrage.ItemIndex:=0 ;
if VH^.CPExoRef.Code<>'' then
   BEGIN
   E_EXERCICE.Value:=VH^.CPExoRef.Code ;
   E_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
   E_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   E_EXERCICE.Value:=VH^.Entree.Code ;
   E_DATECOMPTABLE.Text:=DateToStr(VH^.Entree.Deb) ;
   E_DATECOMPTABLE_.Text:=DateToStr(VH^.Entree.Fin) ;
   END ;
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
AttribTitre ;
E_JOURNAL.DataType:='TTJOURNAUX' ; XX_WHEREAN.Text:='E_ECRANOUVEAU="N"' ;
TGLibelle.Caption:='' ; TTLibelle.Caption:='' ; OldGene:='' ; OldAuxi:='' ;
ChangeMask(LeSolde,V_PGI.OkDecV,'') ;
GereInfoOrig ;
if ctxPCL in V_PGI.PGIContexte then CAvecSimu.Visible:=False ;
  inherited;
  //SetTousCombo(E_JOURNAL) ; SetTousCombo(E_DEVISE) ; SetTousCombo(E_NATUREPIECE) ;
  E_JOURNAL    .ItemIndex := 0;
  E_DEVISE     .ItemIndex := 0;
  E_NATUREPIECE.ItemIndex := 0;

if ctxPCL in V_PGI.PGIContexte then BParamListe.Visible:=False ;
GeneCharge:=False ;
PSQL.TabVisible:=False ;
Q.Manuel:=FALSE ; Q.UpdateCriteres ;
CentreDBGrid(FListe) ;
XX_WHERE1.Text:='' ;
CacheZones ;

  {JP 28/06/06 : FQ 16149 : refonte de la gestion des établissements}
  GereEtablissement;
  OnAfterSelectFiltre := AfterSelectFiltre;;
end;

procedure TFConsEcr.InitCriteres ;
BEGIN
if VH^.Precedent.Code<>'' then E_DATECOMPTABLE.Text:=DateToStr(VH^.Precedent.Deb)
                          else E_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
END ;

procedure TFConsEcr.FListeDblClick(Sender: TObject);
Var {Per,Jal,sNum,}sMode : String ;
    AA  : TActionFiche ;
    M   : RMVT ;
begin
  inherited;
if (Q.Eof) And (Q.Bof) then Exit ;
AA:=taModif ; if Not ToutAcc then AA:=taConsult ;
sMode:=Q.FindField('E_MODESAISIE').AsString ;
if ((sMode<>'-') and (sMode<>'')) then
   BEGIN
   LanceSaisieFolio(Q,AA) ;
   END else
   BEGIN
   if TrouveSaisie(Q,M,'N') then
      BEGIN
      M.NumLigVisu:=Q.FindField('E_NUMLIGNE').AsInteger ;
      LanceSaisie(Q,AA,M) ;
      END else if CAvecSimu.Checked then
      BEGIN
      if TrouveSaisie(Q,M,'S') then
        BEGIN
        M.NumLigVisu:=Q.FindField('E_NUMLIGNE').AsInteger ;
        LanceSaisie(Q,AA,M) ;
        END ;
      END ;
   END ;
ChargeGeneral(True) ; ChargeAuxiliaire(True) ;
// BPY le 24/03/2004 => bug 10476
//If ToutAcc Then BChercheClick(Nil) ;
end;

procedure TFConsEcr.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
if E_EXERCICE.Value='' then BEGIN E_DATECOMPTABLE.Text:=stDate1900 ; E_DATECOMPTABLE_.Text:=stDate2099 ; END ;
end;

procedure TFConsEcr.ChargeGeneral ( Force : boolean ) ;
BEGIN
if ((TOBGene=Nil) or (csDestroying in ComponentState)) then Exit ;
//AfficheLeSoldeUnique ;
// VL
if not Optim then AfficheLeSoldeUnique ;
// VL
  inherited;
if E_GENERAL.Text='' then
   BEGIN
   TOBGene.InitValeurs ; TGLibelle.Caption:='' ;
   E_AUXILIAIRE.ZoomTable:=tzTiers ;
   Exit ;
   END ;
if ((E_GENERAL.Text=OldGene) and (Not Force)) then Exit ;
if E_GENERAL.ExisteH<=0 then
   BEGIN
   if Not GChercheCompte(E_GENERAL,Nil) then TGLibelle.Caption:='' ;
   END ;
if Not TOBGene.SelectDB('"'+E_GENERAL.Text+'"',Nil,False) then
   BEGIN
   TOBGene.InitValeurs ;
   END else
   BEGIN
   if ((TOBGene.GetValue('G_COLLECTIF')<>'X') and (TOBTiers.GetValue('T_AUXILIAIRE')<>'')) then BEGIN E_AUXILIAIRE.Text:='' ; ChargeAuxiliaire(False) ; END ;
   END ;
TGLibelle.Caption:=TOBGene.GetValue('G_LIBELLE') ;
E_AUXILIAIRE.ZoomTable:=QuelZoomTableTNat(TOBGene.GetValue('G_NATUREGENE')) ;
//AfficheLeSoldeUnique ;
// VL
if not Optim then AfficheLeSoldeUnique ;
// VL
END ;

procedure TFConsEcr.E_GENERALExit(Sender: TObject);
begin
//RR gestion de la confidentialité
if EstSQLConfidentiel('GENERAUX',E_GENERAL.Text) or estsqlconfidentiel('TIERS',E_AUXILIAIRE.text) then
begin
  //PGIBox(traduirememoire('Compte confidentiel !' + #13#10 + #13#10 + ' Accés refusé'));
  Caption:='Le compte ' + E_GENERAL.TEXT + traduirememoire(' est confidentiel')  ;
  //Caption:=traduirememoire(' Accés confidentiel !!!')  ;
  UpdateCaption(Self) ;
  exit;
end
else
  ChargeGeneral(False) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 15/07/2002
Modifié le ... :   /  /
Description .. : LG - 15/07/2002 - utilisation de la fct NATUREGENAUXOK
Suite ........ : à la place de CControleGeneAvecTiers
Mots clefs ... :
*****************************************************************}
procedure TFConsEcr.ChargeAuxiliaire ( Force : boolean ) ;
var lTobGene, lTobTiers: Tob;
BEGIN
if ((TOBTiers=Nil) or (csDestroying in ComponentState)) then Exit ;
//AfficheLeSoldeUnique ;
// VL
if not Optim then AfficheLeSoldeUnique ;
// VL
  inherited;
if E_AUXILIAIRE.Text='' then BEGIN TOBTiers.InitValeurs ; TTLibelle.Caption:='' ; Exit ; END ;
if ((E_AUXILIAIRE.Text=OldAuxi) and (Not Force)) then Exit ;

// GC - 27/02/2002
if (E_AUXILIAIRE.Text <> '') and (E_GENERAL.Text <> '') then
begin
  lTobGene  := TOB.Create('GENERAUX',Nil,-1) ;
  lTobTiers := TOB.Create('TIERS',Nil,-1) ;
  lTobGene.SelectDB('"'+E_GENERAL.Text+'"',Nil,False);
  lTobTiers.SelectDB('"'+E_AUXILIAIRE.Text+'"',Nil,False);
//  if not CControleGeneAvecTiers( lTobGene.GetValue('G_NATUREGENE'), lTobTiers.GetValue('T_NATUREAUXI')) then
  if not NATUREGENAUXOK( lTobGene.GetValue('G_NATUREGENE'), lTobTiers.GetValue('T_NATUREAUXI')) then
  begin
    E_GENERAL.Text := '';
    ChargeGeneral( False );
  end;
  if lTobGene <> nil then lTobGene.Free;
  if lTobTiers <> nil then lTobTiers.Free;
end;
// Fin GC

if E_AUXILIAIRE.ExisteH<=0 then
   BEGIN
   if Not GChercheCompte(E_AUXILIAIRE,Nil) then TTLibelle.Caption:='' ;
   END ;
if Not TOBTiers.SelectDB('"'+E_AUXILIAIRE.Text+'"',Nil,False) then
   BEGIN
   TOBTiers.InitValeurs ;
   END else
   BEGIN
   if TOBGene.GetValue('G_COLLECTIF')<>'X' then BEGIN E_GENERAL.Text:=TOBTiers.GetValue('T_COLLECTIF') ; ChargeGeneral(False) ; END ;
   END ;
TTLibelle.Caption:=TOBTiers.GetValue('T_LIBELLE') ;
//AfficheLeSoldeUnique ;
// VL
if not Optim then AfficheLeSoldeUnique ;
// VL
END ;

procedure TFConsEcr.E_AUXILIAIREExit(Sender: TObject);
begin
//RR gestion de la confidentialité
if EstSQLConfidentiel('GENERAUX',E_GENERAL.Text) or estsqlconfidentiel('TIERS',E_AUXILIAIRE.text) then
begin
  Caption:='Le compte ' + E_GENERAL.TEXT + traduirememoire(' est confidentiel')  ;
  //Caption:=traduirememoire(' Accés confidentiel !!!')  ;
  UpdateCaption(Self) ;
  exit;
end
else
  ChargeAuxiliaire(False) ;
end;

procedure TFConsEcr.E_GENERALEnter(Sender: TObject);
begin
  inherited;
  OldGene:=E_GENERAL.Text ;
end;

procedure TFConsEcr.E_AUXILIAIREEnter(Sender: TObject);
begin
  inherited;
OldAuxi:=E_AUXILIAIRE.Text ;
end;

Procedure TFConsEcr.AfficheLeSolde2 (T : THNumEdit ; TD,TC : Double ; PasFait : Boolean = TRUE) ;
BEGIN
If Not Optim Then AfficheLeSolde(T,TD,TC) Else
  BEGIN
  If PasFait Then
    BEGIN
    LESOLDE.Visible:=FALSE ; TPASFAIT.Visible:=TRUE ;
    END Else
    BEGIN
    LESOLDE.Visible:=TRUE ; TPASFAIT.Visible:=FALSE ;
    AfficheLeSolde(T,TD,TC) ;
    END ;
  END ;
END ;

procedure TFConsEcr.AfficheLeSoldeUnique ;
Var Gene,Auxi,{SQL,}StC,chD,chC : String ;
    QQ        : TQuery ;
    XD,XC     : Double ;
    sReq      : String ;
BEGIN
Gene:='' ; Auxi:='' ; XD:=0 ; XC:=0 ;
if ((TOBTiers.GetValue('T_AUXILIAIRE')<>'') and (E_AUXILIAIRE.Text<>'')) then Auxi:=TOBTiers.GetValue('T_AUXILIAIRE') ;
if ((TOBGene.GetValue('G_GENERAL')<>'') and (E_GENERAL.Text<>'')) then Gene:=TOBGene.GetValue('G_GENERAL') ;
if Not Optim And (((Auxi<>'') or (Gene<>''))) then
   BEGIN
    chD:='E_DEBIT' ; chC:='E_CREDIT' ;
   SReq:='SELECT SUM('+chD+'), SUM('+chC+') FROM ECRITURE '+RecupWhereCritere(Pages) ;
   QQ:=OpenSQL(sReq,True) ;
   if Not QQ.EOF then
      BEGIN
      XD:=QQ.Fields[0].AsFloat ; XC:=QQ.Fields[1].AsFloat ;
      END ;
   Ferme(QQ) ;
   END ;
AfficheLeSolde2(LeSolde,XD,XC) ;
If Optim Then UpdateLaCaption Else
  BEGIN
  StC:=Copy(Caption,1,Pos(':',Caption)) ;
  if ((E_AUXILIAIRE.Text<>'') and (E_GENERAL.Text<>'')) then
   BEGIN
   StC:=StC+' '+E_GENERAL.Text+'  '+E_AUXILIAIRE.Text+'  '+TTLIBELLE.Caption+'   '+LESOLDE.Text + '  du ' + E_DateComptable.Text + ' au ' + E_DateComptable_.Text ;
   END else
   BEGIN
   if E_AUXILIAIRE.Text<>'' then StC:=StC+'  '+E_AUXILIAIRE.Text+'  '+TTLIBELLE.Caption+'   '+LESOLDE.Text + '  du ' + E_DateComptable.Text + ' au ' + E_DateComptable_.Text else
     if E_GENERAL.Text<>'' then StC:=StC+' '+E_GENERAL.Text+'  '+TGLIBELLE.Caption+'   '+LESOLDE.Text + '  du ' + E_DateComptable.Text + ' au ' + E_DateComptable_.Text
                           else StC:=StC ;
   END ;
  Caption:=StC ; UpdateCaption(Self) ;
  END ;
END ;

procedure TFConsEcr.GotoEntete ;
Var i : integer ;
BEGIN
if E_GENERAL.CanFocus then E_GENERAL.SetFocus else
 if E_AUXILIAIRE.CanFocus then E_AUXILIAIRE.SetFocus else
    BEGIN
    for i:=0 to Pages.ActivePage.ControlCount-1 do if Pages.ActivePage.Controls[i] is TWinControl then
        if TWinControl(Pages.ActivePage.Controls[i]).CanFocus then
        BEGIN
        TWinControl(Pages.ActivePage.Controls[i]).SetFocus ;
        Break ;
        END ;
    END ;
END ;

procedure TFConsEcr.FormCreate(Sender: TObject);
begin
  inherited;
TOBGene:=TOB.Create('GENERAUX',Nil,-1) ;
TOBTiers:=TOB.Create('TIERS',Nil,-1) ;
Q.Manuel:=True ; ToutAcc:=True ;
MemoStyle:=msBook ; FocusUD:=False ;
end;

procedure TFConsEcr.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : boolean ;
    lKey : Word;
begin
  lKey:=0;
  if Key = VK_F5 then
  begin
    lKey := Key;
    Key := 0;
  end;
  inherited;
  if lKey = VK_F5 then Key := lKey;
Vide:=(Shift=[]) ;
Case Key of
   VK_NEXT  : BEGIN
              if FListe.Focused then Exit ;
              If Optim Then Exit ;
              Key:=0 ;
              if E_GENERAL.Focused then GeneSuivant(True) else if E_AUXILIAIRE.Focused then AuxSuivant(True) ;
              END ;
   VK_PRIOR : BEGIN
              if FListe.Focused then Exit ;
              If Optim Then Exit ;
              Key:=0 ;
              if E_GENERAL.Focused then GeneSuivant(False) else if E_AUXILIAIRE.Focused then AuxSuivant(False) ;
              END ;
   VK_F5 : if Vide then
              BEGIN
              Key:=0 ;
              if FListe.Focused then FListeDblClick(Nil) ;
              if E_GENERAL.Focused then BEGIN if E_GENERAL.Text='' then E_GENERAL.DblClick else ChargeGeneral(True) ; END ;
              if E_AUXILIAIRE.Focused then BEGIN if E_AUXILIAIRE.Text='' then E_AUXILIAIRE.DblClick else ChargeAuxiliaire(True) ; END ;
              END ;
   VK_F6 : if Vide then BEGIN Key:=0 ; if FListe.Focused then GotoEntete else FListe.SetFocus ; END ;
  VK_F11 : BEGIN
           Key:=0 ; PopZ.PopUp(Mouse.CursorPos.X,Mouse.CursorPos.Y) ;
           END ;
{AltC}67 : BEGIN
           if ssAlt in Shift then BEGIN Key:=0 ; InfosCompClick(Nil) ; END ; 
           END ;
   END ;
end;

procedure TFConsEcr.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
TOBGene.Free ; TOBGene:=Nil ;
TOBTiers.Free ; TOBTiers:=Nil ;
PurgePopUp(PopZ) ;
end;

Function TFConsEcr.PositionneExo : TExoDate ;
Var LeExo : TExoDate ;
BEGIN
LeExo:=VH^.Entree ;
if E_EXERCICE.Value=VH^.Encours.Code then LeExo:=VH^.Encours else
 if E_EXERCICE.Value=VH^.Suivant.Code then LeExo:=VH^.Suivant else
   if E_EXERCICE.Value=VH^.Precedent.Code then LeExo:=VH^.Precedent ;
Result:=LeExo ;
END ;

{Lettrage manuel}
Function TFConsEcr.AjouteCrit : String ;
Var j,i    : integer ;
    P      : TTabSheet ;
    Nam    : String ;
    StEcr  : HString;
    C      : TControl ;
BEGIN
StEcr:='' ;
for j:=0 to Pages.PageCount-1 do
    BEGIN
    P:=Pages.Pages[j] ;
    for i:=0 to P.ControlCount-1 do
        BEGIN
        C:=P.Controls[i] ; if C.Tag=1 then Continue ;
        Nam:=C.Name ; if Nam[Length(Nam)]='_' then System.Delete(Nam,Length(Nam),1) ;
        if Copy(Nam,1,2)='E_' then
           BEGIN
           Q.Control2Criteres(Nam,StEcr,C,P) ;
           END ;
        END ;
    END ;
if Copy(StEcr,2,3)='AND' then System.Delete(StEcr,1,5) ; if Copy(StEcr,2,2)='OR' then System.Delete(StEcr,1,4) ;
if Trim(StEcr)<>'' then StEcr:=' AND (('+StEcr+' AND E_ETATLETTRAGE="AL") OR (E_ETATLETTRAGE="PL")) ' ;
Result:=StEcr ;
END ;

procedure TFConsEcr.EtudieModeLettrage ( Var R : RLETTR ) ;
Var ii : integer ;
BEGIN
if R.DeviseMvt=V_PGI.DevisePivot then
   BEGIN
   R.CritDev:=V_PGI.DevisePivot ;
   END else
   BEGIN
   if R.DeviseMvt=R.CritDev then
      BEGIN
      {Paquet en devise --> Lettrage devise}
      R.LettrageDevise:=True ;
      END else
      BEGIN
      {Paquet en devise (non explicite) --> poser la question}
      ii:=HM.Execute(11,'',' '+RechDom('ttDevise',R.DeviseMvt,False)) ;
      R.LettrageDevise:=(ii=mrYes) ;
      if R.LettrageDevise then R.CritDev:=R.DeviseMvt else R.Distinguer:=False ;
      END ;
   END ;
END ;

procedure TFConsEcr.LETTRAGEMClick(Sender: TObject);
Var R : RLETTR ;
    AA : TActionFiche ;
begin
if CAvecSimu.Checked then Exit ;
  inherited;
FillChar(R,Sizeof(R),#0) ; AA:=taModif ; if Not ToutAcc then AA:=taConsult ;
R.General:=E_GENERAL.Text ; R.Auxiliaire:=E_AUXILIAIRE.Text ; R.Appel:=tlMenu ;
R.CritDev:=E_DEVISE.Value ;
if ((Q.EOF) and (Q.BOF)) then R.DeviseMvt:=V_PGI.DevisePivot else R.DeviseMvt:=Q.FindField('E_DEVISE').AsString ;
R.GL:=Nil ; R.CritMvt:=AjouteCrit ;
EtudieModeLettrage(R) ;
LettrageManuel(R,True,AA) ;
BChercheClick(Nil) ;
end;

{DéLettrage manuel}
procedure TFConsEcr.DelettreClick(Sender: TObject);
Var M  : RMVT ;
    R  : RLETTR ;
    QQ : TQuery ;
    AA : TActionFiche ;
    Trouv : boolean ;
begin
Trouv := False;
if CAvecSimu.Checked then Exit ;
  inherited;
AA:=taModif ;
if ((Q.EOF) and (Q.BOF)) then Exit ;
if TrouveSaisie(Q,M,'N') then
   BEGIN
   M.NumLigne:=Q.FindField('E_NUMLIGNE').AsInteger ;
   M.NumEche:=Q.FindField('E_NUMECHE').AsInteger ;
   QQ:=OpenSQL('Select * from ECRITURE WHERE '+WhereEcriture(tsGene,M,True),True) ;
   if Not QQ.EOF then
      BEGIN
      FillChar(R,Sizeof(R),#0) ; Trouv:=True ;
      R.General:=QQ.FindField('E_GENERAL').AsString ;
      R.Auxiliaire:=QQ.FindField('E_AUXILIAIRE').AsString ;
      R.CritDev:=QQ.FindField('E_DEVISE').AsString ; R.DeviseMvt:=R.CritDev ;
      R.CodeLettre:=QQ.FindField('E_LETTRAGE').AsString ;
      R.LettrageDevise:=(QQ.FindField('E_LETTRAGEDEV').AsString='X') ;
      R.Appel:=tlMenu ; R.ToutSelDel:=True ;
      END ;
   Ferme(QQ) ;
   if Trouv then
      BEGIN
      if Not ToutAcc then AA:=taConsult ;
      LettrageManuel(R,False,AA) ;
      BChercheClick(Nil) ;
      END ;
   END ;
end;

{DéLettrage automatiue}
procedure TFConsEcr.DelettreAutoClick(Sender: TObject);
Var SQL : String ;
begin
if CAvecSimu.Checked then Exit ;
  inherited;
if HM.Execute(10,Caption,'')<>mrYes then Exit ;
SQL:='UPDATE ECRITURE SET E_LETTRAGE="", E_ETATLETTRAGE="AL", E_ETAT="0000000000", E_DATEPAQUETMIN="'+UsDateTime(iDate1900)+'", E_DATEPAQUETMAX="'+UsDateTime(iDate1900)+'", '
    +'E_REFLETTRAGE="", E_COUVERTURE=0, E_LETTRAGEDEV="-" ' ;
SQL:=SQL+' WHERE E_GENERAL="'+E_GENERAL.Text+'" AND E_AUXILIAIRE="'+E_AUXILIAIRE.Text+'" ' ;
ExecuteSQL(SQL) ;
BChercheClick(Nil) ;
end;

{Lettrage automatique}
procedure TFConsEcr.LettrageAClick(Sender: TObject);
begin
if CAvecSimu.Checked then Exit ;
  inherited;
RapprochementAuto(E_GENERAL.Text,E_AUXILIAIRE.Text) ;
BChercheClick(Nil) ;
end;

{Pointage du compte}
procedure TFConsEcr.PointageGClick(Sender: TObject);
Var StG : String ;
begin
  inherited;
StG:=TOBGene.GetValue('G_GENERAL') ;
if StG<>'' then
   BEGIN
{$IFNDEF CCMP}
  CPLanceFiche_PointageMul( StG);
   BChercheClick(Nil) ;
{$ENDIF}
   END ;
end;

procedure TFConsEcr.DepointageGClick(Sender: TObject);
Var StG : String ;
begin
  inherited;
StG:=TOBGene.GetValue('G_GENERAL') ;
if StG<>'' then
   BEGIN
{$IFNDEF CCMP}
  CPLanceFiche_PointageMul( StG );
   BChercheClick(Nil) ;
{$ENDIF}
   END ;
end;

{Modif écriture}
procedure TFConsEcr.ModifEcrClick(Sender: TObject);
begin
  inherited;
FListeDBLClick(Nil) ;
end;

{Informations complémentaires ligne}
procedure TFConsEcr.InfosCompClick(Sender: TObject);
Var M   : RMVT ;
    QQ  : TQuery ;
    OBM : TOBM ;
    RC  : R_COMP ;
    AA  : TActionFiche ;
    ModBN,Okok : Boolean ;
    tobtmp : TOB ; //SG6 25/01/05
begin
if Not InfosComp.Enabled then Exit ;
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ;

Okok:=TrouveSaisie(Q,M,'N') ;
if ((Not Okok) and (CAvecSimu.Checked)) then Okok:=TrouveSaisie(Q,M,'S') ;
if Okok then
   BEGIN
   OBM:=Nil ; AA:=taModif ; if Not ToutAcc then AA:=taConsult ;
   M.NumLigne:=Q.FindField('E_NUMLIGNE').AsInteger ;
   M.NumEche:=Q.FindField('E_NUMECHE').AsInteger ;
   QQ:=OpenSQL('Select * from ECRITURE WHERE '+WhereEcriture(tsGene,M,True),True) ;

   if Not QQ.EOF then
      BEGIN
      OBM:=TOBM.Create(EcrGen,'',false) ;
      OBM.ChargeMvt(QQ) ;
      if Not TMemoField(QQ.FindField('E_BLOCNOTE')).IsNull then OBM.M.Assign(TMemoField(QQ.FindField('E_BLOCNOTE'))) ;
      {b FP 15/11/2005 FQ15311 Il faut remplir tobtmp avant de fermer le dataset QQ}
      tobtmp := TOB.Create('ECRITURE',nil,-1);
      tobtmp.SelectDB('',QQ);
      {e FP 15/11/2005 FQ15311}
      END ;
      Ferme(QQ) ;
      if OBM <> nil then
      begin
        RC.StLibre := 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
        RC.StComporte := 'XXXXXXXXXX';
        ModBN := True;
        RC.Conso := True;
        RC.Attributs := False;
        RC.MemoComp := nil;
        RC.Origine := -1;
        RC.TOBCompl := nil ;
        //SG6 25/01/05 FQ 15311 On pass par un tob
        {FP 15/11/2005 FQ15311  tobtmp := TOB.Create('ECRITURE',nil,-1);
        tobtmp.SelectDB('',QQ);}
        if SaisieComplement(TOBM(tobtmp), EcrGen, AA, ModBN, RC, False, True) then
        begin
          tobtmp.UpdateDb;
          BChercheClick(Nil) ;
        end;
        FreeAndNil(tobtmp);
     end;
   END ;
end;

{Nouvelle écriture pièce}
procedure TFConsEcr.SaisieEcrClick(Sender: TObject);
Var OldSC : Boolean ;
begin
  inherited;
OldSC:=VH^.BouclerSaisieCreat ; VH^.BouclerSaisieCreat:=False ;
MultiCritereMvt(taCreat,'N',False) ;
VH^.BouclerSaisieCreat:=OldSC ;
ChargeGeneral(True) ; ChargeAuxiliaire(True) ;
If Not Optim Then BChercheClick(Nil) ;
end;

{Nouvelle écriture folio}
Procedure TFConsEcr.SaisieBorClick(Sender: TObject);
begin
  inherited;
SaisieFolio(taModif) ;
ChargeGeneral(True) ; ChargeAuxiliaire(True) ;
If Not Optim Then BChercheClick(Nil) ;
end;

{Modif compte général}
procedure TFConsEcr.ModifGENEClick(Sender: TObject);
Var AA : TActionFiche ;
begin
if TOBGene.GetValue('G_GENERAL')='' then Exit ;
  inherited;
AA:=taModif ; if ((not ExJaiLeDroitConcept(TConcept(ccGenModif),False)) or (Not ToutAcc)) then AA:=taConsult ;
FicheGene(Nil,'',TOBGene.GetValue('G_GENERAL'),AA,0) ;
end;

{Modif compte auxiliaire}
procedure TFConsEcr.ModifAuxClick(Sender: TObject);
Var AA : TActionFiche ;
begin
if TOBTiers.GetValue('T_AUXILIAIRE')='' then Exit ;
  inherited;
AA:=taModif ; if ((not ExJaiLeDroitConcept(TConcept(ccAuxModif),False)) or (Not ToutAcc)) then AA:=taConsult ;
FicheTiers(Nil,'',TOBTiers.GetValue('T_AUXILIAIRE'),AA,0) ;
end;

{Mouvements analytiques du compte}
procedure TFConsEcr.AnalytiquesClick(Sender: TObject);
{$IFNDEF CCMP}
Var D1,D2 : TDateTime ;
    Crit  : TCritEDT ;
    AA    : TActionFiche ;
{$ENDIF}
begin
  inherited;
{$IFNDEF CCMP}
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=StrToDate(E_DATECOMPTABLE.text) ; D2:=StrToDate(E_DATECOMPTABLE_.text) ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.Exo:=PositionneExo ;
Crit.SCpt1:=E_GENERAL.Text ;
AA:=taModif ; if Not ToutAcc then AA:=taConsult ;
MultiCritereAnaZoom(AA,Crit) ;
{$ENDIF}
end;

{Liste des immos du compte}
procedure TFConsEcr.ImmosClick(Sender: TObject);
{$IFDEF AMORTISSEMENT}
{$ENDIF}
begin
  inherited;
{$IFDEF AMORTISSEMENT}
AMLanceFiche_ListeDesImmobilisations ( E_GENERAL.Text, False , taConsult) ;
{$ENDIF}
end;

{Cumuls du compte général}
procedure TFConsEcr.CumulsGENEClick(Sender: TObject);
Var LeExo : TExoDate ;
begin
  inherited;
LeExo:=PositionneExo ;
CumulCpteMensuel(fbGene,E_GENERAL.Text,TOBGene.GetValue('G_LIBELLE'),LeExo) ;
end;

{Grand-livre général simple}
procedure TFConsEcr.GLGENEClick(Sender: TObject);
Var
    ACritEdt : ClassCritEdt;
    D1,D2 : TdateTime;
    Etab : string;
begin
  inherited;
  ACritEdt := ClassCritEdt.Create;
  try
    Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
    D1:=StrToDate(E_DATECOMPTABLE.text) ; D2:=StrToDate(E_DATECOMPTABLE_.text) ;
    ACritEdt.CritEdt.Date1:=D1 ;
    ACritEdt.CritEdt.DateDeb:=ACritEdt.CritEdt.Date1 ;
    ACritEdt.CritEdt.Date2:=D2 ;
    ACritEdt.CritEdt.DateDeb:=ACritEdt.CritEdt.Date2 ;
    ACritEdt.CritEdt.Exo:=PositionneExo ;
    ACritEdt.CritEdt.Cpt1  := E_General.Text;
    ACritEdt.CritEdt.Cpt2  := E_General.Text;
    ACritEdt.CritEdt.SCpt1 := '';
    ACritEdt.CritEdt.SCpt2 := '';
    Etab:=EtabForce ; if Etab<>'' then ACritEdt.CritEdt.Etab:=Etab ;
    TheData := ACritEdt;
    CPLanceFiche_CPGLGENE();
  finally
    ACritEdt.Free;
    TheData := nil;
  end;
end;

{Justif de solde TID / TIC}
procedure TFConsEcr.JustifGENEClick(Sender: TObject);
(*
Var Crit : TCritEdt ;
    OldD1,D1{,D2} : TDateTime ;
begin
  inherited;
Fillchar(Crit,SizeOf(Crit),#0) ;
Crit.Exo:=PositionneExo ;
D1:=StrToDate(E_DATECOMPTABLE_.text) ;
Crit.Date1:=D1 ;
Crit.DateDeb:=Crit.Date1 ;
Crit.NatureEtat:=neJU ;
InitCritEdt(Crit) ;
Crit.JU.OnTiers:=False ;
Crit.Cpt1:=E_GENERAL.text ; Crit.Cpt2:=Crit.Cpt1 ;
Crit.JU.TriePar:=0 ;
Crit.JU.Lettrage:=0 ;
Crit.JU.EnSituation:=False ;
OldD1:=StrToDate(E_DATECOMPTABLE.text) ; E_DATECOMPTABLE.Text:=DateToStr(iDate1900) ;
Crit.SQLPLUS:=' AND '+Q.CRITERES+' ' ;
E_DATECOMPTABLE.Text:=DateToStr(OldD1) ;
JustSoldeZoom(Crit) ;
*)
begin
CPLanceFiche_JustSold;
end;

{Cumuls auxiliaire}
procedure TFConsEcr.CumulsAUXClick(Sender: TObject);
Var LeExo : TExoDate ;
begin
  inherited;
LeExo:=PositionneExo ;
CumulCpteMensuel(fbAux,E_AUXILIAIRE.Text,TOBTiers.GetValue('T_LIBELLE'),LeExo) ;
end;

{Grand-livre auxiliaire simple}
procedure TFConsEcr.GLAUXClick(Sender: TObject);
var
    D1,D2 : TDateTime ;
    ACritEdt : ClassCritEdt;
    Etab : string;
begin
  inherited;
  ACritEdt := ClassCritEdt.Create;
  try
    Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
    D1:=StrToDate(E_DATECOMPTABLE.text) ; D2:=StrToDate(E_DATECOMPTABLE_.text) ;
    ACritEdt.CritEdt.Date1:=D1 ;
    ACritEdt.CritEdt.DateDeb:=ACritEdt.CritEdt.Date1 ;
    ACritEdt.CritEdt.Date2:=D2 ;
    ACritEdt.CritEdt.DateFin:=ACritEdt.CritEdt.Date2 ;
    ACritEdt.CritEdt.Exo:=PositionneExo ;
    ACritEdt.CritEdt.Cpt1  := E_AUXILIAIRE.text;
    ACritEdt.CritEdt.Cpt2  := E_AUXILIAIRE.text;
    ACritEdt.CritEdt.SCpt1 := '';
    ACritEdt.CritEdt.SCpt2 := '';
    Etab:=EtabForce ; if Etab<>'' then ACritEdt.CritEdt.Etab:=Etab ;
    TheData := ACritEdt;
    CPLanceFiche_CPGLAUXI('');
  finally
    ACritEdt.Free;
    TheData := nil;
  end;
end;

{Grand-livre auxiliaire en situation}
procedure TFConsEcr.GLAUXSITUClick(Sender: TObject);
Var
    D1,D2 : TDateTime ;
    ACritEdt : ClassCritEdt;
    Etab : string;
begin
  inherited;
  ACritEdt := ClassCritEdt.Create;
  try
    Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
    D1:=StrToDate(E_DATECOMPTABLE.text) ; D2:=StrToDate(E_DATECOMPTABLE_.text) ;
    ACritEdt.CritEdt.Date1:=D1 ;
    ACritEdt.CritEdt.DateDeb:=ACritEdt.CritEdt.Date1 ;
    ACritEdt.CritEdt.Date2:=D2 ;
    ACritEdt.CritEdt.DateFin:=ACritEdt.CritEdt.Date2 ;
    ACritEdt.CritEdt.Cpt1  := E_AUXILIAIRE.text;
    ACritEdt.CritEdt.Cpt2  := E_AUXILIAIRE.text;
    ACritEdt.CritEdt.SCpt1 := '';
    ACritEdt.CritEdt.SCpt2 := '';
    ACritEdt.CritEdt.GL.EnDateSituation := True;
    Etab:=EtabForce ; if Etab<>'' then ACritEdt.CritEdt.Etab:=Etab ;
    TheData := ACritEdt;
    CPLanceFiche_CPGLAUXI('');
  finally
    ACritEdt.Free;
    TheData := nil;
  end;
end;

{Justif de solde auxiliaire}
procedure TFConsEcr.JUSTIFAUXClick(Sender: TObject);
(*
Var Crit : TCritEdt ;
    OldD1,D1{,D2} : TDateTime ;
begin
  inherited;
Fillchar(Crit,SizeOf(Crit),#0) ;
Crit.Exo:=PositionneExo ;
D1:=StrToDate(E_DATECOMPTABLE_.text) ;
Crit.Date1:=D1 ;
Crit.DateDeb:=Crit.Date1 ;
Crit.NatureEtat:=neJU ;
InitCritEdt(Crit) ;
Crit.JU.OnTiers:=TRUE ;
Crit.Cpt1:=E_AUXILIAIRE.text ; Crit.Cpt2:=Crit.Cpt1 ;
Crit.SCpt1:=E_GENERAL.text ; Crit.SCpt2:=Crit.SCpt1 ;
Crit.JU.TriePar:=0 ;
Crit.JU.Lettrage:=0 ;
Crit.JU.EnSituation:=False ;
OldD1:=StrToDate(E_DATECOMPTABLE.text) ; E_DATECOMPTABLE.Text:=DateToStr(iDate1900) ;
Crit.SQLPLUS:=' AND '+Q.CRITERES+' ' ;
E_DATECOMPTABLE.Text:=DateToStr(OldD1) ;
JustSoldeZoom(Crit) ;
*)
begin
CPLanceFiche_JustSold;
end;

Function TFConsEcr.OkLettre : boolean ;
Var FF : TField ;
BEGIN
Result:=False ;
if ((Q.EOF) and (Q.BOF)) then Exit ;
FF:=Q.FindField('E_LETTRAGE') ; if FF=Nil then Exit ;
if FF.AsString='' then Exit ;
Result:=True ;
END ;

Function TFConsEcr.OkFicheImmo : boolean ;
Var FF : TField ;
BEGIN
Result:=False ;
if ((Q.EOF) and (Q.BOF)) then Exit ;
FF:=Q.FindField('E_IMMO') ; if FF=Nil then Exit ;
if FF.AsString='' then Exit ;
Result:=True ;
END ;

procedure TFConsEcr.EnabledMenus ;
BEGIN
// Lettrages
LettrageM.Enabled:=False ;
if ((TOBTiers.GetValue('T_AUXILIAIRE')<>'') and (TOBGene.GetValue('G_GENERAL')<>'')) then LettrageM.Enabled:=True else
   BEGIN
   if ((TOBGene.GetValue('G_NATUREGENE')='TID') or (TOBGene.GetValue('G_NATUREGENE')='TIC')) and
       (TOBGene.GetValue('G_LETTRABLE')='X') then LettrageM.Enabled:=True ;
   END ;
if ((FromSaisie) or (CAvecSimu.Checked)) then LettrageM.Enabled:=False ;
LettrageA.Enabled:=LettrageM.Enabled ;
DelettreAuto.Enabled:=LettrageM.Enabled ;
Delettre.Enabled:=((LettrageM.Enabled) and (OkLettre)) ;
// Modifs fiches
ModifAux.Enabled:=TOBTiers.GetValue('T_AUXILIAIRE')<>'' ;
ModifGene.Enabled:=TOBGene.GetValue('G_GENERAL')<>'' ;
{$IFDEF SANSIMMO}
Immos.Enabled:=FALSE ;
{$ELSE}
Immos.Enabled:=(TOBGene.GetValue('G_NATUREGENE')='IMO') and ((VH^.OkModImmo) or (V_PGI.VersionDemo)) ;
{$ENDIF}
FicheImmo.Enabled:=((Immos.Enabled) and (OkFicheImmo)) ;
// Ecritures
ModifEcr.Enabled:=Not ((Q.EOF) and (Q.BOF)) ;
InfosComp.Enabled:=Not ((Q.EOF) and (Q.BOF)) ;
analrevis.Enabled:=Not ((Q.EOF) and (Q.BOF)) ;
Analytiques.Enabled:=(TOBGene.GetValue('G_VENTILABLE')='X') ;
SaisieEcr.Enabled:=((DroitEcritures) and (ToutAcc)) ;
SaisieBOR.Enabled:=((DroitEcritures) and (ToutAcc)) ;
// Pointage
PointageG.Enabled:=(TOBGene.GetValue('G_POINTABLE')='X') ;
DePointageG.Enabled:=PointageG.Enabled ;
//Etats
JustifGENE.Enabled:=(TOBGene.GetValue('G_LETTRABLE')='X') ;
CumulsAUX.Enabled:=(TOBTiers.GetValue('T_AUXILIAIRE')<>'') ;
GLAUX.Enabled:=CumulsAUX.Enabled ;
GLAUXSITU.Enabled:=CumulsAUX.Enabled ;
JustifAUX.Enabled:=CumulsAUX.Enabled ;
JustifAUXSITU.Enabled:=CumulsAUX.Enabled ;

if CtxPcl in V_PGI.PGICOntexte then
begin
  IccDetail.Enabled := (VH^.OkModIcc) and
                       Presence('generaux', 'g_general', E_GENERAL.Text) and
                       Presence('iccgeneraux', 'icg_general', E_GENERAL.Text);

  IccCalcul.Enabled := (VH^.OkModIcc) and IccDetail.Enabled;
  EmpruntCRE.Enabled := (VH^.OkModCRE) and (Copy(E_GENERAL.Text,1,2)='16')
    and (ExisteSQL ('SELECT EMP_NUMCOMPTE FROM FEMPRUNT WHERE EMP_NUMCOMPTE="'+E_GENERAL.Text+'"'));
end
else
begin
  IccDetail.Enabled := Presence('generaux', 'g_general', E_GENERAL.Text) and
                       Presence('iccgeneraux', 'icg_general', E_GENERAL.Text);

  IccCalcul.Enabled := IccDetail.Enabled;
  EmpruntCRE.Enabled := False;
end;

END ;

procedure TFConsEcr.PopACPTPopup(Sender: TObject);
begin
  inherited;
EnabledMenus ;
end;

procedure TFConsEcr.PopAEdtPopup(Sender: TObject);
begin
  inherited;
EnabledMenus ;
end;

procedure TFConsEcr.PopZPopup(Sender: TObject);
Var i : integer ;
    TS,TD : TMenuItem ;
    OldCap : String ;
begin
  inherited;
PurgePopUp(PopZ) ; EnabledMenus ; OldCap:='-' ;
for i:=0 to PopACpt.Items.Count-1 do
    BEGIN
    TS:=TMenuItem(PopACpt.Items[i]) ;
    if TS.Enabled then
       BEGIN
       if ((TS.Caption='-') and (OldCap='-')) then Continue ;
       TD:=TMenuItem.Create(PopZ) ; OldCap:=TS.Caption ;
       TD.Caption:=TS.Caption ; TD.OnClick:=TS.OnClick ;
       PopZ.Items.Add(TD) ;
       END ;
    END ;
for i:=0 to PopAEdt.Items.Count-1 do
    BEGIN
    TS:=TMenuItem(PopAEdt.Items[i]) ;
    if TS.Enabled then
       BEGIN
       if ((TS.Caption='-') and (OldCap='-')) then Continue ;
       TD:=TMenuItem.Create(PopZ) ;
       TD.Caption:=TS.Caption ; TD.OnClick:=TS.OnClick ;
       PopZ.Items.Add(TD) ;
       END ;
    END ;
if POPZ.Items.Count>0 then
   BEGIN
   TD:=POPZ.Items[POPZ.Items.Count-1] ;
   if TD.Caption='-' then BEGIN POPZ.Items.Remove(TD) ; TD.Free ; {TD:=Nil ;} END ;
   END ;
end;

procedure TFConsEcr.GeneSuivant ( Suiv : boolean ) ;
Var Q : TQuery ;
    SQL,CptGene,NewCpt,StC : String ;
BEGIN
CptGene:=E_GENERAL.Text ; NewCpt:=CptGene ;
if Suiv then SQL:='SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL>"'+CptGene+'"'
        else SQL:='SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL<"'+CptGene+'"' ;
Case RDefil.ItemIndex of
   0 : SQL:=SQL+'' ;
   1 : SQL:=SQL+' AND G_TOTALDEBIT-G_TOTALCREDIT<>0 ' ;
   2 : SQL:=SQL+' AND (G_TOTALDEBIT<>0 OR G_TOTALCREDIT<>0) ' ;
   END ;

if Suiv then SQL:=SQL+' ORDER BY G_GENERAL ' else SQL:=SQL+' ORDER BY G_GENERAL DESC ' ;
Q:=OpenSQL(SQL,True) ;
if Not Q.EOF then NewCpt:=Q.Fields[0].AsString ;
Ferme(Q) ;
if NewCpt<>CptGene then
   BEGIN
    if  EstSQLConfidentiel('TIERS', E_AUXILIAIRE.Text) or EstSQLCOnfidentiel('GENERAUX',NewCpt)  then
    begin
      Caption:='Le compte ' + NewCpt + traduirememoire(' est confidentiel')  ;
      //Caption:=traduirememoire(' Accés confidentiel !!!')  ;
      UpdateCaption(Self) ;
      E_GENERAL.Text:=NewCpt ;
      OldGene:=E_GENERAL.Text ;
    end
    else
    begin
      E_GENERAL.Text:=NewCpt ; ChargeGeneral(True) ; if not Optim then BChercheClick(Nil) ;
      OldGene:=E_GENERAL.Text ;
    end ;
   END ;
//VL
if Optim then begin
	StC:=Copy(Caption,1,Pos(':',Caption)) ;
  if E_AUXILIAIRE.Text<>'' then StC:=StC+'  '+E_AUXILIAIRE.Text+'  '+TTLIBELLE.Caption else
  if E_GENERAL.Text<>'' then StC:=StC+' '+E_GENERAL.Text+'  '+TGLIBELLE.Caption;
  Caption:=StC ; UpdateCaption(Self) ;
  LESOLDE.Visible:=FALSE ; TPASFAIT.Visible:=TRUE ;
end;
//VL
END ;

Function TFConsEcr.RecupWhereDataType : String ;
Var St : String ;
    ii  : integer ;
    tz  : TZoomTable ;
BEGIN
Result:='' ; St:='' ;
tz:=E_AUXILIAIRE.ZoomTable ;
Case tz of
  tzTToutDebit  : St:='TZTTOUTDEBIT' ;
  tzTToutCredit : St:='TZTTOUTCREDIT' ;
  tzTSalarie    : St:='TZTSALARIE' ;
  tzTiers       : St:='TZTTOUS' ;
  else St:='TZTTOUS' ;
  END ;
ii:=TTToNum(St) ;
if ii>0 then
begin
  Result:=V_PGI.DECombos[ii].Where ;
  //Si on récupère un &#@ on retourne rien
  if Result='&#@' then Result := '' ;
end ;
END ;

procedure TFConsEcr.AuxSuivant ( Suiv : boolean ) ;
Var Q : TQuery ;
    SQL,CptAux,NewCpt,sWhere,StC : String ;
BEGIN
CptAux:=E_AUXILIAIRE.Text ; NewCpt:=CptAux ;
if Suiv then SQL:='SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE>"'+CptAux+'" '
        else SQL:='SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE<"'+CptAux+'" ' ;
sWhere:=RecupWhereDataType ;
if sWhere<>'' then SQL:=SQL+' AND '+sWhere ;
Case RDefil.ItemIndex of
   0 : SQL:=SQL+'' ;
   1 : SQL:=SQL+' AND T_TOTALDEBIT-T_TOTALCREDIT<>0 ' ;
   2 : SQL:=SQL+' AND (T_TOTALDEBIT<>0 OR T_TOTALCREDIT<>0) ' ;
   END ;

if Suiv then SQL:=SQL+' ORDER BY T_AUXILIAIRE ' else SQL:=SQL+' ORDER BY T_AUXILIAIRE DESC ' ;
Q:=OpenSQL(SQL,True) ;
if Not Q.EOF then NewCpt:=Q.Fields[0].AsString ;
Ferme(Q) ;
if NewCpt<>CptAux then
   BEGIN
    if  EstSQLConfidentiel('TIERS', NewCpt) or EstSQLCOnfidentiel('GENERAUX',E_GENERAL.Text)  then
    begin
      Caption:='Le compte ' + NewCpt + traduirememoire(' est confidentiel')  ;
      //Caption:=traduirememoire(' Accés confidentiel !!!')  ;
      UpdateCaption(Self) ;
      E_AUXILIAIRE.Text:=NewCpt ;
      OldGene:=E_AUXILIAIRE.Text ;
    end
    else
    begin
      E_AUXILIAIRE.Text:=NewCpt ; ChargeAuxiliaire(True) ; if Not Optim then BChercheClick(Nil) ;
      OldAuxi:=E_AUXILIAIRE.Text ;
    end;
   END ;
//VL
if Optim then begin
	StC:=Copy(Caption,1,Pos(':',Caption)) ;
  if E_AUXILIAIRE.Text<>'' then StC:=StC+'  '+E_AUXILIAIRE.Text+'  '+TTLIBELLE.Caption else
  if E_GENERAL.Text<>'' then StC:=StC+' '+E_GENERAL.Text+'  '+TGLIBELLE.Caption;
  Caption:=StC ; UpdateCaption(Self) ;
  LESOLDE.Visible:=FALSE ; TPASFAIT.Visible:=TRUE ;
end;
//VL
END ;

procedure TFConsEcr.RActionClick(Sender: TObject);
begin
  inherited;
ToutAcc:=(RAction.ItemIndex=1) ;
end;

procedure TFConsEcr.ROpposeClick(Sender: TObject);
begin
  inherited;
If Not Optim Then BChercheClick(Nil) ;
end;

procedure TFConsEcr.BNouvRechClick(Sender: TObject);
Var Aux,Gene : String ;
begin
Gene:=E_GENERAL.Text ; Aux:=E_AUXILIAIRE.Text ;
  inherited;
E_GENERAL.Text:=Gene ; E_AUXILIAIRE.Text:=Aux ;
EtatLettrage.ItemIndex:=0 ; EtatLettrageChange(Nil) ; 
end;

procedure TFConsEcr.EtatLettrageChange(Sender: TObject);

begin
  inherited;
Case Etatlettrage.ItemIndex of
   0 : XX_WHERELET.Text:='' ;
   1 : XX_WHERELET.Text:='E_ETATLETTRAGE="PL" OR E_ETATLETTRAGE="TL"' ;
   2 : XX_WHERELET.Text:='E_ETATLETTRAGE<>"PL" AND E_ETATLETTRAGE<>"TL"' ;
   3 :
if VH^.CPIFDEFCEGID then
Begin
   XX_WHERELET.Text:='(E_ETATLETTRAGE="PL" OR E_ETATLETTRAGE="AL")' ;
end else begin
   XX_WHERELET.Text:='E_ETATLETTRAGE="PL"' ;
end ;
   4 : XX_WHERELET.Text:='E_ETATLETTRAGE="TL"' ;
   END ;
end;

procedure TFConsEcr.AuxiUpClick(Sender: TObject);
begin
  inherited;
AuxSuivant(False) ;
end;

procedure TFConsEcr.AuxiDownClick(Sender: TObject);
begin
  inherited;
AuxSuivant(True) ;
end;


procedure TFConsEcr.GeneUpClick(Sender: TObject);
begin
  inherited;
GeneSuivant(False) ;
end;

procedure TFConsEcr.GeneDownClick(Sender: TObject);
begin
  inherited;
GeneSuivant(True) ;
end;

procedure TFConsEcr.FicheImmoClick(Sender: TObject);
{$IFDEF AMORTISSEMENT}
Var FF : TField ;
    NumI : String ;
{$ENDIF}
begin
  inherited;
{$IFDEF AMORTISSEMENT}
if ((Q.EOF) and (Q.BOF)) then Exit ;
FF:=Q.FindField('E_IMMO') ; if FF=Nil then Exit ;
NumI:=FF.AsString ; if NumI='' then Exit ;
AMLanceFiche_FicheImmobilisation(NumI,taConsult,'');
{$ENDIF}
end;

procedure TFConsEcr.FormResize(Sender: TObject);
begin
  inherited;
if ((Not IsInside(Self)) and (HMTrad.ActiveResize) and (HMTrad.ResizeDBGrid)) then
   BEGIN
   HMTrad.ResizeDBGridColumns(FListe) ;
   FListe.Refresh ;
   END ;
end;

procedure TFConsEcr.CAvecSimuClick(Sender: TObject);
begin
  inherited;
XX_WHEREQUALIF.Text:='E_QUALIFPIECE="N"' ;
if cAvecSimu.Checked then XX_WHEREQUALIF.Text:='E_QUALIFPIECE="N" OR E_QUALIFPIECE="S"' ;
end;

procedure TFConsEcr.ANALREVISClick(Sender: TObject);
Var WhereSQL : String ;
begin
  inherited;
WhereSQL:=RecupWhereCritere(Pages) ;
AGLLanceFiche('CP','CPCONSULTREVIS','','',WhereSQL) ;
end;

procedure TFConsEcr.JUSTIFAUXSITUClick(Sender: TObject);
(*
Var Crit : TCritEdt ;
    OldD1,D1{,D2} : TDateTime ;
begin
  inherited;
Fillchar(Crit,SizeOf(Crit),#0) ;
Crit.Exo:=PositionneExo ;
D1:=StrToDate(E_DATECOMPTABLE_.text) ;
Crit.Date1:=D1 ;
Crit.DateDeb:=Crit.Date1 ;
Crit.NatureEtat:=neJU ;
InitCritEdt(Crit) ;
Crit.JU.OnTiers:=TRUE ;
Crit.Cpt1:=E_AUXILIAIRE.text ; Crit.Cpt2:=Crit.Cpt1 ;
Crit.SCpt1:=E_GENERAL.text ; Crit.SCpt2:=Crit.SCpt1 ;
Crit.JU.TriePar:=0 ;
Crit.JU.Lettrage:=0 ;
Crit.JU.EnSituation:=TRUE ;
OldD1:=StrToDate(E_DATECOMPTABLE.text) ; E_DATECOMPTABLE.Text:=DateToStr(iDate1900) ;
Crit.SQLPLUS:=' AND '+Q.CRITERES+' ' ;
E_DATECOMPTABLE.Text:=DateToStr(OldD1) ;
JustSoldeZoom(Crit) ;
*)
begin
CPLanceFiche_JustSold;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2001
Modifié le ... :   /  /
Description .. : Affiche le détail du compte ICC
Mots clefs ... : ICC
*****************************************************************}
procedure TFConsEcr.ICCDETAILClick(Sender: TObject);
begin
  inherited;
  AGLLanceFiche('CP', 'ICCFICHEGENERAUX', '', E_GENERAL.Text, 'ACTION=MODIFICATION;' +
                                                              E_DATECOMPTABLE.Text + ';' +
                                                              E_DATECOMPTABLE_.Text);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2001
Modifié le ... :   /  /    
Description .. : Appel la fenetre de calcul des intérêts ICC
Mots clefs ... : ICC
*****************************************************************}
procedure TFConsEcr.ICCCALCULClick(Sender: TObject);
begin
  inherited;
  AGLLanceFiche('CP', 'ICCPARAMETRE', '', '', E_DATECOMPTABLE.Text  + ';' +
                                              E_DATECOMPTABLE_.Text + ';' +
                                              E_GENERAL.Text);
end;

procedure TFConsEcr.E_GENERALKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  if Key = VK_F5 then E_GENERAL.OnDblClick( nil );
end;

procedure TFConsEcr.EmpruntCREClick(Sender: TObject);
begin
  inherited;
  AGLLanceFiche('FP', 'FMULEMPRUNT', '', '', E_GENERAL.Text);
end;

{---------------------------------------------------------------------------------------}
procedure TFConsEcr.GereEtablissement;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(E_ETABLISSEMENT) then begin
    {Si l'on ne gère pas les établissement ...}
    if not VH^.EtablisCpta  then begin
      {... on affiche l'établissement par défaut}
      E_ETABLISSEMENT.Value := VH^.EtablisDefaut;
      {... on désactive la zone}
      E_ETABLISSEMENT.Enabled := False;
    end

    {On gère l'établisement, donc ...}
    else begin
      {... On commence par regarder les restrictions utilisateur}
      PositionneEtabUser(E_ETABLISSEMENT);
      {... s'il n'y a pas de restrictions, on reprend le paramSoc
       JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
      if E_ETABLISSEMENT.Value = '' then begin
        {... on affiche l'établissement par défaut
        E_ETABLISSEMENT.Value := VH^.EtablisDefaut;
        {... on active la zone
        E_ETABLISSEMENT.Enabled := True;
      end;}
    end;
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TFConsEcr.ControlEtab;
{---------------------------------------------------------------------------------------}
var
  Eta : string;
begin
  if not Assigned(E_ETABLISSEMENT) then Exit;
  {S'il n'y a pas de gestion des établissement, logiquement, on ne force pas l'établissement !!!}
  if not VH^.EtablisCpta then Exit;
  
  Eta := EtabForce;
  {S'il y a une restriction utilisateur et qu'elle ne correspond pas au contenu de la combo ...}
  if (Eta <> '') and (Eta <> E_ETABLISSEMENT.Value) then begin
    {... on affiche l'établissement des restrictions}
    E_ETABLISSEMENT.Value := Eta;
    {... on désactive la zone}
    E_ETABLISSEMENT.Enabled := False;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TFConsEcr.AfterSelectFiltre;
{---------------------------------------------------------------------------------------}
begin
  ControlEtab;
end;

end.

