unit wMnu;

interface

uses
  Windows,
  Classes,
  SysUtils,
  HCtrls,
  Hent1,
  uTob,
  Dialogs,
  Forms,
  Controls,
  comCtrls,
  stdCtrls,
  Vierge,
  AglIsoflex,
  {$IFNDEF EAGLCLIENT}
    db,
    {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
    Fe_Main,                                                           
    Fiche,
    FichList,
    Mul,
    AdoDB,
  {$ELSE}
    MainEagl,
    eFiche,
    eFichList,
    eMul,
  {$ENDIF}
  HMsgBox,
  UTom,
  Hdb,
  M3FP,
  Graphics,
  Menus,
  wCommuns,
  HTB97
  ;

procedure wSetMnuLoupe     (Const MultiSelected: boolean; Const Argument: string; Const Pm: TPopUpMenu);
procedure wSetMnuHisto     (Const MultiSelected: boolean; Const Argument: string; Const Pm: TPopUpMenu);
procedure wSetMnuAction    (Const MultiSelected: Boolean; Const Argument: string; Const Pm: TPopUpMenu);
procedure wSetMnuUtilitaire(Const MultiSelected: Boolean; Const Argument: string; Const Pm: TPopUpMenu);
procedure wSetMnuAnalyse   (Const MultiSelected: Boolean; Const Argument: string; Const Pm: TPopUpMenu);
procedure wSetMnuBaseTech  (Const MultiSelected: Boolean; Const Argument: string; Const Pm: TPopUpMenu);
{$IFDEF GCGC}
procedure wSetItemsMnuLoupeArticle(Const PmLoupe: TPopupMenu; Const LpNAT_OnClick: TNotifyEvent; Const Article: String; Const Sender: TForm);
{$ENDIF GCGC}
function  wIsEmptyPopup(Popup: TPopUpMenu): Boolean;

{$IFDEF STK}
type
  TFluxCanShowAchat = (sfaNone, sfaYes, sfaNo);

  TPopupMenuFlux = Class(TPopupMenu)
  private
    FLastPmFlux: TPopupMenuFlux;
    FFluxCanShowAchat: TFluxCanShowAchat;
    FEcran: TComponent;
    FFromPiece: Boolean;
    function  GetFlux: String;
    procedure SetFlux(Flux: String);
    procedure OnClick(Sender: TObject);
    function  GetCanShowAchat: Boolean;
    procedure SetFromPiece(const Value: Boolean);
  public
    constructor Create(aEcran: TComponent; Mn_OnClick: TNotifyEvent; Argument: String; Flux: String = ''); reintroduce;
    destructor Destroy; override;
    property Flux: String Read GetFlux Write SetFlux;
    property Ecran: TComponent Read FEcran Write FEcran;
    property FromPiece: Boolean Read FFromPiece write SetFromPiece;
    property CanShowAchat: Boolean Read GetCanShowAchat;
    function GetFluxFromSender(Sender: TObject):String;
  end;
{$ENDIF STK}


implementation

Uses
  {$If defined (GPAOLIGHT) or defined (CRM)}
    wAction,
  {$Ifend}
  {$IFDEF GPAOLIGHT}
    EntGP,
    wNomeLig,
  {$ENDIF GPAOLIGHT}
  {$IFDEF GPAO}
    wConfigurator,
  {$ENDIF GPAO}
  {$IFDEF STK}
    StockUtil,
    EmplacementUtil,
    StkNature,
    Dispo,
    Ent1,
  {$ENDIF STK}
  {$IFDEF EDI}
    EdiLigne,
  {$ENDIF EDI}
  EntGC,
  UtilPgi,
  {$IFDEF GCGC}
  UtilArticle,
  FactUtil,
  {$ENDIF GCGC}
  ParamSoc
  ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 22/05/2002
Modifié le ... : 13/10/2002: Retire les images
Description .. : Gestion du menu loupe
Mots clefs ... :
*****************************************************************}
procedure wSetMnuLoupe(Const MultiSelected: Boolean; Const Argument: string; Const Pm: TPopUpMenu);
var
  i: Integer;
  {$IFDEF QUALITE}
    DemDerogNum, PlanCorrNum, IdAction, LigneOrdre: integer;
    lRQN: boolean;
    ClePiece, Groupe, Critere: String;
  {$ENDIF QUALITE}
  {$IFDEF GPAO}
    OrdrePere : Integer;
    Coditi, Circuit: string;
    RubriquePDR, SectionPDR: string;
  {$ENDIF GPAO}
  EtatPhase, EtatOpe,
  LotInterne, SerieInterne,
  TypeOrdre, Article, TypeArticle, Depot : String;
  lCodePort, lTarifSpecial,
  lTiers, lTarifTiers, lTiersOrigine, lTiersActivite, lTiersEnseigne, lTiersGroupe,
  {$IFDEF AFFAIRE}
  lAffaire,
  {$ENDIF AFFAIRE}
  lArticle, lTarifArticle,
  lCatalogue,
  lFamilleValo, lFamilleNiv1,
  lFamilleNiv2, lFamilleNiv3,
  lNatureTra, lRessource,
  lTarifResso, lSecteurGeo,
  lCodePostal, lPays, lRegion,
  lModeExp, lWithPiece,
{$IFDEF STK}

  {$IFDEF GPAO}
  lMarque, lChoixQualite,
  {$ENDIF GPAO}
  lIndiceArticle,
{$ENDIF STK}
  lFormuleMt : Boolean;

  function wSetMn(Const MnLpItem: TMenuItem): boolean;
  var
    j: Integer;
  begin
    Result := true;
    if Assigned(MnLpItem) then
    begin
      with MnLpItem do
      begin
        {--- Commun à toutes les options ---}
        if (UpperCase(Name) = UpperCase('mnSGED'))or (UpperCase(Name) = UpperCase('mnSGEDL'))then
        begin
          Visible := AglIsoflexPresent;
        end
        else if (UpperCase(Name) = UpperCase('MnLpPieceFromWOT')) or (UpperCase(Name) = UpperCase('MnLpLigneFromWOL')) then
        begin
          Visible := not MultiSelected and lWithPiece;
        end
        else if (UpperCase(Name) = UpperCase('MnLpCodePort')) then
        begin
          Visible := (not MultiSelected) and (lCodePort) and (JAiLeDroitTag(65401) or JAiLeDroitTag(654011)); ;
        end
        else if (UpperCase(Name) = UpperCase('MnLpNatureTra')) then
        begin
          Visible := (not MultiSelected) and (lNatureTra) and (JAiLeDroitTag(126910));
        end
        else if (UpperCase(Name) = UpperCase('MnLpArticle')) then
        begin
          Visible := (not MultiSelected) and (lArticle) and (JAiLeDroitGestion('GA') or JAiLeDroitConsult('GA') );
        end
        else if (UpperCase(Name) = UpperCase('MnLpCatalogu')) then
        begin
          Visible := (not MultiSelected) and (lCatalogue) and (JAiLeDroitGestion('GCA') or JAiLeDroitConsult('GCA') );
        end
        else if (UpperCase(Name) = UpperCase('LpIndiceArticle')) then
        begin
          {$IFDEF STK}
            {$IFDEF GPAO}
              //GP_DS_GP14395_20071009
              //GP_JLS_GP14648_20071212
              Visible := ((Pos(TypeArticle,'MAR;CNS') > 0) or (TypeArticle = '')) and StkGereIndiceArticle and (not MultiSelected);
            {$ELSE GPAO}
              Visible := false;
            {$ENDIF GPAO}
          {$ELSE STK}
            Visible := false;
          {$ENDIF STK}
        end
        else if (UpperCase(Name) = UpperCase('MnLpFamilleTarifArticle')) then
        begin
          Visible := (not MultiSelected) and (lTarifArticle) and (JAiLeDroitGestion('TAR') or JAiLeDroitConsult('TAR'));
        end
        else if (UpperCase(Name) = UpperCase('MnLpFamilleValo')) then
        begin
          Visible := (not MultiSelected) and (lFamilleValo) and (JAiLeDroitGestion('FVS') or JAiLeDroitConsult('FVS'));
        end
        else if (UpperCase(Name) = UpperCase('MnLpFamilleNiveau1')) then
        begin
          Caption := RechDom('GCLIBFAMILLE', 'LF1', False);
          Visible := (not MultiSelected) and (lFamilleNiv1) and (JAiLeDroitGestion('FN1') or JAiLeDroitConsult('FN1'));
        end
        else if (UpperCase(Name) = UpperCase('MnLpFamilleNiveau2')) then
        begin
          Caption := RechDom('GCLIBFAMILLE', 'LF2', False);
          Visible := (not MultiSelected) and (lFamilleNiv2) and (JAiLeDroitGestion('FN2') or JAiLeDroitConsult('FN2'));
        end
        else if (UpperCase(Name) = UpperCase('MnLpFamilleNiveau3')) then
        begin
          Caption := RechDom('GCLIBFAMILLE', 'LF3', False);
          Visible := (not MultiSelected) and (lFamilleNiv3) and (JAiLeDroitGestion('FN3') or JAiLeDroitConsult('FN3'));
        end
        {$IFDEF STK}
          {$IFDEF GPAO}
          else if (UpperCase(Name) = UpperCase('MnLpMarque')) then
          begin
            Visible := (not MultiSelected) and (lMarque) and StkGereMarque and (JAiLeDroitGestion('GM') or JAiLeDroitConsult('GM')) ;
          end
          else if (UpperCase(Name) = UpperCase('MnLpChoixQualite')) then
          begin
            Visible := (not MultiSelected) and (lChoixQualite) and StkGereChoixQualite and (JAiLeDroitGestion('GCQ') or JAiLeDroitConsult('GCQ')) ;
          end
          {$ENDIF GPAO}
          else if (UpperCase(Name) = UpperCase('MnLpIndiceArticle')) then
          begin
            //GP_DS_GP14395_20071009
            //GP_JLS_GP14648_20071212
            Visible := ((Pos(TypeArticle,'MAR;CNS') > 0) or (TypeArticle = '')) and (not MultiSelected) and (lIndiceArticle) and StkGereIndiceArticle;
          end
        {$ENDIF STK}
        else if (UpperCase(Name) = UpperCase('MnLpRessource')) then
        begin
          Visible := (not MultiSelected) and (lRessource) and (JAiLeDroitTag(126630));
        end
        else if (UpperCase(Name) = UpperCase('MnLpTarifRessource')) then
        begin
          Visible := (not MultiSelected) and (lTarifResso) and (JAiLeDroitTag(65652));
        end
        else if (UpperCase(Name) = UpperCase('MnLpTarifsSpeciaux')) then
        begin
          Visible := (not MultiSelected) and (lTarifSpecial) and (JAiLeDroitGestion('YTP') or JAiLeDroitConsult('YTP')) ;
        end
        //GP_DS_AFFAIRE_20070925
        else if (UpperCase(Name) = UpperCase('MnLpAffaire')) then
        begin
          {$IFDEF AFFAIRE}
            Visible := (not MultiSelected) and (lAffaire);
          {$ELSE AFFAIRE}
            Visible := False;
          {$ENDIF AFFAIRE}
        end
        else if (UpperCase(Name) = 'MNLPTIERS') or (UpperCase(Name) = 'MNPROSPECT') or (UpperCase(Name) = 'MNCONTACTS') then
        begin
          Visible := (not MultiSelected) and (lTiers) and (JAiLeDroitGestion('T') or JAiLeDroitConsult('T')) ;
        end
        else if (UpperCase(Name) = UpperCase('MnLpFamilleTarifTiers')) then
        begin
          Visible := (not MultiSelected) and (lTarifTiers) and (JAiLeDroitGestion('TRC') or JAiLeDroitConsult('TRC')) ;
        end
        else if (UpperCase(Name) = UpperCase('MnLpTiersOrigine')) then
        begin
          Visible := (not MultiSelected) and (lTiersOrigine) and (JAiLeDroitGestion('GOR') or JAiLeDroitConsult('GOR')) ;
        end
        else if (UpperCase(Name) = UpperCase('MnLpTiersActivite')) then
        begin
          Visible := (not MultiSelected) and (lTiersActivite) and (JAiLeDroitGestion('SCC') or JAiLeDroitConsult('SCC')) ;
        end
        else if (UpperCase(Name) = UpperCase('MnLpTiersEnseigne')) then
        begin
          Visible := (not MultiSelected) and (lTiersEnseigne) and (JAiLeDroitGestion('REN') or JAiLeDroitConsult('REN')) ;
        end
        else if (UpperCase(Name) = UpperCase('MnLpTiersGroupe')) then
        begin
          Visible := (not MultiSelected) and (lTiersGroupe) and (JAiLeDroitGestion('T') or JAiLeDroitConsult('T')) ;
        end
        else if (UpperCase(Name) = UpperCase('MnLpSecteurGeo')) then
        begin
          Visible := (not MultiSelected) and (lSecteurGeo) and (JAiLeDroitTag(65111));
        end
        else if (UpperCase(Name) = UpperCase('MnLpCodePostal')) then
        begin
          Visible := (not MultiSelected) and (lCodePostal) and (JAiLeDroitTag(65906));
        end
        else if (UpperCase(Name) = UpperCase('MnLpPays')) then
        begin
          Visible := (not MultiSelected) and (lPays) and (JAiLeDroitTag(65904));
        end
        else if (UpperCase(Name) = UpperCase('MnLpRegion')) then
        begin
          Visible := (not MultiSelected) and (lRegion) and (JAiLeDroitTag(65905));
        end
        else if (UpperCase(Name) = UpperCase('MnLpModeExp')) then
        begin
          Visible := (not MultiSelected) and (lModeExp) and (JAiLeDroitTag(65102));
        end
        else if UpperCase(Name) = UpperCase('MnLpBlocNote') then
        begin
          Visible := not MultiSelected;
        end
        else if (UpperCase(Name) = 'MnProperties') or (UpperCase(Name) = 'MnPropertiesWLS')then
        begin
          Visible := not MultiSelected;
        end
        else if UpperCase(Name) = UpperCase('MnLpColisage') then
        begin
          Visible := not MultiSelected;
        end
        else if UpperCase(Name) = UpperCase('MnLpPiece') then
        begin
          Visible := not MultiSelected;
        end
{$IFDEF QUALITE}
        else if UpperCase(Name) = UpperCase('MnLpPieceQNC') then
        begin
          Visible := not MultiSelected and (ClePiece<>'');
        end
{$ENDIF QUALITE}
        {--- Nouvelle Gestion des Stocks ---}
        else if (UpperCase(Name) = UpperCase('MnLpDispo')) or (UpperCase(Name) = UpperCase('MnDispo'))
             or (UpperCase(Name) = UpperCase('LpDispo'))   or (UpperCase(Name) = UpperCase('LpDispoDetail'))
//GP_BUG810_DS_GP14335 20071008
             or (pos(UpperCase('MnLpDispo'), UpperCase(Name)) > 0) then
        begin
          {$IFDEF STK}
            // GC_20071205_DM_GC12431_DEBUT
//            Visible := (not MultiSelected) and (JaiLeDroitConceptStock);
            Visible := (not MultiSelected) and (ExJaiLeDroitConcept) and (VH_GC.GCAchatStockSeria);
            // GC_20071205_DM_GC12431_FIN
          {$ELSE STK}
            Visible := false;
          {$ENDIF STK}
            Enabled := visible;
        end
        // GC_20071205_DM_GC12431_DEBUT
        {$IFDEF STK}
        else if (pos(UpperCase('mnEditCatalogueAchat'), UpperCase(Name)) > 0) then
          Visible := VH_GC.GCAchatStockSeria
        else if (pos(UpperCase('mnPiecesAch'), UpperCase(Name)) > 0) then
          Visible := VH_GC.GCAchatStockSeria
        {$ENDIF STK}
        // GC_20071205_DM_GC12431_FIN
        else if (pos(UpperCase('MnLpRefOrigine'), UpperCase(Name)) > 0) then
        begin
          {$IFDEF STK}
            Visible := true;
          {$ELSE STK}
            Visible := false;
          {$ENDIF STK}
        end
        else if pos(UpperCase('LpFicheLot'), UpperCase(Name)) > 0 then
        begin
          {$IFDEF STK}
            Visible := StkGereFicheLot and (not MultiSelected) and (LotInterne <> '');
            if LotInterne <> '' then
              Caption := TraduireMemoire('Fiche du lot') + ': ' + LotInterne;
          {$ELSE STK}
            Visible := false;
          {$ENDIF STK}
        end
        else if pos(UpperCase('LpFicheSerie'), UpperCase(Name)) > 0 then
        begin
          {$IFDEF STK}
            Visible := StkGereFicheSerie and (not MultiSelected) and (SerieInterne <> '');
            if SerieInterne <> '' then
              Caption := TraduireMemoire('Fiche du n° de série') + ': ' + SerieInterne;
          {$ELSE STK}
            Visible := false;
          {$ENDIF STK}
        end
        else if pos(UpperCase('LpNomeLot'), UpperCase(Name)) > 0 then
        begin
          {$IFDEF STK}
            Visible := StkGereFicheLot and (not MultiSelected) and (LotInterne <> '');
          {$ELSE STK}
            Visible := false;
          {$ENDIF STK}
        end
        else if pos(UpperCase('LpNomeSer'), UpperCase(Name)) > 0 then
        begin
          {$IFDEF STK}
            Visible := StkGereFicheSerie and (not MultiSelected) and (SerieInterne <> '');
          {$ELSE STK}
            Visible := false;
          {$ENDIF STK}
        end
        else if pos(UpperCase('LpEmplLot'), UpperCase(Name)) > 0 then
        begin
          {$IFDEF STK}
            Visible := StkGereFicheLot and (not MultiSelected) and (LotInterne <> '');
          {$ELSE STK}
            Visible := false;
          {$ENDIF STK}
        end
        else if pos(UpperCase('LpEmplSer'), UpperCase(Name)) > 0 then
        begin
          {$IFDEF STK}
            Visible := StkGereFicheSerie and (not MultiSelected) and (SerieInterne <> '');
          {$ELSE STK}
            Visible := false;
          {$ENDIF STK}
        end
        else if UpperCase(Name) = UpperCase('MnLpOrdrePere') then
        begin
          {$IFDEF GPAO}
            Visible := wGereSousOrdres and (OrdrePere <> 0) and (not MultiSelected) and (JAiLeDroitGestion('WOL') or JAiLeDroitConsult('WOL'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('MnLpNomeWOL') then
        begin
          {$IFDEF GPAO}
            Visible := wGereSousOrdres and (not MultiSelected) and (JAiLeDroitGestion('WOL') or JAiLeDroitConsult('WOL'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('MnLpCasWOL') then
        begin
          {$IFDEF GPAO}
            Visible := wGereSousOrdres and (not MultiSelected) and (JAiLeDroitGestion('WOL') or JAiLeDroitConsult('WOL'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('MnLpPlanDeCharge') then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and JAiLeDroitConsult('WPH');
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if (UpperCase(Name) = UpperCase('MnLpBasesTech')) or (UpperCase(Name) = UpperCase('MnBasesTechniques')) then
        begin
          {$IFDEF GPAO}
            //GP_JLS_GP14648_20071212
            Visible := ((Pos(TypeArticle,'MAR;CNS') > 0) or (TypeArticle = ''));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Pos(UpperCase('LpVisuArt'), UpperCase(Name)) > 0  then
        begin
          {$IFDEF GPAO}
            //GP_DS_GP14395_20071010
            //GP_JLS_GP14648_20071212
            Visible := ((Pos(TypeArticle,'MAR;CNS') > 0) or (TypeArticle = '')) and (JAiLeDroitGestion('WNT') or JAiLeDroitConsult('WNT'))
                                                                                and (JAiLeDroitGestion('WGT') or JAiLeDroitConsult('WGT'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Pos(UpperCase('LpCasEmploi'), UpperCase(Name)) > 0 then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WNT') or JAiLeDroitConsult('WNT'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Pos(UpperCase('LpNome'), UpperCase(Name)) > 0  then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WNT') or JAiLeDroitConsult('WNT'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Pos(UpperCase('LpVisuNome'), UpperCase(Name)) > 0 then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WNT') or JAiLeDroitConsult('WNT'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Pos(UpperCase('LpVisuGamme'), UpperCase(Name)) > 0 then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WGT') or JAiLeDroitConsult('WGT'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Pos(UpperCase('LpCasEmploi_PRD'), UpperCase(Name)) > 0  then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WNT') or JAiLeDroitConsult('WNT'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Pos(UpperCase('LpNome_PRD'), UpperCase(Name)) > 0 then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WNT') or JAiLeDroitConsult('WNT'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Pos(UpperCase('LpVisuNome_PRD'), UpperCase(Name)) > 0  then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WNT') or JAiLeDroitConsult('WNT'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Pos(UpperCase('LpGamme_PRD'), UpperCase(Name)) > 0 then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WGT') or JAiLeDroitConsult('WGT'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Pos(UpperCase('VisuGamme_PRD'), UpperCase(Name)) > 0 then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WGT') or JAiLeDroitConsult('WGT'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Pos(UpperCase('LpGamme_SAV'), UpperCase(Name)) > 0  then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WGT') or JAiLeDroitConsult('WGT'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Pos(UpperCase('LpVisuGamme_SAV'), UpperCase(Name)) > 0  then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WGT') or JAiLeDroitConsult('WGT'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('MnLpCircuit') then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (Circuit<>'') and (JAiLeDroitGestion('QCI') or JAiLeDroitConsult('QCI'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if (UpperCase(Name) = UpperCase('MnLegendeWOL')) or (UpperCase(Name) = UpperCase('MnLegendeWOT')) or (UpperCase(Name) = UpperCase('MnLegendeWOP')) or (UpperCase(Name) = UpperCase('MnLegendeWOB')) or (UpperCase(Name) = UpperCase('MnLegendeWOG')) or (UpperCase(Name) = UpperCase('MnLegendeWOG')) or (UpperCase(Name) = UpperCase('MnLegendeWLS')) then
        begin
          Visible := True;
        end
        else if UpperCase(Name) = UpperCase('MnSaisieQteWOG') then
        begin
          V_PGI.GraphList.GetBitmap(91-1, BitMap);
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (EtatOpe <> 'DEC') and (JAiLeDroitGestion('WOP') or JAiLeDroitConsult('WOP'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if (pos(UpperCase('MnSaisieTps'), UpperCase(Name)) > 0) then
        begin
          V_PGI.GraphList.GetBitmap(90-1, BitMap);
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (EtatOpe <> 'DEC') and (JAiLeDroitGestion('WOP') or JAiLeDroitConsult('WOP'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if (UpperCase(Name) = UpperCase('MnLpItineraire')) then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (Coditi<>'') and (JAiLeDroitGestion('QIT') or JAiLeDroitConsult('QIT')) and (not VH_GC.OASeria);
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if (UpperCase(Name) = UpperCase('MnLpGamme')) then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WGT') or JAiLeDroitConsult('WGT')) AND (not VH_GC.OASeria);
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if (UpperCase(Name) = UpperCase('MnLignesCSP')) then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (TypeOrdre = 'STP') and (EtatPHASE <> 'DEC');
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        {--- Gestion du S.A.V. ---}
        else if (pos(UpperCase('LpArtNat'), UpperCase(Name)) > 0) then
        begin
          {$IFDEF GPAO}
            //GP_DS_GP14395_20071010
            //GP_JLS_GP14648_20071212
            Visible := ((Pos(TypeArticle,'MAR;CNS') > 0) or (TypeArticle = '')) and (not MultiSelected);
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('MnLpCasEmploi_SAV') then
        begin
          {$IFDEF SAV}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WNT') or JAiLeDroitConsult('WNT'));
          {$ELSE SAV}
            Visible := false;
          {$ENDIF SAV}
        end
        else if UpperCase(Name) = UpperCase('MnLpNome_SAV') then
        begin
          {$IFDEF SAV}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WNT') or JAiLeDroitConsult('WNT'));
          {$ELSE SAV}
            Visible := false;
          {$ENDIF SAV}
        end
        else if UpperCase(Name) = UpperCase('MnLpVisuNome_SAV') then
        begin
          {$IFDEF SAV}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WNT') or JAiLeDroitConsult('WNT'));
          {$ELSE SAV}
            Visible := false;
          {$ENDIF SAV}
        end
        else if UpperCase(Name) = UpperCase('MnLpRessource') then
        begin
          {$IFDEF SAV}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('ARS') or JAiLeDroitConsult('ARS'));
          {$ELSE SAV}
            Visible := false;
          {$ENDIF SAV}
        end
        else if UpperCase(Name) = UpperCase('MnLpParc') then
        begin
          {$IFDEF SAV}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WPC') or JAiLeDroitConsult('WPC'));
          {$ELSE SAV}
            Visible := false;
          {$ENDIF SAV}
        end
        else if UpperCase(Name) = UpperCase('MnLpParcNome') then
        begin
          {$IFDEF SAV}
            V_PGI.GraphList.GetBitmap(55-1, BitMap);
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WPC') or JAiLeDroitConsult('WPC'));
          {$ELSE SAV}
            Visible := false;
          {$ENDIF SAV}
        end
        else if UpperCase(Name) = UpperCase('MnLpInfosCompl') then
        begin
          {$IFDEF SAV}  
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WPC') or JAiLeDroitConsult('WPC')) and GetParamSoc('SO_RTGESTINFOS005');
          {$ELSE SAV}
            Visible := false;
          {$ENDIF SAV}
        end
        else if UpperCase(Name) = UpperCase('MnLpInfosComplLOT') then
        begin
          {$IFDEF STK}
            Visible := (not MultiSelected) and StkGereFicheLot and GetParamSoc('SO_RTGESTINFOS00H');
          {$ELSE STK}
            Visible := false;
          {$ENDIF STK}
        end
        else if UpperCase(Name) = UpperCase('MnLpInfosComplSER') then
        begin
          {$IFDEF STK}
            Visible := (not MultiSelected) and StkGereFicheSerie and GetParamSoc('SO_RTGESTINFOS00J');
          {$ELSE STK}
            Visible := false;
          {$ENDIF STK}
        end
{$IFDEF QUALITE}
        else if UpperCase(Name) = UpperCase('MnLpInfosComplRQN') then
        begin
          Visible := (not MultiSelected) and GetParamSoc('SO_RTGESTINFOS00K');
        end
        else if UpperCase(Name) = UpperCase('MnLpInfosComplRQD') then
        begin
          Visible := (not MultiSelected) and GetParamSoc('SO_RTGESTINFOS00N');
        end
        else if UpperCase(Name) = UpperCase('MnFicheInfos') then
        begin
          Visible := (not MultiSelected) and GetParamSoc('SO_RTGESTINFOS001');
        end
        else if UpperCase(Name) = UpperCase('MnLpInfosComplRQD') then
        begin
          Visible := (not MultiSelected) and GetParamSoc('SO_RTGESTINFOS00N');
        end
{$ENDIF QUALITE}
        else if UpperCase(Name) = UpperCase('MnLpInterv') then   // L'acces aux interventions n'est possible qu'en GP
        begin
            Visible := (not MultiSelected) and (ctxGRC in V_PGI.PGIContexte);
        end
        {--- Gestion de l E.D.I. --- Début --- }
        else if (UpperCase(Name)= UpperCase('MnLpFichierLu')) then
        begin
          {$IFDEF EDI}
            Visible := (not MultiSelected);
          {$ELSE EDI}
            Visible := False;
          {$ENDIF EDI}
        end
        else if (UpperCase(Name)= UpperCase('MnConfig')) then
        begin
          {$IFDEF EDI}
            Visible := (not MultiSelected);
          {$ELSE EDI}
            Visible := False;
          {$ENDIF EDI}
        end
        else if (UpperCase(Name)= UpperCase('MnLpMessageEDI')) then
        begin
          {$IFDEF EDI}
            Visible := (not MultiSelected);
          {$ELSE EDI}
            Visible := False;
          {$ENDIF EDI}
        end
        {--- Gestion de l E.D.I. --- Fin --- }
        else if (UpperCase(Name)= UpperCase('MnDispoMulti')) then
        begin
           Visible := (EstBaseMultiSoc and TablePartagee ('ARTICLE'));
        end
        else if UpperCase(Name)= 'MNSAV' then
          Visible := V_Pgi.SAV
{$IFDEF QUALITE}
				{ Qualité }
        else if UpperCase(Name) = UpperCase('MnDemDerog') then
        begin
          V_PGI.GraphList.GetBitmap(102-1, BitMap);
          Visible := (not MultiSelected) and (DemDerogNum <> 0);
        end
        else if UpperCase(Name) = UpperCase('MnLpOrdre') then
        begin
          V_PGI.GraphList.GetBitmap(108-1, BitMap);
          Visible := (not MultiSelected) and (LigneOrdre <> 0) and (lNatureTra);
        end
        else if UpperCase(Name) = UpperCase('MnLpNonConf') then
        begin
//          V_PGI.GraphList.GetBitmap(117-1, BitMap);
          Visible := (not MultiSelected) and (lRQN) and (JAiLeDroitGestion('RQN'));
        end
        else if UpperCase(Name) = UpperCase('MnNotationFour') then
        begin
          Visible := (not MultiSelected) and (JAiLeDroitGestion('RQT'));
        end
        else if UpperCase(Name) = UpperCase('MnLpPlanCorr') then
        begin
//          V_PGI.GraphList.GetBitmap(102-1, BitMap);
          Visible := (not MultiSelected) and (PlanCorrNum <> 0);
        end
        else if UpperCase(Name) = UpperCase('MnLpAction') then
        begin
//          V_PGI.GraphList.GetBitmap(90-1, BitMap);
          Visible := (not MultiSelected) and (IdAction <> 0);
        end
        else if UpperCase(Name) = UpperCase('MnLpGroupe') then
        begin
          Visible := not MultiSelected and (Groupe<>'');
        end
        else if UpperCase(Name) = UpperCase('MnLpCritere') then
        begin
          Visible := not MultiSelected and (Critere<>'');
        end
{$ENDIF QUALITE}
        else if (UpperCase(Name) = UpperCase('MnLpFormuleMt')) then
        begin
          Visible := (not MultiSelected) and (lFormuleMt) and (JAiLeDroitGestion('YTP') or JAiLeDroitConsult('YTP')) ;
        end
        else if (pos(UpperCase('MnLpRubrique'), UpperCase(Name)) > 0) then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and GereRubriquesPdr and (RubriquePDR<>'') and (RubriquePDR<>'NON AFFECTE') ; ;
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if (pos(UpperCase('MnLpSection'), UpperCase(Name)) > 0) then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and GereSectionsPdr and (SectionPDR<>'') and (SectionPDR<>'NON AFFECTE') ;
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
//GP_20071015_MM_GP14362
        else if (pos(UpperCase('MnLpPdR'), UpperCase(Name)) > 0) then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WPE'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if (pos(UpperCase('MnLpPdrOrdre'), UpperCase(Name)) > 0) then
        begin
          {$IFDEF GPAO}
            Visible := (not MultiSelected) and (JAiLeDroitGestion('WOL'));
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        ;

        if Visible and (Count > 0) then
        begin
          Visible := false;
          for j := 0 to Pred(Count) do
            Visible := wSetMn(Items[j]) or Visible;
        end;
        Result := Visible;
      end
    end
  end;

begin
	ChargeImageList;

  { Argument }
  lCodePort     := (Pos('CODEPORT'      , Argument)<=0) or (GetArgumentString(Argument, 'CODEPORT'      )<>'');
  lTarifSpecial := (Pos('TARIFSPECIAL'  , Argument)<=0) or (GetArgumentString(Argument, 'TARIFSPECIAL'  )<>'');
  lTiers        := (Pos('TIERS'         , Argument)<=0) or (GetArgumentString(Argument, 'TIERS'         )<>'');
  lTarifTiers   := (Pos('TARIFTIERS'    , Argument)<=0) or (GetArgumentString(Argument, 'TARIFTIERS'    )<>'');
  lTiersOrigine := (Pos('ORIGINETIERS'  , Argument)<=0) or (GetArgumentString(Argument, 'ORIGINETIERS'  )<>'');
  lTiersActivite:= (Pos('SECTEUR'       , Argument)<=0) or (GetArgumentString(Argument, 'SECTEUR'       )<>'');
  lTiersEnseigne:= (Pos('CODEENSEIGNE'  , Argument)<=0) or (GetArgumentString(Argument, 'CODEENSEIGNE'  )<>'');
  lTiersGroupe  := (Pos('SOCIETEGROUPE' , Argument)<=0) or (GetArgumentString(Argument, 'SOCIETEGROUPE' )<>'');
  //GP_DS_AFFAIRE_20070925
  {$IFDEF AFFAIRE}
  lAffaire      := (Pos('AFFAIRE'       , Argument)<=0) or (GetArgumentString(Argument, 'AFFAIRE'       )<>'');
  {$ENDIF AFFAIRE}
  lNatureTra    := (Pos('NATURETRAVAIL' , Argument)<=0) or (GetArgumentString(Argument, 'NATURETRAVAIL' )<>'');
  lArticle      := (Pos('ARTICLE'       , Argument)<=0) or (GetArgumentString(Argument, 'ARTICLE'       )<>'');
  lTarifArticle := (Pos('TARIFARTICLE'  , Argument)<=0) or (GetArgumentString(Argument, 'TARIFARTICLE'  )<>'');
  lFamilleValo  := (Pos('FAMILLEVALO'   , Argument)<=0) or (GetArgumentString(Argument, 'FAMILLEVALO'   )<>'');
  lFamilleNiv1  := (Pos('FAMILLENIV1'   , Argument)<=0) or (GetArgumentString(Argument, 'FAMILLENIV1'   )<>'');
  lFamilleNiv2  := (Pos('FAMILLENIV2'   , Argument)<=0) or (GetArgumentString(Argument, 'FAMILLENIV2'   )<>'');
  lFamilleNiv3  := (Pos('FAMILLENIV3'   , Argument)<=0) or (GetArgumentString(Argument, 'FAMILLENIV3'   )<>'');
{$IFDEF STK}
  {$IFDEF GPAO}
  lMarque       := (Pos('MARQUE'        , Argument)<=0) or (GetArgumentString(Argument, 'MARQUE'        )<>'');
  lChoixQualite := (Pos('CHOIXQUALITE'  , Argument)<=0) or (GetArgumentString(Argument, 'CHOIXQUALITE'  )<>'');
  {$ENDIF GPAO}
  lIndiceArticle:= (Pos('INDICEARTICLE' , Argument)<=0) or (GetArgumentString(Argument, 'INDICEARTICLE' )<>'');
{$ENDIF STK}
  lRessource    := (Pos('RESSOURCE'     , Argument)<=0) or (GetArgumentString(Argument, 'RESSOURCE'     )<>'');
  lTarifResso   := (Pos('TARIFRESSOURCE', Argument)<=0) or (GetArgumentString(Argument, 'TARIFRESSOURCE')<>'');
  lSecteurGeo   := (Pos('SECTEURGEO'    , Argument)<=0) or (GetArgumentString(Argument, 'SECTEURGEO'    )<>'');
  lCodePostal   := (Pos('CODEPOSTAL'    , Argument)<=0) or (GetArgumentString(Argument, 'CODEPOSTAL'    )<>'');
  lPays         := (Pos('PAYS'          , Argument)<=0) or (GetArgumentString(Argument, 'PAYS'          )<>'');
  lRegion       := (Pos('REGION'        , Argument)<=0) or (GetArgumentString(Argument, 'REGION'        )<>'');
  lModeExp      := (Pos('MODEEXP'       , Argument)<=0) or (GetArgumentString(Argument, 'MODEEXP'       )<>'');
  lFormuleMt    := (Pos('FORMULEMT'     , Argument)<=0) or (GetArgumentString(Argument, 'FORMULEMT'     )<>'');
  lCatalogue    := (Pos('CATALOGUE'     , Argument)<=0) or (GetArgumentString(Argument, 'CATALOGUE'     )<>'');
  lWithPiece    := GetArgumentBoolean(Argument, 'WITHPIECE');

  Article       := GetArgumentString(Argument, 'ARTICLE');
  TypeArticle   := GetArgumentString(Argument, 'TYPEARTICLE');
  Depot         := GetArgumentString(Argument, 'DEPOT');
  EtatOpe       := GetArgumentString(Argument, 'ETATOPE');
  TypeOrdre     := GetArgumentString(Argument, 'TYPEORDRE');
  EtatPhase     := GetArgumentString(Argument, 'ETATPHASE');
  LotInterne    := GetArgumentString(Argument, 'LOTINTERNE');
  SerieInterne  := GetArgumentString(Argument, 'SERIEINTERNE');

  {$IFDEF GPAO}
    OrdrePere := GetArgumentInteger(Argument, 'ORDREPERE');
    Coditi        := GetArgumentString(Argument, 'CODITI');
    Circuit       := GetArgumentString(Argument, 'CIRCUIT');
    RubriquePDR   := GetArgumentString(Argument, 'RUBRIQUEPDR');
    SectionPDR    := GetArgumentString(Argument, 'SECTIONPDR');
  {$ENDIF GPAO}
{$IFDEF QUALITE}
  DemDerogNum		:= GetArgumentInteger(Argument, 'DEMDEROGNUM');
  lRQN  				:= iif(Pos('QNCNUM', Argument)>0, iif(GetArgumentInteger(Argument, 'QNCNUM')<>0,True,False),True);
  PlanCorrNum		:= GetArgumentInteger(Argument, 'PLANCORRNUM');
  IdAction			:= GetArgumentInteger(Argument, 'IDACTION');
  LigneOrdre		:= GetArgumentInteger(Argument, 'LIGNEORDRE');
  ClePiece	 		:= GetArgumentString(Argument , 'CLEPIECE');
  Groupe		 		:= GetArgumentString(Argument , 'GROUPE');
  Critere 	 		:= GetArgumentString(Argument , 'CRITERE');
{$ENDIF QUALITE}

	for i := 0 to Pm.Items.Count - 1 do
  begin
    wSetMn(Pm.Items[i]);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 22/05/2002
Modifié le ... :   /  /
Description .. : Gestion du menu historique
Mots clefs ... :
*****************************************************************}
procedure wSetMnuHisto(Const MultiSelected: Boolean; Const Argument: string; Const Pm: TPopUpMenu);
var
	i: integer;
  TypeLien, ModeConso : string;

  function wSetMn(Const Mn: TMenuItem): boolean;
  var
    j: Integer;
  begin
    Result := true;
    if Assigned(Mn) then
    begin
      with Mn do
      begin
        if UpperCase(Name) = UpperCase('HiConsoWOB') then
        begin
          V_PGI.GraphList.GetBitmap(53-1, BitMap);
//GP_20080220_PCO_GP14733 deb
          Visible := not MultiSelected and (((TypeLien='COM') and (ModeConso <> 'NON')) or (TypeLien='CPD'));
//GP_20080220_PCO_GP14733 fin

        end
        else if UpperCase(Name) = UpperCase('HiConsoWOP') then
        begin
          V_PGI.GraphList.GetBitmap(53-1, BitMap);
          Visible := not MultiSelected;
        end
        else if UpperCase(Name) = UpperCase('HiConsoWOL') then
        begin
          V_PGI.GraphList.GetBitmap(53-1, BitMap);
          Visible := not MultiSelected;
        end
        else if UpperCase(Name) = UpperCase('HiCoProdWOB') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(52-1, BitMap);
            Visible := PrdGereCoProduction and (not MultiSelected) and (TypeLien = 'CPD') and (ModeConso <> 'NON');
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('HiCoProdWOP') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(52-1, BitMap);
            Visible := PrdGereCoProduction and (not MultiSelected)
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('HiCoProdWOL') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(52-1, BitMap);
            Visible := PrdGereCoProduction and (not MultiSelected)
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if Name = 'HiAffectWOB' then
        begin
          Visible := false;
          {$IFDEF GPAOLIGHT}
            V_PGI.GraphList.GetBitmap(49-1, BitMap);
            Visible := (not MultiSelected) and (TypeLien = 'COM') and (ModeConso <> 'NON');
          {$ENDIF GPAOLIGHT}
        end
        else if Pos('HiRecep', Name) > 0 then
        begin
          V_PGI.GraphList.GetBitmap(41-1, BitMap);
          Visible := not MultiSelected;
        end
        else if Pos('HiDates', Name) >0 then
        begin
          V_PGI.GraphList.GetBitmap(23-1, BitMap);
          Visible := not MultiSelected;
        end
        else	if Name = 'HistoRetoursWPC' then
        begin
          V_PGI.GraphList.GetBitmap(75-1, BitMap);
          Visible := not MultiSelected;
        end
        else	if Name = 'HistoInterventionsWPC' then
        begin
          V_PGI.GraphList.GetBitmap(81-1, BitMap);
          Visible := not MultiSelected;
        end
        else	if Name = 'HistoMouvementsWPC' then
        begin
          V_PGI.GraphList.GetBitmap(71-1, BitMap);
          Visible := not MultiSelected;
        end
        else	if Pos('HiPeriode', Name) > 0 then
        begin
          Caption := GetArgumentString(Argument, UpperCase(Name), false);
          Visible := (not MultiSelected) and (Pos(UpperCase(Name), Argument) > 0);
        end
        ;

        if Visible and (Count > 0) then
        begin
          Visible := false;
          for j := 0 to Pred(Count) do
            Visible := wSetMn(Items[j]) or Visible;
        end;
        Result := Visible;
      end
    end
  end;

begin
	ChargeImageList;

  { Contexte }
  TypeLien  := GetArgumentValue(Argument, 'TYPELIEN');
  ModeConso := GetArgumentValue(Argument, 'MODECONSO');

	for i := 0 to Pm.Items.Count - 1 do
    wSetMn(Pm.Items[i])
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 22/05/2002
Modifié le ... :   /  /
Description .. : Gestion du menu Action
Mots clefs ... :
*****************************************************************}
procedure wSetMnuAction(Const MultiSelected: Boolean; Const Argument: string; Const Pm: TPopUpMenu);
var
	i 	: integer;
  EtatTet, EtatLig, EtatPhase, EtatBes, EtatLas, EtatRev, EtatParc, EtatVer, EtatOpe, EtatRes,
  EtatSta, ModeConso, QAccSais, QLanSais, sAction, Fourni,EtatLasst, MiseEnProd, TypeLien: string;
  {$If defined (GPAOLIGHT) or defined (CRM)}
  IsSST: boolean;
  PiecesCSPEnCours: Boolean;
  {$Ifend}

  function wSetMn(Const Mn: TMenuItem): boolean;
  var
    j: Integer;
  begin
    Result := true;
    if Assigned(Mn) then
    begin
      with Mn do
      begin
        {$If defined (GPAOLIGHT) or defined (CRM)}
        if sAction = 'CONSULTATION' then
        begin
          if Pos(UpperCase(Name), 'MNJOURNALACTION;MNJOURNALREVISION;MNPARAMS') = 0 then
            Visible := False;
        end
        else
        begin
          { WOT }
          if upperCase(Name) = 'MNFERMEWOT' then
          begin
            V_PGI.GraphList.GetBitmap(5-1, BitMap);
            Visible := JAiLeDroitAction(wtaFermeWOT) and ((EtatTet = 'OUV') or MultiSelected);
          end
          else if UpperCase(Name) = 'MNSOLDEWOT' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := JAiLeDroitAction(wtaSoldeWOT) and ((Pos(EtatTet, 'OUV;FER') >0) or MultiSelected);
          end
          { WOL }
          else if UpperCase(Name) = 'MNDECOMPOSEWOL' then
          begin
            V_PGI.GraphList.GetBitmap(42-1, BitMap);
            Visible := JAiLeDroitAction(wtaDecomposeWOL) and (((MiseEnProd <> 'ALL') and (Pos(EtatLig, 'DCL') > 0)) or Multiselected);
          end
          else if UpperCase(Name) = 'MNLANCEWOL' then
          begin
            V_PGI.GraphList.GetBitmap(44-1, BitMap);
            Visible := JAiLeDroitAction(wtaLanceWOL) and (({(MiseEnProd <> 'ALL') and} (Pos(EtatLig, 'DEC') > 0)) or Multiselected);
          end
          else if UpperCase(Name) = 'MNRECEPTIONWOL' then
          begin
            V_PGI.GraphList.GetBitmap(41-1, BitMap);
            Visible := JAiLeDroitAction(wtaReceptionWOL)
                       and (MiseEnProd <> 'ALL')
                       and (Pos(EtatLig, 'LAN;REC;TER') > 0);
          end
          else if UpperCase(Name) = 'MNTERMINEWOL' then
          begin
            V_PGI.GraphList.GetBitmap(43-1, BitMap);
            Visible := JAiLeDroitAction(wtaTermineWOL) and (((MiseEnProd <> 'ALL')
                       and (Pos(EtatLig, 'LAN;REC')  > 0))
                       or Multiselected);
          end
          else if UpperCase(Name) = 'MNSOLDEWOL' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := JAiLeDroitAction(wtaSoldeWOL) and (((MiseEnProd <> 'ALL') and (Pos(EtatLig, 'DCL;DEC;LAN;REC;TER') > 0)) or (Pos(EtatSTA, 'TER') > 0) or Multiselected);
          end
          { WOP }
          else if UpperCase(Name) = 'MNLANCEWOP' then
          begin
            V_PGI.GraphList.GetBitmap(44-1, BitMap);
            Visible := JAiLeDroitAction(wtaLanceWOP) and (( (not isSST ) and (Pos(EtatLig, 'LAN;REC;TER') > 0) and (Pos(EtatPhase, 'DEC;LAN;REC') > 0) and (ValeurI(QAccSais)>ValeurI(QLanSais))) or Multiselected);
          end
          else if UpperCase(Name) = 'MNLANPARTIELWOP' then
          begin
            V_PGI.GraphList.GetBitmap(44-1, BitMap);
            Visible := JAiLeDroitAction(wtaLanceWOP) and  (not isSST) and (Pos(EtatLig, 'LAN;REC;TER') > 0) and (Pos(EtatPhase, 'DEC;LAN;REC') > 0) and (not Multiselected);
          end
          else if UpperCase(Name) = 'MNRECEPTIONWOP' then
          begin
            V_PGI.GraphList.GetBitmap(41-1, BitMap);
            Visible := JAiLeDroitAction(wtaReceptionWOP) and (( (not isSST) and (Pos(EtatLig, 'LAN;REC;TER') > 0) and (Pos(EtatPhase, 'LAN;REC') > 0)) or Multiselected);
          end
          else if UpperCase(Name) = 'MNRECPARTIELLEWOP' then
          begin
            V_PGI.GraphList.GetBitmap(41-1, BitMap);
            Visible := JAiLeDroitAction(wtaReceptionWOP) and (not isSST) and (Pos(EtatLig, 'LAN;REC;TER') > 0) and (Pos(EtatPhase, 'LAN;REC;TER') > 0) and (not Multiselected);
          end
          else if UpperCase(Name) = 'MNTERMINEWOP' then
          begin
            V_PGI.GraphList.GetBitmap(43-1, BitMap);
            Visible := JAiLeDroitAction(wtaTermineWOP) and (( (not isSST) and (Pos(EtatLig, 'LAN;REC;TER') > 0) and (Pos(EtatPhase, 'DEC;LAN;REC') > 0)) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNSOLDEWOP' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := JAiLeDroitAction(wtaSoldeWOP) and (( ((not isSST) or (not PiecesCSPEnCours)) and (Pos(EtatLig, 'LAN;REC;TER') > 0) and (Pos(EtatPhase, 'DEC;LAN;REC;TER') > 0)) or MultiSelected);
          end
          {  WOB }
          else if UpperCase(Name) = 'MNINTEGREWOB' then
          begin
            V_PGI.GraphList.GetBitmap(52-1, BitMap);
            Visible := JAiLeDroitAction(wtaConsoManWOB)
                       and (Pos(EtatLig  , 'LAN;REC;TER') > 0)
                       and (Pos(EtatPhase, 'LAN;REC;TER') > 0)
                       and (Pos(EtatBes  , 'CON;RUP;TER') > 0)
                       and (ModeConso <> 'NON')
                       and (TypeLien = 'COM')
                       and (not MultiSelected);
          end
//GP_20080125_PCO_GP14733 deb
          else if UpperCase(Name) = 'MNINTEGREAUTOWOB' then
          begin
            V_PGI.GraphList.GetBitmap(52-1, BitMap);
//GP_20080701_PCO_GP15140 deb
            Visible := JaiLeDroitTag(120348);
//GP_20080701_PCO_GP15140 fin

          end
//GP_20080125_PCO_GP14733 fin
          else if UpperCase(Name) = 'MNUNDORECEPWOB' then
          begin
            V_PGI.GraphList.GetBitmap(53-1, BitMap);
            Visible := JAiLeDroitAction(wtaConsoManWOB)
                       and (Pos(EtatLig  , 'LAN;REC;TER') > 0)
                       and (Pos(EtatPhase, 'LAN;REC;TER') > 0)
                       and (Pos(EtatBes  , 'CON;TER') > 0)
                       and (TypeLien = 'CPD')
                       and (not MultiSelected);
          end
          else if UpperCase(Name) = 'MNCONSOMANUWOB' then
          begin
            V_PGI.GraphList.GetBitmap(53-1, BitMap);
            Visible := JAiLeDroitAction(wtaConsoManWOB)
                       and (Pos(EtatLig, 'LAN;REC;TER') > 0)
                       and (Pos(EtatPhase, 'LAN;REC;TER') > 0)
                       and (Pos(EtatBes  , 'BES;CON;TER;RUP') > 0)
                       and (ModeConso <> 'NON')
                       and (TypeLien = 'COM')
                       and (not MultiSelected);
          end
//GP_20080125_PCO_GP14733 deb
          else if UpperCase(Name) = 'MNCONSOAUTOWOB' then
          begin
            V_PGI.GraphList.GetBitmap(53-1, BitMap);
//GP_20080701_PCO_GP15140 deb
            Visible := JAiLeDroitTag(120346);
//GP_20080701_PCO_GP15140 fin
          end
          else if UpperCase(Name) = 'MNCONSOAUTOAFFWOB' then
          begin
            V_PGI.GraphList.GetBitmap(49-1, BitMap);
//GP_20080701_PCO_GP15140 deb
            Visible := JAiLeDroitTag(120347);
//GP_20080701_PCO_GP15140 fin
          end
//GP_20080125_PCO_GP14733 fin
          else if UpperCase(Name) = 'MNRECEPMANUWOB' then
          begin
            Visible := false;
            {$IFDEF GPAO}
              V_PGI.GraphList.GetBitmap(52-1, BitMap);
              Visible := JAiLeDroitAction(wtaConsoManWOB)
                         and (Pos(EtatLig  , 'LAN;REC;TER') > 0)
                         and (Pos(EtatPhase, 'LAN;REC;TER') > 0)
                         and (Pos(EtatBes  , 'BES;CON;TER') > 0)
                         and (TypeLien = 'CPD')
                         and (not MultiSelected);
            {$ENDIF GPAO}
          end
          else if UpperCase(Name) = 'MNTERMINEWOB' then
          begin
            V_PGI.GraphList.GetBitmap(43-1, BitMap);
            Visible := JAiLeDroitAction(wtaTermineWOB) and (((Pos(EtatLig, 'LAN;REC') > 0) and (Pos(EtatPhase, 'LAN;REC') > 0) and (Pos(EtatBes, 'BES;CON') > 0)) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNSOLDEWOB' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := JAiLeDroitAction(wtaSoldeWOB) and (((Pos(EtatLig, 'LAN;REC;TER') > 0) and (Pos(EtatPhase, 'LAN;REC;TER') > 0) and (Pos(EtatBes, 'BES;CON;RUP;TER') > 0)) or MultiSelected);
          end
          { WOG }
          else if UpperCase(Name) = 'MNTERMINEWOG' then
          begin
            V_PGI.GraphList.GetBitmap(43-1, BitMap);
            Visible := JAiLeDroitAction(wtaTermineWOG) and ((Pos(EtatOpe, 'LAN;COM') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNSOLDEWOG' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := JAiLeDroitAction(wtaSoldeWOG) and ((Pos(EtatOpe, 'DEC;LAN;COM;TER') > 0) or MultiSelected);
          end
          { WOR }
          else if UpperCase(Name) = 'MNTERMINEWOR' then
          begin
            V_PGI.GraphList.GetBitmap(43-1, BitMap);
            Visible := JAiLeDroitAction(wtaTermineWOR) and ((Pos(EtatRes, 'LAN;COM') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNSOLDEWOR' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := JAiLeDroitAction(wtaSoldeWOR) and ((Pos(EtatRes, 'DEC;LAN;COM;TER') > 0) or MultiSelected);
          end
          { WLS }
          else if UpperCase(Name) = 'MNSELECTIONWLS' then
          begin
            V_PGI.GraphList.GetBitmap(47-1, BitMap);
            Visible := (Pos(EtatLas, 'DES;SEL;SOL') > 0) or MultiSelected;
          end
          else if UpperCase(Name) = 'MNDESELECTIONWLS' then
          begin
            Visible := (Pos(EtatLas, 'SEL') > 0) or MultiSelected;
          end
          else if UpperCase(Name) = 'MNSOLDEWLS' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := (Pos(EtatLas, 'DES;SEL') > 0) or MultiSelected;
          end
          else if (UpperCase(Name) = 'MNAFFECTATIONWLS') or (UpperCase(Name)='MNINITIALISATIONULS') or (UpperCase(Name) = 'MNRAZEROQTES') then
          begin
            Visible := ((Pos(EtatLas, 'SEL') > 0) and (Pos(EtatBes  , 'BES;CON;TER;RUP') > 0))
                    or ((Pos(EtatLasst, 'SEL') > 0) and (Pos(EtatBes  , 'BES;CON;TER;RUP') > 0))
                    or MultiSelected;
          end
          else if UpperCase(Name) = 'MNCONSOWLS' then
          begin
            V_PGI.GraphList.GetBitmap(53-1, BitMap);
            Visible := ((Pos(EtatLas, 'SEL') > 0)
                       and (JAiLeDroitAction(wtaConsoLasWOB))
                       and (Pos(EtatBes  , 'BES;CON;TER;RUP') > 0))
                       or MultiSelected;
          end
          { WNT }
          else if UpperCase(Name) = 'MNVALIDEWNT' then
          begin
            V_PGI.GraphList.GetBitmap(47-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and JAiLeDroitAction(wtaValideWNT) and ((Pos(EtatRev, 'MOD;PER') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNVALIDEDATEWNT' then
          begin
            V_PGI.GraphList.GetBitmap(89-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and JAiLeDroitAction(wtaValideDateWNT) and ((Pos(EtatRev, 'MOD;PER') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNMINEURWNT' then
          begin
            V_PGI.GraphList.GetBitmap(45-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and JAiLeDroitAction(wtaMineurWNT) and ((Pos(EtatRev, 'VAL') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNMAJEURWNT' then
          begin
            V_PGI.GraphList.GetBitmap(5-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and JAiLeDroitAction(wtaMajeurWNT) and ((Pos(EtatRev, 'VAL') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNPERIMEWNT' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and JAiLeDroitAction(wtaPerimeWNT) and ((Pos(EtatRev, 'VAL;MOD') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNPERIMEDATEWNT' then
          begin
            V_PGI.GraphList.GetBitmap(88-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and JAiLeDroitAction(wtaPerimeDateWNT) and ((Pos(EtatRev, 'VAL;MOD') > 0) or MultiSelected);
          end
          { WGT }
          else if UpperCase(Name) = 'MNVALIDEWGT' then
          begin
            V_PGI.GraphList.GetBitmap(47-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and JAiLeDroitAction(wtaValideWGT) and ((Pos(EtatRev, 'MOD;PER') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNVALIDEDATEWGT' then
          begin
            V_PGI.GraphList.GetBitmap(89-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and JAiLeDroitAction(wtaValideDateWGT) and ((Pos(EtatRev, 'MOD;PER') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNMINEURWGT' then
          begin
            V_PGI.GraphList.GetBitmap(45-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and JAiLeDroitAction(wtaMineurWGT) and ((Pos(EtatRev, 'VAL') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNMAJEURWGT' then
          begin
            V_PGI.GraphList.GetBitmap(5-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and JAiLeDroitAction(wtaMajeurWGT) and ((Pos(EtatRev, 'VAL') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNPERIMEWGT' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and JAiLeDroitAction(wtaPerimeWGT) and ((Pos(EtatRev, 'VAL;MOD') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNPERIMEDATEWGT' then
          begin
            V_PGI.GraphList.GetBitmap(88-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and JAiLeDroitAction(wtaPerimeDateWGT) and ((Pos(EtatRev, 'VAL;MOD') > 0) or MultiSelected);
          end
          { GIA }
          else if UpperCase(Name) = 'MNVALIDEGIA' then
          begin
            V_PGI.GraphList.GetBitmap(47-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and ((Pos(EtatRev, 'MOD;PER') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNVALIDEDATEGIA' then
          begin
            V_PGI.GraphList.GetBitmap(89-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and ((Pos(EtatRev, 'MOD;PER') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNMINEURGIA' then
          begin
            V_PGI.GraphList.GetBitmap(45-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and ((Pos(EtatRev, 'VAL') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNMAJEURGIA' then
          begin
            V_PGI.GraphList.GetBitmap(5-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and ((Pos(EtatRev, 'VAL') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNPERIMEGIA' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and ((Pos(EtatRev, 'VAL;MOD') > 0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNPERIMEDATEGIA' then
          begin
            V_PGI.GraphList.GetBitmap(88-1, BitMap);
            Visible := (sAction <> 'CONSULTATION') and ((Pos(EtatRev, 'VAL;MOD') > 0) or MultiSelected);
          end
          { WPC }
          else if upperCase(Name) = 'MNMISEHS' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := JAiLeDroitAction(wtaMiseHSWPC) and ((EtatParc = 'ES') or MultiSelected);
          end
          else if upperCase(Name) = 'MNCHANGEVER' then
          begin
            V_PGI.GraphList.GetBitmap(-1, BitMap);
            Visible := ((EtatParc = 'ES') and (not MultiSelected));
          end
          { WVS }
          else if upperCase(Name) = 'MNPERIMEWVS' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := JAiLeDroitAction(wtaPerimeWVS) and ((EtatVer = 'ACT') or MultiSelected);
          end
          { DIVERS }
          else if UpperCase(Name) = 'MNSEPAR' then
          begin
            Visible := not MultiSelected;
          end
          else if UpperCase(Name) = 'MNFOURNIUCF' then
          begin
            V_PGI.GraphList.GetBitmap(52-1, BitMap);
            Visible := (Pos(Fourni, '2') > 0) or MultiSelected;
          end
          else if UpperCase(Name) = 'MNNONFOURNIUCF' then
          begin
            V_PGI.GraphList.GetBitmap(53-1, BitMap);
            Visible := (Pos(Fourni, '1') > 0) or MultiSelected;
          end
          else if UpperCase(Name) = 'MNSELECTIONULS' then
          begin
            V_PGI.GraphList.GetBitmap(47-1, BitMap);
            Visible := (Pos(EtatLasst, 'DES') > 0) or MultiSelected;
          end
          else if UpperCase(Name) = 'MNDESELECTIONULS' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := (Pos(EtatLasst, 'SEL') > 0) or MultiSelected;
          end
{$IFDEF QUALITE}
          else if UpperCase(Name) = 'MNCLOTURERQN' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := JAiLeDroitAction(wtaClotureRQN) and ((Pos(EtatLIG, '0EN') >0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNCLOTURERAC' then
          begin
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := JAiLeDroitAction(wtaClotureRAC) and ((Pos(EtatLIG, 'NRE;PRE;REA') >0) or MultiSelected);
          end
          else if UpperCase(Name) = 'MNACCEPTERRQD' then
          begin
            V_PGI.GraphList.GetBitmap(37-1, BitMap);
            Visible := JAiLeDroitTag(124731);
          end
          else if UpperCase(Name) = 'MNREFUSERRQD' then
          begin
            V_PGI.GraphList.GetBitmap(35-1, BitMap);
            Visible := JAiLeDroitTag(124732);
          end
          else if UpperCase(Name) = 'MNENCOURSRQD' then
          begin
            V_PGI.GraphList.GetBitmap(29-1, BitMap);
            Visible := JAiLeDroitTag(124732);
          end
{$ENDIF QUALITE}
        end;
        if UpperCase(Name) = 'MNPARAMS' then
        begin
          V_PGI.GraphList.GetBitmap(8-1, BitMap);
        end
        else if Pos(UpperCase(Name), 'MNJOURNALACTION;MNJOURNALREVISION') > 0 then
        begin
          V_PGI.GraphList.GetBitmap(7-1, BitMap);
          Visible := not MultiSelected;
        end
        ;
        {$Ifend}
        if Visible and (Count > 0) then
        begin
          Visible := false;
          for j := 0 to Pred(Count) do
            Visible := wSetMn(Items[j]) or Visible;
        end;
        Result := Visible;
      end
    end
  end;

begin
	ChargeImageList;

  { Argument }
  sAction     := GetArgumentValue(Argument, 'ACTION');
  EtatTet     := GetArgumentValue(Argument, 'ETATTET');
  EtatLig     := GetArgumentValue(Argument, 'ETATLIG');
  EtatPhase   := GetArgumentValue(Argument, 'ETATPHASE');
  EtatBes     := GetArgumentValue(Argument, 'ETATBES');
  TypeLien    := GetArgumentValue(Argument, 'TYPELIEN');
  EtatLas     := GetArgumentValue(Argument, 'ETATLAS');
  EtatParc    := GetArgumentValue(Argument, 'ETATPARC');
  EtatVer     := GetArgumentValue(Argument, 'ETATVER');
  EtatOpe     := GetArgumentValue(Argument, 'ETATOPE');
  EtatRes     := GetArgumentValue(Argument, 'ETATRES');
  EtatLasst   := GetArgumentValue(Argument, 'ETATLASST');
  EtatSta     := GetArgumentValue(Argument, 'ETATSTA');
  EtatRev     := GetArgumentValue(Argument, 'ETATREV');
  ModeConso   := GetArgumentValue(Argument, 'MODECONSO');
  QAccSais    := GetArgumentValue(Argument, 'QACCSAIS');
  QLanSais    := GetArgumentValue(Argument, 'QLANSAIS');
  Fourni      := GetArgumentValue(Argument, 'FOURNI');
  MiseEnProd  := GetArgumentValue(Argument, 'MISEENPROD');
  {$IFDEF GPAOLIGHT}
    IsSST       := GetArgumentValue(Argument, 'TYPEORDRE') = 'STP';
    
    {  Existe des pièces de nature "CSP" (commandes de ss-traitance de phase) en cours }
    PiecesCSPEncours := (IsSST) and (GetArgumentBoolean(Argument, 'PIECESCSPENCOURS'));
  {$ENDIF GPAOLIGHT}

  { Image + Visible }
  for i := 0 to Pm.Items.Count-1 do
    wSetMn(Pm.Items[i])
end;

{ TPopupMenuFlux }

{$IFDEF STK}
constructor TPopupMenuFlux.Create(aEcran: TComponent; Mn_OnClick: TNotifyEvent; Argument: String; Flux: String = '');
var
  sql, s: string;
  i     : Integer;
  tabTM : Array of TMenuItem;
  t     : Tob;
begin
  Ecran := aEcran;
  FromPiece := (GetArgumentString(Argument, 'FROM') = 'PIECE');
  inherited Create(aEcran);

  Name := 'PMFLUX';

  sql := 'SELECT *'
       + ' FROM COMMUN'
       + ' WHERE CO_TYPE="GSF"'
{$IFDEF GPAO}
       + ' AND CO_LIBRE LIKE "%GP%"'
{$ELSE GPAO}
       + ' AND CO_LIBRE LIKE "%GC%"'
{$ENDIF GPAO}
       + ' ORDER BY CO_ABREGE'
       ;

  t := Tob.Create('COMMUN', nil, -1);
  try
    t.LoadDetailFromSQL(Sql);
    for i := 0 to Pred(t.Detail.Count) do
    begin
      s := GetArgumentValue(Argument, 'UNITE' + String(T.Detail[i].GetString('CO_CODE')));
      SetLength(tabTM, Length(tabTM) + 1);
      tabTM[Length(tabTM) - 1] := TMenuItem.Create(Self);
      tabTM[Length(tabTM) - 1].Caption := String(TraduireMemoire(T.Detail[i].GetString('CO_LIBELLE'))) + iif(s <> '', ' (' + RechDom('WUNITE', s, False) + ')', '');
      tabTM[Length(tabTM) - 1].Name := 'MN' + String(T.Detail[i].GetString('CO_CODE'));
    end;

    Items.Add(tabTM);

    SetLength(tabTM, 0);

    for i := 0 to Pred(Items.Count) do
    begin
      Items[i].RadioItem := True;
      Items[i].GroupIndex := 1;
      Items[i].OnClick := Mn_OnClick;
      if Flux <> '' then
        Items[i].Checked := Items[i].Name = 'MN' + Flux
      else
        Items[0].Checked := True;
    end;

    SetFlux(Flux);
    if Assigned(TForm(Ecran).FindComponent('BTFLUX')) then
    begin
      if Ecran is TFMul then
      begin
        TToolBarButton97(TFMul(Ecran).LaTOF.GetControl('BTFLUX')).DropdownMenu := Self;
        TToolBarButton97(TFMul(Ecran).LaTOF.GetControl('BTFLUX')).OnClick := Self.OnClick;
      end
      else if Ecran is TFFiche then
      begin
        TToolBarButton97(TFFiche(Ecran).OM.GetControl('BTFLUX')).DropDownMenu := Self;
        TToolBarButton97(TFFiche(Ecran).OM.GetControl('BTFLUX')).OnClick := Self.OnClick;
      end
      else if Ecran is TFVierge then
      begin
        TToolBarButton97(TFVierge(Ecran).FindComponent('BTFLUX')).DropDownMenu := Self;
        TToolBarButton97(TFVierge(Ecran).FindComponent('BTFLUX')).OnClick := Self.OnClick;
      end;
    end;

  finally
    FLastPmFlux := VH_GC.PmFLux;
    VH_GC.PmFLux := Self;
    t.Free;
  end;
end;
{$ENDIF STK}

{$IFDEF STK}
destructor TPopupMenuFlux.Destroy;
begin
  inherited;

  VH_GC.PmFLux := FLastPmFlux;
end;
{$ENDIF STK}

{$IFDEF STK}
function TPopupMenuFlux.GetFlux: String;
var
  i: Integer;
begin
  Result := '';
  i := 0;
  while (Result = '') and (i < Items.Count) do
  begin
    if Items[i].Checked then
      Result := StringReplace(Items[i].Name, 'MN', '', [rfIgnoreCase]);
    Inc(i);
  end;
end;
{$ENDIF STK}

{$IFDEF STK}
function TPopupMenuFlux.GetFluxFromSender(Sender: TObject): String;
begin
  if Sender is TMenuItem then
    Result := StringReplace(TMenuItem(Sender).Name, 'MN', '', [rfIgnoreCase])
  else
    Result := '';
end;
{$ENDIF STK}

{$IFDEF STK}
procedure TPopupMenuFlux.OnClick(Sender: TObject);
var
  i: Integer;
  Find: Boolean;
begin
  inherited;
  { Désactive le Flux Achat si pas accessible }
  i := -1; Find := False;
  while (i < Items.Count -1) and not Find do
  begin
    Inc(i);
    Find := (Items[i].Name = 'MNACH');
  end;
  if Find then
    Items[i].Enabled := CanShowAchat;
end;
{$ENDIF STK}

{$IFDEF STK}
procedure TPopupMenuFlux.SetFlux(Flux: String);
var
  i: Integer;
begin
  {$IFNDEF GPAO}
  if Flux = 'PRO' then Flux := 'STO';
  if Flux = 'CON' then Flux := 'STO';
  {$ENDIF !GPAO}

  for i := 0 to Pred(Items.Count) do
    Items[i].Checked := Items[i].Name = 'MN' + Flux;
end;
{$ENDIF STK}

procedure wSetMnuUtilitaire(Const MultiSelected: Boolean; Const Argument: string; Const Pm: TPopUpMenu);
var
{$IFDEF STK}
  Contremarque: Boolean;
  TenueStock  : Boolean;
  DateCloture :tDateTime;
{$ENDIF STK}
{$IFDEF GPAO}
  OrdrePere                                       : Integer;
  PeriodeSauvOK																		: Boolean;
{$ENDIF GPAO}
	i                                               : Integer;
  EtatBes, EtatMvt, EtatPhase, EtatLig, ModeConso : string;
  Depot, Article, TypeOrdre, TypeLien             : String;
  StatutDispo, StatutFlux, RefAffectation         : string;
  {$IFDEF EDI}
    IdentifiantELI: Integer;
    NaturePieceG: String;
    ValuesELI: MyArrayValue;
  {$ENDIF EDI}
  QualifMvt: string;

{$IFDEF STK}
  function GetCleGQ: tCleGQ;
  begin
    Result.Depot       := Depot;
    Result.Article     := Article;
  end;
{$ENDIF STK}

  function wSetMn(Const Mn: TMenuItem): boolean;
  var
    j: Integer;
  begin
    Result := true;
    if Assigned(Mn) then
    begin
      with Mn do
      begin
        if UpperCase(Name) = UpperCase('UtRattacheWOL') then
        begin
          V_PGI.GraphList.GetBitmap(18-1, BitMap);
          Visible := MultiSelected or ((TypeOrdre <> 'STA') and (TypeOrdre <> 'VTE'))
        end
        else if UpperCase(Name) = Uppercase('UtSubDiviserWOL') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(51-1, BitMap);
            Visible := (not MultiSelected)  and (Pos(EtatLig, 'DCL;DEC') > 0) and
                       (TypeOrdre <> 'STA') and (TypeOrdre <> 'VTE')
                       and (OrdrePere = 0)
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
{$IFDEF STK}
        else if UpperCase(Name) = UpperCase('UtAffectWOL') then
        begin
          Visible := false;
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(49-1, BitMap);
//GP_20080125_PCO_GP14733 deb
            Visible := JAiLeDroitTag(-32240)
                       and (Pos(EtatLig, 'DCL;DEC;LAN;REC;TER;TER') > 0);
//GP_20080125_PCO_GP14733 fin
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('UtCtrlBTechWOL') then
        begin
          Visible := false;
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(60-1, BitMap);
            Visible := Multiselected or (Pos(EtatLig, 'DEC') > 0);
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('UtCompareWNLWOB') then
        begin
          Visible := false;
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(66-1, BitMap);
            Visible := (not Multiselected) and (Pos(EtatLig, 'DEC;LAN;REC;TER;SOL') > 0);
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('UtCompareWGLWOG') then
        begin
          Visible := false;
          {$IFDEF GPAO GPAO}
            V_PGI.GraphList.GetBitmap(66-1, BitMap);
            Visible := (not Multiselected) and (Pos(EtatLig, 'DEC;LAN;REC;TER;SOL') > 0);
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('UtBilanWOL') then
        begin
          V_PGI.GraphList.GetBitmap(69-1, BitMap);
          Visible := (Pos(EtatLig, 'DEC;LAN;REC;TER;SOL') > 0);
        end
        else if UpperCase(Name) = UpperCase('UtAffectWOP') then
        begin
          Visible := false;
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(49-1, BitMap);
//GP_20080125_PCO_GP14733 deb
            Visible := JAiLeDroitTag(-32240)
                       and (Pos(EtatPhase, 'DEC;LAN;REC;TER') > 0);
//GP_20080125_PCO_GP14733 fin
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('UtChgDatePeremptionGSL') then
        begin
          V_PGI.GraphList.GetBitmap(95-1, BitMap);
          Visible := StkGereDatePeremption;
        end
        else if UpperCase(Name) = UpperCase('UtAffectStkWOB') then
//GP_20080125_PCO_GP14733 deb
        begin
          V_PGI.GraphList.GetBitmap(49-1, BitMap);
          Visible := ((JAiLeDroitTag(-32240)
                     and  (TypeLien = 'COM')
                     and (Pos(ModeConso, 'LAS;LAN;REC'    ) > 0)
                     and (Pos(EtatBes  , 'BES;CON;TER;RUP') > 0))
                     or (JAiLeDroitTag(-32240) and MultiSelected))
        end
        else if UpperCase(Name) = UpperCase('UtAnnuleAffectStkWOB') then
        begin
          V_PGI.GraphList.GetBitmap(47-1, BitMap);
          Visible := ((JAiLeDroitTag(-32240)
                     and (TypeLien = 'COM')
                     and (Pos(ModeConso, 'LAS;LAN;REC'    ) > 0)
                     and (Pos(EtatBes  , 'BES;CON;TER;RUP') > 0))
                     or (JAiLeDroitTag(-32240) and MultiSelected))
        end
//GP_20080125_PCO_GP14733 fin
        else if UpperCase(Name) = UpperCase('UtRuptureWOL') then
        begin
          Visible := false;
          {$IFDEF GPAOLIGHT}
            V_PGI.GraphList.GetBitmap(48-1, BitMap);
            Visible := (not MultiSelected) and (Pos(EtatLig, 'LAN;REC;TER') > 0);
          {$ENDIF GPAOLIGHT}
        end
        else if UpperCase(Name) = UpperCase('UtRuptureWOP') then
        begin
          Visible := false;
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(48-1, BitMap);
            Visible := (not MultiSelected) and (Pos(EtatPhase, 'LAN;REC;TER') > 0);
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('UtRuptureWOB') then
        begin
          Visible := false;
          {$IFDEF GPAOLIGHT}
            V_PGI.GraphList.GetBitmap(48-1, BitMap);
            Visible := (not MultiSelected)
                   and (TypeLien = 'COM')
                   and (Pos(EtatBes, 'BES;CON;RUP;TER') > 0)
          {$ENDIF GPAOLIGHT}
        end
        else if UpperCase(Name) = UpperCase('UtConsoRupAutoGSM') then
        begin
          Visible := false;
          {$IFDEF GPAOLIGHT}
            V_PGI.GraphList.GetBitmap(50-1, BitMap);
            Visible := MultiSelected or (EtatMvt = 'FER');
          {$ENDIF GPAOLIGHT}
        end
        else if UpperCase(Name) = UpperCase('UtConsoRupManuGSM') then
        begin
          Visible := false;
          {$IFDEF GPAOLIGHT}
            V_PGI.GraphList.GetBitmap(53-1, BitMap);
            Visible := (not MultiSelected) and (EtatMvt = 'FER');
          {$ENDIF GPAOLIGHT}
        end
        else if UpperCase(Name) = UpperCase('UtSoldeRupGSM') then
        begin
          Visible := false;
          {$IFDEF GPAOLIGHT}
            V_PGI.GraphList.GetBitmap(46-1, BitMap);
            Visible := MultiSelected or (EtatMvt = 'FER');
          {$ENDIF GPAOLIGHT}
        end
        else if UpperCase(Name) = UpperCase('UtBlocageGQ') then
        begin
          V_PGI.GraphList.GetBitmap(68-1, BitMap);
          Visible := JAiLeDroitTag(32231) and StkGereStatutDispo;
        end
        else if UpperCase(Name) = UpperCase('UtDeBlocageGQ') then
        begin
          V_PGI.GraphList.GetBitmap(47-1, BitMap);
          Visible := JAiLeDroitTag(32232) and StkGereStatutDispo;
        end
        else if UpperCase(Name) = UpperCase('UtEntreeGQ') then
        begin
          V_PGI.GraphList.GetBitmap(52-1, BitMap);
          Visible := JAiLeDroitTag(32221);
        end
        else if UpperCase(Name) = UpperCase('UtSortieGQ') then
        begin
          V_PGI.GraphList.GetBitmap(53-1, BitMap);
          Visible := JAiLeDroitTag(32222)
        end
        else if UpperCase(Name) = UpperCase('UtJustifGQ') then
        begin
          V_PGI.GraphList.GetBitmap(8-1, BitMap);
          Visible := VH_GC.GCIfDefCEGID and JAiLeDroitTag(32204) and GereSerie(Article);
        end
        else if UpperCase(Name) = UpperCase('UtChangement') then
        begin
          V_PGI.GraphList.GetBitmap(51-1, BitMap);
        end
        else if UpperCase(Name) = UpperCase('UtChgEmplaceGQ') then
        begin
          Visible := GereEmplacement(GetCleGQ)
                 and JAiLeDroitTag(32223)
        end
        else if UpperCase(Name) = UpperCase('UtChgIndiceArticleGQ') then
        begin
          Visible := GereIndiceArticle(Article)
        end
{$IFDEF GPAO}
        else if UpperCase(Name) = UpperCase('UtChgMarqueGQ') then
        begin
          Visible := GereMarque(Article)
        end
        else if UpperCase(Name) = UpperCase('UtChgChoixQualiteGQ') then
        begin
          Visible := GereChoixQualite(Article)
        end
{$ENDIF GPAO}
        else if UpperCase(Name) = UpperCase('UtChgEmplaceGQD') then
        begin
          Visible := GereEmplacement(GetCleGQ)
                 and JAiLeDroitTag(32223)
                 and ((Pos(StatutDispo, GetFieldFRomGSN('GSN_SDISPOPICKING', 'CEM')) > 0) or (GetFieldFRomGSN('GSN_SDISPOPICKING', 'CEM') = ''))
                 and ((Pos(StatutFlux , GetFieldFromGSN('GSN_SFLUXPICKING' , 'CEM')) > 0) or (GetFieldFromGSN('GSN_SFLUXPICKING' , 'CEM') = ''))
        end
        else if UpperCase(Name) = UpperCase('UtChgIndiceArticleGQD') then
        begin
          Visible := GereIndiceArticle(Article)
                 and ((Pos(StatutDispo, GetFieldFRomGSN('GSN_SDISPOPICKING', 'CIA')) > 0) or (GetFieldFRomGSN('GSN_SDISPOPICKING', 'CIA') = ''))
                 and ((Pos(StatutFlux , GetFieldFromGSN('GSN_SFLUXPICKING' , 'CIA')) > 0) or (GetFieldFromGSN('GSN_SFLUXPICKING' , 'CIA') = ''))
        end
{$IFDEF GPAO}
        else if UpperCase(Name) = UpperCase('UtChgMarqueGQD') then
        begin
          Visible := GereMarque(Article)
                 and ((Pos(StatutDispo, GetFieldFRomGSN('GSN_SDISPOPICKING', 'CMQ')) > 0) or (GetFieldFRomGSN('GSN_SDISPOPICKING', 'CMQ') = ''))
                 and ((Pos(StatutFlux , GetFieldFromGSN('GSN_SFLUXPICKING' , 'CMQ')) > 0) or (GetFieldFromGSN('GSN_SFLUXPICKING' , 'CMQ') = ''))
        end
        else if UpperCase(Name) = UpperCase('UtChgChoixQualiteGQD') then
        begin
          Visible := GereChoixQualite(Article)
                 and ((Pos(StatutDispo, GetFieldFRomGSN('GSN_SDISPOPICKING', 'CCQ')) > 0) or (GetFieldFRomGSN('GSN_SDISPOPICKING', 'CCQ') = ''))
                 and ((Pos(StatutFlux , GetFieldFromGSN('GSN_SFLUXPICKING' , 'CCQ')) > 0) or (GetFieldFromGSN('GSN_SFLUXPICKING' , 'CCQ') = ''))
        end
{$ENDIF GPAO}
        else if UpperCase(Name) = UpperCase('UtBlocageGQD') then
        begin
          V_PGI.GraphList.GetBitmap(68-1, BitMap);
          Visible := StkGereStatutDispo
                 and JAiLeDroitTag(32231)
                 and ((Pos(StatutDispo, GetFieldFRomGSN('GSN_SDISPOPICKING', 'CBS')) > 0) or (GetFieldFRomGSN('GSN_SDISPOPICKING', 'CBS') = ''))
                 and ((Pos(StatutFlux , GetFieldFromGSN('GSN_SFLUXPICKING' , 'CBS')) > 0) or (GetFieldFromGSN('GSN_SFLUXPICKING' , 'CBS') = ''))
        end
        else if UpperCase(Name) = UpperCase('UtDeBlocageGQD') then
        begin
          V_PGI.GraphList.GetBitmap(47-1, BitMap);
          Visible := StkGereStatutDispo
                 and JAiLeDroitTag(32232)
                 and ((Pos(StatutDispo, GetFieldFRomGSN('GSN_SDISPOPICKING', 'CDS')) > 0) or (GetFieldFRomGSN('GSN_SDISPOPICKING', 'CDS') = ''))
                 and ((Pos(StatutFlux , GetFieldFromGSN('GSN_SFLUXPICKING' , 'CDS')) > 0) or (GetFieldFromGSN('GSN_SFLUXPICKING' , 'CDS') = ''))
        end
        else if UpperCase(Name) = UpperCase('UtEntreeGQD') then
        begin
          V_PGI.GraphList.GetBitmap(52-1, BitMap);
          Visible := JAiLeDroitTag(32221)
                 and (Pos(StatutDispo, 'LBR;BLQ') > 0);
        end
        else if UpperCase(Name) = UpperCase('UtSortieGQD') then
        begin
          V_PGI.GraphList.GetBitmap(53-1, BitMap);
          Visible := JAiLeDroitTag(32222)
                 and ((Pos(StatutDispo, GetFieldFRomGSN('GSN_SDISPOPICKING', 'SEX')) > 0) or (GetFieldFRomGSN('GSN_SDISPOPICKING', 'SEX') = ''))
                 and ((Pos(StatutFlux , GetFieldFromGSN('GSN_SFLUXPICKING' , 'SEX')) > 0) or (GetFieldFromGSN('GSN_SFLUXPICKING' , 'SEX') = ''))
        end
        else if UpperCase(Name) = UpperCase('UtAffectGQD') then
        begin
          V_PGI.GraphList.GetBitmap(49-1, BitMap);
          Visible := JAiLeDroitTag(-32240)
                 and ((Pos(StatutDispo, GetFieldFromGSN('GSN_SDISPOPICKING', 'CAF')) > 0) or (GetFieldFromGSN('GSN_SDISPOPICKING', 'CAF') = ''))
                 and ((Pos(StatutFlux , GetFieldFromGSN('GSN_SFLUXPICKING' , 'CAF')) > 0) or (GetFieldFromGSN('GSN_SFLUXPICKING' , 'CAF') = ''))
        end
        else if UpperCase(Name) = UpperCase('UtAffectGQD_T') then
        begin
//GP_20080125_PCO_GP14733 deb
          Visible := JAiLeDroitTag(-32240)
                     and ((not Contremarque) or (StkFinesseContremarque = 'CLI'));
        end
        else if UpperCase(Name) = UpperCase('UtAffectGQD_GP') then
        begin
          Visible := JAiLeDroitTag(-32240)
                     and ((not Contremarque) or (StkFinesseContremarque = 'PIE'));
        end
        else if UpperCase(Name) = UpperCase('UtAffectGQD_GL') then
        begin
          Visible := JAiLeDroitTag(-32240)
                     and ((not Contremarque) or (StkFinesseContremarque = 'LIG'));
//GP_20080125_PCO_GP14733 fin
        end
        else if UpperCase(Name) = UpperCase('UtDesAffectGQD') then
        begin
          V_PGI.GraphList.GetBitmap(47-1, BitMap);
          Visible := JAiLeDroitTag(32249)
                and (RefAffectation <> '')
                and ((Pos(StatutFlux, GetFieldFromGSN('GSN_SFLUXPICKING', 'CAF')) > 0) or (GetFieldFromGSN('GSN_SFLUXPICKING', 'CAF') = ''))
        end
        else if UpperCase(Name) = UpperCase('UtAffectGQD_Prod') then
//GP_20080125_PCO_GP14733 deb
        begin
          Visible := false;
          {$IFDEF GPAO}
            Visible := JAiLeDroitTag(-32240)
                       and (Not Contremarque);
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('UtAffectGQD_Aff') then
        begin
          Visible := JAiLeDroitTag(-32240)
                     and (not Contremarque);
        end
//GP_20080125_PCO_GP14733 fin
        else if UpperCase(Name) = UpperCase('UtAffectGQ') then
        begin
          V_PGI.GraphList.GetBitmap(49-1, BitMap);
          Visible := JAiLeDroitTag(-32240);
        end
        else if UpperCase(Name) = UpperCase('UtAffectGQ_T') then
        begin
//GP_20080125_PCO_GP14733 deb
          Visible := JAiLeDroitTag(-32240)
                     and ((not Contremarque) or (StkFinesseContremarque = 'CLI'));
        end
        else if UpperCase(Name) = UpperCase('UtAffectGQ_GP') then
        begin
          Visible := JAiLeDroitTag(-32240)
                     and ((not Contremarque) or (StkFinesseContremarque = 'PIE'));
        end
        else if UpperCase(Name) = UpperCase('UtAffectGQ_GL') then
        begin
          Visible := JAiLeDroitTag(-32240)
                     and ((not Contremarque) or (StkFinesseContremarque = 'LIG'));
        end
        else if UpperCase(Name) = UpperCase('UtAffectGQ_Prod') then
        begin
          Visible := false;
          {$IFDEF GPAO}
            Visible := JAiLeDroitTag(-32240)
                       and (not Contremarque);
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('UtAffectGQ_Aff') then
        begin
          Visible := JAiLeDroitTag(-32240)
                     and (not Contremarque);
        end
//GP_20080125_PCO_GP14733 fin
        else if UpperCase(Name) = UpperCase('UtArchivageGQ') then
        begin
          V_PGI.GraphList.GetBitmap(98-1, BitMap);
          Visible := JAiLeDroitTag(32203);
        end
        else if Pos(UpperCase('UtRevalorisation'), UpperCase(Name)) > 0 then
        begin
          V_PGI.GraphList.GetBitmap(77-1, BitMap);
          // GC_20071207_DM_GC12431_DEBUT
//          Visible := JAiLeDroitTag(60405);
          Visible := JAiLeDroitTag(60405) and (VH_GC.GCAchatStockSeria);
          // GC_20071205_DM_GC12431_FIN
        end
        else if (Pos(UpperCase('UtModifPrixArchiveFromGQ'), UpperCase(Name)) > 0)  then
        begin
          V_PGI.GraphList.GetBitmap(90-1, BitMap);
          Visible := JAiLeDroitTag(60405) and (DateCloture <> iDate1900);
        end
        else if (Pos(UpperCase('UtModifPrixArchiveFromGSM'), UpperCase(Name)) > 0)  then
        begin
          V_PGI.GraphList.GetBitmap(90-1, BitMap);
          Visible := JAiLeDroitTag(60405) and (DateCloture <> iDate1900) and (QualifMvt='CLO') and (EtatMvt<>'ARC');
        end
        else if (Pos(UpperCase('UtModifPrixMvtFromGSM'), UpperCase(Name)) > 0)  then
        begin
          V_PGI.GraphList.GetBitmap(90-1, BitMap);
          Visible := JAiLeDroitTag(60405) and (DateCloture <> iDate1900) and (pos(QualifMvt,'EEX/EIN')>0) and (EtatMvt<>'ARC');
        end
        else if Pos(UpperCase('UtRecalculCompteurs'), UpperCase(Name)) > 0 then
        begin
          V_PGI.GraphList.GetBitmap(91-1, BitMap);
          // GC_20071205_DM_GC12431_DEBUT
//          Visible := TenueStock and JAiLeDroitTag(60402);
          Visible := TenueStock and JAiLeDroitTag(60402) and (VH_GC.GCAchatStockSeria);
          // GC_20071205_DM_GC12431_FIN
        end
{$ENDIF STK}
        else if UpperCase(Name) = UpperCase('UtCompareWNT') then
        begin
          V_PGI.GraphList.GetBitmap(66-1, BitMap);
        end
{$IFDEF GPAO}
        else if UpperCase(Name) = UpperCase('IxaGenereWND') then
        begin
          V_PGI.GraphList.GetBitmap(50-1, BitMap);
          Visible := wGereConfigurator;
        end
        else if UpperCase(Name) = UpperCase('IxaReGenereWND') then
        begin
          V_PGI.GraphList.GetBitmap(51-1, BitMap);
          Visible := wGereConfigurator;
        end
{$ENDIF GPAO}        
        else if UpperCase(Name) = UpperCase('UtCompareWPL') then
        begin
          V_PGI.GraphList.GetBitmap(66-1, BitMap);
        end
        {$IFDEF GPAO}
	        else if UpperCase(Name) = UpperCase('utCalculWPL') then
  	      begin
    	      V_PGI.GraphList.GetBitmap(99-1, BitMap);
     		    Visible := PeriodeSauvOK;
        	end
        {$ENDIF GPAO}
        else if UpperCase(Name) = UpperCase('UtCompareWGT') then
        begin
          V_PGI.GraphList.GetBitmap(66-1, BitMap);
        end
        else if UpperCase(Name) = UpperCase('MnRecalculTConstWOG') then
        begin
          V_PGI.GraphList.GetBitmap(91-1, BitMap);
        end
        else if UpperCase(Name) = UpperCase('MnUtConfigurator') then
        begin
          {$IFDEF GPAO}
            Visible := wGereConfigurator and wExistConfigurator;
            if Visible then
              OnClick := Vh_Gp.Configurators.LastCfx.CallCfx
          {$ELSE  GPAO}
            Visible := False;
          {$ENDIF GPAO}
        end
        else if (UpperCase(Name)= UpperCase('MnUtDispatchEDI')) then
        begin
          {$IFDEF EDI}
            Visible := (not MultiSelected);
            if Visible then
            begin
              SetLength(ValuesELI, 0);
              ValuesELI := wGetSQLFieldsValues(['ELI_ARTICLE', 'ELI_DEPOT'], 'EDILIGNE', WhereELI('', IdentifiantELI));
              if Length(ValuesELI) = 2 then
              begin
                Article := ValuesELI[0];
                Depot   := ValuesELI[1];
                Visible := (Article <> '') and GereDetail(GetCleGQ());
              end
              else
                Visible := False;
              if Visible then
                Caption := TitreDetailStock(NaturePieceG)
            end;
          {$ELSE EDI}
            Visible := False;
          {$ENDIF EDI}
        end
        else if (UpperCase(Name)= UpperCase('UtNonConf')) then
        begin
          {$IFDEF QUALITE}
            Visible := (not MultiSelected) and JAiLeDroitTag(124110);
          {$ELSE}
            Visible := False;
          {$ENDIF QUALITE}
        end
        else if UpperCase(Name) = UpperCase('UtChangeCircuit') then
          Visible := {$IFDEF GPAO} MultiSelected {$ELSE  GPAO} False {$ENDIF GPAO}
        ;

        if Visible and (Count > 0) then
        begin
          Visible := false;
          for j := 0 to Pred(Count) do
            Visible := wSetMn(Items[j]) or Visible;
        end;
        Result := Visible;
      end
    end
  end;

begin
  { Récupère les valeurs de l'argument }
  QualifMvt      := GetArgumentString(Argument, 'QUALIFMVT');
  EtatLig        := GetArgumentString(Argument, 'ETATLIG');
  EtatPhase      := GetArgumentString(Argument, 'ETATPHASE');
  EtatBes        := GetArgumentString(Argument, 'ETATBES');
  EtatMvt        := GetArgumentString(Argument, 'ETATMVT');
  ModeConso      := GetArgumentString(Argument, 'MODECONSO');
  Depot          := GetArgumentString(Argument, 'DEPOT');
  Article        := GetArgumentString(Argument, 'ARTICLE');
  StatutDispo    := GetArgumentString(Argument, 'STATUTDISPO');
  StatutFlux     := GetArgumentString(Argument, 'STATUTFLUX');
  RefAffectation := GetArgumentString(Argument, 'REFAFFECTATION');
  TypeOrdre      := GetArgumentString(Argument, 'TYPEORDRE');
  TypeLien       := GetArgumentString(Argument, 'TYPELIEN');
  {$IFDEF STK}
    Contremarque := GetArgumentBoolean(Argument, 'CONTREMARQUE');
    TenueStock   := GetArgumentBoolean(Argument, 'TENUESTOCK') or (Pos('TENUESTOCK', Argument) <= 0);
    DateCloture  := GetArgumentDateTime(Argument, 'DATECLOTURE');
  {$ENDIF STK}
  {$IFDEF GPAO}
    OrdrePere   := GetArgumentInteger (Argument, 'ORDREPERE');
    PeriodeSauvOK:= GetArgumentBoolean(Argument, 'PERIODESAUVOK');
  {$ENDIF GPAO}
  {$IFDEF EDI}
    NaturePieceG   := GetArgumentString (Argument, 'NATUREPIECEG'  );
    IdentifiantELI := GetArgumentInteger(Argument, 'IDENTIFIANTELI');
  {$ENDIF EDI}

	ChargeImageList;
	for i := 0 to Pm.Items.Count-1 do
  begin
    wSetMn(Pm.Items[i]);
  end;
end;

procedure wSetMnuAnalyse(Const MultiSelected: Boolean; Const Argument: string; Const Pm: TPopUpMenu);
var
  i: integer;
  TypeLien, ModeConso: string;
  
  function wSetMn(Const Mn: TMenuItem): boolean;
  var
    j: Integer;
  begin
    Result := true;
    if Assigned(Mn) then
    begin
      with Mn do
      begin
        if UpperCase(Name) = UpperCase('AnCubeGQ') then
        begin
          V_PGI.GraphList.GetBitmap(58-1, BitMap);
          Visible := true;
        end
        else if UpperCase(Name) = UpperCase('AnCubeGQD') then
        begin
          V_PGI.GraphList.GetBitmap(58-1, BitMap);
          Visible := true;
        end
        else if UpperCase(Name) = UpperCase('AnCubeGSM') then
        begin
          V_PGI.GraphList.GetBitmap(58-1, BitMap);
          Visible := true;
        end
        else if UpperCase(Name) = UpperCase('AnCubeWOL') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(58-1, BitMap);
            Visible := true;
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('AnCubeSPRWOL') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(53-1, BitMap);
            Visible := (not MultiSelected) and (JaiLeDroitConceptStock);
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
          Enabled := visible;
        end
        else if UpperCase(Name) = UpperCase('AnCubeWOP') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(58-1, BitMap);
            Visible := true;
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('AnCubeWOB') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(58-1, BitMap);
            Visible := (not MultiSelected) and (TypeLien = 'COM') and (ModeConso <> 'NON');
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('AnCubeWOG') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(58-1, BitMap);
            Visible := true;
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('AnCubeWOR') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(58-1, BitMap);
            Visible := true;
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('AnAvanceWOP') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(58-1, BitMap);
            Visible := (not MultiSelected) ;
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('AnAvanceWOG') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(58-1, BitMap);
            Visible := (not MultiSelected);
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        ;

        if Visible and (Count > 0) then
        begin
          Visible := false;
          for j := 0 to Pred(Count) do
            Visible := wSetMn(Items[j]) or Visible;
        end;
        Result := Visible;

      end;
    end;
  end;

begin
  { Récupère les valeurs de l'argument }
  TypeLien  := GetArgumentValue(Argument, 'TYPELIEN');
  ModeConso := GetArgumentValue(Argument, 'MODECONSO');

	ChargeImageList;
	for i := 0 to Pm.Items.Count-1 do
  begin
    wSetMn(Pm.Items[i]);
  end;
end;

procedure wSetMnuBaseTech (Const MultiSelected: Boolean; Const Argument: string; Const Pm: TPopUpMenu);
var
  i      : integer;
  EtatLig: string;

  function wSetMn(Const Mn: TMenuItem): boolean;
  var
    j: Integer;
  begin
    Result := true;
    if Assigned(Mn) then
    begin
      with Mn do
      begin
        if UpperCase(Name) = UpperCase('MnBtCircuit') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(61-1, BitMap);
            Visible := (not MultiSelected) and (EtatLig <> 'DCL');
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('MnBtBesoins') then
        begin
          V_PGI.GraphList.GetBitmap(78-1, BitMap);
          Visible := (not MultiSelected) and (EtatLig <> 'DCL');
        end
        else if UpperCase(Name) = UpperCase('MnBtCoproduits') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(24-1, BitMap);
            Visible := (not MultiSelected) and PrdGereCoProduction and (EtatLig <> 'DCL');
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if (UpperCase(Name) = UpperCase('MnBtGamme')) or (UpperCase(Name) = UpperCase('MnBtGammeFSL')) then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(79-1, BitMap);
            Visible := (not MultiSelected) and (EtatLig <> 'DCL');
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end
        else if UpperCase(Name) = UpperCase('MnBtBac') then
        begin
          {$IFDEF GPAO}
            V_PGI.GraphList.GetBitmap(71-1, BitMap);
            Visible := (not MultiSelected) and (EtatLig <> 'DCL') and (EtatLig <> 'DEC') and  ( (pos('BC',uppercase(GetParamMESSuiviOP))>0) or (GetParamMESSuiviOP='') )     ;
          {$ELSE GPAO}
            Visible := false;
          {$ENDIF GPAO}
        end;
        ;

        if Visible and (Count > 0) then
        begin
          Visible := false;
          for j := 0 to Pred(Count) do
            Visible := wSetMn(Items[j]) or Visible;
        end;
        Result := Visible;

      end;
    end;
  end;

begin
  { Récupère les valeurs de l'argument }
  EtatLig := GetArgumentString (Argument, 'ETATLIG');

	ChargeImageList;
	for i := 0 to Pm.Items.Count-1 do
  begin
    wSetMn(Pm.Items[i]);
  end;
end;

{$IFDEF GCGC}
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. : Permet de créer dynamiquement des sous menus du popup loupe des articles
Suite ........ : vers d'éventuelles nomenclature et gammes
Mots clefs ... :
*****************************************************************}
procedure wSetItemsMnuLoupeArticle(Const PmLoupe: TPopupMenu; Const LpNAT_OnClick: TNotifyEvent; Const Article: String; Const Sender: TForm);
var
  i       : Integer;
  sSql    : String;
  TobWAN  : Tob;
  TabTM   : Array[1..2] of Array[1..2] of Array of TMenuItem;
  GereParc: Boolean;

  function GetMn : String;
  begin
    if Pos('WARTICLE_FIC', Sender.Name) > 0 then
      Result := ''
    else
      Result := 'Mn'
  end;

  procedure GereIsoflex(MenuIso: TMenuItem);
  begin
    if MenuIso <> nil then MenuIso.Visible := AglIsoflexPresent;
  end;

begin
  with Sender do
  begin
    if Assigned(FindComponent(GetMn + 'LpVisuNome_PRD')) then
      for i := TMenuItem(FindComponent(GetMn + 'LpVisuNome_PRD')).Count - 1 downto 0 do
        TMenuItem(FindComponent(GetMn + 'LpVisuNome_PRD')).Items[i].Free;
    if Assigned(FindComponent(GetMn + 'LpVisuGamme_PRD')) then
      for i := TMenuItem(FindComponent(GetMn + 'LpVisuGamme_PRD')).Count - 1 downto 0 do
        TMenuItem(FindComponent(GetMn + 'LpVisuGamme_PRD')).Items[i].Free;
    if Assigned(FindComponent(GetMn + 'LpVisuNome_SAV')) then
      for i := TMenuItem(FindComponent(GetMn + 'LpVisuNome_SAV')).Count - 1 downto 0 do
        TMenuItem(FindComponent(GetMn + 'LpVisuNome_SAV')).Items[i].Free;
    if Assigned(FindComponent(GetMn + 'LpVisuGamme_SAV')) then
      for i := TMenuItem(FindComponent(GetMn + 'LpVisuGamme_SAV')).Count - 1 downto 0 do
        TMenuItem(FindComponent(GetMn + 'LpVisuGamme_SAV')).Items[i].Free;
    if Assigned(FindComponent(GetMn + 'SGED')) then
      GereIsoFlex(tMenuItem(FindComponent(GetMn + 'SGED')));

    GereParc := StrToBool_ (wGetFieldFromGA('GA_GEREPARC', Article));
    if Assigned(FindComponent(GetMn + 'lpParc')) then
      TMenuItem(FindComponent(GetMn + 'lpParc')).Visible := GereParc;

    {$IFDEF STK}
      { Indice article }
      if Assigned(FindComponent(GetMn + 'LpIndiceArticle')) and (not StkGereIndiceArticle) then
        TMenuItem(FindComponent(GetMn + 'LpIndiceArticle')).Free;
    {$ENDIF STK}

    TobWAN := Tob.Create('WAN', nil, -1);
    try
      sSql := 'SELECT WAN_NATURETRAVAIL, WNA_LIBELLE'
            + ' FROM WARTNAT'
            + ' LEFT JOIN WNATURETRAVAIL ON (WNA_NATURETRAVAIL=WAN_NATURETRAVAIL)'
            + ' WHERE WAN_ARTICLE="' + Article + '"'
            + ' ORDER BY WAN_NATURETRAVAIL'
            ;
      if TobWAN.LoadDetailDBFromSql('WAN', sSql) then
      begin
        if TobWAN.Detail.Count > 1 then
        begin
          if wGetFieldFromGA('GA_GEREPARC', Article) = wTrue then
          begin
            SetLength(TabTM[1, 1], 0);
            for i := 0 to Pred(TobWAN.Detail.Count) do
            begin
              SetLength(TabTM[1, 1], i + 1);
              TabTM[1, 1, i]         := TMenuItem.Create(PmLoupe);
              TabTM[1, 1, i].Name    := 'Mn' + TobWAN.Detail[i].GetString('WAN_NATURETRAVAIL');
              TabTM[1, 1, i].Caption := TobWAN.Detail[i].GetString('WNA_LIBELLE');
              TabTM[1, 1, i].OnClick := LpNAT_OnClick
            end
          end
        end;
        TMenuItem(FindComponent(GetMn + 'LpVisuNome_PRD')).Add(TabTM[1, 1]);
        {$IFDEF GPAO}
          TMenuItem(FindComponent(GetMn + 'LpVisuGamme_PRD')).Add(TabTM[1, 1]);
        {$ENDIF GPAO}
      end
    finally
      TobWAN.Free
    end
  end;

end;
{$ENDIF GCGC}

{ Permet de savoir si le Popup menu est "vide" i.e. tous les items sont invisibles }
function wIsEmptyPopup(Popup: TPopUpMenu): Boolean;
var
  i: Integer;
begin
  i := 0;
  Result := True;
  while Result and (i < Popup.Items.Count) do
  begin
    if (Popup.Items[i].Caption <> '-') then
      Result := not Popup.Items[i].Visible;
    Inc(i)
  end
end;

{$IFDEF STK}
function TPopupMenuFlux.GetCanShowAchat: Boolean;
var
  Article: hString;
begin
  if (FFluxCanShowAchat = sfaNone) then
  begin
    Article := UtilArticle.GetArticle(Ecran);
    if (Article <> '') and not FromPiece and ArticleWithFicheCatalogu(Article) then
      FFluxCanShowAchat := sfaNo
    else
      FFluxCanShowAchat := sfaYes;
  end;
  Result := not (FFluxCanShowAchat = sfaNo);
end;
{$ENDIF STK}

{$IFDEF STK}
procedure TPopupMenuFlux.SetFromPiece(const Value: Boolean);
begin
  if (Value <> FFromPiece) then
  begin
    FFluxCanShowAchat := sfaNone;
    FFromPiece := Value;
  end;
end;
{$ENDIF STK}

end.
