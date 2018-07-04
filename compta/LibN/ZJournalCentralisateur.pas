{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :
Modifié le ... : 15/03/2006
Description .. :
Mots clefs ... :
*****************************************************************}
unit ZJournalCentralisateur;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, ExtCtrls, ComCtrls, ImgList, ToolWin,
  HPanel,
  UIUtil, Menus,
{$IFDEF EAGLCLIENT}
  MainEagl,      // AGLLanceFiche
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB,
  FE_Main,       // AGLLanceFiche
{$ENDIF}
  Buttons,
  UTOB,
  HEnt1,
  Ent1,          // ListePeriode
  hctrls,        // OpenSql
  Saisie,
  SaisUtil,      // QuellePeriode
  UtilSais,
  HmsgBox,       // pour le JaiLeDroitConcept
  Lookup,
  HSysMenu,
  uTof,
  StdCtrls,
  Mask,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Grids, HFLabel, HImgList, TntStdCtrls, TntGrids;

type
  TTriTw    = ( ttPeriode, ttJal );    // stocke le mode de tri courant pour le TreeView, par defaut TriExo
  TTypeNode = ( tnExo, tnJal, tnPer );

  // structure contenant les info pour la gestion de l'interface.
  // ces données sont rattachées au Data du TreeView et du ListView
  PInfoTW = ^RInfoTW;
  RInfoTW = record

  // Exercice
  TExo            : TExoDate ;
  EX_Validee      : string  ;
  JL_Validee      : string  ;

  InAnnee         : Word ;
  InMois          : Word ;

  J_JOURNAL       : string  ;
  J_LIBELLE       : string  ;
  J_NATUREJAL     : string  ;
  J_MODESAISIE    : string  ;
  J_FERME         : Boolean ;

  // Variables de calcul
  Libelle         : string  ;
  Debit           : Double  ;
  Credit          : Double  ;
  Mouvement       : Integer ;

  // Boolean pour les icones dans le TreeView
  BoClosDefinitif : boolean ;
  BoClosPeriodique: boolean ;
  BoValide        : boolean ; // indicateur de Période validée 
  BoFerme         : boolean ;
  BoAvecFille     : boolean ;
  BoVide          : boolean ;
  BoTriExo        : boolean ;
  BoASupprimer    : boolean ;
  BoUtilise       : boolean ;

  InId            : integer ;
  InNiv           : integer ;

  TypeNode        : TTypeNode;
  StInfoLibre     : string ;

end;

  TFJournalCentralisateur = class(TForm)
    DockBottom: TDock97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    Splitter: TSplitter;
    HMTrad: THSystemMenu;
    POPAFF: TPopupMenu;
    MMAFFPERIODE: TMenuItem;
    MMAFFJAL: TMenuItem;
    POPZoom: TPopupMenu;
    MMZoomJal: TMenuItem;
    FLISTE: THGrid;
    Toolbar971: TToolbar97;
    BZoom: TToolbarButton97;
    ImXP: THImageList;
    PopF11: TPopupMenu;
    AFFPeriode: TMenuItem;
    AFFJournal: TMenuItem;
    N1: TMenuItem;
    NouvelElement: TMenuItem;
    AccesDetail: TMenuItem;
    Reclassement: TMenuItem;
    AffNoeudVide: TMenuItem;
    N3: TMenuItem;
    Imprimer: TMenuItem;
    JALCECR: TMenuItem;
    JALCENTRAL: TMenuItem;
    BALGENE: TMenuItem;
    GLECR: TMenuItem;
    BNew: TToolbarButton97;
    Dock971: TDock97;
    PEntete: TToolWindow97;
    BCherche: TToolbarButton97;
    TEXERCICE: THLabel;
    EXERCICE: THValComboBox;
    TETABLISSEMENT: THLabel;
    ETABLISSEMENT: THMultiValComboBox;
    imTSE: THImageList;
    TW: TTreeView;
    AppliquerLesCriteres: TMenuItem;
    AccesJournal: TMenuItem;
    BAFF: TToolbarButton97;
    MMAFFNOEUDVIDE: TMenuItem;
    FlashingLabel1: TFlashingLabel;

    procedure FormCreate           (Sender : TObject);
    procedure FormShow             (Sender : TObject);
    procedure FormKeyDown          (Sender : TObject; var Key: Word;Shift: TShiftState);

    procedure TWExpanded           (Sender : TObject; Node : TTreeNode);
    procedure TWChange             (Sender : TObject; Node : TTreeNode);
    procedure TWClick              (Sender : TObject);

    procedure MMAFFPERIODEClick    (Sender : TObject);
    procedure MMAFFJALClick        (Sender : TObject);
    procedure MMZoomJalClick       (Sender : TObject);
    procedure MMAFFNOEUDVIDEClick  (Sender : TObject);

    procedure EXERCICEChange       (Sender : TObject);
    procedure ETABLISSEMENTChange  (Sender : TObject);

    procedure SplitterMoved        (Sender : TObject);
    procedure OnPostDrawCellFListe (ACol, ARow : Integer; Canvas : TCanvas; AState : TGridDrawState);
    procedure OnGetCellCanvasFListe(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState);
    procedure FListeDblClick       (Sender : TObject);
    procedure AccesDetailClick     (Sender : TObject);

    procedure FormClose            (Sender : TObject; var Action : TCloseAction);
    procedure NouvelElementClick   (Sender : TObject);
    procedure AFFPeriodeClick      (Sender : TObject);
    procedure AFFJournalClick      (Sender : TObject);
    procedure BNewClick            (Sender : TObject);
    procedure JALCECRClick         (Sender : TObject);
    procedure JALCENTRALClick      (Sender : TObject);
    procedure BALGENEClick         (Sender : TObject);
    procedure GLECRClick           (Sender : TObject);
    procedure PopF11Popup          (Sender : TObject);
    procedure BChercheClick        (Sender : TObject);
    procedure AffNoeudVideClick    (Sender : TObject);
    procedure AppliquerLesCriteresClick(Sender: TObject);
    procedure AccesJournalClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);

  private
    { Déclarations privées }
   FTriTw      : TTriTw ;        // stocke le mode de tri courant pour le TreeView, par defaut TriExo
   FCurrentId  : integer;        // utilise pour identifie un noueds ds la treeview

   FBoAfficheVide : Boolean;

{$IFNDEF EAGLCLIENT}
   FBoToolBarUp : Boolean;
{$ENDIF}

   FBoLoading     : Boolean;
   FBoAutoRefresh : Boolean;

   FStSqlEtab : string; // Condition sur la combo Etablissement

   // Variables de repositionnement sur le TreeView
   FLastAnnee : word;
   FLastMois  : word;
   FLastJour  : word;
   FLastJal   : string;

   procedure AjouteItemVide             ( vNode : TTreeNode );      // ajoute un noeud a supprimer par la suite au noeud courant, sert a faire apparaitre le '+' sur celui ci, le treeview etant remplis que qd on le selectionne
   procedure SuppItemVide               ( vNode : TTreeNode );      // supprime le noeud a supprimer avant de rajoute un element un noeud
   procedure RemplirTreeviewExo ;                                   // remplit le treeview avec l'ensemble des exo du dossiers
   procedure RemplirTreeviewPeriode     ( vRootNode : TTreeNode );  // remplit le treeview avec l'ensemble des periodes de l'exo
   procedure RemplirTreeviewPeriodeJal  ( vRootNode : TTreeNode );  // remplit le treeview avec l'ensemble des journal des periodes de l'exo
   procedure RemplirTreeviewJal         ( vRootNode : TTreeNode );  // remplit le treeview avec l'ensemble des journaux de l'exo
   procedure RemplirTreeviewJalPeriode  ( vRootNode : TTreeNode );  // remplit le treeview avec l'ensemble des periodes de l'exo

   procedure InitDefaut;
   procedure InitFListe;
   procedure Creation;

   procedure RemplirFListe              ( vNode : TTreeNode );
   procedure RemplitFListeAvecInfoTW    ( vNode : TTreeNode );
   procedure RemplitFListeAvecMouvement ( vNode : TTreeNode );
   procedure OuvertureSaisie            ( vNode : TTreeNode ; vBoCreation : Boolean );

   procedure VideTW ;
   procedure ChangeTri(vTriTw : TTriTw) ;
   procedure RemplirTW(vNode: TTreeNode);
   procedure EnabledControl;
   //function  GetLibJal      ( vP : PInfoTW ) : string;
   //function  GetLibExo      ( vP : PInfoTW ) : string;
   function  MakeSQLBor( vP : PInfoTW ) : string ;

   procedure DoubleCliqueFListe;
   procedure ZoomJal;
  end;

procedure CPLanceFiche_JournalCentralisateur ;
procedure CPInitPInfoTW ( var vP : PInfoTW ) ;

implementation

uses uLibEcriture,
     uLibExercice,  // CInitComboExercice
     CritEdt,       // TCritEdt
     AglInit,       // TheData
     uLibWindows,   // TraductionTHMultiValComboBox
     SaisBor,       // LanceSaisieFolio
     ParamSOc,      // GetParamSocSecur
{$IFDEF COMPTA}
     uTofCPJalEcr,  // CPLanceFiche_CPJALECR
     uTofCPGlGene,  // CPLanceFiche_CPGLGENE('')
     uTofCPMulMvt,  // MultiCritereMvt(taCreat, 'N', False)
     CPBalGen_Tof,  // CPLanceFiche_BalanceGeneral
{$ENDIF}
     CPJournal_TOM;

{$R *.DFM}

const
 cImageBureau      = 0 ;
 cImageDoss        = 1 ;
 cImageDossVide    = 2 ;
 cImageDossValide  = 3 ;
 cImageClot        = 4 ;
 cImageDossFerme   = 5 ;

 cStTitreTWDoss    = 'Dossier';
 cStTitreTWExo     = 'Exercice';

 cStBordereau      = 'Folio';
 cStPiece          = 'Pièce';
 cStLibre          = 'Folio';
 cStPer            = 'Période';
 cStExo            = 'Exercice';
 cStJournal        = 'Journal';
 cStExoClos        = 'Exercice clos';
 cStPeriodeVal     = 'Période validée';
 cStPeriodeClo     = 'Période cloturée';

 cStNouveauJournal = 'Nouveau journal';
 cStNouveauBor     = 'Nouveau folio';
 cStNouvellePiece  = 'Nouvelle pièce';

 cErrPasEnr        = 'Aucun Enregistrment sélectionné';
 cErrPasInfo       = 'Aucune info. de dispo pour ce noeud';
 cErrPasEnrT       = 'Aucun enregistrement !';
 cErrMvtValide     = 'Vous ne pouvez pas modifier ce bordereau,'+#10#13+'il contient des éléments pointés, lettrés ou validés !';
 cErrPrg           = - 100 ;
 cErrTransfert     = - 101 ;

 cMaxCol           = 5;
 cColLibelle       = 1;
 cColDebit         = 2;
 cColCredit        = 3;
 cColMouvement     = 4;
 cColNumeroPiece   = 5;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function _DecodeModeSaisie( vStModeSaisie : string ) : string ;
begin
  if vStModeSaisie = 'BOR' then
    Result := cStBordereau
  else
    if ( vStModeSaisie = '-' ) or ( vStModeSaisie = '' ) then
      Result := cStPiece
    else
      Result := cStLibre ;

  Result := TraduireMemoire(Result) ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... : 13/09/2006    
Description .. : GCO - 13/09/2006 - FQ 18269 - Ajout FStSqlETab
Mots clefs ... : 
*****************************************************************}
function TFJournalCentralisateur.MakeSQLBor( vP : PInfoTW ) : string ;
begin
  Result := 'SELECT E_NUMEROPIECE NUMEROPIECE, SUM(E_DEBIT) DEBIT, ' +
            'SUM(E_CREDIT) CREDIT, COUNT(*) MOUVEMENT FROM ECRITURE WHERE ' +
            'E_EXERCICE = "' + vP^.TExo.Code + '" AND ' +
            'E_JOURNAL = "' + vP^.J_JOURNAL + '" AND ' + FStSqlEtab +
            'E_DATECOMPTABLE >= "' + USDateTime(EncodeDateBor(vP^.InAnnee, vP^.InMois , vP^.TExo))+ '" AND ' +
            'E_DATECOMPTABLE <= "' + USDateTime(FinDeMois(EncodeDateBor(vP^.InAnnee, vP^.InMois , vP^.TExo))) + '" AND ' +
            'E_QUALIFPIECE = "N" GROUP BY E_NUMEROPIECE ORDER BY E_NUMEROPIECE';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_JournalCentralisateur;
var
  lF : TFJournalCentralisateur ;
  PP : THPanel ;
begin
  try
    lF:=TFJournalCentralisateur.Create(Application) ;
    PP:=FindInsidePanel ;
    if PP=nil then
    begin
      try
        lF.ShowModal ;
      finally
        lF.Free ;
      end ;
      // Screen.Cursor:=SyncrDefault ;
    end
    else
    begin
      InitInside(lF, PP) ;
      lF.Show ;
    end;
  finally

  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function UpperCaseLettre1(vStChaine: string): string;
begin
  Result := UpperCase(vStChaine[1]) + LowerCase(Copy(vStChaine, 2, length(vStChaine)));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/02/2006
Modifié le ... : 30/03/2006
Description .. :
Mots clefs ... :
*****************************************************************}
function _FabriqTitre( vP : PInfoTW ; vTri : TTriTw ) : string ;
var lStPeriode : string;
begin
  if vP = nil then exit ;

  Result := 'Saisie';

  if (vP^.BoClosDefinitif) or (vP^.BoClosPeriodique) then
    lStPeriode := cStPeriodeClo
  else
    if vP^.BoValide then
      lStPeriode := cStPeriodeVal
    else
      lStPeriode := cStPer;
 
  if vTri = ttPeriode then
  begin
    if (vP^.InAnnee <> 0) and (vP^.InMois <> 0) then
      Result := Result + ' - ' + lStPeriode + ' ' + UpperCaseLettre1(FormatDateTime('mmmm yyyy',EncodeDateBor(vP^.InAnnee,vP^.InMois,vP^.TExo)));
    if length(vP^.J_LIBELLE) > 0 then
      Result :=  Result + ' - (' + vP^.J_JOURNAL + ') ' + vP^.J_LIBELLE;
  end // if
  else
  begin
    if length(vP^.J_LIBELLE) > 0 then
      Result := Result + ' - (' + vP^.J_JOURNAL + ') ' + vP^.J_LIBELLE;
    if (vP^.InAnnee <> 0) and (vP^.InMois <> 0) then
      Result := Result + ' - ' + lStPeriode + ' ' + UpperCaseLettre1(FormatDateTime('mmmm yyyy',EncodeDateBor(vP^.InAnnee,vP^.InMois,vP^.TExo)));
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/02/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure _AffecteImageNode( vNode : TTreeNode ; vP : PInfoTW ) ;
begin
  if (vP^.BoClosDefinitif) or (vP^.BoClosPeriodique) then
  begin
    vNode.ImageIndex := cImageClot;
  end
  else
    if vP^.BoFerme then
    begin
      vNode.ImageIndex := cImageDossFerme;
    end
    else
      if vP^.BoValide then
      begin
        vNode.ImageIndex := cImageDossValide;
      end
      else
      begin
        if vP^.BoAvecFille then
        begin
          vNode.ImageIndex := cImageDoss;
        end
        else
        begin
          vNode.Imageindex := cImageDossVide;
        end;
      end;

  vNode.SelectedIndex := vNode.ImageIndex;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CPInitPInfoTW ( var vP : PInfoTW ) ;
begin
 vP^.TExo.Code         := '' ;
 vP^.EX_Validee        := '' ;

 vP^.TExo.Deb          := 0 ;
 vP^.TExo.Fin          := 0 ;
 vP^.InAnnee           := 0 ;
 vP^.InMois            := 0 ;

 vP^.J_JOURNAL         := '' ;
 vP^.J_LIBELLE         := '' ;
 vP^.J_NATUREJAL       := '' ;
 vP^.J_MODESAISIE      := '' ;
 vP^.J_FERME           := False;

 vP^.JL_Validee        := '' ;
 
 vP^.Libelle           := '';
 vP^.Debit             := 0 ;
 vP^.Credit            := 0 ;
 vP^.Mouvement         := 0 ;

 vP^.BoClosDefinitif   := False ;
 vP^.BoClosPeriodique  := False ;
 vP^.BoValide          := False ;
 vP^.BoFerme           := False ;

 vP^.InId              := -1 ;
 vP^.BoVide            := false;
 vP^.BoAvecFille       := False;

 vP^.InNiv             := -1 ;
 vP^.BoTriExo          := true ;
 vP^.BoASupprimer      := false ;
 vP^.TypeNode          := tnExo ;
 vP^.BoUtilise         := false ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function _IsValidJournalPeriode( vStExo , vStValideN , vStValideN1 : string ;  vDtDate : TDateTime ; vInfo : TInfoEcriture) : integer ;
var
 lBoExoN     : boolean ;
 lInIndexPer : integer ; // indice de la periode pour les champs valideen ( chaine de 24 car representant deux annees )
 lDtDate     : TDateTime ;
 lStValide   : string ;
begin
 result  := RC_PASERREUR ;

 lBoExoN := vStExo = vInfo.Exercice.EnCours.Code ;
 if ( not lBoExoN ) and ( vStExo <>vInfo.Exercice.Suivant.Code ) then Exit ; // on est pas sur l'exo en cours ni le suivant, il est donc forcement clos

 if lBoExoN then
  begin
   lStValide := vStValideN ;
   lDtDate   := vInfo.Exercice.EnCours.Deb ;
  end
   else
    begin
     lStValide := vStValideN1 ;
     lDtDate   := vInfo.Exercice.Suivant.Deb ;
    end ;

 lInIndexPer := QuellePeriode(vDtDate , lDtDate ) ;
 if ( lInIndexPer > 0 ) and ( lStValide[lInIndexPer] = 'X' ) then
  result := RC_JALVALID ;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.FormCreate(Sender: TObject);
var lLastDate : TDateTime;
begin
  if V_PGI.ModeTSE or V_PGI.Mode256cl or V_PGI.WinRemote then
  begin
    TW.Images := imTSE;
    POpF11.Images := imTSE;
  end
  else
  begin
    TW.Images := imXP;
    POpF11.Images := imXP;
  end;

{$IFNDEF EAGLCLIENT}
  FBoToolBarUp      := V_Pgi.ToolBarreUp;
  V_Pgi.ToolBarreUp := False;
{$ENDIF}

  CInitComboExercice(EXERCICE);

  // Ajout du menu Bouton droit de la compta
  AddMenuPop(PopF11, '', '');

  //PgiInfo(GetParamSocSecur('SO_ZNATCOMPL', '', False), 'SO_ZNATCOMPL');
  FBoLoading := True;

  // Chargement des infos pour le repositionnment par défaut à l'ouverture
  // GCO - 23/06/2006 - FQ 18389 + FQ 19334 pour rechargemment des paramsoc
  lLastDate := GetParamSocSecur('SO_ZLASTDATE', idate1900, True);
  DecodeDate( lLastDate, FLastAnnee, FLastMois, FLastJour);
  FLastJal  := GetParamSocSecur('SO_ZLASTJAL', '', True);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/02/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.FormShow(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
 // V_Pgi.ToolBarreUp := FBoToolBarUp;
{$ENDIF}
  InitDefaut;

  // Fin du chargement de l'écran
  FBoLoading := False;

  // Remplissage du TreeView, et calcul
  RemplirTreeviewExo ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.InitDefaut;
begin
  InitFListe;

  // GCO - 01/06/2006 - FQ
  if GetParamSocSecur('SO_CPJALCENTRALSURJOURNAL', False) then
    FTriTw := TTJal
  else
    FTriTw := TTPeriode;

  ChangeTri(FTriTw);

  FBoAfficheVide := True;
  MMAFFNoeudVide.checked := True;

  if CtxPCl in V_Pgi.PgiContexte then
  begin
    if VH^.CPExoRef.Code <> '' then
      Exercice.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
    else
      Exercice.Value := CExerciceVersRelatif(VH^.Encours.Code);

    BCherche.Visible := False;
    AppliquerLesCriteres.Visible := False;
    FBoAutoRefresh := True;
  end
  else
  begin
    Exercice.Value := CExerciceVersRelatif(VH^.Entree.Code);
    FBoAutoRefresh := GetParamSocSecur('SO_BDEFGCPTE', False);
  end;

  PositionneEtabUser(Etablissement);
  if Etablissement.Value = '' then
  begin
    Etablissement.SelectAll;
    Etablissement.Text := TraduireMemoire('<<Tous>>');
  end;

  if Etablissement.Items.Count = 1 then
  begin
    TEtablissement.Visible := False;
    Etablissement.Visible := False;
  end;

  FCurrentId := 0;
  EnabledControl;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.InitFListe;
begin
  FListe.ColCount := cMaxCol + 1;

  FListe.Cells[cColLibelle, 0]     := TraduireMemoire('Intitulé');
  FListe.ColAligns[cColLibelle]    := TaLeftJustify;
  FListe.ColWidths[cColLibelle]    := 300;

  FListe.Cells[cColDebit, 0]       := TraduireMemoire('Débit');
  FListe.ColTypes[cColDebit]       := 'R';
  FListe.ColFormats[cColDebit]     := StrfMask(V_PGI.OkDecV, '', True);
  FListe.ColAligns[cColDebit]      := TaRightJustify;
  FListe.ColWidths[cColDebit]      := 125;

  FListe.Cells[cColCredit, 0]      := TraduireMemoire('Crédit');
  FListe.ColTypes[cColCredit]      := 'R';
  FListe.ColFormats[cColCredit]    := StrfMask(V_PGI.OkDecV, '', True);
  FListe.ColAligns[cColCredit]     := TaRightJustify;
  FListe.ColWidths[cColCredit]     := 125;

  FListe.Cells[cColMouvement, 0]   := TraduireMemoire('Mouvements');
  FListe.ColTypes[cColMouvement]   := 'R';
  FListe.ColFormats[cColMouvement] := '###0';
  FListe.ColAligns[cColMouvement]  := TaRightJustify;
  FListe.ColWidths[cColMouvement]  := 90;

  FListe.Cells[cColNumeroPiece, 0]   := TraduireMemoire('NumeroPiece');
  FListe.ColTypes[cColNumeroPiece]   := 'R';
  FListe.ColFormats[cColNumeroPiece] := '###0';
  FListe.ColAligns[cColNumeroPiece]  := TaRightJustify;

{$IFDEF TT}
  FListe.ColWidths[cColNumeroPiece]  := 90;
{$ELSE}
  FListe.ColWidths[cColNumeroPiece]  := -1;
{$ENDIF}

  FListe.PostDrawCell  := OnPostDrawCellFListe;
  FListe.GetCellCanvas := OnGetCellCanvasFListe;
  //fListe.OnColumnWidthsChanged := OnColumnWidthsChangedFListe;

  HMTrad.ResizeGridColumns(FListe);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE            
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.EXERCICEChange(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ; 
  if not FBoAutoRefresh then Exit;
  RemplirTreeviewExo ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.ETABLISSEMENTChange(Sender: TObject);
var lSt1, lSt2 : string;
begin
  if csDestroying in ComponentState then Exit ;
  
  FStSqlEtab := '';

  if (not Etablissement.Tous) then
  begin
    TraductionTHMultiValComboBox( Etablissement, lSt1, lSt2, 'E_ETABLISSEMENT', False);
    if lSt1 <> '' then
      FStSqlEtab := lSt1 + ' AND ';
  end;

  if not FBoAutoRefresh then Exit;
  RemplirTreeviewExo ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/02/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.RemplirTreeviewExo;
var
  lQ        : TQuery ;
  lNodeDoss : TTreeNode ;
  lP        : PInfoTW ;
begin
  if FBoLoading then Exit;

  lNodeDoss := nil;
  VideTW ;
  try
    try
      lQ := OpenSQL('SELECT EX_EXERCICE, EX_LIBELLE, EX_ETATCPTA, EX_VALIDEE, ' +
                    'EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE ' +
                    'EX_EXERCICE = "' + CRelatifVersExercice(Exercice.Value) + '"', True) ;
      New(lP) ;
      CPInitPInfoTW(lP) ;
      lP^.TExo.Code          := lQ.findField('EX_EXERCICE').asString ;
      lP^.Libelle            := lQ.findField('EX_LIBELLE').asString ;
      lP^.TExo.Deb           := lQ.findField('EX_DATEDEBUT').asDateTime ;
      lP^.TExo.Fin           := lQ.findField('EX_DATEFIN').asDateTime ;
      lP^.EX_Validee         := lQ.findField('EX_VALIDEE').asString;

      lP^.BoClosDefinitif    := (lQ.findField('EX_ETATCPTA').asString <> 'OUV') and
                                (lQ.findField('EX_ETATCPTA').asString <> 'CPR');

      lP^.InNiv              := 0 ;
      lP^.BoTriExo           := true ;
      lP^.BoVide             := true ;
      lP^.BoASupprimer       := false ;
      lP^.InId               := FCurrentId ;
      lP^.TypeNode           := tnExo ;
      lNodeDoss              := TW.Items.AddChildObject(nil, lP^.Libelle, LP);
      lNodeDoss.ImageIndex   := cImageBureau ;
      AjouteItemVide(lNodeDoss);

      inc(FCurrentId ) ;
    except
      on E: Exception do PgiError(E.Message, '- Procedure RemplirTreeviewExo - ');
    end;
  finally
    Ferme(lQ) ;
  end;

  if lNodeDoss <> nil then
    lNodeDoss.Expand(False);

  if Tw.CanFocus then
    TW.SetFocus;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/02/2006
Modifié le ... : 18/09/2007
Description .. : - LG - 18/09/2007 - FB 20625 - correction pour les cloture 
Suite ........ : periodique sur les exercices decales
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.RemplirTreeviewPeriode( vRootNode : TTreeNode ) ;
var
  lNode      : TTreeNode;
  lNodeTemp  : TTreeNode;
  lPRoot     : PInfoTW;
  lP         : PInfoTW;
  lAnnee     : word;
  lMois      : word;
  lJour      : word;
  lNbMois    : word;
  i          : integer;
  lTobPeriode: Tob;
begin
  lPRoot := vRootNode.Data;
  if not Assigned(lPRoot) then Exit;

  TW.Items.BeginUpdate;
  lNodeTemp := nil;
  try
    try
      lPRoot^.BoVide := False;
      SuppItemVide(vRootNode);

      inc(FCurrentId) ;

      lTobPeriode := Tob.Create('', nil, -1);
      lTobPeriode.LoadDetailFromSQL('SELECT MONTH(E_DATECOMPTABLE) MOIS, ' +
        'YEAR(E_DATECOMPTABLE) ANNEE, SUM(E_DEBIT) DEBIT, SUM(E_CREDIT) CREDIT, ' +
        'COUNT(*) MOUVEMENT FROM ECRITURE WHERE ' +
        'E_EXERCICE = "' + lPRoot^.TExo.Code + '" AND ' + FStSqlEtab +
        'E_QUALIFPIECE = "N" ' +
        'GROUP BY YEAR(E_DATECOMPTABLE),MONTH(E_DATECOMPTABLE) ' +
        'ORDER BY YEAR(E_DATECOMPTABLE),MONTH(E_DATECOMPTABLE)');

      lPRoot^.Debit     := 0;
      lPRoot^.Credit    := 0;
      lPRoot^.Mouvement := 0;

      NombrePerExo(lPRoot^.TExo, lMois, lAnnee, lNbMois) ;
      for i := 0 to lNbMois -1 do
      begin
        DecodeDate(PlusMois(lPRoot^.TExo.Deb, i), lAnnee, lMois, lJour);

        lP := nil;
        if FBoAfficheVide then
        begin
          New(lP) ;
          CPInitPInfoTW( lP ) ;
        end;

        if lTobPeriode.Detail.Count > 0 then
        begin
          if (lTobPeriode.Detail[0].GetInteger('ANNEE') = lAnnee) and
             (lTobPeriode.Detail[0].GetInteger('MOIS') = lMois) then
          begin
            if not FBoAfficheVide then
            begin
              New(lP);
              CPInitPInfoTW(lP);
            end;

            lP^.Debit       := lTobPeriode.Detail[0].GetDouble('DEBIT');
            lP^.Credit      := lTobPeriode.Detail[0].GetDouble('CREDIT');
            lP^.Mouvement   := lTobPeriode.Detail[0].GetInteger('MOUVEMENT');
            lP^.BoAvecFille := True;

            lPRoot^.Debit     := lPRoot^.Debit +  lP^.Debit;
            lPRoot^.Credit    := lPRoot^.Credit + lP^.Credit;
            lPRoot^.Mouvement := lPRoot^.Mouvement + lP^.Mouvement;

            lTobPeriode.Detail[0].Free;
          end;
        end;

        if (lP = nil) then Continue;

        lP^.TExo      := lPRoot^.TExo ;
        lP^.InAnnee   := lAnnee;
        lP^.InMois    := lMois;
        lP^.InNiv     := 1 ;
        lP^.BoTriExo  := True;
        lP^.BoVide    := True;

        lP^.BoASupprimer       := False ;
        lP^.InId               := FCurrentId ;

        // FQ 17694
        lP^.BoClosDefinitif    := lPRoot^.BoClosDefinitif;
        lP^.BoClosPeriodique   := (( VH^.DateCloturePer > 0 ) and ( DebutDeMois(EncodeDateBor(lP^.InAnnee, lP^.InMois,lP^.TExo )) <= VH^.DateCloturePer )) and
                                  ( ctxExercice.QuelExoDate(VH^.DateCloturePer).Code = lP^.TExo.Code ) ; // FB 20625
        lP^.TypeNode           := tnPer ;
        LP^.BoValide           := lPRoot^.EX_Validee[i+1] = 'X';
        lP^.Libelle            := UpperCaseLettre1(FormatDateTime('mmmm yyyy',EncodeDateBor(lP^.InAnnee,lP^.InMois,lP^.TExo)));
        lNode                  := TW.Items.AddChildObject(vRootNode, lP^.Libelle,lP) ;
        _AffecteImageNode(lNode,lP) ;
        AjouteItemVide(lNode) ;

        if (lAnnee = FLastAnnee) and (lMois = FLastMois) then
        begin
          lNodeTemp := lNode;
        end;
      end;

    except
      on E: Exception do PgiError(E.Message, '- Procedure RemplirTreeviewPeriode - ');
    end;
  finally

    if assigned(lNodeTemp) then
    begin
      TW.Selected  := lNodeTemp;
      TW.Selected.Expand(False);
    end;

    TW.Items.EndUpdate;
    FreeAndNil(lTobPeriode);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.RemplirTreeviewJal( vRootNode : TTreeNode ) ;
var lNode      : TTreeNode;
    lNodeTemp  : TTreeNode;
    lPRoot     : PInfoTW;
    lP         : PInfoTW;
    lTobListeJal   : Tob;
    lTobJal        : Tob;
    lTobJalAvecEcr : Tob;
    i          : integer;
begin
  lPRoot := vRootNode.Data ;
  if not assigned(lPRoot) then exit ;

  TW.Items.BeginUpdate;
  lNodeTemp := nil;
  try
    try
      lPRoot^.BoVide := false;

      lTobListeJal := Tob.Create('', nil, -1);
      lTobListeJal.LoadDetailFromSQL('SELECT J_JOURNAL, J_LIBELLE, J_MODESAISIE, ' +
        'J_NATUREJAL, J_VALIDEEN, J_VALIDEEN1, J_FERME FROM JOURNAL ORDER BY J_JOURNAL');

      if lTobListeJal.Detail.Count <> 0 then SuppItemVide(vRootNode);
      inc(FCurrentId ) ;

      lTobJalAvecEcr := Tob.Create('', nil, -1);
      lTobJalAvecEcr.LoadDetailFromSql('SELECT E_JOURNAL, SUM(E_DEBIT) DEBIT, ' +
        'SUM(E_CREDIT) CREDIT, COUNT(*) MOUVEMENT FROM ECRITURE WHERE ' +
        'E_QUALIFPIECE = "N" AND ' + FstSqlEtab +
        'E_EXERCICE = "' + lPRoot^.TExo.Code + '" ' +
        'GROUP BY E_JOURNAL ORDER BY E_JOURNAL', True);

      lPRoot^.Debit     := 0;
      lPRoot^.Credit    := 0;
      lPRoot^.Mouvement := 0;

      for i := 0 to lTobListeJal.Detail.Count -1 do
      begin
        lP := nil;

        lTobJal := lTobListeJal.Detail[i];

        // Pas d'affichage des journaux de type CLO ou ODA
        if (lTobJal.GetString('J_NATUREJAL') = 'CLO') or
           (lTobJal.GetString('J_NATUREJAL') = 'ODA') then Continue;

        // GCO - 26/10/2006 - FQ 18494
        // Ne pas afficher les journaux fermés sans mouvement
        if (lTobJal.GetString('J_FERME') = 'X')  then
        begin
           if lTobJalAvecEcr.Detail.Count = 0 then
             Continue
           else
             if (lTobJalAvecEcr.Detail[0].GetString('E_JOURNAL') <> lTobJal.GetString('J_JOURNAL')) then
               Continue;
        end;       

        if FBoAfficheVide then
        begin
          New(lP) ;
          CPInitPInfoTW( lP ) ;
        end;

        if lTobJalAvecEcr.Detail.Count > 0 then
        begin
          if lTobJalAvecEcr.Detail[0].GetString('E_JOURNAL') = lTobJal.GetString('J_JOURNAL') then
          begin
            if not FBoAfficheVide then
            begin
              New(lP);
              CPInitPInfoTW(lP);
            end;

            lP^.Debit       := lTobJalAvecEcr.Detail[0].GetDouble('DEBIT');
            lP^.Credit      := lTobJalAvecEcr.Detail[0].GetDouble('CREDIT');
            lP^.Mouvement   := lTobJalAvecEcr.Detail[0].GetInteger('MOUVEMENT');
            lP^.BoAvecFille := True;

            lTobJalAvecEcr.Detail[0].Free;

            // Stockage des cumuls dans le noeud Parent
            lPRoot^.Debit     := lPRoot^.Debit + lP^.Debit;
            lPRoot^.Credit    := lPRoot^.Credit + lP^.Credit;
            lPRoot^.Mouvement := lPRoot^.Mouvement + lP^.Mouvement;
          end;
        end;

        if (lP = nil) then Continue;

        lP^.TExo          := lPRoot^.TExo ;
        lP^.InAnnee       := lPRoot^.InAnnee ;
        lP^.InMois        := lPRoot^.InMois ;
        lP^.EX_Validee    := lPRoot^.EX_Validee;

        lP^.J_JOURNAL     := lTobJal.GetString('J_JOURNAL');
        lP^.J_LIBELLE     := lTobJal.GetString('J_LIBELLE');
        lP^.J_NATUREJAL   := lTobJal.GetString('J_NATUREJAL');
        lP^.J_MODESAISIE  := lTobJal.GetString('J_MODESAISIE');
        lP^.J_FERME       := lTobJal.GetBoolean('J_FERME');

        if lP^.TExo.Code = VH^.Encours.Code then
          lP^.JL_VALIDEE := lTobJal.GetString('J_VALIDEEN')
        else
          if lP^.TExo.Code = VH^.Suivant.Code then
           lP^.JL_VALIDEE := lTobJal.GetString('J_VALIDEEN1');

        lP^.Libelle       := lTobJal.GetString('J_LIBELLE');

        lP^.InNiv              := 1 ;
        lP^.BoTriExo           := False ;
        lP^.BoVide             := True ;
        lP^.BoASupprimer       := False ;
        lP^.InId               := FCurrentId ;
        lP^.BoClosDefinitif    := lPRoot^.BoClosDefinitif ;
        lP^.BoClosPeriodique   := lPRoot^.BoClosPeriodique ;
        lP^.BoValide           := lP^.J_NATUREJAL = 'ANO' ;
        lP^.BoFerme            := (lTobJal.GetString('J_FERME')= 'X');
        lP^.TypeNode           := tnJal ;
        lNode                  := TW.Items.AddChildObject(vRootNode, lP^.Libelle,lP) ;
        _AffecteImageNode(lNode,lP) ;
        AjouteItemVide(lNode) ;

        if (lP^.J_Journal = FLastJal) and
           ((EncodeDate(FLastAnnee, FLastMois, FLastJour) >= lP^.TExo.Deb) and
            (EncodeDate(FLastAnnee, FLastMois, FLastJour) <= lP^.TExo.Fin)) then
        begin
          lNodeTemp := lNode;
        end;
      end;

    except
      on E: Exception do PgiError(E.Message, '- Procedure RemplirTreeviewJal - ');
    end;
  finally

    if assigned(lNodeTemp) then
    begin
      TW.Selected  := lNodeTemp;
      TW.Selected.Expand(False);
    end;

    TW.Items.EndUpdate;
    FreeAndNil(lTobListeJal);
    FreeAndNil(lTobJalAvecEcr);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... : 21/02/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.RemplirTreeviewPeriodeJal( vRootNode : TTreeNode ) ;
var
 lNode          : TTreeNode ;
 lNodeTemp      : TTreeNode ;
 lPRoot         : PInfoTW ;
 lP             : PInfoTW ;
 lTobListeJal   : Tob;
 lTobJal        : Tob;
 lTobJalAvecEcr : Tob;
 i : integer;
begin
  lPRoot := vRootNode.Data ;
  if not assigned(lPRoot) then Exit;

  TW.Items.BeginUpdate;
  lNodeTemp := nil;
  try
    try
      lPRoot^.BoVide := false;

      lTobListeJal := Tob.Create('', nil, -1);
      lTobJalAvecEcr := Tob.Create('', nil, -1);

      lTobListeJal.LoadDetailFromSQL('SELECT J_JOURNAL, J_LIBELLE, J_MODESAISIE, ' +
        'J_NATUREJAL, J_VALIDEEN, J_VALIDEEN1, J_FERME FROM JOURNAL ' +
        'ORDER BY J_JOURNAL');

      if lTobListeJal.Detail.Count <> 0 then SuppItemVide(vRootNode);
      inc(FCurrentId ) ;

      // Pas de besoin de faire la requête si le lPRoot^.BoAvecFille = True
      if lPRoot^.BoAvecFille then
      begin

        lTobJalAvecEcr.LoadDetailFromSQL('SELECT E_JOURNAL, SUM(E_DEBIT) DEBIT, ' +
          'SUM(E_CREDIT) CREDIT, COUNT(*) MOUVEMENT FROM ECRITURE WHERE ' +
          'E_QUALIFPIECE="N" AND ' +
          'E_EXERCICE = "' + lPRoot^.TExo.Code + '" AND ' + FStSqlEtab +
          'YEAR(E_DATECOMPTABLE) = ' + IntToStr(lPRoot^.InAnnee) + ' AND ' +
          'MONTH(E_DATECOMPTABLE) = ' + IntToStr(lPRoot^.InMois) +  ' ' +
          'GROUP BY E_JOURNAL ORDER BY E_JOURNAL');
      end;

      for i := 0 to lTobListeJal.Detail.Count -1 do
      begin
        lP := nil;

        lTobJal := lTobListeJal.Detail[i];

        // Pas d'affichage des journaux de type CLO ou ODA
        if (lTobJal.GetString('J_NATUREJAL') = 'CLO') or
           (lTobJal.GetString('J_NATUREJAL') = 'ODA') then Continue;


        // Affiche le journal ANO sur la première période de l'exercice uniquement
        if (lTobJal.GetString('J_NATUREJAL') = 'ANO') and
           (vRootNode.Index <> 0) then Continue;

        // GCO - 26/10/2006 - FQ 18494
        // Ne pas afficher les journaux fermés sans mouvement
        if (lTobJal.GetString('J_FERME') = 'X') then
        begin
           if (lTobJalAvecEcr.Detail.Count = 0) then
             Continue
           else
             if (lTobJalAvecEcr.Detail[0].GetString('E_JOURNAL') <> lTobJal.GetString('J_JOURNAL')) then
               Continue;
        end;

        if FBoAfficheVide then
        begin
          New(lP) ;
          CPInitPInfoTW( lP ) ;
        end;

        if lTobJalAvecEcr.Detail.Count > 0 then
        begin
          if lTobJalAvecEcr.Detail[0].GetString('E_JOURNAL') = lTobJal.GetString('J_JOURNAL') then
          begin
            if not FBoAfficheVide then
            begin
              New(lP);
              CPInitPInfoTW(lP);
            end;

            lP^.Debit     := lTobJalAvecEcr.Detail[0].GetDouble('DEBIT');
            lP^.Credit    := lTobJalAvecEcr.Detail[0].GetDouble('CREDIT');
            lP^.Mouvement := lTobJalAvecEcr.Detail[0].GetInteger('MOUVEMENT');
            lP^.BoAvecFille := True;

            lTobJalAvecEcr.Detail[0].Free;
          end;
        end;

        if (lP = nil) then Continue;

        lP^.TExo               := lPRoot^.TExo ;
        lP^.InAnnee            := lPRoot^.InAnnee ;
        lP^.InMois             := lPRoot^.InMois ;
        lP^.J_JOURNAL          := lTobJal.GetString('J_JOURNAL');
        lP^.J_LIBELLE          := lTobJal.GetString('J_LIBELLE');
        lP^.J_NATUREJAL        := lTobJal.GetString('J_NATUREJAL');
        lP^.J_MODESAISIE       := lTobJal.GetString('J_MODESAISIE');
        lP^.J_FERME            := lTobJal.GetBoolean('J_LIBELLE');

        if lP^.TExo.Code = VH^.Encours.Code then
          lP^.JL_VALIDEE := lTobJal.GetString('J_VALIDEEN')
        else
          if lP^.TExo.Code = VH^.Suivant.Code then
            lP^.JL_VALIDEE := lTobJal.GetString('J_VALIDEEN1');

        lP^.Libelle            := lTobJal.GetString('J_LIBELLE');

        lP^.InNiv              := 2 ;
        lP^.BoTriExo           := True ;
        lP^.BoVide             := False;
        lP^.BoASupprimer       := False;
        lP^.InId               := FCurrentId ;
        lP^.BoClosDefinitif    := lPRoot^.BoClosDefinitif;
        lP^.BoClosPeriodique   := lPRoot^.BoClosPeriodique;

        // GCO - 03/05/2006 - FQ 17848, FQ 18175
        lP^.BoValide           := lPRoot^.BoValide or (lP^.J_NATUREJAL = 'ANO');
        if lP^.JL_VALIDEE <> '' then
          // GCO - 01/10/2007 - FQ 21528 - on cherchait dans 0+1 au lieu 4+1
          // car avec l'affichage que des mouvementés, "Mai" se retrouvait en
          // première position dans vRootNode, maintenant je me base sur
          // InMois qui lui ne change jamais après chargmenet de la structure
          //lP^.BoValide := lP^.BoValide or (lP^.JL_VALIDEE[vRootNode.Index+1] = 'X');
          lP^.BoValide := lP^.BoValide or (lP^.JL_VALIDEE[lP^.InMois] = 'X');

        lP^.BoFerme            := (lTobJal.GetString('J_FERME') = 'X');
        lP^.TypeNode           := tnJal ;
        lNode                  := TW.Items.AddChildObject(vRootNode, lP^.Libelle, lP) ;
        _AffecteImageNode(lNode,lP) ;

        if (lPRoot^.InAnnee = FLastAnnee) and
           (lPRoot^.InMois = FLastMois) and
           (lP^.J_Journal = FLastJal) then
        begin
          lNodeTemp := lNode;
        end;
      end;

    except
      on E: Exception do PgiError(E.Message, '- Procedure RemplirTreeviewPeriodeJal - ');
    end;

  finally

    if assigned(lNodeTemp) then
    begin
      TW.Selected  := lNodeTemp;
      RemplirFListe(TW.Selected);
    end;

    TW.Items.EndUpdate;
    FreeAndNil(lTobListeJal);
    FreeAndNil(lTobJalAvecEcr);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/02/2006
Modifié le ... : 18/09/2007
Description .. : - LG - 18/09/2007 - FB 20625 - correction pour les cloture 
Suite ........ : periodique sur les exercices decales
Suite ........ : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.RemplirTreeviewJalPeriode( vRootNode : TTreeNode ) ;
var
  lTobPeriodeAvecEcr : Tob;
  lNode     : TTreeNode;
  lNodeTemp : TTreeNode;
  lPRoot  : PInfoTW;
  lP      : PInfoTW;
  lAnnee  : word;
  lMois   : word;
  lJour   : word;
  lNbMois : word;
  i       : integer;
begin
  lPRoot := vRootNode.Data ;
  if not assigned(lPRoot) then exit ;

  TW.Items.BeginUpdate;

  lNodeTemp := nil;
  try
    try
      lPRoot^.BoVide := false ;
      SuppItemVide(vRootNode) ;
      inc(FCurrentId) ;

      lTobPeriodeAvecEcr := Tob.Create('', nil, -1);

      if lPRoot^.BoAvecFille then
      begin
        lTobPeriodeAvecEcr.LoadDetailFromSql('SELECT YEAR(E_DATECOMPTABLE) ANNEE, ' +
          'MONTH(E_DATECOMPTABLE) MOIS, SUM(E_DEBIT) DEBIT, SUM(E_CREDIT) CREDIT, ' +
          'COUNT(*) MOUVEMENT FROM ECRITURE WHERE '+
          'E_EXERCICE = "' + lPRoot^.TExo.Code + '" AND ' +
          'E_JOURNAL = "' + lPRoot^.J_JOURNAL + '" AND ' + FStSqlEtab +
          'E_QUALIFPIECE = "N" ' +
          'GROUP BY YEAR(E_DATECOMPTABLE),MONTH(E_DATECOMPTABLE) ' +
          'ORDER BY YEAR(E_DATECOMPTABLE),MONTH(E_DATECOMPTABLE)');
      end;

      lPRoot^.Debit     := 0;
      lPRoot^.Credit    := 0;
      lPRoot^.Mouvement := 0;

      NombrePerExo(lPRoot^.TExo, lMois, lAnnee, lNbMois) ;
      for i := 0 to lNbMois -1 do
      begin
        DecodeDate(PlusMois(lPRoot^.TExo.Deb, i), lAnnee, lMois, lJour);

        lP := nil;

        // Affiche le journal ANO sur la première période de l'exercice uniquement
        if (lPRoot^.J_NATUREJAL = 'ANO') and (i <> 0) then Continue;

        // GCO - 26/10/2006 - FQ 18494
        if (lPRoot^.J_FERME ) then
        begin
          if lTobPeriodeAvecEcr.Detail.Count = 0 then
            Continue
          else
            if (not ((lTobPeriodeAvecEcr.Detail[0].GetInteger('ANNEE') = lAnnee) and
                     (lTobPeriodeAvecEcr.Detail[0].GetInteger('MOIS') = lMois))) then
              Continue;
        end;

        if FBoAfficheVide then
        begin
          New(lP) ;
          CPInitPInfoTW( lP ) ;
        end;

        if lTobPeriodeAvecEcr.Detail.Count > 0 then
        begin
          if (lTobPeriodeAvecEcr.Detail[0].GetInteger('ANNEE') = lAnnee) and
             (lTobPeriodeAvecEcr.Detail[0].GetInteger('MOIS') = lMois) then
          begin
            if not FBoAfficheVide then
            begin
              New(lP);
              CPInitPInfoTW(lP);
            end;

            lP^.Debit       := lTobPeriodeAvecEcr.Detail[0].GetDouble('DEBIT');
            lP^.Credit      := lTobPeriodeAvecEcr.Detail[0].GetDouble('CREDIT');
            lP^.Mouvement   := lTobPeriodeAvecEcr.Detail[0].GetInteger('MOUVEMENT');
            lP^.BoAvecFille := True;

            // Stockage des cumuls dans le noeud Parent
            lPRoot^.Debit     := lPRoot^.Debit + lP^.Debit;
            lPRoot^.Credit    := lPRoot^.Credit + lP^.Credit;
            lPRoot^.Mouvement := lPRoot^.Mouvement + lP^.Mouvement;

            lTobPeriodeAvecEcr.Detail[0].Free;
          end;
        end;

        if (lP = nil) then Continue;

        lP^.TExo     := lPRoot^.TExo;
        lP^.InAnnee  := lAnnee;
        lP^.InMois   := lMois;
        lP^.InNiv    := 2 ;

        lP^.J_JOURNAL    := lPRoot^.J_JOURNAL;
        lP^.J_LIBELLE    := lPRoot^.J_LIBELLE;
        lP^.J_MODESAISIE := lPRoot^.J_MODESAISIE;
        lP^.J_NATUREJAL  := lPRoot^.J_NATUREJAL;
        lP^.J_FERME      := lPRoot^.J_FERME;

        lP^.JL_Validee   := lPRoot^.JL_Validee;

        lP^.BoTriExo     := True ;
        lP^.BoVide       := True ;
        lP^.BoASupprimer := false ;
        lP^.InId         := FCurrentId ;

        // FQ 17694
        lP^.BoClosDefinitif  := lPRoot^.BoClosDefinitif;
        lP^.BoClosPeriodique := (( VH^.DateCloturePer > 0 ) and ( DebutDeMois(EncodeDateBor(lP^.InAnnee, lP^.InMois,lP^.TExo )) <= VH^.DateCloturePer )) and
                                 ( ctxExercice.QuelExoDate(VH^.DateCloturePer).Code = lP^.TExo.Code ) ;  // FB 20625

        // GCO - 03/05/2006 - FQ 17848, FQ 18175
        if lPRoot^.JL_Validee <> '' then
          lP^.BoValide  := lPRoot^.JL_Validee[i+1] = 'X';

        lP^.BoFerme      := lPRoot^.BoFerme;
        lP^.TypeNode     := tnPer ;
        lP^.Libelle      := UpperCaseLettre1(FormatDateTime('mmmm yyyy',EncodeDateBor(lP^.InAnnee,lP^.InMois,lP^.TExo)));
        //if not lP^.BoValide then
        // lP^.BoValide          := ( _IsValidJournalPeriode( lP^.TExo.Code , lP^.J_VALIDEEN  , lP^.J_VALIDEEN1 , EncodeDateBor(lP^.InAnnee,lP^.InMois,lP^.TExo),FInfoEcr )  <> RC_PASERREUR ) or ( lP^.J_NATUREJAL = 'ANO' ) ;
        lNode := TW.Items.AddChildObject(vRootNode, lP^.Libelle, lP);
        _AffecteImageNode(lNode,lP) ;

        if (lP^.J_Journal = FLastJal) and
           (lP^.InAnnee = FLastAnnee) and
           (lP^.InMois = FLastMois) then
        begin
          lNodeTemp := lNode;
        end;

      end;

    except
      on E: Exception do PgiError(E.Message, '- Procedure RemplirTreeviewJalPeriode - ');
    end;
  finally
    if assigned(lNodeTemp) then
    begin
      TW.Selected  := lNodeTemp;
      TW.Selected.Expand(False);
    end;

    TW.Items.EndUpdate;
    FreeAndNil(lTobPeriodeAvecEcr);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.SplitterMoved(Sender: TObject);
begin
  HMTrad.ResizeGridColumns(FListe);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.RemplirFListe( vNode : TTreeNode );
begin
  if ( vNode = nil ) or ( vNode.Data = nil ) then exit ;

  RemplirTW(vNode) ;

  if PInfoTW(vNode.Data).InNiv = 2 then
  begin
    RemplitFListeAvecMouvement( vNode );
  end
  else
  begin
    RemplitFListeAvecInfoTW( vNode );
  end;

  HMTrad.ResizeGridColumns(FListe);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.RemplitFListeAvecMouvement(vNode: TTreeNode);
var lP : PInfoTW;
    lTob : Tob;
    lTobFille : Tob;
    i : integer;
    lTotDeb : double;
    lTotCre : double;
    lTotMvt : integer;
    lStLibelle : string;
begin
  lP := PInfoTW(vNode.Data);
  if lP = nil then Exit;

  // GCO - 23/06/2006 - FQ 18389
  FLastJal   := LP^.J_JOURNAL;
  FLastAnnee := LP^.InAnnee;
  FLastMois  := LP^.InMois;

  FListe.VidePile(False);
  lStLibelle := _DecodeModeSaisie(lP^.J_ModeSaisie);
  lTob := Tob.Create('', nil, -1);

  // Pas besion de faire la requête si pas de mouvement dans le noeud
  if lP^.BoAvecFille then
  begin
    lTob.LoadDetailFromSQL(MakeSQLBor(lP), False);
  end;

  lTotDeb := 0;
  lTotCre := 0;
  lTotMvt := 0;

  if lTob.Detail.Count <> 0 then
  begin
    FListe.RowCount := lTob.Detail.Count + 3;
    for i := 0 to lTob.Detail.Count -1 do
    begin
      lTob.Detail[i].AddChampSupValeur('LIBELLE', lStLibelle + ' ' + lTob.Detail[i].GetString('NUMEROPIECE'));
      lTotDeb := lTotDeb + lTob.Detail[i].GetDouble('DEBIT');
      lTotCre := lTotCre + lTob.Detail[i].GetDouble('CREDIT');
      lTotMvt := lTotMvt + lTob.Detail[i].GetInteger('MOUVEMENT');
    end;
  end
  else
    FListe.RowCount := 3;

  lTobFille := Tob.Create('', lTob, -1);
  lTobFille.AddChampSupValeur('LIBELLE', 'Total ' + lP^.Libelle + ' ' + PInfoTW(vNode.Parent.Data)^.Libelle);
  lTobFille.AddChampSupValeur('DEBIT',  lTotDeb);
  lTobFille.AddChampSupValeur('CREDIT', lTotCre);
  lTobFille.AddChampSupValeur('MOUVEMENT', lTotMvt);
  lTobFille.AddChampSupValeur('NUMEROPIECE', 0);

  lTobFille := Tob.Create('', lTob, -1);
  lTobFille.AddChampSupValeur('LIBELLE', 'Total ' + PInfoTw(TW.TopItem.Data)^.Libelle);
  lTobFille.AddChampSupValeur('DEBIT', PInfoTw(TW.TopItem.Data)^.Debit);
  lTobFille.AddChampSupValeur('CREDIT', PInfoTw(TW.TopItem.Data)^.Credit);
  lTobFille.AddChampSupValeur('MOUVEMENT', PInfoTw(TW.TopItem.Data)^.Mouvement);
  lTobFille.AddChampSupValeur('NUMEROPIECE', 0);

  lTob.PutGridDetail(FListe, False, False, 'LIBELLE;DEBIT;CREDIT;MOUVEMENT;NUMEROPIECE');

  if FListe.RowCount > 3 then
    FListe.Row := FListe.RowCount - 3;

  FreeAndNil(lTob);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.RemplitFListeAvecInfoTW( vNode: TTreeNode);
var lP : PInfoTW;
    i : integer;
    lTob : Tob;
    lTobFille : Tob;
    lTotDeb : double;
    lTotCre : double;
    lTotMvt : integer;
begin
  lP := PInfoTW(vNode.Data);
  if lP = nil then Exit;

  lTob := Tob.Create('', nil, -1);
  FListe.VidePile(False);

  FListe.RowCount := vNode.Count + 3;

  lTotDeb := 0;
  lTotCre := 0;
  lTotMvt := 0;

  for i:= 0 to vNode.Count - 1 do
  begin
    lTobFille := Tob.Create('', lTob, -1);
    lTobFille.AddChampSupValeur('LIBELLE', PInfoTW(vNode.Item[i].Data)^.Libelle);
    lTobFille.AddChampSupValeur('DEBIT', PInfoTW(vNode.Item[i].Data)^.Debit);
    lTobFille.AddChampSupValeur('CREDIT', PInfoTW(vNode.Item[i].Data)^.Credit);
    lTobFille.AddChampSupValeur('MOUVEMENT', PInfoTW(vNode.Item[i].Data)^.Mouvement);

    lTotDeb := lTotDeb + PInfoTW(vNode.Item[i].Data)^.Debit;
    lTotCre := lTotCre + PInfoTW(vNode.Item[i].Data)^.Credit;
    lTotMvt := lTotMvt + PInfoTW(vNode.Item[i].Data)^.Mouvement;
  end;

  lTobFille := Tob.Create('', lTob, -1);
  if lP^.TypeNode <> tnExo then
  begin
    lTobFille.AddChampSupValeur('LIBELLE', 'Total ' + lP^.Libelle);
    lTobFille.AddChampSupValeur('DEBIT',  lTotDeb);
    lTobFille.AddChampSupValeur('CREDIT', lTotCre);
    lTobFille.AddChampSupValeur('MOUVEMENT', lTotMvt);
  end
  else
  begin
    lTobFille.AddChampSupValeur('LIBELLE', '');
    lTobFille.AddChampSupValeur('DEBIT',  '');
    lTobFille.AddChampSupValeur('CREDIT', '');
    lTobFille.AddChampSupValeur('MOUVEMENT', '');
  end;

  lTobFille := Tob.Create('', lTob, -1);
  lTobFille.AddChampSupValeur('LIBELLE', 'Total ' + PinfoTw(TW.TopItem.Data)^.Libelle);
  lTobFille.AddChampSupValeur('DEBIT',  PinfoTw(TW.TopItem.Data)^.Debit);
  lTobFille.AddChampSupValeur('CREDIT', PinfoTw(TW.TopItem.Data)^.Credit);
  lTobFille.AddChampSupValeur('MOUVEMENT', PinfoTw(TW.TopItem.Data)^.Mouvement);

  lTob.PutGridDetail(FListe, False, False, 'LIBELLE;DEBIT;CREDIT;MOUVEMENT');

  FreeAndNil(lTob);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. : Pour dessiner le + dans le TreeNode
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.AjouteItemVide( vNode : TTreeNode) ;
var lP : PInfoTW;
begin
  New(lP);
  CPInitPInfoTW( lP );
  lP^.InNiv        := 0;
  lP^.BoVide       := True;
  lP^.BoASupprimer := True;
  TW.Items.AddChildObject(vNode,'<Création>' ,lP);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.VideTW ;
var i : integer ;
 lP : PInfoTW ;
begin
 for i := 0 to TW.Items.Count - 1  do
  begin
   lP := TW.Items[i].Data ;
   if lP <> nil then Dispose(lP) ;
   TW.Items[i].Data :=  nil ;
  end ; // for
 //if assigned(TW.Items[i].Data) then Dispose(TW.Items[i].Data) ;
  TW.Items.Clear ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.SuppItemVide( vNode : TTreeNode);
var lP    : PInfoTW ;
    lNode : TTreeNode ;
begin
  lNode := vNode.getFirstChild;
  if lNode = nil then Exit;
  lP := lNode.Data;
  if not assigned(lP) then Exit;
  if lP^.BoASupprimer then
  begin
    Dispose(lP);
    lNode.Delete;
  end; // if
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.RemplirTW( vNode : TTreeNode) ;
var lP : PInfoTW ;
begin
  lP := vNode.Data ;
  if not assigned(lP) then Exit ;
  if not lP^.BoVide then Exit ;

  if FTriTw = ttPeriode then
  begin
    case lP^.InNiv of
      0 : RemplirTreeviewPeriode(vNode);
      1 : RemplirTreeviewPeriodeJal(vNode);
    else
    end;
  end
  else
  begin
    case lP^.InNiv of
      0 : RemplirTreeviewJal(vNode);
      1 : RemplirTreeviewJalPeriode(vNode);
    else
    end;  
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.TWClick(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;
  if TW.Selected = nil then Exit;
  RemplirFListe(TW.Selected);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.ChangeTri( vTriTw : TTriTw ) ;
begin
  MMAFFPeriode.Checked := (vTriTw = ttPeriode);
  MMAFFJal.Checked     := (vTriTw = ttJal);
  if FTriTw = vTriTw then Exit;
  FTriTw := vTriTw ;
  SetParamSoc('SO_CPJALCENTRALSURJOURNAL', FTriTw = TTJal);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.EnabledControl ;
var lP : PInfoTW ;
begin
  if not assigned(TW.Selected) then Exit ;
  if not assigned(TW.Selected.Data) then Exit;

  lP := PInfoTW(TW.Selected.Data);

  Self.Caption := _FabriqTitre(lP, FTriTw);
  UpdateCaption(Self);

  BNew.Enabled  := False;
  BZoom.Enabled := False;

  if lP^.TypeNode = tnExo then Exit;

  if FTriTw = ttPeriode then
  begin
    if lP^.TypeNode = tnPer then
    begin
      BNew.Hint := cStNouveauJournal;
      BNew.Enabled := True;
    end
    else
    if lP^.TypeNode = tnJal then
    begin
      BZoom.Enabled := True;
      // GCO - 07/06/2006 - FQ 18271
      if (lP^.J_NATUREJAL <> 'ANO') then
      begin
        if (not lP^.BoClosDefinitif) and (not lP^.BoClosPeriodique) and (not lP^.BoFerme) then
        begin
          if (lP^.J_MODESAISIE <> '-') and (lP^.J_MODESAISIE <> '') then
            BNew.Hint := cStNouveauBor
          else
            BNew.Hint := cStNouvellePiece;
          BNew.Enabled := True;
        end;
      end;
    end;
  end
  else
  begin // Tri sur Journal
    if lP^.TyPeNode = tnJal then
    begin // Sur un Noeud Journal
      BZoom.Enabled := True;
      BNew.Hint := cStNouveauJournal;
      BNew.Enabled := True;
    end
    else
    begin // Sur un Noeud Période
      if lP^.TypeNode = tnPer then
      begin
        if (not lP^.BoClosDefinitif) and (not lP^.BoClosPeriodique) and (not lP^.BoFerme) then
        begin
          // GCO - 07/06/2006 - FQ 18271
          if (lP^.J_NATUREJAL <> 'ANO') then
          begin
            if (lP^.J_MODESAISIE <> '-') and (lP^.J_MODESAISIE <> '') then
              BNew.Hint := cStNouveauBor
            else
              BNew.Hint := cStNouvellePiece;
            BNew.Enabled := True;
          end;
        end;
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.PopF11Popup(Sender: TObject);
var lP : PInfoTW ;
begin
  AffPeriode.Checked   := MMAFFPeriode.Checked;
  AffJournal.Checked   := MMAFFJal.Checked;
  AffNoeudVide.Checked := MMAFFNoeudVide.Checked;

  if not assigned(TW.Selected) then Exit ;
  if not assigned(TW.Selected.Data) then Exit;

  lP := PInfoTW(TW.Selected.Data);
  AccesDetail.Enabled := (FListe.RowCount > 3);

  AccesDetail.Caption := 'Accéder au détail';

  if lP^.InNiv = 2 then
  begin
    if StrToInt(FListe.Cells[cColNumeroPiece, FListe.Row]) > 0 then
    begin
      if (lP^.J_MODESAISIE <> '-') and (lP^.J_MODESAISIE <> '') then
        AccesDetail.Caption := AccesDetail.Caption + ' du folio '
      else
        AccesDetail.Caption := AccesDetail.Caption + ' de la pièce ';
      AccesDetail.Caption := AccesDetail.Caption + FListe.Cells[cColNumeroPiece, FListe.Row];
    end;  
  end;

  NouvelElement.Caption := BNew.Hint;
  NouvelElement.Enabled := BNew.Enabled;
  AccesJournal.Enabled  := BZoom.Enabled;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.FormKeyDown(Sender: TObject; var Key: Word ; Shift: TShiftState);
begin
 if csDestroying in ComponentState then Exit ;

  case Key of

    VK_F5  : if (FListe.Focused) or (Tw.Focused) then
               DoubleCliqueFListe;

    VK_F9  : BChercheClick(nil);

    VK_F11 : PopF11.Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);

    VK_ESCAPE : if Shift=[] then
                 begin
                  key := 0 ;
                  Close ;
                  if IsInside(Self) then
                   begin
                    CloseInsidePanel(Self) ;
                    THPanel(Self.parent).InsideForm := nil;
                    THPanel(Self.parent).VideToolBar;
                   end ;  
                 end ; 
    (*74 : if Shift = [ssAlt] then
         begin
           Key := 0;
           ZoomJal;
         end;*)

    78 : if Shift = [ssCtrl] then // Ctrl + N
         begin
           Key := 0;
           Creation;
         end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.TWChange(Sender: TObject; Node: TTreeNode);
begin
 if csDestroying in ComponentState then Exit ;
 EnabledControl ;
 TWClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.MMAFFPERIODEClick(Sender: TObject) ;
begin
  ChangeTri(ttPeriode);
  RemplirTreeviewExo;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.MMAFFJALClick(Sender: TObject);
begin
  ChangeTri(ttJal);
  RemplirTreeviewExo;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.MMAFFNOEUDVIDEClick(Sender: TObject);
begin
  MMAFFNoeudVide.Checked := not MMAFFNoeudVide.Checked;
  FBoAfficheVide := MMAFFNoeudVide.Checked;
  RemplirTreeviewExo;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.AffNoeudVideClick(Sender: TObject);
begin
  MMAFFNOEUDVIDEClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
(*
function TFJournalCentralisateur.GetLibJal( vP : PInfoTW ) : string ;
begin
  if not MMAFFJAL.Checked then
    Result := vP^.J_JOURNAL
  else
    Result := vP^.J_JOURNAL + ' - ' + vP^.J_LIBELLE ;
end;

////////////////////////////////////////////////////////////////////////////////
function TFJournalCentralisateur.GetLibExo( vP : PInfoTW ) : string ;
begin
  if not MMCodeExo.Checked then  
    Result := cStTitreTWExo + vP^.TExo.Code
  else
    Result := vP^.Libelle;
end;
*)
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.MMZoomJalClick(Sender: TObject);
begin
  ZoomJal;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.AccesJournalClick(Sender: TObject);
begin
  MMZoomJalClick(nil);
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.ZoomJal ;
var lP      : PInfoTW ;
    lAction : TActionFiche ;
begin
  if TW.Selected = nil then exit ;
  lP := PInfoTW(TW.Selected.Data) ;
  if lP = nil then exit ;

  if lP^.J_JOURNAL = '' then exit ;
  lAction := taModif ;
  if not ExJaiLeDroitConcept(TConcept(ccJalModif), FALSE) then lAction := taConsult ;

  FicheJournal(nil, '', lP^.J_JOURNAL,lAction,0) ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.FListeDblClick(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;
  DoubleCliqueFListe;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.AccesDetailClick(Sender: TObject);
begin
  DoubleCliqueFListe;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.DoubleCliqueFListe;
var lPRoot : PInfoTW;
begin
  if not assigned(TW.Selected) then Exit;
  if not Assigned(TW.Selected.Data) then Exit;

  lPRoot := TW.Selected.Data;

  if lPRoot^.InNiv = 2 then
  begin
    if StrToInt(FListe.Cells[cColNumeroPiece, FListe.Row]) > 0 then
    begin
      OuvertureSaisie(TW.Selected, False);
    end;
  end
  else
  begin
    if FListe.Row-1 < Tw.Selected.Count then
    begin
      TW.Selected.Expand(False);
      TW.Selected.Item[FListe.Row-1].Selected := True;
      TW.Selected.Expand(False);
      RemplirFListe(TW.Selected);
      TW.SetFocus;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/03/2006
Modifié le ... : 09/05/2007
Description .. : - LG - 14/04/2006 - FB 17805 - on presente le dernier jour 
Suite ........ : du mois pour la saisie
Suite ........ : - LG - 27/04/2006 - FB 17873 -  Positionner le numéro de 
Suite ........ : folio  à + 1 et être placé dedans en position de saisie (dans 
Suite ........ : la grille). Attention le paramètre journal (incrémenter le 
Suite ........ : bordereau) ne doit pas entre en ligne de compte. C'est +1 
Suite ........ : par rapport au dernier numéro utilisé dans le couple 
Suite ........ : journal/période. 
Suite ........ : - LG - 02/06/2006 - FB 18216 - on passe le derneir jour du 
Suite ........ : mois pour ne ps ettre gene par les exercice decale (ne 
Suite ........ : commencant pas au 1 du mois )
Suite ........ : - LG - 09/05/2007 - FB 19097 - on utilise la fct 
Suite ........ : GetNewNumJal pour calculer le prochain numereo de folio 
Suite ........ : et on passe l'etablissment a la saisir bordereau
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.OuvertureSaisie( vNode : TTreeNode ; vBoCreation: Boolean);
var lP : PInfoTw;
    lQuery : TQuery;
    lParFolio : RParFolio;
    lInMaxJour : integer ;
    lStValeur  :string ;
begin
  if not assigned(vNode) then Exit;
  if not Assigned(vNode.Data) then Exit;

  {$IFNDEF EAGLCLIENT}
  V_Pgi.ToolBarreUp := FBoToolBarUp;
 {$ENDIF}

  lP := vNode.Data;

  if vBoCreation then
  begin
    if (lP^.J_MODESAISIE <> '-') and (lP^.J_MODESAISIE <> '') then
    begin
      FillChar(lParFolio, Sizeof(lParFolio), #0);
      lInMaxJour                := CGetDaysPerMonth(DateToStr(EncodeDate(lP^.InAnnee, lP^.InMois, 1)),lP^.TExo) ;
      lParFolio.ParPeriode      := DateToStr(EncodeDate(lP^.InAnnee, lP^.InMois, lInMaxJour));
      lParFolio.ParCodeJal      := lP^.J_Journal;
      lParFolio.ParCreatCentral := True;
      lParFolio.ParNumFolio     := intToStr( GetNewNumJal(lP^.J_Journal,true,EncodeDate(lP^.InAnnee, lP^.InMois, lInMaxJour) ) - 1 ) ;
      lStValeur                 := ETABLISSEMENT.Value ;
      lParFolio.ParEta          := ReadTokenSt(lStValeur) ;
      ChargeSaisieFolio(lParFolio, TaCreat);
    end
    else
     begin
      lInMaxJour := CGetDaysPerMonth(DateToStr(EncodeDate(lP^.InAnnee, lP^.InMois, 1)),lP^.TExo) ;
      MultiCritereMvt(taCreat, 'N;'+lP^.J_Journal+';'+DateToStr(EncodeDate(lP^.InAnnee, lP^.InMois, lInMaxJour)), False);
     end ;
  end
  else
  begin // Mode Consultation
    try
      lQuery := OpenSql('SELECT E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, ' +
                  'E_NUMLIGNE FROM ECRITURE WHERE ' +
                  'E_EXERCICE = "' + lP^.TExo.Code + '" AND ' +
                  'E_DATECOMPTABLE >= "' + USDateTime(EncodeDate(lP^.InAnnee, lP^.InMois, 1)) + '" AND ' +
                  'E_DATECOMPTABLE <= "' + USDateTime(FinDeMois(EncodeDate(lP^.InAnnee, lP^.InMois, 1))) + '" AND ' +
                  'E_JOURNAL = "' + lP^.J_Journal + '" AND ' +
                  'E_NUMEROPIECE = ' + FListe.Cells[cColNumeroPiece, FListe.Row] + ' AND ' +
                  'E_NUMLIGNE = 1 AND E_NUMECHE <= 1 AND E_QUALIFPIECE = "N" ' +
                  'ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE', True);

      if (lP^.J_MODESAISIE <> '-') and (lP^.J_MODESAISIE <> '') then
        LanceSaisieFolio(lQuery, taModif,true)
      else
       TrouveEtLanceSaisie(lQuery, taModif, 'N');

    finally
      Ferme(lQuery);
    end;
  end;

  {$IFNDEF EAGLCLIENT}
  V_Pgi.ToolBarreUp := false ;
 {$ENDIF}

  // Mise à jour du Solde
  //RemplitFListeAvecMouvement(vNode);

  BchercheClick(nil);

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.OnPostDrawCellFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var TheRect : TRect;

begin
  If ARow<=0 Then Exit ;
  if (FListe.ColTypes[ACol] = 'R') and
     ((Trim(FListe.Cells[ACol, ARow]) = '') or
      (Valeur(FListe.Cells[ACol, ARow]) = 0)) then
  begin
    TheRect := FListe.CellRect(ACol, ARow);
    Canvas.TextRect(TheRect, TheRect.Left, TheRect.Top, '');
    Canvas.Brush.Color := FListe.FixedColor;
    Canvas.Brush.Style := bsBDiagonal;
    Canvas.Pen.Color := FListe.FixedColor;
    Canvas.Pen.Mode := pmCopy;
    Canvas.Pen.Style := psClear;
    Canvas.Pen.Width := 1;
    Canvas.Rectangle(TheRect);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.OnGetCellCanvasFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if aRow = 0 then Exit;
  if (FListe.RowCount - FListe.FixedRows - aRow) < 2 then
    Canvas.Font.Style := [fsBold];
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/02/2006
Modifié le ... : 09/05/2007
Description .. : - LG - 09/05/2007 - FB 19405 - on reaffecte la valeur 
Suite ........ : original a la sortie
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 {$IFNDEF EAGLCLIENT}
  V_Pgi.ToolBarreUp := FBoToolBarUp;
 {$ENDIF}
 VideTW ;
 Self.caption:='' ;
 UpdateCaption(self) ;
 if ( Parent is THPanel ) then Action := caFree ;
end;

///////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/02/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.NouvelElementClick(Sender: TObject);
begin
  Creation;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.AFFPeriodeClick(Sender: TObject);
begin
  MMAFFPeriodeClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.AFFJournalClick(Sender: TObject);
begin
  MMAFFJALClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.BNewClick(Sender: TObject);
begin
  Creation;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.BChercheClick(Sender: TObject);
begin
  RemplirTreeViewExo;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.Creation;
var lP : PInfoTW ;
begin
  if not assigned(TW.Selected) then Exit;
  if not assigned(TW.Selected.Data) then Exit;

  lP := PInfoTW(TW.Selected.Data);

  if FTriTw = ttPeriode then
  begin
    if lP^.TypeNode = tnPer then
    begin
      FicheJournal(nil, '', '' , TaCreat, 0);
    end
    else
    if lP^.TypeNode = tnJal then
    begin
      if (not lP^.BoClosDefinitif) and (not lP^.BoClosPeriodique) and (not lP^.BoFerme) then
        OuvertureSaisie( TW.Selected, True);
    end;
  end
  else
  begin // Tri sur Journal
    if lP^.TyPeNode = tnJal then
    begin // Sur un Noeud Journal
    //  if lP^.BoAvecFille then
    //  begin
    //    if (not lP^.BoClos) and (not lP^.BoValide) and (not lP^.BoFerme) then
    //    SaisieEcriture(lP)
    //  end
    //  else
        FicheJournal(nil, '', '' , TaCreat, 0);
    end
    else
    begin // Sur un Noeud Période
      if lP^.TypeNode = tnPer then
      begin
        if (not lP^.BoClosDefinitif) and (not lP^.BoClosPeriodique) and (not lP^.BoFerme) then
          OuvertureSaisie( TW.Selected, True);
      end;
    end;
  end;

  // GCO - 23/06/2006 - FQ 18463
  RemplirTreeViewExo;
end;

(*
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.SaisieEcriture ( vP : PInfoTW ) ;
var lParFolio : RParFolio;
begin
  if not assigned(vP) then Exit;

  FillChar(lParFolio, Sizeof(lParFolio), #0);
  lParFolio.ParPeriode := DateToStr(EncodeDate(vP^.InAnnee, vP^.InMois, 1));
  lParFolio.ParCodeJal := vP^.J_Journal;
  lParFolio.ParCentral := True;

  if (vP^.J_MODESAISIE <> '-') and (vP^.J_MODESAISIE <> '') then
    ChargeSaisieFolio(lParFolio, TaCreat)
  else
    MultiCritereMvt(taCreat, 'N', False);
end;
*)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.JALCECRClick(Sender: TObject);
var lP : PInfoTw;
    lCritEdt : ClassCritEdt;
begin
{$IFDEF COMPTA}
  if not assigned(TW.Selected) then Exit;
  if not assigned(TW.Selected.Data) then Exit;

  lP := PInfoTW(TW.Selected.Data);

  lCritEdt := ClassCritEdt.Create;
  Fillchar(lCritEdt.CritEdt, SizeOf(lCritEdt.CritEdt), #0);
  lCritEdt.CritEdt.Cpt1     := lP^.J_JOURNAL;

  lCritEdt.CritEdt.Exo.Code := lP^.TExo.Code;
  lCritEdt.CritEdt.Date1    := lP^.TExo.Deb;
  lCritEdt.CritEdt.Date2    := lP^.TExo.Fin;

  if lP^.InAnnee <> 0 then
  begin
    lCritEdt.CritEdt.Date1    := EncodeDate(lP^.InAnnee, lP^.InMois, 1);
    lCritEdt.CritEdt.Date2    := FinDeMois(lCritEdt.CritEdt.Date1);
  end;

  TheData := lCritEdt;
  CPLanceFiche_CPJALECR;
  FreeAndNil(lCritEdt);
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.JALCENTRALClick(Sender: TObject);
{$IFDEF COMPTA}
var lP : PInfoTw;
    lCritEdt : ClassCritEdt;
{$ENDIF}    
begin
{$IFDEF COMPTA}
  if not assigned(TW.Selected) then Exit;
  if not assigned(TW.Selected.Data) then Exit;

  lP := PInfoTW(TW.Selected.Data);

  lCritEdt := ClassCritEdt.Create;
  Fillchar(lCritEdt.CritEdt, SizeOf(lCritEdt.CritEdt), #0);
  lCritEdt.CritEdt.Cpt1     := lP^.J_JOURNAL;
  lCritEdt.CritEdt.Cpt2     := lP^.J_JOURNAL;

  lCritEdt.CritEdt.Exo.Code := lP^.TExo.Code;
  lCritEdt.CritEdt.Date1    := lP^.TExo.Deb;
  lCritEdt.CritEdt.Date2    := lP^.TExo.Fin;

  if lP^.InAnnee <> 0 then
  begin
    lCritEdt.CritEdt.Date1    := EncodeDate(lP^.InAnnee, lP^.InMois, 1);
    lCritEdt.CritEdt.Date2    := FinDeMois(lCritEdt.CritEdt.Date1);
  end;

  TheData := lCritEdt;
  if FTriTw = TTPeriode then
    AGLLanceFiche('CP', 'EPJALGEN', '', 'JCD', 'JAC')
  else
    AGLLanceFiche('CP', 'EPJALGEN', '', 'JCJ', 'JAC');

  FreeAndNil(lCritEdt);
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.BALGENEClick(Sender: TObject);
{$IFDEF COMPTA}
var lP : PInfoTw;
    lCritEdt : ClassCritEdt;
{$ENDIF}
begin
{$IFDEF COMPTA}
  if not assigned(TW.Selected) then Exit;
  if not assigned(TW.Selected.Data) then Exit;

  lP := PInfoTW(TW.Selected.Data);
  lCritEdt := ClassCritEdt.Create;
  Fillchar(lCritEdt.CritEdt, SizeOf(lCritEdt.CritEdt), #0);

  lCritEdt.CritEdt.Exo.Code := lP^.TExo.Code;
  lCritEdt.CritEdt.Date1    := lP^.TExo.Deb;
  lCritEdt.CritEdt.Date2    := lP^.TExo.Fin;

  if lP^.InAnnee <> 0 then
  begin
    lCritEdt.CritEdt.Date1 := EncodeDate(lP^.InAnnee, lP^.InMois, 1);
    lCritEdt.CritEdt.Date2 := FinDeMois(lCritEdt.CritEdt.Date1);
  end;

  TheData := lCritEdt;
  CPLanceFiche_BalanceGeneral;
  FreeAndNil(lCritEdt);
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFJournalCentralisateur.GLECRClick(Sender: TObject);
{$IFDEF COMPTA}
var lP : PInfoTw;
    lCritEdt : ClassCritEdt;
{$ENDIF}
begin
{$IFDEF COMPTA}
  if not assigned(TW.Selected) then Exit;
  if not assigned(TW.Selected.Data) then Exit;

  lP := PInfoTW(TW.Selected.Data);
  lCritEdt := ClassCritEdt.Create;
  Fillchar(lCritEdt.CritEdt, SizeOf(lCritEdt.CritEdt), #0);

  lCritEdt.CritEdt.Exo.Code := lP^.TExo.Code;
  lCritEdt.CritEdt.Date1    := lP^.TExo.Deb;
  lCritEdt.CritEdt.Date2    := lP^.TExo.Fin;

  if lP^.InAnnee <> 0 then
  begin
    lCritEdt.CritEdt.Date1 := EncodeDate(lP^.InAnnee, lP^.InMois, 1);
    lCritEdt.CritEdt.Date2 := FinDeMois(lCritEdt.CritEdt.Date1);
  end;

  TheData := lCritEdt;
  CPLanceFiche_CPGLGENE('');
  FreeAndNil(lCritEdt);
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.TWExpanded(Sender: TObject; Node: TTreeNode);
begin
  if csDestroying in ComponentState then Exit ;
  if Node.Data = nil then Exit;
  RemplirTW(Node);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.AppliquerLesCriteresClick(Sender: TObject);
begin
  BChercheClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFJournalCentralisateur.BAbandonClick(Sender: TObject);
begin
  Close;
  if IsInside (Self) then THPanel (parent) .CloseInside;
end;

////////////////////////////////////////////////////////////////////////////////

end.
