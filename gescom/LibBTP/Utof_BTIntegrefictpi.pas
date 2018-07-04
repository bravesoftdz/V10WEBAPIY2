{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 16/03/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : INTEGREFICTPI ()
Mots clefs ... : TOF;INTEGREFICTPI
*****************************************************************}
Unit Utof_BTIntegrefictpi ;

Interface

Uses vierge,
     StdCtrls,
     Controls, 
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     FE_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms,
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Utob,
     HTB97,
     Windows,
     UtilsEtat,
		 hPDFPrev,
     hPDFViewer,
     uPDFBatch,
     Graphics,
     Grids,
     uEntCommun,
     UTOF ;

Type
  TOF_INTEGREFICTPI = Class (TOF)
    //
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
//
  Private
    //
    FichierImport   : THEdit;
    DateMvt         : THEdit;
    Analyse         : THEdit;
    Depot           : THEdit;
    //
    Client          : THEdit;
    Affaire         : THEdit;
    Affaire0        : THEdit;
    Affaire1        : THEdit;
    Affaire2        : THEdit;
    Affaire3        : THEdit;
    Avenant         : THEdit;
    CAB_Affaire     : THEdit;
    //
    CodeTiers       : THEdit;
    CAB_Tiers       : THEdit;
    Juridique       : THLabel;
    Nom             : THLabel;
    ADR1            : THLabel;
    ADR2            : THLabel;
    ADR3            : THLabel;
    ADR4            : THLabel;
    CP              : THLabel;
    Ville           : THLabel;
    //
    TNumCde         : THLabel;
    TNumBL          : THLabel;
    TNumBesoin      : THLabel;
    NumCde          : THEdit;
    NumBL           : THEdit;
    NumBesoin       : THEdit;

    //
    BTSelectNumCde  : TToolBarButton97;
    BTSelectNumBL   : TToolBarButton97;
    BTSelectBesoin  : TToolBarButton97;
    BtSelectLigne   : TToolBarButton97;
    //
    CodeArticle     : THEdit;
    Article         : THEdit;
    Qte             : THNumEdit;
    CoefUSUV        : THEdit;
    CoefUAUS        : THEdit;
    PAHT            : THEdit;
    PAUA            : THEdit;
    DPA             : THEdit;
    PMAP            : THEdit;
    CAB_Article     : THEdit;
    RefInterne      : THEdit;
    RefExterne      : THEdit;
    //
    Aff_Ferme       : TCheckBox;
    Frs_Ferme       : TCheckBox;
    Art_Ferme       : TCheckBox;
    Art_TenueStock  : TCheckBox;
    ChkIntegre      : TCheckBox;
    ChkConsoFact    : TCheckBox;
    //
    Lib_Depot       : THLabel;
    Lib_Affaire     : THLabel;
    TGPAFFAIRE3     : THLabel;
    Lib_Tiers       : THLabel;
    Lib_Article     : THLabel;
    Lib_Responsable : THLabel;
    Lib_Domaine     : THLabel;
    LIB_TenueStock  : THLabel;
    CaptionErreur   : THLabel;
    Lib_RefInterne  : THLabel;
    Lib_RefExterne  : THLabel;
    Lib_TypeAffaire : THLabel;

    //
    PageGrille      : TPageControl;
    PToutes         : TTabSheet;
    PAnomalies      : TTabSheet;
    //
    Pages           : TPageControl;
    PGeneralite     : TTabSheet;
    PAffaire        : TTabSheet;
    PFournisseur    : TTabSheet;
    PArticle        : TTabSheet;
    PEdition        : TTabSheet;
    PInventaire     : TTabSheet;
    //
    TypeMVT         : THValComboBox;
    Art_QualifCAB   : THValComboBox;
    Frs_QualifCAB   : THValComboBox;
    Aff_QualifCAB   : THValComboBox;
    QualifUS        : THValComboBox;
    QualifUA        : THValComboBox;
    QualifUV        : THValComboBox;
    TypeAffaire     : THValCombobox;
    //
    OkAccepte       : Boolean;
    //
    FichierHisto    : String;
    Question        : String;
    //
    Grille          : THGrid;
    Grille1         : THGrid;
    Grille2         : THGrid;
    Grille3         : THGrid;
    //
    BesoinWin       : TToolWindow97;
    //
    FF              : string;
    //
    fColNamesGrille     : string;
    Falignement         : string;
    Ftitre              : string;
    fLargeur            : string;
    //
    fColNamesGrille1    : string;
    FalignementGrille1  : string;
    FtitreGrille1       : string;
    fLargeurGrille1     : string;
    //
    fColNamesGrille2    : string;
    FalignementGrille2  : string;
    FtitreGrille2       : string;
    fLargeurGrille2     : string;
    //
    fColNamesGrille3    : string;
    FalignementGrille3  : string;
    FtitreGrille3       : string;
    fLargeurGrille3     : string;
    //
    Domaine             : string;
    Etablissement       : string;
    LibErreur           : String;
    //
    QQ                  : TQuery ;
    StSQL               : String ;
    //
    TOBGrille       : TOB;
    TobAnomalie     : TOB;
    TobValide       : TOB;
    TobRejete       : TOB;
    TobInventaire   : TOB;
    TobStock        : TOB;
    TobBesoin       : TOB;
    //
    TOBPiece        : TOB;
    TobLigne        : TOB;
    //
    TOBAffaire      : TOB;
    TOBTiers        : TOB;
    TobArticle      : TOB;
    TobDepot        : TOB;
    //
    BtImport        : TToolBarButton97;
    //
    BtSelect1       : TToolBarButton97;
    BtSelect2       : TToolBarButton97;
    BtSelect3       : TToolBarButton97;
    //
    BtEfface1       : TToolBarButton97;
    BtEfface2       : TToolBarButton97;
    BtEfface3       : TToolBarButton97;
    //
    BtSelectCAB1    : TToolBarButton97;
    BtSelectCAB2    : TToolBarButton97;
    BtSelectCAB3    : TToolBarButton97;
    //
    BGeneration     : TToolBarButton97;

    //Variable névcessaire pour la gestion de l'état
    BParamEtat    : TToolBarButton97;
    //
    ChkApercu     : TCheckBox;
    ChkReduire    : TCheckBox;
    ChkCouleur    : TCheckBox;
    //
    FETAT         : THValComboBox;
    TEtat         : ThLabel;
    //
    OptionEdition : TOptionEdition;
    TheType       : String;
    TheNature     : String;
    TheTitre      : String;
    TheModele     : String;
    //
    FETAT1        : THValComboBox;
    TEtat1        : ThLabel;
    BParamEtat1   : TToolBarButton97;
    //
    OptionEdition1: TOptionEdition;
    TheType1      : String;
    TheNature1    : String;
    TheTitre1     : String;
    TheModele1    : String;
    //
    Idef          : Integer;
    //
    Cancel        : Boolean;
    //
    //
    //
    procedure AdditionQteInventaire(CodeArticle: String; QteArt : Double; NoLig : Integer);
    procedure AffichageGrilleAnomalie;
    procedure AffichageGrilleInventaire;
    procedure AnalyseOnExit(Sender: Tobject);
    function  ArchivagefichierTPI: Boolean;
    //
    procedure BtEffaceOnclick(Sender: TObject);
    procedure BTGenerationOnClick(Sender: TObject);
    procedure BtImportOnClick(Sender : TObject);
    procedure BTSelectBesoinOnClick(Sender: TObject);
    procedure BtSelectCABOnClick(Sender: TObject);
    procedure BtSelectLigneOnClick(sender: TObject);
    procedure BTSelectNumCdeOnClick(Sender: TObject);
    procedure BtSelectOnClick(Sender: TObject);
    //
    procedure ChargeBL(TOBL: TOB; CodeBarre, NaturePiece : String);
    procedure ChargeEtatRejete;
    procedure ChargeEtatValide;
    procedure ChargeGrilleToScreen(Arow: Integer; TOBL : TOB);
    procedure ChargeInfoAffaire(TOBL : TOB);
    procedure ChargeInfoArticle(TOBL : TOB);
    procedure ChargeInfoTiers(TOBL : TOB);
    procedure ChargeLigneCommentaire(TypeMvt : string;TOBLIGNE, TOBL: TOB);
    procedure ChargementLigneTOB(TOBL: TOB; StLig: String);
    procedure ChargeTobwithScreen;
    procedure ChargementStockArticle(TobTempo: TOB);
    procedure ChargeTOB;
    procedure ChargeTobAffaire(TOBL : TOB; CodeBarre: string);
    procedure ChargeTobArticle(TOBL : TOB; CodeBarre: string);
    Function  ChargeTOBDepot(TOBL: TOB; CodeDepot: string) : Boolean;
    procedure ChargeTOBGRILLEwithTOBAFFAIRE(Tobl, TOBAffaire: TOB);
    procedure ChargeTOBGRILLEwithTOBARTICLE(Tobl, TOBArticle: TOB);
    procedure ChargeTOBGRILLEwithTOBTIERS(Tobl, TOBTiers: TOB);
    procedure ChargeTobLigne(STypeMVT : string; TOBL, TOBLigne: TOB);
    procedure ChargeTobTiers(TOBL : TOB; CodeBarre: string);
    procedure CloseBesoinWin(Sender: Tobject);
    Procedure CodeBarreOnExit(Sender : Tobject);
    procedure CodeOnExit(Sender : TObject);
    procedure ConsoFactOnClick(Sender: Tobject);
    procedure ControleCBT(TOBL: TOB);
    procedure ControleChamp(Champ, Valeur: String);
    function  ControleFichierArchi : Boolean;
    procedure ControleFichierImport;
    function  ControleRepertoire(DefaultPath : String) : String;
    function  ControleStockAvantIntegration(TobTempo: TOB): Boolean;
    procedure CreateEnteteDoc(Nature : String; TOBL : Tob);
    procedure CreateLigne(TobLigne: TOB);
    procedure CreateTOB;
    procedure CreateTOBPiece;
    //
    procedure DateMVTOnExit(Sender: Tobject);
    procedure DefinieGrid;
    procedure DepotOnExit(Sender: TObject);
    procedure DestroyTOB;
    //
    procedure EditeRejete;
    procedure EditeValide;
    //
    Function  GenerationDeLaPiece(TOBL: TOB; TypeMvt: String) : Boolean;
    function  GenerationMsgAnomalie(TOBL: TOB): String;
    procedure GestionAnomalie(TOBLGrille : TOB);
    procedure GestionGrilleInventaire;
    procedure GestionModif(Zone, NewCode, OldCode: string; IndLigne: Integer);
    function  GestionRejet(Tobl: TOB): Boolean;
    procedure GestionRuptureBesoin(TOBL: TOB; AncienNumero: String);
    procedure GestionRuptureChantier(TOBL: TOB; AncienAff: String);
    procedure GestionRuptureDate(TOBL: TOB; AncienDateMvt: TDatetime);
    procedure GestionRuptureFournisseur(TOBL: TOB; AncienFrs: String);
    procedure GestionRuptureNumCde(TOBL: TOB; AncienNumero: String);
    procedure GestionRuptureTypeMvt(TOBL: TOB; AncienTypeMvt: String; var LigToDel: Integer);
    procedure GestionValide(Tobl: TOB);
    procedure GetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas;AState: TGridDrawState);
    procedure GetObjects;
    procedure GrilleRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure Grille1RowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    //
    procedure ImportElipsis(Sender: TObject);
    procedure InitZoneAffaire;
    Procedure InitZoneAffaireToTob(TOBL : TOB);
    procedure InitZoneArticle;
    Procedure InitZoneArticleToTob(TOBL : TOB);
    Procedure InitZoneEcran;
    procedure InitzoneFournisseur;
    procedure InitZoneFournisseurToTob(TOBL: TOB);
    procedure InitZoneTOB(TOBL: TOB);
    procedure IntegreOnClick(Sender: Tobject);
    procedure InventaireDblClick(Sender: TObject);
    //
    function  MAJInventaire(TOBL : TOB) : Boolean;
    procedure MAJLigneInventaire(TOBResult, TOBL : TOB);
    procedure MajTobRejet(TOBL : TOB;TypeErr, Motif: String);
    procedure MajTobValide;
    procedure ModificationEnSerie(Zone, OldCode, Newcode: String; TOBL: TOB; IndLigne: Integer);
    procedure ModificationSimple(zone, Newcode: String; TOBL: TOB;IndLigne: Integer);
    //
    procedure NumeroOnExit(Sender: Tobject);
    //
    procedure OnChangeEtat(Sender: Tobject);
    procedure OnChangeEtat1(Sender: Tobject);
    procedure OnClickApercu(Sender: Tobject);
    procedure OnClickReduire(Sender: Tobject);
    //
    procedure ParamEtat(Sender: TOBJect);
    procedure ParamEtat1(Sender: TOBJect);
    //
    procedure QteOnExit(Sender: Tobject);
    //
    procedure RechercheAffaire;
    procedure RechercheArticle;
    procedure RechercheCABAffaire;
    procedure RechercheCABArticle;
    procedure RechercheCABFournisseur;
    procedure RechercheFournisseur;
    procedure RechercheInventaireArticle(TOBL : TOB);
    procedure RecupLignePiecePrecedente(TOBL, TobLigne: TOB);
    Procedure RefOnExit(Sender : Tobject);
    procedure ReinitialiseTobPiece;
    //
    procedure SelectLigneBesoin(Sender : TObject);
    procedure SetGrilleEvents(Etat: boolean);
    procedure SetScreenEvents;
    procedure SoustractionQteInv(CodeArticle: String; QteArt: Double; NoLig : Integer);

    procedure TriTableGrille;
    procedure TypeAffaireOnchange(Sender: Tobject);
    procedure TypeMVTOnchange(Sender: TObject);

  end ;
  //
  Type TGestionGS =
  Record
       fColNamesGS     : string;
       FalignementGS   : string;
       FtitreGS        : string;
       fLargeurGS      : string;
       GS              : THGrid;
       TOBG            : TOB;
  End;

  //
  procedure ChargementGrille(GestionGS : TGestionGS);
  procedure DessineGrille(GestionGS : TGestionGS);


  //
  Var GestionGS : TGestionGS;
      GestionG3 : TGestionGS;
  //
  const
    SG_TYPEMVT      = 1;
    SG_DATEMVT      = 2;
    SG_DEPOT        = 3;
    SG_FOURNISSEUR  = 4;
    SG_AFFAIRE      = 5;
    SG_ARTICLE      = 6;
    SG_QTE          = 7;
    SG_INTEGRE      = 8;

Implementation
uses Paramsoc,
     Dialogs,
     AffaireUtil,
     FactUtil,
     GerePiece,
     FileCtrl,
     FactComm;

procedure TOF_INTEGREFICTPI.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_INTEGREFICTPI.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_INTEGREFICTPI.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_INTEGREFICTPI.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_INTEGREFICTPI.OnArgument (S : String ) ;
var i       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
    ChoixRepertoire : string;
begin

  Inherited ;

  Cancel := false;

  Ecran.caption := 'Intégration fichier LSE';

  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  CreateTOB;
  //
  Critere := S;
  //
  While (Critere <> '') do
  BEGIN
    i:=pos(':',Critere);
    if i = 0 then i:=pos('=',Critere);
    if i <> 0 then
       begin
       Champ:=copy(Critere,1,i-1);
       Valeur:=Copy (Critere,i+1,length(Critere)-i);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;

  //Gestion des évènement des zones écran
  SetScreenEvents;

  //Gestion des évènement de Grille
  SetGrilleEvents(True);

  if V_PGI.OkDecQ>0 then
  begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecQ-1 do
    begin
       FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;

  PageGrille.ActivePage := PToutes;
  //
  PAnomalies.TabVisible   := False;
  //
  PGeneralite.TabVisible  := True;
  PInventaire.TabVisible  := False;

  //On se positionne sur le bon onglet
  Pages.ActivePage := PGeneralite;
  Pages.enabled := False;

  InitZoneEcran;

  // définition de la grille
  DefinieGrid;

  GestionGS.fColNamesGS   := fColNamesGrille;
  GestionGS.FalignementGS := Falignement;
  GestionGS.FtitreGS      := Ftitre;
  GestionGS.fLargeurGS    := fLargeur;
  GestionGS.GS            := Grille;

  //dessin de la grille
  DessineGrille(GestionGS);

  FichierImport.Text := ControleRepertoire(GetParamSocSecur('SO_REPRECUPFIC', '', False));

  if FichierImport.text = '' then
  begin
    if PGIAsk('Importation impossible le répertoire de récupération n''existe pas. Voulez-vous en sélectionner un ?', 'Erreur Répertoire Stockage')=Mryes then
    begin
      //SelectDirectory(ChoixRepertoire,[sdAllowCreate,sdPerformCreate,sdPrompt],0)
      If SelectDirectory('Sélection de répertoire', '', ChoixRepertoire) then
      begin
        FichierImport.Text := ChoixRepertoire + '\lsecab.txt';
      end;
    end
  end
  else
  begin
    fichierImport.text := FichierImport.text + '\lsecab.txt';
  end;

  ControleFichierImport;

  if FichierImport.Text = 'C:\' Then BtImport.Visible :=False;

  ChargeEtatRejete;

  ChargeEtatValide;

  OkAccepte := GetParamSocSecur('SO_BTINTERDIREACHATS', False);

end ;

Procedure TOF_INTEGREFICTPI.ControleFichierImport;
begin

  If not FileExists(FichierImport.text) then
  Begin
    PGIError('Aucun fichier à traiter...','Erreur Importation');
    FichierImport.Text := 'C:\';
  end;

end;

function TOF_INTEGREFICTPI.ControleFichierArchi : Boolean;
Var StDate    : String;
    StTime    : String;
    FileName  : String;
begin

  Result := False;

  StDate := DateToStr(NowH);
  StTime := TimeToStr(Nowh);
  //
  StDate := FindEtReplace(StDate,'/','',True);
  StTime := FindEtReplace(Sttime,':','',True);
  //
  FileName      := ExtractFileName(FichierImport.text);
  if FileName = '' then exit;
  FileName      := Copy(FileName,1, Length(FileName)-4);

  FichierHisto  := ControleRepertoire(GetParamSocSecur('SO_REPARCHFIC', '', False));
  FichierHisto  := FichierHisto + '\lsecab_' + StDate + '_' + StTime + '.Arc';

  if FileExists(FichierHisto) then
  begin
    if FichierHisto = '' then
    begin
      if PGIAsk('Archivage impossible le répertoire d''Archivage n''existe pas. Voulez-vous en sélectionner un ?', 'Erreur Répertoire Archivage')=Mryes then
      begin
        if not SelectDirectory(FichierHisto,[sdAllowCreate,sdPerformCreate,sdPrompt],0) then Exit;
      end
      else Exit;
    end;
    if ControleFichierArchi then Result := True;
  end
  else
    Result := True;

end;

Function TOF_INTEGREFICTPI.ControleRepertoire(DefaultPath : String) : String;
var ChoixRepertoire : string;
begin

  Result := defaultPath;

  if Result <> '' then
  begin
    if not DirectoryExists(Result) then Result := '';
  end
  else
  begin
    if SelectDirectory('Sélection de répertoire', '', ChoixRepertoire) then Result := ChoixRepertoire;
    //If SelectDirectory(ChoixRepertoire,[sdAllowCreate,sdPerformCreate,sdPrompt],0) Then Result := ChoixRepertoire;
  end;

end;

procedure TOF_INTEGREFICTPI.OnClose ;
begin
  Inherited ;

  DestroyTOB;

  FreeAndNil(OptionEdition);
  
end ;

procedure TOF_INTEGREFICTPI.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_INTEGREFICTPI.OnCancel () ;
begin
  Inherited ;
end ;

  //Chargement des zones ecran dans des zones programme
Procedure TOF_INTEGREFICTPI.GetObjects;
begin
  //
  Pages         := TPageControl(GetControl('PAGES'));
  PageGrille    := TPageControl(GetControl('PAGEGRILLE'));
  //
  Ptoutes       := TTabSheet(GetControl('PTOUTES'));
  PAnomalies    := TTabSheet(GetControl('PANOMALIES'));
  //
  PGeneralite   := TTabSheet(GetControl('GENERALITES'));
  PAffaire      := TTabSheet(GetControl('AFFAIRE'));
  PFournisseur  := TTabSheet(GetControl('FOURNISSEUR'));
  PArticle      := TTabSheet(GetControl('ARTICLE'));
  PEdition      := TTabSheet(GetControl('EDITION'));
  PInventaire   := TTabSheet(GetControl('INVENTAIRE'));
  //
  FichierImport := THEdit(GetControl('FICHIERIMPORT'));
  //
  TypeMVT       := THValcombobox(GetControl('TYPEMVT'));
  DateMvt       := THEdit(GetControl('DATEMVT'));
  Analyse       := THEdit(GetControl('ANALYSE'));
  Depot         := THEdit(GetControl('DEPOT'));
  //
  //chargement de la grille
  Grille        := THGrid(GetControl('GRILLE'));
  Grille1       := THGrid(GetControl('GRILLE1'));
  Grille2       := THGrid(GetControl('GRILLE2'));
  //
  //chargement des zones écran liées à l'affaire
  Client        := THEdit(GetControl('TIERS'));
  //
  Affaire       := THEdit(GetControl('AFF_AFFAIRE'));
  Affaire0      := THEdit(GetControl('AFFAIRE0'));
  Affaire1      := THEdit(GetControl('AFF_AFFAIRE1'));
  Affaire2      := THEdit(GetControl('AFF_AFFAIRE2'));
  Affaire3      := THEdit(GetControl('AFF_AFFAIRE3'));
  Avenant       := THEdit(GetControl('AFF_AVENANT'));
  //
  CAB_Affaire   := THEdit(GetControl('AFF_CODEBARRE'));
  //
  Lib_Depot       := THLabel(GetControl('LIBDEPOT'));
  Lib_Affaire     := THLabel(GetControl('LIBAFFAIRE'));
  TGPAFFAIRE3     := THLabel(GetControl('TGPAFFAIRE3'));
  Lib_Tiers       := THLabel(GetControl('LIBTIERS'));
  LIB_TenueStock  := THLabel(Getcontrol('LIBTENUESTOCK'));
  Lib_Responsable := THLabel(GetControl('LIB_RESPONSABLE'));
  Lib_Domaine     := THLabel(GetControl('LIB_DOMAINE'));
  CaptionErreur   := THLabel(GetControl('LIBERREUR'));
  //
  Aff_Ferme       := TCheckBox(GetControl('AFF_FERME'));
  //
  //chargement des zones écran liées au fournisseur
  //
  CodeTiers       := THEdit(GetControl('T_TIERS'));
  CAB_Tiers       := THEdit(GetControl('BT1_CODEBARRE'));;
  Juridique       := THLabel(GetControl('JURIDIQUE'));
  Nom             := THLabel(GetControl('T_NOM'));
  ADR1            := THLabel(GetControl('ADR1'));
  ADR2            := THLabel(GetControl('ADR2'));
  ADR3            := THLabel(GetControl('ADR3'));
  ADR4            := THLabel(GetControl('ADR4'));
  CP              := THLabel(GetControl('CP'));
  Ville           := THLabel(GetControl('VILLE'));
  Frs_Ferme       := TCheckBox(GetControl('T_FERME'));
  NumCde          := THEdit(GetControl('NUMCDE'));
  NumBL           := THEdit(GetControl('NUMBL'));
  NumBesoin       := THEdit(GetControl('NUMBESOIN'));
  TNumCde         := THLabel(GetControl('TNUMCDE'));
  TNumBL          := THLabel(GetControl('TNUMBL'));
  TNumBesoin      := THLabel(GetControl('TNUMBESOIN'));
  //
  //chargement des zones écran liées à l'article
  //
  CodeArticle     := THEdit(GetControl('GA_CODEARTICLE'));
  Article         := THEdit(GetControl('GA_ARTICLE'));
  Qte             := THNumEdit(GetControl('QTE'));
  CoefUSUV        := THEdit(GetControl('GA_COEFCONVQTEVTE'));
  CoefUAUS        := THEdit(GetControl('GA_COEFCONVQTEACH'));
  PAHT            := THEdit(GetControl('GA_PAHT'));
  PAUA            := THEdit(GetControl('GA_PAUA'));
  DPA             := THEdit(GetControl('GA_DPA'));
  PMAP            := THEdit(GetControl('GA_PMAP'));
  CAB_Article     := THEdit(GetControl('GA_CODEBARRE'));
  Lib_Article     := THLabel(GetControl('LIBARTICLE'));
  RefInterne      := THEdit(GetControl('REFINTERNE'));
  RefExterne      := THEdit(GetControl('REFEXTERNE'));
  Lib_RefInterne  := THLabel(GetControl('LIB_REFINTERNE'));
  Lib_RefExterne  := THLabel(GetControl('LIB_REFEXTERNE'));
  Lib_TypeAffaire := THLabel(GetControl('STTYPEAFFAIRE'));
  //
  Art_Ferme       := TCheckBox(Getcontrol('GA_FERME'));
  Art_TenueStock  := TcheckBox(GetControl('GA_TENUESTOCK'));
  ChkIntegre      := TcheckBox(GetControl('CHKINTEGRE'));
  ChkConsoFact    := TCheckBox(GetControl('CHKCONSOFACT'));
  //
  Art_QualifCAB   := THValComboBox(GetControl('GA_QUALIFCODEBARRE'));
  Frs_QualifCAB   := THValComboBox(GetControl('BT1_QUALIFCODEBARRE'));
  Aff_QualifCAB   := THValComboBox(GetControl('AFF_QUALIFCODEBARRE'));
  QualifUS        := THValComboBox(GetControl('GA_QUALIFUNITESTO'));
  QualifUA        := THValComboBox(GetControl('GA_QUALIFUNITEACH'));
  QualifUV        := THValComboBox(GetControl('GA_QUALIFUNITEVTE'));
  TypeAffaire     := THValComboBox(GetControl('TYPEAFFAIRE'));
  //
  BtImport        := TToolBarButton97(Getcontrol('BTIMPORT'));
  //
  BtSelect1       := TToolBarButton97(Getcontrol('BSelect1'));
  BtSelect2       := TToolBarButton97(Getcontrol('BSelect2'));
  BtSelect3       := TToolBarButton97(Getcontrol('BSelect3'));
  //
  BtEfface1       := TToolBarButton97(Getcontrol('BEfface1'));
  BtEfface2       := TToolBarButton97(Getcontrol('BEfface2'));
  BtEfface3       := TToolBarButton97(Getcontrol('BEfface3'));
  //
  BtSelectCAB1    := TToolBarButton97(Getcontrol('BSelectCAB1'));
  BtSelectCAB2    := TToolBarButton97(Getcontrol('BSelectCAB2'));
  BtSelectCAB3    := TToolBarButton97(Getcontrol('BSelectCAB3'));
  //
  BtSelectNUMCDE  := TToolBarButton97(Getcontrol('BSelectNUMCDE'));
  BtSelectNUMBL   := TToolBarButton97(Getcontrol('BSelectNUMBL'));
  BtSelectBESOIN  := TToolBarButton97(Getcontrol('BSelectBESOIN'));
  BtSelectLigne   := TToolBarButton97(Getcontrol('BSelectLIGNE'));
  //
  BGeneration     := TToolbarButton97(GetControl('BTGENERATION'));
  //
  BParamEtat      := TToolbarButton97(GetCONTROL('BPARAMETAT'));
  TEtat           := ThLabel(ecran.FindComponent('TEtat'));
  FEtat           := ThValComboBox(ecran.FindComponent('FEtat'));
  //
  BParamEtat1     := TToolbarButton97(GetCONTROL('BPARAMETAT1'));
  TEtat1          := ThLabel(ecran.FindComponent('TEtat1'));
  FEtat1          := ThValComboBox(ecran.FindComponent('FEtat1'));
  //
  ChkApercu       := TCheckBox(Ecran.FindComponent('fApercu'));
  ChkReduire      := TCheckBox(Ecran.FindComponent('fReduire'));
  ChkCouleur      := TCheckBox(Ecran.FindComponent('fCouleur'));
  //
end;
  //
Procedure TOF_INTEGREFICTPI.CreateTOB;
begin

  TobGrille     := Tob.create('LA GRILLE', nil, -1);
  TobAnomalie   := Tob.Create('ANOMALIES', nil, -1);
  TobInventaire := Tob.Create('INVENTAIRES', nil, -1);

  TOBRejete     := TOB.Create('REJETE', nil, -1);
  TOBValide     := TOB.Create('VALIDE', nil, -1);

  TOBPiece      := TOB.Create('LA PIECE', nil,  -1);

  TobStock      := TOB.Create('STOCK', nil,  -1);

end;

procedure TOF_INTEGREFICTPI.InitZoneTOB(TOBL : TOB);
begin

  //zone nécessaire à la grille
  TOBL.AddChampSupValeur('NOLIG', '');
  TOBL.AddChampSupValeur('TYPEMVT', '');
  TOBL.AddChampSupValeur('DATEMVT', iDate1900);
  TOBL.AddChampSupValeur('DEPOT', '');
  TOBL.AddChampSupValeur('LIB_DEPOT', '');
  TOBL.AddChampSupValeur('CODEAFFAIRE', '');
  TOBL.AddChampSupValeur('CODEFOURNISSEUR', '');
  TOBL.AddChampSupValeur('CODEARTICLE', '');
  TOBL.AddChampSupValeur('QTE', '0.00');
  TOBL.AddChampSupValeur('INTEGRE', 'X');
  //
  TOBL.AddChampSupValeur('ANALYSE', '');
  //
  TOBL.AddChampSupValeur('ERREUR',      '-');
  TOBL.AddChampSupValeur('ANOMALIE',    '');
  TOBL.AddChampSupValeur('ANOMALIEAFF', '');
  TOBL.AddChampSupValeur('ANOMALIEBC',  '');
  TOBL.AddChampSupValeur('LIBANOMALIEAFF', '');
  TOBL.AddChampSupValeur('LIBANOMALIEBC',  '');
  TOBL.AddChampSupValeur('ANOMALIEFRS', '');
  TOBL.AddChampSupValeur('LIBANOMALIEFRS', '');
  TOBL.AddChampSupValeur('ANOMALIEMAR', '');
  TOBL.AddChampSupValeur('LIBANOMALIEMAR', '');
  TOBL.AddChampSupValeur('ANOMALIECLI', '');
  TOBL.AddChampSupValeur('LIBANOMALIECLI', '');
  //
  TOBL.AddChampSupValeur('TYPEERREUR', '');

  //Zone nécessaire à l'affaire
  TOBL.AddChampSupValeur('CLIENT', '');
  TOBL.AddChampSupValeur('LIB_CLIENT', '...');

  TOBL.AddChampSupValeur('AFF_AFFAIRE', '');
  TOBL.AddChampSupValeur('LIB_AFFAIRE', '');
  //
  TOBL.AddChampSupValeur('CAB_Affaire',   '');
  TOBL.AddChampSupValeur('AFF_QUALIFCAB', '');
  //
  TOBL.AddChampSupValeur('Lib_Responsable', '...');
  TOBL.AddChampSupValeur('Domaine',        '');
  TOBL.AddChampSupValeur('Etablissement',  '');
  TOBL.AddChampSupValeur('Lib_Domaine', '...');
  //
  //zone nécessaire au fournisseur
  TOBL.AddChampSupValeur('T_TIERS',        '');
  TOBL.AddChampSupValeur('CAB_Tiers',      '');
  TOBL.AddChampSupValeur('Frs_QualifCAB',  '');
  TOBL.AddChampSupValeur('Juridique',   '...');
  TOBL.AddChampSupValeur('Nom',         '...');
  TOBL.AddChampSupValeur('ADR1',        '...');
  TOBL.AddChampSupValeur('ADR2',        '...');
  TOBL.AddChampSupValeur('ADR3',        '...');
  TOBL.AddChampSupValeur('ADR4',        '...');
  TOBL.AddChampSupValeur('CP',          '...');
  TOBL.AddChampSupValeur('Ville',       '...');

  //Zone nécessaire à l'article
  TOBL.AddChampSupValeur('GA_ARTICLE',     '');
  TOBL.AddChampSupValeur('Lib_Article', '...');
  //
  TOBL.AddChampSupValeur('CAB_Article',    '');
  TOBL.AddChampSupValeur('Art_QualifCAB',  '');
  //
  TOBL.AddChampSupValeur('QUALIFUNITESTO',  '');
  TOBL.AddChampSupValeur('QUALIFUNITEACH',  '');
  TOBL.AddChampSupValeur('QUALIFUNITEVTE',  '');
  //
  TOBL.AddChampSupValeur('CoefUSUV',       '');
  TOBL.AddChampSupValeur('CoefUAUS',       '');
  TOBL.AddChampSupValeur('PAHT',       '0.00');
  TOBL.AddChampSupValeur('PAUA',       '0.00');
  TOBL.AddChampSupValeur('DPA',        '0.00');
  TOBL.AddChampSupValeur('PMAP',       '0.00');

  TOBL.AddChampSupValeur('TENUESTOCK',    '-');
  TOBL.AddChampSupValeur('ART_FERME',     '-');
  TOBL.AddChampSupValeur('AFF_FERME',     '-');
  TOBL.AddChampSupValeur('FRS_FERME',     '-');

  TOBL.AddChampSupValeur('NUMBL',         '0');
  TOBL.AddChampSupValeur('NUMCDE',        '0');
  TOBL.AddChampSupValeur('NUMBESOIN',     '0');
  TOBL.AddchampSupValeur('DATEBESOIN', idate1900);
  TOBL.AddchampSupValeur('MORELINEBC',    '-');

  TOBL.AddChampSupValeur('REFINTERNE',    '');
  TOBL.AddChampSupValeur('REFEXTERNE',    '');
  TOBL.AddChampSupValeur('NATPIECEPREC',  '');
  TOBL.AddChampSupValeur('PIECEGENEREE',  '');
  TOBL.AddChampSupValeur('PIECEPRECEDENTE', '');
  TOBL.AddChampSupValeur('FACTURABLE',      'N');
                 
end;

Procedure TOF_INTEGREFICTPI.Controlechamp(Champ, Valeur : String);
begin

end;

procedure TOF_INTEGREFICTPI.SetScreenEvents;
begin

  FichierImport.OnElipsisClick := ImportElipsis;

  TypeMVT.OnChange      := TypeMVTOnchange;
  TypeAffaire.OnChange  := TypeAffaireOnChange;
  DateMvt.OnExit        := DateMVTOnExit;
  Analyse.OnExit        := AnalyseOnExit;
  Depot.OnExit          := DepotOnExit;
  Qte.OnExit            := QteOnExit;
  //
  ChkIntegre.OnClick    := IntegreOnClick;
  ChkConsoFact.OnClick  := ConsoFactOnClick;
  //
  Affaire1.Onexit       := CodeOnExit;
  Affaire2.OnExit       := CodeOnExit;
  Affaire3.OnExit       := CodeOnExit;
  CodeTiers.Onexit      := CodeOnExit;
  CodeArticle.Onexit    := CodeOnExit;
  //
  //NumBL.OnExit          := NumeroOnExit;
  NumCde.OnExit         := NumeroOnExit;
  NumBesoin.OnExit      := NumeroOnExit;
  //
  CAB_Affaire.Onexit    := codeBarreOnExit;
  CAB_Tiers.OnExit      := codeBarreOnExit;
  CAB_Article.OnExit    := codeBarreOnExit;

  BtImport.OnClick      := BtImportOnClick;

  BtSelect1.OnClick     := BtSelectOnClick;
  BtEfface1.OnClick     := BtEffaceOnclick;
  BtSelectCAB1.OnClick  := BtSelectCABOnClick;
  //
  BtSelect2.OnClick     := BtSelectOnClick;
  BtEfface2.OnClick     := BtEffaceOnclick;
  BtSelectCAB2.OnClick  := BtSelectCABOnClick;
  //
  BtSelect3.OnClick     := BtSelectOnClick;
  BtEfface3.OnClick     := BtEffaceOnclick;
  BtSelectCAB3.OnClick  := BtSelectCABOnClick;
  //
  BtSelectNUMCDE.OnClick:= BTSelectNumCdeOnClick;
  BtSelectBESOIN.OnClick:= BTSelectBesoinOnClick;
  BtSelectLigne.OnClick := BTSelectLigneOnClick;
  //
  BGeneration.OnClick   := BTGenerationOnClick;
  //
  BParamEtat.OnClick    := ParamEtat;
  BParamEtat1.OnClick   := ParamEtat1;
  //
  refInterne.OnExit     := RefOnExit;
  refExterne.OnExit     := RefOnExit;
  //
  if assigned(ChkApercu)    then ChkApercu.OnClick    := OnClickApercu;
  if assigned(ChkReduire)   then ChkReduire.OnClick   := OnClickReduire;

  if assigned(FETAT)        then FETAT.OnChange       := OnChangeEtat;
  if assigned(FETAT1)       then FETAT1.OnChange      := OnChangeEtat1;

end;

procedure TOF_INTEGREFICTPI.SetGrilleEvents (Etat : boolean);
begin

  if Etat then
  begin
    Grille.OnRowEnter   := GrilleRowEnter;
    Grille.GetCellCanvas:= GetCellCanvas;
    Grille1.OnRowEnter  := Grille1RowEnter;
    Grille2.OnDblClick  := InventaireDblClick;
  end
  else
  begin
    Grille.OnRowEnter   := Nil;
    Grille.GetCellCanvas:= nil;
    Grille1.OnRowEnter  := Nil;
    Grille2.OnDblClick  := Nil;
  end;

end;

procedure TOF_INTEGREFICTPI.DestroyTOB;
begin

  FreeAndNil(TOBGrille);
  FreeAndNil(TobAnomalie);
  FreeAndNil(TobInventaire);

  FreeAndNil(TOBPIECE);

  FreeAndNil(Tobrejete);
  FreeandNil(TobValide);

  FreeAndNil(TobStock);

end;

procedure TOF_INTEGREFICTPI.DefinieGrid;
begin

  // Définition de la liste de saisie pour la grille Détail
  fColNamesGrille   := 'SEL;TYPEMVT;DATEMVT;DEPOT;CODEFOURNISSEUR;CODEAFFAIRE;CODEARTICLE;QTE;INTEGRE';
  Falignement       := 'C.0  ---;G.0  ---;C.0  ---;G.0  ---;G.0  ---;G.0  ---;G.0  ---;D/2  -X-;C.0  ---';
  Ftitre            := ' ;Type Mvt;Date Mvt;Dépot;Fournisseur;Chantier;Article;Qté;A intégrer';
  fLargeur          := '2;25;10;15;17;17;30;8;2';
  //
  // Définition de la liste de saisie pour la grille anomalie
  fColNamesGrille1  := 'SEL;TYPEERREUR;DEPOT;CODEFOURNISSEUR;CODEAFFAIRE;CODEARTICLE;QTE;INTEGRE';
  FalignementGrille1:= 'C.0  ---;G.0  ---;G.0  ---;G.0  ---;G.0  ---;G.0  ---;D/2  -X-;C.0  ---';
  FtitreGrille1     := ' ;Type Erreur;Dépot;Fournisseur;Chantier;Article;Qté;A intégrer';
  fLargeurGrille1   := '2;50;15;17;17;30;8;5';
  //
  // Définition de la liste de saisie pour la grille Inventaire
  fColNamesGrille2  := 'SEL;GIE_CODELISTE;GIE_LIBELLE;GIE_DEPOT;GIE_DATEINVENTAIRE;GIE_TYPEINVENTAIRE';
  FalignementGrille2:= 'C.0  ---;C.0  ---;G.0  ---;G.0  ---;G.0  ---;C.0  ---';
  FtitreGrille2     := ' ;Code Inv;Désignation;Dépot;Date Inv.;Type Inv.';
  fLargeurGrille2   := '2;15;30;15;15;10';
  //
  // Définition de la liste de saisie pour la grille Inventaire
  fColNamesGrille3  := 'SEL;GL_CODEARTICLE;GL_LIBELLE;GL_QTEFACT;GL_QTERESTE;GL_PUHTDEV';
  FalignementGrille2:= 'C.0  ---;G.0  ---;G.0  ---;D/2  -X-;D/2  -X-;D/2  -X-';
  FtitreGrille3     := ' ;Code Art;Désignation;Qté Fact.;Qté Reste.;PU HT';
  fLargeurGrille3   := '2;30;100;20;20;20';

end;


procedure TOF_INTEGREFICTPI.ImportElipsis(Sender : TObject);
var TT  : TOpenDialog;
    rep : string;
begin

  Rep := FichierImport.text;

  //ouverture du sélecteur de fichier windows dans le répertoire des modèles...
	TT := TOpenDialog.Create(Ecran);
  TRY
    TT.DefaultExt := '.txt';
    TT.Filter := 'Modèle Texte (*.txt)|*.txt';
    if (rep = '') or (rep = 'C:\') then Rep := GetParamSocSecur('SO_REPRECUPFIC', '', False);
    //Rep := ExtractFilePath(rep);
    TT.InitialDir := rep;
    if TT.Execute then
    begin
      FichierImport.text := TT.FileName;
    end;
  FINALLY
  	TT.Free;
  end;

  if FichierImport.Text <> '' Then BtImport.Visible := True;

end;

procedure TOF_INTEGREFICTPI.BtImportOnClick(Sender: TObject);
begin

  //Chargement de la Tob en fonction des élément préchargés.
  ChargeTOB;

  //Vérification si des lignes ont été rejetées et affichage de l'onglet anomalie
  //AffichageGrilleAnomalie;

  //On se positionne sur le bon onglet
  Pages.ActivePage := PGeneralite;

  //chargement de la grille avec les éléments de la TOB
  //GestionGS.TOBG  := TOBGrille;
  //GestionGS.GS    := Grille;
  //
  //ChargementGrille(GestionGS);

  //Chargement de la première ligne
  if TOBGrille.detail.count <> 0 then
  begin
    //on Se positionne sur la première ligne de la grille
    Grille.row := 1;
    GrilleRowEnter(self, Grille.Row, Cancel, true);
    Pages.enabled := True;
  end;

  //Création de la TOBPiece pour génération ultérieure des Documents
  CreateTOBPiece;

  FichierImport.Enabled := False;
  BtImport.Visible      := False;

end;

procedure TOF_INTEGREFICTPI.AffichageGrilleAnomalie;
var GestionG1 : TGestionGS;
    NoLig     : Integer;
    TOBLAna   : TOB;
begin

  CaptionErreur.Caption := '';

  NoLig := TobGrille.detail[Grille.row-1].GetValue('NOLIG');

  if Tobanomalie.Detail.count <> 0 then
  begin

    //Affichage de l'onglet anomalie
    PAnomalies.TabVisible   := True;
    //chargement de la grille avec les éléments de la TOB
    GestionG1.fColNamesGS   := fColNamesGrille1;
    GestionG1.FalignementGS := FalignementGrille1;
    GestionG1.FtitreGS      := FtitreGrille1;
    GestionG1.fLargeurGS    := fLargeurGrille1;
    GestionG1.TOBG          := Tobanomalie;
    GestionG1.GS            := Grille1;
    //
    DessineGrille(GestionG1);
    //
    GestionG1.GS.row := 1;
    //GestionG1.TOBG.PutGridDetail(GestionG1.GS,False,True,GestionG1.fColNamesGS);
    ChargementGrille(GestionG1);
    //
    Grille.InvalidateRow(Grille.Row);
    //
    Toblana := Tobanomalie.findfirst(['NOLIG'],[NoLig],false);
    if Toblana <> nil then CaptionErreur.Caption := TOBlAna.GetValue('TYPEERREUR');
  end
  else
  begin
    //suppression de l'onglet anomalie
    PAnomalies.TabVisible := False;
    Tobanomalie.ClearDetail;
    Grille.InvalidateRow(Grille.Row);
  end;

end;

procedure TOF_INTEGREFICTPI.BtSelectOnClick(Sender: TObject);
begin

  if Pages.ActivePage = PAffaire          then
  begin
    RechercheAffaire;
    Grille.Cells [SG_AFFAIRE,Grille.Row] := TOBGrille.Detail[Grille.row-1].GetString('CODEAFFAIRE');
  end
  else if Pages.ActivePage = PFournisseur then
  begin
    RechercheFournisseur;
    Grille.Cells [SG_FOURNISSEUR,Grille.Row] := TOBGrille.Detail[Grille.row-1].GetString('CODEFOURNISSEUR');
  end
  else if Pages.ActivePage = PArticle     then
  Begin
    RechercheArticle;
    Grille.Cells [SG_ARTICLE,Grille.Row] := TOBGrille.Detail[Grille.row-1].GetString('CODEARTICLE');
  end;

end;

procedure TOF_INTEGREFICTPI.BtEffaceOnclick(Sender: TObject);
begin

  if Pages.ActivePage = PAffaire          then
  begin
    InitZoneAffaire;
    InitZoneAffaireToTob(TOBGrille.Detail[Grille.Row-1]);
    //Remise à blanc de la zone dans la TOBGrille
    if TOBGRille.Detail.count <> 0 then Grille.Cells[SG_AFFAIRE,Grille.Row] := '';
    TypeAffaire.SetFocus;
  end
  else if Pages.ActivePage = PFournisseur then
  begin
    InitzoneFournisseur;
    InitZoneFournisseurToTob(TOBGrille.Detail[Grille.Row-1]);
    //Remise à blanc de la zone dans la TOBGrille
    if TOBGRille.Detail.count <> 0 then Grille.Cells[SG_FOURNISSEUR,Grille.Row] := '';
    CodeTiers.SetFocus;
  end
  else if Pages.ActivePage = PArticle     then
  Begin
    SoustractionQteInv(CodeArticle.Text, StrToFloat(Qte.Text),TOBGRille.Detail[Grille.Row-1].GetValue('NOLIG'));
    InitZoneArticle;
    InitZoneArticleToTob(TOBGrille.Detail[Grille.Row-1]);
    //Remise à blanc de la zone dans la TOBGrille
    if TOBGRille.Detail.count <> 0 then Grille.Cells[SG_ARTICLE,Grille.Row] := '';
    CodeArticle.SetFocus;
  end;

end;

procedure TOF_INTEGREFICTPI.BtSelectCABOnClick(Sender: TObject);
begin

  if Pages.ActivePage = PAffaire          then
  Begin
    RechercheCABAffaire;
    Grille.Cells[SG_AFFAIRE,Grille.Row] := TOBGrille.Detail[Grille.Row-1].GetString('CODEAFFAIRE');
  end
  else if Pages.ActivePage = PFournisseur then
  begin
    RechercheCABFournisseur;
    Grille.Cells[SG_FOURNISSEUR,Grille.Row] := TOBGrille.Detail[Grille.Row-1].GetString('CODEFOURNISSEUR');
  end
  else if Pages.ActivePage = PArticle     then
  begin
    RechercheCABArticle;
    Grille.Cells[SG_ARTICLE,Grille.Row] := TOBGrille.Detail[Grille.Row-1].GetString('CODEARTICLE');
  end;

end;

Procedure TOF_INTEGREFICTPI.BTSelectNumCdeOnClick(Sender : TObject);
Var StRange     : string;
    Retour      : string;
    OldTiers    : string;
begin

  StRange :=  'GP_NATUREPIECEG=CF;GP_VENTEACHAT=ACH';

  //StRange := StRange + ';AFFAIRE0= ' + TypeAffaire.Value;

  OldTiers  := Affaire.Text;

  If NumCde.Text <> ''   then StRange := StRange + ';GP_NUMERO='   + NumCde.Text + ';GP_NUMERO_='+ NumCde.Text;

  //If (Affaire.Text <> '') And (Affaire.Text <> 'A') then StRange := stRange + ';GP_AFFAIRE='  + Affaire.Text;
  If CodeTiers.Text <> ''   then StRange := stRange + ';GP_TIERS='    + CodeTiers.Text;

  //If RefInterne.Text <> ''  then StRange := StRange + ';GP_REFINTERNE=' + RefInterne.Text;
  //If RefExterne.Text <> ''  then StRange := StRange + ';GP_REFEXTERNE=' + RefExterne.Text;

  Retour := AGLLanceFiche('BTP', 'BTPIECENEG_MUL',stRange ,'','RECUP;CONSULTATION;STATUT=AFF');

  NumCde.Text     := ReadtokenSt(Retour);
  Affaire.Text    := ReadtokenSt(Retour);
  CodeTiers.Text  := ReadtokenSt(Retour);

  TOBGrille.detail[Grille.Row - 1].PutValue('NUMCDE', NumCDE.Text);

  ChargeBL(TOBGrille.Detail[Grille.Row - 1], NumCde.text, 'CF');

end;

Procedure TOF_INTEGREFICTPI.BTSelectBesoinOnClick(Sender : TObject);
Var StRange     : string;
    Retour      : string;
    OldAffaire  : string;
begin

  StRange :=  'GP_NATUREPIECEG=CBT;GP_VENTEACHAT=VEN';
  StRange := StRange + ';AFFAIRE0= ' + TypeAffaire.Value;

  OldAffaire  := Affaire.Text;

  If NumBesoin.Text <> ''   then StRange := StRange + ';GP_NUMERO='   + NumBesoin.Text + ';GP_NUMERO_='+ NumBesoin.Text;

  If (Affaire.Text <> '') And (Affaire.Text <> 'A') then StRange := stRange + ';GP_AFFAIRE='  + Affaire.Text;
  If (Affaire0.Text <> '')  then StRange := stRange + ';GP_AFFAIRE0='  + Affaire0.Text;
  If (Affaire1.Text <> '')  then StRange := stRange + ';GP_AFFAIRE1='  + Affaire1.Text;
  If (Affaire2.Text <> '')  then StRange := stRange + ';GP_AFFAIRE2='  + Affaire2.Text;
  If (Affaire3.Text <> '')  then StRange := stRange + ';GP_AFFAIRE3='  + Affaire3.Text;

  If CodeTiers.Text <> ''   then StRange := stRange + ';GP_TIERS='    + CodeTiers.Text;

  If RefInterne.Text <> ''  then StRange := StRange + ';GP_REFINTERNE=' + RefInterne.Text;
  If RefExterne.Text <> ''  then StRange := StRange + ';GP_REFEXTERNE=' + RefExterne.Text;

  Retour := AGLLanceFiche('BTP', 'BTPIECENEG_MUL',stRange ,'','RECUP;CONSULTATION;STATUT=AFF');

  NumBesoin.Text    := ReadtokenSt(Retour);
  if NumBesoin.text <> '' then
  begin
    Affaire.Text    := ReadtokenSt(Retour);
    CodeTiers.Text  := ReadtokenSt(Retour);
  end;

  TOBGrille.detail[Grille.Row - 1].PutValue('NUMBESOIN', NumBesoin.Text);

  //ChargeTobAffaire(TOBGrille.detail[Grille.Row - 1],  Affaire.Text);

  ControleCBT(TOBGrille.Detail[Grille.Row - 1]);

  //If NumBesoin.Text <> '' then
  //begin
  //  Question := Question + 'à cette affaire ?';
  //  GestionModif('CODEAFFAIRE',Affaire.text,OldAffaire,SG_AFFAIRE);
  //end;

end;

procedure TOf_INTEGREFICTPI.TypeMVTOnchange(Sender: TObject);
Var TOBL      : TOB;
    ZoneSt    : string;
begin

  if TypeMVT.Value= '' then exit;

  TOBL := TobGrille.Detail[Grille.row-1];

  if TypeMvt.Value=TOBL.GetString('TYPEMVT') then exit;

  //ON MODIFIE LA LIGNE DE LA GRILLE AVEC LA NOUVELLE VALEUR DU TYPE mvt !!!
  TOBL.PutValue('TYPEMVT', TypeMVT.Value);

  Grille.Cells [SG_TYPEMVT,Grille.Row] := RechDom('BTTYPEMVT',TOBL.GetString('TYPEMVT'),False);

  //on réinitialise toutes les valeur d'anomalie à part le dépôt
  TOBL.PutValue('ANOMALIE',    '');
  TOBL.PutValue('ANOMALIEAFF', '');
  TOBL.PutValue('ANOMALIEBC',  '');
  TOBL.PutValue('ANOMALIEFRS', '');
  TOBL.PutValue('ANOMALIECLI', '');
  TOBL.PutValue('ANOMALIEMAR', '');

  //On charge l'onglet Affaire en fonction du type de Mouvement
  if (TypeMVT.Value = 'EA') then
  begin
    ZoneSt         := TOBL.GetString('CodeAFFAIRE');
    ChargeTobAffaire(TOBL, ZoneSt);
  end;

  //On charge l'onglet Fournisseur en fonction du type de Mouvement
  if (TypeMVT.Value = 'CD') then
  begin
    ZoneSt         := TOBL.GetString('CodeAFFAIRE');
    ChargeTobAffaire(TOBL, ZoneSt);
    ZoneSt         := TOBL.GetString('CodeFOURNISSEUR');
    ChargeTOBTiers(TOBL, ZoneSt);
  end;

  if TypeMVT.Value = 'BL' then
  begin
    NumCde.text    := TOBL.GetString('NUMCDE');
    ChargeBL(TOBL, NumCde.text, 'CF');
    If NumCde.text = '' then
    begin
      ZoneSt         := TOBL.GetString('CodeFOURNISSEUR');
      ChargeTOBTiers(TOBL, ZoneSt);
    end;
  end
  else if TypeMVT.Value = 'SA' then
  begin
    NumBesoin.text := TOBL.GetString('NUMBESOIN');
    if NumBesoin.text = '' then
    begin
      ZoneSt         := TOBL.GetString('CodeAFFAIRE');
      ChargeTobAffaire(TOBL, ZoneSt);
    end
    else
      ControleCBT(TOBL);
  end;

  GestionAnomalie(TOBGrille.Detail[Grille.row-1]);

  AffichageGrilleAnomalie;

End;

procedure TOf_INTEGREFICTPI.TypeAffaireOnchange(Sender: TObject);
var TOBL : TOB;
begin

  if TypeAffaire.Value= '' then exit;

  if  TobGrille.Detail.Count > 0 then TOBL := TobGrille.Detail[Grille.row-1] else exit;

  if TypeAffaire.Value=Affaire0.text then exit;

  //le type d'affaire est différent
  InitZoneAffaire;

  InitZoneAffaireToTob(TOBL);

  Affaire0.text := TypeAffaire.Value;

end;

procedure  TOF_INTEGREFICTPI.DateMVTOnExit(Sender : Tobject);
Var Indice  : Integer;
    SaveLig : Integer;
begin

  SaveLig := Grille.row;

  if DateMvt.Text <> TobGrille.detail[Grille.row-1].GetString('DATEMVT') then
  begin
    if PgiAsk('Voulez-vous appliquer cette modification à l''ensemble des lignes ?','Modification Date')=MrYes then
    begin
      For Indice := 0 to TOBGrille.detail.count -1 do
      begin
        TobGrille.Detail[Indice].PutValue('DATEMVT', DateMVT.Text);
        Grille.Cells [SG_DATEMVT,Indice+1] := TOBGrille.Detail[Indice].GetString('DATEMVT');
      end;
    end
    else
    begin
      TobGrille.Detail[Grille.row-1].PutValue('DATEMVT', DateMVT.Text);
      Grille.Cells [SG_DATEMVT,Grille.Row] := TOBGrille.Detail[Grille.row-1].GetString('DATEMVT');
    end;
  end;

  Grille.row := SaveLig;

end;

procedure  TOF_INTEGREFICTPI.AnalyseOnExit(Sender : Tobject);
Var Indice : Integer;
    SaveLig : Integer;
begin

  SaveLig := Grille.row;

  if Analyse.Text ='' then exit;

  IF Analyse.Text <> TobGrille.detail[Grille.row-1].GetString('ANALYSE') then
  begin
    if PgiAsk('Voulez-vous appliquer cette modification à l''ensemble des lignes ?','Modification Date')=MrYes then
    begin
      For Indice := 0 to TOBGrille.detail.count -1 do
      begin
        TobGrille.Detail[Indice].PutValue('ANALYSE', Analyse.Text);
      end;
    end
    else
    begin
      TobGrille.Detail[Grille.row-1].PutValue('ANALYSE', Analyse.Text);
    end;
  end;

  Grille.row := SaveLig;

end;

procedure TOf_INTEGREFICTPI.DepotOnExit(Sender: TObject);
Var Indice : Integer;
    SaveLig: Integer;
begin

  SaveLig := Grille.row;

  if Depot.text = '' then exit;

  if ChargeTobDepot(TobGrille.Detail[Grille.row-1], Depot.text) then
  begin
    if TobGrille.detail.count = 1 then
    begin
      TobGrille.Detail[Grille.row-1].PutValue('DEPOT', Depot.Text);
      Grille.Cells [SG_DEPOT,Grille.Row] := TOBGrille.Detail[Grille.row-1].GetString('DEPOT');
      TobGrille.Detail[Grille.row-1].PutValue('ANOMALIE','');
      GestionAnomalie(TOBGrille.Detail[Grille.row-1]);
    end
    else
    begin
      IF Depot.Text <> TobGrille.detail[Grille.row-1].GetString('DEPOT') then
      begin
        if PgiAsk('Voulez-vous appliquer cette modification à l''ensemble des lignes ?','Modification Date')=MrYes then
        begin
          For Indice := 0 to TOBGrille.detail.count -1 do
          begin
            Grille.Row := Indice+1;
            TobGrille.Detail[Indice].PutValue('DEPOT', Depot.Text);
            Grille.Cells [SG_DEPOT,Indice+1] := TOBGrille.Detail[Indice].GetString('DEPOT');
            TobGrille.Detail[Indice].PutValue('ANOMALIE','');
            GestionAnomalie(TOBGrille.Detail[Grille.row-1]);
          end;
        end
        else
        begin
          TobGrille.Detail[Grille.row-1].PutValue('DEPOT', Depot.Text);
          Grille.Cells [SG_DEPOT,Grille.Row] := TOBGrille.Detail[Grille.row-1].GetString('DEPOT');
          TobGrille.Detail[Grille.row-1].PutValue('ANOMALIE','');
          GestionAnomalie(TOBGrille.Detail[Grille.row-1]);
        end;
      end;
    end;
  end
  else
  begin
    TobGrille.Detail[Grille.row-1].PutValue('ANOMALIE', 'DEPOT');
    TobGrille.Detail[Grille.row-1].PutValue('ERREUR',   'X');
    GestionAnomalie(TOBGrille.Detail[Grille.row-1]);
  end;

  //DessineGrille(GestionGS);
  //chargementGrille(GestionGS);

  Grille.row := SaveLig;

  AffichageGrilleAnomalie;

end;

procedure TOf_INTEGREFICTPI.QteOnExit(Sender: TObject);
Var QteArt      : Double;
begin

  QteArt := TobGrille.Detail[Grille.row-1].GetDouble('QTE');

  SoustractionQteInv(CodeArticle.text, QteArt,TobGrille.Detail[Grille.row-1].GetValue('NOLIG'));

  //ON MODIFIE LA LIGNE DE LA GRILLE AVEC LA NOUVELLE VALEUR !!
  TobGrille.Detail[Grille.row-1].PutValue('QTE', Qte.text);
  Grille.Cells [SG_QTE,Grille.Row] := TOBGrille.Detail[Grille.row-1].GetString('QTE');

  AdditionQteInventaire(CodeArticle.text, StrToFloat(Qte.Text), TobGrille.Detail[Grille.row-1].GetValue('NOLIG'));

end;

procedure TOf_INTEGREFICTPI.IntegreOnClick(Sender : Tobject);
Var ToblA : TOB;
begin

  if ChkIntegre.checked = True then
  begin
    TobGrille.Detail[Grille.row-1].PutValue('INTEGRE', 'X');
    GestionAnomalie(TobGrille.Detail[Grille.row-1]);
  end
  else
  begin
    TobGrille.Detail[Grille.row-1].PutValue('INTEGRE', '-');
    TobGrille.Detail[Grille.row-1].PutValue('ERREUR',  '-');
    //On annule l'anomalie si on l'intègre pas de façon à ne pas avoir de message à la fin...
    ToblA := TobAnomalie.FindFirst(['NOLIG'], [TOBGrille.Detail[Grille.row- 1].GetValue('NOLIG')], False);
    If ToblA <> nil then FreeAndNil(ToblA);
  end;

  Grille.Cells [SG_INTEGRE,Grille.Row] := TOBGrille.Detail[Grille.row-1].GetString('INTEGRE');

  AffichageGrilleAnomalie;

end;

//FV1 : 23/08/2017 - FS#2648 - DELABOUDINIERE - Rendre les consos associées à un appel facturable
procedure TOf_INTEGREFICTPI.ConsoFactOnClick(Sender : Tobject);
begin

  if TOBGrille = nil then exit;

  if ChkConsoFact.checked then
    TobGrille.Detail[Grille.row-1].PutValue('FACTURABLE', 'A')
  else
    TobGrille.Detail[Grille.row-1].PutValue('FACTURABLE', 'N');

end;


Procedure TOF_INTEGREFICTPI.RefOnExit(Sender : Tobject);
Var TOBL      : TOB;
    ARow      : Integer;
begin

  if TOBGrille = nil then exit;

  ARow := Grille.Row;

  if ((ARow <= 0) or (ARow > TOBGrille.Detail.Count)) then Exit;

  TOBL := TOBGrille.Detail[ARow - 1];

  if TOBL = nil then Exit;

  TOBL.PutValue('REFEXTERNE', RefExterne.text);
  TOBL.PutValue('REFINTERNE', RefInterne.text);

end;

Procedure TOF_INTEGREFICTPI.NumeroOnExit(Sender : Tobject);
var TOBL      : TOB;
    OldNumCde : string;
    OldNumBL  : string;
    OldBesoin : string;
    ARow      : Integer;
begin

  if TOBGrille = nil then exit;

  ARow := Grille.Row;

  if ((ARow <= 0) or (ARow > TOBGrille.Detail.Count)) then Exit;

  TOBL := TOBGrille.Detail[ARow - 1];

  if TOBL = nil then Exit;

  OldNumCde := TOBL.GetString('NUMCDE');
  OldNumBL  := TOBL.GetString('NUMBL');
  OldBesoin := TOBL.GetString('NUMBESOIN');

  if TypeMVT.Value = 'BL' then
  begin
    TOBL.PutValue('NUMCDE', NumCde.Text);
    ChargeBL(TOBL, NumCde.text, 'CF');
    If NumCde.text = '' then
    begin
      ChargeTOBTiers(TOBL, TOBL.GetString('CodeFOURNISSEUR'));
    end;
  end
  else if TypeMVT.Value = 'SA' then
  begin
    TOBL.PutValue('NUMBESOIN', NumBesoin.Text);
    ControleCBT(TOBL);
    if NumBesoin.text = '' then
    begin
      ChargeTobAffaire(TOBL, TOBL.GetString('CodeAFFAIRE'));
      if TOBL.GetString('ANOMALIEAFF')='' then ChargeInfoAffaire(TOBL);
    end;
  end;

  GestionAnomalie(TOBGrille.Detail[Grille.row-1]);

  AffichageGrilleAnomalie;

end;

Procedure TOF_INTEGREFICTPI.CodeOnExit(Sender : Tobject);
Var IP      : Integer;
    TOBL    : Tob;
    ARow    : Integer;
    OldCode : String;
begin

  ARow := Grille.Row;

  if ((ARow <= 0) or (ARow > TOBGrille.Detail.Count)) then Exit;

  TOBL := TOBGrille.Detail[ARow - 1];

  if TOBL = nil then Exit;

  if Pages.ActivePage = PAffaire          then
  begin
    OldCode       := TOBL.GetString('CODEAFFAIRE');
    Affaire.text  := DechargeCleAffaire(Affaire0, Affaire1, Affaire2, Affaire3, Avenant, Client.text, Taconsult, False, True, false, IP);
    ChargeTobAffaire(TOBL, Affaire.text);
    Question := Question + 'à cet Affaire ?';
    GestionModif('CODEAFFAIRE',Affaire.Text,OldCode,SG_AFFAIRE);
  end
  else if Pages.ActivePage = PFournisseur then
  begin
    if oldcode <> CodeTiers.text then
    begin
      OldCode := TOBL.GetString('CODEFOURNISSEUR');
      ChargeTobTiers(TOBL, Codetiers.text);
      Question := Question + 'à ce Fournisseur ?';
      GestionModif('CODEFOURNISSEUR',Codetiers.Text,OldCode,SG_FOURNISSEUR);
    end;
  end
  else if Pages.ActivePage = PArticle     then
  begin
    OldCode := TOBL.GetString('CODEARTICLE');
    ChargeTobArticle(TOBL, OldCode);
    Question := Question + 'à cet Article ?';
    GestionModif('CODEARTICLE',CodeArticle.Text,OldCode,SG_ARTICLE);
  end;

end;

Procedure TOF_INTEGREFICTPI.CodeBarreOnExit(Sender : TObject);
Var IP      : Integer;
    TOBL    : Tob;
    OldCode : string;
begin

  TOBL := TobGrille.Detail[Grille.row-1];

  if Pages.ActivePage = PAffaire          then
  begin
    OldCode       := TOBL.GetString('CODEAFFAIRE');
    Affaire.text  := DechargeCleAffaire(Affaire0, Affaire1, Affaire2, Affaire3, Avenant, Client.text, Taconsult, False, True, false, IP);
    ChargeTobAffaire(TOBL, Affaire.text);
    Question := Question + 'à cet Affaire ?';
    GestionModif('CODEAFFAIRE',Affaire.Text,OldCode,SG_AFFAIRE);
  end
    //GestionAffaire(CAB_Affaire.text,TOBL.GetString('CAB_AFFAIRE'), TOBL)
  else if Pages.ActivePage = PFournisseur then
  begin
    OldCode := TOBL.GetString('CODEFOURNISSEUR');
    if TOBL.GetValue('TYPEMVT') = 'BL' then
    begin
      if OldCode <> CodeTiers.text then
      begin
        PGIBox('Modification impossible ce fournisseur appartient à une réception', 'Modification Impossible');
        Codetiers.text := OldCode;
      end;
    end
    else
    begin
      ChargeTobTiers(TOBL, Codetiers.text);
      Question := Question + 'à ce Fournisseur ?';
      GestionModif('CODEFOURNISSEUR',Codetiers.Text,OldCode,SG_FOURNISSEUR);
    end;
  end
    //GestionFournisseur(CAB_Tiers.text,TOBL.GetString('CAB_Tiers'), TOBL)
  else if Pages.ActivePage = PArticle     then
  begin
    OldCode := TOBL.GetString('CODEARTICLE');
    ChargeTobArticle(TOBL, CodeArticle.text);
    Question := Question + 'à cet Article ?';
    GestionModif('CODEARTICLE',CodeArticle.Text,OldCode,SG_ARTICLE);
  end;

end;

Procedure TOF_INTEGREFICTPI.GestionModif(Zone, NewCode, OldCode : string; IndLigne : Integer);
Var TobL        : TOB;
    TOBC        : TOB;
begin

  Tobl    := TobGrille.Detail[Grille.row-1];
  //
  Question := 'Voulez-vous appliquer cette modification à l''ensemble des lignes affectées ';

  //Si le nouveau code = l'ancien aucune modifs
  if NewCode <> OldCode then
  begin
    //si l'ancien code était à blanc pas de modif en série
    if (OldCode = '') Or (OldCode = 'A') then
      ModificationSimple(zone, Newcode, TOBL, IndLigne)
    else
    begin
      //s'il n'y a pas au moins deux ligne avec le même code pas de modification en série
      TOBC := TOBGrille.FindFirst([zone],[OldCode],False);
      if TOBC <> Nil then
      begin
        TOBC := TOBGrille.FindNext([zone],[OldCode],False);
        if TOBC = Nil then
          ModificationSimple(zone, Newcode, TOBL, IndLigne)
        else
        begin
          if PgiAsk(Question,'Modification en Série')=MrNo then
            ModificationSimple(zone, Newcode, TOBL,IndLigne)
          else
            ModificationenSerie(zone,OldCode, Newcode, TOBL, IndLigne);
        end;
      end
      else
      begin
        ModificationSimple(zone, Newcode, TOBL, IndLigne)
      end;
    end;
  end
  else
  begin
    GestionAnomalie(TOBGrille.Detail[Grille.Row-1]);
  end;

  AffichageGrilleAnomalie;

end;

Procedure TOF_INTEGREFICTPI.ModificationSimple(zone, Newcode : String; TOBL : TOB;IndLigne : Integer);
begin

  //pas de gestion en série...
  Grille.Cells [IndLigne,Grille.Row] := NewCode;

  if Zone = 'CODEARTICLE' then
  begin
    //chargement des infos article avec la nouvelle saisie
    ChargeTOBArticle(TOBL, NewCode);
    if TOBL.GetString('ANOMALIEMAR')='' then ChargeInfoArticle(TOBL);
  end
  else if Zone = 'CODEFOURNISSEUR' then
  begin
    //chargement des infos Tiers avec la nouvelle saisie
    ChargeTOBTiers(TOBL, NewCode);
    if TOBL.GetString('ANOMALIEFRS')='' then ChargeInfoTiers(TOBL);
  end
  else if Zone = 'CODEAFFAIRE' then
  begin
    //chargement des infos Affaire avec la nouvelle saisie
    ChargeTOBAffaire(TOBL, NewCode);
    if TOBL.GetString('ANOMALIEAFF')='' then ChargeInfoAffaire(TOBL);
  end
  else if Zone = 'NUMERO' then
  begin
    if TypeMVT.Value = 'BL' then
    Begin
      ChargeBL(TOBL, NewCode, 'CF');
      If NumCde.text = '' then
      begin
        ChargeTOBTiers(TOBL, TOBL.GetString('CodeFOURNISSEUR'));
        if TOBL.GetString('ANOMALIEFRS')='' then ChargeInfoTiers(TOBL);
      end;
    end
    else if TypeMVT.Value = 'SA' then
    begin
      ControleCBT(TOBL);
      if NumBesoin.text = '' then
      begin
        ChargeTobAffaire(TOBL, TOBL.GetString('CodeAFFAIRE'));
        if TOBL.GetString('ANOMALIEAFF')='' then ChargeInfoAffaire(TOBL);
      end;
      ChargeInfoTiers(TOBL);
    end;
  end;

  GestionAnomalie(TOBGrille.Detail[Grille.Row-1]);

  Grille.Cells [IndLigne, Grille.Row] := NewCode;

end;

Procedure TOF_INTEGREFICTPI.ModificationEnSerie(Zone, OldCode, Newcode : String; TOBL : TOB;IndLigne : Integer);
Var TobLInvArt  : TOB;
    Indice      : Integer;
    SaveLig     : Integer;
begin

  SaveLig := Grille.row;

  //Suppression de la ligne d'inventaire pour l'ancien article
  if Zone = 'CODEARTICLE' then
  begin
    TobLInvArt  := TobInventaire.FindFirst(['CODEARTICLE'],[OldCode],True);
    if TOBLinvArt <> nil then FreeAndNil(TOBLInvArt);
  end;

  //
  //On boucle sur la tobgrille pour mettre à jour les codes
  For Indice := 0 to TOBGrille.Detail.Count -1 do
  begin
    if TobGrille.Detail[Indice].GetString(Zone) = OldCode then
    begin
      Grille.Row := Indice + 1;
      TobGrille.Detail[Indice].PutVALUE(Zone, NewCode);
      Grille.Cells [IndLigne, Grille.Row] := NewCode;
      if Zone = 'CODEAFFAIRE' then
        ChargeTOBAffaire(TobGrille.Detail[Grille.row-1], NewCode)
      else if Zone = 'CODEFOURNISSEUR' then        //chargement des infos fournisseurs avec la nouvelle saisie
        ChargeTobTiers(TobGrille.Detail[Grille.row-1], NewCode)
      else if Zone = 'CODEARTICLE' then           //chargement des infos article avec la nouvelle saisie
        ChargeTOBArticle(TobGrille.Detail[Grille.row-1], NewCode);
      GestionAnomalie(TOBGrille.Detail[Grille.row-1]);
    end;
  end;

  ChargeGrilleToScreen(Grille.Row, TobGrille.Detail[Grille.Row-1]);

  Grille.row := SaveLig;

end;

procedure TOF_INTEGREFICTPI.GrilleRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
Var NoLig  : String;
    TOBLAna: Tob;
begin

  if Grille.RowCount = 0 then Exit;

  NoLig := TOBGrille.detail[ou-1].GetValue('NOLIG');

  ChargeGrilleToScreen(Ou, TOBGrille.detail[ou-1]);

  CaptionErreur.Caption := '';

  //positionnement sur la grille Anomalie pour garder la cohérence d'affichage
  TOBLAna := Tobanomalie.Findfirst(['NOLIG'],[noLig], false);
  if tobLAna <> nil then
  begin
    Grille1.Row := TOBLAna.GetIndex + 1;
    CaptionErreur.Caption := TOBLAna.GetValue('TYPEERREUR');
  end;

  if TobGrille.detail[Ou-1].GetValue('MORELINEBC') = 'X' then
  begin
    BtSelectLigne.Visible := true;
    if TobGrille.detail[Ou-1].GetValue('ANOMALIEBC') = 'X' then BtSelectLigneOnClick(self);
  end
  else
    BtSelectLigne.Visible := False;

end;

procedure TOF_INTEGREFICTPI.Grille1RowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
Var TypeErr : string;
begin

  if Grille1.RowCount = 0 then Exit;

  //positionnement sur la grille toutes
  //il faut dans un premier temps trouver la ligne dans la tob...
  if TOBGrille.Detail.count > 0 then
  begin
    Grille.Row := StrToInt(TOBanomalie.detail[ou-1].GetValue('NOLIG'));
    GrilleRowEnter(self, Grille.Row, Cancel, true);
  end
  else
  begin
    if TobRejete.Detail.Count > 0 then
    begin
      TypeErr := Tobrejete.detail[Ou-1].GetString('TYPEERREUR');
      if TypeErr = 'GEN' then
      begin
        PGeneralite.TabVisible  := True;
        PArticle.TabVisible     := False;
        PFournisseur.TabVisible := False;
        PAffaire.TabVisible     := False;
        PEdition.TabVisible     := False;
        PInventaire.TabVisible  := False;
      end
      else if TypeErr = 'ART' then
      begin
        PGeneralite.TabVisible  := False;
        PArticle.TabVisible     := True;
        PFournisseur.TabVisible := False;
        PAffaire.TabVisible     := False;
        PEdition.TabVisible     := False;
        PInventaire.TabVisible  := False;
      end
      else if TypeErr = 'AFF' then
      begin
        PGeneralite.TabVisible  := False;
        PArticle.TabVisible     := False;
        PFournisseur.TabVisible := False;
        PAffaire.TabVisible     := True;
        PEdition.TabVisible     := False;
        PInventaire.TabVisible  := False;
      end
      else if TypeErr = 'FRS' then
      begin
        PGeneralite.TabVisible  := False;
        PArticle.TabVisible     := False;
        PFournisseur.TabVisible := True;
        PAffaire.TabVisible     := False;
        PEdition.TabVisible     := False;
        PInventaire.TabVisible  := False;
      end
      else if TypeErr = 'INV' then
      begin
        GestionGrilleInventaire;
      end
    end;
  end;

end;

Procedure TOF_INTEGREFICTPI.GestionGrilleInventaire;
begin

  PGeneralite.TabVisible  := False;
  PArticle.TabVisible     := False;
  PFournisseur.TabVisible := False;
  PAffaire.TabVisible     := False;
  PEdition.TabVisible     := False;
  //
  AffichageGrilleInventaire;
  //
  PInventaire.TabVisible  := True;

end;

Procedure TOF_INTEGREFICTPI.InventaireDblClick(Sender: TObject);
Var TOBLigInvent : TOB;
    TOBL         : TOB;
begin

  With GestionGS do
  begin
    TOBLigInvent := TOBG.detail[GS.Row +1];
    TOBL         := TobGrille.Detail[Grille.row+1];
    //
    MAJLigneInventaire(TOBLigInvent, TOBL);
  end;

end;

procedure TOF_INTEGREFICTPI.ChargeTOB;
Var FSource  : TextFile;
    StLig    : string;
    TOBL     : TOB;
    RImport  : string;
    ZoneSt   : String;
    NoLig    : Integer;
    Indice   : Integer;
    TypeMvt  : String;
begin

  RImport   := FichierImport.text;

  NoLig     := 0;

  AssignFile(FSource,RImport);   //Associe la variable F au fichier texte
  try
    Reset(FSource); //Ouvre le fichier sans le modifier

    While Not eof(FSource) do
    Begin
        TOBL := TOB.Create('Les Lignes',TOBGrille,-1);
        LibErreur := '';
        Readln(FSource, StLig); //Lit une ligne du fichier text jusqu'au prochain saut de ligne
        //on initialise les zones de la TOB
        InitZoneTOB(TOBL);
        //On charge la tob avec les valeurs du fichier text
        ChargementLigneTOB(TOBL,StLig);
    end;

  finally
    CloseFile(FSource); //Ferme l'association entre la variable F et le fichier texte
  end;

  TriTableGrille;

  for Indice := 0 To TobGrille.Detail.count -1 do
  begin
    TobL := TobGrille.detail[Indice];

    TOBL.PutValue('ERREUR', '-');

    //Oncharge le Numéro de ligne
    inc(NoLig);
    TOBL.PutValue('NOLIG', NoLig);
    //
    TypeMvt := TOBL.GetString('TYPEMVT');
    //
    //On Contrôle l'existence du dépot
    ZoneSt        := TOBL.GetString('DEPOT');
    if not ChargeTobDepot(TOBL, ZoneSt) then
    begin
      TOBL.PutValue('ANOMALIE', 'DEPOT');
      TOBL.PutValue('ERREUR',   'X');
    end;

    //On charge l'onglet Affaire en fonction du type de Mouvement
    if (TypeMVT = 'EA') then
    begin
      ZoneSt      := TOBL.GetString('CODEAFFAIRE');
      ChargeTobAffaire(TOBL, ZoneSt);
    end
    else if TypeMVT = 'SA' then
    begin
      NumBesoin.text := TOBL.GetString('NUMBESOIN');
      if NumBesoin.text = '' then
      begin
        ZoneSt         := TOBL.GetString('CodeAFFAIRE');
        ChargeTobAffaire(TOBL, ZoneSt);
      end
      else
        ControleCBT(TOBL);
    end;

    //On charge l'onglet Fournisseur en fonction du type de Mouvement
    if (TypeMVT = 'CD') then
    begin
      ZoneSt      := TOBL.GetString('CODEAFFAIRE');
      ChargeTobAffaire(TOBL, ZoneSt);
      ZoneSt      := TOBL.GetString('CodeFOURNISSEUR');
      ChargeTOBTiers(TOBL, ZoneSt);
    end;

    if (TypeMVT = 'BL') then
    Begin
      ZoneSt      := TOBL.GetString('NUMCDE');
      ChargeBL(TOBL, ZoneSt, 'CF');
      If NumCde.text = '' then
      begin
        ChargeTOBTiers(TOBL, TOBL.GetString('CodeFOURNISSEUR'));
        if TOBL.GetString('ANOMALIEFRS')='' then ChargeInfoTiers(TOBL);
      end;
    End;

    //On charge l'onglet Article
    ZoneSt        := TOBL.GetString('CodeARTICLE');
    ChargeTobArticle(TOBL, ZoneSt);
    //
    GestionAnomalie(TOBL);
  end;

  Grille.row := 1;
  TOBGrille.PutGridDetail(Grille,False,True,GestionGS.fColNamesGS);

end;

Function TOF_INTEGREFICTPI.GenerationMsgAnomalie(TOBL : TOB) : String;
Var MotifAnomalie : string;
begin

  Result := '';

  TOBL.PutValue('ERREUR', '-');
  //
 if (TOBL.GetValue('ANOMALIE')    <> '') OR
    (TOBL.GetValue('ANOMALIEAFF') <> '') OR
    (TOBL.GetValue('ANOMALIEFRS') <> '') OR
    (TOBL.GetValue('ANOMALIEMAR') <> '') OR
    (TOBL.GetValue('ANOMALIEBC')  <> '') OR
    (TOBL.GetValue('ANOMALIECLI') <> '') then TOBL.PutValue('ERREUR', 'X');

  if TOBL.getString('ERREUR') = '-' Then Exit;

  IF TOBL.GetString('ANOMALIE')='STOCK' then
  begin
    if Result = '' then Result := 'Erreur Intégration :';
    MotifAnomalie := 'Stock Physique Article ' + TOBL.GetString('CODEARTICLE') + ' négatif après intégration';
    Result := Result + ' ' + MotifAnomalie;
  end;

  //contrôle si des anomalies et chargement du libellé d'erreur
  IF TOBL.GetString('ANOMALIE')='DEPOT' then
  begin
    if Result = '' then Result := 'Erreur sur la ligne :';
    MotifAnomalie := 'Dépôt Inexistant';
    Result := Result + ' ' + MotifAnomalie;
  end;

  IF (TOBL.GetString('TYPEERREUR') = 'GEN') OR
     (TOBL.GetString('TYPEERREUR') = 'ART') OR
     (TOBL.GetString('TYPEERREUR') = 'AFF') OR
     (TOBL.GetString('TYPEERREUR') = 'FRS') OR
     (TOBL.GetString('TYPEERREUR') = 'INV') then
  begin
    if Result = '' then Result := 'Erreur sur la ligne :';
    MotifAnomalie := TOBL.GetString('MOTIF');
    Result := Result + ' ' + MotifAnomalie;
  end;

  IF TOBL.GetString('ANOMALIEAFF')='X' then
  begin
    if Result = '' then Result := 'Erreur sur la ligne :';
    MotifAnomalie := TOBL.GetString('LIBANOMALIEAFF');
    if motifAnomalie = '' then
      Result := Result + 'Anomalie Affaire'
    else
      Result := Result + ' ' + MotifAnomalie;
  end;

  IF TOBL.GetString('ANOMALIEBC')='X' then
  begin
    if Result = '' then Result := 'Erreur sur la ligne :';
    MotifAnomalie := TOBL.GetString('LIBANOMALIEBC');
    if motifAnomalie = '' then
      Result := Result + 'Anomalie Livraison sur Besoin de chantier '
    else
      Result := Result + ' ' + MotifAnomalie;
  end;


  IF TOBL.GetString('ANOMALIECLI')='X' then
  begin
    if Result = '' then Result := 'Erreur sur la ligne :';
    MotifAnomalie := TOBL.GetString('LIBANOMALIECLI');
    if motifAnomalie = '' then
      Result := Result + 'Anomalie Client'
    else
      Result := Result + ' ' + MotifAnomalie;
  end;

  IF TOBL.GetString('ANOMALIEFRS')='X' then
  begin
    if Result = '' then Result := 'Erreur sur la ligne :';
    MotifAnomalie := TOBL.GetString('LIBANOMALIEFRS');
    if motifAnomalie = '' then
      Result := Result + 'Anomalie Fournisseur'
    else
      Result := Result + ' ' + MotifAnomalie;
  end;

  IF (TOBL.GetString('ANOMALIEMAR')='X') OR
     (TOBL.GetString('ANOMALIEMAR')='S') OR
     (TOBL.GetString('ANOMALIEMAR')='F') then
  begin
    if Result = '' then Result := 'Erreur sur la ligne :';
    MotifAnomalie := TOBL.GetString('LIBANOMALIEMAR');
    if motifAnomalie = '' then
      Result := Result + 'Anomalie Marchandise'
    else
      Result := Result + ' ' + MotifAnomalie;
  end;

end;

Function TOF_INTEGREFICTPI.ChargeTOBDepot(TOBL : TOB; CodeDepot : string) : Boolean;
Var CodeArticle : String;
    Qteinv      : Double;
    Qte         : double;
    ToblInvArt  : TOB;
begin

  result := False;

  StSQL := 'SELECT GDE_LIBELLE FROM DEPOTS WHERE GDE_DEPOT="' + CodeDepot + '"';
  QQ := OpenSQL(STSQL,False);

  //Chargement des zones écran....
  if Not QQ.Eof then
  begin
    TOBL.PutVALUE('LIB_DEPOT',  QQ.Findfield('GDE_LIBELLE').AsString);
    Lib_Depot.Caption := QQ.Findfield('GDE_LIBELLE').AsString;
    //
    //Recherche de la ligne Article/Dépôt dans l'inventaire pour Mise à jour de la quantité de la ligne
    CodeArticle := TOBL.GetString('CODEARTICLE');
    TobLInvArt  := TobInventaire.FindFirst(['CODEARTICLE','CODEDEPOT'],[CodeArticle,CodeDepot],True);
    //
    if TOBLinvArt <> nil then
    begin
      QteInv  := TOBLinvArt.GetDouble('QTE');
      QTE     := TOBL.GetDouble('QTE');
      QteInv  := QteInv + Qte;
      TOBLInvArt.PutValue('QTE', QteInv);
    end;
    result := True;
  end;

  Ferme(QQ);

end;


Procedure TOF_INTEGREFICTPI.ChargementLigneTOB(TOBL : TOB;StLig : String);
Var StDate  : string;
    StQte   : String;
    DateMvt : TDateTime;
    CodeArt : String;
    Codeana : string;
    CodeAff : String;
    CodeFrs : String;
    CodeDep : String;
    NumeroCD: String;
    NumeroBL: String;
    NumeroBC: String;
    Qte     : Double;
    TypeMvt : string;
begin

  //On découpe l'enreg et on positionne les zones dans la TOBGrille

  if StLig = '' then Exit;

  TypeMVT := copy(StLig,1,2);

  StDate  := Copy(StLig,9,2) + '/' + copy(StLig,7,2) + '/' + copy(StLig,3,4);
  DateMvt := StrToDate(StDate);
  if DateMvt = iDate1900 then DateMvt := now;
  //
  CodeArt := Trim(copy(StLig,11,20));
  //
  StQte   := Trim(FindEtReplace(Copy(StLig,31,12),'.',',',True));
  if IsNumeric(StQte, True) Then Qte := StrToFloat(StQte) else Qte := 0;

  if TypeMvt = 'BL' then NumeroBL:= Trim(Copy(StLig,43,20));

  if (TypeMvt = 'SA') OR (TypeMvt = 'EA') OR (TypeMvt = 'CD') then
  begin
    CodeAff := Trim(Copy(StLig,43,20));
    if CodeAff = '' then CodeAff := TypeAffaire.Value;
  end;

  //Modification pour client LSEBAT DELATTRE (Code ana à 15 et pas 10)
  //FV1 : 14/12/20018 - FS#2764 - SFD ETS SERVAIS - En gestion TPI erreur dépôt inexistant et les livraisons ne seraient pas créés
  if Length(StLig) <= 113 then
  begin
    CodeAna := Trim(Copy(StLig,63,10));
    CodeDep := Trim(Copy(StLig,73,20));

    if TypeMvt = 'BL' then  NumeroCD:= Trim(Copy(StLig,93,20));

    if TypeMvt = 'SA' then  NumeroBC:= Trim(Copy(StLig,93,20));

    If TypeMvt = 'CD' then CodeFrs := Trim(Copy(StLig,93,20));
  end
  else
  begin
    CodeAna := Trim(Copy(StLig,63,15));
    CodeDep := Trim(Copy(StLig,78,20));

    if TypeMvt = 'BL' then  NumeroCD:= Trim(Copy(StLig,98,20));

    if TypeMvt = 'SA' then  NumeroBC:= Trim(Copy(StLig,98,20));

    If TypeMvt = 'CD' then CodeFrs := Trim(Copy(StLig,98,20));
  end;
  //
  TOBL.PutValue('TYPEMVT',      TypeMvt);
  TOBL.PutValue('DATEMVT',      DateTimeToStr(DateMvt));
  TOBL.PutValue('CODEARTICLE',  CodeArt);
  TOBL.PutValue('QTE',          Qte);
  TOBL.PutValue('CODEAFFAIRE',  CodeAff);
  TOBL.PutValue('ANALYSE',      CodeAna);
  TOBL.PutValue('DEPOT',        CodeDep);
  TOBL.PutValue('CODEFOURNISSEUR', CodeFrs);
  //
  TOBL.PutValue('NUMBL',         NumeroBL);
  TOBL.PutValue('NUMCDE',        NumeroCD);
  TOBL.PutValue('NUMBESOIN',     NumeroBC);
  //
  NumBL.Text      := NumeroBL;
  NumCde.Text     := NumeroCD;
  NumBesoin.Text  := NumeroBC;
  //
end;

procedure TOF_INTEGREFICTPI.ChargeGrilleToScreen(Arow : Integer; TOBL : TOB);
Var ValQte      : Double;
begin

  InitZoneEcran;
  //
  PAffaire.TabVisible      := False;
  PFournisseur.TabVisible  := False;
  PArticle.TabVisible      := False;
  //
  RefInterne.Visible       := false;
  RefExterne.Visible       := false;
  Lib_RefInterne.Visible   := false;
  Lib_RefExterne.Visible   := false;
  //
  TNumCde.Visible          := False;
  NumCde.Visible           := False;
  TNumBL.Visible           := False;
  NumBL.Visible            := False;
  BTSelectNumCde.Visible   := False;
  TNumBesoin.Visible       := False;
  NumBesoin.Visible        := False;
  BTSelectBesoin.Visible   := False;
  //
  LIB_TenueStock.visible  := False;
  Art_TenueStock.Checked  := False;
  //
  Art_Ferme.Checked       := False;
  Frs_Ferme.Checked       := False;
  Aff_Ferme.Checked       := False;
  //
  CodeTiers.Enabled        := True;
  CAB_Tiers.Enabled        := True;

  //On charge l'onglet Généralité
  TypeMVT.Value := TOBL.GetString('TYPEMVT');
  DateMvt.Text  := TOBL.GetString('DATEMVT');
  Analyse.Text  := TOBL.GetString('ANALYSE');
  Depot.Text    := TOBL.GetString('DEPOT');
  if not ChargeTobDepot(TOBL, Depot.Text) then
  begin
    TOBL.PutValue('ANOMALIE', 'DEPOT');
    TOBL.PutValue('ERREUR',   'X');
  end;

  if TOBL.GetString('INTEGRE') = 'X' then
    ChkIntegre.Checked := True
  else
  begin
    if TOBL.Getstring('ANOMALIE') <> 'STOCK' then ChkIntegre.Checked := False;
  end;

  Refinterne.text := TOBL.GetValue('REFINTERNE');
  RefExterne.text := TOBL.GetValue('REFEXTERNE');

  //FV1 : 20/09/2016 - FS#2132 - DELABOUDINIERE : Gestion TPI, diverses corrections

  ValQte        := StrToFloat(TOBL.GetString('QTE'));
  //if ValQte     <> 0 then Qte.enabled := False else Qte.enabled := true;
  Qte.enabled   := true;
  Qte.Text      := FloatToStr(ValQte);

  //On charge l'onglet Affaire
  if Pos(TypeMvt.Value, 'BT;EA;SA;CD') > 0 then
  begin
    ChargeInfoAffaire(TOBL);
  end;

  if (TypeMVT.Value = 'BL') then
  begin
    ChargeInfoAffaire(TOBL);
  end;

  //On charge l'onglet fournisseur
  if Pos(TypeMvt.Value, 'BL;CD') > 0 then ChargeInfoTiers(TOBL);

  if Pos(TypeMvt.Value, 'SA;CD;BL;EA') > 0 then
  begin
    RefInterne.Visible       := True;
    RefExterne.Visible       := True;
    Lib_RefInterne.Visible   := True;
    Lib_RefExterne.Visible   := True;
  end;

  //chargement à partir de TOBL le reste est fait dans charge tob...
  Lib_Depot.caption := TOBL.GetString('LIB_DEPOT');

  //On charge l'onglet Article
  ChargeInfoArticle(TOBL);

  Pages.ActivePage := PGeneralite;

end;

procedure TOF_INTEGREFICTPI.ChargeTobWithScreen;
Var TOBL : TOB;
begin

  TOBL := tobGrille.detail[Grille.row -1];

  TOBL.PutValue('TYPEMVT',    TypeMVT.Value);
  TOBL.PutValue('DATEMVT',    DateMvt.Text);
  TOBL.PutValue('ANALYSE',    Analyse.Text);
  TOBL.PutValue('DEPOT',      Depot.Text);
  TOBL.PutValue('LIB_DEPOT',  Lib_Depot.caption);

  TOBL.PutValue('REFINTERNE', Refinterne.text);
  TOBL.PutValue('REFEXTERNE', RefExterne.text);

  //Onglet Affaire
  TOBL.PutValue('CLIENT',         Client.Text);
  TOBL.PutValue('LIB_CLIENT',     Lib_Tiers.caption);
  TOBL.PutValue('CODEAFFAIRE',    Affaire.Text);
  TOBL.PutValue('NUMBESOIN',      NUMBesoin.text);
  TOBL.PutValue('LIB_AFFAIRE',    Lib_Affaire.Caption);
  TOBL.PutValue('CAB_Affaire',    CAB_Affaire.Text);
  TOBL.PutValue('Aff_QualifCAB',  Aff_QualifCAB.Value);
  TOBL.PutValue('Lib_Responsable',Lib_Responsable.Caption);
  TOBL.PutValue('DOMAINE',        Domaine);
  TOBL.PutValue('ETABLISSEMENT',  Etablissement);
  TOBL.PutValue('LIB_DOMAINE',    Lib_Domaine.Caption);

  //Onglet Fournisseur
  TOBL.PutValue('CODEFOURNISSEUR',CodeTiers.Text);
  TOBL.PutValue('NUMCDE',         NUMCDE.text);
  TOBL.PutValue('NUMBL',          NUMBL.text);
  TOBL.PutValue('CAB_Tiers',      CAB_Tiers.Text);
  TOBL.PutValue('Frs_QualifCAB',  Frs_QualifCAB.Value);
  TOBL.PutValue('Juridique',      Juridique.Caption);
  TOBL.PutValue('Nom',            Nom.caption);
  TOBL.PutValue('ADR1',           ADR1.Caption);
  TOBL.PutValue('ADR2',           ADR2.Caption);
  TOBL.PutValue('ADR3',           ADR3.Caption);
  TOBL.PutValue('ADR4',           ADR4.Caption);
  TOBL.PutValue('CP',             CP.Caption);
  TOBL.PutValue('Ville',          Ville.Caption);

  //Onglet Article
  TOBL.PutValue('CODEARTICLE',    CodeArticle.text);
  TOBL.PutValue('GA_ARTICLE',     Article.text);
  //
  TOBL.PutValue('Lib_Article',    Lib_Article.Caption);
  //
  TOBL.PutValue('CAB_Article',    CAB_Article.text);

  TOBL.PutValue('Art_QualifCAB',  Art_QualifCAB.Value);
  //
  TOBL.PutValue('QUALIFUNITESTO', QualifUS.Value);
  TOBL.PutValue('QUALIFUNITEACH', QualifUA.Value);
  TOBL.PutValue('QUALIFUNITEVTE', QualifUV.Value);
  //
  TOBL.PutValue('CoefUSUV',       CoefUSUV.Text);
  TOBL.PutValue('CoefUAUS',       CoefUAUS.text);
  TOBL.PutValue('PAHT',           PAHT.Text);
  TOBL.PutValue('PAUA',           PAUA.text);
  TOBL.PutValue('DPA',            DPA.text);
  TOBL.PutValue('PMAP',           PMAP.text);

  TOBL.PutValue('QTE',            Qte.Text);

  if TypeAffaire.Value = 'W' then
  Begin
    if ChkConsoFact.Checked then
      TOBL.PutValue('FACTURABLE',   'A')
    else
      TOBL.PutValue('FACTURABLE',   'N')
  end
  else
    TOBL.PutValue('FACTURABLE', 'N');

end;


procedure TOF_INTEGREFICTPI.ChargeTobAffaire(TOBL : TOB; CodeBarre : string);
Var TypeMvt : string;
begin

  TOBL.PutValue('ANOMALIEAFF', '');
  TOBL.PutValue('LIBANOMALIEAFF','');

  TypeMvt       := TOBL.GetString('TYPEMVT');

  if (TypeMVT = 'SA') Or (TypeMVT = 'EA') then
  begin
    if (codeBarre = '') Or (CodeBarre = 'A')  then
    Begin
      //On gère une anomalie car le code Affaire est obligatoire...
      TOBL.PutValue('ANOMALIEAFF', 'X');
      TOBL.PutValue('ERREUR', 'X');
      if TypeMVT = 'SA' then
        TOBL.PutValue('LIBANOMALIEAFF', TypeAffaire.Text + ' obligatoire sur Livraison sur Chantier ')
      else if TypeMVT = 'EA' then
        TOBL.PutValue('LIBANOMALIEAFF', TypeAffaire.Text + ' obligatoire sur Retour Chantier ');
      Exit;
    end;
  end;

  if (CodeBarre = '') Or (CodeBarre = 'A') then exit;

  //Affaire0.Text := TypeAffaire.Value;
  if (copy(CodeBarre, 1, 1) <> 'A') AND (copy(CodeBarre, 1, 1) <> 'I') AND (copy(CodeBarre, 1, 1) <> 'W') then
  begin
    //Codebarre := 'A' + CodeBarre;
    //Contrôle unicité du code affaire hors détermination du type d'affaire (A, I ou W)
    StSQL := 'SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE LIKE "%' + CodeBarre + '%" ';
    QQ := OpenSQL(StSQL, False);
    if QQ.Eof then
    begin
      //On gère une anomalie car le code Affaire est obligatoire...
      TOBL.PutValue('CODEAFFAIRE', CodeBarre);
      TOBL.PutValue('ANOMALIEAFF', 'X');
      TOBL.PutValue('ERREUR', 'X');
      TOBL.PutValue('LIBANOMALIEAFF','Chantier Inexistant');
      Exit;
    end
    else
    begin
      if QQ.RecordCount > 1 then
      begin
        TOBL.PutValue('CODEAFFAIRE', CodeBarre); //'A' +
        TOBL.PutValue('ANOMALIEAFF', 'X');
        TOBL.PutValue('ERREUR', 'X');
        TOBL.PutValue('LIBANOMALIEAFF','Code affaire non unique');
        Exit;
      end
      else
      begin
        CodeBarre := QQ.FindField('AFF_AFFAIRE').AsString;
      end;
    end;
    Ferme(QQ);
  end;     

  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect1, TaModif, CodeBarre, True);

  if (Affaire1.Text = '') AND (Affaire2.Text = '') AND (Affaire3.Text = '') then
  begin
    if (TypeMVT = 'SA') Or (TypeMVT = 'EA') then
    begin
      //On gère une anomalie car le code Affaire est obligatoire...
      TOBL.PutValue('ANOMALIEAFF', 'X');
      TOBL.PutValue('ERREUR', 'X');
      if TypeMVT = 'SA' then
        TOBL.PutValue('LIBANOMALIEAFF', TypeAffaire.Text + ' obligatoire sur Livraison sur Chantier ')
      else if TypeMVT = 'EA' then
        TOBL.PutValue('LIBANOMALIEAFF', TypeAffaire.Text + ' obligatoire sur Retour Chantier ');
    end;
    TOBL.PutValue('CODEAFFAIRE', copy(CodeBarre, 1, 1));
    Exit;
  end;

  //FV1 - 16/12/2015 : FS#1833 - GESTION TPI : Modification dans la récupération du code chantier.
  TOBAffaire := TOB.Create('CAB_AFFAIRE',nil,-1);

  StSQL := 'SELECT *  FROM AFFAIRE ';
  StSQL := StSQL + ' WHERE AFF_AFFAIRE LIKE "%' + CodeBarre + '%" ';

  QQ    := OpenSQL(StSQL, False);

  //Chargement des zones écran....
  if Not QQ.Eof then
  begin
    TOBAffaire.SelectDB('UNE AFFAIRE',QQ);
    //FV1 : 23/08/2017 -  FS#2648 - DELABOUDINIERE - Rendre les consos associées à un appel facturable
    if TypeMvt = 'SA' then
    begin
      if Affaire0.Text = 'W' then
      begin
        If GetParamSocSecur('SO_CONSOFACT', False) = True then
          TOBL.PutValue('FACTURABLE', 'A')
        else
          TOBL.PutValue('FACTURABLE', 'N')
      end
      else
      begin
        TOBL.PutValue('FACTURABLE', 'N')
      end;
    end;
    //
    if (TypeMVT = 'CD') And (OkAccepte) then
    begin
      if ((Affaire0.Text = 'A') Or (Affaire0.Text = 'I')) And (TOBAffaire.GetValue('AFF_ETATAFFAIRE')<> 'ACP') then
      begin
        TOBL.PutValue('ANOMALIEAFF', 'X');
        TOBL.PutValue('ERREUR', 'X');
        if Affaire0.Text = 'A' then
          TOBL.PutValue('LIBANOMALIEAFF','Chantier Non Accepté')
        else
          TOBL.PutValue('LIBANOMALIEAFF','Contrat Non Accepté');
        ChargeTOBGRILLEwithTOBAFFAIRE(TOBL, TOBAFFAIRE);
      end
      else
        ChargeTOBGRILLEwithTOBAFFAIRE(TOBL, TOBAFFAIRE);
    end
    else
      ChargeTOBGRILLEwithTOBAFFAIRE(TOBL, TOBAFFAIRE);
  end
  else
  begin
    TOBL.PutValue('CODEAFFAIRE', CodeBarre);
    TOBL.PutValue('ANOMALIEAFF', 'X');
    TOBL.PutValue('ERREUR', 'X');
    //
    if Affaire0.Text = 'A' then
      TOBL.PutValue('LIBANOMALIEAFF','Chantier Inexistant')
    else if Affaire0.Text = 'I' then
      TOBL.PutValue('LIBANOMALIEAFF','Contrat Inexistant')
    else
      TOBL.PutValue('LIBANOMALIEAFF','Appel Inexistant');
    //
    //TOBL.PutValue('LIBANOMALIEAFF','Chantier Inexistant');
  end;

  Ferme(QQ);

  FreeAndNil(TobAffaire);

end;

procedure TOF_INTEGREFICTPI.ChargeTOBGRILLEwithTOBAFFAIRE(TOBL,TOBAFFAIRE : TOB);
Begin

  TOBL.PutVALUE('CLIENT',      TOBAffaire.GetString('AFF_TIERS'));

  TOBL.PutValue('ANOMALIECLI', '');
  TOBL.PutValue('LIBANOMALIECLI', '');

  StSQL := 'SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="CLI" AND T_TIERS="' + TOBAffaire.GetString('AFF_TIERS') + '"';
  QQ := Opensql(StSQL,False);
  if not QQ.eof then
    TOBL.PutVALUE('LIB_CLIENT', QQ.Findfield('T_LIBELLE').asString)
  else
  begin
    TOBL.PutValue('ANOMALIECLI', 'X');
    TOBL.PutValue('ERREUR', 'X');
    TOBL.PutValue('LIBANOMALIECLI','Client Inexistant');
    //TOBL.PutVALUE('CLIENT', '');
    TOBL.PutValue('LIB_CLIENT','');
  end;
  Ferme(QQ);

  TOBL.PutVALUE('CODEAFFAIRE', TOBAffaire.GetString('AFF_AFFAIRE'));
  TOBL.PutVALUE('AFF_AFFAIRE', TOBAffaire.GetString('AFF_AFFAIRE'));
  TOBL.PutVALUE('LIB_AFFAIRE', TOBAffaire.GetString('AFF_LIBELLE'));
  //
  IF TOBAffaire.FieldExists('BCB_CODEBARRE')        then TOBL.PutVALUE('CAB_Affaire',   TOBAffaire.GetString('BCB_CODEBARRE'));
  IF TOBAffaire.FieldExists('BCB_QUALIFCODEBARRE')  then TOBL.PutVALUE('AFF_QUALIFCAB', TOBAffaire.GetString('BCB_QUALIFCODEBARRE'));
  //
  TOBL.PutVALUE('Lib_Responsable'  , RechDom('AFLRESSOURCE',TOBAffaire.GetString('AFF_RESPONSABLE'),False));
  TOBL.PutVALUE('Domaine'          , TOBAffaire.GetString('AFF_DOMAINE'));
  TOBL.PutVALUE('Etablissement'    , TOBAffaire.GetString('AFF_ETABLISSEMENT'));
  TOBL.PutVALUE('Lib_Domaine'      , RechDom('AFLRESSOURCE',TOBAffaire.GetString('AFF_DOMAINE'), False));
  //
  TOBL.PutVALUE('AFF_FERME',   TOBAffaire.GetString('AFF_FERME'));
  //
end;

procedure TOF_INTEGREFICTPI.ChargeInfoAffaire(TOBL : TOB);
begin

  PAffaire.TabVisible := True;
  PAffaire.Enabled    := True;

  Client.Text         := TOBL.GetString('CLIENT');
  Lib_Tiers.caption   := TOBL.GetString('LIB_CLIENT');
  //
  Affaire.Text        := TOBL.GetString('CODEAFFAIRE');

  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect1, TaModif, Affaire.Text, True);

  if (Affaire0.text = '') Or (Pos(Affaire0.text, 'A;I;W')=0) then
  Begin
    Typeaffaire.Value := 'A';
    Affaire0.text     := Typeaffaire.Value;
    Affaire.Text      := Affaire.Text;
  end
  else
    TypeAffaire.value := Affaire0.text;

  NUMBesoin.Visible       := TOBL.GetString('TYPEMVT') = 'SA';
  TNumBesoin.Visible      := TOBL.GetString('TYPEMVT') = 'SA';
  BTSelectBesoin.Visible  := TOBL.GetString('TYPEMVT') = 'SA';

  CAB_Affaire.Visible     := True;
  Aff_QualifCAB.Visible   := True;

  if (Typeaffaire.Value = 'W') then
    PAffaire.Caption := 'Appel/Intervention'
  else if (Typeaffaire.Value = 'A') then
    PAffaire.Caption := 'Chantier'
  else  if (Typeaffaire.Value = 'I') then
    PAffaire.Caption := 'Contrat'
  Else
    PAffaire.Caption := 'Chantiers';

  //TOBL.PutValue('CODEAFFAIRE', Affaire.text);
  //FV1 : 23/08/2017 - FS#2648 - DELABOUDINIERE - Rendre les consos associées à un appel facturable
  ChkConsoFact.visible := False;
  //
  if TOBL.GetString('TYPEMVT') = 'SA' then
  begin
    NUMBesoin.text       := TOBL.GetString('NUMBESOIN');
  end;
  //
  Lib_Affaire.Caption := TOBL.GetString('LIB_AFFAIRE');
  //
  CAB_Affaire.Text    := TOBL.GetString('CAB_Affaire');
  Aff_QualifCAB.Value := TOBL.GetString('Aff_QualifCAB');
  if CAB_Affaire.text = '' then
  begin
    CAB_Affaire.Visible   := False;
    Aff_QualifCAB.Visible := False;
    BtSelectCAB1.Visible  := False;
    TGPAFFAIRE3.Visible   := False;
  end;
  //
  Lib_Responsable.Caption := TOBL.GetString('Lib_Responsable');
  Domaine                 := TOBL.GetString('DOMAINE');
  Etablissement           := TOBL.GetString('ETABLISSEMENT');
  Lib_Domaine.Caption     := TOBL.GetString('LIB_DOMAINE');
  //
  If TOBL.GetString('AFF_FERME')='X' then Aff_Ferme.Checked := True else Aff_Ferme.Checked := False;

end;

procedure TOF_INTEGREFICTPI.ChargeTobTiers(TOBL : TOB; CodeBarre : string);
begin

  TOBL.PutValue('ANOMALIEFRS', '');
  TOBL.PutValue('LIBANOMALIEFRS', '');

  TypeMvt.Value := TOBL.GetString('TYPEMVT');

  if TypeMVT.Value = 'CD' then
  begin
    if codeBarre = '' then
    Begin
      //On gère une anomalie car le code fournisseur est obligatoire...
      TOBL.PutValue('ANOMALIEFRS', 'X');
      TOBL.PutValue('ERREUR', 'X');
      TOBL.PutValue('LIBANOMALIEFRS', 'Fournisseur Obligatoire sur Commande fournisseur');
      Exit;
    end;
  end;
  
  TOBTiers   := TOB.Create('CAB_TIERS  ',nil,-1);

  StSQL := 'SELECT * FROM TIERS '+
           'LEFT JOIN BTIERS ON BT1_AUXILIAIRE=T_AUXILIAIRE '+
           'WHERE T_NATUREAUXI="FOU" AND T_TIERS="' + CodeBarre + '"';
  QQ    := OpenSQL(StSQL, False);

  //Chargement des zones écran....
  if Not QQ.Eof then
  begin
    TOBTiers.SelectDB('UN TIERS',QQ);
    ChargeTOBGRILLEwithTOBTIERS(TobL, TOBTiers);
  end
  else
  begin
    TOBL.PutValue('ANOMALIEFRS', 'X');
    TOBL.PutValue('ERREUR', 'X');
    TOBL.PutValue('LIBANOMALIEFRS', 'Fournisseur Inexistant');
  end;

  Ferme(QQ);

  FreeAndNil(TobTiers);

end;

procedure TOF_INTEGREFICTPI.ControleCBT(TOBL : TOB);
Var Cledoc    : R_CLEDOC;
    StD       : string;
    CodeAff   : string;
begin

  TOBL.PutValue('ANOMALIEBC',    '');
  TOBL.putValue('LIBANOMALIEBC', '');

  if TOBL.GetString('TYPEMVT') <> 'SA' then Exit;

  if TOBL.GetString('NUMBESOIN') = ''  then Exit;

  CodeAff := TOBL.GetString('CODEAFFAIRE');

  //Lecture de la pièce associé au besoin de chantier pour vérification si
  //plusieurs lignes avec le même code article dans un seul et même besoin chantier
  StSQL := 'SELECT GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, GL_INDICEG, GL_NUMORDRE ';
  StSQL := StSQL + ', GL_DATEPIECE, GL_AFFAIRE, GL_TIERS, GL_QTERESTE ';
  StSQL := StSQL + ' FROM LIGNE ';
  StSQL := StSQL + 'WHERE GL_NATUREPIECEG="CBT" ';

  if (IsNumeric(TOBL.GetValue('NUMBESOIN'))) then StSQL := StSQL + '  AND GL_NUMERO=' + TOBL.GetString('NUMBESOIN');
  
  StSQL := StSQL + '  AND GL_INDICEG=0';
  if TOBL.GetString('CODEARTICLE') <> '' then  StSQL := StSQL + ' AND GL_CODEARTICLE="' + TOBL.GetString('CODEARTICLE') + '"';
  if TOBL.GetString('CODEAFFAIRE') <> '' then  StSQL := StSQL + ' AND GL_AFFAIRE LIKE "' + TOBL.GetString('CODEAFFAIRE') + '%"';
  //
  QQ    := OpenSQL(StSQL, False);

  If Not QQ.Eof then
  begin
    Affaire.text        := QQ.Findfield('GL_AFFAIRE').asString;
    if QQ.RecordCount > 1 then
    begin
      TOBL.PutValue('ANOMALIEBC', 'X');
      TOBL.PutValue('ERREUR', 'X');
      TOBL.PutValue('LIBANOMALIEBC', 'Plusieurs lignes sur cet article pour ce besoin de chantier');
      TOBL.PutValue('MORELINEBC', 'X');
    end
    else
    begin
      if QQ.FindField('GL_QTERESTE').AsFloat <= 0 then
      begin
        //
        TOBL.PutValue('ERREUR', 'X');
        TOBL.PutValue('ANOMALIEBC', 'X');
        TOBL.PutValue('LIBANOMALIEBC', 'La ligne sur ce besoin de chantier est totalement soldée');
        TOBL.PutValue('MORELINEBC', '-');
        TOBL.PutValue('INTEGRE', '-');
      end
      else
      begin
        Cledoc.NaturePiece  := QQ.FindField('GL_NATUREPIECEG').AsString;
        Cledoc.DatePiece    := QQ.FindField('GL_DATEPIECE').AsDateTime;
        StD := FormatDateTime('ddmmyyyy', cledoc.DatePiece);
        Cledoc.Souche       := QQ.FindField('GL_SOUCHE').AsString;
        Cledoc.NumeroPiece  := QQ.FindField('GL_NUMERO').AsInteger;
        Cledoc.Indice       := QQ.FindField('GL_INDICEG').AsInteger;
        Cledoc.NumOrdre     := QQ.FindField('GL_NUMORDRE').AsInteger;
        //
        TOBL.PutValue('DATEBESOIN',       Cledoc.DatePiece);
        TOBL.PutValue('PIECEPRECEDENTE',  StD + ';' + cledoc.NaturePiece  + ';' + Cledoc.Souche  + ';' + IntToStr(cledoc.NumeroPiece) + ';' + IntToStr(cledoc.Indice) + ';' + IntToStr(cledoc.NumOrdre) + ';');
        TOBL.PutValue('NATPIECEPREC',     Cledoc.NaturePiece);
        TOBL.PutValue('MORELINEBC', '-');
        TOBL.PutValue('ANOMALIEBC', '');
        TOBL.PutValue('LIBANOMALIEBC', '');
      end;
      //
    end;
  end
  else
  Begin
    TOBL.PutValue('ERREUR', 'X');
    TOBL.PutValue('ANOMALIEBC', 'X');
    TOBL.PutValue('LIBANOMALIEBC', 'Besoin de chantier Inexistant');
    TOBL.PutValue('MORELINEBC', '-');
  end;

  ChargeTobAffaire(TOBL, Affaire.text);
  if TOBL.GetString('ANOMALIEAFF')='' then ChargeInfoAffaire(TOBL);

  //GestionGs.TOBG.PutGridDetail(GestionGs.GS,False,True,GestionGs.fColNamesGS);

  //Grille.OnRowEnter(Self, Grille.Row, Cancel, True);

  Ferme(QQ);

end;

procedure TOF_INTEGREFICTPI.ChargeBL(TOBL : TOB; CodeBarre, NaturePiece : String);
Var CodeTiers : String;
    Codeaff   : String;
begin

  if TOBL.GetString('TYPEMVT') <> 'BL' then Exit;

  if CodeBarre = ''  then exit;

  StSQL := 'SELECT GP_REFINTERNE, GP_REFEXTERNE, GP_TIERS, GP_AFFAIRE FROM PIECE WHERE GP_NATUREPIECEG="' + NaturePiece + '" ';
  StSQL := StSQL + '  AND GP_NUMERO=' + CodeBarre;
  StSQL := StSQL + '  AND GP_INDICEG=0';
  QQ    := OpenSQL(StSQL, False,1,'',True);

  //chargement du Numéro de Fournisseur
  if Not QQ.eof then
  begin
    CodeTiers  := QQ.Findfield('GP_TIERS').asString;
    CodeAff    := QQ.Findfield('GP_AFFAIRE').asString;
    TOBL.PutValue('REFINTERNE'  , QQ.Findfield('GP_REFINTERNE').asString);

    RefInterne.Text := QQ.Findfield('GP_REFINTERNE').asString;
    RefExterne.Text := QQ.Findfield('GP_REFEXTERNE').asString;

    TOBL.PutValue('NATPIECEPREC', NaturePiece);

    ChargeTobAffaire(TOBL, CodeAff);
    if TOBL.GetString('ANOMALIEAFF')='' then ChargeInfoAffaire(TOBL);

    if Affaire.text <> '' then
    begin
      PAffaire.Enabled      := False;
      BtSelect1.Visible     := false;
      BtEfface1.Visible     := False;
      BtSelectCAB1.Visible  := False;
    end;

    ChargeTobTiers(TOBL, CodeTiers);

    //GestionGs.TOBG.PutGridDetail(GestionGs.GS,False,True,GestionGs.fColNamesGS);

    //Grille.OnRowEnter(Self, Grille.Row, Cancel, True);
  end
  else
  begin
    NumCde.Text := '';
    TOBL.PutValue('ANOMALIEFRS', 'X');
    TOBL.PutValue('ERREUR', 'X');
    if NaturePiece = 'CF' then TOBL.PutValue('LIBANOMALIEFRS', 'Commande fournisseur Inexistante')
  end;

  Ferme(QQ);

end;

procedure TOF_INTEGREFICTPI.ChargeTOBGRILLEwithTOBTIERS(TOBL, TOBTiers : TOB);
begin

  TOBL.PutValue('T_TIERS', TOBTiers.GetString('T_TIERS'));
  TOBL.PutValue('CODEFOURNISSEUR', TOBTiers.GetString('T_TIERS'));

  IF TOBTiers.FieldExists('BCB_CODEBARRE') Then
    TOBL.PutVALUE('CAB_Tiers',TOBTiers.GetString('BCB_CODEBARRE'))
  else
    TOBL.PutVALUE('CAB_Tiers',TOBTiers.GetString('BT1_CODEBARRE'));
  //
  IF TOBTiers.FieldExists('BCB_QUALIFCODEBARRE') Then
    TOBL.PutVALUE('Frs_QualifCAB',TOBTiers.GetString('BCB_QUALIFCODEBARRE'))
  else
    TOBL.PutVALUE('Frs_QualifCAB',TOBTiers.GetString('BT1_QUALIFCODEBARRE'));

  TOBL.PutVALUE('Juridique', RechDom('TTFORMEJURIDIQUE',TOBTiers.GetString('T_JURIDIQUE'),False));

  TOBL.PutVALUE('Nom',  TOBTiers.GetString('T_LIBELLE'));
  TOBL.PutVALUE('ADR1', TOBTiers.GetString('T_PRENOM'));
  TOBL.PutVALUE('ADR2', TOBTiers.GetString('T_ADRESSE1'));
  TOBL.PutVALUE('ADR3', TOBTiers.GetString('T_ADRESSE2'));
  TOBL.PutVALUE('ADR4', TOBTiers.GetString('T_ADRESSE3'));
  TOBL.PutVALUE('CP',   TOBTiers.GetString('T_CODEPOSTAL'));
  TOBL.PutVALUE('Ville',TOBTiers.GetString('T_VILLE'));

  TOBL.PutVALUE('Frs_Ferme', TobTiers.GetString('FRS_FERME'));

end;

procedure TOF_INTEGREFICTPI.ChargeInfoTiers(TOBL : TOB);
begin

  PFournisseur.TabVisible  := True;
  PFournisseur.enabled     := True;

  CodeTiers.Text    := TOBL.GetString('CODEFOURNISSEUR');

  NUMCDE.Visible    := TOBL.GetString('TYPEMVT') = 'BL';
  NUMBL.Visible     := TOBL.GetString('TYPEMVT') = 'BL';
  TNumCde.Visible   := TOBL.GetString('TYPEMVT') = 'BL';
  TNumBL.Visible    := TOBL.GetString('TYPEMVT') = 'BL';

  BtSelectNUMCDE.Visible := TOBL.GetString('TYPEMVT') = 'BL';

  if TOBL.GetString('TYPEMVT') = 'BL' then
  begin
    NUMCDE.text       := TOBL.GetString('NUMCDE');
    NUMBL.text        := TOBL.GetString('NUMBL');
  end;

  CAB_Tiers.Text      := TOBL.GetString('CAB_Tiers');
  //
  Frs_QualifCAB.Value := TOBL.GetString('Frs_QualifCAB');

  Juridique.Caption   := TOBL.GetString('Juridique');

  Nom.caption         := TOBL.GetString('Nom');
  ADR1.Caption        := TOBL.GetString('ADR1');
  ADR2.Caption        := TOBL.GetString('ADR2');
  ADR3.Caption        := TOBL.GetString('ADR3');
  ADR4.Caption        := TOBL.GetString('ADR4');
  CP.Caption          := TOBL.GetString('CP');
  Ville.Caption       := TOBL.GetString('Ville');

  If TOBL.GetString('FRS_FERME')='X' then Frs_Ferme.Checked := True else Frs_Ferme.Checked := False;

end;

procedure TOF_INTEGREFICTPI.ChargeTobArticle(TOBL : TOB; CodeBarre : string);
begin

  TOBarticle   := TOB.Create('CAB_ARTICLE',nil,-1);

  TOBL.PutValue('ANOMALIEMAR', '');
  TOBL.PutValue('LIBANOMALIEMAR', '');

  //Lecture du fichier CodeBarre pour récupération du code affaire
  StSQL := 'SELECT * FROM BTCODEBARRE LEFT JOIN ARTICLE ON BCB_IDENTIFCAB=GA_ARTICLE ';
  StSQL := StSQL + ' WHERE BCB_NATURECAB= GA_TYPEARTICLE ';
  StSQL := StSQL + '   AND BCB_CODEBARRE="' + CodeBarre + '"';
  QQ    := OpenSQL(StSQL, False);

  if QQ.EOF then
  begin
    StSQL := 'SELECT * FROM ARTICLE WHERE GA_TYPEARTICLE="MAR" ';
    StSQL := StSQL + '  AND (GA_CODEARTICLE="' + CodeBarre + '" ';
    StSQL := StSQL + '   OR GA_ARTICLE="' + CodeBarre + '")';
    QQ    := OpenSQL(StSQL, False);
  end;

  //Chargement des zones écran....
  if Not QQ.Eof then
  begin
    TOBarticle.SelectDB('UN ARTICLE',QQ);
    ChargeTOBGRILLEwithTOBArticle(TobL, TOBArticle);
    RechercheInventaireArticle(TOBL);
    CodeArticle.text := TOBL.GetString('CODEARTICLE');
    Article.text     := TOBL.GetString('GA_ARTICLE');
  end
  else
  Begin
    TOBL.PutValue('ANOMALIEMAR', 'X');
    TOBL.PutValue('ERREUR', 'X');
    TOBL.PutValue('LIBANOMALIEMAR', 'Article Inexistant');
  end;
    
  Ferme(QQ);

  FreeAndNil(TOBarticle);

end;

Procedure TOF_INTEGREFICTPI.RechercheInventaireArticle(TOBL : TOB);
Var TOBLInvArt  : TOB;
    TOBLInv     : TOB;
    TOBLigInv   : TOB;
    CodeArticle : String;
    CodeDepot   : String;
begin

  CodeArticle := TOBL.GetString('CODEARTICLE');
  CodeDepot   := TOBL.GetString('DEPOT');

  //On vérifie que l'article n'existe pas déjà dans la TOBInventaire
  TobLInvArt  := TobInventaire.FindFirst(['CODEARTICLE'],[CodeArticle],True);
  if TOBLinvArt = nil then  //il n'existe pas
  Begin
    TOBLInvArt := TOB.Create('ARTINV', TobInventaire,-1);
    TOBLInvArt.AddChampSupValeur('CODEARTICLE',CodeArticle);
    TOBLInvArt.AddChampSupValeur('QTECALC', 0.00);
  end;

  //Il existe on vérifie s'il a déjà des lignes d'inventaire par dépôt
  TobLInv  := TobLInvArt.FindFirst(['GIL_CODEARTICLE','GIL_DEPOT'],[CodeArticle, Depot.text],True);
  if TOBLinv = nil then //Aucune ligne pour cet article/Dépot
  Begin
    //on lit la liste d'inventaire de l'Article chargé
    StSQL := 'SELECT *  FROM LISTEINVENT LEFT JOIN LISTEINVLIG ';
    StSQL := StSQL + '    ON GIE_CODELISTE=GIL_CODELISTE ';
    StSQL := STSQL + ' WHERE GIE_VALIDATION="-" ';
    StsQL := StSQL + '   AND GIE_UTILISATEUR="' + V_PGI.User + '"';
    StSQL := StSQL + '   AND GIL_CODEARTICLE="' + CodeArticle + '"';
    //
    QQ := OpenSQL(StsQl, False);
    //
    if not QQ.eof then
    begin
      TOBLinv := TOB.Create ('lignes inventaire', TOBLInvArt, -1);
      TOBLInv.SelectDB('Ligne Inventaire', QQ);
      TOBLInv.AddChampSupValeur('QTECALC', 0.00);
      TOBLInv.AddChampSupValeur('NOLIG', TOBL.GetValue('NOLIG'));
      AdditionQteInventaire(CodeArticle, TOBL.GetDouble('QTE'), TOBL.GetValue('NOLIG'));
    end;
    Ferme(QQ);
  end
  else //Il existe au moins une liste pour cet article/dépot
  begin
    //Cette liste correspond à la ligne que nous sommes en train de traiter
    If TobLInv.GetValue('NOLIG')=TOBL.GetValue('NOLIG') then
      AdditionQteInventaire(CodeArticle, TOBL.GetDouble('QTE'),TOBL.GetValue('NOLIG'))
    else
    begin //elle ne correspond pas on la duplique
      TOBLiginv := TOB.Create ('lignes inventaire', TOBLInvArt, -1);
      TOBLIginv.Dupliquer(TobLInv, True, True);
      TOBLInv.PutValue('NOLIG', TOBL.GetValue('NOLIG'));
      TOBLInv.PutValue('QTECALC', 0.00);
      AdditionQteInventaire(CodeArticle, TOBL.GetDouble('QTE'),TOBL.GetValue('NOLIG'));
    end;
  end;

end;

procedure TOF_INTEGREFICTPI.ChargeTOBGRILLEwithTOBArticle(TOBL, TOBArticle : TOB);
begin

  TOBL.PutValue('CODEARTICLE',    TobArticle.GetString('GA_CODEARTICLE'));
  TOBL.PutValue('GA_ARTICLE',     TobArticle.GetString('GA_ARTICLE'));
  TOBL.PutVALUE('Article'    ,    TobArticle.GetString('GA_ARTICLE'));
  TOBL.PutVALUE('Lib_Article',    TOBArticle.GetString('GA_LIBELLE'));
  //
  IF TobArticle.FieldExists('BCB_CODEBARRE') Then
    TOBL.PutVALUE('CAB_Article',  TOBArticle.GetString('BCB_CODEBARRE'))
  else
    TOBL.PutVALUE('CAB_Article',  TOBArticle.GetString('GA_CODEBARRE'));

  IF TobArticle.FieldExists('BCB_QUALIFCODEBARRE') Then
    TOBL.PutVALUE('Art_QualifCAB',TOBArticle.GetString('BCB_QUALIFCODEBARRE'))
  else
    TOBL.PutVALUE('Art_QualifCAB',TOBArticle.GetString('GA_QUALIFCODEBARRE'));
  //
  TOBL.PutVALUE('QUALIFUNITESTO', TobArticle.GetString('GA_QUALIFUNITESTO'));
  TOBL.PutVALUE('QUALIFUNITEACH', TobArticle.GetString('GA_QUALIFUNITEACH'));
  TOBL.PutVALUE('QUALIFUNITEVTE', TobArticle.GetString('GA_QUALIFUNITEVTE'));
  //
  TOBL.PutVALUE('CoefUSUV',       TobArticle.GetString('GA_COEFCONVQTEVTE'));
  TOBL.PutVALUE('CoefUAUS',       TobArticle.GetString('GA_COEFCONVQTEACH'));
  TOBL.PutVALUE('PAHT',           TobArticle.GetString('GA_PAHT'));
  TOBL.PutVALUE('PAUA',           TobArticle.GetString('GA_PAUA'));
  TOBL.PutVALUE('DPA',            TobArticle.GetString('GA_DPA'));
  TOBL.PutVALUE('PMAP',           TobArticle.GetString('GA_PMAP'));

  TOBL.PutVALUE('ART_FERME',      TobArticle.GetString('GA_FERME'));

  If TOBL.GetString('ART_FERME')='X'    then
  begin
    TOBL.PutValue('ANOMALIEMAR','F');
    TOBL.PutValue('ERREUR', 'X');
    TOBL.PutValue('LIBANOMALIEMAR'   ,'Art. Fermé');
  end;

  TOBL.PutVALUE('TENUESTOCK',     TobArticle.GetString('GA_TENUESTOCK'));

  If TOBL.GetString('TENUESTOCK')='-'  then
  Begin
    TOBL.PutValue('ANOMALIEMAR','S');
    TOBL.PutValue('ERREUR', 'X');
    TOBL.PutValue('LIBANOMALIEMAR','Art. non tenu en Stock');
  end;

end;

procedure TOF_INTEGREFICTPI.ChargeInfoArticle(TOBL : TOB);
begin

  PArticle.TabVisible  := True;

  CodeArticle.text     := TOBL.GetString('CODEARTICLE');
  Article.text         := TOBL.GetString('GA_ARTICLE');
  //
  Lib_Article.Caption  := TOBL.GetString('Lib_Article');
  //
  CAB_Article.text     := TOBL.GetString('CAB_Article');

  Art_QualifCAB.Value  := TOBL.GetString('Art_QualifCAB');
  //
  QualifUS.Value       := TOBL.GetString('QUALIFUNITESTO');
  QualifUA.Value       := TOBL.GetString('QUALIFUNITEACH');
  QualifUV.Value       := TOBL.GetString('QUALIFUNITEVTE');
  //
  CoefUSUV.Text        := TOBL.GetString('CoefUSUV');
  CoefUAUS.text        := TOBL.GetString('CoefUAUS');
  PAHT.Text            := TOBL.GetString('PAHT');
  PAUA.text            := TOBL.GetString('PAUA');
  DPA.text             := TOBL.GetString('DPA');
  PMAP.text            := TOBL.GetString('PMAP');

  IF CodeArticle.text = '' then
  begin
    TOBL.PutValue('ANOMALIEMAR', 'X');
    TOBL.PutValue('ERREUR', 'X');
  end;

  if TOBL.GetString('ANOMALIEMAR')='X' then exit;

  If TOBL.GetString('GA_FERME')='X'    then
    Art_Ferme.Checked := True
  else
    Art_Ferme.Checked := False;
  //
  If TOBL.GetString('TENUESTOCK')='X'  then
  begin
    LIB_TenueStock.visible := False;
    Art_TenueStock.checked := True;
  end
  else
  begin
    Art_TenueStock.checked := True;
    LIB_TenueStock.visible := false;
    //
    if PGIAsk('Cet Article : ' + Tobl.GetString('CODEARTICLE') + ' n''est pas tenu en stock, voulez-vous le modifier ?', 'Intégration TPI')=Mryes then
    begin
      StSQL := 'UPDATE ARTICLE SET GA_TENUESTOCK="X" WHERE GA_TYPEARTICLE="MAR" AND GA_ARTICLE="' + Article.Text + '"';
      If ExecuteSQL(StSQL)=0 then
      begin
        PgiError('La mise à jour de l''article a echouée !','Mise à jour Article');
        TOBL.PutValue('ANOMALIEMAR','S');
        TOBL.PutValue('ERREUR', 'X');
        TOBL.PutValue('LIBANOMALIEMAR', 'Art. non tenu en Stock');
      end
      else
      begin
        Art_TenueStock.checked := True;
        LIB_TenueStock.visible := false;
        TOBL.PUTVALUE('TENUESTOCK','X');
        TOBL.PutValue('ANOMALIEMAR','');
        TOBL.PutValue('LIBANOMALIEMAR', '');
        GestionAnomalie(TOBGrille.Detail[Grille.row-1]);
        AffichageGrilleAnomalie;
      end;
    end;
  end;

  //FV1 : 23/08/2017 - FS#2648 - DELABOUDINIERE - Rendre les consos associées à un appel facturable
  If TypeAffaire.Value = 'W' then
  begin
    ChkConsoFact.visible := True;
  end;

  if TOBL.GetString('FACTURABLE') = 'A' then
    ChkConsoFact.Checked := True
  else
    ChkConsoFact.Checked := False;

end;

Procedure TOF_INTEGREFICTPI.InitZoneEcran;
begin

  //On charge l'onglet Généralité
  TypeMVT.Value := '';
  DateMvt.Text  := '01/01/1900';
  Analyse.Text  := '';
  Depot.Text    := '';
  //
  InitZoneAffaire;
  //
  InitZoneFournisseur;
  //
  InitZoneArticle;

end;

procedure TOf_INTEGREFICTPI.InitZoneAffaire;
begin
  //
  //if TypeMVT.Value = 'BT' then
  //  TypeAffaire.Value := 'W'
  //else
  //  TypeAffaire.Value := 'A';
  //
  Client.Text   := '';
  Affaire.Text  := '';
  Affaire0.Text := TypeAffaire.Value;
  Affaire1.text := '';
  Affaire2.text := '';
  Affaire3.Text := '';
  Avenant.Text  := '';
  //
  Lib_Affaire.Caption     := '';
  CAB_Affaire.Text        := '';
  Aff_QualifCAB.Value     := '';
  Lib_Responsable.Caption := '...';
  Lib_Domaine.Caption     := '...';
  //

end;

Procedure TOF_INTEGREFICTPI.InitZoneAffaireToTob(TOBL : TOB);
begin

  TOBL.PutVALUE('CLIENT',      '');

  TOBL.PutValue('ANOMALIECLI', '');
  TOBL.PutValue('ANOMALIEAFF', '');

  TOBL.PutValue('ANOMALIEBC',  '');
  TOBL.PutValue('LIBANOMALIEBC','');

  TOBL.PutVALUE('LIB_CLIENT',  '...');
  //TOBL.PutVALUE('CODEAFFAIRE', '');
  TOBL.PutVALUE('AFF_AFFAIRE', '');
  TOBL.PutVALUE('LIB_AFFAIRE', '...');
  //
  TOBL.PutVALUE('CAB_Affaire',   '');
  TOBL.PutVALUE('AFF_QUALIFCAB', '');
  //
  TOBL.PutVALUE('Lib_Responsable'  , '...');
  TOBL.PutVALUE('Domaine'          , '...');
  TOBL.PutVALUE('Etablissement'    , '...');
  TOBL.PutVALUE('Lib_Domaine'      , '...');
  //
  TOBL.PutVALUE('AFF_FERME',   '-');
  //
end;

procedure TOf_INTEGREFICTPI.InitZoneFournisseur;
begin

  CodeTiers.Text      := '';
  CAB_Tiers.Text      := '';
  Frs_QualifCAB.Value := '';
  //
  Juridique.Caption   := '...';
  Nom.caption         := '...';
  ADR1.Caption        := '...';
  ADR2.Caption        := '...';
  ADR3.Caption        := '...';
  ADR4.Caption        := '...';
  CP.Caption          := '...';
  Ville.Caption       := '...';

end;

Procedure TOF_INTEGREFICTPI.InitZoneFournisseurToTob(TOBL : TOB);
begin

  TOBL.PutValue('T_TIERS',          '');
  //TOBL.PutValue('CODEFOURNISSEUR',  '');

  TOBL.PutVALUE('CAB_Tiers',        '');
  //
  TOBL.PutVALUE('Frs_QualifCAB',    '');

  TOBL.PutVALUE('Juridique',        '');

  TOBL.PutVALUE('Nom',  '');
  TOBL.PutVALUE('ADR1', '');
  TOBL.PutVALUE('ADR2', '');
  TOBL.PutVALUE('ADR3', '');
  TOBL.PutVALUE('ADR4', '');
  TOBL.PutVALUE('CP',   '');
  TOBL.PutVALUE('Ville','');

  TOBL.PutVALUE('Frs_Ferme', '');

end;

Procedure TOF_INTEGREFICTPI.InitZoneArticle;
Begin
  //
  CodeArticle.text     := '';
  Article.text         := '';
  CAB_Article.text     := '';
  Lib_Article.Caption  := '...';
  Art_QualifCAB.Value  := '';
  QualifUV.Value       := '';
  QualifUS.Value       := '';
  QualifUA.Value       := '';
  CoefUSUV.Text        := '';
  CoefUAUS.text        := '';
  PAHT.Text            := '0.00';
  PAUA.text            := '0.00';
  DPA.text             := '0.00';
  PMAP.text            := '0.00';
  //
  //FV1 : 23/08/2017 - FS#2648 - DELABOUDINIERE - Rendre les consos associées à un appel facturable
  if TOBGrille.detail.count > 0 then
  begin
    if TOBGrille.detail[Grille.Row - 1].GetString('FACTURABLE') = 'A' then
      ChkConsoFact.Checked := True
    else
      ChkConsoFact.Checked := False;
  end;
    
end;

Procedure TOF_INTEGREFICTPI.InitZoneArticleToTob(TOBL : TOB);
begin

  //TOBL.PutValue('CODEARTICLE',    '');
  TOBL.PutVALUE('Article'    ,    '');
  TOBL.PutVALUE('Lib_Article',    '');
  //
  TOBL.PutVALUE('CAB_Article',    '');
  TOBL.PutVALUE('Art_QualifCAB',  '');
  //
  TOBL.PutVALUE('QUALIFUNITESTO', '');
  TOBL.PutVALUE('QUALIFUNITEACH', '');
  TOBL.PutVALUE('QUALIFUNITEVTE', '');
  //
  TOBL.PutVALUE('CoefUSUV',       '');
  TOBL.PutVALUE('CoefUAUS',       '');
  TOBL.PutVALUE('PAHT',           0);
  TOBL.PutVALUE('PAUA',           0);
  TOBL.PutVALUE('DPA',            0);
  TOBL.PutVALUE('PMAP',           0);

  TOBL.PutVALUE('GA_FERME',       '-');

  TOBL.PutVALUE('TENUESTOCK',     '-');

  TOBL.PutValue('FACTURABLE',     'N');

end;

procedure TOF_INTEGREFICTPI.RechercheAffaire;
Var StChamps  : String;
    CodeBarre : String;
    IP        : Integer;
begin

  StChamps := Affaire.text;
  if StChamps = '' then StChamps := TOBGrille.detail[Grille.row-1].GetString('CODEAFFAIRE');

  //ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect1, TaModif, Affaire.text, True);

  if Affaire.text = '' then Affaire.text  := DechargeCleAffaire(Affaire0, Affaire1, Affaire2, Affaire3, Avenant, Client.text, Taconsult, False, True, false, IP);

  if GetAffaireEnteteSt(AFFAIRE0, Affaire1, Affaire2, Affaire3, AVENANT, Client, CodeBarre, false, false, false, True, true,'') then AFFAIRE.text := CodeBarre;

  If Affaire.Text <> '' then
  begin
    //ChargeTobAffaire(TOBGrille.detail[Grille.row-1], Affaire.text);
    Question := Question + 'à cette affaire ?';
    GestionModif('CODEAFFAIRE',Affaire.text,StChamps,SG_AFFAIRE);
  end;

end;

Procedure TOF_INTEGREFICTPI.RechercheFournisseur;
Var StChamps  : string;
begin

  CodeTiers.SetFocus;

  StChamps  := CodeTiers.Text;
  if stChamps = '' then Stchamps := TobGrille.Detail[Grille.row-1].getString('CODEFOURNISSEUR');

  CodeTiers.Text := AGLLanceFiche('GC','GCFOURNISSEUR_MUL','T_TIERS=' + CodeTiers.text +';T_NATUREAUXI=FOU','','SELECTION');

  Question := Question + 'à ce Fournisseur ?';
  GestionModif('CODEFOURNISSEUR',Codetiers.Text, StChamps, SG_FOURNISSEUR);

end;

Procedure TOF_INTEGREFICTPI.RechercheArticle;
Var StChamps  : string;
    SWhere    : string;
begin

  if CodeArticle.text <> '' then
  begin
    SoustractionQteInv(CodeArticle.text, StrtoFloat(Qte.text),TobGrille.Detail[Grille.row-1].getValue('NOLIG'));
  end;

	sWhere := ' AND ((GA_TYPEARTICLE="MAR") AND GA_TENUESTOCK="X")';

  StChamps := CodeArticle.Text;
  if stChamps = '' then Stchamps := TobGrille.Detail[Grille.row-1].getString('CODEARTICLE');

  if CodeArticle.Text <> '' then
    sWhere := 'GA_CODEARTICLE=' + Trim(Copy(CodeArticle.text, 1, 18)) + ';XX_WHERE=' + sWhere
  else
    sWhere := 'XX_WHERE=' + sWhere;

	CodeArticle.text := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', sWhere+';RECHERCHEARTICLE');
  CodeArticle.text := Trim(Copy(CodeArticle.text, 1, 18));

  If CodeArticle.Text <> '' then
  begin
    Question := question + 'à cet Article ?';
    GestionModif('CODEARTICLE',CodeArticle.Text, StChamps,SG_ARTICLE);
  end;

end;

procedure TOf_INTEGREFICTPI.RechercheCABAffaire;
Var StChamps : String;
begin

  Stchamps := Cab_Affaire.text;
  if stChamps = '' then Stchamps := TobGrille.Detail[Grille.row-1].getString('CAB_AFFAIRE');

  CAB_Affaire.Text := AGLLanceFiche('BTP', 'BTCODEBARRE_MUL', '', '', '');

  If CAB_Affaire.Text <> '' then
  begin
    Question := Question + 'à cette affaire ?';
    GestionModif('CODEAFFAIRE',CAB_Affaire.text,StChamps,SG_AFFAIRE);
  end;

end;

procedure TOf_INTEGREFICTPI.RechercheCABFournisseur;
begin

end;

procedure TOf_INTEGREFICTPI.RechercheCABArticle;
begin

end;

procedure TOF_INTEGREFICTPI.GestionAnomalie(TOBLGrille: TOB);
Var TOBL    : TOB;
begin

  TOBLGrille.PutValue('TYPEERREUR', '');

  TOBL := TobAnomalie.FindFirst(['NOLIG'], [TOBLGrille.GetValue('NOLIG')], false);
  If TOBL <> nil then TOBL.Free;

  LibErreur := '';
  LibErreur := GenerationMsgAnomalie(TOBLGrille);
  //
  if LibErreur <> '' then
  begin
    TOBL := TOB.Create('Ligne Anomalie', TOBAnomalie,-1);
    TOBL.Dupliquer(TOBLGrille, True, true);
    TOBL.PutValue('TYPEERREUR', LibErreur);
    TOBLGrille.PutValue('TYPEERREUR', LibErreur);
  end;

  TobAnomalie.Detail.Sort('NOLIG');
  
end;

Procedure TOF_INTEGREFICTPI.BTGenerationOnClick(Sender: TObject);
Var Indice    : Integer;
    TOBL      : TOB;
    TobTempo  : TOB;

    SCodeFrs  : String; //Ancien Code
    CodeFrs   : String; //Nouveau Code

    STypeMvt  : String;
    TypeMvt   : String;

    SCodeAff  : String;
    CodeAff   : String;

    SDateMvt  : TDateTime;
    DateMvt   : TDateTime;

    SNumCde   : String;
    NumCde    : String;
    NumBesoin : String;
    SNumBesoin: String;

    LigToDel  : Integer;

    OkStop    : Boolean;
begin

  BGeneration.enabled := False;

  OkStop := False;

  ChargeTobWithScreen;

  //On charge le stock des articles avec leurs entrées totales et leurs sorties...
  ChargementStockArticle(TobGrille);

  //Lecture de la tob pour déterminer les article dont la qté de stock physique se retrouvera en négatif à la fin de l'intégration
  OkStop := ControleStockAvantIntegration(TOBGrille);

  {if OkStop then
  begin
    Pages.ActivePage      := PGeneralite;
    AffichageGrilleAnomalie(TobAnomalie);
    BGeneration.enabled   := True;
    Grille1.Row           := 1;
    Exit;
  end;}

  //
  if TobAnomalie.detail.count <> 0 then
  begin
    if PGIAsk('Attention, La liste de vos anomalies d''intégration n''est pas vide... voulez-vous tout de même continuer ?', 'Anomalies')= MrNo then
    begin
      //On se positionne sur les bonnes pages et les bons onglets
      Pages.ActivePage    := PGeneralite;
      Grille1.Row         := 1;
      BGeneration.enabled := True;
      exit;
    end;
  end;

  //
  if not ArchivageFichierTPI then
  begin
    PGIError('Une erreur c''est produit lors de l''archivage de ' + fichierImport.text + ' en ' + FichierHisto, 'Erreur d''archivage');
    //On se positionne sur les bonnes pages et les bons onglets
    Pages.ActivePage      := PGeneralite;
    AffichageGrilleAnomalie;
    Grille1.Row           := 1;
    BGeneration.enabled   := True;
    Exit;
  end;

  //après Archivage on supprime le fichier import
  if not SysUtils.DeleteFile(fichierImport.text) then
  begin
    PGIError('Une erreur s''est produite lors de la suppression de ' +  fichierImport.text, 'Suppression');
    //On se positionne sur les bonnes pages et les bons onglets
    Pages.ActivePage      := PGeneralite;
    BGeneration.enabled   := True;
    Exit;
  end;

  LigToDel  := 1;
  TOBTempo := Tob.Create('TEMPO',nil,-1);
  TobTempo.Dupliquer(TobGrille,true, true);

  //génération de la tob pour génération des documents en fonction des lignes lues
  For indice := 0 to TobTempo.detail.count -1 do
  begin
    TOBL      := TOBTempo.detail[indice];
    //
    if TOBL.GetValue('ANOMALIE') = 'STOCK' then
    Begin
      MajTobRejet(TOBL, 'REJ', TOBL.GetValue('TYPEERREUR'));
      Continue;
    end;

    if TOBL.GetValue('INTEGRE') = '-' then
    Begin
      MajTobRejet(TOBL, 'REJ', 'Rejeté à l''intégration par l''utilisateur');
      Continue;
    end;
    //
    TypeMvt  := TOBL.GetString('TYPEMVT');
    CodeFrs  := TOBL.GetString('CODEFOURNISSEUR');
    CodeAff  := TOBL.GetString('CODEAFFAIRE');
    DateMvt  := StrToDate(TOBL.GetString('DATEMVT'));
    NumCde   := TOBL.GetString('NUMCDE');
    NumBesoin:= TOBL.GetString('NUMBESOIN');

    //
    if not GestionRejet(TOBL) then
    Begin
      STypeMvt  := TOBL.GetString('TYPEMVT');
      SCodeFrs  := TOBL.GetString('CODEFOURNISSEUR');
      SCodeAff  := TOBL.GetString('CODEAFFAIRE');
      SDateMvt  := StrToDate(TOBL.GetString('DATEMVT'));
      SNumCde   := TOBL.GetString('NUMCDE');
      SNumBesoin:= TOBL.GetString('NUMBESOIN');
      LigToDel  := LigToDel + 1;
      continue;
    end;
    //
    //Rupture sur piéce en fonction du type de mouvement, du fournisseur et de l'affaire....
    //
    if StypeMvt <> TypeMvt then
    begin
      GestionRuptureTypeMvt(TOBL, STypeMvt, LigToDel);
      STypeMvt  := TOBL.GetString('TYPEMVT');
      SCodeFrs  := TOBL.GetString('CODEFOURNISSEUR');
      SCodeAff  := TOBL.GetString('CODEAFFAIRE');
      SDateMvt  := StrToDate(TOBL.GetString('DATEMVT'));
      SNumCde   := TOBL.GetString('NUMCDE');
      SNumBesoin:= TOBL.GetString('NUMBESOIN');
      if TypeMvt = 'IN' then Continue;
    end
    else
    begin
      if TypeMvt = 'IN' then //Si on est sur un inventaire il faut le créer quelque soit la rupture...
      begin
        if not GenerationDeLaPiece(TOBL, TypeMvt) then
           LigToDel := LigToDel + 1
        else
        begin
          Grille.DeleteRow(LigToDel);
        end;
        //
        if LigToDel < TobGrille.Detail.count -1 then TOBGrille.Detail[LigToDel].Free;
        Continue;
      end
      else if  (TypeMvt = 'SA') then
      begin
        if SNumBesoin <> NumBesoin then //rupture su Numéro de besoin
        begin
          if SCodeAff <> CodeAff then //Rupture sur l'affaire
          begin
            GestionRuptureChantier(TOBL, SCodeAff);
            SCodeAff  := TOBL.GetString('CODEAFFAIRE');
            SNumBesoin := TOBL.GetString('NUMBESOIN');
          end
          else //On garde la même affaire
          begin
            If DateMvt <> SDateMvt then //Même Besoin mais pas la même date rupture
            begin
              GestionRuptureDate(TOBL, SDateMvt);
              SDateMvt  := StrToDate(TOBL.GetString('DATEMVT'));
              SNumBesoin := TOBL.GetString('NUMBESOIN');
            end
            else //On Garde la même date
            begin
              GestionRuptureBesoin(TOBL, SNumBesoin);
              SNumBesoin := TOBL.GetString('NUMBESOIN');
            end;
          end;
        end
        else
        begin
          if SCodeAff <> CodeAff then //Rupture sur l'affaire
          begin
            GestionRuptureChantier(TOBL, SCodeAff);
            SCodeAff  := TOBL.GetString('CODEAFFAIRE');
            SNumBesoin := TOBL.GetString('NUMBESOIN');
          end
          else
          begin
            If DateMvt <> SDateMvt then //Même Besoin mais pas la même date rupture
            begin
               GestionRuptureDate(TOBL, SDateMvt);
               SDateMvt  := StrToDate(TOBL.GetString('DATEMVT'));
            end;
          end;
        end
      end
      else if (TypeMvt = 'BL') then
      begin
        if SNumCde <> NumCde then
        begin
          if SCodeFrs <> CodeFrs   then
          begin
            GestionRuptureFournisseur(TOBL, SCodeFrs);
            SCodeFrs  := TOBL.GetString('CODEFOURNISSEUR');
            SNumCde   := TOBL.GetString('NUMCDE');
          end
          else
          begin
            GestionRuptureNumCde(TOBL, SNumCde);
            SNumCde := TOBL.GetString('NUMCDE');
          end;
        end
        else if DateMvt <> SDateMvt then
        begin
           GestionRuptureDate(TOBL, SDateMvt);
           SDateMvt  := StrToDate(TOBL.GetString('DATEMVT'));
        end
        else if SCodeFrs <> CodeFrs   then
        begin
          GestionRuptureFournisseur(TOBL, SCodeFrs);
          SCodeFrs  := TOBL.GetString('CODEFOURNISSEUR');
        end
        else
        begin
          If SCodeAff <> CodeAff then
          begin
            GestionRuptureChantier(TOBL, SCodeAff);
            SCodeAff  := TOBL.GetString('CODEAFFAIRE');
          end;
        end;
      end
      else if (TypeMvt='ED') Or (TypeMvt = 'SD') then
      begin
        if DateMvt <> SDateMvt then
        begin
           GestionRuptureDate(TOBL, SDateMvt);
           SDateMvt  := StrToDate(TOBL.GetString('DATEMVT'));
        end
        else if SCodeFrs <> CodeFrs   then
        begin
          GestionRuptureFournisseur(TOBL, SCodeFrs);
          SCodeFrs  := TOBL.GetString('CODEFOURNISSEUR');
        end
        else
        begin
          If SCodeAff <> CodeAff then
          begin
            GestionRuptureChantier(TOBL, SCodeAff);
            SCodeAff  := TOBL.GetString('CODEAFFAIRE');
          end;
        end;
      end
      else if (TypeMvt = 'EA') OR (TypeMvt = 'CD') then
      begin
        if SCodeFrs <> CodeFrs   then
        begin
          GestionRuptureFournisseur(TOBL, SCodeFrs);
          SCodeFrs  := TOBL.GetString('CODEFOURNISSEUR');
        end
        else
        begin
          If SCodeAff <> Codeaff then
          begin
            GestionRuptureChantier(TOBL, SCodeAff);
            SCodeAff  := TOBL.GetString('CODEAFFAIRE');
          end;
        end;
      end;
    end;

    //création des lignes qui vont bien....
    TobLigne := TOB.Create('LES LIGNES', TOBPiece,-1);

    //Création des champs de la tob ligne de piece
    CreateLigne(TobLigne);

    //Chargement des champs de la tob ligne de piece
    ChargeTobLigne(STypeMVT, TOBL, TOBLigne);
    //
    //Chargement de la tob de validation de la ligne
    GestionValide(TOBL);
    //
    STypeMvt    := TOBL.GetString('TYPEMVT');
    SCodeFrs    := TOBL.GetString('CODEFOURNISSEUR');
    SCodeAff    := TOBL.GetString('CODEAFFAIRE');
    SDateMvt    := StrToDate(TOBL.GetString('DATEMVT'));
    SNumCde     := TOBL.GetString('NUMCDE');
    SNumBesoin  := TOBL.GetString('NUMBESOIN');

    //
    Grille.DeleteRow(LigToDel);
    if LigToDel < TobGrille.Detail.count -1 then TOBGrille.Detail[LigToDel].Free;
    //
  end;

  if TobPiece.detail.count <> 0 then
  begin
    if CreatePieceFromTob(TOBPiece, nil, nil,nil) then
      MajTobValide
    else
      MajTobRejet(TOBL, 'PIE', 'Erreur en création de la pièce : ' + RechDom('GPP',TOBL.GetValue('GP_NATUREPIECEG'),False) + '-' + IntToStr(TOBL.GetValue('GP_NUMERO')));
  end;
  //
  FreeAndNil(TobTempo);
  //
  ReinitialiseTobPiece;
  //
  //edition des lignes rejetées
  EditeRejete;

  EditeValide;
  //
  if TobRejete.detail.count <> 0 then
  begin
    Grille.row := 1;
    //
    Tobanomalie.ClearDetail;
    TobGrille.ClearDetail;
    //
    TobGrille.Dupliquer(Tobrejete,true, true);
    //
    GestionGS.TOBG  := TOBGrille;
    GestionGS.GS    := Grille;
    //
    ChargementGrille(GestionGS);
    //
    For indice := 0 to TobGrille.detail.count -1 do
    begin
      Grille.Row := Indice+1;
      TobGrille.detail[Indice].putValue('NOLIG', Grille.Row);
      GestionAnomalie(TOBGrille.Detail[Grille.row-1]);
    end;
    //
    AffichageGrilleAnomalie;
    //
    Grille.row := 1;
    //
    TobRejete.ClearDetail;
    TobValide.ClearDetail;
    //
    PGIBox('Traitement terminé avec des erreurs', 'Intégration fichier : ' + FichierImport.text);
    Pages.enabled := False;
  end
  else
  Begin
    PGIBox('Traitement terminé sans erreur', 'Intégration fichier : ' + FichierImport.text);
    Pages.enabled := False;
  end;

end;

Procedure TOF_INTEGREFICTPI.CreateLigne(TobLigne : TOB);
begin

  TobLigne.AddChampSupValeur('TYPELIGNE',   '');
  TobLigne.AddChampSupValeur('AFFAIRE',     '');
  TobLigne.AddChampSupValeur('CODEARTICLE', '');
  TobLigne.AddChampSupValeur('ARTICLE',     '');
  TobLigne.AddChampSupValeur('LIBELLE',     '');
  TobLigne.AddChampSupValeur('QTEFACT',      0);
  TobLigne.AddChampSupValeur('QUALIFQTE',   '');
  TobLigne.AddChampSupValeur('PUHTDEV',      0);
  TobLigne.AddChampSupValeur('PAHT',         0);
  TobLigne.AddChampSupValeur('DPA',          0);
  TobLigne.AddChampSupValeur('PMAP',         0);
  TobLigne.AddChampSupValeur('FACTURABLE', 'N');

  TobLigne.AddChampSupValeur('DATELIVRAISON', idate1900);
  TobLigne.AddChampSupValeur('DEPOT',       '');
  TobLigne.AddChampSupValeur('PIECEPRECEDENTE', '');
  TobLigne.AddChampSupValeur('GL_PIECEPRECEDENTE', '');
  TobLigne.AddChampSupValeur('PIECEORIGINE', '');

end;

Procedure TOF_INTEGREFICTPI.ChargeTobLigne(STypeMVT : String; TOBL, TOBLigne : TOB);
Var DateLiv   : String;
    Qte       : Double;
    Pa        : Double;
begin

  TobLigne.PutValue('TYPELIGNE',   'ART');
  TobLigne.PutValue('AFFAIRE',     TOBL.GetString('CODEAFFAIRE'));
  TobLigne.PutValue('CODEARTICLE', TOBL.GetString('CODEARTICLE'));
  TobLigne.PutValue('ARTICLE',     TOBL.GetString('GA_ARTICLE'));
  TobLigne.PutValue('LIBELLE',     TOBL.GetString('LIB_ARTICLE'));

  Qte := TOBL.GetDouble('QTE');
  Pa  := TOBL.GetDouble('PAHT');

  if STypeMvt = 'EA' then
  Begin
    if QTE > 0 then Qte := Qte * -1;
  end
  else if STypeMvt = 'SA' then
  Begin
    if Qte < 0 then Qte := Qte * -1;
  end;

  TobLigne.PutValue('QTEFACT',     Qte);
  TobLigne.PutValue('QUALIFQTE',   TOBL.GetString('QUALIFUNITEACH'));
  TobLigne.PutValue('PUHTDEV',     TOBL.GetDouble('PAUA'));
  TobLigne.PutValue('PAHT',        TOBL.GetDouble('PAHT'));
  TobLigne.PutValue('DPA',         TOBL.GetDouble('DPA'));
  TobLigne.PutValue('PMAP',        TOBL.GetDouble('PMAP'));
  TobLigne.PutValue('DATELIVRAISON', idate1900);

  if STypeMVT = 'CD' then
  begin
    DateLiv := DateToStr(Now);
    TobLigne.PutValue('DATELIVRAISON',DateLiv);
  end
  else if StypeMVT = 'BL' then
  begin
    RecupLignePiecePrecedente(TOBL, TobLigne);
  end
  else if StypeMVT = 'SA' then
  begin
    TobLigne.PutValue('PIECEPRECEDENTE',    TOBL.GetValue('PIECEPRECEDENTE'));
    TobLigne.PutValue('GL_PIECEPRECEDENTE', TOBL.GetValue('PIECEPRECEDENTE'));
  end;

  if Depot.text = '' then
    TobLigne.PutValue('DEPOT', GetParamSocSecur('SO_GCDEPOTDEFAUT', '', False))
  else
    TobLigne.PutValue('DEPOT', Depot.text);

  TobLigne.PutValue('FACTURABLE',     TOBL.GetString('FACTURABLE'));

end;

Procedure TOF_INTEGREFICTPI.RecupLignePiecePrecedente(TOBL, TobLigne : TOB);
VAR TOBLignePiece : TOB;
begin

  If TypeMVT.Value <> 'BL' then Exit;

  StSQL := 'SELECT GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, GL_INDICEG, GL_AFFAIRE, ';
  StSQL := STSQL + ' GL_NUMORDRE, GL_PIECEORIGINE, GL_DATEPIECE, GL_IDENTIFIANTWOL FROM LIGNE';
  //
  StSQL := StSQL + ' WHERE GL_NATUREPIECEG="CF" AND GL_NUMERO=' + InttoStr(TOBL.GetValue('NUMCDE'));
  //
  StSQL := StSQL + '   AND GL_ARTICLE="' + TOBL.GetString('GA_ARTICLE') + '"';

  QQ := OpenSQL(STSQL, False,1,'',True);

  if Not QQ.eof then
  begin
    TOBLignePiece := Tob.Create('LIGNE',nil,-1);
    TOBLignePiece.selectDB('LIGNE', QQ);
    TOBLigne.PutValue('PIECEPRECEDENTE', EncodeRefPiece(TOBLignePiece));
    TOBLigne.PutValue('PIECEORIGINE',    EncodeRefPiece(TOBLignePiece));
    if TOBLigne.GetString('GL_AFFAIRE') <> '' then
    begin
      IF TOBLIGNE.GetValue('GL_IDENTIFIANTWOL') = 1 then
      begin
        StSQL := 'UPDATE LIGNE SET GL_IDENTIFIANTWOL=0 ';
        StSQL := StSQL + ' WHERE GL_NATUREPIECEG="CF" AND GL_NUMERO=' + InttoStr(TOBL.GetValue('NUMCDE'));
        StSQL := StSQL + '   AND GL_ARTICLE="' + TOBL.GetString('GA_ARTICLE') + '"';
      end;
    end;
    TOBLignePiece.free;
  end;

  Ferme(QQ);

end;


Procedure TOF_INTEGREFICTPI.GestionRuptureTypeMvt(TOBL : TOB; AncienTypeMvt: String; Var LigToDel : Integer);
begin

  if (AncienTypeMvt <> '') then
  Begin
    if CreatePieceFromTob(TOBPiece, nil, nil,nil) then
      MajTobValide
    else
      MajTobRejet(TOBL, 'PIE', 'Erreur en création de la pièce : ' + RechDom('GPP',TOBL.GetValue('GP_NATUREPIECEG'),False) + '-' + IntToStr(TOBL.GetValue('GP_NUMERO')));
  end;

  ReinitialiseTobPiece;

  if TOBL.GetString('TYPEMVT') = 'IN' then
  begin
    if not GenerationDeLaPiece(TOBL,  TOBL.GetString('TYPEMVT')) then
       LigToDel := LigToDel + 1
    else
    Begin
      Grille.DeleteRow(LigToDel);
    end;
    if LigToDel < TobGrille.Detail.count -1 then TOBGrille.Detail[LigToDel].Free;
  end
  else
  begin
    if Not GenerationDeLaPiece(TOBL,  TOBL.GetString('TYPEMVT')) then Exit;
  end;

end;

Procedure TOF_INTEGREFICTPI.GestionRuptureFournisseur(TOBL : TOB; AncienFrs : String);
begin

  If AncienFrs <> ''     then
  begin
    if CreatePieceFromTob(TOBPiece, nil, nil,nil) then
      MajTobValide
    else
      MajTobRejet(TOBL, 'PIE', 'Erreur en création de la pièce : ' + RechDom('GPP',TOBL.GetValue('GP_NATUREPIECEG'),False) + '-' + IntToStr(TOBL.GetValue('GP_NUMERO')));
  end;

  ReinitialiseTobPiece;

  if Not GenerationDeLaPiece(TOBL, TOBL.GetString('TYPEMVT')) then Exit;

end;

Procedure TOF_INTEGREFICTPI.GestionRuptureChantier(TOBL : TOB; AncienAff : String);
begin

  If AncienAff <> ''     then
  begin
    if CreatePieceFromTob(TOBPiece, nil, nil,nil) then
      MajTobValide
    else
      MajTobRejet(TOBL, 'PIE', 'Erreur en création de la pièce : ' + RechDom('GPP',TOBL.GetValue('GP_NATUREPIECEG'),False) + '-' + IntToStr(TOBL.GetValue('GP_NUMERO')));
  end;

  ReinitialiseTobPiece;

  if Not GenerationDeLaPiece(TOBL, TOBL.GetString('TYPEMVT')) then Exit;

end;


Procedure TOF_INTEGREFICTPI.GestionRuptureDate(TOBL : TOB; AncienDateMvt : TDateTime);
begin

  If AncienDateMvt <> idate1900 then
  begin
    if CreatePieceFromTob(TOBPiece, nil, nil,nil) then
      MajTobValide
    else
      MajTobRejet(TOBL, 'PIE', 'Erreur en création de la pièce : ' + RechDom('GPP',TOBL.GetValue('GP_NATUREPIECEG'),False) + '-' + IntToStr(TOBL.GetValue('GP_NUMERO')));
  end;

  ReinitialiseTobPiece;

  if Not GenerationDeLaPiece(TOBL, TOBL.GetString('TYPEMVT')) then Exit;

end;

Procedure TOF_INTEGREFICTPI.GestionRuptureNumCde(TOBL : TOB; AncienNumero : String);
begin

  If AncienNumero <> '' then
  begin
    // Création TOB pour ligne de pièce
    TobLigne := Tob.Create ('LES LIGNES', TobPiece, -1);
    CreateLigne(TobLigne);
    ChargeLigneCommentaire('BL', TOBLIGNE, TOBL);
  end;

end;

Procedure TOF_INTEGREFICTPI.GestionRuptureBesoin(TOBL : TOB; AncienNumero : String);
begin

  If AncienNumero <> '' then
  begin
    // Création TOB pour ligne de pièce
    TobLigne := Tob.Create ('LES LIGNES', TobPiece, -1);
    CreateLigne(TobLigne);
    ChargeLigneCommentaire('SA', TOBLIGNE, TOBL);
  end;

end;


Procedure TOF_INTEGREFICTPI.TriTableGrille;
Var Indice : Integer;
begin

  //If TypeMVT.Value = 'BL' then
  //  TobGrille.Detail.Sort('TYPEMVT;DEPOT;DATEMVT;CODEFOURNISSEUR;CODEAFFAIRE;NUMCDE')
  //Else If TypeMVT.Value = 'SA' then
  //  TobGrille.Detail.Sort('TYPEMVT;DEPOT;DATEMVT;CODEFOURNISSEUR;CODEAFFAIRE;NUMBESOIN')
  //else

  TobGrille.Detail.Sort('TYPEMVT;DEPOT;CODEAFFAIRE;CODEFOURNISSEUR;DATEMVT;NUMCDE;NUMBESOIN');

  Grille.row := 1;

  Tobanomalie.ClearDetail;
  //
  //GestionGS.TOBG  := TOBGrille;
  //GestionGS.GS    := Grille;
  //
  //ChargementGrille(GestionGS);
  //
  //For indice := 0 to TobGrille.detail.count -1 do
  //begin
  //  Grille.Row := Indice+1;
  //  TobGrille.detail[Indice].putValue('NOLIG', Grille.Row);
  //  GestionAnomalie(TOBGrille.Detail[Indice]); //?????
  //end;
  //
  Grille.row := 1;
  //
  GestionAnomalie(TOBGrille.Detail[0]);
  //
end;

Procedure TOF_INTEGREFICTPI.ReinitialiseTobPiece;

begin

  TobPiece.ClearDetail;

  TOBPiece.InitValeurs;

end;

Procedure TOF_INTEGREFICTPI.MajTobValide;
Var TOBLValide  : TOB;
begin

  if TobValide = nil then Exit;

  //Mise à jour de la TobValide
  if (TobPiece.GetString('TYPEMVT')='SA') or (TobPiece.GetString('TYPEMVT')='EA') then
    TOBLValide := TobValide.FindFirst(['TYPEMVT','CLIENT','CODEAFFAIRE','DATEMVT'],[TobPiece.GetString('TYPEMVT'),TobPiece.GetString('TIERS'),TobPiece.GetString('AFFAIRE'),TobPiece.GetString('DATEPIECE')], False)
  else
    TOBLValide := TobValide.FindFirst(['TYPEMVT','CODEFOURNISSEUR','CODEAFFAIRE','DATEMVT'],[TobPiece.GetString('TYPEMVT'),TobPiece.GetString('TIERS'),TobPiece.GetString('AFFAIRE'),TobPiece.GetString('DATEPIECE')], False);

  Repeat
    if ToblValide = nil then Exit;
    TOBLValide.PutValue('PIECEGENEREE', RechDom('GCNATUREPIECEG', TObPiece.GetString('GP_NATUREPIECEG'), False) + ' N° ' + IntToStr(TobPiece.GetValue('GP_NUMERO')) + ' du ' + DateToStr(TobPiece.GetValue('GP_DATEPIECE')));
  if TobPiece.GetString('TYPEMVT')='SA' then
    TOBLValide := TobValide.FindNext(['TYPEMVT','CLIENT','CODEAFFAIRE','DATEMVT'],[TobPiece.GetString('TYPEMVT'),TobPiece.GetString('TIERS'),TobPiece.GetString('AFFAIRE'),TobPiece.GetString('DATEPIECE')], False)
  else
    TOBLValide := TobValide.FindNext(['TYPEMVT','CODEFOURNISSEUR','CODEAFFAIRE','DATEMVT'],[TobPiece.GetString('TYPEMVT'),TobPiece.GetString('TIERS'),TobPiece.GetString('AFFAIRE'),TobPiece.GetString('DATEPIECE')], False);
  until TOBLValide = nil;

end;

Function TOF_INTEGREFICTPI.GenerationDeLaPiece(TOBL : TOB; TypeMvt : String) : Boolean;
Begin

  Result := true;

  If      TypeMvt = 'IN' then       ///gestion des inventaires
  begin
    Result := MAJInventaire(TOBL);
    exit;
  end
  Else If TypeMvt = 'SA' then      ///Affectation Livraison sur chantier (LBT)
    CreateEnteteDoc('LBT', TOBL)
  Else If TypeMvt = 'EA' then      ///Affectation Retour chantier (BFC)
    CreateEnteteDoc('BFC', TOBL)
  Else If TypeMvt = 'SD' then      ///Affectation Sorties Exceptionnelles (SEX)
  begin
    CreateEnteteDoc('SEX', TOBL);
    Exit;
  end
  Else If TypeMvt = 'ED' then      ///Affectation Entrées Exceptionnelles (EEX)
  Begin
    CreateEnteteDoc('EEX', TOBL);
    Exit;
  end
  Else If TypeMvt = 'CD' then      ///Affectation Commandes Fournisseurs (CF)
    CreateEnteteDoc('CF',  TOBL)
  Else If TypeMvt = 'BT' then      ///Affectation Livraison sur Chantier (???)
    CreateEnteteDoc('LBT', TOBL)
  Else If TypeMvt = 'BL' then      ///Affectation Réception fournisseur
    CreateEnteteDoc('BLF', TOBL);

  If (TypeMvt = 'BL') Or (TypeMvt = 'SA') then
  begin
    // Création TOB pour ligne de pièce
    TobLigne := Tob.Create ('LES LIGNES', TobPiece, -1);
    CreateLigne(TobLigne);
    ChargeLigneCommentaire(TypeMvt, TOBLIGNE, TOBL);
  end;

end;

procedure TOF_INTEGREFICTPI.ChargeLigneCommentaire(TypeMvt : String; TOBLIGNE, TOBL : TOB);
Var Refp : String;
begin

  TobLigne.PutValue('TYPELIGNE',   'COM');
  TobLigne.PutValue('CODEARTICLE', '');
  TobLigne.PutValue('ARTICLE',     '');

  if TypeMvt = 'BL' then
  Begin
    if TOBL.GetString('NUMCDE') <> '' then
    begin
      RefP := RechDom('GCNATUREPIECEG', TOBL.GetValue('NATPIECEPREC'),False);
      RefP := RefP + ' N° ' + TOBL.GetString('NUMCDE');
      RefP := RefP + ' du ' + DateToStr(TOBPiece.GetValue('DATEPIECE'));
      RefP := RefP + ' '    + TOBPiece.GetString('REFINTERNE');
    end;
  end
  else if TypeMvt = 'SA' then
  begin
    if TOBL.GetString('NUMBESOIN') <> '' then
    begin
      RefP := RechDom('GCNATUREPIECEG', TOBL.GetValue('NATPIECEPREC'),False);
      RefP := RefP + ' N° ' + TOBL.GetString('NUMBESOIN');
      RefP := RefP + ' du ' + DateToStr(TOBL.GetValue('DATEBESOIN'));
    end;
  end
  else
  begin
    if TOBPiece.GetString('NUMERO') <> '' then
    begin
      RefP := RechDom('GCNATUREPIECEG', TOBPiece.GetString('NATUREPIECEG'),False);
      RefP := RefP + ' N° ' + TOBPiece.GetString('NUMERO');
      RefP := RefP + ' du ' + DateToStr(TOBPiece.GetValue('DATEPIECE'));
      RefP := RefP + ' '    + TOBPiece.GetString('REFINTERNE');
    end;
  end;
  //
  RefP:=Copy(RefP,1,70) ;
  //
  TobLigne.PutValue('LIBELLE', RefP);
  //
  TobLigne.PutValue('QTEFACT',      0.00);
  TobLigne.PutValue('QUALIFQTE',   '');
  TobLigne.PutValue('PUHTDEV',      0.00);
  TobLigne.PutValue('PAHT',         0.00);
  TobLigne.PutValue('DPA',          0.00);
  TobLigne.PutValue('PMAP',         0.00);
  TobLigne.PutValue('DATELIVRAISON',idate1900);
  //
  if Depot.text = '' then
    TobLigne.PutValue('DEPOT', GetParamSocSecur('SO_GCDEPOTDEFAUT', '', False))
  else
    TobLigne.PutValue('DEPOT', Depot.text);

end;

Function TOF_INTEGREFICTPI.MAJInventaire(TOBL : TOB) : Boolean;
Var TOBLInvArt : TOB;
    TOBLInv    : TOB;
begin

  Result := false;

  //on recherche les lignes d'inventaire correspondant à l'article
  TobLInvArt := TobInventaire.FindFirst(['CODEARTICLE'],[TOBL.GetString('CODEARTICLE')],False);
  If TOBLInvArt = nil then
    MajTobRejet(TOBL,'INV', 'Aucune liste d''inventaire pour ' + TOBL.GetString('CODEARTICLE'))
  else
  begin
    //On Recherche la lignes correspondantes à l'Article/dépôt
    TobLInv := TOBLInvArt.findFirst(['GIL_CODEARTICLE','GIL_DEPOT'],[TOBL.GetString('CODEARTICLE'),TOBL.GetString('DEPOT')],False);
    if TobLInv <> nil then
    begin
      if TOBLinv.getValue('GIE_VALIDATION')='X' then
      begin
        MajTobRejet(TOBL,'INV', 'Liste d''inventaire déjà validée : ' + TOBLinv.GetString('GIE_CODELISTE'));
      end
      else if TOBLinv.getValue('GIL_SAISIINV')='X' then
        MajTobRejet(TOBL,'INV', 'Article déjà inventorié : ' + TOBLinv.GetString('GIL_CODEARTICLE'))
      else
      begin
        Result := True;
        MAJLigneInventaire(TOBLInv, TOBL);
      end;
      FreeAndNil(ToblInv);
    end
    else
      MajTobRejet(TOBL,'INV', 'Aucune liste d''inventaire sur ce dépôt ' + TOBL.GetString('CODEDEPOT'));
  end;

end;

procedure TOF_INTEGREFICTPI.AffichageGrilleInventaire;
var GestionG2 : TGestionGS;
    TOBLInv   : TOB;
begin

  //On Recherche si par hasard l'article possède une Liste Inventaire
  TobLInv := TobInventaire.FindFirst(['CODEARTICLE','CODEDEPOT'],[TOBGrille.Detail[Grille.row+1].GetString('CODEARTICLE'),TOBGrille.Detail[Grille.row+1].GetString('DEPOT')],True);

  //Si Oui...
  if TOBLInv <> nil then
  begin
    //chargement de la grille avec les éléments de la TOB
    GestionG2.fColNamesGS   := fColNamesGrille1;
    GestionG2.FalignementGS := FalignementGrille1;
    GestionG2.FtitreGS      := FtitreGrille1;
    GestionG2.fLargeurGS    := fLargeurGrille1;
    GestionG2.TOBG          := TobLInv;
    GestionG2.GS            := Grille2;
    //
    DessineGrille(GestionG2);
    //
    ChargementGrille(GestionG2);
    //
    Grille2.row := 1;
  end;

end;

procedure TOF_INTEGREFICTPI.MAJLigneInventaire(TOBResult, TOBL : TOB);
Var Qte        : Double;
begin

  Qte := TOBL.Getdouble('QTE');

  StSQL := 'UPDATE LISTEINVLIG SET GIL_INVENTAIRE=' + FloatToStr(Qte);
  StSQL := StSQL + ', GIL_DATESAISIE="' + USDATETIME(StrToDateTime(TOBL.GetString('DATEMVT'))) + '" ';
  StSQL := StSQL + ', GIL_SAISIINV="X" ';
  //if TOBL.GetString('CODEFOURNISSEUR') <> '' then
  //  StSQL := StSQL + ', GIL_FOURNISSEUR="'    + TOBL.GetString('CODEFOURNISSEUR') + '" ';
  StSQL := StSQL + '  WHERE GIL_CODELISTE="'  + TOBResult.GetString('GIE_CODELISTE') + '" ';
  StSQL := STSQL + '    AND GIL_DEPOT="'      + TOBL.GetString('DEPOT') + '" ';
  StSQL := StSQL + '    AND GIL_CODEARTICLE="'+ TOBL.GetString('CODEARTICLE') + '"';

  If ExecuteSQL(StSQL) = 0 then
  begin
    PgiError('La mise à jour de l''inventaire a échouée !','Mise à jour Inventaire');
    MajTobRejet(TOBL, 'INV', 'La mise à jour de l''inventaire a échouée');
  end
  else GestionValide(TOBL);

end;

Procedure TOF_INTEGREFICTPI.CreateEnteteDoc(Nature : String; TOBL : TOB);
Var DateMouv : TDateTime;
begin

  DateMouv := StrToDate(TOBL.GetValue('DATEMVT'));

  TobPiece.PutValue ('ETABLISSEMENT',TOBL.GetString('Etablissement'));
  TobPiece.PutValue ('DOMAINE',      TOBL.GetString('Domaine'));
  TobPiece.PutValue ('DATEPIECE',    DateMouv);
  TobPiece.PutValue ('CREEPAR',      'TPI');
  TobPiece.PutValue ('ORIGINE',      'CODEBARRE');
  TobPiece.PutValue ('REFEXTERNE',   '');
  TobPiece.PutValue ('CLEDOC',       '');
  TobPiece.PutValue ('TYPEMVT',      TOBL.GetString('TYPEMVT'));

  TobPiece.PutValue ('NATUREPIECEG', Nature);

  IF (TOBL.GetString('TYPEMVT') = 'SA') OR (TOBL.GetString('TYPEMVT') = 'EA') OR (TOBL.GetString('TYPEMVT') = 'BT') THEN
  begin
    TobPiece.PutValue ('AFFAIRE',    TOBL.GetString('CODEAFFAIRE'));
    TobPiece.PutValue ('TIERS',      TOBL.GetString('CLIENT'));
    TobPiece.PutValue ('REFINTERNE', TOBL.GetString('REFINTERNE'));
    TobPiece.PutValue ('REFEXTERNE', TOBL.GetString('REFEXTERNE'));
  end
  else IF (TOBL.GetString('TYPEMVT') = 'CD') then
  begin
    TobPiece.PutValue ('AFFAIRE',    TOBL.GetString('CODEAFFAIRE'));
    TobPiece.PutValue ('TIERS',      TOBL.GetString('CODEFOURNISSEUR'));
    TobPiece.PutValue ('REFINTERNE', TOBL.GetString('REFINTERNE'));
    TobPiece.PutValue ('REFEXTERNE', TOBL.GetString('REFEXTERNE'));
  end
  else IF (TOBL.GetString('TYPEMVT') = 'BL') then
  Begin
    TobPiece.PutValue ('AFFAIRE',    TOBL.GetString('CODEAFFAIRE'));
    TobPiece.PutValue ('TIERS',      TOBL.GetString('CODEFOURNISSEUR'));
    TobPiece.PutValue ('REFINTERNE', TOBL.GetString('REFINTERNE'));
    TobPiece.PutValue ('REFEXTERNE', 'Bon Livraison N°' + TOBL.GetString('NUMBL'));
  end
  else
  begin
    TobPiece.PutValue ('AFFAIRE',    '');
    TobPiece.PutValue ('TIERS',      '');
  end;

end;

Procedure TOF_INTEGREFICTPI.CreateTOBPiece;
begin

  TobPiece.AddChampSupValeur ('NATUREPIECEG', '');
  TobPiece.AddChampSupValeur ('AFFAIRE',      '');
  TobPiece.AddChampSupValeur ('TIERS',        '');
  TobPiece.AddChampSupValeur ('ETABLISSEMENT','');
  TobPiece.AddChampSupValeur ('DOMAINE',      '');
  TobPiece.AddChampSupValeur ('DATEPIECE',    '');
  TobPiece.AddChampSupValeur ('REFINTERNE',   '');
  TobPiece.AddChampSupValeur ('REFEXTERNE',   '');
  TobPiece.AddChampSupValeur ('CREEPAR',      'TPI');

  TobPiece.AddChampSupValeur ('ORIGINE',      'CODEBARRE');
  TobPiece.AddChampSupValeur ('TYPEMVT',      '');
  TobPiece.AddChampSupValeur ('TYPEMVT',      '');

  TobPiece.AddChampSupValeur ('CLEDOC',       '');

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 26/03/2015
Modifié le ... : 26/03/2015
Description .. : Gestion des rejets et des validations
Mots clefs ... : ETAT;REJET;VALIDATION
*****************************************************************}

Function TOF_INTEGREFICTPI.GestionRejet(TOBL : TOB) : Boolean;
Var Motif       : String;
    TypeMvt     : string;
    TypeErr     : string;
begin

  Result := True;
  Motif  := '';
  TypeErr := '';
  TypeMvt := Tobl.GetString('TYPEMVT');

  If TypeMvt = '' then
  begin
    TypeErr := 'GEN';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif   := Motif + ' Type de Mouvement inconnu : + ' + TypeMvt;
  end;

  If TOBL.GetString('ANOMALIE') = 'DEPOT' then
  begin
    TypeErr := 'DEP';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + ' Code Dépôt inexistant';
  end;

  If TOBL.GetString('CODEARTICLE') = '' then
  begin
    TypeErr := 'ART';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + ' Code Article non renseigné';
  end;

  //If (TypeMvt = 'CD') AND (TOBL.GetString('CODEFOURNISSEUR') = '') then //commande Fournisseur
  //begin
  //  TypeErr := 'FRS';
  //  if Motif = '' then Motif := 'Erreur :';
  //  Motif := Motif + ' Code fournisseur à blanc sur Commande fournisseur';
  //end;

  If ((TypeMvt = 'SA') OR (TypeMvt = 'EA')) AND (TOBL.GetString('CODEAFFAIRE') = '') then
  begin
    TypeErr := 'AFF';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + ' Le Code Affaire n''est pas renseigné';
  end;

  If (TOBL.GetString('LIB_ARTICLE') = '...') then
  begin
    TypeErr := 'ART';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + ' L''Article est inexistant';
  end;

  If (TOBL.GetString('ANOMALIECLI') = 'X') then
  begin
    TypeErr := 'CLI';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + TOBL.GetString('LIBANOMALIECLI');
  end;

  If  (TOBL.GetString('ANOMALIEMAR') = 'X') Or
      (TOBL.GetString('ANOMALIEMAR') = 'F') Or
      (TOBL.GetString('ANOMALIEMAR') = 'S') then
  begin
    TypeErr := 'ART';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + TOBL.GetString('LIBANOMALIEMAR');
  end;

  If (TOBL.GetString('ANOMALIEFRS') = 'X') then
  begin
    TypeErr := 'FRS';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + TOBL.GetString('LIBANOMALIEFRS');
  end;

  If (TOBL.GetString('ANOMALIEAFF') = 'X') then
  begin
    TypeErr := 'AFF';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + TOBL.GetString('LIBANOMALIEAFF');
  end;

  If (TOBL.GetString('ANOMALIEBC') = 'X') then
  begin
    TypeErr := 'BC';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + TOBL.GetString('LIBANOMALIEBC');
  end;

  if TOBL.GetString('GA_FERME') = 'X' then
  begin
    TypeErr := 'ART';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + ' Article ' + TOBL.GetString('CODEARTICLE') + ' Fermé';
  end;

  if TOBL.GetString('AFF_FERME') = 'X' then
  begin
    TypeErr := 'AFF';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + ' L''affaire ' + TOBL.GetString('CODEAFFAIRE')  + ' est fermée';
  end;

  if TOBL.GetString('FRS_FERME') = 'X' then
  begin
    TypeErr := 'FRS';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + ' Le fournisseur ' + TOBL.GetString('CODEFOURNISSEUR') + ' est fermé';
  end;

  if (TOBL.GetString('TENUESTOCK') = '-') then
  begin
    TypeErr := 'ART';
    if Motif = '' then Motif := 'Erreur :' Else Motif := Motif + ' / ';
    Motif := Motif + ' Article ' +  TOBL.GetString('CODEARTICLE') + ' non tenu en Stock';
  end;

  if Motif <> '' then
  begin
    MajTobRejet(TOBL, TypeErr, Motif);
    Result := False;
  end;

end;

Procedure TOF_INTEGREFICTPI.MajTobRejet(TOBL : TOB;TypeErr, Motif : String);
Var  TobLRejete  : TOB;
begin

    TObLRejete := TOB.Create('REJET', TOBRejete,-1);
    TObLRejete.Dupliquer(TOBL, True, true);
    TobLRejete.PutValue('TYPEERREUR', TypeErr);
    TobLRejete.AddChampSupValeur('MOTIF', Motif);

end;

Procedure TOF_INTEGREFICTPI.GestionValide(TOBL : TOB);
Var TobLValide  : TOB;
begin

  TObLValide := TOB.Create('VALIDE', TobValide,-1);
  TObLValide.Dupliquer(TOBL, True, true);

end;


procedure TOF_INTEGREFICTPI.ChargeEtatRejete;
begin
  TheType   := 'E';
  TheNature := 'BCB';
  TheModele := 'BCR';
  TheTitre  := 'Edition des lignes Rejetées';

  OptionEdition := TOptionEdition.Create(TheType, TheNature, TheModele, TheTitre, '', ChkApercu.Checked, ChkReduire.Checked, True, False, False, Pages, fEtat);

  OptionEdition.first := True;
  OptionEdition.ChargeListeEtat(fEtat,Idef);

  if fetat.itemindex  >= 0 then TheModele := FETAT.values[fetat.itemindex];

end;

procedure TOF_INTEGREFICTPI.ChargeEtatValide;
begin

  TheType1   := 'E';
  TheNature1 := 'BCB';
  TheModele1 := 'BCV';
  TheTitre1  := 'Edition des Lignes Validées';

  OptionEdition1 := TOptionEdition.Create(TheType1, TheNature1, TheModele1, TheTitre1, '', ChkApercu.Checked, ChkReduire.Checked, True, False, False, Pages, FETAT1);

  OptionEdition1.first := True;
  OptionEdition1.ChargeListeEtat(fEtat1,Idef);

  if fetat1.itemindex  >= 0 then TheModele1 := FETAT1.values[fetat1.itemindex];

end;

procedure TOF_INTEGREFICTPI.ParamEtat(Sender : TOBJect);
begin

  OptionEdition.Appel_Generateur

end;
procedure TOF_INTEGREFICTPI.ParamEtat1(Sender : TOBJect);
begin

  OptionEdition1.Appel_Generateur

end;

procedure TOF_INTEGREFICTPI.OnClickApercu(Sender : Tobject);
begin
  OptionEdition.Apercu  := ChkApercu.checked;
  OptionEdition1.Apercu := ChkApercu.checked;
end;

procedure TOF_INTEGREFICTPI.OnClickReduire(Sender : Tobject);
begin
  OptionEdition.DeuxPages  := ChkReduire.checked;
  OptionEdition1.DeuxPages := ChkReduire.checked;
end;

Procedure TOF_INTEGREFICTPI.OnChangeEtat(Sender : Tobject);
Begin

   OptionEdition.Modele := FETAT.values[fetat.itemindex];

   TheModele := OptionEdition.Modele;

end;

Procedure TOF_INTEGREFICTPI.OnChangeEtat1(Sender : Tobject);
Begin

   OptionEdition1.Modele := FETAT1.values[fetat1.itemindex];

   TheModele1 := OptionEdition1.Modele;

end;

procedure TOF_INTEGREFICTPI.EditeRejete;
begin

  if TobRejete.detail.count = 0 then exit;

  if TobValide.detail.count <> 0 then StartPdfBatch;

  OptionEdition.Apercu    := ChkApercu.Checked;
  OptionEdition.DeuxPages := ChkReduire.Checked;
  OptionEdition.Spages    := Pages;

  if OptionEdition.LanceImpression('', TobRejete) < 0 then V_PGI.IoError:=oeUnknown;

end;
procedure TOF_INTEGREFICTPI.EditeValide;
begin

  if TobValide.detail.count = 0 then exit;

  OptionEdition1.Apercu    := ChkApercu.Checked;
  OptionEdition1.DeuxPages := ChkReduire.Checked;
  OptionEdition1.Spages    := Pages;

  if OptionEdition1.LanceImpression('', TobValide) < 0 then V_PGI.IoError:=oeUnknown;

  if TobRejete.detail.count <> 0 then
  begin
    CancelPDFBatch ;
    PreviewPDFFile('',GetMultiPDFPath,True);
  end;

end;


function TOF_INTEGREFICTPI.ArchivagefichierTPI : Boolean;
Var FileName : string;
    FileNameH : String;
begin

  Result := False;

  FileName := FichierImport.text;

  if ControleFichierArchi then
  begin
    FileNameH := FichierHisto;
    if CopyFile(PChar(FileName), PChar(FileNameH), False) then Result := True; //on copie le fichier dans le répertoire des archives
  end;

end;

Procedure TOF_INTEGREFICTPI.SoustractionQteInv(CodeArticle : String; QteArt : Double;NoLig : Integer);
Var TobLinvArt  : TOB;
    TOBLInv     : TOB;
    QteCalc     : Double;
begin

  //Modification de la ligne d'inventaire pour l'ancien article/dépôt
  TobLInvArt   := TobInventaire.FindFirst(['CODEARTICLE','CODEDEPOT'],[CodeArticle,Depot.text],True);

  if TOBLinvArt <> nil then
  begin
    QteCalc    := TOBLinvArt.GetDouble('QTECALC');
    QteCalc    := QteCalc - QteArt;
    TOBLInvArt.PutValue('QTECALC', FloatToStr(QteCalc));
    //si la quantité Calculée est égale à zéro on supprime l'enregistrement
    If QteCalc = 0 then
      FreeAndNil(TOBLinvArt)
    else
    begin
      //Recherche par dépot
      TOBLInv   := TobLinvArt.Findfirst(['GIL_CODEARTICLE','GIL_DEPOT','NOLIG'], [CodeArticle,Depot.text,NoLig], False);
      If TOBLinv <> nil then
      begin
        QteCalc := TOBLinvArt.GetDouble('QTECALC');
        QteCalc := QteCalc - QteArt;
        TOBLInv.PutValue('QTECALC', QteCalc);
      end;
    end;
  end;

end;

Procedure TOF_INTEGREFICTPI.AdditionQteInventaire(CodeArticle : String; QteArt : Double; Nolig : Integer);
Var TobLinvArt  : TOB;
    TOBLInv     : TOB;
    QteCalc     : Double;
begin

  //Modification de la ligne d'inventaire pour l'ancien article/dépôt
  TobLInvArt  := TobInventaire.FindFirst(['CODEARTICLE'],[CodeArticle],False);

  //Ligne inventaire générale par Article
  if TOBLinvArt <> nil then
  begin
    QteCalc := TOBLinvArt.GetDouble('QTECALC');
    QteCalc := QteCalc + QteArt;
    TOBLInvArt.PutValue('QTECALC', QteCalc);
    //Recherche par dépôt
    TOBLInv   := TobLinvArt.FindFirst(['GIL_CODEARTICLE','GIL_DEPOT','NOLIG'], [CodeArticle,Depot.text,Nolig], True);
    If TOBLinv <> nil then
    begin
      QteCalc := TOBLinv.GetDouble('QTECALC');
      QteCalc := QteCalc + QteArt;
      TOBLInv.PutValue('QTECALC', QteCalc);
    end;
  end;

end;

procedure TOF_INTEGREFICTPI.GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
var TOBL: TOB;
begin

  if ACol < Grille.FixedCols then Exit;
  if Arow < Grille.Fixedrows then Exit;

  if ((ARow <= 0) or (ARow > TOBGrille.Detail.Count)) then Exit;

  TOBL := TOBGrille.Detail[ARow - 1];

  if TOBL = nil then Exit;

  {*
  if (TOBL.GetValue('ANOMALIE')    <> '') OR
     (TOBL.GetValue('ANOMALIEAFF') <> '') OR
     (TOBL.GetValue('ANOMALIEFRS') <> '') OR
     (TOBL.GetValue('ANOMALIEMAR') <> '') OR
     (TOBL.GetValue('ANOMALIEBC')  <> '') OR
     (TOBL.GetValue('ANOMALIECLI') <> '') then
  *}

  if TOBL.GetString('ERREUR') = 'X' then
  begin
  	Canvas.Font.Color := clRed;
  end
  else
  begin
    Canvas.Font.Color := clWindowText;
  end;

end;


Procedure  TOF_INTEGREFICTPI.ChargementStockArticle(TobTempo : TOB);
Var Ind     : Integer;
    TOBL    : TOB;
    TOBLSTK : TOB;
    Qte     : Double;
    QteE    : Double;
    QteS    : Double;
    TypeMvt : string;
begin

  for Ind := 0 to TobTempo.Detail.count -1 do
  begin
    //
    Qte   := 0;
    QteE  := 0;
    QteS  := 0;
    //
    TOBL    := TobTempo.detail[Ind];
    TypeMvt := TOBL.GetValue('TYPEMVT');
    //On ne traitera que les pièce de type sortie
    if (TypeMvt <> 'SA') AND (TypeMvt <> 'SD') then continue;
    TOBLSTK := TobStock.FindFirst(['CODEARTICLE','CODEDEPOT'],[TOBL.GetValue('GA_ARTICLE'), TOBL.GetValue('DEPOT')], False);
    if TOBLSTK = nil then
    begin
      StSQL := 'SELECT GQ_ARTICLE AS CODEARTICLE, GQ_DEPOT AS CODEDEPOT, GQ_PHYSIQUE AS QTEDISPO, 0 AS QTEENTREE, 0 AS QTESORTIE , GQ_PHYSIQUE AS NEWSTOCK ';
      StSQL := StSQL + ' FROM DISPO ';
      StSQL := StSQL + 'WHERE GQ_ARTICLE="' + TOBL.GetValue('GA_ARTICLE') + '"';
      StSQL := StSQL + '  AND GQ_DEPOT="'   + TOBL.GetValue('DEPOT') + '"';
      QQ := OpenSQL(StSQL, False,-1, '', true);
      If Not QQ.EOF then
      begin
        TOBLSTK := Tob.Create('DISPOSTK', TobStock, -1);
        TOBLSTK.SelectDB('STKDISPO',QQ, True);
      end
      else
      begin
        TOBLSTK := Tob.Create('DISPOSTK', TobStock, -1);
        TOBLSTK.AddchampSupValeur('CODEARTICLE', TOBL.GetValue('CODEARTICLE'));
        TOBLSTK.AddchampSupValeur('CODEDEPOT',   TOBL.GetValue('DEPOT'));
        TOBLSTK.AddchampSupValeur('QTEDISPO',    0);
        TOBLSTK.AddchampSupValeur('QTEENTREE',   0);
        TOBLSTK.AddchampSupValeur('QTESORTIE',   0);
        TOBLSTK.AddchampSupValeur('NEWSTOCK',    0);
      end;
      Ferme(QQ);
    end;
    if TOBL.GetString('INTEGRE') = 'X' then
    begin
      Qte := TOBLSTK.GetValue('NEWSTOCK');
      If  Pos(TypeMVT, 'EA;ED;CD') > 0 then QteE := TOBL.GetValue('QTE');
      If  Pos(TypeMVT, 'SA;SD')    > 0 then QteS := TOBL.GetValue('QTE');
      //Calcul du Nouveau Stock Physique
      Qte := (Qte + QteE) - QteS;
      //
      TOBLSTK.PutValue('QTEENTREE',  TOBLSTK.GetValue('QTEENTREE') + QteE);
      TOBLSTK.PutValue('QTESORTIE',  TOBLSTK.GetValue('QTESORTIE') + QteS);
      TOBLSTK.PutValue('NEWSTOCK', Qte);
    end;
  end;

end;

Function TOF_INTEGREFICTPI.ControleStockAvantIntegration(TobTempo : TOB) : Boolean;
Var Indice  : Integer;
    TOBLSTK : TOB;
    TOBL    : TOB;
    TOBLTmp : TOB;
    TmpArt  : String;
    TmpDepot: String;
    NewStock: Double;
begin

  Result := False;

  //il faut partir de la TobGrille...
  For indice := 0 to TobTempo.detail.count -1 do
  begin
    TOBL      := TobTempo.Detail[Indice];
    TmpArt    := TobL.GetValue('CODEARTICLE');
    TmpDepot  := TobL.GetValue('DEPOT');
    //On ne prend que les mouvement de type sortie
    if (TOBL.GetValue('TYPEMVT') <> 'SA') AND (TOBL.GetValue('TYPEMVT') <> 'SD') then continue;
    //on ne contrôle pas les lignes que l'on ne désire pas intégrer
    if  TOBL.GetVALUE('INTEGRE') =  '-'   Then Continue;
    //On recherche dans la table des stock la ligne correspondante à l'article
    TOBLSTK := TOBStock.Findfirst(['CODEARTICLE', 'CODEDEPOT'],[TOBL.GetValue('GA_ARTICLE'),TOBL.GetValue('DEPOT')], False);
    if TOBLSTK <> nil then
    begin
      NewStock := TOBLSTK.GetValue('NEWSTOCK');
      If  NewStock < 0 then
      begin
        if PGIAsk('Attention : l''article ' + TmpArt + ' aura une quantité négative en stock sur le dépôt ' + TmpDepot + ' après intégration du fichier' + CHR(13) + ' Voulez-vous l''intégrer quand même ?')=MrNo then
        begin
          Result := True;
          //On boucle pour rendre l'ensemble des lignes de l'article non intégrable...
          TOBLTmp := TobTempo.FindFirst(['GA_ARTICLE', 'DEPOT'],[TOBL.GetValue('GA_ARTICLE'),TOBL.GetValue('DEPOT')], False);
          while TOBLTmp <> nil do
          begin
            TOBLTmp.PutVALUE('INTEGRE', '-');
            TOBLTmp.PutValue('ANOMALIE', 'STOCK');
            GestionAnomalie(TOBLTmp);
            Grille.Row := TOBLTMP.GetValue('NOLIG');
            Grille.InvalidateRow(Grille.Row);
            TOBLTMP := TobTempo.FindNext(['GA_ARTICLE', 'DEPOT'],[TOBL.GetValue('GA_ARTICLE'),TOBL.GetValue('DEPOT')], False);
          end;
        end;
      end;
    end;  
  end;

  if result then AffichageGrilleAnomalie;
  
end;

procedure TOF_INTEGREFICTPI.BtSelectLigneOnClick(sender: TObject);
begin

  if TypeMVT.Value <> 'SA'  then Exit;

  If NumBesoin.text = ''    then Exit;

  //Lecture de la pièce associé au besoin de chantier
  StSQL := 'SELECT GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, GL_INDICEG, GL_NUMORDRE, GL_DATEPIECE, GL_CODEARTICLE, GL_ARTICLE, GL_LIBELLE, GL_QTEFACT, GL_QTERESTE, GL_PUHTDEV FROM LIGNE ';
  StSQL := StSQL + 'WHERE GL_NATUREPIECEG="CBT" ';
  StSQL := StSQL + '  AND GL_NUMERO=' + NumBesoin.text;
  StSQL := StSQL + '  AND GL_INDICEG=0';
  If Depot.Text   <> '' then  StSQL := StSQL + ' AND GL_DEPOT="' + Depot.Text + '"';
  if CodeArticle.Text <> '' then  StSQL := StSQL + ' AND GL_CODEARTICLE="' + CodeArticle.Text + '"';
  StSQL := StSQL + '  AND GL_QTERESTE > 0';
  //
  QQ    := OpenSQL(StSQL, False,1,'',True);

  If Not QQ.Eof then
  begin
    TobBesoin := TOB.create('BESOIN',nil, -1);
    TobBesoin.LoadDetailDBFromSQL('Lignes Besoin', StSQL);

    //Création de la ttoolWindows qui recevra la grille
    BesoinWin         := TToolWindow97.create (application);
    BesoinWin.Parent  := Ecran;
    BesoinWin.OnClose := CloseBesoinWin;
    //
    Grille3           := THGrid.Create(BesoinWin);
    Grille3.Parent    := BesoinWin;
    Grille3.Align     := alClient;
    Grille3.Visible   := True;
    Grille3.Options   := Grille3.Options + [GoRowSelect];

    Grille3.OnDblClick:= SelectLigneBesoin;
    //
    //chargement de la grille avec les éléments de la TOB
    GestionG3.fColNamesGS   := fColNamesGrille3;
    GestionG3.FalignementGS := FalignementGrille3;
    GestionG3.FtitreGS      := FtitreGrille3;
    GestionG3.fLargeurGS    := fLargeurGrille3;
    GestionG3.TOBG          := TOBBesoin;
    GestionG3.GS            := Grille3;
    //
    BesoinWin.ClientWidth  := Round(Ecran.Width/2);
    BesoinWin.ClientHeight := Round(Ecran.Height/2);
    BesoinWin.Top          := Round(Ecran.Top + (Ecran.Height / 2));
    BesoinWin.Left         := Round(Ecran.Left+ (Ecran.Width  / 2));
    //
    BesoinWin.Caption      := TraduireMemoire('Ensemble des lignes Article ' + CodeArticle.Text + ' sur Besoin de Chantier N°' + NumBesoin.Text);
    //
    BesoinWin.visible := True;
    //
    DessineGrille(GestionG3);
    //
    ChargementGrille(GestionG3);
    //
    Grille3.row := 1;
    //
  end;

  Ferme(QQ);

end;

procedure TOF_INTEGREFICTPI.SelectLigneBesoin(Sender : Tobject);
Var Cledoc  : R_Cledoc;
    TOBL    : TOB;
    Std     : string;
begin

  With GestionG3 do
  begin
    TOBL := TobGrille.Detail[Grille.row-1];
    //
    Cledoc.NaturePiece  := TOBG.detail[GS.Row -1].GetValue('GL_NATUREPIECEG');
    Cledoc.DatePiece    := TOBG.detail[GS.Row -1].GetValue('GL_DATEPIECE');
    StD := FormatDateTime('ddmmyyyy', cledoc.DatePiece);

    Cledoc.Souche       := TOBG.detail[GS.Row -1].GetValue('GL_SOUCHE');
    Cledoc.NumeroPiece  := TOBG.detail[GS.Row -1].GetValue('GL_NUMERO');
    Cledoc.Indice       := TOBG.detail[GS.Row -1].GetValue('GL_INDICEG');
    Cledoc.NumOrdre     := TOBG.detail[GS.Row -1].GetValue('GL_NUMORDRE');
    //
    TOBL.PutValue('PIECEPRECEDENTE',  StD + ';' + cledoc.NaturePiece  + ';' + Cledoc.Souche  + ';' + IntToStr(cledoc.NumeroPiece) + ';' + IntToStr(cledoc.Indice) + ';' + IntToStr(cledoc.NumOrdre) + ';');
    TOBL.PutValue('NATPIECEPREC',     Cledoc.NaturePiece);
    TOBL.PutValue('DATEBESOIN',       Cledoc.DatePiece);
    //
    TOBL.PutValue('ANOMALIEBC',  '');
    TOBL.PutValue('LIBANOMALIEBC',  '');
    if TobL.GetValue('MORELINEBC') = 'X' then
      BtSelectLigne.Visible := true
    else
      BtSelectLigne.Visible := False;
  end;
  //
  GestionAnomalie(TOBL);
  AffichageGrilleAnomalie;
  //
  BesoinWin.Visible := False;
  //
  FreeAndNil(BesoinWin);

end;

procedure TOF_INTEGREFICTPI.CloseBesoinWin(Sender : Tobject);
begin
  //
  BesoinWin.Visible := False;
  //
  FreeAndNil(BesoinWin);
  //
end;

procedure DessineGrille(GestionGs : TGestionGS);
var st              : String;
    lestitres       : String;
    lesalignements  : String;
    leslargeurs     : String;
    FF              : String;
    alignement      : String;
    Nam             : String;
    lalargeur       : Integer;
    letitre         : String;
    lelement        : string;
    //
    Obli            : Boolean;
    OkLib           : Boolean;
    OkVisu          : Boolean;
    OkNulle         : Boolean;
    OkCumul         : Boolean;
    Sep             : Boolean;
    //Okimg           : boolean;
    //
    dec             : Integer;
    NbCols          : integer;
    indice          : Integer;
    //
    TailleCar       : Integer;
begin
  //
  //Calcul du nombre de Colonnes du Tableau en fonction des noms
  st := GestionGS.fColNamesGS;

  NbCols := 0;

  repeat
    lelement := READTOKENST (st);
    if lelement <> '' then
    begin
      inc(NbCols);
    end;
  until lelement = '';
  //
  GestionGS.GS.ColCount := Nbcols;
  //
  st              := GestionGS.fColNamesGS;
  lesalignements  := GestionGS.FalignementGS;
  lestitres       := GestionGS.FtitreGS;
  leslargeurs     := GestionGS.FlargeurGS;
  TailleCar       := GestionGS.GS.Canvas.TextWidth('W');

  //Mise en forme des colonnes
  for indice := 0 to Nbcols -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    //
    alignement  := ReadTokenSt(lesalignements);
    lalargeur   := StrToInt(readtokenst(leslargeurs));
    letitre     := readtokenst(lestitres);
    //
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    //
    GestionGS.GS.cells[Indice,0]     := leTitre;
    GestionGS.GS.ColNames [Indice]   := Nam;
    //
    //Alignement des cellules
    if copy(Alignement,1,1)='G'       then //Cadré à Gauche
      GestionGS.GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D'  then  //Cadré à Droite
      GestionGS.GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C'  then
      GestionGS.GS.ColAligns[indice] := taCenter; //Cadré au centre

    If Nam = 'TYPEMVT'    then GestionGS.GS.ColFormats[indice]:='CB=BTTYPEMVT'
    Else If Nam = 'DEPOT' then GestionGS.GS.ColFormats[indice]:='CB=GCDEPOT';

    //Colonne visible ou non
    if OkVisu then
  		GestionGS.GS.ColWidths[indice] := lalargeur * TailleCar
    else
    	GestionGS.GS.ColWidths[indice] := -1;

    //Affichage d'une image ou du texte
    {*
    okImg := (copy(Alignement,8,1)='X');
    if (OkLib) or (okImg) then
    begin
    	Grille.ColFormats[indice] := 'CB=' + Get_Join(Nam);
      if OkImg then
      begin
      	Grille.ColDrawingModes[Indice]:= 'IMAGE';
      end;
    end;
    *}

    if (Dec<>0) or (Sep) then
    begin
    	if OkNulle then
        GestionGS.GS.ColFormats[indice] := FF+';; ;' //'#'
      else
      	GestionGS.GS.ColFormats[indice] := FF; //'#';
    end;
  end ;

end;

Procedure ChargementGrille(GestionGS : TGestionGS);
var indice : integer;
begin

  With GestionGs do
  begin
    if TOBG.Detail.count <> 0 then
      GS.RowCount := TOBG.detail.count + 1
    else
      Exit;
    //
    GS.DoubleBuffered := true;
    GS.BeginUpdate;
    //
    TRY
      GS.SynEnabled := false;
      //for Indice := 0 to TOBG.detail.count -1 do
      //begin
      //  GS.row := Indice+1;
        TOBG.PutGridDetail(GS,False,True,fColNamesGS);
      //end;
    FINALLY
      GS.SynEnabled := true;
      GS.EndUpdate;
    END;
    TFVierge(Tform).HMTrad.ResizeGridColumns(GS);
  end;

end;

Initialization
  registerclasses ( [ TOF_INTEGREFICTPI ] ) ;
end.

