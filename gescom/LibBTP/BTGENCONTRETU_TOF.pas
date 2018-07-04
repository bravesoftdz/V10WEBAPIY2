{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/02/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTGENCONTRETU ()
Mots clefs ... : TOF;BTGENCONTRETU
*****************************************************************}
Unit BTGENCONTRETU_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
     M3FP,
{$ELSE}
     MainEagl,
{$ENDIF}
		 menus,
		 graphics,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     Hpanel,
     HEnt1,
     HMsgBox,
     Vierge,
     ExtCtrls,
//     ExComCtrls,
		 AglInit,
     UTOB,
     UTOF,
     UtilGenContreEtude,
     BTStructChampSup ;

CONST MAXPHASE = 9;

Type
  TOF_BTGENCONTRETU = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	ChoixPopupPar : TNodeAttachMode;
  	Validate : boolean;
  	TreeDevis,TreeCtrEtude : TTreeView;
    MnAddPar,MnDelPar : TMenuItem;
    PopupPar : TpopupMenu;
    TOBDevis,TOBOuvrage,TOBIntermediaire,TOBNewParag,TOBCtrEtude,TOBLienInterDoc,TOBOUvrageResult : TOB;
    PN1,PN2,PN3,PN4,PN5,PN6,PN7,PN8,PN9 : integer;  // paragraphe du document
    Level : integer;
    RootDevis,RootCtrEtude : TTreeNode;
    TheListImage : TImageList;
    ListParag : Tlist;
    TheRecepteurPop : TTreeNode;

	  ParamEmetteur,ParamDestinataire : TTreeNode;

  	POPCHOIXDEPPAR : TPopupMenu;
    POPCTREDTUDE : TPopupMenu;


    function GenereTobIntermediaire1: boolean;
    procedure GetControls;
    procedure GetData;
    procedure AddChampInterm(TOBD: TOB);
    function InsereOuvrageInterm(TOBO, TOBL, TOBPere: TOB): TOB;
    function InserePhaseInterm(TOBDevis, TOBL, TOBPere: TOB): TOB;
    function FindDetailOuvrage(TOBL: TOB): TOB;
    procedure InitValues;
    procedure PositionneLevel (TOBPere: TOB);
    procedure IncrementeLevel  ;
    //
    function EncodePhase(TOBS: TOB) : string;
    procedure DecodePhase(Phase: string);
    //
    function EncodeClef(TOBPere : TOB) : string;
    //
    procedure TraiteDetailOuvrage(TOBL, TOBOUV, TOBPere: TOB);
    procedure ChargeComponents;
    function InsereArticleInterm(TOBO, TOBL, TOBpere: TOB): TOB;
    procedure chargementTreeDevis(TreeDevis: TTreeView; Pere: TTreeNode; TobIntermediaire: TOB);
    function DefiniLibelle(TOBL: TOB): string;
    procedure AffecteImage(NodeSuite: TTreeNode; TOBSuite: TOB);
    procedure ChargeListImage;
    procedure SetMethode;
  	procedure TreeDevisDragDrop(Sender, Source: TObject; X, Y: Integer);
		procedure TreeDevisDragOver (Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
		procedure TreeCtrEtudeDragDrop(Sender, Source: TObject; X, Y: Integer);
		procedure TreeCtrEtudeDragOver (Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    function VerifAutorisationDeplacement(Source: TObject): boolean;
    function verifElement(TheSelectedItem: TTreeNode): boolean;
    procedure DevisVerCtrEtude(Emetteur, Destinataire: TTreeNode);
    function AjouteLigneCtrEtude(Emetteur,Destinataire: TTreeNode): TTreeNode;
    procedure SupprimeEmetteurDevis(Element: TTreeNode);
//    procedure AfficheCoordonnees(X, Y: integer; TheRecepteur: TTreeNode);
    procedure InterCtrEtude(Emetteur, Destinataire: TTreeNode; X,Y : integer);
    function EncodeOuvrage(TOBO: TOB): string;
    function VerifNode(NodePere, Emetteur: TTreeNode): TTreeNode;
    procedure InsereTreeDevis(NodePere: TTreenode; TheTOB: TOB);
    function FindNode(NodePere: TTreeNode; TheClef: string): TTreeNode;
    function FindLastParag: TOB;
    function TrouvePositionTreeView(NodePere: TTreeNode;
      UniqueId: string): TTreeNode;
    procedure PopCtrEtudeBeforePopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
  	procedure MNAddParagrapheClick (sender : TObject);
  	procedure MNDeleteParagrapheClick (sender : TObject);
    procedure CtrEtudeBeforeEditing(Sender: TObject; Node: TTreeNode;
      var AllowEdit: Boolean);
    function DevisToutTransfere : boolean;
    function NoeudNonVide(Depart: TTreeNode): boolean;
    procedure ValideLaCtrEtude;
    procedure AddChampsSupOuvrages(TOBO,TOBCtrEtude : TOB);
    procedure AjouteFilleCtrEtude(Emetteur,Destinataire: TTreeNode);
    function GetQteAplat (Emetteur : TTreeNode) : double;
    function DetailOuvrageDejaPresent (TOBGeniteur : TOB; Source : TObject ; IDSel : string) : boolean;
    function TrouveGeniteur(IDSel: string): TOB;
    function VerifExistID(TOBControle: TOb; Controle: TTreeNode): boolean;
    procedure TraiteChoixParagraphe(Emetteur, Destinataire: TTreeNode; X,Y : integer);
    procedure DeplacePar(Sender: TObject);
    procedure DeplaceParB(Sender: TObject);
    procedure InsertParDansPar(Sender: TObject);
    procedure AddparagFin(Sender: TObject);
  end ;

function SetContreEtude (TOBDEVIS,TobOuvrage : TOB) : boolean;

Implementation

uses  Factvariante,
      FactTOB,
      FactComm,
      FactUtil,
      FactGrp,
      FactPiece,
      UentCommun,
      UtilTOBPiece;

function SetContreEtude (TOBDEVIS,TobOuvrage : TOB) : boolean;
var st : string;
begin
result := true;
TheTob := TOBDevis;
TheTob.Data := TOBOuvrage;
st := AGLLanceFiche('BTP', 'BTGENCONTRETU', '', '', 'ACTION=MODIFICATION');
TheTob := nil;
end;


procedure TOF_BTGENCONTRETU.AddChampInterm (TOBD : TOB);
begin
  TOBD.addChampsupValeur ('PROVENANCE','DEVIS',false);
  TOBD.addChampsupValeur ('PHASE','',false);
  TOBD.addChampsupValeur ('NUMLIGNE',0,false);
  TOBD.addChampsupValeur ('INDICENOMEN',0,false);
  TOBD.addChampsupValeur ('DECOMPOSITIONOUVRAGE','',false);
  TOBD.addChampsupValeur ('ARTICLE','',false);
  TOBD.addChampsupValeur ('TYPE','',false);
  TOBD.addChampsupValeur ('LIBELLE','',false);
  TOBD.addChampsupValeur ('QTE',0.0,false);
  TOBD.addChampsupValeur ('FROMDEVIS','-',false);
  TOBD.addChampsupValeur ('CLEFPERE','',false);
  TOBD.addChampsupValeur ('UNIQUEID','',false);
  TOBD.addChampSupValeur ('ADDED','-');
  TOBD.addChampsupValeur ('QTEINIT',0.0,false);
  TOBD.addChampsupValeur ('QTEDUDETAIL',1,false);
  TOBD.addChampsupValeur ('PRIXPOSEPAPA',0,false);
  TOBD.addChampsupValeur ('NATURETRAVAIL','',false);
end;

procedure TOF_BTGENCONTRETU.DecodePhase (Phase : string);
var TheDecodeStr,TheValue : string;
		Niveau : integer;
begin
  PN1 := 0; PN2 := 0; PN3 := 0;
  PN4 := 0; PN5 := 0; PN6 := 0;
  PN7 := 0; PN8 := 0; PN9 := 0;
  Niveau := 1;
	TheDecodeStr := StringReplace (Phase,'|',';',[rfReplaceAll]);
  repeat
    TheValue := READTOKENST (TheDecodeStr);
    if TheValue <> '' then
    begin
    	if Niveau = 1 then PN1 := StrTOInt (TheValue)
      else if Niveau = 2 then PN2 := StrTOInt (TheValue)
      else if Niveau = 3 then PN3 := StrTOInt (TheValue)
      else if Niveau = 4 then PN4 := StrTOInt (TheValue)
      else if Niveau = 5 then PN5 := StrTOInt (TheValue)
      else if Niveau = 6 then PN6 := StrTOInt (TheValue)
      else if Niveau = 7 then PN7 := StrTOInt (TheValue)
      else if Niveau = 8 then PN8 := StrTOInt (TheValue)
      else if Niveau = 9 then PN9 := StrTOInt (TheValue);
      inc(Niveau);
    end;
  until Thevalue = '';
  if PN1 > 0 then level := 0;
  if PN2 > 0 then level := 1;
  if PN3 > 0 then level := 2;
  if PN4 > 0 then level := 3;
  if PN5 > 0 then level := 4;
  if PN6 > 0 then level := 5;
  if PN7 > 0 then level := 6;
  if PN8 > 0 then level := 7;
  if PN9 > 0 then level := 8;
//
end;

function TOF_BTGENCONTRETU.EncodePhase(TOBS : TOB) : string;
begin
  result := Format ('%d|%d|%d|%d|%d|%d|%d|%d|%d|',[PN1,PN2,PN3,PN4,PN5,PN6,PN7,PN8,PN9]);
end;

function TOF_BTGENCONTRETU.InserePhaseInterm (TOBDevis,TOBL,TOBPere : TOB) : TOB;
var TOBS : TOB;
begin
  TOBS := TOB.Create ('LA LIGNE',TOBPere,-1);
  // TODO AJOUTER LA MISE A JOUR DES ZONES SUP
  AddChampInterm (TOBS);
  TOBS.putValue('PHASE',EncodePhase(TOBS));
  TOBS.PutValue ('NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE'));
  TOBS.PutValue ('TYPE','PAR');
  TOBS.putValue ('DECOMPOSITIONOUVRAGE',EncodeOuvrage(TOBL));
  TOBS.putValue ('LIBELLE',TOBL.GetValue('GL_LIBELLE'));
  TOBS.PutValue ('CLEFPERE',EncodeClef (TOBpere));
  TOBS.PutValue ('UNIQUEID',EncodeClef (TOBS));
  TOBS.PutValue ('NATURETRAVAIL',TOBL.GetString('GLC_NATURETRAVAIL'));
  TOBS.data := TOBL;
  result := TOBS;
end;

function TOF_BTGENCONTRETU.FindLastParag : TOB;
begin
	if ListParag.Count = 0 then BEGIN result := TOBIntermediaire; exit; END;
	result := ListParag.Items [ListParag.Count-1];
end;

function TOF_BTGENCONTRETU.EncodeOuvrage (TOBO : TOB) : string;
var N1,N2,N3,N4,N5 : integer;
begin
  result := '|||||';
	if TOBO.NomTable <> 'LIGNEOUV' then exit;
  N1 := TOBO.GetValue('BLO_N1');
  N2 := TOBO.GetValue('BLO_N2');
  N3 := TOBO.GetValue('BLO_N3');
  N4 := TOBO.GetValue('BLO_N4');
  N5 := TOBO.GetValue('BLO_N5');
  result := Format ('%d|%d|%d|%d|%d|',[N1,N2,N3,N4,N5]);
end;

function TOF_BTGENCONTRETU.InsereArticleInterm (TOBO,TOBL,TOBpere : TOB) : TOB;
var TOBS : TOB;
begin
  TOBS := TOB.Create ('LA LIGNE',TOBPere,-1);
  // TODO AJOUTER LA MISE A JOUR DES ZONES SUP
  AddChampInterm (TOBS);
  TOBS.putValue('PHASE',EncodePhase(TOBS));
  TOBS.PutValue ('NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE'));
  TOBS.PutValue ('INDICENOMEN',TOBL.GetValue('GL_INDICENOMEN'));
  if TOBO <> nil then
  begin
  	TOBS.putValue ('ARTICLE',TOBO.GetValue('BLO_CODEARTICLE'));
  	TOBS.putValue ('LIBELLE',TOBO.GetValue('BLO_LIBELLE'));
  	TOBS.putValue ('DECOMPOSITIONOUVRAGE',EncodeOuvrage(TOBO));
   	TOBS.PutValue ('TYPE',TOBO.GetValue('BLO_TYPEARTICLE'));
   	TOBS.PutValue ('QTE',TOBO.GetValue('BLO_QTEFACT'));
   	TOBS.PutValue ('QTEINIT',TOBO.GetValue('BLO_QTEFACT'));
   	TOBS.PutValue ('QTEDUDETAIL',TOBO.GetValue('BLO_QTEDUDETAIL'));
  	TOBS.PutValue ('NATURETRAVAIL',TOBO.GetString('BLO_NATURETRAVAIL'));
    TOBS.data := TOBO;
  end else
  begin
  	TOBS.putValue ('ARTICLE',TOBL.GetValue('GL_CODEARTICLE'));
  	TOBS.putValue ('LIBELLE',TOBL.GetValue('GL_LIBELLE'));
  	TOBS.putValue ('DECOMPOSITIONOUVRAGE',EncodeOuvrage(TOBL));
   	TOBS.PutValue ('QTE',TOBL.GetValue('GL_QTEFACT'));
   	TOBS.PutValue ('TYPE',TOBL.GetValue('GL_TYPEARTICLE'));
   	TOBS.PutValue ('QTEINIT',TOBL.GetValue('GL_QTEFACT'));
  	TOBS.PutValue ('NATURETRAVAIL',TOBL.GetString('GLC_NATURETRAVAIL'));
    TOBS.data := TOBL;
  end;
  TOBS.PutValue ('CLEFPERE',EncodeClef (TOBpere));
  TOBS.PutValue ('UNIQUEID',EncodeClef (TOBS));
  result := TOBS;
end;

function TOF_BTGENCONTRETU.InsereOuvrageInterm (TOBO,TOBL,TOBPere : TOB) : TOB;
var TOBS : TOB;
begin
  TOBS := TOB.Create ('LA LIGNE',TOBPere,-1);
  // TODO AJOUTER LA MISE A JOUR DES ZONES SUP
  AddChampInterm (TOBS);
  TOBS.putValue('PHASE',EncodePhase(TOBS));
  TOBS.PutValue ('NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE'));
  TOBS.PutValue ('INDICENOMEN',TOBL.GetValue('GL_INDICENOMEN'));
  if TOBO <> nil then
  begin
  	TOBS.putValue ('ARTICLE',TOBO.GetValue('BLO_CODEARTICLE'));
  	TOBS.putValue ('LIBELLE',TOBO.GetValue('BLO_LIBELLE'));
  	TOBS.putValue ('DECOMPOSITIONOUVRAGE',EncodeOuvrage(TOBO));
   	TOBS.PutValue ('TYPE',TOBO.GetValue('BLO_TYPEARTICLE'));
   	TOBS.PutValue ('QTE',TOBO.GetValue('BLO_QTEFACT'));
   	TOBS.PutValue ('QTEINIT',TOBO.GetValue('BLO_QTEFACT'));
   	TOBS.PutValue ('QTEDUDETAIL',TOBO.GetValue('BLO_QTEDUDETAIL'));
  	TOBS.PutValue ('NATURETRAVAIL',TOBO.GetString('BLO_NATURETRAVAIL'));
    if (TOBO.getValue('BLO_TYPEARTICLE')='ARP') and (TOBO.detail.count > 0) then TOBS.PutValue ('PRIXPOSEPAPA',1);
    TOBS.data := TOBO;
  end else
  begin
  	TOBS.putValue ('ARTICLE',TOBL.GetValue('GL_CODEARTICLE'));
  	TOBS.putValue ('LIBELLE',TOBL.GetValue('GL_LIBELLE'));
  	TOBS.putValue ('DECOMPOSITIONOUVRAGE',EncodeOuvrage(TOBL));
   	TOBS.PutValue ('TYPE',TOBL.GetValue('GL_TYPEARTICLE'));
   	TOBS.PutValue ('QTE',TOBL.GetValue('GL_QTEFACT'));
   	TOBS.PutValue ('QTEINIT',TOBL.GetValue('GL_QTEFACT'));
  	TOBS.PutValue ('NATURETRAVAIL',TOBL.GetString('GLC_NATURETRAVAIL'));
    if TOBL.getValue('GL_TYPEARTICLE')='ARP' then TOBS.PutValue ('PRIXPOSEPAPA',1);
    TOBS.data := TOBL;
  end;
  TOBS.PutValue ('CLEFPERE',EncodeClef (TOBpere));
  TOBS.PutValue ('UNIQUEID',EncodeClef (TOBS));
  result := TOBS;
end;

procedure TOF_BTGENCONTRETU.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTGENCONTRETU.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTGENCONTRETU.OnUpdate ;
var IoErr : TIOErr;
    OneDoc : r_cledoc;
begin
  Inherited ;
  if RootCtrEtude.Count = 0 then exit; // on sort sans rien faire
  if not DevisToutTransfere then
  begin
    PGIBox ('Il reste des élements dans le Devis',ecran.caption);
		TFVierge(Ecran).ModalResult := 0;
    exit;
  end;
  DefinitPieceCtrEtude (TOBDevis,TOBCtrEtude,TOBOuvrage,TOBOuvrageResult,TOBLienInterDoc,RootCtrEtude);
  IoErr := TRANSACTIONS (ValideLaCtrEtude,0);
  if IOerr <> OeOk then
  begin
  	PgiBox ('Erreur lors de la validation',ecran.caption);
    exit;
  end;
  OneDoc := TOB2Cledoc (TOBDEVIS);
  //
  if ExecuteSql ('UPDATE AFFAIRE SET AFF_PREPARE="X" '+
              'WHERE '+
              'AFF_AFFAIRE=('+
              'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+WherePiece(OneDoc,ttdPiece,false)+')')=0  then
  begin
  	PgiBox ('Erreur lors de la validation',ecran.caption);
    IoErr := OeUnknown;
    exit;
  end;
  //
	Validate := true;
end ;

procedure TOF_BTGENCONTRETU.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTGENCONTRETU.OnArgument (S : String ) ;
begin
  Inherited ;
  MemoriseChampsSupLigneETL ('DBT' ,True);
  MemoriseChampsSupLigneOUV  ('DBT');
  MemoriseChampsSupPIECETRAIT;
	TOBIntermediaire := TOB.Create ('INITIAL',nil,-1);
	TOBLienInterDoc := TOB.Create ('LIEN INTER DOC',nil,-1);
	TOBNewParag := TOB.Create ('__NEW',nil,-1);
  TOBCtrEtude := TOB.Create ('PIECE',nil,-1);
  TOBOuvrageResult := TOB.Create ('LES OUVRAGES SORTI',nil,-1);
  AddChampInterm (TOBintermediaire);
  GetData;
  GetControls;
  InitValues;
	ChargeListImage;
	GenereTobIntermediaire1;
	ChargeComponents;
  SetMethode;
end ;

procedure TOF_BTGENCONTRETU.InitValues;
begin
  PN1 := 0;
  PN2 := 0;
  PN3 := 0;
  PN4 := 0;
  PN5 := 0;
  PN6 := 0;
  PN7 := 0;
  PN8 := 0;
  PN9 := 0;
  Level := 0;
end;

function TOF_BTGENCONTRETU.FindDetailOuvrage (TOBL : TOB) : TOB;
begin
	result := nil;
	if TOBL.getValue('GL_INDICENOMEN') <> 0 then
  begin
  	result := TOBOuvrage.detail[TOBL.getValue('GL_INDICENOMEN')-1];
  end;
end;

function TOF_BTGENCONTRETU.GenereTobIntermediaire1: boolean;
var Indice : integer;
    TOBL,TOBI,TOBOUV : TOB;
    TOBPere : TOB;
begin
  Result := true;
  TOBPere := TOBINTERMEDIAIRE;
  for Indice := 0 to TobDevis.detail.count -1 do
  begin
    TOBL := TOBDevis.detail[Indice];
    // VARIANTE
    if IsVariante (TOBL) then continue;
    //-- si cotraitance ligne pas prise en compte...
    if (TOBL.GetString('GLC_NATURETRAVAIL')='001') then continue;
    // --
    AddLesSupLigne (TOBL,false);

    if IsDebutParagraphe(TOBL) then
    BEGIN
      IncrementeLevel;
      TOBI := InserePhaseInterm (TOBDevis,TOBL,TOBPere);
      TOBPere := TOBI;
      ListParag.Add (TOBI);
      continue;
    END;

    if IsFinParagraphe(TOBL) then
    BEGIN
    	level := 0 ; // init
			ListParag.Delete (ListParag.Count-1);
      TOBPere := FindLastParag;
      TOBI := TOBPere.detail[TOBpere.detail.count -1];
      PositionneLevel (TOBI);
      continue;
    END;

    if (TOBL.GetValue('GL_TYPEARTICLE')='OUV') OR (TOBL.GetValue('GL_TYPEARTICLE')='ARP') then
    BEGIN
      // Ajout de la ligne d'ouvrage
      TOBI := InsereOuvrageInterm (nil,TOBL,TOBPere);
      // recherche de son detail
      TOBOUV := FindDetailOuvrage (TOBL);
      if TOBOUv.detail.count > 0 then
      begin
        // traite le detail ..plein d'enfants cachés
        TraiteDetailOuvrage (TOBL,TOBOUV,TOBI);
      end;
      // pere indigne suivant ...  ;o))
      continue;
    END;

    if (TOBL.GetValue('GL_TYPEARTICLE')='MAR') OR
    (TOBL.GetValue('GL_TYPEARTICLE')='PRE') or
    (TOBL.GetValue('GL_TYPEARTICLE')='POU') or
    (TOBL.GetValue('GL_TYPEARTICLE')='FRA') then
    BEGIN
      InsereArticleInterm (nil,TOBL,TOBPere);
    END;
  end;

  //paragraphe à blanc...
  SupprimeOuvrageVide(TOBPere);

  // libere de la memoire

//  TOBOuvrage.clearDetail;
end;

procedure TOF_BTGENCONTRETU.GetControls;
begin
	PopupPar := TpopupMenu(getControl('PCHOIXPOPPAR'));
  TreeDevis := TTreeView(GetControl('TREEDEVIS'));
  TreeCtrEtude := TTreeView(GetControl('TREECTRETUDE'));
	TheListImage := TImageList.Create (TForm(ecran));
  ListParag := Tlist.Create;
  MnAddPar := TmenuItem(GetCOntrol('MnAddParag'));
  MnDelPar := TmenuItem(GetCOntrol('MnSupparag'));
  POPCHOIXDEPPAR := TPopupMenu(getControl('POPCHOIXDEPPAR'));
  POPCTREDTUDE := TPopupMenu(getControl('POPCTREDTUDE'));

end;

procedure TOF_BTGENCONTRETU.SetMethode;
begin
  TreeDevis.OnDragDrop := TreeDevisDragDrop;
  TreeDevis.OnDragOver := TreeDevisDragOver;
  TreeCtrEtude.OnDragDrop := TreeCtrEtudeDragDrop;
  TreeCtrEtude.OnDragOver := TreeCtrEtudeDragOver;
  TreeCtrEtude.OnContextPopup  := PopCtrEtudeBeforePopup;
  TreeCtrEtude.onEditing := CtrEtudeBeforeEditing;
  MnAddPar.OnClick := MNAddParagrapheClick;
  MnDelPar.OnClick := MNDeleteParagrapheClick;
  TmenuItem(GetControl('MnChoixInsPar')).OnClick := InsertParDansPar;
  TmenuItem(GetControl('MnChoixDepPar')).OnClick := DeplacePar;
  TmenuItem(GetControl('MnChoixDepParB')).OnClick := DeplaceParB;
  TmenuItem(GetControl('MnAddParagFin')).OnClick := AddParagFin;
end;

procedure TOF_BTGENCONTRETU.GetData;
begin
	TOBDevis := LaTOB;
  TOBOuvrage := TOB(LaTOB.data);
end;

procedure TOF_BTGENCONTRETU.OnClose ;
begin
  Inherited ;
  if (RootCtrEtude.Count > 0) and (not Validate) then
  begin
  	if PgiAsk ('Abandon de la génération ?',ecran.caption)<>mryes then BEGIN TFVierge(Ecran).ModalResult := 0; exit; END;
  end;
	TheListImage.free;
	TOBIntermediaire.free;
	TOBNewParag.free;
  TOBCtrEtude.free;
	TOBLienInterDoc.free;
  TOBOuvrageresult.free;
end ;

procedure TOF_BTGENCONTRETU.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTGENCONTRETU.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTGENCONTRETU.PositionneLevel (TOBpere : TOB);
begin
	DecodePhase (TOBPere.getValue('PHASE'));
end;


procedure TOF_BTGENCONTRETU.IncrementeLevel;
begin
	inc(level);
	if Level < 1 then PN1 := 0;
  if Level < 2 then PN2 := 0;
  if Level < 3 then PN3 := 0;
  if Level < 4 then PN4 := 0;
  if Level < 5 then PN5 := 0;
  if Level < 6 then PN6 := 0;
  if Level < 7 then PN7 := 0;
  if Level < 8 then PN8 := 0;
  if Level < 9 then PN9 := 0;
  Case level of
   1 : inc(PN1);
   2 : inc(PN2);
   3 : inc (PN3);
   4 : inc (PN4);
   5 : inc (PN5);
   6 : inc (PN6);
   7 : inc (PN7);
   8 : inc (PN8);
   9 : inc (PN9);
  end;
end;

function TOF_BTGENCONTRETU.EncodeClef (TOBpere : TOB) : string;
var Ligne,IndiceNomen : integer;
begin
	result := '';
  Ligne := TOBPere.getValue('NUMLIGNE');
  IndiceNomen := TOBPere.getValue('INDICENOMEN');
	if (TobPere.GetValue('PHASE') = '') and (ligne = 0) then exit;
  result := format ('%s/%d/%d/%s',[TOBPere.GetValue('PHASE'),Ligne,IndiceNOmen,
                                   TOBPere.GetValue('DECOMPOSITIONOUVRAGE')]);
end;

procedure TOF_BTGENCONTRETU.TraiteDetailOuvrage (TOBL,TOBOUV,TOBPere : TOB);
var Indice : integer;
		TOBO,TOBI : TOB;
begin
	For Indice := 0 To TOBOuv.detail.count -1 do
  begin
    //-- si cotraitance ligne pas prise en compte...
    if (TOBOuv.detail[Indice].GetString('BLO_NATURETRAVAIL')='001') then continue;
  	TOBO := TOBOuv.detail[Indice];
  	TOBI := InsereOuvrageInterm (TOBO,TOBL,TOBPere);
    if TOBO.detail.count > 0 then
    begin
      TraiteDetailOuvrage(TOBL,TOBO,TOBI);
    end else
    if (TOBO.GetValue('BLO_TYPEARTICLE')='MAR') OR
       (TOBO.GetValue('BLO_TYPEARTICLE')='PRE') or
       (TOBO.GetValue('BLO_TYPEARTICLE')='POU') or
       (TOBO.GetValue('BLO_TYPEARTICLE')='FRA') then
    BEGIN
    	InsereArticleInterm (TOBO,TOBL,TOBI);
    END;
  end;
end;

procedure TOF_BTGENCONTRETU.ChargeComponents;
begin
  // initialisation de base
  Rootdevis := TreeDevis.items.Add (nil,'DEVIS');
  RootDevis.ImageIndex := 0;
  RootDevis.SelectedIndex := 0;
  RootCtrEtude := TreeCtrEtude.items.Add (nil,'Contre Etude');
  RootCtrEtude.ImageIndex := 0;
  RootCtrEtude.SelectedIndex := 0;
  // chargement du TV_DEVIS
  chargementTreeDevis (TreeDevis,Rootdevis,TobIntermediaire);
  TreeDevis.FullExpand;
  TreeDevis.Refresh;
  TreeDevis.Selected := rootdevis;
end;

procedure TOF_BTGENCONTRETU.chargementTreeDevis(TreeDevis : TTreeView;Pere : TTreeNode; TobIntermediaire :TOB);
VAR Indice : integer;
    TOBSuite : TOB;
    NodeSuite : TTreeNode;
begin
if TobIntermediaire.detail.count = 0 then exit;
if TobIntermediaire.NomTable = 'LIGNE' then exit;
for Indice := 0 to TobIntermediaire.detail.count -1 do
    begin
    TOBSuite := TobIntermediaire.detail[Indice];
    NodeSuite := TreeDevis.items.Addchild (Pere, DefiniLibelle(TOBSuite));
    NodeSuite.data := ToBSuite;
    AffecteImage (NodeSuite,TOBSuite);
    if (TOBSUITE.GetValue('TYPE')='PAR') or (TOBSUITE.GetValue('TYPE')='OUV') or (TOBSUITE.GetValue('TYPE')='ARP') then
       begin
       chargementTreeDevis (TreeDevis,NodeSuite,TobSuite);
       end;
    end;
end;

function TOF_BTGENCONTRETU.DefiniLibelle(TOBL: TOB): string;
begin
	if (TOBL.GetValue('TYPE') = 'PAR') then
  begin
    result := copy(trim(TOBL.GetValue('LIBELLE')),1,70);
  end else
  begin
    result := '';
    if TOBL.GetString('NATURETRAVAIL')='002' then Result := result + '[Sous traitance] ';
    result := result + '['+TOBL.GetValue('ARTICLE')+'] '+ copy(trim(TOBL.GetValue('LIBELLE')),1,35);
    if TOBL.GetValue('QTE') <> 0 then result := result + '  ( '+FloatToStr(TOBL.GetValue('QTE'))+' )';
  end;
end;


procedure TOF_BTGENCONTRETU.ChargeListImage;
var UneImage : Timage;
begin
  UneImage := TImage(GetControl('RIEN'));
  if TheListImage.add(TBitMap(UneImage.Picture.Bitmap ),nil) < 0 then exit;
  UneImage := TImage(GetControl('PARAGRAPHE'));
  if TheListImage.add(TBitMap(UneImage.Picture.Bitmap ),nil) < 0 then exit;
  UneImage := TImage(GetControl('MARCHANDISE'));
  if TheListImage.add(TBitMap(UneImage.Picture.Bitmap ),nil) < 0 then exit;
  UneImage := TImage(GetControl('PRESTATION'));
  if TheListImage.add(TBitMap(UneImage.Picture.Bitmap),nil) < 0 then exit;
  UneImage := TImage(GetControl('FRAIS'));
  if TheListImage.add(TBitMap(UneImage.Picture.Bitmap),nil) < 0 then exit;
  UneImage := TImage(GetControl('OUVRAGE'));
  if TheListImage.add(TBitMap(UneImage.Picture.Bitmap),nil) < 0 then exit;
  UneImage := TImage(GetControl('PRIXPOSE'));
  if TheListImage.add(TBitMap(UneImage.Picture.Bitmap),nil) < 0 then exit;
  UneImage := TImage(GetControl('PRIXPOSEPAPA'));
  if TheListImage.add(TBitMap(UneImage.Picture.Bitmap),nil) < 0 then exit;
  //
  TreeDevis.Images := TheListImage;
  TreeCtrEtude.Images := TheListImage;
end;

procedure TOF_BTGENCONTRETU.AffecteImage (NodeSuite : TTreeNode ; TOBSuite : TOB);
begin
  if TOBSUite.GetValue('TYPE') = 'PAR' then
  begin
    NodeSuite.ImageIndex := 1;  // Paragraphe
    NodeSuite.SelectedIndex := 1;  // Paragraphe
  end else if TOBSUite.GetValue('TYPE') = 'MAR' then
  begin
    NodeSuite.ImageIndex := 2; // Marchandise
    NodeSuite.SelectedIndex := 2;
  end else if TOBSUite.GetValue('TYPE') = 'PRE' then
  begin
    NodeSuite.ImageIndex := 3; // prestation
    NodeSuite.SelectedIndex := 3;
  end else if TOBSUite.GetValue('TYPE') = 'FRS' then
  begin
    NodeSuite.ImageIndex := 4; // Frais
    NodeSuite.SelectedIndex := 4;
  end else if TOBSUite.GetValue('TYPE') = 'OUV' then
  begin
    NodeSuite.ImageIndex := 5; // ouvrage
    NodeSuite.SelectedIndex := 5;
  end else if TOBSUite.GetValue('TYPE') = 'ARP' then
  begin
  	if TOBSUite.GetValue('PRIXPOSEPAPA')= 1 then
    begin
      NodeSuite.ImageIndex := 7; // Prix posé pere
      NodeSuite.SelectedIndex := 7;
    end else
    begin
      NodeSuite.ImageIndex := 6; // Prix posé sous detail
      NodeSuite.SelectedIndex := 6;
    end;
  end;
end;

procedure TOF_BTGENCONTRETU.TreeCtrEtudeDragDrop(Sender, Source: TObject;X, Y: Integer);
var Destinataire,Emetteur,TTI : TTreeNode;
		indice,II : integer;
    TT : Tlist;
begin
	// se produit lorsqu'un objet est relaché dans le treeview Contre etude
  // Sender est le treeview contre etude ...
  TT := Tlist.create;
  TTreeView(Source).GetSelections(TT);
  For Indice := 0 to TT.Count -1 do
  begin
    TTI := TT.Items[Indice];
    if TTI.Count > 0 then
    begin
    	For II := 0 to TTI.Count -1 do
      begin
      	TTreeView(source).Deselect ( TTI.Item [II]);
      end;
    end;
  end;
  //
  TT.Clear;
  TTreeView(Source).GetSelections(TT);
  for Indice := 0 to TT.Count -1 do
  begin
    Emetteur := TT.Items [indice] ;
    Destinataire := TTreeView(Sender).GetNodeAt (X,Y);
    if Destinataire = nil then exit;
    if Source = TreeDevis then
    begin
      // deplacement du devis vers contre etude
      DevisVerCtrEtude (Emetteur,Destinataire);
      SupprimeEmetteurDevis (Emetteur);
    end else
    begin
      // deplacement inter contre etude
      InterCtrEtude (Emetteur,Destinataire,X,Y);
    end;
  end;
  TT.Free;
end;

procedure TOF_BTGENCONTRETU.InterCtrEtude (Emetteur,Destinataire : TTreeNode; X,Y : integer);
var ModeAjout : TNodeAttachMode;
		QteAplat : double;
    RectCtrEtude : Trect;
begin
	QteAplat := 0;

  if (Emetteur.data <> nil) and (TOB(Emetteur.data).GetValue('TYPE') <> 'PAR') then
  begin
    QteAPlat := GetQteAplat (Emetteur);
    TOB(Emetteur.data).PutValue('QTE',QteAPlat);
    Emetteur.Text := DefiniLibelle (TOB(Emetteur.data));
  end;

  if (Emetteur.data <> nil) and (TOB(Emetteur.data).GetValue('TYPE') = 'PAR') and
  	 (Destinataire.data <> nil) and (TOB(Destinataire.data).GetValue('TYPE') = 'PAR') then
  begin
    TraiteChoixParagraphe (Emetteur,Destinataire,X,Y);
    exit;
  end;

	if Destinataire = RootCtrEtude  then
  begin
  	ModeAjout := naAddChildFirst;
  end else
  begin
    if (Destinataire.data <> nil) and (TOB(Destinataire.data).GetValue('TYPE')='PAR') then
    begin
      ModeAjout := naAddChildFirst;
    end else
    begin
      ModeAjout := naAddFirst ;
    end;
  end;
	Emetteur.MoveTo (Destinataire,ModeAjout);
end;

procedure TOF_BTGENCONTRETU.AjouteFilleCtrEtude (Emetteur,Destinataire : TTreeNode);
var Indice : integer;
		TheDest : TTreeNode;
begin
  TheDest := TreeCtrEtude.Items.AddChild (Destinataire,Emetteur.text);
  TheDest.data := Emetteur.data;

  AffecteImage (TheDest,TOB(TheDest.data));

  if Emetteur.Count > 0 then
  begin
  	for Indice := 0 to Emetteur.count -1 do
    begin
    	AjouteFilleCtrEtude (Emetteur.Item [Indice],TheDest);
    end;
  end;
end;

function TOF_BTGENCONTRETU.AjouteLigneCtrEtude (Emetteur,Destinataire : TTreeNode) : TTreeNode;
var TheDestinataire : TTreeNode;
		TOBEmetteur : TOB;
    QteAplat : double;
begin
	TheDestinataire := Destinataire;
	if (TheDestinataire.data <> nil) and (TOB(TheDestinataire.data).GetValue ('TYPE') <> 'PAR') then
  begin
  	// si ce n'est pas un paragraphe ... on le recherche
    repeat
    	TheDestinataire := TheDestinataire.parent;
      if (TheDestinataire.data <> nil) and (TOB(TheDestinataire.data).GetValue ('TYPE') = 'PAR') then break;
    until TheDestinataire.data = nil;
  end;
  // si ce n'est pas un paragraphe on calcule la quantité réelle
  if TOB(Emetteur.data).GetValue('TYPE') <> 'PAR' then QteAPlat := GetQteAplat (Emetteur)
  																								else QteAPlat := 0;
 	// c'est un paragraphe ou le debut de la contre etude on l'ajoute comme une fille
  result := TreeCtrEtude.Items.AddChild (TheDestinataire,Emetteur.text);
//  end;
  result.data := Emetteur.data;
//  if QteAplat <> 0 then
  begin
  	TOB(result.data).putValue('QTE',QteAplat);
    result.Text := DefiniLibelle (TOB(result.data));
  end;
  AffecteImage (result,TOB(result.data));
end;

procedure TOF_BTGENCONTRETU.SupprimeEmetteurDevis (Element : TTreeNode);
var Detail : TTreeNode;
		Indice : integer;
begin
	if Element.Count > 0 then
  begin
  	Indice := 0;
    repeat
    	Detail := Element.Item [Indice];
      SupprimeEmetteurDevis (Detail);
    until Indice >= Element.count;
  end;
  if Element <> RootDevis then Element.Delete;
end;

procedure TOF_BTGENCONTRETU.DevisVerCtrEtude (Emetteur,Destinataire : TTreeNode);
var TheDest : TTreeNode;
		Indice : integer;
begin
	TheDest := Destinataire;
	if Emetteur.data <> nil then
  begin
  	TheDest := AjouteLigneCtrEtude (Emetteur,Destinataire);
  end;
  for Indice := 0 to Emetteur.Count -1 do
  begin
    AjouteFilleCtrEtude (Emetteur.Item [Indice],TheDest);
  end;
  Destinataire.Expand (false);
end;
{
procedure TOF_BTGENCONTRETU.AfficheCoordonnees (X,Y: integer;TheRecepteur : TTreeNode);
var OUINON : string;
begin
	if TheRecepteur = nil then OUINON := '-' else OUINON := 'X';
	THLabel(getControl('COORDONNESOURIS')).Caption := Format ('%d / %d  : %s',[X,Y,OUINON]);
end;
}
procedure TOF_BTGENCONTRETU.TreeCtrEtudeDragOver(Sender, Source: TObject;X, Y: Integer; State: TDragState; var Accept: Boolean);
var	TheRecepteur : TTreeNode;
		LaTOB : TOB;
begin
	// se produit lorsqu'on essaye de deplacer un element du treeview contre etude
  // Source est l'object qui glisse
  // sender est l'emplacement desire pour relacher l'objet
  accept := VerifAutorisationDeplacement (Source);
  if not Accept then exit;
  TheRecepteur := TTreeView(Sender).GetNodeAt (X,Y);
  if TheRecepteur = nil then BEGIN Accept := false; exit; end;
  LaTOB := TOB(TheRecepTeur.Data); if LaTOB = nil then exit;
//  if (LaTOB.GetValue('TYPE')='OUV') or (laTOB.GetValue('TYPE')='ARP') then Accept := false
  if (LaTOB.GetValue('TYPE')='OUV') then Accept := false
  																	else Accept := VerifElement (TheRecepteur);
end;

function TOF_BTGENCONTRETU.verifElement (TheSelectedItem : TTreeNode) : boolean;
var ThePere : TTreeNode;
		LeTOB : TOB;
begin
	result := true;
  ThePere := TheSelectedItem.Parent;
  if ThePere <> nil then
  begin
    LeTOB := TOB(ThePere.data);
    if LeTOB = nil then exit;
    if (LeTOB.GetValue('TYPE')='OUV') or (leTOB.GetValue('TYPE')='ARP') then
    BEGIN
    	result := false;
      exit;
    END;
    result := VerifElement(ThePere);
  end;
end;

function TOF_BTGENCONTRETU.VerifAutorisationDeplacement (Source : TObject) : boolean;
var TheSelectedItem : TTreeNode;
		LeTOB : TOB;
    LETypeArticle : string;
    IdSel : string;
    Indice : integer;
    TOBGeniteur : TOB;
begin
  //
  result := true;
  TheSelectedItem := TTreeview(Source).Selected; if TheSelectedItem = nil then exit;
  // on autorise le deplacement d'un objet s'il n'est pas un article de type marchandise ou prix pose
  // provenant d'un ouvrage - ouvrage uniquement
  leTOB := TOB(TheSelectedItem.data); if LeTOB = nil then exit;
  LeTypeArticle := LeTOB.GetValue('TYPE');
  IdSel := LeTOB.GetValue('UNIQUEID');
  if (LeTypeArticle = 'OUV') or (LeTYpeArticle = 'ARP') or (LeTypeArticle= 'PAR') Then
  begin
  	// trouve le pere de niveau le plus important (geniteur...)
    TOBgeniteur := TrouveGeniteur (IDSel);
    if TOBGeniteur <> nil then
    begin
  		if DetailOuvrageDejaPresent (TOBGeniteur,Source,IDSel) then BEGIN result := false; exit; END;
    end;
  end;
  if (Source = TreeCtrEtude) and (Pos(LeTypeArticle,'PAR;OUV') = 0) then
  begin
    result := verifElement (TheSelectedItem);
  end;
end;

function TOF_BTGENCONTRETU.FindNode (NodePere : TTreeNode ; TheClef : string) : TTreeNode;
var Indice : integer;
		NodeSuite,NodeFille : TTreeNode;
begin
	result := nil;
  for Indice := 0 to NodePere.count -1 do
  begin
    NodeSuite := NodePere.item[Indice];
    if NodeSuite.data = nil then break;
    if TOB(NodeSuite.data).GetValue('UNIQUEID')=TheClef then BEGIN Result := NodeSuite; break; END;
    if NOdeSuite.Count > 0 then
    begin
    	NodeFille := FindNode (NodeSuite,TheClef);
      if NodeFille <> nil then BEGIN result := NodeFille; break; END;
    end;
  end;
end;

procedure TOF_BTGENCONTRETU.InsereTreeDevis (NodePere : TTreenode; TheTOB : TOB);
var ThePere : string;
    Pere,NodeSuite : TTreeNode;
    TheTOBPere,TheTOBCurrent : TOB;
    Copain : TTreeNode;
begin
	if (TheTOB = nil) or (TheTOB.GetValue('UNIQUEID')='') then exit;
  // on essaye de retrouver papa dans la foule
  ThePere := TheTOB.GetValue('CLEFPERE');
  if Thepere = '' then
  begin
  	Pere := RootDevis;
  end else
  begin
  	TheTOBPere := TOBIntermediaire.FindFirst (['UNIQUEID'],[ThePere],true);
  	if TheTOBPere = nil then exit;
    InsereTreeDevis(NodePere,TheTOBPere);
    Pere := FindNode (NodePere,ThePere);
  	if Pere = nil then Pere := RootDevis;
  end;
  //
  if TheTOB.GetValue('ADDED')='X' then exit;
  NodeSuite := FindNode (NodePere,TheTOB.GetValue('UNIQUEID'));
  if NodeSuite = nil then
  begin
  	//
    Copain := TrouvePositionTreeView (Pere,TheTOB.GetValue('UNIQUEID'));
    if Copain = nil then
    begin
    	NodeSuite := TreeDevis.items.Addchild (Pere, DefiniLibelle(TheTOB));
    end else
    begin
    	NodeSuite := TreeDevis.items.Insert (Copain, DefiniLibelle(TheTOB));
    end;
    NodeSuite.data := TheTOB;
    TheTOBCurrent := TOB(NodeSuite.data);
    if TheTOBCurrent.getValue('QTE') <> TheTOBCurrent.getValue('QTEINIT') then
    begin
    	TheTOBCurrent.PutValue('QTE',TheTOBCurrent.getValue('QTEINIT'));
      NodeSuite.Text := DefiniLibelle (TheTOBCurrent);
    end;
    AffecteImage (NodeSuite,TheTOB);
  end;
end;

function TOF_BTGENCONTRETU.TrouvePositionTreeView (NodePere : TTreeNode;UniqueId : string) : TTreeNode;
var Indice : integer;
		NodCurr : TTreeNode;
    TOBCurr : TOB;
begin
	result := nil;
  for Indice := 0 to NodePere.count -1 do
  begin
    NodCurr := NodePere.Item [Indice];
    TOBCurr := TOB(NodCurr.data);
    if TOBCurr.GetValue ('UNIQUEID') > UniqueId then BEGIN result := NodCurr; break; END;
  end;
end;

function TOF_BTGENCONTRETU.VerifNode (NodePere : TTreeNode ;Emetteur : TTreeNode) : TTreeNode;
var Indice : integer;
begin
	result := nil;
  if Emetteur.data <> nil then
  begin
  	InsereTreeDevis (NodePere,TOB(Emetteur.data));
  end;
  if Emetteur.Count > 0 then
  begin
  	for Indice := 0 to Emetteur.count-1 do
    begin
    	VerifNode (NodePere,Emetteur.Item[Indice]  ); // on le recherche ou on recree l'arborescence dans le treeview de depart
    end;
  end;
end;

procedure TOF_BTGENCONTRETU.TreeDevisDragDrop(Sender, Source: TObject; X,Y: Integer);
var emetteur : TTreeNode;
begin
  Emetteur := TTreeView(Source).Selected;
  VerifNode (RootDevis, Emetteur); // on le recherche ou on recree l'arborescence dans le treeview de depart
  RootDevis.Expand(true);
  if Emetteur = RootCtrEtude then RootCtrEtude.DeleteChildren
  													 else Emetteur.Delete;
end;

procedure TOF_BTGENCONTRETU.TreeDevisDragOver(Sender, Source: TObject; X,Y: Integer; State: TDragState; var Accept: Boolean);
var emetteur{, destinataire} : TTreeNode;
		LeTOB : TOB;
    LeTypeArticle : string;
    Indice : integer;
begin
	accept := true;
  if Source = TreeDevis then BEGIN accept := false; exit; END;
  for Indice := 0 to TTreeView(Source).SelectionCount -1 do
  begin
    Emetteur := TTreeView(Source).Selections [indice];
  //  Destinataire := TTreeView(Sender).GetNodeAt (X,Y);
    if (Emetteur.data = nil) and (emetteur <> RootCtrEtude) then BEGIN accept := false; break; END;
    if emetteur.data = nil then exit;
    leTOB := TOB(Emetteur.data); if LeTOB = nil then break;
    LeTypeArticle := LeTOB.GetValue('TYPE');
  end;
(*
  if Pos(LeTypeArticle,'PAR') = 0 then
  begin
    accept := verifElement (Emetteur);
  end;
*)
end;

procedure TOF_BTGENCONTRETU.PopCtrEtudeBeforePopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
var LeType : string;
begin
	if Sender <> TreeCtrEtude then exit;
  TheRecepteurPop := TTreeView(Sender).GetNodeAt (MousePos.X,MousePos.Y);
  if TheRecepteurPop = nil then
  begin
  	TTreeView(Sender).PopupMenu := nil;
    exit;
  end else
  begin
  	TTreeView(Sender).PopupMenu := POPCTREDTUDE;
  end;
  
  if TheRecepteurPop.data = nil then
  begin
    MnAddPar.enabled := true;
    MnDelPar.enabled := true;
  	exit;
  end;
  LeTYpe :=  TOB(TheRecepteurPop.data).GetValue('TYPE');
  if pos(LeType,'PAR') = 0 then
  begin
    MnAddPar.enabled := false;
    MnDelPar.enabled := false;
  end else
  begin
    MnAddPar.enabled := true;
    MnDelPar.enabled := true;
  end;
end;

procedure TOF_BTGENCONTRETU.MNAddParagrapheClick(sender: TObject);
var TheTreeNode,TheFather : TTreeNode;
		TOBS : TOB;
    ClefPere : string;
begin
	ClefPere := '';
  TheFather := TreeCtrEtude.Selected;
  if TheFather.data <> nil then ClefPere := TOB(TheFather.data).GetValue('UNIQUEID');
  TOBS := TOB.Create ('LA LIGNE',TOBNewParag,-1);
  AddChampInterm (TOBS);
  TOBS.putValue('ADDED','X');
  TOBS.putValue('PHASE',EncodePhase(TOBS));
  TOBS.PutValue ('NUMLIGNE',0);
  TOBS.PutValue ('TYPE','PAR');
  TOBS.putValue ('DECOMPOSITIONOUVRAGE','');
  TOBS.putValue ('LIBELLE','Nouveau Paragraphe');
  TOBS.PutValue ('CLEFPERE',ClefPere);
  TOBS.PutValue ('UNIQUEID','');
  TheTreeNode := TreeCtrEtude.Items.AddChildFirst (TheRecepteurPop,'Nouveau Paragraphe');
  TheTreeNode.data := TOBS;
  AffecteImage (TheTreeNode,TOBS);
end;


procedure TOF_BTGENCONTRETU.AddparagFin (Sender : TObject);
var TheTreeNode,TheFather : TTreeNode;
		TOBS : TOB;
    ClefPere : string;
begin
	ClefPere := '';
  TheFather := TreeCtrEtude.Selected;
  if TheFather.data <> nil then ClefPere := TOB(TheFather.data).GetValue('UNIQUEID');
  TOBS := TOB.Create ('LA LIGNE',TOBNewParag,-1);
  AddChampInterm (TOBS);
  TOBS.putValue('ADDED','X');
  TOBS.putValue('PHASE',EncodePhase(TOBS));
  TOBS.PutValue ('NUMLIGNE',0);
  TOBS.PutValue ('TYPE','PAR');
  TOBS.putValue ('DECOMPOSITIONOUVRAGE','');
  TOBS.putValue ('LIBELLE','Nouveau Paragraphe');
  TOBS.PutValue ('CLEFPERE',ClefPere);
  TOBS.PutValue ('UNIQUEID','');
  if TheRecepteurPop <> RootCtrEtude then
  begin
  	TheTreeNode := TreeCtrEtude.Items.Addchild (TheRecepteurPop,'Nouveau Paragraphe');
  end else
  begin
		TheTreeNode := TreeCtrEtude.Items.Addchild (TheRecepteurPop,'Nouveau Paragraphe');
  end;
  TheTreeNode.data := TOBS;
  AffecteImage (TheTreeNode,TOBS);
end;

procedure TOF_BTGENCONTRETU.MNDeleteParagrapheClick(sender: TObject);
var TheTreeNode,TheNode : TTreeNode;
		TheFather : TTreeNode;
    indice : integer;
begin
  TheTreeNode := TreeCtrEtude.Selected;
  if TheTreeNode.data = nil then exit;
  if TOB(TheTreeNode.data).getValue('TYPE')<>'PAR' then exit;
  TheFather := TheTreeNode.Parent;
  if TheTreeNode.count > 0 then
  begin
  	Indice := 0;
    repeat
      TheNode := TheTreeNode.item[Indice];
      TheNode.MoveTo (TheFather,naAddChild);
    until indice >= TheTreeNode.count;
  end;
  TheTreeNode.Delete;
end;

procedure TOF_BTGENCONTRETU.CtrEtudeBeforeEditing(Sender: TObject; Node: TTreeNode; var AllowEdit: Boolean);
begin
  if (Node.data = nil) then BEGIN AllowEdit := false; exit; END;
  if (TOB(Node.data).getValue('ADDED')='-') then BEGIN AllowEdit := false; exit; END;
end;

function TOF_BTGENCONTRETU.DevisToutTransfere: boolean;
begin
	result := true;
  if RootDevis.count = 0 then exit;
  result := NoeudNonVide (RootDevis);
end;

function TOF_BTGENCONTRETU.NoeudNonVide (Depart : TTreeNode)  : boolean;
var Indice : integer;
		Fille : TTreeNode;
begin
	result := true;
  for Indice := 0 to Depart.count -1 do
  begin
    Fille := Depart.Item[Indice];
    if (Fille.data <> nil) and
    	 (TOB(Fille.data).GetValue('TYPE')<>'PAR') and
       (TOB(Fille.data).GetValue('TYPE')<>'OUV') and
       (TOB(Fille.data).GetValue('TYPE')<>'ARP') then
    begin
    	result := false;
      break;
    end;
    if Fille.count > 0 then
    begin
    	result := NoeudNonVide (Fille);
    end;
    if not result then break;
  end;
end;

procedure TOF_BTGENCONTRETU.ValideLaCtrEtude;
var TOBO : TOB;
		TOBOuvrages : TOB;
begin
	TOBOUvrages := TOB.Create('LES OUVRAGES DES DOCUMENTS',nil,-1);
	TOBO := TOB.create('OUVRAGE DE LA PIECE',TOBOuvrages,-1);
  TOBO.Dupliquer (TOBOUvrageResult,true,true);
  AddChampsSupOuvrages(TOBO,TOBCtrEtude);
  CreerPiecesFromLignes (TOBCtrEtude,'DEVTOCTE',Date,true,true,nil,TOBOuvrages);
	TOBOUvrages.free;
end;

procedure TOF_BTGENCONTRETU.AddChampsSupOuvrages(TOBO, TOBCtrEtude: TOB);
begin
  TOBO.AddChampSupValeur ('NATUREPIECEG',TOBCtrEtude.getVAlue('GP_NATUREPIECEG'));
  TOBO.AddChampSupValeur ('SOUCHE',TOBCtrEtude.getVAlue('GP_SOUCHE'));
  TOBO.AddChampSupValeur ('NUMERO',TOBCtrEtude.getVAlue('GP_NUMERO'));
end;

function TOF_BTGENCONTRETU.GetQteAplat(Emetteur: TTreeNode): double;
var StopIt : boolean;
		Pere,CurrentNode : TTreeNode;
begin
	StopIt := false;
	result := 1;
	CurrentNode := Emetteur;
	repeat
    if CurrentNode.data <> nil then
    begin
    	if TOB(CurrentNode.data).getValue('TYPE') <> 'PAR' then
      begin
      	result := result * (TOB(CurrentNode.data).getValue('QTE')/TOB(CurrentNode.data).getValue('QTEDUDETAIL'));
      end;
    end;
  	if (CurrentNode.parent <> nil) then
    begin
    	CurrentNode := CurrentNode.parent;
    end else
    begin
    	StopIt := true;
    end;
  until StopIt;
end;

function TOF_BTGENCONTRETU.VerifExistID (TOBControle : TOb ; Controle : TTreeNode) : boolean;
var Indice : integer;
		CurrTOb : TOB;
    NodeCurr : TTreeNode;
begin
	result := false;
  for Indice := 0 to TOBControle.detail.count -1 do
  begin
  	CurrTOB := TOBControle.detail[Indice];
    if CurrTOB.detail.count > 0 then
    begin
    	result:= VerifExistID (CurrTOB,Controle) ;
      if result then exit;
    end else
    begin
    	NodeCurr := FindNode (Controle,CurrTOB.GetValue('UNIQUEID'));
      if NodeCurr <> nil then BEGIN result := true; Exit; END;
    end;
  end;
end;

function TOF_BTGENCONTRETU.DetailOuvrageDejaPresent(TOBGeniteur : TOB; Source: TObject; IDSel: string): boolean;
var Controle : TTreeNode;
		Indice : integer;
begin
	result := false;
  if Source = TreeDevis  then
  begin
  	Controle := RootCtrEtude;
    result := VerifExistID (TOBGeniteur,Controle);
  end else
  begin
  	result := false;
  end;
end;

function TOF_BTGENCONTRETU.TrouveGeniteur (IDSel : string) : TOB;
var TOBCur,TOBPere : TOB;
begin
  result := TOBIntermediaire.FindFirst (['UNIQUEID'],[IDSel],true);
  repeat
    if TOBCur <> nil then
    begin
      TOBPere := TOBCur.Parent;
      if (TOBPere.getValue('TYPE')='OUV') or
      	 (TOBpere.GetValue('TYPE')='ARP') or
      	 (TOBpere.GetValue('TYPE')='PAR') then result:= TOBpere;
      TOBCur := TOBpere;
    end else break;
  until TOBCur.GetValue('TYPE')='';
end;


procedure TOF_BTGENCONTRETU.TraiteChoixParagraphe (Emetteur,Destinataire : TTreeNode; X,Y : integer);
var Pos: Tpoint;
begin
	if Emetteur = destinataire then exit;
  ParamEmetteur := Emetteur;
  ParamDestinataire := Destinataire;
  Pos := TreeCtrEtude.ClientToScreen (Point( X,Y));
  TPopupMenu(GetControl('POPCHOIXDEPPAR')).Popup (Pos.X,Pos.Y);
end;

procedure TOF_BTGENCONTRETU.InsertParDansPar (Sender : TObject);
begin
	ParamEmetteur.MoveTo (ParamDestinataire,naAddChildFirst);
  ParamEmetteur := nil;
  ParamDestinataire := nil;
end;

procedure TOF_BTGENCONTRETU.DeplacePar (Sender : TObject);
begin
	ParamEmetteur.MoveTo (ParamDestinataire,naInsert);
  ParamEmetteur := nil;
  ParamDestinataire := nil;
end;


procedure TOF_BTGENCONTRETU.DeplaceParB (Sender : TObject);
var NewDestinataire : TTreeNode;
		ModeAjout : TNodeAttachMode;
begin
	modeAjout := NaInsert;
	NewDestinataire := paramDestinataire.GetNextSibling;
  if NewDestinataire = nil then
  begin
  	NewDestinataire := paramDestinataire.Parent;
    modeAjout := naAddchild;
  end;
	ParamEmetteur.MoveTo (NewDestinataire,Modeajout);
  ParamEmetteur := nil;
  ParamDestinataire := nil;
end;

Initialization
  registerclasses ( [ TOF_BTGENCONTRETU ] ) ;
end.

