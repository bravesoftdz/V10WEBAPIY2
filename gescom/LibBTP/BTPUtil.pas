unit BTPUtil;

interface

uses Classes,HEnt1,Ent1,EntGC,UTob, AGLInit, SaisUtil, UtilPGI,forms,Math,HrichOle,graphics,controls,dialogs,
{$IFDEF EAGLCLIENT}
     UtileAGL,MaineAGL,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}EdtREtat,FE_Main,uPDFBatch,
{$ENDIF}
		 BTFactImprTob,
     UtilPhases,HRichEdt,
     FactTOB,TiersUtil,
     UTofGCDatePiece,FactUtil,HCtrls,ParamSoc,SysUtils,FactComm,uRecupSQLModele,
     hPDFPrev,hPDFViewer,hmsgbox,CalcOLEGenericBTP,UTofListeInv,FactCommBtp,uEntCommun
    , DB
    ;

Const { Concepts }
      //bt500 modification des TAG du menu concept pour éviter confusion avec Gescom et Affaire
      bt510=510; //Modif Prospect
      bt511=511; //Modif Client
      bt512=512; //acces au sous-détail
      bt513=513; //accès à la remise
      bt514=514; //accès à l'escompte
      bt515=515; //accès à la simulation de rentabilité
      bt516=516; //accès à la suppression de pièce
      bt517=517; //accès au coeff FG en simulation de rentabilité
      bt518=518; //Modif Fournisseur
      Bt519=519; //Validation définitve en saisie de reste à dépenser
      Bt520=520; //GUINIER - Modification des pièce de vente
      bt500=500; //Affichage commande en génération
      bt521=521; //Gestion des montants de productions
      Bt522=522; //GUINIER - Déblocage RAD
      Bt530=530; //GUINIER - Suivi des cdes fournisseurs


Type
  BTTraitChantier = (BTTContrEtud,BTTModif,BTTSuppress);

	StInfosAffaire = record
    TypFac : string;
    ModeFact : boolean;
  end;

  TGenereLivraison = class
    private
    TOBPieces,ThePieceGenere,TheAffaires,TOBClients : TOB;
    TOBLivrDirecte : TOB;
    TOBresult : TOB;
    CompteRendu : boolean;
    DatePriseEnCompte : TDateTime;
    Duplication : Boolean;
    Transformation : Boolean;
    Action : TActionFiche;
//    LesNumPieces : string;
    procedure ChargeTobPieces(LaTob: TOB);
    procedure MiseAjourPieceprec(TOBpiece: TOB);
    procedure AjusteLesLignes;
    procedure SetInfosLivraison(TOBL: TOB);
    procedure SetInfosClient (TOBL : TOB);
    procedure AddLivraisonDirecte (cledoc : r_cledoc;  QteRecep : double);
    function MakeWhere: string;
    function AddDetail(TOBLD: TOB): string;
    function DejaLivre(Piece: string): boolean;
    public
    TheLignesALivrer : TOB;
    constructor create;
    destructor destroy; override;
    procedure GenereThepieces ;
    procedure DecortiqueBesoinLivraison;
    procedure TransformeAchatEnVente (TOBL : TOB);

  end;

var TheImpressionViaTOB : TImprPieceViaTOB;
		TOBDispoLoc : TOB;
    FromDispoGlob,FromDispoArt : boolean;
function SetPlusAffaire0 (TypeForce : string = ''): string;
Function GetEtatAffaire (Codessaff : string) : String ;
Procedure MajSousAffaire(TobPiece,TOBAcomptes:TOB; CodeAffaireAvenant : string;Action : TActionFiche;Duplic:Boolean;SaisieAvanc:boolean;ModeTrait:string=''; CodeEtat : String='');
Function RenvoieTypeFact (Codessaff : string) : String ;
Procedure MajTypeFact (TobPiece : TOB; TypFac : string);
Function RenvoieCodeReg (codeaffaire : string) : String ;
Function RenvoieTypeRes (codeprestation : string) : String ;
Function RenvoieNumPiece (codeaffaire : string) : String ;
Function GetMiniAvancAcompte (CodessAff : string) : Integer;
Procedure MajApresGeneration(TobPiece,TOBPieceRG,TOBBasesRG,TOBAcomptes : Tob; DEV : Rdevise;NumSituation : Integer=-1);
Procedure MajQtefacture(TobPiece,TobPiece_ori : Tob);
Procedure BtMajSDP(CleDoc : R_CleDoc;EnTOBUniquement : boolean = false);
Procedure BtMajTenueStock(CleDoc : R_CleDoc);
Function DerniereSituation(TOBPiece : TOB; Quefacture : Boolean=false) : Boolean ;
Procedure ModifSituation(TobPiece,TOBOuvrage,TobPieceRG,TOBPieceRg_O,TOBBasesRG,TOBAcomptes : Tob; DEV : RDevise; var DocArecalc : r_cledoc );
procedure ReajusteSituation (TOBPiece,TOBAcomptes,TOBPieceRG,TOBBasesRG : TOB; DEV : RDevise);
function SupprimeSituation(TobPiece,TOBOuvrage,TobPieceRG,TOBBasesRG,TOBAcompte : Tob; DEV : RDevise ) : boolean;
//Function CalculMontantRevient (TOBPiece, TOBporcs : TOB; DEV:RDevise;InclusStInFG : boolean; var ExistFG : boolean) : double ;
Function BTPSupprimePieceFrais (CleDoc : R_CLEDOC):boolean; overload;
Function BTPSupprimePieceFrais (TOBPiece : TOB):boolean; overload
Function BTPSupprimePiece (CleDoc : R_CLEDOC;AvecAdresses:boolean=true):boolean;
Function Origine_EXCEL(TOBPiece : TOB) : Boolean;
Function Lexical_RechArt ( GS : THGrid ; TitreSel,NaturePieceG,DomainePiece,SelectFourniss : String; var CodeArticle : string  ) : boolean ;
function CodeLaPiece (cledoc : r_cledoc) : string;
Procedure ImprimePieceBTP (cledoc : r_cledoc; Bapercu : boolean; Modele, Req ,TypeFacturation : string; DGD : boolean = false; ImpressionViaTOB : TImprPieceViaTOB = nil);
function MajMontantAcompte(TOBPiece,TOBAcomptes : TOB) : boolean;
procedure MajAffaireApresGeneration (TobPiece,TOBPieceRG,TOBBasesRG,TOBAcomptes : TOB;DEV: Rdevise; FinTravaux : boolean=false; NumSituation : integer=-1);
procedure MajQtesAvantSaisie (TobPiece : TOB);
function ControleAffaireRef (CodeAffaire : string) : boolean;
procedure GetMontantsAcomptes (TOBAcomptes : TOB;var MontantAcompte : double ; var MontantRegl : double);
function ControleChantierBTP (TOBPiece : TOB; Mode : BTTraitChantier) : boolean;
procedure ReajusteQteReste (TOBGenere : TOB);
// NEW One
function ISPrepaLivFromAppro : boolean;
//function ISGenereLivFromAppro : boolean;
procedure GenereLivraisonClients (ThePieceGenere : TOB;Action : TActionFiche;transfoPiece,DuplicPiece: boolean;DemandeDate: boolean = true;CompteRendu : boolean=true; TOBresult : TOB= nil);
function LivraisonVisible : boolean;
Procedure InitParPieceBTP;
function ExisteDocAchatLigne (theLigne : String) : boolean;
function ISAliveLine(TOBpiece: TOB): boolean;
Function GetQtelivrable (TOBL : TOB;gestionConso : TGestionPhase; TOBART : TOB = nil; PriseSurStock : boolean=true) : double;
Function GetQtelivrableHlien (TOBL : TOB;gestionConso : TGestionPhase; TOBART : TOB = nil) : double;
function NaturepieceOKPourOuvrage (TOBPiece : TOB) : boolean;
function IsPrestationInterne(Article : String) : boolean;
function TiersModifiableInDocument(TOBPiece : TOB) : Boolean;
function IsRetourFournisseur (TObpiece : TOB) : boolean;
function IsRetourClient(TOBPiece : TOB) : Boolean;
function ControleDateDocument ( StD : String ; Naturepiece : string) : integer ;
function IsSamePieceOrFromAchat (TOBL : TOB) : boolean;
function GetContact (TOBContact : TOB;LeTiers : string;NumeroContact : integer) : boolean; overload;
function GetContact (LeTiers : string;NumeroContact : integer; var libcontact : string) : boolean; overload;
function GetCodeBQ (TOBBanqueCP : TOB; CodeAuxiliaire: String; NumeroRIB : Integer) : boolean; Overload;
function GetCodeBQ (Auxiliaire : String; NumeroRIB : Integer; var RefBancaire : string) : boolean; overload;
Function  RenvoieInfosAffaireDevis (Codessaff : string) : StInfosAffaire ;
Procedure InitCalcDispo;
procedure InitCalcDispoArticle;
procedure FinCalcDispo;
function IsPrevisionChantier (TOBpiece : TOB) : boolean;
function GetNumCompteur (Souche : string; ODate : TDateTime; var TheNumero : integer) : boolean;
function isPieceGerableFraisDetail(naturePiece : string) : boolean;
function IsContreEtude (TOBpiece : TOB) : boolean;
function GetTauxFg (TOBporcs : TOB) : double;
procedure RestitueAcompte(Cledoc : R_cledoc; Journal : string; Numecr : integer ;MontantDev,Montant : double);
procedure DefiniPvLIgneModifiable (TOBpiece,TOBPOrcs: TOB);
function GetCoefFC (TOBpiece : TOB) : double; overload;
function GetCoefFC (MontantAppliquable,MontantFraisChantier : double) : double; overload;
Function GetContrat (Affaire : string) : string;
procedure DefiniDejaFacture (TOBPiece,TOBOuvrages : TOB;DEV : rdevise; TypeFacturation : string);
procedure RecalculeCoefFrais (TOBPiece,TOBporcs : TOB);
function GetMontantFraisDetail (TOBPiece : TOB; Var ExistFg : boolean) : double;
procedure  SetLigneFacture (TOBL,TOBD : TOB);
procedure InitChampsSup(TOBL : TOB; Champs,tipe : string);
procedure AddChampsSupTable (TOBCur : TOB;PrefixeTable : string);
function getInfoPuissance (NbLig : integer) : integer;
procedure SetItemHvalComboBox (Combo : THValComboBox; Valeur : string);
//
Procedure LectBanque(TOBBQE : TOB; NumeroRIB : String);
Procedure LectRIB(TOBBQE : TOB; Auxiliaire : string; NumeroRIB : Integer);
procedure AppliqueFontDefaut ( TheControl : THRichEditOle);
procedure DecodeFont( StrFont : string; TheFont : TFont) ;
function EncodeFont (Font : Tfont) : string;
function SelectBrushColor (Element : THedit; F : Tform) : boolean;
function SelectFonte(Zone : Thedit; F : TForm) : boolean;
function SelectFontColor (Element : THedit; F : Tform) : boolean;
function SelectColor (Element : THedit; F : Tform) : boolean;
function BTGetSectionStock(Depot : string) : string;
function isParamGeneDepot : Boolean;
function GetShare (NomTable : string) : string;
procedure GetSituationSuivante (TOBpiece : TOB; var NCd : R_CLEDOC);
procedure ValideSituation (TOBPiece,TOBOuvrage : TOB );
//
Function CtrlDateLimiteSaisie(DateSaisie : TDateTime) : Boolean;
//
procedure ChargePieceRepart(AffaireDevis : String; TobLigneFac : TOB);
Procedure RecherchePieceDansLigneFac(TOBMemo, TobLigneFac : TOB);
Procedure TraitementPiecebyMemoFact(TOBMemo : Tob);
procedure DestruitPrepaFact(Nature, Souche  : String;  Numero  : Integer);

//--- GUINIER ---
Function CtrlOkReliquat(TOBL : TOB; Prefixe : String) : Boolean;
procedure AnnuleSituation (TOBpiece,TOBOUvrage,TOBPieceRg : TOB);
Function FormatMultiValComboforSQL(Champ : String) : String;
function ExRichToString (TheObj : ThRichEdit) : string;
function GetPiecesVenteBTP (ForWhat : integer = 3; ForPlus : boolean=true): string;

type
  Publipost = class
  public
    class function GetPrefixFromTableName(TableName : string) : string;
    class function GetTranslateData(FieldName, Code : string) : string;
    class procedure SetData(TobData : TOB; FF : TField);
    class procedure SetFieldsList(Prefix, TableName : string; var FieldsList, SqlFieldsList : string);
    class procedure AddFieldsInTob(TobData : TOB; TableName, SqlFieldsList, Where : string; WithData: Boolean);
    class procedure NewModel(InputFilePath : string; TobFields : TOB; FieldsList : string);
    class procedure UpdateModel(InputFilePath : string; TobFields : TOB; FieldsList : string);
    class function DeleteModel(InputFilePath, Title : string) : Boolean;
    class procedure MergeExecute(TableName, Where, TemplateDoc, OtputDoc : string); 

  end;

implementation

uses
  AffaireUtil
  , Facture
  , FactRg
  , FactGrp
  , FactureBTP
  , TntWideStrings
  , UtilTOBPiece
  , factOuvrage
  , StockUtil
  , FactPiece
  , Mask
  , CbpMCD
  , CbpEnumerator
  , UtilWord
  , wCommuns
  , CommonTools
  ;

function GetPiecesVenteBTP (ForWhat : integer = 3; ForPlus : boolean=true): string;
begin
  if ForPlus then
  begin
    if ForWhat = 2 then Result := 'AND GPP_NATUREPIECEG IN ("FBC","ABC","CC","BLC")'
    else if ForWhat = 1 then Result := 'AND GPP_NATUREPIECEG IN ("FBT","B00","FBP","ABT","APR","ABP","DBT","BAC","DAC","ETU")'
    else Result := 'AND GPP_NATUREPIECEG IN ("FBT","B00","FBP","FBC","ABT","APR","ABP","ABC","DBT","BAC","DAC","ETU","CC","BLC")'
  end else
  begin
    if ForWhat = 2 then Result := 'AND GP_NATUREPIECEG IN ("FBC","ABC","CC","BLC")'
    else if ForWhat = 1 then Result := 'AND GP_NATUREPIECEG IN ("FBT","B00","FBP","ABT","APR","ABP","DBT","BAC","DAC","ETU")'
    else Result := 'AND GP_NATUREPIECEG IN ("FBT","B00","FBP","FBC","ABT","APR","ABP","ABC","DBT","BAC","DAC","ETU","CC","BLC")'
  end;
end;

function ExRichToString (TheObj : ThRichEdit) : string;
var Ipos : integer;
begin
  result := RichToString(TheObj);

  Ipos := Pos(''#$D#$A'\par '#$D#$A'\par \',result);
  While Ipos >0 do
  begin
    result := StringReplace (result,''#$D#$A'\par '#$D#$A'\par \',''#$D#$A'\par \',[]);
    Ipos := Pos(''#$D#$A'\par '#$D#$A'\par \',result);
  end;

  Ipos := Pos(' '#$D#$A'\par '#$D#$A'\par }',REsult);
  While Ipos >0 do
  begin
    result := StringReplace (result,' '#$D#$A'\par '#$D#$A'\par }',' '#$D#$A'\par }',[]);
    Ipos := Pos(' '#$D#$A'\par '#$D#$A'\par }',result);
  end;
end;


procedure AppliqueFontDefaut ( TheControl : THRichEditOle);
var TheFont : Tfont;
		QQ : Tquery;
    OkOk : boolean;
begin
  OKOK := false;
  TheFont := TFont.Create;
  //
  QQ := OpenSQL ('SELECT BFS_FONT FROM BFONTS WHERE BFS_APPLICATION="BLOC"',true,1,'',true);
  if not QQ.eof then
  begin
		Okok := true;
  	DecodeFont (QQ.findField('BFS_FONT').AsString,TheFont);
  end;
  ferme (QQ);
  //
  if not OkOk then
  begin
  	DecodeFont (V_PGI.FontMemo,TheFont);
  end;
  //
  if (TheFont.Name <> '') and (TheControl <> nil) then
  begin
  	TheControl.Font.Name := TheFont.Name;
  	TheControl.Font.Size := TheFont.Size;
  	TheControl.Font.Style  := TheFont.Style;
  end;
end;

procedure DecodeFont( StrFont : string; TheFont : TFont) ;
var TheDescFont : string;
		TheFontname,Attrib,Taille,AttribPlus : string;
begin
  TheFont.name := '';
  if StrFont ='' then exit;
	TheDescFont := StrFont;
  TheFontName := '';
  Attrib := '';
  AttribPlus := '';
  taille := '';
  TheFontName := readTokenSt(TheDescFont);
  TheFont.name := TheFontName;
  if TheDescFont <> '' then
  begin
  	TheFont.Size := StrToInt(readTokenSt(TheDescFont));
  end;
  if TheDescFont <> '' then Attrib := readTokenSt(TheDescFont);
  if Attrib <> '' then
  begin
		 if Attrib = 'I' Then TheFont.Style := [fsItalic] else
		 if Attrib = 'B' Then TheFont.Style := [fsBold];
  end;
  if TheDescFont <> '' then AttribPlus := readTokenSt(TheDescFont);
  if AttribPlus <> '' then
  begin
		 if AttribPlus = 'I' Then TheFont.Style := TheFont.Style + [fsItalic] else
		 if AttribPlus = 'B' Then TheFont.Style := TheFont.Style + [fsBold];
  end;
end;

function EncodeFont(Font: Tfont): string;
begin
  result := Font.name+';'+IntToStr(Font.size);
  if fsItalic in Font.Style then result := result+';I';
  if fsBold in Font.Style then result := result + ';B';
end;

function SelectFonte(Zone : Thedit; F : TForm) : boolean;
Var T			: TFontDialog;
Begin
  result := false;
  T := TFontDialog.Create(F);
  T.Options := T.Options- [fdEffects];

  T.Font.Size  := Zone.Font.Size;
	T.Font.Style := Zone.Font.Style;
  T.Font.Color := Zone.Font.Color;
  T.Font.Name  := Zone.Font.Name;

  if T.Execute then
  begin
		Zone.Font.Size := T.Font.size;
    Zone.Font.Style := T.Font.Style;
    Zone.Font.Name  := T.Font.Name;
    result := true;
  end;
  T.Free;
end;

function SelectBrushColor (Element : THedit; F : Tform) : boolean;
Var T	: TColorDialog;
begin
  result := false;
  //
  T := TColorDialog.Create(F);
  T.color := Element.Color;
  T.CustomColors.Clear;
  T.CustomColors.Add('ColorA=' + InttoHex(Element.Color, 6));
  //
  if T.Execute then
  begin
    Element.color := T.Color;
    Element.Refresh;
  	result := true;
  end;
  T.Free;

end;

function SelectFontColor (Element : THedit; F : Tform) : boolean;
Var T	: TColorDialog;
begin
  result := false;
  //
  T := TColorDialog.Create(F);
  T.color := Element.Font.Color;
  T.CustomColors.Clear;
  T.CustomColors.Add('ColorA=' + InttoHex(Element.Font.Color, 6));
  //
  if T.Execute then
  begin
    Element.font.color := T.Color;
    Element.Refresh;
  	result := true;
  end;
  T.Free;

end;

function SelectColor (Element : THedit; F : Tform) : boolean;
Var T	: TColorDialog;
begin
  result := false;
  //
  T := TColorDialog.Create(F);
  T.color := Element.Color;
  T.CustomColors.Clear;
  T.CustomColors.Add('ColorA=' + InttoHex(Element.Color, 6));
  //
  if T.Execute then
  begin
    Element.color := T.Color;
    Element.Refresh;
  	result := true;
  end;
  T.Free;

end;

function FindPieceTraitOrig (TOBPieceTraitDevis: TOB; Cledoc : r_cledoc;Fournisseur : string) : TOB;
begin
  Result := TOBPieceTraitDevis.findFirst(['BPE_NATUREPIECEG','BPE_SOUCHE','BPE_NUMERO','BPE_INDICEG','BPE_FOURNISSEUR'],
                                         [Cledoc.NaturePiece,Cledoc.Souche,cledoc.NumeroPiece,Cledoc.Indice,Fournisseur],True);
end;

function ChargePieceTraitOrig(TOBPieceTraitDevis: TOB; cledoc : R_CLEDOC ; Fournisseur : string): TOB;
var Sql : string;
    QQ : TQuery;
begin
  Result := nil;
  Sql := 'SELECT * FROM PIECETRAIT WHERE '+WherePiece(cledoc,ttdPieceTrait,True)+' AND BPE_FOURNISSEUR="'+Fournisseur+'"';
  QQ := OpenSQL(Sql,True,1,'',True);
  if not QQ.eof then
  begin
    result := TOB.Create ('PIECETRAIT',TOBPieceTraitDevis,-1);
    result.SelectDB('',QQ);
  end;
  Ferme(QQ);
end;

procedure AjoutSuprAvancementOuv (TOBPieceTraitDevis,TOBL,TOBOuvrage,TOBENREG : TOB; Avancement : boolean);
var indiceNomen,II : integer;
		TOBOUV,TOBOO,TOBD : TOB;
    cledoc,CleDocPre : r_cledoc;
    Req : string;
    QQ : Tquery;
    pourcent,NatureTravailOuv : double;
    TOBPTraitDevisL,TOBOP,TOBPLAT : TOB;
    RefPieceCur : string;
begin
  TOBPLAT := TOB.Create ('L OUVRAGE PLAT',nil,-1);
  TRY
    IndiceNomen := TOBL.GetValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
    NatureTravailOuv := VALEUR(TOBL.GetValue('GLC_NATURETRAVAIL'));
    TOBOUV := TOBOuvrage.detail[IndiceNomen-1]; if TOBOUV = nil then exit;
    for II := 0 to TOBOUV.detail.count -1 do
    begin
      TOBOO := TOBOUV.detail[II];
      if TOBOO.GEtValue('BLO_PIECEPRECEDENTE') = '' then exit;
      DecodeRefPieceOUv (TOBOO.GEtValue('BLO_PIECEPRECEDENTE'),cledoc);
      Req := 'SELECT * FROM LIGNEFAC WHERE '+ WherePiece (cledoc,ttdLigneFac,false)+' AND BLF_NUMORDRE=0 AND BLF_UNIQUEBLO='+IntToStr(Cledoc.UniqueBlo);
      QQ:= OpenSql (Req,true,-1,'',true);
      if not QQ.eof then
      begin
        TOBD := TOB.Create ('LIGNEFAC',TOBENREG,-1);
        TOBD.SelectDB('',QQ);
        if Avancement then
        begin
          TOBD.PutValue('BLF_QTEDEJAFACT',TOBOO.GetValue('BLF_QTEDEJAFACT'));
          TOBD.PutValue('BLF_QTECUMULEFACT',TOBOO.GetValue('BLF_QTEDEJAFACT'));
          //
          TOBD.PutValue('BLF_MTDEJAFACT',TOBOO.GetValue('BLF_MTDEJAFACT'));
          TOBD.PutValue('BLF_MTCUMULEFACT',TOBOO.GetValue('BLF_MTDEJAFACT'));
          //FV1 - 25/11/20015 - FS#1802 - BENETEAU - impossible de supprimer une facture message d'erreur
          if TOBD.GetValue('BLF_QTEMARCHE') <> 0 then
          begin
          Pourcent := arrondi((TOBD.GetValue('BLF_QTECUMULEFACT') / TOBD.GetValue('BLF_QTEMARCHE')) * 100,2);
          TOBD.PutValue('BLF_POURCENTAVANC',Pourcent);
          end
        end else
        begin
          TOBD.PutValue('BLF_QTEDEJAFACT',TOBD.GetValue('BLF_QTEDEJAFACT')-TOBOO.GetValue('BLF_QTESITUATION'));
          //
          TOBD.PutValue('BLF_MTDEJAFACT',TOBD.GetValue('BLF_MTDEJAFACT')-TOBOO.GetValue('BLF_MTSITUATION'));
          //FV1 - 25/11/20015 - FS#1802 - BENETEAU - impossible de supprimer une facture message d'erreur
          if TOBD.GetValue('BLF_QTEMARCHE') <> 0 then
          begin
          Pourcent := arrondi((TOBD.GetValue('BLF_QTEDEJAFACT') / TOBD.GetValue('BLF_QTEMARCHE')) * 100,2);
          TOBD.PutValue('BLF_POURCENTAVANC',0);
        end;
      end;
      end;
      ferme(QQ);
    end;
    //
    if NatureTravailOuv >= 10 then
    begin
      RefPieceCur := EncodeRefPiece(TOBL);
      DecodeRefPiece (RefPieceCur,cledoc);
      DecodeRefPiece(TOBL.getValue('GL_PIECEPRECEDENTE'),cledocPre);

      QQ := OpenSQL('SELECT * FROM LIGNEOUVPLAT WHERE '+WherePiece(cledoc,ttdLigneOUVP,True,True),True,-1,'',True);
      TOBPLAT.LoadDetailDB('LIGNEOUVPLAT','','',QQ,false);
      ferme (QQ);
      for II := 0 to TOBPLAT.Detail.count -1 do
      begin
        TOBOP := TOBPLAT.Detail [II];
        if (VALEUR(TOBOO.GetValue('BOP_NATURETRAVAIL'))<10) then
        begin
          TOBPTraitDevisL := FindPieceTraitOrig (TOBPieceTraitDevis,cledocPre,TOBOP.GetValue('BOP_FOURNISSEUR'));
          if TOBPTraitDevisL = nil then
          begin
            TOBPTraitDevisL := ChargePieceTraitOrig(TOBPieceTraitDevis,cledocPre,TOBOP.GetValue('BOP_FOURNISSEUR'));
          end;
          if TOBPTraitDevisL <> nil then
          begin
            TOBPTraitDevisL.PutValue('BPE_MONTANTFAC',TOBPTraitDevisL.GetValue('BPE_MONTANTFAC')-TOBOP.GetValue('BOP_TOTALTTCDEV'));
          end;
        end;
      end;
    end;
  FINALLY
    TOBPLAT.Free;
  END;
end;





function SetPlusAffaire0 (TypeForce : string = ''): string;
begin
	if TypeForce = '' then
  begin
    result := ' AND ((CO_CODE="A")';
//    if VH_GC.BTSeriaAO then Result := Result + '  OR (CO_CODE="P")';
    if VH_GC.BTSeriaContrat then result := result + ' OR (CO_CODE="I")';
    if VH_GC.BTSeriaIntervention then result := result + ' OR (CO_CODE="W")';
    result := result + ')';
  end else
  begin
  	result := ' AND (CO_CODE="'+TypeForce+'")';
  end
end;

Function GetEtatAffaire (Codessaff : string) : String ;
{
Var Q : TQuery ;
    Req, EtatAffaire : String;
}
BEGIN
// OPTIMISATIONS LS
{
EtatAffaire := 'ENC';
//récupération du type de facturation pour l'affaire liée au devis
Req:='SELECT AFF_ETATAFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="'+Codessaff+'"' ;
Q:=OpenSQL(Req,TRUE) ;
if Not Q.EOF then EtatAffaire:=Q.Fields[0].AsString;
Ferme(Q) ;

result := EtatAffaire;
}
result := GetChampsAffaire (CodeSSAff,'AFF_ETATAFFAIRE');
if Result= '' then result := 'ENC';

end;

function ControleAffaireRef (CodeAffaire : string) : boolean;
var nbpiece : integer;
    CleDocAffaire : R_CLEDOC;
    Req : string;
begin
result := true;
if (CodeAffaire <> '') then
   begin
   NbPiece := SelectPieceAffaire(CodeAffaire, 'AFF', CleDocAffaire);
   if NbPiece = 1 then
      begin
      Req := 'UPDATE AFFAIRE SET AFF_MULTIPIECES = "-" WHERE AFF_AFFAIRE='+'"'+CodeAffaire+'"';//MODIFBTP190101
      if (ExecuteSQL(Req) < 1) then result := false;
      end;
   end;
end;

procedure ReinitEltAffaire (TOBAffaire : TOB);
begin
TOBAffaire.putValue('AFF_ACOMPTE',0);
TOBAffaire.putValue('AFF_ACOMPTEREND',0);
TOBAffaire.PutValue('AFF_MULTIPIECES', '-'); //MODIFBTP290301
TOBAffaire.PutValue('AFF_STATUTAFFAIRE', '');
TOBAffaire.PutValue('AFF_ETATAFFAIRE', 'ENC');
if TOBAffaire.GetValue('AFF_GENERAUTO') = '' then TOBAffaire.PutValue('AFF_GENERAUTO', VH_GC.AFFGenerAuto);
end;

Procedure MajSousAffaire(TobPiece,TOBAcomptes:TOB; CodeAffaireAvenant : string;Action : TActionFiche;Duplic:Boolean;SaisieAvanc:boolean;ModeTrait:string=''; CodeEtat : String='');
Var CodeAffaire, Req, annee : string;
    NbPiece, i, Numpiece : Integer;
    Part0, Part1, Part2, Part3,Avenant : String;
    CleDocAffaire : R_CLEDOC;
    TOBAffaire : TOB;
    TOBAffaire2 : TOB;
    codessaff, MultiPiece,SqlSup : String;
    Q : TQuery ;
    Acompte,Regl : double;
BEGIN

CodeAffaire := TobPiece.GetValue('GP_AFFAIRE');
// Contrôle nombre de devis pour l'affaire si code affaire renseigné sinon toujours -1
// Attention la pièce courante est déjà créée donc SelectPieceAffaire renvoie au moins 1
if (CodeAffaire <> '') and (TobPiece.Getvalue ('GP_NATUREPIECEG') = VH_GC.AFNatAffaire) then
  NbPiece := SelectPieceAffaire(CodeAffaire, 'AFF', CleDocAffaire) // Lancement avec AFF pour lecture des devis
else
  NbPiece := -1;
// Chargement clé document pour la pièce en cours
FillChar(CleDocAffaire,Sizeof(CleDocAffaire),#0) ;
CleDocAffaire.NaturePiece:= TOBPIECE.GetValue('GP_NATUREPIECEG');
CleDocAffaire.DatePiece:= TOBPIECE.GetValue('GP_DATEPIECE');
CleDocAffaire.Souche:= TOBPIECE.GetValue('GP_SOUCHE');
CleDocAffaire.NumeroPiece:= TOBPIECE.GetValue('GP_NUMERO');
CleDocAffaire.Indice:= TOBPIECE.GetValue('GP_INDICEG');
Annee:=FormatDateTime('yy',CleDocAffaire.DatePiece) ;

// création TOB Affaire
TOBAffaire:=TOB.Create('AFFAIRE',Nil,-1) ;
RemplirTOBAffaire(CodeAffaire,TobAffaire);
TOBAffaire2:=TOB.Create('AFFAIRE',Nil,-1);
TOBAffaire2.Dupliquer (TOBAffaire, True, true);
ReinitEltAffaire (TOBAffaire2);
if CodeAffaire = '' then
begin
	TOBAffaire.PutValue('AFF_GENERAUTO', 'DIR');
	TOBAffaire2.PutValue('AFF_GENERAUTO', 'DIR');
end;
// Récupération code sous-affaire du devis
Req := 'SELECT GP_NUMERO, GP_AFFAIREDEVIS,GP_REFINTERNE FROM PIECE WHERE GP_AFFAIRE = "'+ CodeAffaire +
         '" AND GP_NATUREPIECEG = "' + CleDocAffaire.NaturePiece + '" AND GP_NUMERO = '+ IntToStr(CleDocAffaire.NumeroPiece);
Q := OpenSQL(Req, True,-1,'',true);
if Not Q.EOF then
    begin
    Codessaff:=Q.FindField('GP_AFFAIREDEVIS').AsString;
    end else
    begin
    Codessaff := '';
    end;

Ferme(Q);
Numpiece:= CleDocAffaire.NumeroPiece;

if NbPiece = 2 then
    begin
    if (Action = taCreat) or (Duplic = True) then
       begin
       Req := 'SELECT GP_NUMERO, GP_AFFAIREDEVIS,GP_REFINTERNE FROM PIECE WHERE GP_AFFAIRE = "'+ CodeAffaire +
           '" AND GP_NATUREPIECEG = "' + CleDocAffaire.NaturePiece + '" AND GP_NUMERO <> '+ IntToStr(CleDocAffaire.NumeroPiece);
       Q := OpenSQL(Req, True,-1,'',true);
       if Not Q.EOF then
          begin
          Numpiece:= Q.FindField('GP_NUMERO').AsInteger;
          Codessaff:=Q.FindField('GP_AFFAIREDEVIS').AsString;
          end else
          begin
          Numpiece := 0;
          Codessaff := '';
          end;
       Ferme(Q);
       end;
    // En création d'avenant, on met à jour le code sous-affaire
    // avec le nouveau code affaire créé
    if (CodeAffaireAvenant <> '00') then
       begin
       CodeAffaireAvenant := codessaff;
       end;

    end;  // fin du traitement particulier pour le deuxième devis

// BRL FQ10118 6/10/03
// maj affaire multipièces
if  (TobPiece.Getvalue ('GP_NATUREPIECEG') = VH_GC.AFNatAffaire) and (CodeAffaire <> '') then
    begin
    if NbPiece > 1 then MultiPiece:='X'
    else                MultiPiece:='-';
    Req := 'UPDATE AFFAIRE SET AFF_MULTIPIECES = "'+MultiPiece+'" WHERE AFF_AFFAIRE='+'"'+CodeAffaire+'"';
    ExecuteSQL(Req);
    end;

GetMontantsAcomptes (TOBAcomptes,Acompte,regl);

// Traitement de création d'un nouveau devis
  if CledocAffaire.Naturepiece = 'DAP' then
    begin
    Codessaff := CodeAffaire;
    Req := 'UPDATE AFFAIRE SET '+
           'AFF_DATEMODIF ="'+ USDATETIME(NowH) +'"'+
           'WHERE AFF_AFFAIRE ="'+Codessaff+'"';
    ExecuteSQL(Req);
  end else if (Action = taCreat) or (Duplic = True) then {or ((Action = taModif) and (V_PGI.SAV)) : pour recréer fiche affaire-devis cf pb partage beneteau}
  begin
      // création sous-affaire pour nouveau devis
      if (CodeAffaireAvenant = '00') then
         begin // cas d'un devis initial
         Codessaff := Format ('%.8d',[CleDocAffaire.NumeroPiece]);
         Codessaff := 'Z'+CleDocAffaire.NaturePiece+CledocAffaire.Souche+Codessaff+'00';
         end else
         begin // cas d'un avenant
         Codessaff := Copy(CodeAffaireAvenant,2,14);
         Avenant := '01';
         Part0 := 'Z'+Codessaff+Avenant;
        // boucle pour verifier si la sous-affaire existe déjà
         while (existeaffaire (Part0,'') = true) Do
             begin
             i := StrToInt (Avenant);
             Inc(i);
             Avenant := Format ('%.2d',[i]);
             Part0 := 'Z'+Codessaff+Avenant;
             end;
         Codessaff := Part0;
         end;

      BTPCodeAffaireDecoupe(Codessaff, Part0, Part1, Part2, Part3,Avenant,taConsult,False);
      TOBAffaire2.PutValue('AFF_AFFAIRE', Codessaff);
      TOBAffaire2.PutValue('AFF_AFFAIRE0', 'Z');
      TOBAffaire2.PutValue('AFF_AFFAIRE1', Part1);
      TOBAffaire2.PutValue('AFF_AFFAIRE2', Part2);
      TOBAffaire2.PutValue('AFF_AFFAIRE3', Part3);
      TOBAffaire2.PutValue('AFF_AVENANT', Avenant);
      TOBAffaire2.PutValue('AFF_AFFAIREINIT', CodeAffaire);
      TOBAffaire2.PutValue('AFF_AFFAIREREF', CodeAffaire);
      TOBAffaire2.PutValue('AFF_TIERS', TOBPIECE.GetValue('GP_TIERS'));
      TOBAffaire2.PutValue('AFF_LIBELLE', TOBPIECE.GetValue('GP_REFINTERNE'));
      if ModeTrait = 'ETUTODEV' then TOBAffaire2.PutValue('AFF_ETATAFFAIRE','ACP');

      if CodeEtat <> '' then
        TOBAffaire2.PutValue('AFF_ETATAFFAIRE', CodeEtat);

      if CleDocAffaire.NaturePiece = 'DE' then TOBAffaire2.PutValue('AFF_ETATAFFAIRE','ENC');
      // AJOUT LS POUR GESTION LIGNE A ZERO
      TOBAffaire2.PutValue('AFF_OKSIZERO', TOBPIECE.GetValue('AFF_OKSIZERO'));
      // AJOUT BRL POUR TYPE DE FACTURATION
      TOBAffaire2.PutValue('AFF_GENERAUTO', TOBPIECE.GetValue('AFF_GENERAUTO'));
      // --
      if Acompte <> 0 then TOBAffaire2.PutValue('AFF_ACOMPTE',Acompte);
      TOBAffaire2.InsertDB(Nil);
    end else
    begin
      // Traitement de modification d'un devis
      // maj dans sous-affaire
    Req := 'UPDATE AFFAIRE SET '+
           'AFF_OKSIZERO="'+TOBPiece.getValue('AFF_OKSIZERO')+'",'+
           'AFF_GENERAUTO ='+'"'+ TOBPiece.getValue('AFF_GENERAUTO') +'",'+
           'AFF_AFFAIREREF ='+'"'+ CodeAffaire +'",'+
           'AFF_AFFAIREINIT ='+'"'+ CodeAffaire +'",'+
           'AFF_ACOMPTE='+StringReplace(FloatToStr(acompte),',','.',[rfReplaceAll])+' '+
           'WHERE AFF_AFFAIRE ='+'"'+Codessaff+'"';
      ExecuteSQL(Req);
    end;

  // Maj code sous-affaire dans pièce
Req := 'UPDATE PIECE SET GP_AFFAIREDEVIS ='+'"'+ Codessaff+'" WHERE '+WherePiece(CleDocAffaire,ttdPiece,False) ;
ExecuteSQL(Req);
TOBPiece.PutValue('GP_AFFAIREDEVIS', Codessaff);

  // suppression des TOB Affaire
TOBAffaire.free;
TOBAffaire := Nil;
TOBAffaire2.free;
TOBAffaire2 := Nil;
END;

function MajMontantAcompte(TOBPiece,TOBAcomptes : TOB) : boolean;
var TOBAffaire : TOB;
    Req : string;
    Q : TQuery;
    Acompte,Regl : double;
begin
result := true;
if (TOBPiece.GetValue ('GP_AFFAIREDEVIS') <> '') and (TOBPiece.GetValue('GP_NATUREPIECEG')=GetParamSoc('SO_AFNATAFFAIRE')) then
   begin
   result := false;
   Req := 'SELECT AFF_AFFAIRE,AFF_ACOMPTE FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBPiece.GetValue('GP_AFFAIREDEVIS')+'"' ;
   Q := OpenSql (Req,true,-1,'',true);
   if not Q.eof then
      begin
      TobAffaire := TOB.Create ('AFFAIRE',nil,-1);
      TOBAffaire.selectdb ('',Q);
      ferme (Q);
      GetMontantsAcomptes (TOBAcomptes,Acompte,regl);
      if Acompte > 0 then
         begin
         TOBAffaire.PutValue('AFF_ACOMPTE',Acompte);
         if TOBAffaire.updateDb (true) then result := true;
         end else Result := true;
      TOBAffaire.free;
      end else Ferme(Q);
   end;
end;

procedure MajAffaireApresGeneration (TobPiece,TOBPieceRG,TOBBasesRG,TOBAcomptes : TOB;DEV: Rdevise; FinTravaux : boolean=false; NumSituation : integer=-1);
var Req : String;
    Q : Tquery;
    TOBAffaire : TOB;
begin
  if (TobPiece <> NIL) and ((TOBPIECE.GetValue('GP_NATUREPIECEG') = VH_GC.AFNatAffaire) or (TOBPIECE.GetValue('GP_NATUREPIECEG') =VH_GC.AFNatProposition)) then Exit;

  if (not FinTravaux) then MajApresGeneration(TobPiece,TOBPieceRG,TOBBasesRG,TOBAcomptes,DEV,NumSituation);

  if TobPiece = nil then exit;
  if TOBPiece.GetValue('GP_AFFAIREDEVIS') = '' then exit;
  TobAffaire := TOB.Create ('AFFAIRE',nil,-1);
  Req := 'SELECT AFF_AFFAIRE,AFF_ACOMPTEREND FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBPiece.GetValue('GP_AFFAIREDEVIS')+'"' ;
  Q := OpenSql (Req,true,-1,'',true);
  TRY
    if not Q.eof then
    begin
      TOBAffaire.selectdb ('',Q);
      TOBAffaire.PutValue('AFF_ACOMPTEREND',TOBAffaire.GetValue('AFF_ACOMPTEREND')+TOBPiece.GetValue('GP_ACOMPTE'));
      TOBAffaire.updateDb (true);
    end;
  FINALLY
    TOBAffaire.free;
  ferme (Q);
  END;
end;

// fonction de mise à jour du type de facturation dans l'affaire liée au devis
Procedure  MajTypeFact (TobPiece : TOB; TypFac : string);
Var Req, Codessaff : String;
    TobLocal : TOB;
BEGIN
  // Maj dans la TOB pièce (utile pour pour maj en validation de pièce : cf majsousaffaire)
  TOBPiece.PutValue('AFF_GENERAUTO', TypFac);

  // Maj dans la TOB des affaires (utile pour l'affichage en complément document)
  Codessaff := TOBPiece.GetValue('GP_AFFAIREDEVIS');
  if Codessaff = '' then exit;
  TOBlocal := TOBlesAffaires.findfirst(['AFF_AFFAIRE'],[Codessaff],true);
  if TOBlocal <> nil then
  begin
    TOBlocal.putvalue('AFF_GENERAUTO', TypFac);
  end;
end;

// fonction retournant le type de facturation lié au devis
Function  RenvoieTypeFact (Codessaff : string) : String ;
Var Q : TQuery ;
    Req, TypFac : String;
    TobLocal : TOB;
BEGIN
  if Codessaff = '' then BEGIN result := 'DIR'; exit; end;
  //
  TOBlocal := TOBlesAffaires.findfirst(['AFF_AFFAIRE'],[Codessaff],true);
  if TOBlocal = nil then
  begin
    //récupération du type de facturation pour l'affaire liée au devis
    Req:='SELECT AFF_GENERAUTO FROM AFFAIRE WHERE AFF_AFFAIRE="'+Codessaff+'"' ;
    Q:=OpenSQL(Req,TRUE,-1,'',true) ;
    if Not Q.EOF then TypFac:=Q.Fields[0].AsString
                 else TypFac := 'DIR';
    Ferme(Q) ;
    result := TypFac;
  end else
  begin
    result := TOBlocal.getValue('AFF_GENERAUTO');
  end;

end;

Function  RenvoieInfosAffaireDevis (Codessaff : string) : StInfosAffaire ;
Var Q : TQuery ;
    Req : string ;
    TheResult : stInfosAffaire;
BEGIN

//récupération du type de facturation pour l'affaire liée au devis
Req:='SELECT AFF_OKSIZERO,AFF_GENERAUTO FROM AFFAIRE WHERE AFF_AFFAIRE="'+Codessaff+'"' ;
Q:=OpenSQL(Req,TRUE,-1,'',true) ;
if Not Q.EOF then
begin
	TheResult.TypFac:=Q.Fields[1].AsString;
  TheResult.ModeFact := (Q.Fields[0].AsString='X');
end else
begin
	TheResult.TypFac := 'DIR';
  TheResult.ModeFact := false;
end;
Ferme(Q) ;

result := TheResult;

end;
// fonction retournant le code regroupement sur facture d'une affaire
Function  RenvoieCodeReg (codeaffaire : string) : String ;
Var Q : TQuery ;
    Req, CodReg : String;
BEGIN

Req:='SELECT AFF_REGROUPEFACT FROM AFFAIRE WHERE AFF_AFFAIRE="'+Codeaffaire+'"' ;
Q:=OpenSQL(Req,TRUE,-1,'',true);
if Not Q.EOF then CodReg:=Q.Fields[0].AsString
else CodReg := 'AUC';
Ferme(Q) ;

result := CodReg;

end;

// fonction retournant le N° de devis depuis le champ affairedevis
Function  RenvoieNumPiece (codeaffaire : string) : String ;
Var Q : TQuery ;
    Req : String;
    NumPiece : Integer;
BEGIN

Req:='SELECT GP_NUMERO FROM PIECE WHERE GP_AFFAIREDEVIS="'+Codeaffaire+'" AND GP_NATUREPIECEG="DBT"' ;
Q:=OpenSQL(Req,TRUE,-1,'',true);
if Not Q.EOF then NumPiece:=Q.Fields[0].AsInteger
else NumPiece := 0;
Ferme(Q) ;

result := IntToStr(NumPiece);

end;

// fonction retournant le type de ressource d'une prestation
Function  RenvoieTypeRes (codeprestation : string) : String ;
Var Q : TQuery ;
    Req, TypeRes : String;
BEGIN
TypeRes:='';
Req:='SELECT BNP_TYPERESSOURCE FROM NATUREPREST WHERE BNP_NATUREPRES=(';
Req:=Req+'SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE="'+Codeprestation+'")';
Q:=OpenSQL(Req,TRUE,-1,'',true);
if Not Q.EOF then TypeRes:=Q.Fields[0].AsString;
Ferme(Q) ;

result := TypeRes;

end;

Procedure MajQtefacture(TobPiece,TobPiece_ori : Tob);
Var i : integer;
    TOBL : TOB;
    TypeFacturation : string;
Begin
  // Après génération de facture,
  // Maj dans le devis de la qté réellement facturée (en cumul pour les situations)
  if (Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC;')=0) then exit;
  for i:=0 to TOBPiece_ori.Detail.Count-1 do
  begin
    TOBL:=TOBPiece_ori.Detail[i] ;
    // Correction fiche qualité 11545 par LS
    // if (TOBL.GetValue('GL_REFARTSAISIE') <> '') then
    if (TOBL.GetValue('GL_ARTICLE') <> '') then
    begin
      TypeFacturation := RenvoieTypeFact(TOBPiece_ori.GetValue('GP_AFFAIREDEVIS'));
      if (TypeFacturation = 'AVA') or (TypeFacturation = 'DAC') then
      begin
        // en situation, stockage qté cumulée
        TOBL.PutValue('GL_QTESIT',TOBL.GetValue('GL_QTEPREVAVANC'));
      end else
      begin
        // en facture directe, cumul + stockage qté facturée
        TOBL.PutValue('GL_QTESIT',TOBL.GetValue('GL_QTESIT')+TOBL.GetValue('GL_QTEPREVAVANC')) ;
        TOBL.PutValue('GL_QTEPREVAVANC',0) ;
        TOBL.PutValue('GL_POURCENTAVANC',0) ;
      end;
      // MODIFBRL 13/11
      TOBL.PutValue('GL_QTERESTE',TOBL.GetValue('GL_QTEFACT')); {newpiece}
      //--- GUINIER ---
      if CtrlOkReliquat(TOBL, 'GL') then TOBL.PutValue('GL_MTRESTE', TOBL.GetValue('GL_MONTANTHTDEV')); {newpiece}
    end;
  end;
  TOBPiece_ori.UpdateDB (False);
End;

procedure GetMontantsAcomptes (TOBAcomptes : TOB;var MontantAcompte : double ; var MontantRegl : double);
var Indice : integer;
begin
  MontantAcompte := 0;
  MontantRegl := 0;
  if TOBAcomptes = nil then Exit;
  if TOBAcomptes.detail.count = 0  then Exit;
  for Indice := 0 to TOBAcomptes.detail.count -1 do
    begin
    if TOBAcomptes.detail[Indice].getValue('GAC_ISREGLEMENT')='X' then
       begin
       MontantRegl := MontantRegl + TOBAcomptes.detail[Indice].getValue('GAC_MONTANTDEV');
       end else
       begin
       MontantAcompte := MontantAcompte + TOBAcomptes.detail[Indice].getValue('GAC_MONTANTDEV');
       end;
  //			MontantAcompte := MontantAcompte + TOBAcomptes.detail[Indice].getValue('GAC_MONTANTDEV');
    end;
end;

Procedure MajApresGeneration(TobPiece,TOBPieceRG,TOBBasesRG ,TOBAcomptes: Tob; DEV : Rdevise;NumSituation : Integer=-1);
Var Req: string;
    numsit,IndiceSit : integer;
    tottva : double;
    Q : TQuery ;
    TTC,XP,XD,XE,TXD,TXP,TXE : Double;
    TypeFacturation : string;
    MontantRegl,MontantAcompte : double;
    Okok : Boolean;
Begin
  // Après génération de facture,
  // Si situation, Maj table Situations
//  if (GetInfoParPiece(TOBPiece.getString('GP_NATUREPIECEG'),'GPP_ESTAVOIR')='X') then exit;
  if TOBPieceRG = nil then Exit;
  TypeFacturation := RenvoieTypeFact (TOBPiece.GetValue('GP_AFFAIREDEVIS'));
  if (TypeFacturation = 'AVA') or (TypeFacturation = 'DAC') then
  begin
    if NumSituation = -1 then
    begin
  Req:='SELECT BST_NUMEROSIT FROM BSITUATIONS WHERE '+
       'BST_SSAFFAIRE="'+ TOBPiece.GetValue('GP_AFFAIREDEVIS') + '" '+
          ' AND BST_VIVANTE="X" '+
       'ORDER BY BST_SSAFFAIRE,BST_NUMEROSIT DESC';
  Q:=OpenSQL(Req,TRUE,-1,'',true);
  if Q.EOF then
    numsit := 0
  else
    begin
        //inutile et pb pour eagl    Q.Findfirst;
    numsit := Q.Fields[0].AsInteger;
    end;
  Ferme(Q) ;
      Inc(numsit); IndiceSit := 0;
      //
      Req:='SELECT BST_INDICESIT,BST_VIVANTE FROM BSITUATIONS WHERE '+
          'BST_SSAFFAIRE="'+ TOBPiece.GetValue('GP_AFFAIREDEVIS') + '" AND '+
          'BST_NUMEROSIT='+ IntToStr(NumSit) + ' '+
          'ORDER BY BST_SSAFFAIRE,BST_NUMEROSIT,BST_INDICESIT DESC';
      Q:=OpenSQL(Req,TRUE,-1,'',true);
      if not Q.eof then
      begin
        IndiceSit := Q.Fields[0].AsInteger;
        Inc(IndiceSit);
      end;
      ferme (Q);
    end else
    begin
      numsit := NumSituation; IndiceSit := 0;
      if POs(TobPiece.getString('GP_NATUREPIECEG'),'ABT;ABP')<0 then
      begin
      Req:='SELECT BST_INDICESIT,BST_VIVANTE FROM BSITUATIONS WHERE '+
          'BST_SSAFFAIRE="'+ TOBPiece.GetValue('GP_AFFAIREDEVIS') + '" AND '+
            'BST_NATUREPIECEG='+ TOBPiece.GetValue('GP_NATUREPIECEG') + ' '+
          'BST_NUMEROSIT='+ IntToStr(NumSituation) + ' '+
          'ORDER BY BST_SSAFFAIRE,BST_NUMEROSIT,BST_INDICESIT DESC';
      Q:=OpenSQL(Req,TRUE,-1,'',true);
      if not Q.eof then
      begin
        IndiceSit := Q.Fields[0].AsInteger;
        Inc(IndiceSit);
    end;
      ferme (Q);
    end;
    end;
  // Maj table situations
  GetMontantsAcomptes (TOBAcomptes,MontantAcompte,MontantRegl);
  GetMontantRG (TOBPieceRG,TOBBasesRG,XD,XP,DEV,True);
  GetcumultaxesRG (TOBBasesRG,TOBPieceRG,TXD,TXP,DEV);
  TTC := TOBPiece.GetValue('GP_TOTALTTCDEV') - XD - TXD;
  tottva := TTC - TOBPiece.GetValue('GP_TOTALHTDEV');
  Req := 'INSERT INTO BSITUATIONS '+
      '(BST_NATUREPIECE,BST_SOUCHE,BST_NUMEROFAC,BST_NUMEROSIT,BST_INDICESIT,BST_DATESIT,'+
      'BST_AFFAIRE,BST_SSAFFAIRE,BST_MONTANTHT,BST_MONTANTTVA,BST_MONTANTTTC,BST_MONTANTACOMPTE,'+
      'BST_MONTANTREGL,BST_VIVANTE) '+
         ' VALUES '+
         '("'+
         TOBPiece.GetValue('GP_NATUREPIECEG')+
         '","'+
         TOBPiece.GetValue('GP_SOUCHE')+
         '",'+
         IntToStr(TOBPiece.GetValue('GP_NUMERO'))+
         ','+
         IntToStr(numsit)+
      ','+
      IntToStr(IndiceSit)+
         ',"'+
         usdatetime(StrToDate(TOBPiece.GetValue('GP_DATEPIECE')))+
         '","'+
         TOBPiece.GetValue('GP_AFFAIRE')+
         '","'+
         TOBPiece.GetValue('GP_AFFAIREDEVIS')+
         '",'+
         StringReplace(FloatToStr(TOBPiece.GetValue('GP_TOTALHTDEV')),',','.',[rfReplaceAll])+
         ','+
         StringReplace(FloatToStr(tottva),',','.',[rfReplaceAll])+
         ','+
         StringReplace(FloatToStr(TTC),',','.',[rfReplaceAll])+
         ','+
         StringReplace(FloatToStr(MontantAcompte),',','.',[rfReplaceAll])+
         ','+
         StringReplace(FloatToStr(MontantRegl),',','.',[rfReplaceAll])+
      ',"X")';
  ExecuteSQL(Req);
  end;
End;

procedure MajLignesFacOuvr (TOBL,TOBOuvrage : TOB);
var TOBO ,TOBOO, TOBODF : TOB;
    II,Indicenomen : integer;
    QQ : Tquery;
    CD : R_CleDoc ;
    Req,refPiece : string;
    Pourcent : double;
    created : boolean;
begin
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;

  TOBOO := TOBOuvrage.detail[IndiceNomen-1];
  TOBODF := TOB.Create('LIGNEFAC',nil,-1);

  for II := 0 TO TOBOO.detail.count -1 do
  begin
    TOBO := TOBOO.detail[II];
    RefPiece := EncodeRefPieceOuv(TOBO);
    DecodeRefPieceOUv(RefPiece,CD);
    created := false;
    Req := 'SELECT * FROM LIGNEFAC WHERE '+ WherePiece (CD,ttdLigneFac,false)+' AND BLF_NUMORDRE=0 AND BLF_UNIQUEBLO='+IntToStr(CD.UniqueBlo);
    QQ:= OpenSql (Req,true,-1,'',true);
    if not QQ.eof then
    begin
      TOBODF.SelectDB('',QQ);
    end else
    begin
      TOBODF.putValue('BLF_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG'));
      TOBODF.putValue('BLF_SOUCHE',TOBL.GetValue('GL_SOUCHE'));
      TOBODF.putValue('BLF_DATEPIECE',TOBL.GetValue('GL_DATEPIECE'));
      TOBODF.putValue('BLF_NUMERO',TOBL.GetValue('GL_NUMERO'));
      TOBODF.putValue('BLF_INDICEG',TOBL.GetValue('GL_INDICEG'));
      TOBODF.putValue('BLF_NUMORDRE',0);
      TOBODF.putValue('BLF_UNIQUEBLO',TOBO.GetValue('BLO_UNIQUEBLO'));
      TOBODF.putValue('BLF_MTMARCHE',TOBO.GetValue('BLF_MTMARCHE'));
      TOBODF.putValue('BLF_MTDEJAFACT',TOBO.GetValue('BLF_MTDEJAFACT'));
      TOBODF.putValue('BLF_QTEMARCHE',TOBO.GetValue('BLF_QTEMARCHE'));
      TOBODF.putValue('BLF_QTEDEJAFACT',TOBO.GetValue('BLF_QTEDEJAFACT'));
      created := true;
    end;
    //
    TOBODF.putValue('BLF_MTCUMULEFACT',TOBO.GetValue('BLF_MTCUMULEFACT'));
    TOBODF.putValue('BLF_MTSITUATION',TOBO.GetValue('BLF_MTSITUATION'));
    TOBODF.putValue('BLF_QTECUMULEFACT',TOBO.GetValue('BLF_QTECUMULEFACT'));
    TOBODF.putValue('BLF_QTESITUATION',TOBO.GetValue('BLF_QTESITUATION'));
    TOBODF.putValue('BLF_POURCENTAVANC',TOBO.GetValue('BLF_POURCENTAVANC'));
    TOBODF.putValue('BLF_POURCENTAVANC',TOBO.GetValue('BLF_POURCENTAVANC'));
    TOBODF.putValue('BLF_MTPRODUCTION',0);
    TOBODF.SetAllModifie (true);
    //
    if TOBODF.GetValue('BLF_MTSITUATION') = 0 then
    begin
      if not Created then TOBODF.DeleteDB(false); // le montant de situation est devenu nul alors qu'il existait préalablement
    end else
    begin
      if Created then
      begin
        if (TOBODF.GetValue ('BLF_MTSITUATION')<> 0) or (TOBODF.GetValue ('BLF_MTMARCHE')<> 0) then
        begin
          TOBODF.setAllModifie(true);
          TOBODF.InsertDB(nil,false) // situation positionné alors qu'elle n'existait pas avant
        end;
      end else TOBODF.UpdateDB(false);    // mise à jour
    end;
  end;

  TOBODF.free;

end;

procedure ValideSituation (TOBPiece,TOBOuvrage : TOB );
var TOBLIGNEFAC,TOBLN,TOBL : TOB;
    NaturePiece : string;
    i : integer;
begin
  TOBLIGNEFAC := TOB.create ('LES LIGNES FACTURATION',nil,-1);

  NaturePiece := TOBPiece.GetString('GP_NATUREPIECEG');

  TRY
    if (Pos(NaturePiece,'FBT;DAC;FBP;BAC;')>0) and (TOBPiece.GetValue('GP_AFFAIREDEVIS')<>'') then
    begin
      for i:=0 to TOBPiece.Detail.Count-1 do
      begin
        TOBL:=TOBPiece.Detail[i] ;
        if IsSousDetail(TOBL) then continue;
        if (TOBL.getValue('BLF_MTMARCHE') <> 0) or (TOBL.getValue('BLF_MTSITUATION')<>0) then
        begin
          TOBLN := TOB.Create ('LIGNEFAC',TOBLIGNEFAC,-1);
          SetLigneFacture (TOBL,TOBLN);
          if IsOuvrage(TOBL) then MajLignesFacOuvr(TOBL,TOBOUvrage);
        end;
      end;
    end;
    if TOBLIGNEFAC.detail.count > 0 then TOBLIGNEFAC.InsertDB (nil);
  FINALLY
    TOBLIGNEFAC.Free;
  END;

end;

//FV1 ! 01/12/2015 - FS#355 - DELABOUDINIERE : En impression situation, ajouter les avenants
procedure ChargePieceRepart(AffaireDevis : String; TobLigneFac : TOB);
var StSQl         : string;
    QQ            : TQuery;
    ind           : Integer;
    Ind_2         : Integer;
    Ind_3         : Integer;
    TOBMemo       : Tob;
    TobPieceG     : Tob;
    TobRepartFac  : Tob;
    TobLRepartFac : Tob;
    TOBL          : Tob;
    AffDevis      : string;
    Devis         : String;
    PiecePrec     : String;
    Nature        : string;
    souche        : string;
    DatePiece     : string;
    Numero        : Integer;
    Avenant       : Integer;
    Indice        : Integer;
    Budget        : Double;
    MtCumulFact   : Double;
    MtDejaFact    : Double;
    MtSituation   : Double;
begin

  if AffaireDevis= '' then Exit;

  if TobLigneFac = nil then Exit;

  //on décompose le GP_AffaireDevis
  Nature        := copy(AffaireDevis,2,3);
  souche        := copy(AffaireDevis,5,3);
  Numero        := StrToInt(Trim(copy(AffaireDevis,8,8)));
  Avenant       := StrToInt(copy(AffaireDevis,16,2));
  AffaireDevis  := Nature + ';' + Souche + ';' + IntToStr(Numero) + ';' + IntToStr(Avenant);

  //Lecture du dossier de facturation pour création de l'enreg et mise à jour du budget
  StSQl := 'SELECT * FROM BTMEMOFACTURE WHERE BMF_DEVISPRINC="' + AffaireDevis + ';"';
  QQ    := OpenSQL(StSQl, False);

  TOBMemo := TOB.Create('Dossier Facture',nil, -1);
  TobMemo.LoadDetailDb('BTMEMOFACTURE','','', QQ, True);

  Ferme(QQ);

  For ind := 0 to TOBMemo.Detail.Count -1 do
  begin
    //
    TOBL := TOBMemo.detail[ind];
    //
    TOBL.AddChampSupValeur('REPART', '');
    TOBL.AddChampSupValeur('BUDGET', 0.00);
    //
    TOBL.AddChampSupValeur('NATURE', '');
    TOBL.AddChampSupValeur('SOUCHE', '');
    TOBL.AddChampSupValeur('NUMERO', 0);
    TOBL.AddChampSupValeur('INDICEG',0);
    //
    MtCumulFact := 0;
    MtSituation := 0;
    MtDejaFact  := 0;
    //
    //lecture de la pièce pour récupération des N° Avenant...
    TraitementPiecebyMemoFact(TobL);
    //
  end;

  RecherchePieceDansLigneFac(TobMemo, TobLigneFac);

  FreeAndNil(TobMemo);

end;

Procedure TraitementPiecebyMemoFact(TOBMemo : Tob);
var TobPieceG : TOB;
    UnePiece  : TOB;
    Devis     : string;
    Nature    : string;
    Souche    : string;
    StSql     : string;
    AffDevis  : string;
    Numero    : String;
    Avenant   : String;
    Budget    : Double;
    QQ        : TQuery;
begin

  Devis     := TobMemo.GetString('BMF_DEVIS');
  //
  Nature    := READTOKENST(Devis);
  souche    := READTOKENST(Devis);
  Numero    := READTOKENST(Devis);
  Avenant   := READTOKENST(Devis);
  //
  TOBMemo.PutValue('NATURE', Nature);
  TOBMemo.PutValue('SOUCHE', Souche);
  TOBMemo.PutValue('NUMERO', Numero);
  TOBMemo.PutValue('INDICEG',Avenant);

  StSQl := 'SELECT GP_AFFAIREDEVIS, GP_TOTALHTDEV FROM PIECE ';
  StSQl := StSQl + ' WHERE GP_NATUREPIECEG="' + Nature;
  StSQL := StSQL + '"  AND GP_SOUCHE="'       + souche;
  StSQL := StSQL + '"  AND GP_NUMERO='        + Numero;

  QQ    := OpenSQL(StSQl, False);

  if not QQ.eof then
  begin
    //chargement du devis initial et/ou avenant
    AffDevis  := QQ.Findfield('GP_AFFAIREDEVIS').AsString;;
    Nature    := copy(AffDevis, 2,3);
    souche    := copy(AffDevis, 5,3);
    Numero    := copy(AffDevis, 8,8);
    Avenant   := copy(AffDevis,16,2);
    //
    Budget    := QQ.findfield('GP_TOTALHTDEV').AsFloat;
    //
    //on recherge la zone REPART du dossier de facturation avec l'avenant du Devis trouvé
    //On charge la zone budget également.
    TOBMemo.PutValue('REPART', Avenant);
    TOBMemo.PutValue('BUDGET', Budget);
  end;

  ferme(QQ);

end;

Procedure RecherchePieceDansLigneFac(TOBMemo, TobLigneFac : TOB);
Var ind           : Integer;
    TOBLFac       : Tob;
    Tobl,TOBLMemo : Tob;
    TobRepartFac  : Tob;
    PiecePrec     : String;
    Nature        : string;
    souche        : string;
    DatePiece     : string;
    Numero,Indice : Integer;
    Avenant       : String;
    Ligne         : Integer;

    MtMarche      : double;
    Budget        : Double;
    MtCumulFact   : Double;
    MtDejaFact    : Double;
    MtSituation   : Double;
begin

  //
  TobRepartFac := TOB.Create('Les Repartitions', nil, -1);

  //lecture du LigneFac pour chargement des autres valeur...
  For Ind := 0 To TobLigneFac.detail.count -1 do
  begin
    TOBLFac   := TobLigneFac.detail[ind];

    //si nous somme sur un ouvrage on ne traite pas le premier enreg
    if TOBLFac.FieldExists('ISOUVRAGE') then continue;
    //
    PiecePrec := TOBLFac.GetValue('PIECEPRECEDENTE');
    //
    if PiecePrec = '' then exit;
    //
    DatePiece := ReadTokenSt(PiecePrec);
    Nature    := ReadTokenSt(PiecePrec);
    Souche    := ReadTokenSt(PiecePrec);
    Numero    := StrToInt(ReadTokenSt(PiecePrec));
    Indice    := StrToInt(ReadTokenSt(PiecePrec));
    //
    Ligne     := StrToInt(ReadTokenSt(PiecePrec)); //

    //chargement de la piece dans le dossier de facture à partir des élément du LigneFac
    TOBLMemo := TOBMemo.FindFirst(['NATURE','SOUCHE','NUMERO','INDICEG'],[Nature,souche,Numero,indice], True);
    if TOBLMemo <> Nil then
    begin
      Avenant := TOBLMemo.getString('REPART');
      //On cherche si la pièce n'est pas déjà dans la table des répartitions
      TOBL := TobRepartFac.FindFirst(['BPR_NATUREPIECEG','BPR_SOUCHE','BPR_NUMERO','BPR_AVENANT'],[TOBLFac.GetString('BLF_NATUREPIECEG'),TOBLFac.GetString('BLF_SOUCHE'),TOBLFac.GetString('BLF_NUMERO'),Avenant],True);
      if TOBL = nil then
      begin
        //création de la Tob de répartition
        TOBL := TOB.Create('BTPIECEREPART', TobRepartFac, -1);
        TOBL.PutValue('BPR_NATUREPIECEG',   TOBLFac.GetString('BLF_NATUREPIECEG'));
        TOBL.PutValue('BPR_SOUCHE',         TOBLFac.GetString('BLF_SOUCHE'));
        TOBL.PutValue('BPR_NUMERO',         TOBLFac.GetInteger('BLF_NUMERO'));
        TOBL.PutValue('BPR_INDICEG',        TOBLFac.GetInteger('BLF_INDICEG'));
        TOBL.PutValue('BPR_AVENANT',        Avenant);
        TOBL.PutValue('BPR_BUDGET',         TOBLMemo.GetDouble('BUDGET'));
        TOBL.SetDouble('BPR_CUMULFACT',     0);
        TOBL.PutValue('BPR_SITPRECEDENTE',  0);
        TOBL.PutValue('BPR_SITENCOURS',     0);
      end;
      //chargement de repartFac
      TOBL.SetDouble('BPR_CUMULFACT',     TOBL.GetDouble('BPR_CUMULFACT')     + TOBLFac.GetDouble('BLF_MTCUMULEFACT'));
      TOBL.SetDouble('BPR_SITPRECEDENTE', TOBL.GetDouble('BPR_SITPRECEDENTE') + TOBLFac.GetDouble('BLF_MTDEJAFACT'));
      TOBL.SetDouble('BPR_SITENCOURS',    TOBL.GetDouble('BPR_SITENCOURS')    + TOBLFac.GetDouble('BLF_MTSITUATION'));
    end;
    //
  end;

  //Ecriture de la table Repartition
  TobRepartFac.InsertDB(nil);
  FreeAndNil(TobRepartFac);

end;

procedure DestruitPrepaFact(Nature, Souche  : String;  Numero  : Integer);
Var StSQL   : string;
    nb      : integer;
begin

  if not GetParamSocSecur('SO_PIECEREPART', False) then exit;

  //il suffit simplement de lire al table et de supprimer ses enregistrements.
  If (Pos(Nature, 'FBT;DAC;FBP;BAC;'))=0 then
  begin
//    V_PGI.IoError := oeUnknown;
    Exit;
  end;

  //Requete de suppression
  StSQL := 'DELETE BTPIECEREPART WHERE BPR_NATUREPIECEG="' + Nature + '" ';
  StSQL := StSQL + 'AND BPR_SOUCHE="' + Souche + '" ';
  StSQL := StSQL + 'AND BPR_NUMERO= ' + IntToStr(Numero) + '';

  Nb := ExecuteSQL(StSQL);

  if Nb < 0 then V_PGI.IoError := oeUnknown;

end;

Procedure ModifSituation(TobPiece,TOBOuvrage,TobPieceRG,TOBPieceRg_O,TOBBasesRG,TOBAcomptes : Tob; DEV : RDevise; var DocArecalc : r_cledoc );
Var i               : integer;
    ind             : Integer;
    Ind_2           : Integer;
    NbSit : integer;
    NumPiecePrec    : Integer;
    //
    CD              : R_CleDoc ;
    //
    QQ              : TQuery;
    //
    Req             : string;
    StPrec          : string;
    StOrig          : string;
    QteSit          : string;
    QtePrevavanc    : string;
    PourcentSit     : string;
    NaturePiece     : string;
    TypeFacturation : string;
    DossierFac : string;
    //
    IsLastSituation : Boolean;
    Avancement      : Boolean;
    Situations      : Boolean;
    OkCont          : boolean;
    OKOK            : boolean;
    Repartition     : Boolean;
    //
    XP              : double;
    XD              : double;
    XE              : double;
    TXP             : double;
    TXD             : double;
    TXE             : double;
    Pourcent        : double;
    TTC             : double;
    MontantRegl     : Double;
    MontantAcompte  : double;
    PrevFact        : Double;
    CurrFact        : Double;
    DejaFact        : double;
    Marche,TVA : double;
    //
    TOBSITSUITE     : TOB;
    TOBLSITSUITE    : TOB;
    TOBL            : TOB;
    TOBSIT          : TOB;
    TOBLF           : TOB;
    TOBLIGNEFAC     : TOB;
    TOBLS           : TOB;
    TOBLD           : TOB;
    TOBLN           : TOB;
    TOBMEMFAC       : TOB;
    RefPieceSel : string;
  //
  //
  //
  Procedure CreateTobSituationSuivante(TobLigne, TOBLigneFact : TOB);
  Var StSQL         : string;
      QQ            : TQuery;
      Affairedevis  : string;
      Nature        : String;
      souche        : String;
      PiecePrec     : String;
      DatePiece     : string;
      Numero        : Integer;
      Indice        : Integer;
      TOBMemo       : TOB;
      TOBLMemo      : TOB;
      ind_3         : Integer;
      Avenant       : String;
  begin

    TOBLSITSUITE := Tob.create('LIGNEFAC', TOBSITSUITE, -1);
    //
    TOBLSITSUITE.Dupliquer(TOBLigneFact, false, true);
    //
    TOBLSITSUITE.PutValue('PIECEPRECEDENTE',  TobLigne.GetString('GL_PIECEPRECEDENTE'));
    //
    //Lecture de la piece avec le numéro de la suivante
    StSQL := 'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE ';
    StSQL := StSQL + 'GP_NATUREPIECEG="' + TOBLSITSUITE.GetValue('BLF_NATUREPIECEG') + '" ';
    StSQL := StSQL + 'AND GP_SOUCHE="'   + TOBLSITSUITE.GetValue('BLF_SOUCHE')           + '" ';
    StSQL := StSQL + 'AND GP_NUMERO='    + IntToStr(TOBLSITSUITE.GetValue('BLF_NUMERO')) + '  ';
    StSQL := StSQL + 'AND GP_INDICEG='   + IntToStr(TOBLSITSUITE.GetValue('BLF_INDICEG'))+ '  ';
    QQ := OpenSQL(StSQL,False);
    if not QQ.Eof then
    begin
      //on décompose le GP_AffaireDevis
      Affairedevis := QQ.findfield('GP_AFFAIREDEVIS').AsString;
      Nature        := copy(AffaireDevis,2,3);
      souche        := copy(AffaireDevis,5,3);
      Numero        := StrToInt(Trim(copy(AffaireDevis,8,8)));
      Indice        := StrToInt(copy(AffaireDevis,16,2));
      AffaireDevis  := Nature + ';' + Souche + ';' + IntToStr(Numero) + ';' + IntToStr(Indice);
      Ferme(QQ);

      //Lecture du dossier de facturation pour création de l'enreg et mise à jour du budget
      StSQl := 'SELECT * FROM BTMEMOFACTURE WHERE BMF_DEVISPRINC="' + AffaireDevis + ';"';
      QQ    := OpenSQL(StSQl, False);

      TOBMemo := TOB.Create('Dossier Facture',nil, -1);
      TobMemo.LoadDetailDb('BTMEMOFACTURE','','', QQ, True);

      Ferme(QQ);

      For ind_3 := 0 to TOBMemo.Detail.Count -1 do
      begin
        //
        TOBLMemo := TOBMemo.detail[ind_3];
        //
        TOBLMemo.AddChampSupValeur('REPART', '');
        TOBLMemo.AddChampSupValeur('BUDGET', 0.00);
        //
        TOBLMemo.AddChampSupValeur('NATURE', '');
        TOBLMemo.AddChampSupValeur('SOUCHE', '');
        TOBLMemo.AddChampSupValeur('NUMERO', 0);
        TOBLMemo.AddChampSupValeur('INDICEG',0);
        //lecture de la pièce pour récupération des N° Avenant...
        TraitementPiecebyMemoFact(TOBLMemo);
      end;
      //
      PiecePrec := TOBLSITSUITE.GetValue('PIECEPRECEDENTE');
      //
      if PiecePrec = '' then exit;
      //
      DatePiece := ReadTokenSt(PiecePrec);
      Nature    := ReadTokenSt(PiecePrec);
      Souche    := ReadTokenSt(PiecePrec);
      Numero    := StrToInt(ReadTokenSt(PiecePrec));
      Indice    := StrToInt(ReadTokenSt(PiecePrec));
      //
      //Ligne     := StrToInt(ReadTokenSt(PiecePrec));
      //
      //chargement de la piece dans le dossier de facture à partir des éléments du LigneFac
      TOBLMemo := TOBMemo.FindFirst(['NATURE','SOUCHE','NUMERO','INDICEG'],[Nature,souche,Numero,indice], True);
      if TOBLMemo <> Nil then
      begin
        Avenant := TOBLMemo.getString('REPART');
        //Requete de suppression des éléments de la Situation suivante
        StSQL := 'DELETE BTPIECEREPART WHERE ';
        StSQL := StSQL + '    BPR_NATUREPIECEG="'   + TOBLSITSUITE.GetValue('BLF_NATUREPIECEG')      + '" ';
        StSQL := StSQL + 'AND BPR_SOUCHE="'         + TOBLSITSUITE.GetValue('BLF_SOUCHE')            + '" ';
        StSQL := StSQL + 'AND BPR_NUMERO= '         + IntToStr(TOBLSITSUITE.GetValue('BLF_NUMERO'))  + ' ';
        StSQL := StSQL + 'AND BPR_INDICEG= '        + IntToStr(TOBLSITSUITE.GetValue('BLF_INDICEG')) + ' ';
        StSQL := StSQL + 'AND BPR_AVENANT= '        + Avenant           + '';
        //
        ExecuteSQL(StSQL);
      end;
      Ferme (QQ);
      //
      FreeAndNil(TOBMemo);
    end;
  end;


  {*****************************************************************
  Auteur  ...... :
  Créé le ...... : 10/12/2015
  Modifié le ... :   /  /
  Description .. : Mise à jour Ligne Devis Ouvrage
  *****************************************************************}
	procedure MajLigneDevisOuv ( TOBL,TOBOuvrage : TOB; Avancement : boolean);
  var TOBO ,TOBOO, TOBODD : TOB;
  		II,Indicenomen,NbSit : integer;
      QQ : Tquery;
    	CD : R_CleDoc ;
      Req : string;
      Pourcent : double;
      created : boolean;
      PrevFact,CurrFact,DejaFact : double;
  begin

    IndiceNomen := TOBL.getValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;

    TOBOO := TOBOuvrage.detail[IndiceNomen-1];

    TOBODD := TOB.Create ('LIGNEFAC',nil,-1);

    for II := 0 TO TOBOO.detail.count -1 do
    begin
       //
    	TOBO := TOBOO.detail[II];
      //
      if TOBO.getValue('BLO_PIECEPRECEDENTE')='' then continue;
      //
      DecodeRefPieceOUv(TOBO.getValue('BLO_PIECEPRECEDENTE'),CD) ;
      //
      Req := 'SELECT *, '' AS PIECEPRECEDENTE FROM LIGNEFAC WHERE '+ WherePiece (CD,ttdLigneFac,false)+' AND BLF_NUMORDRE=0 AND BLF_UNIQUEBLO='+IntToStr(CD.UniqueBlo);
      QQ:= OpenSql (Req,true,-1,'',true);
      created := false;
      if not QQ.eof then
      begin
        TOBODD.SelectDb ('',QQ);
      end
      else
      begin
      	TOBODD.InitValeurs(false);
        TOBODD.putValue('BLF_NATUREPIECEG',   CD.NaturePiece );
        TOBODD.putValue('BLF_SOUCHE',         CD.Souche );
        TOBODD.putValue('BLF_DATEPIECE',      CD.DatePiece);
        TOBODD.putValue('BLF_NUMERO',         CD.NumeroPiece );
        TOBODD.putValue('BLF_INDICEG',        CD.Indice );
        TOBODD.putValue('BLF_NUMORDRE',       0);
        TOBODD.putValue('BLF_UNIQUEBLO',      CD.UniqueBlo);
        TOBODD.putValue('BLF_MTMARCHE',       TOBO.GetValue('BLF_MTMARCHE'));
        TOBODD.putValue('BLF_QTEMARCHE',      TOBO.GetValue('BLF_QTEMARCHE'));
        created := true;
      end;
      ferme (QQ);

      if Avancement then
      begin
        TOBODD.PutValue('BLF_QTEDEJAFACT',    TOBO.GetValue('BLF_QTECUMULEFACT'));
        TOBODD.PutValue('BLF_QTECUMULEFACT',  TOBO.GetValue('BLF_QTECUMULEFACT'));
      //
        TOBODD.PutValue('BLF_MTDEJAFACT',     TOBO.GetValue('BLF_MTCUMULEFACT'));
        TOBODD.PutValue('BLF_MTCUMULEFACT',   TOBO.GetValue('BLF_MTCUMULEFACT'));
      //
        if TOBODD.GetValue('BLF_MTMARCHE') <> 0 then Pourcent := arrondi((TOBODD.GetValue('BLF_MTCUMULEFACT') / TOBODD.GetValue('BLF_MTMARCHE')) * 100,2)
                                                else Pourcent := 0;
        TOBODD.PutValue('BLF_POURCENTAVANC',  Pourcent);
      if TOBODD.getvalue('BLF_MTCUMULEFACT')=0 then
      begin
      	if not created then
        begin
      		TOBODD.SetAllModifie(true);
        	TOBODD.DeleteDB(false);
        end;
      end else
      begin
      	if created then
        begin
          if (TOBODD.GetValue('BLF_MTMARCHE')<>0) OR (TOBODD.GetValue('BLF_MTSITUATION')<>0) then
          begin
            TOBODD.SetAllModifie(true);
            TOBODD.InsertDB(nil,false)
          end;
        end else TOBODD.UpdateDB(false);
      end;
      end
      else
      begin
        PrevFact := TOBO.GetDouble('OLD_QTESITUATION');
        CurrFact := TOBO.GetDouble('BLF_QTESITUATION');
        DejaFact := TOBODD.GetDouble('BLF_QTEDEJAFACT');
        DejaFact := DejaFact - prevFact + Currfact;
        TOBODD.PutValue('BLF_QTEDEJAFACT',Dejafact);
        TOBODD.PutValue('BLF_QTECUMULEFACT',TOBODD.GetDouble('BLF_QTEDEJAFACT'));
        //----
        PrevFact := TOBO.GetDouble('OLD_MTSITUATION');
        CurrFact := TOBO.GetDouble('BLF_MTSITUATION');
        DejaFact := TOBODD.GetDouble('BLF_MTDEJAFACT');
        DejaFact := DejaFact - prevFact + Currfact;

        TOBODD.PutValue('BLF_MTDEJAFACT',Dejafact);
        TOBODD.PutValue('BLF_MTCUMULEFACT',TOBODD.GetDouble('BLF_MTDEJAFACT'));
        if created then
        begin
          TOBODD.SetAllModifie(true);
          TOBODD.InsertDB(nil,false)
        end
        else
        begin
          TOBODD.UpdateDB (false);
    end;
      end;
    end;
    TOBODD.free;
  end;
  {*****************************************************************
  Auteur  ...... :
  Créé le ...... : 10/12/2015
  Modifié le ... :   /  /
  Description .. : Mise à jour de l'ouvrage de la situation suivante
  *****************************************************************}
  procedure MajOuvSituationSuiv(TOBL,TOBOUVRAGES: TOB; DocArecalc : r_cledoc; Avancement : boolean);
  var II,UniqueBLOSuiv,IndiceNomen : Integer;
      TOBOO,TOBOS,TOBLF,TOBOUV : TOB;
      QQ : TQuery;
      Sql : string;
  begin
    IndiceNomen := TOBL.getValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
    TOBOUV := TOBOuvrage.detail[IndiceNomen-1];
    //
    TOBLF := TOB.Create ('LIGNEFAC',nil,-1);
    TRY
      for II := 0 to TOBOUV.Detail.count -1 do
      begin
        TOBOO := TOBOUV.detail[II];
        // On recherche la ligne de lignefac de la situation suivante
        if TOBOO.GetString('BLO_PIECEORIGINE') <> '' then
        begin
          Sql := 'SELECT *, '' AS PIECEPRECEDENTE FROM LIGNEFAC WHERE '+WherePiece(DocArecalc,ttdLignefac,false)+' AND '+
               'BLF_UNIQUEBLO=(SELECT BLO_UNIQUEBLO FROM LIGNEOUV WHERE '+WherePiece(DocArecalc,ttdOuvrage,false) + ' AND '+
                 'BLO_PIECEORIGINE="'+ TOBOO.GetString('BLO_PIECEORIGINE')+'")';
        end else
        begin
          // gestion de l'anteriorite ou le piece origine n'etait pas positionné
          Sql := 'SELECT *, '' AS PIECEPRECEDENTE FROM LIGNEFAC WHERE '+WherePiece(DocArecalc,ttdLignefac,false)+' AND '+
                 'BLF_UNIQUEBLO='+ TOBOO.GetString('BLO_UNIQUEBLO');

        end;
        QQ := OpenSQL(Sql,true,1,'',true);
        if not QQ.eof then
        begin
          TOBLF.SelectDB('',QQ);
          if Avancement then
          begin
            // calcul des elements de la situation suivante
            TOBLF.SetDouble('BLF_MTDEJAFACT',   TOBOO.GetDouble('BLF_MTCUMULEFACT'));
            TOBLF.SetDouble('BLF_MTSITUATION',  TOBLF.GetDouble('BLF_MTCUMULEFACT')-TOBLF.GetDouble('BLF_MTDEJAFACT'));
            TOBLF.SetDouble('BLF_QTEDEJAFACT',  TOBOO.GetDouble('BLF_QTECUMULEFACT'));
            TOBLF.SetDouble('BLF_QTESITUATION', TOBLF.GetDouble('BLF_QTECUMULEFACT')-TOBLF.GetDouble('BLF_QTEDEJAFACT'));
            //
            TOBLF.UpdateDB (false);
          end else
          begin
            TOBLF.SetDouble('BLF_MTDEJAFACT',   TOBOO.GetDouble('BLF_MTDEJAFACT')+TOBOO.GetDouble('BLF_MTSITUATION'));
            TOBLF.SetDouble('BLF_QTEDEJAFACT',  TOBOO.GetDouble('BLF_QTEDEJAFACT')+TOBOO.GetDouble('BLF_QTESITUATION'));
            TOBLF.UpdateDB (false);
          end;
        end;
      end;
    finally
      TOBLF.free;
    end;
  end;

Begin
  //
  FillChar(DocArecalc,SizeOf(DocArecalc),#0);
  //
  Repartition := GetParamSocSecur('SO_PIECEREPART', False);
  //
  Avancement := (Pos(RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS')),'AVA;DAC;')>0);
  //
  TOBLIGNEFAC := TOB.create ('LES LIGNES FACTURATION',nil,-1);
  TOBLS := TOB.Create ('LIGNE',nil,-1);
  TOBLD := TOB.Create ('LIGNE',nil,-1);
  TOBLF := TOB.Create ('LIGNEFAC',nil,-1);
  //
  if Repartition THEN TOBSITSUITE := TOB.create ('SITUATION SUIVANTE',nil,-1);
  //
  NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
  Situations := ((Pos(TOBPiece.GetValue('AFF_GENERAUTO'),'AVA;DIR;')>0) or (TOBPiece.GetValue('AFF_GENERAUTO')=''));

  // Après modification de facture,
  // Si situation, Maj lignes du devis
  TRY
    if (Pos(NaturePiece,'FBT;DAC;FBP;BAC;')>0) and (TOBPiece.GetValue('GP_AFFAIREDEVIS')<>'') then
  begin
      IsLastSituation := DerniereSituation (TobPiece);
      if (not IsLastSituation) and (Situations) then GetSituationSuivante (TobPiece,DocArecalc);
      //
    for i:=0 to TOBPiece.Detail.Count-1 do
    begin

      TOBL:=TOBPiece.Detail[i] ;

      if IsSousDetail(TOBL) then continue;
      //
        TOBLD.InitValeurs;
        TOBLF.InitValeurs;
        TOBLS.InitValeurs;
        
        //Si on a changé la valeur de la ligne
      if (TOBL <> Nil) and (TOBL.GetValue('QTECHANGE')='X') then
      begin
        if TOBL.GetValue ('GL_TYPELIGNE') <> 'ART' Then Continue;
          //
        StPrec:=TOBL.GetValue('GL_PIECEPRECEDENTE') ;
          StOrig:=TOBL.GetValue('GL_PIECEORIGINE') ;
          //
          if (StPrec <> '') and (IsLastSituation) then
        begin
            OKOK := false;
            // On ne remet a jour l'avancement au niveau du devis que si le document traité est bien la derniere situation
          DecodeRefPiece(StPrec,CD) ;
          Req := 'SELECT * FROM LIGNE WHERE '+ WherePiece (CD,ttdLigne,true,True);
          QQ:= OpenSql (Req,true,-1,'',true);
          if not QQ.eof then
          begin
              TOBLD.SelectDB ('',QQ);
              OKOK := True;
            end;
            Ferme(QQ);
            if Okok then
            begin
            //
            if Avancement then
            begin
            	if TOBL.getValue('GL_QTEPREVAVANC')=0 then
              begin
              	Pourcent := Arrondi(TOBL.GetValue('BLF_POURCENTAVANC'),2);
              end else
              begin
              	Pourcent := Arrondi(TOBL.GetValue('GL_QTESIT')/TOBL.GetValue('GL_QTEPREVAVANC')*100,2);
              end;
                TOBLD.PutValue('GL_QTEPREVAVANC',TOBL.GetValue('GL_QTESIT'));
                TOBLD.PutValue('GL_QTESIT',TOBL.GetValue('GL_QTESIT'));
                TOBLD.PutValue('GL_POURCENTAVANC',Pourcent);
            end else
            begin
                TOBLD.PutValue('GL_QTEPREVAVANC',TOBLD.GetValue('GL_QTEPREVAVANC')-TOBL.GetValue('OLD_QTESIT')+TOBL.GetValue('GL_QTEFACT'));
                TOBLD.PutValue('GL_QTESIT',TOBLD.GetValue('GL_QTEPREVAVANC'));
            end;
              TOBLD.UpdateDB (false);
            //
              OKOK := false;
              Req := 'SELECT *, '' AS PIECEPRECEDENTE FROM LIGNEFAC WHERE '+ WherePiece (CD,ttdLigneFac,true,True);
              QQ:= OpenSql (Req,true,-1,'',true);
              if not QQ.eof then
              begin
                TOBLF.SelectDB ('',QQ);
                OKOK := True;
              end;
              ferme (QQ);
              if OKOK then
              begin
                if avancement then
                begin
                  TOBLF.PutValue('BLF_QTEDEJAFACT',  TOBL.GetValue('BLF_QTECUMULEFACT'));
                TOBLF.PutValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTECUMULEFACT'));
                //
                  TOBLF.PutValue('BLF_MTDEJAFACT',   TOBL.GetValue('BLF_MTCUMULEFACT'));
                  TOBLF.PutValue('BLF_MTCUMULEFACT', TOBL.GetValue('BLF_MTCUMULEFACT'));
                //
                  if TOBL.GetValue('BLF_MTMARCHE') <> 0 then
                  begin
                Pourcent := arrondi((TOBL.GetValue('BLF_MTCUMULEFACT') / TOBL.GetValue('BLF_MTMARCHE')) * 100,2);
                  end else
                  begin
                    Pourcent := 100;
                  end;
                TOBLF.PutValue('BLF_POURCENTAVANC',Pourcent);
                if TOBLF.GetValue('BLF_QTEDEJAFACT') <> 0 then
                begin
                  TOBLF.UpdateDB (false);
                end else
                begin
                  TOBLF.DeleteDB (false);
                end;
                end
                else
                begin
                  PrevFact := TOBL.GetDouble('OLD_QTESITUATION');
                  CurrFact := TOBL.GetDouble('BLF_QTESITUATION');
                  DejaFact := TOBLF.GetDouble('BLF_QTEDEJAFACT');
                  DejaFact := DejaFact - prevFact + Currfact;
                  //
                  TOBLF.SetDouble('BLF_QTEDEJAFACT',Dejafact);
                  TOBLF.SetDouble('BLF_QTECUMULEFACT',TOBLF.GetDouble('BLF_QTEDEJAFACT'));
                  // --------------------
                  PrevFact := TOBL.GetDouble('OLD_MTSITUATION');
                  CurrFact := TOBL.GetDouble('BLF_MTSITUATION');
                  DejaFact := TOBLF.GetDouble('BLF_MTDEJAFACT');
                  //
                  DejaFact := DejaFact - prevFact + Currfact;
                  //
                  TOBLF.SetDouble('BLF_MTDEJAFACT',   DejaFact);
                  TOBLF.SetDouble('BLF_MTCUMULEFACT', TOBLF.GetDouble('BLF_MTDEJAFACT'));
                  TOBLF.UpdateDB (false);
                  //
              end;
              end;
              //
              if IsOuvrage(TOBL) and (TOBL.getValue('GL_INDICENOMEN')>0) then
                MajLigneDevisOuv (TOBL,TOBOUvrage,Avancement);
              end;
          end
          else if (DocArecalc.NumeroPiece<>0) and (Situations) and (not IsLastSituation) and (StPrec <> '')  then
          begin

            OkCont := True;

            // S'il y a une situation suivante --> mise a jour du deja facture + recalcul de la ligne
            Req := 'SELECT * FROM LIGNE WHERE ' + WherePiece (DocArecalc,ttdLigne,false) + ' AND GL_PIECEORIGINE="' + StOrig + '"';
            QQ:= OpenSql (Req,true,-1,'',true);
            if not QQ.eof then
              TOBLS.SelectDB ('',QQ)
            else
              OkCont := false;
            ferme (QQ);

            if OkCont then
            begin
              Req := 'SELECT *, '' AS PIECEPRECEDENTE FROM LIGNEFAC WHERE ' + WherePiece (DocArecalc,ttdLignefac,false);
              Req := Req + ' AND ' + 'BLF_NUMORDRE="' + TOBLS.GetString('GL_NUMORDRE') + '" AND BLF_UNIQUEBLO=0';
              QQ:= OpenSql (Req,true,-1,'',true);
              if not QQ.eof then
                TOBLF.SelectDB ('',QQ)
              else
                OkCont := False;
              ferme (QQ);
          end;

            if not OkCont then continue;

              if Avancement then
              begin
                // calcul des elements de la situation suivante
              TOBLF.SetDouble('BLF_MTDEJAFACT',   TOBL.GetDouble('BLF_MTCUMULEFACT'));
              TOBLF.SetDouble('BLF_MTSITUATION',  TOBLF.GetDouble('BLF_MTCUMULEFACT')-  TOBLF.GetDouble('BLF_MTDEJAFACT'));
              TOBLF.SetDouble('BLF_QTEDEJAFACT',  TOBL.GetDouble('BLF_QTECUMULEFACT'));
              TOBLF.SetDouble('BLF_QTESITUATION', TOBLF.GetDouble('BLF_QTECUMULEFACT')- TOBLF.GetDouble('BLF_QTEDEJAFACT'));
                //
              TOBLS.SetDouble('GL_QTEFACT',       TOBLF.GetDouble('BLF_QTESITUATION'));
              TOBLS.SetDouble('GL_QTERESTE',      TOBLF.GetDouble('BLF_QTESITUATION'));
              TOBLS.SetDouble('GL_QTESTOCK',      TOBLF.GetDouble('BLF_QTESITUATION'));
              TOBLS.SetDouble('GL_QTERELIQUAT',   0);
              //--- GUINIER ---
              if CtrlOkReliquat(TOBLS, 'GL') then
              begin
                TOBLS.SetDouble('GL_MTRESTE',     TOBLF.GetDouble('BLF_MTSITUATION'));
                TOBLS.SetDouble('GL_MTRELIQUAT',  0);
              end;
              //
                TOBLF.UpdateDB (false);
              TOBLS.UpdateDB (false);
            end
            else
              begin
              TOBLF.SetDouble('BLF_MTDEJAFACT', TOBL.GetDouble('BLF_MTDEJAFACT')+TOBL.GetDouble('BLF_MTSITUATION'));
                TOBLF.SetDouble('BLF_QTEDEJAFACT',TOBL.GetDouble('BLF_QTEDEJAFACT')+TOBL.GetDouble('BLF_QTESITUATION'));
              //
                TOBLF.UpdateDB (false);
        end;
            //
            if Repartition      then CreateTobSituationSuivante(TOBLS, TOBLF);
            //
              if IsOuvrage(TOBL) then MajOuvSituationSuiv(TOBL,TOBOUvrage,DocArecalc,Avancement);

      end;
          end;
      //
        //
        //
        if (TOBL <> Nil) then
      begin
        if (TOBL.getValue('BLF_MTMARCHE') <> 0) or (TOBL.getValue('BLF_MTSITUATION')<>0) then
        begin
            TOBLN := TOB.Create ('LIGNEFAC',TOBLIGNEFAC,-1);
            SetLigneFacture (TOBL,TOBLN);
            if IsOuvrage(TOBL) then MajLignesFacOuvr(TOBL, TOBOUvrage);
            if Repartition then TOBLN.AddChampSupValeur('PIECEPRECEDENTE', TOBL.GetString('GL_PIECEPRECEDENTE'));
        end;
      end;
      //
    end;
    //
    if TOBLIGNEFAC.detail.count > 0 then TOBLIGNEFAC.InsertDB (nil);
    //
      //On recalcul les répartitions avec les nouvelles modifications....
      if Repartition then
      begin
        if TOBSITSUITE.detail.count > 0 then
        begin
          //on charge la tob temporaire dans la tob finale
          For ind_2 := 0 TO TobSitSuite.detail.count -1 do
          begin
            TOBLN := TOB.Create ('LIGNEFAC',TOBLIGNEFAC,-1);
            TOBLN.Dupliquer(TOBSITSUITE.detail[ind_2], False, True);
          end;
          TOBSITSUITE.ClearDetail;
        end;
        //on supprimme d'abord les enregistrements correspondant à la pièce en cours....
        DestruitPrepaFact(TobPiece.GetString('GP_NATUREPIECEG'), Tobpiece.GetString('GP_SOUCHE'), TobPiece.GetInteger('GP_NUMERO'));
        //Rechargement du fichier de répartition
        ChargePieceRepart(TOBPIECE.GetString('GP_AFFAIREDEVIS'), TOBLIGNEFAC);
      end;

      //
      ReajusteCautionDocOrigine (TOBPIECE,TOBPieceRG,true,TOBPieceRG_O);
      // Reajustement de BSITUATION si besoin est

      req := 'SELECT * FROM BSITUATIONS WHERE BST_NATUREPIECE="'+TOBPIece.GetValue('GP_NATUREPIECEG')+'" AND ';
      req := Req + 'BST_SOUCHE="'+TOBPIece.GetVAlue('GP_SOUCHE')+'" AND BST_NUMEROFAC="'+inttoStr(TOBPiece.GetValue('GP_NUMERO'))+'"';
      QQ := OpenSql(Req,true,-1,'',true);
      if not QQ.eof then
      begin
        GetMontantsAcomptes (TOBAcomptes,MontantAcompte,MontantRegl);
        GetMontantRG (TOBPieceRG,TOBBasesRG,XD,XP,DEV,True,True);
        GetcumultaxesRG (TOBBasesRG,TOBPieceRG,TXD,TXP,DEV);
        TOBSit := TOB.create ('BSITUATIONS',nil,-1);
        TOBSIt.SelectDB ('',QQ);
        TTC := TOBPiece.GetValue('GP_TOTALTTCDEV') - XD - TXD;
        TOBSIt.putvalue('BST_MONTANTHT',TOBPiece.GetValue('GP_TOTALHTDEV'));
        TOBSIt.putvalue('BST_MONTANTTVA',TTC - TOBPiece.GetValue('GP_TOTALHTDEV'));
        TOBSIt.putvalue('BST_MONTANTTTC',TTC);
        TOBSIT.PutValue('BST_MONTANTREGL',MontantRegl);
        TOBSIT.PutValue('BST_MONTANTACOMPTE',MontantAcompte);
        TOBSit.SetAllModifie (true);
        TOBSIT.UpdateDB (false);
        TOBSIT.free;
      end;
      ferme (QQ);
    end else if (NaturePiece='B00') and (TOBpiece.GetString('GP_ATTACHEMENT')<>'') and (GetparamSocSecur('SO_BTACPT1SIT',false)) then
    begin
      // cas d'un acompte sur devis et considéré comme 1ere Situation
      DecodeRefPiece(TOBpiece.getString('GP_ATTACHEMENT'),CD);
      RefPieceSel := EncoderefSel (CD.NaturePiece,CD.Souche ,CD.NumeroPiece ,CD.Indice);

      DossierFac := ''; Avancement := false; NbSit := 0;

      QQ := OpenSql ('SELECT GP_AFFAIREDEVIS,AFF_GENERAUTO,'+        // DEVIS
                     '(SELECT COUNT(*) FROM BSITUATIONS WHERE BST_SSAFFAIRE=GP_AFFAIREDEVIS) AS NBSIT '+
                     'FROM PIECE '+
                     'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS '+
                     'WHERE '+ WherePiece(CD,ttdPiece,false),true,1,'',true);
      if not QQ.eof then
      begin
        Avancement := (Pos(QQ.fields[1].AsString ,'AVA;DAC;')>0);
        NbSit := QQ.fields[2].AsInteger;
        DossierFac := QQ.fields[0].AsString;
      end;
      ferme (QQ);
      if Avancement then
      begin
        TOBSit := TOB.create ('BSITUATIONS',nil,-1);
        TOBMEMFAC := TOB.create ('BTMEMOFACTURE',nil,-1);
        TRY
          if (NbSit = 0) and (DossierFac<>'') then
          begin
            TOBSit.SetString('BST_NATUREPIECE',TOBPiece.getString('GP_NATUREPIECEG'));
            TOBSit.SetString('BST_SOUCHE',TOBPiece.getString('GP_SOUCHE'));
            TOBSit.SetInteger('BST_NUMEROFAC',TOBPiece.getInteger('GP_NUMERO'));
            if GetparamSocSecur('SO_BTACPTSIT0',false) then TOBSit.SetInteger('BST_NUMEROSIT',0) else TOBSit.SetInteger('BST_NUMEROSIT',1); 
            TOBSit.PutValue('BST_DATESIT',TOBPiece.getValue('GP_DATEPIECE'));
            TOBSit.SetString('BST_AFFAIRE',TOBPiece.getString('GP_AFFAIRE'));
            TOBSit.SetString('BST_SSAFFAIRE',DossierFac);
            TOBSit.SetDouble('BST_MONTANTHT',TOBpiece.getDouble('GP_TOTALHTDEV'));
            TOBSit.SetDouble('BST_MONTANTTTC',TOBpiece.getDouble('GP_TOTALTTCDEV'));
            TOBSit.SetDouble('BST_MONTANTTVA',TOBSit.getDouble('BST_MONTANTTTC')-TOBSit.getDouble('BST_MONTANTHT'));
            TOBSit.SetBoolean('BST_VIVANTE',true);
            TOBSit.SetInteger('BST_INDICESIT',0);
            TOBSIT.SetAllModifie(true);
            TOBSit.InsertDB(nil);
            //
          end else if (NbSit = 1) and (Dossierfac <> '') then
          begin
            QQ := OpenSql ('SELECT * FROM BSITUATIONS WHERE BST_SSAFFAIRE="'+DossierFac+'"',true,1,'',true);
            if not QQ.eof then
            begin
              TOBSIT.SelectDB('',QQ);
            end;
            ferme (QQ);
            if (TOBSIT.GetString('BST_NATUREPIECE')= TOBPiece.getString('GP_NATUREPIECEG')) and
               (TOBSIT.GetString('BST_SOUCHE')= TOBPiece.getString('GP_SOUCHE')) and
               (TOBSIT.GetInteger('BST_NUMEROFAC')= TOBPiece.getInteger('GP_NUMERO')) then
            begin
              TOBSit.SetDouble('BST_MONTANTHT',TOBpiece.getDouble('GP_TOTALHTDEV'));
              TOBSit.SetDouble('BST_MONTANTTTC',TOBpiece.getDouble('GP_TOTALTTCDEV'));
              TOBSit.SetDouble('BST_MONTANTTVA',TOBSit.getDouble('BST_MONTANTTTC')-TOBSit.getDouble('BST_MONTANTHT'));
              TOBSit.UpdateDB(false);
            end;
          end;
          // ----
          QQ := OpenSql ('SELECT 1 FROM BTMEMOFACTURE WHERE BMF_DEVISPRINC="'+RefPieceSel+'" AND '+
                         'BMF_AFFAIRE= "'+TOBPiece.getString('GP_AFFAIRE')+'" AND '+
                         'BMF_DEVIS="'+RefPieceSel+'"',true,1,'',true);
          if QQ.eof then
          begin
            TOBMEMFAC.SetString('BMF_DEVISPRINC',RefPieceSel);
            TOBMEMFAC.SetInteger('BMF_NUMORDRE',1);
            TOBMEMFAC.SetString('BMF_DEVIS',RefPieceSel);
            TOBMEMFAC.SetString('BMF_AFFAIRE',TOBPiece.getString('GP_AFFAIRE'));
            TOBMEMFAC.SetString('BMF_AFFAIRE1',TOBPiece.getString('GP_AFFAIRE1'));
            TOBMEMFAC.SetString('BMF_AFFAIRE2',TOBPiece.getString('GP_AFFAIRE2'));
            TOBMEMFAC.SetString('BMF_AFFAIRE3',TOBPiece.getString('GP_AFFAIRE3'));
            TOBMEMFAC.SetString('BMF_NATUREPIECEG',CD.NaturePiece);
            TOBMEMFAC.SetString('BMF_SOUCHE',CD.Souche);
            TOBMEMFAC.SetInteger('BMF_NUMERO',CD.NumeroPiece);
            TOBMEMFAC.SetInteger('BMF_INDICEG',CD.Indice);
            TOBMEMFAC.SetAllModifie(true);
            TOBMEMFAC.InsertDB(nil);  
          end;
          ferme (QQ);
        FINALLY
          TOBSit.free;
          TOBMEMFAC.free;
        END;
      end;
    end else if (NaturePiece='DBT') and (Avancement) and (TOBpiece.GetString('GP_ATTACHEMENT')<>'') and (GetparamSocSecur('SO_BTACPT1SIT',false)) then
    begin
      DecodeRefPiece(TOBpiece.getString('GP_ATTACHEMENT'),CD);
      DossierFac := TOBPiece.getString('GP_AFFAIREDEVIS'); Avancement := false; NbSit := 0;
      RefPieceSel := EncoderefSel (TOBPiece.getString('GP_NATUREPIECEG'),TOBPiece.getString('GP_SOUCHE') ,TOBPiece.getInteger('GP_NUMERO') ,TOBPiece.getInteger('GP_INDICEG'));

      TOBSit := TOB.create ('BSITUATIONS',nil,-1);
      TOBMEMFAC := TOB.create ('BTMEMOFACTURE',nil,-1);
      TRY
        Marche := 0; TTC := 0; TVA := 0;
        //
        QQ := openSql ('SELECT GP_TOTALHTDEV,GP_TOTALTTCDEV FROM PIECE WHERE '+WherePiece (CD,ttdPiece,false),true,1,'',true);
        if not QQ.eof then
        begin
          Marche := QQ.fields[0].AsFloat;
          TTC := QQ.fields[1].AsFloat;
          TVA := TTC - Marche ;
        end;
        ferme(QQ);

        QQ := OpenSql ('SELECT COUNT(*) FROM BSITUATIONS WHERE BST_SSAFFAIRE="'+TOBPiece.getString('GP_AFFAIREDEVIS')+'"',true,-1,'',true);
        if not QQ.eof then
        begin
          NbSit := QQ.fields[0].AsInteger;
        end;
        ferme (QQ);
        
        if (NbSit = 0) and (DossierFac<>'') then
        begin
          TOBSit.SetString('BST_NATUREPIECE',CD.NaturePiece);
          TOBSit.SetString('BST_SOUCHE',CD.Souche);
          TOBSit.SetInteger('BST_NUMEROFAC',CD.NumeroPiece);
          if GetparamSocSecur('SO_BTACPTSIT0',false) then TOBSit.SetInteger('BST_NUMEROSIT',0) else TOBSit.SetInteger('BST_NUMEROSIT',1); 
          TOBSit.PutValue('BST_DATESIT',CD.DatePiece);
          TOBSit.SetString('BST_AFFAIRE',TOBPiece.getString('GP_AFFAIRE'));
          TOBSit.SetString('BST_SSAFFAIRE',TOBPiece.getString('GP_AFFAIREDEVIS'));
          TOBSit.SetDouble('BST_MONTANTHT',Marche);
          TOBSit.SetDouble('BST_MONTANTTTC',TTC);
          TOBSit.SetDouble('BST_MONTANTTVA',TVA);
          TOBSit.SetBoolean('BST_VIVANTE',true);
          TOBSit.SetInteger('BST_INDICESIT',0);
          TOBSIT.SetAllModifie(true);
          TOBSit.InsertDB(nil)
        end else if (NbSit = 1) and (Dossierfac <> '') then
        begin
          QQ := OpenSql ('SELECT * FROM BSITUATIONS WHERE BST_SSAFFAIRE="'+DossierFac+'"',true,1,'',true);
          if not QQ.eof then
          begin
            TOBSIT.SelectDB('',QQ);
          end;
          ferme (QQ);
          if (TOBSIT.GetString('BST_NATUREPIECE')= CD.NaturePiece) and
             (TOBSIT.GetString('BST_SOUCHE')= CD.Souche) and
             (TOBSIT.GetInteger('BST_NUMEROFAC')= CD.NumeroPiece) then
          begin
            TOBSit.SetDouble('BST_MONTANTHT',Marche);
            TOBSit.SetDouble('BST_MONTANTTTC',TTC);
            TOBSit.SetDouble('BST_MONTANTTVA',TVA);
            TOBSit.UpdateDB(false);
          end;
        end;
        //
        QQ := OpenSql ('SELECT 1 FROM BTMEMOFACTURE WHERE BMF_DEVISPRINC="'+RefPieceSel+'" AND '+
                       'BMF_AFFAIRE= "'+TOBPiece.getString('GP_AFFAIRE')+'" AND '+
                       'BMF_DEVIS="'+RefPieceSel+'"',true,1,'',true);
        if QQ.eof then
        begin
          TOBMEMFAC.SetString('BMF_DEVISPRINC',RefPieceSel);
          TOBMEMFAC.SetInteger('BMF_NUMORDRE',1);
          TOBMEMFAC.SetString('BMF_DEVIS',RefPieceSel);
          TOBMEMFAC.SetString('BMF_AFFAIRE',TOBPiece.getString('GP_AFFAIRE'));
          TOBMEMFAC.SetString('BMF_AFFAIRE1',TOBPiece.getString('GP_AFFAIRE1'));
          TOBMEMFAC.SetString('BMF_AFFAIRE2',TOBPiece.getString('GP_AFFAIRE2'));
          TOBMEMFAC.SetString('BMF_AFFAIRE3',TOBPiece.getString('GP_AFFAIRE3'));
          TOBMEMFAC.SetString('BMF_NATUREPIECEG',TOBPiece.getString('GP_NATUREPIECEG'));
          TOBMEMFAC.SetString('BMF_SOUCHE',TOBPiece.getString('GP_SOUCHE'));
          TOBMEMFAC.SetInteger('BMF_NUMERO',TOBPiece.getInteger('GP_NUMERO'));
          TOBMEMFAC.SetInteger('BMF_INDICEG',TOBPiece.getInteger('GP_INDICEG'));
          TOBMEMFAC.SetAllModifie(true);
          TOBMEMFAC.InsertDB(nil);
        end;
        ferme (QQ);
      FINALLY
        TOBSit.free;
        TOBMEMFAC.free;
      END;
    end;
  FINALLY
    TOBLIGNEFAC.Free;
    TOBLS.free;
    TOBLD.free;
    TOBLF.free;
  END;
End;

procedure AnnuleSituationDetail (TOBL,TOBOUvrage : TOB; Avancement : boolean);
var TOBOUV,TOBO,TOBLF :  TOB;
    II,IndiceNomen : integer;
    StPrec,Req : string;
    Pourcent : double;
    CD : r_cledoc;
    okok : boolean;
    QQ : TQuery;
begin
  TOBLF := TOB.Create ('LIGNEFAC',nil,-1);
  TRY
    OKOK := false;
    IndiceNomen := TOBL.GetInteger('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
    if IndiceNomen > TOBOuvrage.detail.count then exit;
    TOBOUV := TOBOuvrage.detail[IndiceNomen-1];
    if TOBOUV.detail.count = 0 then exit;
    for II := 0 to TOBOUV.detail.count -1 do
    begin
      TOBO := TOBOUV.detail[II];
      StPrec:=TOBO.GetValue('BLO_PIECEPRECEDENTE'); DecodeRefPiece(StPrec,CD); if CD.NaturePiece <> 'DBT' then continue;
      if Stprec <> '' then
      begin
        DecodeRefPieceOUv(StPrec,CD) ;
        Req := 'SELECT * FROM LIGNEFAC WHERE '+ WherePiece (CD,ttdLigneFac,false)+' AND BLF_NUMORDRE=0 AND BLF_UNIQUEBLO='+IntToStr(CD.UniqueBlo);
        QQ:= OpenSql (Req,true,-1,'',true);
        if not QQ.eof then
        begin
          TOBLF.SelectDB ('',QQ);
          OKOK := True;
        end;
        Ferme(QQ);
        if OKOK then
        begin
          TOBLF.PutValue('BLF_QTEDEJAFACT',TOBLF.GetValue('BLF_QTEDEJAFACT')- TOBO.GetValue('BLF_QTESITUATION'));
          TOBLF.PutValue('BLF_QTECUMULEFACT',TOBLF.GetValue('BLF_QTECUMULEFACT')-TOBO.GetValue('BLF_QTESITUATION'));
          //
          TOBLF.PutValue('BLF_MTDEJAFACT',TOBLF.GetValue('BLF_MTDEJAFACT')-TOBO.GetValue('BLF_MTSITUATION'));
          TOBLF.PutValue('BLF_MTCUMULEFACT',TOBLF.GetValue('BLF_MTCUMULEFACT')-TOBO.GetValue('BLF_MTSITUATION'));
          if avancement then
          begin
            if TOBLF.GetValue('BLF_QTEMARCHE') <> 0 then Pourcent := arrondi((TOBLF.GetValue('BLF_QTECUMULEFACT') / TOBLF.GetValue('BLF_QTEMARCHE')) * 100,2)
                                                    else Pourcent := 0;
            TOBLF.PutValue('BLF_POURCENTAVANC',Pourcent);
          end;
          if TOBLF.GetDouble('BLF_MTDEJAFACT')<>0 then TOBLF.UpdateDB (false)
                                                  else TOBLF.DeleteDB(false);  
        end;
      end;
    end;
  FINALLY
    TOBLF.Free;
  END;
end;

procedure AnnuleSituation (TOBpiece,TOBOUvrage,TOBPieceRg : TOB);
var i : integer;
    TOBL : TOB;
    TOBLD: TOB;
    TOBLF : TOB;
    StPrec,StOrig,Req : string;
    CD : R_cledoc;
    OKOK,Avancement : boolean;
    QQ: TQuery;
    Pourcent : double;
begin

  TOBLD       := TOB.Create ('LIGNE',nil,-1);
  TOBLF       := TOB.Create ('LIGNEFAC',nil,-1);
  Avancement := (Pos(RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS')),'AVA;DAC;')>0);
  //
  for i:=0 to TOBPiece.Detail.Count-1 do
  begin

    TOBL:=TOBPiece.Detail[i] ;

    if IsSousDetail(TOBL) then continue;
    //
    TOBLD.InitValeurs;
    TOBLF.InitValeurs;

    if (TOBL <> Nil)  then
    begin
      if TOBL.GetValue ('GL_TYPELIGNE') <> 'ART' Then Continue;
      //
      StPrec:=TOBL.GetValue('GL_PIECEPRECEDENTE') ;
      StOrig:=TOBL.GetValue('GL_PIECEORIGINE') ;
      //
      DecodeRefPiece(StPrec,CD); if CD.NaturePiece <> 'DBT' then StPrec := StOrig;
      if (StPrec <> '') then
      begin
        OKOK := false;
        // On ne remet a jour l'avancement au niveau du devis que si le document traité est bien la derniere situation
        DecodeRefPiece(StPrec,CD) ;
        Req := 'SELECT * FROM LIGNE WHERE '+ WherePiece (CD,ttdLigne,true,True);
        QQ:= OpenSql (Req,true,-1,'',true);
        if not QQ.eof then
        begin
          TOBLD.SelectDB ('',QQ);
          OKOK := True;
        end;
        Ferme(QQ);
        if Okok then
        begin
          TOBLD.SetDouble('GL_QTEPREVAVANC',TOBLD.GetDouble('GL_QTEPREVAVANC')-TOBL.GetValue('GL_QTEFACT'));
          TOBLD.SetDouble('GL_QTESIT',TOBLD.GetDouble('GL_QTEPREVAVANC'));
          if Avancement then
          begin
            if TOBLD.getValue('GL_QTEPREVAVANC')=0 then
            begin
              Pourcent := 0;
            end else
            begin
              if TOBLD.GetValue('GL_QTEFACT') <> 0 then Pourcent := Arrondi(TOBLD.GetValue('GL_QTESIT')/TOBLD.GetValue('GL_QTEFACT')*100,2)
                                                   else if TOBLD.GetValue('GL_QTESIT') <> 0 then Pourcent := 100 else Pourcent := 0;
            end;
            TOBLD.PutValue('GL_POURCENTAVANC',Pourcent);
          end else
          begin
            TOBLD.PutValue('GL_POURCENTAVANC',0);
          end;
          TOBLD.UpdateDB (false);
          //
          OKOK := false;
          Req := 'SELECT *, '' AS PIECEPRECEDENTE FROM LIGNEFAC WHERE '+ WherePiece (CD,ttdLigneFac,true,True);
          QQ:= OpenSql (Req,true,-1,'',true);
          if not QQ.eof then
          begin
            TOBLF.SelectDB ('',QQ);
            OKOK := True;
          end;
          ferme (QQ);
          if OKOK then
          begin
            TOBLF.SetDouble('BLF_QTEDEJAFACT',  TOBLF.getDouble('BLF_QTEDEJAFACT')-TOBL.getDouble('GL_QTEFACT'));
            TOBLF.SetDouble('BLF_QTECUMULEFACT',TOBLF.GetDouble('BLF_QTECUMULEFACT') - TOBL.getDouble('GL_QTEFACT'));
            //
            TOBLF.SetDouble('BLF_MTDEJAFACT',   TOBLF.getDouble('BLF_MTDEJAFACT')-TOBL.getDouble('BLF_MTSITUATION'));
            TOBLF.SetDouble('BLF_MTCUMULEFACT', TOBLF.GetDouble('BLF_MTCUMULEFACT') - TOBL.getDouble('BLF_MTSITUATION'));
            //
            if Avancement then
            begin
              if TOBLF.GetValue('BLF_MTMARCHE') <> 0 then
              begin
                Pourcent := arrondi((TOBLF.GetValue('BLF_MTCUMULEFACT') / TOBLF.GetValue('BLF_MTMARCHE')) * 100,2);
              end else
              begin
                Pourcent := 100;
              end;
              TOBLF.PutValue('BLF_POURCENTAVANC',Pourcent);
            end;
            if TOBLF.GetValue('BLF_QTEDEJAFACT') <> 0 then
            begin
              TOBLF.UpdateDB (false);
            end else
            begin
              TOBLF.DeleteDB (false);
            end;
          end;
          //
          if IsOuvrage(TOBL) and (TOBL.getValue('GL_INDICENOMEN')>0) then
          BEGIN
            AnnuleSituationDetail (TOBL,TOBOUvrage,Avancement);
          END;
        end;
      end
    end;
    //
  end;
  //
  AnnuleCautionUtilise (TOBPIECE,TOBPieceRG);
  //
  req := 'UPDATE BSITUATIONS SET '+
         'BST_VIVANTE="-" '+
         'WHERE '+
         'BST_SSAFFAIRE="'+TOBPIece.GetValue('GP_AFFAIREDEVIS')+'" AND '+
         'BST_NATUREPIECE="'+TOBPIece.GetValue('GP_NATUREPIECEG')+'" AND '+
         'BST_SOUCHE="'+TOBPIece.GetVAlue('GP_SOUCHE')+'" AND '+
         'BST_NUMEROFAC="'+inttoStr(TOBPiece.GetValue('GP_NUMERO'))+'"';
  //
  ExecuteSql (Req);
end;



procedure ReajusteSituation (TOBPiece,TOBAcomptes,TOBPieceRG,TOBBasesRG : TOB; DEV : RDevise);
var Req : string;
    MontantRegl,MontantAcompte,XD,XP,TXD,TXP,TTC : double;
    TOBSIt : TOB;
    QQ : TQuery;
begin
  if Pos(TOBPiece.getString('GP_NATUREPIECEG'),'FBP;FBT;DAC')<=0 then exit;
  //
  req := 'SELECT * FROM BSITUATIONS WHERE BST_NATUREPIECE="'+TOBPIece.GetValue('GP_NATUREPIECEG')+'" AND ';
  req := Req + 'BST_SOUCHE="'+TOBPIece.GetVAlue('GP_SOUCHE')+'" AND BST_NUMEROFAC="'+inttoStr(TOBPiece.GetValue('GP_NUMERO'))+'"';
  QQ := OpenSql(Req,true,-1,'',true);
  if not QQ.eof then
  begin
    GetMontantsAcomptes (TOBAcomptes,MontantAcompte,MontantRegl);
    GetMontantRG (TOBPieceRG,TOBBasesRG,XD,XP,DEV,True,True);
    GetcumultaxesRG (TOBBasesRG,TOBPieceRG,TXD,TXP,DEV);
    TOBSit := TOB.create ('BSITUATIONS',nil,-1);
    TOBSIt.SelectDB ('',QQ);
    TTC := TOBPiece.GetValue('GP_TOTALTTCDEV') - XD - TXD;
    TOBSIt.putvalue('BST_MONTANTHT',TOBPiece.GetValue('GP_TOTALHTDEV'));
    TOBSIt.putvalue('BST_MONTANTTVA',TTC - TOBPiece.GetValue('GP_TOTALHTDEV'));
    TOBSIt.putvalue('BST_MONTANTTTC',TTC);
    TOBSIT.PutValue('BST_MONTANTREGL',MontantRegl);
    TOBSIT.PutValue('BST_MONTANTACOMPTE',MontantAcompte);
    TOBSit.SetAllModifie (true);
    TOBSIT.UpdateDB (false);
    TOBSIT.free;
  end;
  ferme (QQ);
end;


function GetClefEcriture (TOBA: TOB) : string;
begin
  result :='E_JOURNAL="'+TOBA.GetValue('GAC_JALECR')+'" AND E_NUMEROPIECE='+IntToStr(TOBA.GetValue('GAC_NUMECR'));
end;

function ReactiveSituationProvisoirePrec(TOBSitPre : TOB) : boolean;

  function ReactivePiece (TOBSitPre : TOB):boolean;
  var II : Integer;
      Req : string;
      nbr : integer;
  begin
    result := true;
    Req := 'UPDATE PIECE SET GP_VIVANTE="X" '+
           'WHERE '+
           'GP_NATUREPIECEG="'+TOBSitPre.GetString('BST_NATUREPIECE')+'" AND '+
           'GP_SOUCHE="'+TOBSitPre.GetString('BST_SOUCHE')+'" AND '+
           'GP_NUMERO='+TOBSitPre.GetString('BST_NUMEROFAC')+' AND '+
           'GP_INDICEG=0';
    Nbr := ExecuteSQL(Req);
    if (nbr=0) then
    begin
      Result := false;
      exit;
    end;
    Req := 'UPDATE LIGNE SET GL_VIVANTE="X" '+
           'WHERE '+
           'GL_NATUREPIECEG="'+TOBSitPre.GetString('BST_NATUREPIECE')+'" AND '+
           'GL_SOUCHE="'+TOBSitPre.GetString('BST_SOUCHE')+'" AND '+
           'GL_NUMERO='+TOBSitPre.GetString('BST_NUMEROFAC')+' AND '+
           'GL_INDICEG=0';
    Nbr := ExecuteSQL(Req);
    if (nbr=0) then
    begin
      Result := false;
      exit;
    end;
    //
    Req := 'UPDATE BSITUATIONS SET BST_VIVANTE="X" '+
           'WHERE '+
           'BST_NATUREPIECE="'+TOBSitPre.GetString('BST_NATUREPIECE')+'" AND '+
           'BST_SOUCHE="'+TOBSitPre.GetString('BST_SOUCHE')+'" AND '+
           'BST_NUMEROFAC='+TOBSitPre.GetString('BST_NUMEROFAC');
    Nbr := ExecuteSQL(Req);
    if (nbr=0) then
    begin
      Result := false;
      exit;
    end;
  end;

  procedure  EnregPieceAReactiver (CD: R_CLEDOC;TOBPieceReactiv : TOB);
  var TOBL : TOB;
  begin
    TOBL := TOBPieceReactiv.FindFirst(['NATUREPIECEG','SOUCHE','NUMERO','INDICE'],
                                      [CD.NaturePiece,CD.Souche,CD.NumeroPiece,CD.Indice],True);
    if TOBL = nil then
    begin
      TOBL := TOB.Create ('UN DOC',TOBPieceReactiv,-1);
      TOBL.AddChampSupValeur('NATUREPIECEG',CD.NaturePiece);
      TOBL.AddChampSupValeur('SOUCHE',CD.Souche);
      TOBL.AddChampSupValeur('NUMERO',CD.NumeroPiece);
      TOBL.AddChampSupValeur('INDICE',CD.Indice);
    end;
  end;

  FUNCTION TraiteLignesOuv(TOBL : TOB): boolean;
  var QQ : TQuery;
      TOBOUV,TOBOO : TOB;
      Req,StPrec : string;
      II : Integer;
      CD : r_cledoc;
      NewMtCum,NewQteCum,PourcentAvanc : Double;
  begin
    Result := True;
    TOBOUV := TOB.Create ('LES LIGNES OUV',nil,-1);
    //
    Req := 'SELECT BLO_PIECEPRECEDENTE,BLF_QTESITUATION,BLF_MTSITUATION FROM LIGNEOUV '+
           'LEFT JOIN LIGNEFAC ON '+
           'BLF_NATUREPIECEG=BLO_NATUREPIECEG AND '+
           'BLF_SOUCHE=BLO_SOUCHE AND '+
           'BLF_NUMERO=BLO_NUMERO AND '+
           'BLF_INDICEG=BLO_INDICEG AND '+
           'BLF_NUMORDRE=BLO_NUMORDRE '+
           'WHERE '+
           'BLO_NATUREPIECEG="'+TOBL.GetString('GL_NATUREPIECEG')+'" AND '+
           'BLO_SOUCHE="'+TOBL.GetString('GL_SOUCHE')+'" AND '+
           'BLO_NUMERO='+TOBL.GetString('GL_NUMERO')+' AND '+
           'BLO_INDICEG='+TOBL.GetString('GL_NUMERO')+' AND '+
           'BLO_NUMLIGNE='+TOBL.GetString('GL_NUMLIGNE')+' AND '+
           'BLO_N1<>0 AND BLO_N2=0 AND '+
           'BLO_UNIQUEBLO <> 0';
    QQ := OpenSql (Req,True,-1,'',true);
    if not QQ.eof then
    begin
      TOBOUV.LoadDetailDB('LIGNEFAC','','',QQ,false);
    end;
    ferme (QQ);
    for II := 0 to TOBOUV.detail.count -1 do
    begin
      TOBOO := TOBOUV.detail[II];
      StPrec:=TOBOO.GetValue('BLO_PIECEPRECEDENTE') ;
      //
      if (StPrec = '') or (TOBOO.GetValue('BLF_MTSITUATION') = 0) then Continue;
      DecodeRefPieceOUv(StPrec,CD) ;
      if CD.NaturePiece <> 'DBT' then Continue; // Si la pièce précédente n'est pas le devis --> on ne fait rien
      Req := 'SELECT BLF_QTECUMULEFACT,BLF_MTCUMULEFACT,BLF_MTMARCHE FROM LIGNEFAC '+
             'WHERE '+
             'BLF_NATUREPIECEG="'+CD.NaturePiece+'" AND '+
             'BLF_SOUCHE="'+CD.Souche +'" AND '+
             'BLF_NUMERO='+InttoStr(CD.NumeroPiece) +' AND '+
             'BLF_INDICEG='+InttoStr(CD.Indice) +' AND '+
             'BLF_UNIQUEBLO='+InttoStr(CD.UniqueBlo);
      //
      QQ := OpenSQL(Req,True,1,'',true);
      TRY
        if not QQ.eof then
        begin
          NewMtCum :=  QQ.fields[1].asFloat + TOBOO.getDouble('BLF_MTSITUATION');
          NewQteCum :=  QQ.fields[0].asFloat + TOBOO.getDouble('BLF_QTESITUATION');
          PourcentAvanc := 100.0;
          if QQ.fields[2].AsFloat <> 0 then
          begin
            PourcentAvanc := arrondi((NewMtCum / QQ.fields[2].AsFloat) * 100,2);
          end;
          req := 'UPDATE LIGNEFAC SET '+
                 'BLF_QTECUMULEFACT='+STRFPOINT(ARRONDI(NewQteCum,V_PGI.OkDecQ))+',' +
                 'BLF_QTEDEJAFACT='+STRFPOINT(ARRONDI(NewQteCum,V_PGI.OkDecQ))+','+
                 'BLF_MTDEJAFACT='+STRFPOINT(ARRONDI(NewMtCum,V_PGI.OkDecQ))+','+
                 'BLF_MTCUMULEFACT='+STRFPOINT(ARRONDI(NewMtCum,V_PGI.OkDecQ))+','+
                 'BLF_POURCENTAVANC='+STRFPOINT(ARRONDI(Pourcentavanc,V_PGI.OkDecQ))+' '+
                 'WHERE '+
                 'BLF_NATUREPIECEG="'+CD.NaturePiece+'" AND '+
                 'BLF_SOUCHE="'+CD.Souche +'" AND '+
                 'BLF_NUMERO='+InttoStr(CD.NumeroPiece) +' AND '+
                 'BLF_INDICEG='+InttoStr(CD.Indice) +' AND '+
                 'BLF_UNIQUEBLO='+InttoStr(CD.UniqueBlo);
          Result := (ExecuteSQL(Req)=0);
          if not result then exit;
        end;
      FINALLY
        ferme (QQ);
      end;
    end;
  end;


var REq,StPrec : string;
    QQ : TQuery;
    TOBPiece,TOBLIGNEFAC,TOBLIGNEFACDEV,TOBPiecePrec,TOBOuvrage,TOBL,TOBPieceReactiv : TOB;
    II,Nbr : Integer;
    NewQteCum,NewMtCum,Pourcent : double;
    CD : R_cledoc;
    OldQteCum ,OldMtCum, OldQteMar : Double;

begin
  Result := True;
  TOBPiece := TOB.Create ('LES LIGNES',nil,-1);
  TOBLIGNEFAC := TOB.Create ('LES LIGNESFAC',nil,-1);
  TOBOUVRAGE := TOB.Create ('LES LIGNESOUV',nil,-1);
  TOBLIGNEFACDEV := TOB.Create ('LES LIGNESFACDEV',nil,-1);
  TOBPiecePrec := TOB.Create ('LES LIGNES DEV',nil,-1);
  TOBPieceReactiv := TOB.Create ('LES PIECES A REACTIVER',nil,-1);
  //
  TRY
    Req := 'SELECT GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_TYPEARTICLE,GL_INDICENOMEN,'+
           'GL_NUMORDRE,GL_PIECEPRECEDENTE,GL_ARTICLE,GL_LIBELLE,BLF_MTSITUATION,BLF_QTESITUATION,BLF_MTMARCHE '+
           'FROM LIGNE '+
           'LEFT JOIN LIGNEFAC ON '+
           'BLF_NATUREPIECEG=GL_NATUREPIECEG AND '+
           'BLF_SOUCHE=GL_SOUCHE AND '+
           'BLF_NUMERO=GL_NUMERO AND '+
           'BLF_NUMORDRE=GL_NUMORDRE AND '+
           'BLF_UNIQUEBLO=0 '+
           'WHERE '+
           'GL_NATUREPIECEG="'+TOBSitpre.GetString('BST_NATUREPIECE')+'" AND '+
           'GL_SOUCHE="'+TOBSitpre.GetString('BST_SOUCHE')+'" AND '+
           'GL_NUMERO='+TOBSitpre.GetString('BST_NUMEROFAC')+' AND '+
           'GL_INDICEG=0 AND '+
           'GL_TYPELIGNE="ART" AND '+
           'BLF_QTESITUATION <> 0';
    QQ := OpenSQL(REq ,True,-1,'',true);
    if not QQ.eof then
    begin
      TOBPiece.LoadDetailDB('LIGNE','','',QQ,false);
    end;
    Ferme(QQ);
    //
    if TOBPiece.detail.count > 0 then
    begin
      for II := 0 to TOBPiece.detail.count -1 do
      begin
        TOBL:=TOBPiece.Detail[II] ;
        StPrec:=TOBL.GetValue('GL_PIECEPRECEDENTE') ;
        //
        if (StPrec = '') or (TOBL.GetValue('BLF_MTSITUATION') = 0)  then Continue;
        if (TOBL.GetDouble('BLF_MTMARCHE')=0) then Exit;
        DecodeRefPiece(StPrec,CD) ;
        if CD.NaturePiece <> 'DBT' then Continue; // Si la pièce précédente n'est pas le devis --> on ne fait rien
        //
        EnregPieceAReactiver (CD,TOBPieceReactiv);
        // -- > Devis
        Req:='SELECT GL_QTESIT,GL_QTEFACT,BLF_QTECUMULEFACT,BLF_MTCUMULEFACT '+
             'FROM LIGNE '+
             'LEFT JOIN LIGNEFAC ON '+
             'BLF_NATUREPIECEG=GL_NATUREPIECEG AND '+
             'BLF_SOUCHE=GL_SOUCHE AND '+
             'BLF_NUMERO=GL_NUMERO AND '+
             'BLF_INDICEG=GL_INDICEG AND '+
             'BLF_NUMORDRE=GL_NUMORDRE AND '+
             'BLF_UNIQUEBLO=0 '+
             'WHERE '+ WherePiece (CD,ttdLigne,true,True);
        QQ := OpenSql (Req,true,-1,'',true);
        OldQteCum := 0;
        OldMtCum := 0;
        OldQteMar := 0;
        if not QQ.eof then
        begin
          OldQteCum := QQ.fields[0].AsFloat;
          OldMtCum := QQ.fields[3].AsFloat;
          OldQteMar := QQ.fields[1].AsFloat;
        end;
        if OldQteMar <> 0 then
        begin
          //
          NewQteCum := QQ.fields[0].AsFloat + TOBL.GetDouble('BLF_QTESITUATION');
          NewMtCum := QQ.fields[3].AsFloat + TOBL.GetDouble('BLF_MTSITUATION');
          //
          if QQ.fields[1].AsFloat  <> 0 then Pourcent := arrondi((NewQteCum / QQ.fields[1].AsFloat) * 100,2)
                                        else Pourcent := 100;
          //
          REq := 'UPDATE LIGNE SET '+
                 'GL_QTEPREVAVANC='+STRFPOINT(ARRONDI(NewQteCum,V_PGI.okdecQ))+','+
                 'GL_QTESIT='+STRFPOINT(ARRONDI(NewQteCum,V_PGI.okdecQ))+','+
                 'GL_POURCENTAVANC='+STRFPOINT(ARRONDI(Pourcent,2))+' '+
                 'WHERE '+
                 WherePiece (CD,ttdLigne,true,True);
          if (ExecuteSQL(Req)=0) then
          begin
            Result := false;
            exit;
          end;
          //
          REq := 'UPDATE LIGNEFAC SET '+
                 'BLF_MTCUMULEFACT='+STRFPOINT(ARRONDI(NewMtCum,V_PGI.okdecQ))+','+
                 'BLF_MTDEJAFACT='+STRFPOINT(ARRONDI(NewMtCum,V_PGI.okdecQ))+','+
                 'BLF_QTECUMULEFACT='+STRFPOINT(ARRONDI(NewQteCum,V_PGI.okdecQ))+','+
                 'BLF_QTEDEJAFACT='+STRFPOINT(ARRONDI(NewQteCum,V_PGI.okdecQ))+','+
                 'BLF_POURCENTAVANC='+STRFPOINT(ARRONDI(Pourcent,2))+' '+
                 'WHERE '+
                 'BLF_NATUREPIECEG="'+CD.Naturepiece+'" AND '+
                 'BLF_SOUCHE="'+CD.Souche+'" AND '+
                 'BLF_NUMERO='+InttoStr(CD.NumeroPiece)+' AND '+
                 'BLF_INDICEG='+InttoStr(CD.INDICE)+' AND '+
                 'BLF_NUMORDRE='+InttoStr(CD.NumOrdre)+' AND '+
                 'BLF_UNIQUEBLO=0 ';
          //
          Nbr := ExecuteSQL(REq);           
          if nbr = 0 then
          begin
            Result := false;
            exit;
          end;
          if IsOuvrage(TOBL) then Result := TraiteLignesOuv(TOBL);
        end;
      end;
      if Result then
      begin
        Result := ReactivePiece (TOBSitPre);
      end;
    end;
  FINALLY
    TOBPiece.Free;
    TOBLIGNEFAC.Free;
    TOBOUVRAGE.free;
    TOBLIGNEFACDEV.free;
    TOBPiecePrec.free;
    TOBPieceReactiv.Free;
  END;
end;

function SupprimeSituation(TobPiece,TOBOuvrage,TobPieceRG,TOBBasesRG,TOBAcompte : Tob; DEV : RDevise ) : boolean;
Var Req, StPrec: string;
    i,NumSituation,IndiceSit : integer;
    TOBL,TOBSIT,TOBPIece_O,TOBL_O,TOBA,TOBFACTURE_O,TOBSitPre : TOB;
    CD : R_CleDoc ;
    QQ,Q,QQ1 : TQuery;
    Pourcent,reste : double;
    TypeFacturation,EtatLettrage,SSAffaire : string;
    NaturePiece: string;
    NewValeur : double;
    TOBPieceTraitDevis,TOBPTraitDevisL : TOB;
Begin
	TOBPieceTraitDevis := TOB.Create ('LES PIECETRAIT ORIG',nil,-1);
  TOBSitPre := TOB.create ('BSITUATIONS',nil,-1);
  //
  NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
  result := true;
  TypeFacturation := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
  //
  if (Pos(NaturePiece ,'FBT;DAC;FBP;BAC;')>0) then
  begin
  	if not RestitueMontantCautionUtilise (TOBPieceRG) then exit;
  end;
  //
  TOBPiece_O := TOB.Create ('LES LIGNES DEVIS',nil,-1);
  TOBFacture_O := TOB.Create ('LES LIGNES AVANC',nil,-1);
  //
  if ((TypeFacturation='AVA') or (TypeFacturation='DAC')) then
  begin
    if (Pos(NaturePiece,'FBT;DAC;FBP;BAC;')>0) and (TOBPiece.GetValue('GP_AFFAIREDEVIS')<>'') then
    begin
      for i:=0 to TOBPiece.Detail.Count-1 do
      begin
        TOBL:=TOBPiece.Detail[i] ;
        if (TOBL <> Nil) then
        begin
          StPrec:=TOBL.GetValue('GL_PIECEPRECEDENTE') ;
          if (StPrec <> '') and (TOBL.GetValue('GL_QTEFACT') <> 0) then
          begin
            DecodeRefPiece(StPrec,CD) ;
            if CD.NaturePiece <> 'DBT' then Continue; // Si la pièce précédente n'est pas le devis --> on ne fait rien
            //
            Req:='SELECT * FROM LIGNE WHERE '+ WherePiece (CD,ttdLigne,true,True);
            Q := OpenSql (Req,true,-1,'',true);
            if not Q.eof then
            begin
              TOBL_O := TOB.Create ('LIGNE',TOBPiece_O,-1);
              TOBL_O.SelectDB ('',Q);
              //
              NewValeur := TOBL_O.GetValue('GL_QTESIT') - TOBL.GetValue('GL_QTEFACT');
              if ((TOBL_O.GetValue('GL_QTESIT') < 0) and (newValeur > 0)) or ((TOBL_O.GetValue('GL_QTESIT') > 0) and (newValeur < 0)) then
              begin
                newValeur := 0;
              end;
              TOBL_O.PutValue('GL_QTESIT',NewValeur);
              TOBL_O.PutValue('GL_QTEPREVAVANC',NewValeur);
              //
              if TOBL_O.GetValue('GL_QTEFACT') <> 0 then
              begin
                Pourcent := arrondi((TOBL_O.GetValue('GL_QTESIT') / TOBL_O.GetValue('GL_QTEFACT')) * 100,2);
                TOBL_O.PutValue('GL_POURCENTAVANC',Pourcent);
              end;
            end;
            ferme (Q);
            // traitement de mise à jour des lignes fac du devis initial
            Req := 'SELECT * FROM LIGNEFAC WHERE '+ WherePiece (CD,ttdLigneFac,true,True);
            Q := OpenSql (Req,true,-1,'',true);
            if not Q.eof then
            begin
              TOBL_O := TOB.Create ('LIGNEFAC',TOBFacture_O,-1);
              TOBL_O.SelectDB ('',Q);
              //
              TOBL_O.PutValue('BLF_QTEDEJAFACT',TOBL.GetValue('BLF_QTEDEJAFACT'));
              TOBL_O.PutValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTEDEJAFACT'));
              //
              TOBL_O.PutValue('BLF_MTDEJAFACT',TOBL.GetValue('BLF_MTDEJAFACT'));
              TOBL_O.PutValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTDEJAFACT'));
              //
              //FV1 - 25/11/20015 - FS#1802 - BENETEAU - impossible de supprimer une facture message d'erreur
              if TOBL_O.GetValue('BLF_QTEMARCHE') <> 0 then
              begin
                Pourcent := arrondi((TOBL_O.GetValue('BLF_QTECUMULEFACT') / TOBL_O.GetValue('BLF_QTEMARCHE')) * 100,2);
                TOBL_O.PutValue('BLF_POURCENTAVANC',Pourcent);
              end;
            end;
            ferme (Q);
            //
            if VALEUR(TOBL.GetValue('GLC_NATURETRAVAIL'))<10 then
            begin
              TOBPTraitDevisL := FindPieceTraitOrig (TOBPieceTraitDevis,CD,TOBL.GetValue('GL_FOURNISSEUR'));
              if TOBPTraitDevisL = nil then
              begin
                TOBPTraitDevisL := ChargePieceTraitOrig(TOBPieceTraitDevis,CD,TOBL.GetValue('GL_FOURNISSEUR'));
              end;
              if TOBPTraitDevisL <> nil then
              begin
                TOBPTraitDevisL.PutValue('BPE_MONTANTFAC',TOBPTraitDevisL.GetValue('BPE_MONTANTFAC')-TOBL.GetValue('BLF_TOTALTTCDEV'));
              end;
            end;
            //
            if isOuvrage (TOBL) then
            begin
              AjoutSuprAvancementOuv (TOBPieceTraitDevis,TOBL,TOBOuvrage,TOBfacture_O,true);
            end;

          end;
        end;
      end;
      // Suppression de la situation
      req := 'SELECT * FROM BSITUATIONS WHERE '+
             'BST_NATUREPIECE="'+TOBPIece.GetValue('GP_NATUREPIECEG')+'" AND '+
             'BST_SOUCHE="'+TOBPIece.GetVAlue('GP_SOUCHE')+'" AND '+
             'BST_NUMEROFAC="'+inttoStr(TOBPiece.GetValue('GP_NUMERO'))+'"';
      QQ := OpenSql(Req,true,-1,'',true);
      if not QQ.eof then
      begin
        TOBSit := TOB.create ('BSITUATIONS',nil,-1);
        TOBSIt.selectdb ('',QQ);
        NumSituation := TOBSIT.GetInteger('BST_NUMEROSIT');
        indiceSit := TOBSIT.GetInteger('BST_INDICESIT');
        SSAffaire := TOBSIT.GetString('BST_SSAFFAIRE');
        if not TOBSIt.DeleteDB (false) then Result := false;
        if (Result) and (pos(TOBPIece.GetValue('GP_NATUREPIECEG'),'FBP;DAP')>0) and (indiceSit>0) then
        begin
          // on reactive la precedente piece provisoire....
          Req := 'SELECT * FROM BSITUATIONS WHERE '+
                      'BST_SSAFFAIRE="'+SSAffaire+'" AND '+
                      'BST_NUMEROSIT='+InttoStr(NumSituation)+' AND '+
                      'BST_INDICESIT='+InttoStr(indiceSit-1)+' AND '+
                      'BST_VIVANTE="-"';
          QQ1 := OpenSQL(Req,True,1,'',true);
          if not QQ1.eof then
          begin
            TOBSitPre.selectdb('',QQ1);
          end;
          ferme (QQ1);
        end else if (Result) and (pos(TOBPIece.GetValue('GP_NATUREPIECEG'),'FBT;DAC')>0) and (indiceSit>0) then
        begin
          if ExisteSql ('SELECT 1 FROM BSITUATIONS WHERE '+
                 'BST_SSAFFAIRE="'+SSAffaire+'" AND '+
                 'BST_NUMEROSIT='+InttoStr(NumSituation)+' AND '+
                 'BST_INDICESIT=(SELECT MAX(BST_INDICESIT) FROM BSITUATIONS WHERE BST_SSAFFAIRE="'+SSAffaire+'" AND BST_NUMEROSIT='+InttoStr(NumSituation)+') AND '+
                 'BST_VIVANTE="-"') then
          begin
            Req := 'UPDATE BSITUATIONS '+
                   'SET BST_VIVANTE="X" '+
                   'WHERE '+
                   'BST_SSAFFAIRE="'+SSAffaire+'" AND '+
                   'BST_NUMEROSIT='+InttoStr(NumSituation)+' AND '+
                   'BST_INDICESIT=(SELECT MAX(BST_INDICESIT) FROM BSITUATIONS WHERE BST_SSAFFAIRE="'+SSAffaire+'" AND BST_NUMEROSIT='+InttoStr(NumSituation)+') AND '+
                   'BST_VIVANTE="-"';
            ExecuteSql (Req);
        end;
        end;
        TOBSit.free;
      end;
      ferme (QQ);
    end;
  end else if (Pos(NaturePiece, 'FBT;FBP')>0) and (TOBPiece.GetValue('GP_AFFAIREDEVIS')<>'') then
  begin
  	// facture directe
    for i:=0 to TOBPiece.Detail.Count-1 do
    begin
      TOBL:=TOBPiece.Detail[i] ;
      if (TOBL <> Nil) then
      begin
        StPrec:=TOBL.GetValue('GL_PIECEPRECEDENTE') ;
        if IsSousDetail(TOBL) then continue; // on ne traite pas les lignes de gestion des sous détail
        if (StPrec <> '') and (TOBL.GetValue('GL_QTEFACT') <> 0) then
        begin
          DecodeRefPiece(StPrec,CD) ;
          Req:='SELECT * FROM LIGNE WHERE '+ WherePiece (CD,ttdLigne,true,True);
          Q := OpenSql (Req,true,-1,'',true);
          if not Q.eof then
          begin
            TOBL_O := TOB.Create ('LIGNE',TOBPiece_O,-1);
            TOBL_O.SelectDB ('',Q);
            //
            NewValeur := TOBL_O.GetValue('GL_QTESIT') - TOBL.GetValue('GL_QTEFACT');
            if ((TOBL_O.GetValue('GL_QTESIT') < 0) and (newValeur > 0)) or ((TOBL_O.GetValue('GL_QTESIT') > 0) and (newValeur < 0)) then
            begin
              newValeur := 0;
            end;
            TOBL_O.PutValue('GL_QTESIT',NewValeur);
            TOBL_O.PutValue('GL_QTEPREVAVANC',NewValeur);
          end;
          ferme (Q);
          // traitement de mise à jour des lignes fac du devis initial
          Req := 'SELECT * FROM LIGNEFAC WHERE '+ WherePiece (CD,ttdLigneFac,true,True);
          Q := OpenSql (Req,true,-1,'',true);
          if not Q.eof then
          begin
            TOBL_O := TOB.Create ('LIGNEFAC',TOBFacture_O,-1);
            TOBL_O.SelectDB ('',Q);
            //
            TOBL_O.PutValue('BLF_QTEDEJAFACT',TOBL_O.GetValue('BLF_QTEDEJAFACT')-TOBL.GetValue('BLF_QTESITUATION'));
            //
            TOBL_O.PutValue('BLF_MTDEJAFACT',TOBL_O.GetValue('BLF_MTDEJAFACT')-TOBL.GetValue('BLF_MTSITUATION'));
            //
            //FV1 - 25/11/20015 - FS#1802 - BENETEAU - impossible de supprimer une facture message d'erreur
            if TOBL_O.GetValue('BLF_QTEMARCHE') <> 0 then
            begin
            Pourcent := arrondi((TOBL_O.GetValue('BLF_QTEDEJAFACT') / TOBL_O.GetValue('BLF_QTEMARCHE')) * 100,2);
            TOBL_O.PutValue('BLF_POURCENTAVANC',0);
          end;
          end;
          ferme (Q);

          if VALEUR(TOBL.GetValue('GLC_NATURETRAVAIL'))<10 then
          begin
            TOBPTraitDevisL := FindPieceTraitOrig (TOBPieceTraitDevis,CD,TOBL.GetValue('GL_FOURNISSEUR'));
            if TOBPTraitDevisL = nil then
            begin
              TOBPTraitDevisL := ChargePieceTraitOrig(TOBPieceTraitDevis,CD,TOBL.GetValue('GL_FOURNISSEUR'));
            end;
            if TOBPTraitDevisL <> nil then
            begin
              TOBPTraitDevisL.PutValue('BPE_MONTANTFAC',TOBPTraitDevisL.GetValue('BPE_MONTANTFAC')-TOBL.GetValue('BLF_TOTALTTCDEV'));
            end;
          end;
            //
          if isOuvrage (TOBL) then
          begin
            AjoutSuprAvancementOuv (TOBPieceTraitDevis,TOBL,TOBOuvrage,TOBfacture_O,false);
          end;
        end;
      end;
    end;
  end;

  if (Result) and (TOBPieceTraitDevis.detail.count > 0 ) then
  BEGIN
  	if not TOBPieceTraitDevis.UpdateDB (false) then result := false;
  END;
  if (Result) and (TOBPIece_O.detail.count > 0 ) then
  BEGIN
  	if not TOBPiece_O.UpdateDB (false) then result := false;
  END;
  if (Result) and (TOBFACTURE_O.detail.count > 0 ) then
  BEGIN
    (*
    i := 0;
    repeat
      if TOBFACTURE_O.detail[i].getValue('BLF_MTDEJAFACT')=0 then
      begin
        TOBFACTURE_O.detail[i].DeleteDB  (false);
        TOBFACTURE_O.Detail[i].Free;
      end else inc(I);
    until i >= TOBfacture_O.detail.count;
    *)
  	if not TOBFACTURE_O.UpdateDB (false) then result := false;
  END;
  //
  TOBPiece_O.free;
  TOBFacture_O.free;
  TOBPieceTraitDevis.Free;
  //
  if ((TypeFacturation<>'AVA') And (TypeFacturation<>'DAC')) then exit;
  if Result then
  begin
    if (pos(TOBPIece.GetValue('GP_NATUREPIECEG'),'FBP;DAP')>0) and
       (TOBSitPre.GetInteger('BST_NUMEROFAC')<>0) then
    begin
      Result := ReactiveSituationProvisoirePrec(TOBSitPre);
    end;
  end;

  if Result Then
  begin
  	TOBA := TOB.Create('AFFAIRE',nil,-1);
    if TOBPiece.getValue('GP_AFFAIREDEVIS')<> '' then
    begin
      Req := 'SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBPiece.GetValue('GP_AFFAIREDEVIS')+'"';
      QQ := OpenSql (Req,true,-1,'',true);
      if not QQ.eof then
      begin
      	TOBA.SelectDb ('',QQ);
      	TOBA.putValue('AFF_ACOMPTEREND',TOBA.GetValue('AFF_ACOMPTEREND')- TOBPiece.GetValue('GP_ACOMPTEDEV'));
        TOBA.UpdateDB;
        (*
        reste := QQ.findfield ('AFF_ACOMPTEREND').AsFloat;
        Reste := Reste - TOBPiece.GetValue('GP_ACOMPTEDEV');
        Req := 'UPDATE AFFAIRE SET AFF_ACOMPTEREND=';
        if Reste > 0 then Req := REq +StringReplace(StrF00(Reste,V_PGI.OkDecV),',','.',[rfReplaceAll])
                     else Req := REq + '0';
        req := req +' WHERE AFF_AFFAIRE="'+TOBPiece.GetValue('GP_AFFAIREDEVIS')+'"';
        if (ExecuteSQL(req) < 0) then result := false;
        *)
      end;
      QQ.close;
    end;
    TOBA.free;
  end;
  if (Result) and (TOBAcompte.detail.count > 0) then
  begin
    for I := 0 to TOBAcompte.detail.count -1 do
    begin
    	if TOBAcompte.detail[I].GetValue('GAC_PIECEPRECEDENTE')<> '' then
      begin
      	DecodeRefpiece (TOBAcompte.detail[I].GetValue('GAC_PIECEPRECEDENTE'),CD);
        RestitueAcompte(CD,TOBAcompte.detail[I].GetValue('GAC_JALECR'),TOBAcompte.detail[I].GetValue('GAC_NUMECR'),
        								TOBAcompte.detail[I].GetValue('GAC_MONTANTDEV'),TOBAcompte.detail[I].GetValue('GAC_MONTANT'));
      end;
      QQ := OpenSql ('SELECT E_ETATLETTRAGE FROM ECRITURE WHERE '+GetClefecriture(TOBACompte.detail[I]),true,-1,'',true);
      EtatLettrage := QQ.findfield('E_ETATLETTRAGE').AsString;
      ferme (QQ);
      if EtatLettrage = 'TL' then
      BEGIN
        req := 'UPDATE ECRITURE SET E_ETATLETTRAGE="AL" WHERE '+GetClefEcriture (TOBACompte.detail[I]);
        if ExecuteSQL (req) < 0 then result := false;
      END;
    end;
  end;
  TOBSitPre.free;
End;


procedure GetSituationSuivante (TOBpiece : TOB; var NCd : R_CLEDOC);
Var Req: string;
    Num : integer;
    Q : TQuery ;
    SqlPlus,TypeD : string;
Begin
  if not ExisteSQL('SELECT BST_NUMEROFAC FROM BSITUATIONS WHERE '+
      'BST_SSAFFAIRE="'+ TOBPiece.GetValue('GP_AFFAIREDEVIS') +'" '+
      'AND BST_NUMEROSIT > ('+
      'SELECT BST_NUMEROSIT FROM BSITUATIONS WHERE '+
      'BST_SSAFFAIRE="'+ TOBPiece.GetValue('GP_AFFAIREDEVIS') +'" AND '+
      'BST_NATUREPIECE="'+TOBPiece.GeTString('GP_NATUREPIECEG')+'" AND '+
      'BST_NUMEROFAC='+InttoStr(TOBPiece.Getvalue('GP_NUMERO'))+') ') then Exit;
  //    
  // récupération numéro de facture de la dernière situation
  Req:='SELECT BST_NUMEROFAC,BST_NATUREPIECE,BST_SOUCHE FROM BSITUATIONS WHERE '+
      'BST_SSAFFAIRE="'+ TOBPiece.GetValue('GP_AFFAIREDEVIS') +'" '+
      'AND BST_NUMEROSIT > ('+
      'SELECT BST_NUMEROSIT FROM BSITUATIONS WHERE '+
      'BST_SSAFFAIRE="'+ TOBPiece.GetValue('GP_AFFAIREDEVIS') +'" AND '+
      'BST_NATUREPIECE="'+TOBPiece.GeTString('GP_NATUREPIECEG')+'" AND '+
      'BST_NUMEROFAC='+InttoStr(TOBPiece.Getvalue('GP_NUMERO'))+') AND '+
      'BST_VIVANTE="X" '+
      'ORDER BY BST_NUMEROSIT ASC';
  Q:=OpenSQL(Req,TRUE,1,'',true);
  if Not Q.EOF then
  begin
    Ncd.NumeroPiece  := Q.Fields[0].AsInteger;
    Ncd.NaturePiece  := Q.Fields[1].AsString;
    Ncd.Souche       := Q.Fields[2].AsString;
    Ncd.Indice       := 0;
  end;
  Ferme(Q) ;
end;

// Fonction qui vérifie si la facture passée en argument dans la TOB
// est la dernière situation.
Function DerniereSituation(TOBPiece : TOB; Quefacture : Boolean=false) : Boolean ;
Var Req: string;
    Num : integer;
    Q : TQuery ;
    SqlPlus,TypeD : string;
Begin
  result := true;
  if TOBPiece.GetValue('AFF_GENERAUTO')='DIR' then
  begin
    Exit;
  end;
  if QueFacture then SqlPlus := 'AND BST_NATUREPIECE IN ("FBT","DAC") ';
  // récupération numéro de facture de la dernière situation
  Req := 'SELECT BST_NUMEROFAC,BST_NATUREPIECE,BST_NUMEROSIT FROM BSITUATIONS ';
  Req := req + ' WHERE BST_SSAFFAIRE="'+ TOBPiece.GetValue('GP_AFFAIREDEVIS') + '" ';
  req := req + SqlPlus+' AND BST_VIVANTE="X" ';
  if TOBPiece.GetString('NUMEROSIT') <> '' then Req := req + ' AND BST_NUMEROSIT > '+ TOBPiece.GetString('NUMEROSIT');
  Req := Req + ' ORDER BY BST_NUMEROSIT DESC';
  Q:=OpenSQL(Req,TRUE,1,'',true);
  if Not Q.EOF then
  begin
    Num := Q.Fields[0].AsInteger;
    TypeD := Q.Fields[1].AsString;
    if (TOBPiece.getString('GP_NATUREPIECEG')<>TypeD) or (Num <> TOBPiece.GetValue('GP_NUMERO')) then result:=false;
  end;
  Ferme(Q) ;
end;


// Fonction qui vérifie si l'étude provient d'un document EXCEL
Function Origine_EXCEL(TOBPiece : TOB) : Boolean ;
Var Req, RefPiece: string;
    Q : TQuery ;
Begin
result := False;

if TOBPiece.GetValue('GP_NATUREPIECEG') <> VH_GC.AFNatProposition Then Exit;

// récupération Type de document d'origine
RefPiece:=EncodeRefPiece(TOBPiece) ;
Req:='SELECT BDE_TYPE FROM BDETETUDE WHERE BDE_PIECEASSOCIEE="'+RefPiece+'"';
Q:=OpenSQL(Req,TRUE,-1,'',true);
if Not Q.EOF then result:=True;
Ferme(Q) ;

end;

Procedure BtMajSDP(CleDoc : R_CleDoc;EnTOBUniquement : boolean = false);
Var Q : TQuery;
    St : String;
    TOBL : TOB;
    NumOrdre : Integer;
Begin
	if not EnTOBUniquement then
  begin
    //Maj TABLE LIGNE
    Q:=OpenSQL ( 'SELECT GL_ACTIONLIGNE, GL_NUMORDRE FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,True,False),true,-1,'',true);
    if Not Q.EOF then
    begin
      St:=Q.FindField('GL_ACTIONLIGNE').AsString;
      NumOrdre:=Q.FindField('GL_NUMORDRE').AsInteger;
    end else
    begin
      St := '';
      numordre:=0;
    end;
    Ferme(Q);

    if St = 'SDP' then St :=''
    else St := 'SDP';

    ExecuteSQL('UPDATE LIGNE SET GL_ACTIONLIGNE="'+St+'" WHERE '+WherePiece(CleDoc,ttdLigne,True,False)) ;
    TOBL:=TheTOB.FindFirst(['GL_NUMORDRE'],[NumOrdre],False) ;
  end
  else
    begin
    TOBL:=TheTOB.FindFirst(['GL_NUMLIGNE'],[CleDoc.Numligne],False) ;
    end;
  if TOBL<>Nil then
  Begin
  	if TOBL.GetValue('GL_ACTIONLIGNE')='SDP' Then st := '' else st := 'SDP';
    TOBL.PutValue('GL_ACTIONLIGNE',St) ;
    if EnTOBUniquement then
    begin
      if TheImpressionViaTOB <> nil then TheImpressionViaTOB.PreparationEtat;
    End;
  end;
End;

// Fonction appelée depuis l'impression du décisionnel stock
// traitement temporaire en attendant le developpement de la fonctionnalité
Procedure BtMajTenueStock(CleDoc : R_CleDoc);
Var Q : TQuery;
    TenueStock, stArticle, stDepot : String;
    Nb : integer;
    Qte : Double;
Begin
  // Maj TABLE LIGNE : bascule du code tenue en stock
  // et maj du stock disponible
  Q:=OpenSQL ( 'SELECT GL_ARTICLE, GL_DEPOT, GL_TENUESTOCK, GL_QTEFACT FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,True,False),true,-1,'',true);
  if Not Q.EOF then
  begin
    stArticle:=Q.FindField('GL_ARTICLE').AsString;
    stDepot:=Q.FindField('GL_DEPOT').AsString;
    TenueStock:=Q.FindField('GL_TENUESTOCK').AsString;
    Qte:=Q.FindField('GL_QTEFACT').AsFloat;
    if TenueStock = 'X' then
    begin
      TenueStock:='-';
      Qte := Qte * -1.0;
    end else
    begin
      TenueStock:='X';
    end;
    Nb:=ExecuteSQL('UPDATE LIGNE SET GL_TENUESTOCK="'+TenueStock+'" WHERE '+WherePiece(CleDoc,ttdLigne,True,False)) ;
    if Nb>0 then
    begin
      ExecuteSQL('UPDATE DISPO SET GQ_RESERVECLI=GQ_RESERVECLI + "'+
                 StrfPoint(Qte) + '" WHERE ' +
                 'GQ_ARTICLE="' + stArticle + '" AND ' +
                 'GQ_DEPOT="' + stDepot + '" AND ' +
                 'GQ_CLOTURE="-"') ;
    end;
  end;
  Ferme(Q);
End;

function GetMontantFraisDetail (TOBPiece : TOB; Var ExistFg : boolean) : double;
var NatureFrais : string;
		Req : string;
    Q : TQuery;
    cledoc : R_CLEDOC;
begin
  result := 0;
	existFg :=  ExisteFraisChantier (TOBPiece,Cledoc);
  if Existfg then
  begin
  	Req:='SELECT GP_TOTALHTDEV FROM PIECE WHERE '+WherePiece (cledoc,ttdpiece,true) ;
    Q := OpenSql (Req,True,-1,'',true);
    if not Q.eof then result:=Q.Fields[0].AsFloat;
    Ferme (Q);
  end;
end;

function GetTauxFg (TOBporcs : TOB) : double;
var i : integer;
begin
  result:=0;
  for i:=0 to TOBPorcs.Detail.Count-1 do
  BEGIN
    if TOBPorcs.Detail[i].GetValue('GPT_FRAISREPARTIS') = 'X' then
    Begin
      result := result + TOBPorcs.Detail[i].GetValue('GPT_POURCENT');
    End;
  END;
end;

(*
// fonction retournant le coefficient de FG d'une étude
// avec prise en compte des frais généraux et des frais de chantier
Function CalculMontantRevient (TOBPiece, TOBporcs : TOB; DEV:RDevise;InclusStInFG : boolean; var ExistFG : boolean) : double ;
Var  Revient, TotalAchat, TotalFrais, TauxFrais : double ;
    i : integer ;
    Q : TQuery ;
    Req, NatureFrais : String;
    cledoc : r_cledoc;
BEGIN
result:=1;

// Récupération du total de l'étude des frais
// => lire le document de même numéro pour la nature de frais associée
//modif brl 080604 : NatureFrais:=GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_NATPIECEFRAIS');
TotalFrais := GetMontantFraisDetail (TOBPiece, ExistFg);
if (ExistFg) and (TotalFrais = 0) then
begin
	BTPSupprimePieceFrais (TOBpiece);
  ExistFg := false;
end;



// Calcul du pourcentage total de frais répartis
// => Dans TOBPorcs cumul des pourcentage de frais ayant le code FRAISREPARTIS à TRUE
//    puis calcul du coefficient global que réprésentent ces frais sur le total HT du document
TauxFrais := GetTauxFg (TOBporcs);
ExistFG := (tauxFrais <> 0);

  // Calcul déboursé sec hors sous-traitance :
TotalAchat:=TOBPiece.GetValue('DPA_FG');

if (TotalAchat <> 0) and (TauxFrais < 100) then
   begin
   Revient := (TotalAchat+TotalFrais) * (1+(TauxFrais/100));
//---   METHODE LEVAUX incompréhensible !!!
//   Revient := ((TotalAchat+TotalFrais)*100) / (100-TauxFrais);
//---
   result := Revient/TotalAchat;
//   result := Arrondi(Revient,V_PGI.OkdecP);
   end;

if result=0 then result:=1;

end;
*)

Function BTPSupprimePieceFrais (CleDoc : R_CLEDOC):boolean;
//var Naturefrais : string;
begin
result := false;
(* modif brl 080604 :
NatureFrais:=GetInfoParPiece(cledoc.NaturePiece ,'GPP_NATPIECEFRAIS');
if NatureFrais = '' then exit;
CleDoc.NaturePiece:=NatureFrais;
*)
if cledoc.naturepiece <> 'FRC' then CleDoc.NaturePiece:='FRC';
if ExistePiece (cledoc) then BTPSupprimePiece (cledoc);
end;

Function BTPSupprimePieceFrais (TOBPiece : TOB):boolean;
var cledoc : r_cledoc;
begin
	if TOBPiece.GetValue('GP_PIECEFRAIS') = '' then
  begin
    FillChar(CleDoc, Sizeof(CleDoc), #0);
    CleDoc.NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
    CleDoc.DatePiece := TOBPiece.GetValue('GP_DATEPIECE');
    CleDoc.Souche := TOBPiece.GetValue('GP_SOUCHE');
    CleDoc.NumeroPiece := TOBPiece.GetValue('GP_NUMERO');
    CleDoc.Indice := TOBPiece.GetValue('GP_INDICEG');
    BTPSupprimePieceFrais (cledoc);
  end else
  begin
  	DecodeRefPiece (TOBpiece.getValue('GP_PIECEFRAIS'),cledoc);
		BTPSupprimePieceFrais (cledoc);
	end;

end;

Function BTPSupprimePiece (CleDoc : R_CLEDOC;AvecAdresses:boolean=true):boolean;
var nb : Integer;
		Sql : string;
    AffaireDevis : string;
    QQ : TQuery;
begin
  result := true;
  QQ := OpenSql ('SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False),true,-1,'',true);
  if not QQ.eof then
  begin
  	AffaireDevis :=  QQ.FindField  ('GP_AFFAIREDEVIS').AsString;
    ferme (QQ);
    if AffaireDevis <> '' then
    begin
      Sql := 'DELETE FROM AFFAIRE WHERE AFF_AFFAIRE="'+AffaireDevis+'"';
      nb := ExecuteSQL(Sql);
      if nb < 0 then
      begin
        V_PGI.IoError:=oeunknown ;
        result := false;
        Exit ;
      end;
    end;
  end else
  begin
  	Ferme (QQ);
  end;
  Sql := 'DELETE FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False);
  nb := ExecuteSQL(Sql) ;
  if nb < 0 then
     begin
     V_PGI.IoError:=oeunknown ;
     result := false;
     Exit ;
     end;
  Sql := 'DELETE FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False);
  nb := ExecuteSQL(Sql) ;
  if nb < 0 then
     begin
     V_PGI.IoError:=oeunknown ;
     result := false;
     Exit ;
     end;
  Sql := 'DELETE FROM LIGNECOMPL WHERE '+WherePiece(CleDoc,ttdLigneCompl,False);
  nb := ExecuteSQL(Sql) ;
  if nb < 0 then
     begin
     V_PGI.IoError:=oeunknown ;
     result := false;
     Exit ;
     end;
  Sql := 'DELETE FROM PIEDBASE WHERE '+WherePiece(CleDoc,ttdPiedBase,False);
  nb := ExecuteSQL(Sql) ;
  if nb < 0 then
     begin
     V_PGI.IoError:=oeunknown ;
     result := false;
     Exit ;
     end;
  Sql := 'DELETE FROM PIEDECHE WHERE '+WherePiece(CleDoc,ttdEche,False);
  nb := ExecuteSQL(Sql) ;
  if nb < 0 then
     begin
     V_PGI.IoError:=oeunknown ;
     result := false;
     Exit ;
     end;
  Sql := 'DELETE FROM PIEDPORT WHERE '+WherePiece(CleDoc,ttdPorc,False);
  nb := ExecuteSQL(Sql) ;
  if nb < 0 then
     begin
     V_PGI.IoError:=oeunknown ;
     result := false;
     Exit ;
     end;
  Sql:= 'DELETE FROM ACOMPTES WHERE '+WherePiece(CleDoc,ttdAcompte,False);
  nb := ExecuteSQL(Sql) ;
  if nb < 0 then
     begin
     V_PGI.IoError:=oeunknown ;
     result := false;
     Exit ;
     end;
  //
  Sql := 'DELETE FROM LIGNEOUVPLAT WHERE '+WherePiece(CleDoc,ttdOuvrageP,False);
  nb := ExecuteSQL(Sql) ;
  //
  Sql := 'DELETE FROM LIGNEOUV WHERE '+WherePiece(CleDoc,ttdOuvrage,False);
  nb := ExecuteSQL(Sql) ;
  if nb < 0 then
     begin
     V_PGI.IoError:=oeunknown ;
     result := false;
     Exit ;
     end;
  Sql := 'DELETE FROM LIENSOLE WHERE '+WherePiece(CleDoc,ttdLienOle,False);
  nb := ExecuteSQL(Sql) ;
  if nb < 0 then
     begin
     V_PGI.IoError:=oeunknown ;
     result := false;
     Exit ;
     end;
  Sql := 'DELETE FROM LIGNEPHASES WHERE '+WherePiece(CleDoc,ttdLignePhase,False);
  nb := ExecuteSQL(Sql) ;
  //
  Sql := 'DELETE FROM PIECERG WHERE '+WherePiece(CleDoc,ttdretenuG,False);
  nb := ExecuteSQL(Sql) ;
  //
  Sql := 'DELETE FROM PIEDBASERG WHERE '+WherePiece(CleDoc,ttdBaseRG,False);
  nb := ExecuteSQL(Sql) ;
  //
  Sql := 'DELETE FROM BSITUATIONS WHERE '+WherePiece(CleDoc,ttdSit,False);
  nb := ExecuteSQL(Sql) ;
  //
  Sql := 'DELETE FROM PIECETRAIT WHERE '+WherePiece(CleDoc,ttdPieceTrait,False);
  nb := ExecuteSQL(Sql) ;
  //
  Sql := 'DELETE FROM BTPARDOC WHERE '+WherePiece(CleDoc,ttdParDoc,False);
  nb := ExecuteSQL(Sql) ;
  if nb < 0 then
     begin
     V_PGI.IoError:=oeunknown ;
     result := false;
     Exit ;
     end;
  if AvecAdresses then
     begin
       if GetParamSoc('SO_GCPIECEADRESSE') then
       begin
       	 Sql := 'DELETE FROM PIECEADRESSE WHERE '+WherePiece(CleDoc,ttdPieceAdr,False);
         nb := ExecuteSQL(Sql) ;
         if nb < 0 then
            begin
            V_PGI.IoError:=oeunknown ;
            result := false;
            end;
       end else
       begin
       	 Sql := 'DELETE FROM ADRESSES WHERE '+WherePiece(CleDoc,ttdAdresse,False);
         nb := ExecuteSQL(Sql) ;
         if nb < 0 then
            begin
            V_PGI.IoError:=oeunknown ;
            result := false;
            end;
       end;
     end;
  Sql := 'DELETE FROM BLIENEXCEL WHERE BLX_PIECEASSOCIEE="'+CodeLaPiece(cledoc)+'"';
  nb := ExecuteSQL(Sql);
  if nb < 0 then
     begin
     V_PGI.IoError:=oeunknown ;
     result := false;
     end;
end;

Function Lexical_RechArt ( GS : THGrid ; TitreSel,NaturePieceG,DomainePiece,SelectFourniss : String; var CodeArticle : string ) : boolean ;
Var sw,sWhere,CodeArt,StChamps,Libel, Mot : String ;
BEGIN
  result := false;
  CodeArticle := '';
  // recup du libellé saisi
  CodeArt:=GS.Cells[GS.Col,GS.Row] ;

  // recup conditions liées à la nature de pièce
  sWhere:='' ;
  sW:=FabricWhereNatArt(NaturePieceG,DomainePiece,SelectFourniss) ;
  if sw<>'' then sWhere:=sWhere+' AND '+sw ;

  // Préparation de la chaine de recherche
  // - On remplace les points, les virgules, les tirets et les traits de soulignement par des blancs
  // - On remplace ensuite les blancs par des point-virgules
  // - On intégre dans la chaine le libellé complet puis tous les mots de plus de 2 lettres
  StChamps := CodeArt;
  StChamps := StringReplace(StChamps,'.',' ',[rfReplaceAll]); // remplacement des points par des blancs
  StChamps := StringReplace(StChamps,',',' ',[rfReplaceAll]); // remplacement des virgules par des blancs
  StChamps := StringReplace(StChamps,'-',' ',[rfReplaceAll]); // remplacement des tirets par des blancs
  StChamps := StringReplace(StChamps,'_',' ',[rfReplaceAll]); // remplacement des soulignés par des blancs
  StChamps := StringReplace(StChamps,' ',';',[rfReplaceAll]); // remplacement des blancs par un point-virgule

  Libel:=Trim(CodeArt)+';';
  Repeat
    Mot:=ReadTokenSt(StChamps) ;
    if Length(Mot)>4 then Libel:=Libel+Mot+';' ;
  Until (StChamps='') ;
  Libel := FabricWhereToken(Libel,'GA_LIBELLE');     // Création clause where
  Libel := StringReplace(Libel,'="',' LIKE @%',[rfReplaceAll]); // remplacement des blancs par un point-virgule
  Libel := StringReplace(Libel,'"','%"',[rfReplaceAll]); // remplacement des blancs par un point-virgule
  Libel := StringReplace(Libel,'@','"',[rfReplaceAll]); // remplacement des blancs par un point-virgule
  StChamps:='XX_WHERE='+sWhere+';XX_WHERE1= AND '+Libel;

  // Lancement de la recherche
  CodeArt:=AGLLanceFiche('BTP','BTARTICLE_RECH','','',StChamps);

  // traitement du code retour : si le code est à -1 (cas du bouton STOP), on renvoit False
  // et on positionne la ligne au nombre maximum de lignes présentes
  if CodeArt= '-1' then
  Begin
    Result:=False;
    GS.Row:=GS.RowCount-1;
    Exit;
  End;

  if CodeArt<>'' then
  BEGIN
  	Result:=True ;
    GS.Cells[SG_RefArt,GS.Row]:=CodeArt ;
    CodeArticle := CodeArt;
  END ;

END ;

function CodeLaPiece (cledoc : r_cledoc) : string;
var std : string;
begin
StD:=FormatDateTime('ddmmyyyy',Cledoc.DatePiece ) ;
Result:=StD+';'+CleDoc.NaturePiece +';'+cledoc.Souche +';'
          +IntToStr(cledoc.NumeroPiece)+';'+inttostr(cledoc.Indice)+';;'
end;

Procedure ImprimePieceBTP (cledoc : r_cledoc; Bapercu : boolean; Modele, Req ,TypeFacturation : string; DGD : boolean = false; ImpressionViaTOB : TImprPieceViaTOB = nil);
var SQL, SQLOrder : string;
    TableauSituation : String;
    Q : TQuery ;
    UneTOB : TOB;
    TheModele, ModeleR,ModeleS : string;
    NbSSTrait : integer;
    MultiEtat : boolean;
begin
  MultiEtat := false;
  application.bringtofront;


  //FV1 - 26/07/2016 : Impossible d'imprimer des commandes fournisseurs ou des préparation de commandes 
  if ImpressionViaTOB <> nil then
  begin
  if ImpressionViaTOB.TobSsTraitantImp.Detail.Count > 0 then
  begin
    NbSSTrait := ImpressionViaTOB.TobSsTraitantImp.Detail[0].detail.Count - 1;
  end
  else
    NbSSTrait := 0;
  end;

  TableauSituation:='-';

  if (Pos(CleDoc.NaturePiece ,'FBT;DAC;FBP;BAC;')>0) then
  Begin
    //récupération du code impression tableau totalisation de situation
    SQL:='SELECT BPD_IMPTABTOTSIT FROM BTPARDOC WHERE BPD_NATUREPIECE="'+CleDoc.NaturePiece+'" AND BPD_SOUCHE="'+CleDoc.Souche+'" AND BPD_NUMPIECE='+IntToStr(CleDoc.NumeroPiece) ;
    Q:=OpenSQL(SQL,TRUE,-1,'',true);
    if Not Q.EOF then
      TableauSituation:=Q.Fields[0].AsString
    else
    Begin
      Ferme(Q);
      SQL:='SELECT BPD_IMPTABTOTSIT FROM BTPARDOC WHERE BPD_NATUREPIECE="'+CleDoc.NaturePiece+'" AND BPD_SOUCHE="'+CleDoc.Souche+'" AND BPD_NUMPIECE=0' ;
      Q:=OpenSQL(SQL,TRUE,-1,'',true);
      if Not Q.EOF then TableauSituation:=Q.Fields[0].AsString
                   else TableauSituation:='-';
    End;
    Ferme(Q) ;
  End;
  if (TableauSituation = 'X') or ( NbSSTrait > 5) then
  Begin
    if (Modele <> '') or (NbSSTrait > 5) then MultiEtat := true;
    //
    Repeat
      RelancerEtat:=False;
    	if MultiEtat  then StartPdfBatch; // inutile si pas de modèle et impression du tableau seul

      if (ImpressionViaTOB = nil) OR (not ImpressionViaTOB.Usable) then
      begin
        if LanceEtat('E','GPJ',Modele,True,False,False,Nil,Req,'',False)<0 then V_PGI.IoError:=oeUnknown ;
      end else if Modele <> '' then
      begin
        TheModele:=Modele;
        UneTOB := ImpressionViaTOB.TOBAIMPRIMER;
        if LanceEtatTOB ('E','GPJ',TheModele,UneTOB,Bapercu,false,false,nil,'','',false) < 0 then V_PGI.IoError:=oeUnknown ;
      end;

      if V_PGI.IoError=oeOK then
      Begin
        SQLOrder:='';
        if (ImpressionViaTOB = nil) OR (not ImpressionViaTOB.Usable) then
        begin
          ModeleR:=GetParamsoc('SO_BTETATSIR');
          if ModeleR='' then ModeleR:=Copy(Modele,1,2)+'R';
          SQL:=RecupSQLEtat('E','GPJ',ModeleR,'','','',' WHERE '+WherePiece(CleDoc,ttdPiece,False),SQLOrder) ;
          SQL:=SQL+' '+SQLOrder;
      		if multiEtat then LastPdfBatch ;
          LanceEtat('E','GPJ',ModeleR,True,False,False,Nil,Trim(SQL),'',False);
        end else
        begin
        	if DGD then
          begin
          	ModeleR:=GetParamsoc('SO_BTETATDGDRDIR');
          end else
          begin
            if TypeFacturation = 'AVA' then
            begin
            if (Pos(CleDoc.NaturePiece ,'FBP;BAC;')>0) then
            begin
              ModeleR:=GetParamsoc('SO_BTETATFBR');
            end;
            if ModeleR = '' then
            begin
          	ModeleR:=GetParamsoc('SO_BTETATSIRDIR');
          end;
            end else
            begin
              ModeleR:=GetParamsoc('SO_BTETATRECAPDIRECT');
          end;
          end;
          //FV1 - 09/02/2018 : FS#2930 - TEST BL : Pb d'impression dans le cas d'une situation avec plus de 5 sous-traitants
          if TableauSituation = 'X' then
          begin
            if ModeleR='' then ModeleR:=Copy(Modele,1,2)+'R';
            UneTOB := ImpressionViaTOB.TOBRECAP;
            if (MultiEtat) and ( NbSSTrait <= 5) then LastPdfBatch ; // inutile si pas de modèle et impression du tableau seul
            if LanceEtatTOB ('E','GPJ',ModeleR,UneTOB,Bapercu,false,false,nil,'','',false) < 0 then V_PGI.IoError:=oeUnknown ;
            // ---
          end;
          if (NbSSTrait > 5) then
          begin
            ModeleS:=GetParamsoc('SO_BTETATSSTRAITANCE');
            if ModeleS = '' then break;
            UneTOB := ImpressionViaTOB.TobSsTraitantImp;
    			  if (MultiEtat) then LastPdfBatch ; // inutile si pas de modèle et impression du tableau seul
            if LanceEtatTOB ('E','GPJ',ModeleS,UneTOB,Bapercu,false,false,nil,'','',false) < 0 then V_PGI.IoError:=oeUnknown ;
        end;
        end;
      end
      else
      Begin
        if MultiEtat then CancelPDFBatch ;  // inutile si pas de modèle et impression du tableau seul
      end;
      {$IFNDEF EAGLCLIENT}
      // RH le 14/10/2003 ne passe pas à la compile WebAcces
      if MultiEtat then PreviewPDFFile('',GetMultiPDFPath,True); // inutile si pas de modèle et impression du tableau seul
      {$ENDIF}
      //PrintPDF (GetMultiPDFPath,'','');
    Until (Not RelancerEtat) or (V_PGI.IoError=oeUnknown);
  End else
  Begin
    Repeat
      RelancerEtat:=False;
      if (ImpressionViaTOB = nil) OR (not ImpressionViaTOB.Usable) then
      begin
        if LanceEtat('E','GPJ',Modele,BApercu,False,False,Nil,Req,'',False) < 0 then V_PGI.IoError:=oeUnknown ;
      end else
      begin
        TheModele:=Modele;
        If TheModele = '' Then TheModele := ImpressionViaTOB.GetModeleAssocie(Modele);
        if TheModele = '' then break;
        UneTOB := ImpressionViaTOB.TOBAIMPRIMER;
        if LanceEtatTOB ('E','GPJ',TheModele,UneTOB,Bapercu,false,false,nil,'','',false) < 0 then
        V_PGI.IoError:=oeUnknown ;
      end;
    Until (Not RelancerEtat) or (V_PGI.IoError=oeUnknown);
  End;
end;

Function GetMiniAvancAcompte (CodessAff : string) : Integer;
var req : string;
    Q : TQuery;
begin
result := 0;
(* POUR PLUS TARD ???
Req:='SELECT AFF_MINAVANCACC FROM AFFAIRE WHERE AFF_AFFAIRE="'+Codessaff+'"' ;
Q:=OpenSQL(Req,TRUE) ;
if Not Q.EOF then result := Q.Fields[0].AsInteger;
Ferme(Q) ;
*)
end;

function ExisteDocAchatLigne (theLigne : String) : boolean;
var Req : string;
begin
    result := false;
    REq := 'SELECT GL_NUMERO FROM LIGNE WHERE GL_PIECEORIGINE="'+TheLigne+'" AND GL_NATUREPIECEG IN ('+GetPieceAchat (true,true,true,true)+')';
    if ExisteSql (req) then result := true;
end;

function ExisteDocumentAchat (TOBpiece : TOB) : boolean;
var Indice : integer;
    TOBL : TOB;
begin
  result := false;
  for Indice := 0 to TOBpiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[indice];
    if ExisteDocAchatLigne (EncodeRefPiece(TOBL)) then
    begin
      Result := true;
      exit;
    end;
  end;
end;

function ControleChantierBTP (TOBPiece : TOB; Mode : BTTraitChantier) : boolean;
var Indice : integer;
    Req : String;
    QQ : Tquery;
    MessageRet : string;
    Cledoc : r_cledoc;
    refPiece : string;
begin
  result := true;
  if TOBPIece.GetValue('GP_NATUREPIECEG') = 'BCE' then
  begin
    if Mode = BTTSuppress then
    begin
    	RefPiece := EncodeRefPiece (TOBPiece,0,false);
      RefPiece := Copy(RefPiece,1,length(RefPiece)-1); // pour enlever le ; de fin
      Req := 'SELECT DISTINCT GL_NUMERO FROM LIGNE WHERE GL_NATUREPIECEG IN ("PBT") AND GL_PIECEPRECEDENTE LIKE "'+
      			 RefPiece+'%"';
      QQ := OpenSql (Req,True,-1,'',true);
      TRY
        if not QQ.eof then
        begin
        	Result := false;
					PgiBox (TraduireMemoire('Suppression impossible. Une prévision de chantier existe'),'Suppression de pièce');
          exit;
				end;
      FINALLY
      	ferme (QQ);
      END;
    end;
  end;

  if (GetParamSoc('SO_AFNATAFFAIRE') = TOBPIece.GetValue('GP_NATUREPIECEG')) then
  BEGIN
// controle si chantier terminé
	 if Mode = BTTModif then
   begin
   	if ControleChantiertermineBTP (TOBpiece,TOBPiece.getValue('GP_AFFAIRE'),true,false) then
    begin
    	PgiError ('IMPOSSIBLE : Le chantier est clôturé');
    	result := false;
      exit;
    end;
   end;
// Nouvelle gestion
    Req := 'SELECT AFF_PREPARE FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBPiece.getValue('GP_AFFAIREDEVIS')+'"';
    QQ := OpenSql (Req,true,-1,'',true);
    TRY
      if not QQ.eof then
      begin
        if QQ.FindField ('AFF_PREPARE').AsString = 'X' then
        begin
          Result := false;
          if Mode = BTTSuppress then PgiBox (TraduireMemoire('Suppression impossible. Une prévision de chantier existe'),'Suppression de pièce')
          else if Mode = BTTmodif then PgiBox (TraduireMemoire('La prévision de chantier ou la contre étude existe déjà'),'Création de prévision de chantier')
          else if Mode = BTTContrEtud then PgiBox (TraduireMemoire('La prévision de chantier ou la contre étude existe déjà'),'Génération de contre étude');
          exit;
        end;
      end;
    FINALLY
      Ferme (QQ);
    END;
// verif si devis facture
    if Mode = BTTSuppress then
    begin
    	RefPiece := EncodeRefPiece (TOBPiece,0,false);
      RefPiece := Copy(RefPiece,1,length(RefPiece)-1); // pour enlever le ; de fin
      Req := 'SELECT DISTINCT GL_NUMERO FROM LIGNE WHERE GL_NATUREPIECEG IN ("FBT","DAC","ABT") AND GL_PIECEPRECEDENTE LIKE "'+
      			 RefPiece+'%"';
      QQ := OpenSql (Req,True,-1,'',true);
      TRY
        if not QQ.eof then
        begin
        	Result := false;
					PgiBox (TraduireMemoire('Suppression impossible. Une Facture existe'),'Suppression de pièce');
          exit;
				end;
      FINALLY
      	ferme (QQ);
      END;
    end;
// ---------
  END;

  if (TOBPIece.GetValue('GP_NATUREPIECEG')='CBT') then
  BEGIN
    if ExisteDocumentAchat (TOBpiece) then
    begin
      PgiBox (TraduireMemoire('Suppression impossible. Un document d''achat existe pour cette pièce'),'Suppression de pièce');
      result := false;
    end;
    exit;
  END;
  (* GESTION CONSO *)
  if not ControlePiecePhases (TOBPiece,MessageRet) then
  begin
    PgiBox (TraduireMemoire(MessageRet),'Suppression de pièce');
    result := false;
  end;

end;

procedure ReajusteQteReste (TOBGenere : TOB);
var indice : integer;
    TOBL : TOB;
    CleDoc : R_CLEDOC;
    QQ : Tquery;
    req : string;
begin
  for indice := 0 to TOBgenere.detail.count -1 do
  begin
    TOBL := TOBGenere.detail[indice];
    if TOBL.GetValue('GL_PIECEPRECEDENTE') = '' then continue;
    DecodeRefPiece (TOBL.GetValue('GL_PIECEPRECEDENTE'),Cledoc);
    if ExecuteSql ('UPDATE LIGNE SET GL_QTERESTE=0, GL_MTRESTE=0 WHERE '+wherePiece(cledoc,ttdLigne,true,true)) < 0 then //--- GUINIER ---
        V_PGI.IOError := oeUnknown ;
  end;
end;


function ISPrepaLivFromAppro : boolean;
begin
  result := (GetParamSoc ('SO_BTOPTGENREA') = 'GEN');
end;
(*
function ISGenereLivFromAppro : boolean;
begin
  result := (GetParamSoc ('SO_BTOPTLANCEREA') = 'GBL');
end;
*)
function LivraisonVisible : boolean;
begin
  result := GetParamSoc ('SO_BTLIVRVISIBLE');
end;
//
{
procedure ChargeTobPieces (LaTob : TOB;Naturepiece : string);
var Requete : string;
    QQ : Tquery;
begin
  Requete := 'SELECT * FROM PIECE WHERE GP_NATUREPIECEG="'+Naturepiece+'" AND GP_VIVANTE="X"';
  QQ := OpenSql (requete,true);
  if not QQ.eof Then LaTob.LoadDetailDB ('PIECE','','',QQ,false);
  ferme (QQ);
end;
}
procedure GenereLivraisonClients (ThePieceGenere : TOB;Action : TActionFiche;transfoPiece,DuplicPiece: boolean;DemandeDate: boolean = true;CompteRendu : boolean=true; TOBresult : TOB= nil) ;
var Livraison : TGenereLivraison;
    indice : integer;
    TOBL : TOB;
    IOOK : TIOErr;
begin
	//
	//if (Action <> taModif) or (DuplicPiece = True) or (TransfoPiece = false) then exit;
	if (Action = taModif) and (TransfoPiece = false) and (not DuplicPiece)  then exit;
  //
	// Si c'est pas une reception ou une facture fournisseur ---> Get out
	if Pos(ThePieceGenere.GetValue('GP_NATUREPIECEG'),GetPieceAchat (false,false,false)) = 0 then exit;
  Livraison := TGenereLivraison.create ;
  Livraison.CompteRendu := CompteRendu; // mode silencieux
  TRY
  	Livraison.ThePieceGenere := ThePieceGenere;
    // Positionnement des infos
    Livraison.Duplication := DuplicPiece;
    Livraison.Transformation := transfoPiece;
    Livraison.Action := Action;
    Livraison.TOBresult := TOBresult;
    //
		Livraison.DecortiqueBesoinLivraison;
    Livraison.ChargeTobPieces (Livraison.TOBPieces);
    Livraison.datePriseEnCompte := ThePieceGenere.getValue('GP_DATEPIECE');
    if Livraison.TOBPieces.detail.count > 0 then
    begin
      for Indice := 0 to Livraison.TOBPieces.detail.count -1 do
      begin
        TOBL := Livraison.TOBPieces.detail[Indice];
        // FQ 12066
        if TOBL.getValue('GL_INDICENOMEN')<>0 then TOBL.PutValue('GL_INDICENOMEN',0);
        // --
        TOBL.putvalue('GL_PIECEPRECEDENTE',EncodeRefPiece(TOBL));
        TOBL.putvalue('GL_PIECEORIGINE',EncodeRefPiece(TOBL));
        TOBL.addChampSupValeur('ORIGINE',EncodeRefPiece(TOBL,0,false));
        TOBL.addChampSupValeur('REFINTERNE','');
        if TOBL.fieldExists ('GP_REFINTERNE') then
        begin
        	TOBL.PutValue('REFINTERNE',TOBL.GetValue('GP_REFINTERNE'));
        end;
      end;
    end;
		if Livraison.TheLignesALivrer.detail.count > 0 then Livraison.AjusteLesLignes;
    IOOK := Transactions(Livraison.GenereThepieces,1);
  FINALLY
    FreeAndNil(Livraison);
  END;
end;

// Chargement de valeurs particulières pour les natures de pièces communes
// avec d'autres applications.
Procedure InitParPieceBTP ;
Var TOBNat : TOB ;
BEGIN
// Pour la nature REAPPROS, on force l'action de fermeture à TRANSFORMATION
TOBNat:=VH_GC.TOBParPiece.FindFirst(['GPP_NATUREPIECEG'],['REA'],False) ;
if TOBNat<>Nil then
  Begin
  TOBNat.PutValue('GPP_ACTIONFINI','TRA') ;
  TOBNat.UpdateDB;
  End;
TOBNat:=VH_GC.TOBParPiece.FindFirst(['GPP_NATUREPIECEG'],['EEX'],False) ;
if TOBNat<>Nil then
  Begin
  TOBNat.PutValue('GPP_LISTESAISIE','BTSAISIEEX') ;
  TOBNat.UpdateDB;
  End;
TOBNat:=VH_GC.TOBParPiece.FindFirst(['GPP_NATUREPIECEG'],['SEX'],False) ;
if TOBNat<>Nil then
  Begin
  TOBNat.PutValue('GPP_LISTESAISIE','BTSAISIEEX') ;
  TOBNat.UpdateDB;
  End;
{
// Pour la nature BL ClIENT, on force le calcul de rupture de stock à STOCK PHYSIQUE
TOBNat:=VH_GC.TOBParPiece.FindFirst(['GPP_NATUREPIECEG'],['BLC'],False) ;
if TOBNat<>Nil then
  Begin
  TOBNat.PutValue('GPP_CALCRUPTURE','PHY') ;
  TOBNat.UpdateDB;
  End;
}
END ;
{ TGenereLivraison }

constructor TGenereLivraison.create;
begin
TOBPieces := TOB.Create ('LES LIGNES',nil,-1);
TheLignesALivrer := TOB.Create ('LES LIGNES RECEP A LIVRER DIRECTEMENT',nil,-1);
TheAffaires := TOB.create ('LES AFFAIRES',nil,-1);
TOBClients := TOB.Create ('LES TIERS (CLIENTS)',nil,-1);
TOBLivrDirecte := TOB.Create ('LES LIGNES LIVRABLES',nil,-1);
end;

function TGenereLivraison.AddDetail(TOBLD : TOB) : string;
var Indice : integer;
		TOBLDD : TOB;
begin
	result := ' AND (GL_NUMORDRE IN (';
	for Indice := 0 to TOBLD.detail.count -1 do
  begin
  	TOBLDD := TOBLD.detail[Indice];
    if Indice > 0 then result := result + ',';
    result := result +  IntToStr(TOBLDD.GetValue('NUMORDRE'));
  end;
  result := result + '))';
end;

function TGenereLivraison.MakeWhere : string;
var Indice : integer;
		LaChaineWhere,LaChaineQteRelle : string;
    TOBLD : TOB;
begin
  for Indice := 0 to TOBLivrDirecte.detail.count -1 do
  begin
    TOBLD := TOBLivrDirecte.detail[Indice];
    if Indice = 0 Then
    begin
    LaChaineWhere :='((GL_NATUREPIECEG="'+TOBLD.GEtValue('NATURE')+'") AND '+
                    '(GL_NUMERO="'+InttoStr(TOBLD.GEtValue('NUMERO'))+'") AND '+
                    '(GL_INDICEG="'+IntToStr(TOBLD.GEtValue('INDICE'))+'")'+AddDetail(TOBLD)+')'
    end else
    begin
    LaChaineWhere := LAChaineWhere + 'OR ((GL_NATUREPIECEG="'+TOBLD.GEtValue('NATURE')+'") AND '+
                    '(GL_NUMERO="'+InttoStr(TOBLD.GEtValue('NUMERO'))+'") AND '+
                    '(GL_INDICEG="'+IntToStr(TOBLD.GEtValue('INDICE'))+'")'+addDetail(TOBLD)+')'
    end;
    LaChaineWhere := '('+LaChaineWhere+') ' +
                     'AND (GL_QTERESTE <> 0 OR GL_MTRESTE <> 0) AND GL_IDENTIFIANTWOL=-1'; //--- GUINIER ---
  end;
  Result := ' WHERE ' + LaChaineWhere;
end;

procedure TGenereLivraison.ChargeTobPieces (LaTob : TOB);
var Requete : string;
    QQ : Tquery;
    II,JJ,KK : integer;
    BTOB,LTOB,TOBL : TOB;
begin
	if TOBLivrDirecte.detail.count = 0 then exit;
(* precedemment *)
//  Requete := 'SELECT *,PIE.GP_REFINTERNE,PIE.GP_NUMADRESSEFACT,PIE.GP_NUMADRESSELIVR FROM LIGNE '+
//             'LEFT JOIN PIECE PIE ON (PIE.GP_NATUREPIECEG=GL_NATUREPIECEG) and (PIE.GP_NUMERO=GL_NUMERO)' +
//             ' WHERE GL_NATUREPIECEG="'+Naturepiece+'" AND GL_QTERESTE > 0 AND GL_IDENTIFIANTWOL=-1';
(* maintenant *)
	requete := MakeSelectLigneBtp (true,True) + MakeWhere;
  QQ := OpenSql (requete,true,-1,'',true);
  if not QQ.eof Then LaTob.LoadDetailDB ('LIGNE','','',QQ,false);
  ferme (QQ);
  for II := 0 to TOBLivrDirecte.detail.count -1 do
  begin
    BTOB := TOBLivrDirecte.detail[II];  // un besoin de chanteir
    for JJ := 0 TO BTOB.detail.count -1 do
    begin
      LTOB := BTOB.detail[JJ];
      TOBL := LaTob.FindFirst(['GL_NUMERO','GL_NUMORDRE'],[BTOB.GetInteger('NUMERO'),LTOB.getInteger('NUMORDRE')],true);
      if TOBL <> nil then
      begin
        TOBL.SetDouble('GL_QTEFACT',LTOB.GetDouble('QTERECEP'));
      end;
    end;
  end;
end;

destructor TGenereLivraison.destroy;
begin
  inherited;
  TOBPieces.free;
	TheLignesALivrer.free;
  TheAffaires.free;
  TOBClients.free;
	TOBLivrDirecte.free;
end;


procedure TGenereLivraison.MiseAjourPieceprec (TOBpiece : TOB);
var Indice : integer;
    TobPrecTraite : TOB;
    TOBL : TOB;
    cledoc : R_CleDoc;
    TOBD : TOB;
begin
  TOBPrecTraite := TOB.create ('LES DOCUMENTS TRAITES',nil,-1);
  TRY
    // Mise à jour des lignes
    for Indice := 0 to TOBpiece.detail.count -1 do
    begin
      TOBL := TOBpiece.detail[Indice];
      DecodeRefPiece (TOBL.GetValue('GL_PIECEPRECEDENTE'),Cledoc);
      TOBD := TOBPrecTraite.findFirst (['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
                                  [cledoc.NaturePiece,cledoc.Souche,cledoc.NumeroPiece,cledoc.indice],true);
      if TOBD = nil then
      begin
        TOBD := TOB.create ('PIECE',TobPrecTraite,-1);
        TOBD.putvalue ('GP_NATUREPIECEG',cledoc.Naturepiece);
        TOBD.putvalue ('GP_SOUCHE',cledoc.Souche);
        TOBD.putvalue ('GP_NUMERO',cledoc.NumeroPiece);
        TOBD.putvalue ('GP_INDICEG',cledoc.Indice);
        TOBD.LoadDB (true);
      end;
      //UPdateLignePrec (Cledoc);
    end;

    // mise à jour des pièces
    for Indice := 0 to TOBPrecTraite.detail.count -1 do
    begin
      TOBD := TOBPrecTraite.detail[Indice];
      if not ISAliveLine (TOBD) then
      begin
        TOBD.PutValue('GP_VIVANTE','-');
        TOBD.UpdateDB;
      end;
    end;
  FINALLY
    TOBPrecTraite.free;
  END;
end;


procedure TGenereLivraison.SetInfosClient(TOBL: TOB);
var TOBClient : TOB;
		QQ : TQuery;
    TypeCom,CodeRepres : string;
begin
  TypeCom := GetInfoParPiece('LBT', 'GPP_TYPECOMMERCIAL');
  TOBClient := TOBClients.findFirst(['T_TIERS'],[TOBL.GetValue('GL_TIERS')],true);
  if TOBClient = nil then
  begin
  	TOBClient := TOB.Create ('TIERS',TOBClients,-1);
    QQ := OpenSql ('SELECT T_DEVISE,T_FACTUREHT,T_REGIMETVA,T_COMPTATIERS,T_DOMAINE,T_ZONECOM FROM TIERS WHERE T_TIERS="'+TOBL.GetValue('GL_TIERS')+'" AND T_NATUREAUXI="CLI"',true,-1,'',true);
    if not QQ.eof Then
    begin
      TOBClient.selectDb ('',QQ,true);
    end;
    ferme (QQ);
  end;
//  TOBL.PutValue('GP_MODEREGLE',TOBClient.GetValue('T_MODEREGLE'));
  TOBL.PutValue('GL_REGIMETAXE', TOBClient.GetValue('T_REGIMETVA'));
  if TOBClient.FieldExists('T_REPRESENTANT') then CodeRepres := TOBClient.GetValue('T_REPRESENTANT') else CodeRepres := '';
  TOBL.PutValue('GL_REPRESENTANT', CodeRepres);
end;

procedure TGenereLivraison.DecortiqueBesoinLivraison;
var Indice,NbDoc : integer;
		TOBL,TheLigneAlivrer : TOB;
    PieceOrigine,PieceCde,PiecePrecedente : string;
    Cledoc,CledocPrec : R_Cledoc;
    QteRecep : double;
    RatioVente,RatioStk : double;
begin
	// Initialisations
  TOBLivrDirecte.ClearDetail;
  NbDoc := 0;
	TheLignesALivrer.ClearDetail;
  PieceCde := GetNaturePieceCde(false);
  // --
  For indice := 0 to ThePieceGenere.detail.count -1 do
  begin
    TOBL := ThePieceGenere.detail[Indice];

    if TOBL.GetValue('GL_TYPELIGNE') <> 'ART' Then Continue;
    if TOBL.GetValue('GL_ARTICLE')='' Then Continue;

    PieceOrigine := TOBL.GetValue ('GL_PIECEORIGINE');
    PiecePrecedente := TOBL.GetValue ('GL_PIECEPRECEDENTE');
    if Pieceprecedente <> '' then
    begin
    	DecodeRefPiece (PiecePrecedente,cledocPrec);
    end;
    // si transformation d'une réception en facture fournisseur alors on degage
    {
    if (Pos(CledocPrec.NaturePiece ,GetPieceAchat (false,false,false,false)) > 0) and (Not Duplication)  then continue;
    If DejaLivre (PiecePrecedente) then continue; // si la piece precedente est deja livré on ne continue pas
    }
    if (Pos(CledocPrec.NaturePiece ,GetPieceAchat (false,false,false,true)) > 0) and
    	 (Not Duplication) and
       DejaLivre (PiecePrecedente) then continue; // si la piece precedente est deja livré on ne continue pas
    // --
    if PieceOrigine <> '' then
    begin
    	DecodeRefPiece (PieceOrigine,cledoc);
    end;
    if (TOBL.GetValue('GL_AFFAIRE') <> '') AND
    	 ((PieceOrigine='') or (IsSamePieceOrFromAchat(TOBL))) AND
    	 (TOBL.GetValue('GL_IDENTIFIANTWOL')=-1) then
    begin
    	// provenance des achats et livraison sur chantier par fourniseur
      if TOBL.GEtValue('GL_QTEFACT') <> 0 then
      begin
        TheLigneALivrer := TOB.Create ('LIGNE',TheLignesALivrer,-1);
        TheLigneALivrer.Dupliquer (TOBL,false,true); // on mémorise la ligne
        TransformeAchatEnVente (TheLigneALivrer);
      end;
    end else if ((PieceOrigine<>'') and
    				 		(Pos(Cledoc.NaturePiece ,PieceCde)>0)) then
    begin
			if TOBL.GetValue('GL_QTEFACT') <> 0 then
      begin
        // -- cas ou dans la commande la réception n'est pas a faire sur le chantier -- */
        if (PiecePrecedente <> '') and (PiecePrecedente <> PieceOrigine) then
        begin

          if TOBL.GetInteger('GL_IDENTIFIANTWOL')<> -1 then continue;
        end;
        // ------------------
        RatioVente := GetRatio(TOBL, nil, trsVente);
        //
        RatioStk := Ratiovente ;
        QteRecep:=ARRONDI(TOBL.GetValue('GL_QTEFACT') * RatioStk,V_PGI.OkDecQ);
        //
				AddLivraisonDirecte (cledoc,QteRecep);
      end;
    end;
  end;
end;

function IsSamePieceOrFromAchat(TOBL: TOB): boolean;
var cledoc : R_CLEDOC;
begin
	result := false;
	if TOBL.GetValue('GL_PIECEORIGINE') = '' then Exit;
  //
  DecodeRefPiece (TOBL.GetValue('GL_PIECEORIGINE'),Cledoc);
  if ((Cledoc.NaturePiece = TOBL.GetValue('GL_NATUREPIECEG')) and
  	 (Cledoc.Souche = TOBL.GetValue('GL_SOUCHE')) and
     (Cledoc.NumeroPiece = TOBL.GetValue('GL_NUMERO')) and
     (Cledoc.Indice = TOBL.GetValue('GL_INDICEG'))) OR
     ( Pos(CleDoc.NaturePiece,GetNaturePieceCde(false)) = 0 ) then result := true;
end;

procedure TGenereLivraison.SetInfosLivraison(TOBL : TOB);
var Affaire : string;
		TobA : TOB;
    QQ : TQuery;
    OldDevise : string;
    OldTauxDevise : double;
begin
	Affaire := TOBL.GetValue('GL_AFFAIRE');
  TOBA := TheAffaires.FindFirst (['AFF_AFFAIRE'],[Affaire],True)  ;
  if TOBA = nil then
  begin
    TOBA := TOB.Create ('AFFAIRE',TheAffaires,-1);
    QQ := OpenSql ('SELECT AFF_AFFAIRE,AFF_TIERS FROM AFFAIRE WHERE AFF_AFFAIRE="'+affaire+'"',True,-1,'',true);
    if not QQ.eof then
    begin
      TOBA.SelectDb ('',QQ,true);
    end;
    Ferme (QQ);
  end;
  TOBL.PutValue('GL_PIECEORIGINE',EncodeRefPiece (TOBL));
  TOBL.PutValue('GL_PIECEPRECEDENTE','');
  TOBL.PutValue ('GL_TIERS',TOBA.GetValue('AFF_TIERS'));
  OldDevise := TOBL.GetValue('GL_DEVISE');
  OldTauxDevise := TOBL.GetValue('GL_TAUXDEV');
  SetInfosClient (TOBL);
end;

procedure TGenereLivraison.AjusteLesLignes;
var TOBL : TOB;
		Indice : integer;
begin
	for indice := 0 to TheLignesALivrer.detail.count -1 do
  begin
  	TOBL := TheLignesALivrer.detail[Indice];
    SetInfosLivraison(TOBL);
  end;
end;

procedure TGenereLivraison.GenereThepieces;

  function CCInPiece (TOBpieces : TOB) : boolean;
  begin
    result := (TOBPieces.findFirst(['GL_NATUREPIECEG'],['CC'],true)<>nil);
  end;

  procedure decomposePieces (TOBPieces,TOBPCC,TOBPCBT : TOB);
  begin
    repeat
      if TOBPieces.detail[0].GetString('GL_NATUREPIECEG')='CBT' then TOBPieces.detail[0].ChangeParent(TOBPCBT,-1)
                                                                else TOBPieces.detail[0].ChangeParent(TOBPCC,-1);

    until TOBpieces.detail.count =0;
  end;

var Indice : integer;
    TOBpiece : TOB;
    TOBPCC, TOBPCBT : TOB;
begin
  TOBPCC := TOB.Create ('LES PIECES CC',nil,-1);
  TOBPCBT := TOB.Create ('LES PIECES CBT',nil,-1);
  TRY
    if TOBPieces.detail.count > 0 then
    begin
        if CCInPiece (TOBPieces) then
        begin
          decomposePieces (TOBPieces,TOBPCC,TOBPCBT);
          if TOBPCC.detail.count > 0 then
          begin
            CreerPiecesFromLignes(TOBPCC,'CCTOLIVCLI',DatePriseEnCompte,True,false,TOBresult);
          end;
          if TOBPCBT.detail.count > 0 then
          begin
            CreerPiecesFromLignes(TOBPCBT,'CBTTOLIVC',DatePriseEnCompte,True,false,TOBresult);
          end;
        end else
        begin
          CreerPiecesFromLignes(TobPieces,'CBTTOLIVC',DatePriseEnCompte,True,false,TOBresult);
        end;
      end;
    if (TheLignesALivrer.detail.count > 0) and (V_PGI.IOError = OeOk) then
    begin
      CreerPiecesFromLignes(TheLignesALivrer,'RFOTOLIVC',DatePriseEnCompte, CompteRendu,false,TOBResult);
    end;
  FINALLY
  END;
end;

function ISAliveLine (TOBpiece : TOB) : boolean;
var Requete : String;
    QQ : TQuery;
    cledoc : R_cledoc;
begin
  result := false;
  cledoc := TOB2CleDoc (TOBPiece);
  Requete := 'SELECT GL_NUMERO FROM LIGNE WHERE '+WherePiece (Cledoc,ttdLigne,false) +
             ' AND GL_VIVANTE="X" AND (GL_QTERESTE <> 0 OR GL_MTRESTE <> 0) AND GL_TYPELIGNE="ART"'; // --- GUINIER ---
  QQ := OpenSql (requete,true,-1,'',true);
  if not QQ.eof then result := true;
  ferme (QQ);
end;

procedure InitCalcDispo;
begin
	TOBDispoLoc := TOB.Create ('LES STOCKS',nil,-1);
	FromDispoGlob := true;
end;

procedure InitCalcDispoArticle;
begin
	FromDispoArt := true;
end;

procedure FinCalcDispo;
begin
	if TOBDispoLoc <> nil then
  begin
  	TOBDispoLoc.free;
  	TOBDispoLoc := nil;
  end;
	FromDispoGlob := false;
	FromDispoArt := false;
end;

procedure AddChampsSupDispoAff (TOBDEPOT : TOB);
begin
	TOBDepot.AddChampSupValeur  ('_LIVRE',0,false);
end;

Function GetQtelivrable (TOBL : TOB;gestionConso : TGestionPhase; TOBART : TOB = nil; PriseSurStock : boolean=true) : double;
var QTeRecep,QTeDejaLivre,QteLivrable,QteDispo,QteResteAlivrer,QTe,QteTemp : double;
    requete,Moderecep : string;
    QQ : TQuery;
    TOBDepot,TOBA,TOBDR,TOBR,TOBLS : TOB;
    TypeArticle : string;
    RefOrigine : string;
    Created : boolean;
    FUA,FUS,FUV : double;
    MTAchat,PUA,CoefPAPR : double;
    NumMouv : double;
    QteRecepFour,QTeDejaLivrCli,CoefUAUS,COEFUSUV : double;
    Indice : integer;
    UneLoc : TOB;
    Ok_ReliquatMt : Boolean;
begin
  Created := false;
  PUA := 0;
  UneLoc := nil;

  if TOBL.getValue('GL_DPA') <> 0 then
    CoefPAPR := TOBL.getValue('GL_DPR')/TOBL.getValue('GL_DPA')
  else
    CoefPAPR := 0;

  if TOBART <> nil then
    TOBA := TOBART
  else
  begin
    TOBA := TOB.Create ('ARTICLE',nil,-1);
    TOBA.PutVAlue('GA_ARTICLE',TOBL.GetValue('GL_ARTICLE'));
    TOBA.LoadDB;
    Created := true;
  end;

  result := 0;

  if TOBL.GetValue('GL_TYPELIGNE') <> 'ART' then Exit;

  TypeArticle := TOBL.GetValue('GL_TYPEARTICLE');
  Ok_ReliquatMt :=  CtrlOkReliquat(TOBL, 'GL');

  if (TypeArticle = 'PRE') and (IsPrestationInterne(TOBL.GetValue('GL_ARTICLE'))) then
  begin
    // --- GUINIER ---
    if Ok_ReliquatMt then Result := TOBL.GetValue('GL_MTRESTE');
    exit;
  end;
  //
  refOrigine := TOBL.GetValue('GL_PIECEPRECEDENTE');
  if RefOrigine = '' then exit;
  TOBR := TOB.Create ('LES LIGNES',nil,-1);
  // initialisation
  QteResteAlivrer := 0;
  QTeRecep := 0;
  QTeDejaLivre := 0;
  QteLivrable := 0;
  QteDispo := 0;
  MtAchat := 0;
  //
  FUV := RatioMesure('PIE', TOBL.GetValue('GL_QUALIFQTEVTE'));
  if FUV = 0 then FUV := 1;
  FUS := RatioMesure('PIE', TOBL.GetValue('GL_QUALIFQTESTO'));
  if FUS = 0 then FUS := 1;
  COEFUSUV := TOBL.GetDouble('GL_COEFCONVQTEVTE');
  COEFUAUS := TOBL.GetDouble('GL_COEFCONVQTE');

  // calcule les quantités réceptionnées des fourniseurs
  requete := 'SELECT GL_NATUREPIECEG,GL_DATEPIECE,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_NUMORDRE,'+
  					 'GL_QTESTOCK AS QTERECEP,GL_PIECEPRECEDENTE'+
             ',(GL_QTESTOCK*GL_PUHTNET) AS MTACHAT, GL_QUALIFQTEACH,GL_COEFCONVQTE,GL_COEFCONVQTEVTE,'+
             'GL_QUALIFQTEACH,GL_QUALIFQTEVTE,GL_QUALIFQTESTO,'+
             'GL_MONTANTHTDEV, ' + // --- GUINIER ---
             'BCO_NUMMOUV,BCO_INDICE,BCO_QUANTITE,BCO_QTEVENTE,BCO_TRANSFORME,BCO_TRAITEVENTE '+
             'FROM LIGNE '+
             'LEFT JOIN CONSOMMATIONS ON '+
             '(BCO_NATUREPIECEG=GL_NATUREPIECEG) AND (BCO_SOUCHE=GL_SOUCHE) AND (BCO_NUMERO=GL_NUMERO) '+
             'AND (BCO_NUMORDRE=GL_NUMORDRE) '+
             'WHERE GL_PIECEORIGINE="'+RefOrigine+ '" '+
             'AND GL_NATUREPIECEG IN ('+GetPieceAchat+') '+ {AND GL_VIVANTE="X" AND BCO_TRAITEVENTE<>"X"}
             'AND BCO_TRANSFORME<>"X"';
  QQ := OpenSql (requete,true,-1,'',true);
  TRY
    if not QQ.eof then
    begin
    	TOBR.loadDetailDb ('LIGNE','','',QQ,false);
      For Indice := 0 to TOBR.detail.count -1 do
      begin
      	TOBDR := TOBR.detail[Indice];
        NumMouv := TOBDR.GetValue('BCO_NUMMOUV');
        if NumMouv <> 0 then
        begin
          QteRecepFour := TOBDR.GetValue('BCO_QUANTITE');
//          QTeDejaLivrCli := TOBDR.GetValue('BCO_QTEVENTE');
          if TOBDR.GetValue('BCO_TRAITEVENTE') <> 'X' then GestionConso.MemoriseLienReception (TOBL,TOBDR);
        end;
        Qte := TOBDR.GetValue('QTERECEP');
        FUA := RatioMesure('PIE', TOBDR.GetValue('GL_QUALIFQTEACH'));
        if CoefUaUS <> 0 then
        begin
          if COEFUSUV <> 0 then
          begin
        		Qterecep := QteRecep + (Qte * CoefUaUS * CoefUsUV);
        		MtAchat := MtAchat + (TOBDR.GetValue('MTACHAT') / CoefUAUS / COEFUSUV);
          end else
          begin
        		Qterecep := QteRecep + (Qte * CoefUaUS) / FUA * FUS;
        		MtAchat := MtAchat + (TOBDR.GetValue('MTACHAT') / CoefUAUS * FUA / FUS);
          end;
        end else
        begin
        	Qterecep := QteRecep + (Qte * FUA / FUV);
          MtAchat := MtAchat + (TOBDR.GetValue('MTACHAT') / FUA * FUV);
        end;
      end;
    end;
  FINALLY
    Ferme (QQ);
  END;
  //
  (*
  if (TOBL.getValue('GL_TENUESTOCK') = '-') then ModeRecep := 'X' else ModeRecep := '-';
  // deja livre depuis le stock
  requete := 'SELECT GL_QTEFACT,(GL_QTEFACT*GL_DPA) AS MTACHAT '+
             'FROM LIGNE WHERE GL_PIECEORIGINE="'+RefOrigine+ '" '+
             'AND GL_NATUREPIECEG ="LBT" AND GL_TENUESTOCK="'+Moderecep+'"';
  QQ := OpenSql (requete,true,-1,'',true);
  //
  TOBLS := TOB.Create ('LES LIGNES',nil,-1);
  if not QQ.eof then
  begin
    TOBLS.loadDetailDb ('LIGNE','','',QQ,false);
    For Indice := 0 to TOBLS.detail.count -1 do
    begin
      Qterecep := QteRecep + TOBLS.detail[Indice].getValue('GL_QTEFACT');
      MtAchat := MtAchat + TOBLS.detail[Indice].GetValue('MTACHAT');
    end;
  end;
  ferme (QQ);
  TOBLS.free;
  //
  *)
  if QteRecep > 0 then
  begin
  	PUA := Arrondi(MtAchat / QteRecep,V_PGI.okdecP);
  end;

  QteDejaLivre := TOBL.GetValue('GL_QTESTOCK')-TOBL.GetValue('GL_QTERESTE');
  QteResteAlivrer := TOBL.GetValue('GL_QTERESTE');

  if Ok_ReliquatMt then   // --- GUINIER ---
  begin
    QteDejaLivre    := TOBL.GetValue('GL_MONTANTHTDEV')-TOBL.GetValue('GL_MTRESTE');
    QteResteAlivrer := TOBL.GetValue('GL_MTRESTE');
  end;
  (*
  QteTemp := Abs(QteRecep) - (QteDejaLivre);
  if QTeTemp <= 0 then QteLivrable := 0
  								else QteLivrable := QteRecep - QteDejaLivre;
  *)
//  if QteRecep > QteResteAlivrer then QteLivrable := QteResteAlivrer
    QteLivrable := QteRecep;
//  if QTeLivrable < 0 then QTeLivrable := 0;
  // Controle sur le stock
  if (TOBL.getValue('GL_TENUESTOCK') = 'X') and (PriseSurStock) then
  begin
    if (QteResteAlivrer > QteLivrable) then QteLivrable :=QteResteALivrer;

    if FromDispoGlob then
    begin
      TOBDEPOT := TOBDispoLoc.findFirst(['GQ_ARTICLE','GQ_DEPOT'],[TOBL.GetValue('GL_ARTICLE'),TOBL.GetValue('GL_DEPOT')],true);
    end else if FromDispoArt then
    begin
      TOBDEPOT := TOBA.findFirst(['GQ_ARTICLE','GQ_DEPOT'],[TOBL.GetValue('GL_ARTICLE'),TOBL.GetValue('GL_DEPOT')],true);
    end;

    if TOBDepot = nil then
    begin
    	if FromDispoGlob then
      begin
      	TOBDEpot := TOB.create ('DISPO',TOBDispoLoc,-1);
      end else if FromDispoArt then
      begin
      	TOBDEpot := TOB.create ('DISPO',TOBART,-1);
      end else
      begin
				TOBDEpot := TOB.create ('DISPO',nil,-1);
      end;
      Requete := 'SELECT * FROM DISPO WHERE GQ_ARTICLE="'+TOBL.GetValue('GL_ARTICLE')+'" AND '+
                  'GQ_CLOTURE="-" AND GQ_DEPOT="'+ TOBL.GetValue('GL_DEPOT')+'"';
      QQ := OpenSql (Requete,true,-1,'',true);
      TOBDepot.SelectDB ('',QQ);
      AddChampsSupDispoAff (TOBDEPOT);
      ferme (QQ);
    end;
    if not TOBDepot.fieldExists('_LIVRE') then AddChampsSupDispoAff (TOBDepot);
    QteDispo := TOBDepot.GetValue('GQ_PHYSIQUE') - TOBDepot.GetValue('_LIVRE') -
                TOBDepot.GetValue('GQ_PREPACLI') ;
    if CoefUSUV <> 0 then
    begin
    	QteDispo := (QteDispo * CoefUSUV);// + QteResteALivrer;
    end else
    begin
    	QteDispo := (QteDispo * FUS / FUV);// + QteResteALivrer;
    end;
    if QteDispo < 0 then QteDispo := 0;
    // Passage US - UV
    if (QteLivrable > QteDispo) then QteLivrable := QteDispo;
    PUA := TOBdepot.GetValue('GQ_PMAP') * CoefUSUV;
    if (not FromDispoGlob) and (not FromDispoArt) then
    begin
    	TOBDepot.free;
      TOBDepot := nil;
    end;
  end;
  if (QteLivrable <> 0 ) and (PUA > 0) then
  begin
    // on peut livrer a concurence de QTeLivrable
    if not (TOBL.FieldExists ('DPARECUPFROMRECEP')) then TOBL.AddChampSup('DPARECUPFROMRECEP',false);
    TOBL.PutValue('DPARECUPFROMRECEP',PUA);
    if (TOBdepot <> nil) and ((FromDispoGlob) or (FromDispoArt))  then
    begin
    	TOBDepot.putValue('_LIVRE', TOBDepot.GetValue('_LIVRE')+ (QteLivrable));
    end;
  end;
  Result := QteLivrable;
  TOBR.free;
  if Created then TOBA.free;
end;

Function GetQtelivrableHlien (TOBL : TOB;gestionConso : TGestionPhase; TOBART : TOB = nil) : double;
begin
	result := gestionConso.GetQteLivrable (TOBL,True);
end;

function NaturepieceOKPourOuvrage (TOBPiece : TOB) : boolean;
var Naturepiece,Prefixe : string;
begin
  result := false;

  if TOBPiece.NomTable = 'PIECE' Then Prefixe := 'GP'
  else if TOBPiece.NomTable = 'LIGNE' Then Prefixe := 'GL'
  else exit;

  // MODIF FV Pour facturation des appels (articles en prix posés)
  //if copy(TOBPiece.getValue(prefixe+'_AFFAIRE'),1,1)='W' then exit;
  //
  
  Naturepiece := TOBPiece.GetValue (prefixe+'_NATUREPIECEG');
  if (Naturepiece = 'DBT') or (Naturepiece = 'FBT') or
     (Naturepiece = 'FBP') or (Naturepiece = 'BAC') or
     (Naturepiece = 'AFF') or (Naturepiece = 'FAC') or
     (Naturepiece = 'ETU') or (Naturepiece = 'DAC') or
     (Naturepiece = 'BCE') or (Naturepiece = 'AVC') or
     (Naturepiece = GetParamSoc('SO_BTNATBORDEREAUX')) or
     (Naturepiece = 'ABT') or (Naturepiece = 'FRC') or
     (Naturepiece = 'ABP') or
     (Naturepiece = 'DAP') or (Naturepiece = 'FPR') then Result := true;
     
  // Ajout des pices de commerce de details
  if (Naturepiece = 'ABC') or (Naturepiece = 'FBC') or
     (Naturepiece = 'CC') or (Naturepiece = 'BLC') or
     (Naturepiece = 'DE') then result := true;

end;

function IsPrestationInterne(Article : String) : boolean;
var TheType : string;
    Q : Tquery;
begin
  result := false;
  Q := OpenSql ('SELECT BNP_TYPERESSOURCE FROM NATUREPREST WHERE BNP_NATUREPRES = (SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE="'+Article+'")',true,-1,'',true);
  TRY
    if not Q.eof then
    begin
      TheType := trim(Q.FindField('BNP_TYPERESSOURCE').AsString);
      if (TheType = 'SAL') or (TheType = 'MAT') then result := true;
    end;
  FINALLY
    ferme (Q);
  END;
end;

function TiersModifiableInDocument(TOBPiece : TOB) : Boolean;
var NaturePiece : string;
begin
  result := false;
  NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
  if Pos (NaturePiece,'DBT;ETU;CF;DE;CC') = 0 then exit;
  { Ancienne gestion


  if TOBPiece.GetValue('GP_NATUREPIECEG') <> 'DBT' then exit;

  if TOBPiece.GetValue('GP_NATUREPIECEG') <> 'DAP' then exit;
  }

//if TOBPiece.GetValue('GP_AFFAIRE') <> '' then exit;
  if (NaturePiece = 'DBT') or (NaturePiece = 'DE')   then
  begin
    result := ((GetEtatAffaire(TOBPiece.getValue('GP_AFFAIREDEVIS')) = 'ENC') or
              (TOBPiece.getValue('ISDEJAFACT') = '-'));
  end else if NaturePiece = 'ETU' then
  begin
    result := (GetEtatAffaire(TOBPiece.getValue('GP_AFFAIREDEVIS')) = 'ENC');
  end else if (NaturePiece = 'CF') or (NaturePiece = 'CC') then
  begin
    result := (TOBPiece.getValue('GP_DEVENIRPIECE') = '');
  end;
end;

function IsRetourFournisseur (TObpiece : TOB) : boolean;
var Prefixe : string;
begin
	result := false;
  if TOBPiece.NomTable = 'PIECE' Then
  begin
  	Prefixe := 'GP';
  end else if TOBPiece.NomTable = 'LIGNE' Then
  begin
  	Prefixe := 'GL';
  end else Exit;
  if TOBPIece.getValue(Prefixe+'_NATUREPIECEG')='BFA' Then Result := true;
end;

function IsRetourClient(TOBPiece : TOB) : Boolean;
var Prefixe : string;
begin
	result := false;
  if TOBPiece.NomTable = 'PIECE' Then
  begin
  	Prefixe := 'GP';
  end else if TOBPiece.NomTable = 'LIGNE' Then
  begin
  	Prefixe := 'GL';
  end else Exit;
  if TOBPIece.getValue(Prefixe+'_NATUREPIECEG')='BFC' Then Result := true;
end;

function ControleDateDocument ( StD : String ; Naturepiece : string) : integer ;
Var DD  : TDateTime ;
		JNL : string;
BEGIN
Result:=0 ;
if Not IsValidDate(StD) then BEGIN Result:=1 ; Exit ; END ;
{$IFDEF CONSOCERIC}
Exit ;
{$ENDIF}
DD:=StrToDate(StD) ;
(* Ajout LS *)
if GetInfoParPiece(Naturepiece,'GPP_TYPEPASSCPTA') = 'AUC' then Exit;
JNL := GetInfoParPiece(Naturepiece,'GPP_JOURNALCPTA');
(* --------- *)
if DD<VH^.Encours.Deb then BEGIN Result:=2 ; Exit ; END ;
if ((VH^.Suivant.Fin>0) and (DD>VH^.Suivant.Fin)) then BEGIN Result:=2 ; Exit ; END ;
if ((VH^.Suivant.Fin=0) and (DD>VH^.Encours.Fin)) then BEGIN Result:=2 ; Exit ; END ;
if ((VH^.DateCloturePer>0) and (DD<=VH^.DateCloturePer)) then BEGIN Result:=3 ; Exit ; END ;
if ISmoisCloture(StrToDate(StD)) then begin Result := 3; Exit; END;
if JNL <> '' then
begin
	if isMoisJournalCloture (JNL, DD) then begin Result := 45; Exit; End;
end;
{$IFNDEF GCGC}
if ((VH^.NbjEcrAv>0) and (DD<V_PGI.DateEntree) and (V_PGI.DateEntree-DD>VH^.NbjEcrAv)) then BEGIN Result:=5 ; Exit ; END ;
if ((VH^.NbjEcrAp>0) and (DD>V_PGI.DateEntree) and (DD-V_PGI.DateEntree>VH^.NbjEcrAp)) then BEGIN Result:=5 ; Exit ; END ;
{$ENDIF}
END ;

function GetCodeBQ (TOBBanqueCP : TOB; CodeAuxiliaire: String; NumeroRIB : Integer) : boolean;
var Req : string;
		QQ : TQuery;
begin
	result := false;
  Req := 'SELECT R_DOMICILIATION,R_ETABBQ,R_GUICHET,R_NUMEROCOMPTE,R_CLERIB FROM RIB WHERE ' +
         'R_AUXILIAIRE="' + CodeAuxiliaire + '" AND R_NUMERORIB='+ IntToStr(NumeroRIB);
  QQ := OpenSql(Req,True,-1,'',true);
  if not QQ.eof then
  begin
  	TOBBanqueCP.SelectDB ('',QQ,true);
    result := true;
  end;
  ferme (QQ);
end;

function GetCodeBQ (Auxiliaire : String; NumeroRIB : Integer; var RefBancaire : string) : boolean;
var Req : string;
		QQ : TQuery;
begin
	result := false;
  Req := 'SELECT R_DOMICILIATION FROM RIB WHERE ' +
         'R_AUXILIAIRE="'+ Auxiliaire + '"  AND ' +
         'R_NUMERORIB='  + IntToStr(NumeroRIB);
  QQ := OpenSql(Req,True,1,'',true);
  if not QQ.eof then
  begin
    RefBancaire := (QQ.FindField('R_DOMICILIATION').AsString);
    result := true;
  end;
  ferme (QQ);

end;

function GetContact (TOBContact : TOB;LeTiers : string;NumeroContact : integer) : boolean;
var Req : string;
		QQ : TQuery;
begin
	result := false;

  if LeTiers = '' then exit;

  Req := 'SELECT C_CIVILITE,C_NOM,C_PRENOM,C_TELEPHONE,C_RVA FROM CONTACT WHERE C_TYPECONTACT="T" AND '+
         'C_AUXILIAIRE="'+TiersAuxiliaire(LeTiers)+'" AND C_NUMEROCONTACT='+inttoStr(NumeroContact);
  QQ := OpenSql(Req,True,-1,'',true);
  if not QQ.eof then
  begin
  	TOBCOntact.SelectDB ('',QQ,true);
    result := true;
  end;
  ferme (QQ);
end;

function GetContact (LeTiers : string;NumeroContact : integer; var libcontact : string) : boolean;
var Req : string;
		QQ : TQuery;
begin
	result := false;

  if LeTiers = '' then exit;

  Req := 'SELECT C_NOM FROM CONTACT WHERE C_TYPECONTACT="T" AND '+
         'C_AUXILIAIRE="'+TiersAuxiliaire(LeTiers)+'" AND C_NUMEROCONTACT='+inttoStr(NumeroContact);
  QQ := OpenSql(Req,True,1,'',true);
  if not QQ.eof then
  begin
    libcontact := (QQ.FindField('C_NOM').AsString);
    result := true;
  end;
  ferme (QQ);
end;

procedure TGenereLivraison.AddLivraisonDirecte(cledoc: r_cledoc; QteRecep : double);
var TOBLD,TOBLDD : TOB;
begin
	TOBLD := TOBLivrDirecte.findFirst (['NATURE','NUMERO','INDICE'],[Cledoc.NaturePiece,Cledoc.NumeroPiece,Cledoc.Indice],true);
  if TOBLD = nil then
  begin
  	TOBLD := TOB.create ('__LIVREE ',TOBLivrDirecte,-1);
    TOBLD.AddChampSupValeur ('NATURE',Cledoc.NaturePiece);
    TOBLD.AddChampSupValeur ('NUMERO',Cledoc.NumeroPiece);
    TOBLD.AddChampSupValeur ('INDICE',Cledoc.Indice);
  end;
  TOBLDD := TOB.create ('__DETAIL ',TOBLD,-1);
  TOBLDD.AddChampSupValeur ('NUMORDRE',Cledoc.Numordre);
  TOBLDD.AddChampSupValeur ('QTERECEP',QteRecep);
end;

procedure MajQtesAvantSaisie (TobPiece : TOB);
Var ind : integer;
		QteSitRest : double;
    TOBL : TOB;
    Cledoc : R_CLEDOC;
    LastPrec,CurrPiece : string;
begin
   for ind := 0 to TOBpiece.detail.count -1 do
   begin
     if TOBPiece.detail[Ind].getValue('GL_TYPELIGNE')= 'ART' then
     begin
     	 TOBPiece.detail[Ind].putValue('OLD_QTESIT',TOBPiece.detail[Ind].GetValue('GL_QTESIT'));
     	 QteSitRest := Abs(TOBPiece.detail[Ind].GetValue('GL_QTESIT')) - Abs(TOBPiece.detail[Ind].GetValue('GL_QTEFACT'));
       if QteSitRest > 0 then
       begin
         TOBPiece.detail[Ind].PutValue('GL_QTESIT',TOBPiece.detail[Ind].GetValue('GL_QTESIT') -
                                                   TOBPiece.detail[Ind].GetValue('GL_QTEFACT'));
       end else
       begin
       	 TOBPiece.detail[Ind].PutValue('GL_QTESIT',0);
       end;
     end else
     begin
       TOBPiece.detail[Ind].PutValue('GL_QTESIT',0);
       TOBPiece.detail[Ind].PutValue('GL_QTEPREVAVANC',0);
       TOBPiece.detail[Ind].PutValue('GL_POURCENTAVANC',0);
     end;
   end;
end;

procedure DefiniDejaFacture (TOBPiece,TOBOuvrages : TOB;DEV : rdevise; TypeFacturation : string);
Var ind : integer;
		QteSitRest : double;
    TOBL : TOB;
    lastprec,CurrPiece : string;
    Cledoc : R_CLEDOC;
begin
  for ind := 0 to TOBpiece.detail.count -1 do
  begin
    if TOBPiece.detail[Ind].getValue('GL_TYPELIGNE')= 'ART' then
    begin
      TOBL := TOBPiece.detail[Ind];
      // Ancienne gestion (pour rester compatible)
      TOBL.putValue('OLD_QTESIT',TOBL.GetValue('GL_QTESIT'));
(*    MODIF BRL 2/08/2011 FQ 15631
      QteSitRest := Abs(TOBL.GetValue('GL_QTESIT')) - abs(TOBL.GetValue('GL_QTEFACT'));
      if QteSitRest > 0 then
      begin
        TOBL.putValue('DEJAFACT',TOBL.GetValue('GL_QTESIT')-TOBL.GetValue('GL_QTEFACT'));
      end else
      begin
        TOBL.putValue('DEJAFACT',0);
      end;
*)
      TOBL.putValue('DEJAFACT',TOBL.GetValue('GL_QTESIT')-TOBL.GetValue('GL_QTEFACT'));
      // -------
      if TOBL.getValue('BLF_NATUREPIECEG')=''  then
      begin
        if TOBL.GetValue('GL_QTEPREVAVANC')<>0 then
        begin
          // gestion des anciennes facturations
          PositionneAncienneFacturation(TOBL,DEV,TypeFacturation);
          if IsOuvrage(TOBL) then
          begin
            PositionneAncienneFacDetOuv(TOBL,TOBOuvrages,DEV,Typefacturation);
          end;
        end else
        begin
          InitMontantfacturation(TOBL);
        end;
      end;
    end;
 end;
end;

function TGenereLivraison.DejaLivre(Piece: string): boolean;
var Cledoc : R_CLEDOC;
		REQUETE : String;
    QQ : TQuery;
begin
	result := false;
  if Piece = '' then exit;
	DecodeRefPiece (Piece,cledoc);
  Requete := 'SELECT BCO_LIENVENTE FROM CONSOMMATIONS WHERE BCO_NUMMOUV = (SELECT BLP_NUMMOUV FROM LIGNEPHASES WHERE'+
  					 ' BLP_NATUREPIECEG="'+Cledoc.NaturePiece+'"'+
             ' AND BLP_SOUCHE="'+Cledoc.Souche+'" AND BLP_NUMERO='+inttoStr(Cledoc.NumeroPiece)+
             ' AND BLP_NUMORDRE='+IntToStr(cledoc.NumOrdre)+') AND BCO_LIENVENTE <> 0';
  QQ := OpenSql (Requete,True,-1,'',true);
  result := not QQ.eof;
  ferme (QQ);
end;

function IsContreEtude (TOBpiece : TOB) : boolean;
begin
	result := false;
	if TOBPIece.NomTable = 'PIECE' then result := (TOBPiece.getValue('GP_NATUREPIECEG')='BCE')
  else if TOBPIece.NomTable = 'LIGNE' then result := (TOBPiece.getValue('GL_NATUREPIECEG')='BCE')
  else if TOBPIece.NomTable = 'LIGNEOUV' then result := (TOBPiece.getValue('BLO_NATUREPIECEG')='BCE');
end;

function IsPrevisionChantier (TOBpiece : TOB) : boolean;
begin
	result := false;
	if TOBPIece.NomTable = 'PIECE' then result := (TOBPiece.getValue('GP_NATUREPIECEG')='PBT')
  else if TOBPIece.NomTable = 'LIGNE' then result := (TOBPiece.getValue('GL_NATUREPIECEG')='PBT')
  else if TOBPIece.NomTable = 'LIGNEOUV' then result := (TOBPiece.getValue('BLO_NATUREPIECEG')='PBT');
end;

function GetNumCompteur (Souche : string; ODate : TDateTime; var TheNumero : integer) : boolean;
var TheNumPrec : integer;
		QQ : TQuery;
    requete : string;
    CptArret : integer;
    MajOk : boolean;
begin
	result := true;
  CptArret := 1;
  MajOk := false;
	QQ := OpenSQL('SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="BTP" AND SH_SOUCHE="'+Souche+'"', True,-1,'',true);
  if not QQ.eof then
  begin
  	TheNumprec := QQ.findfield('SH_NUMDEPART').AsInteger;
  	ferme (QQ);
    repeat
      TheNumero := TheNumPrec + 1;
      requete := 'UPDATE SOUCHE SET SH_NUMDEPART ='+IntToStr(TheNumero)+' WHERE SH_TYPE="BTP" AND SH_SOUCHE="'+Souche+'" AND'+
                 ' SH_NUMDEPART='+IntToStr(TheNumprec);
      if ExecuteSQL (requete) <= 0 then inc(CptArret) else MajOk := true;
    until (CptArret > 10) or (MajOk);
    result := MajOk;
  end else
  begin
    ferme (QQ);
  	PgiBox ('le compteur '+Souche+' n''existe pas..il faut la creer');
  end
end;

function isPieceGerableFraisDetail(naturePiece : string) : boolean;
begin
  result := (naturePiece = VH_GC.AFNatProposition) or
         		(naturePiece = VH_GC.AFNatAffaire) or
         		(naturePiece = 'ETU') or
         		(naturePiece = 'PBT') or
            (naturePiece = 'BCE') ;
end;

procedure RestitueAcompte(Cledoc : R_cledoc; Journal : string; Numecr : integer ;MontantDev,Montant : double);
var QQ : TQuery;
		TOBA : TOB;
begin
	TOBA := TOB.Create('ACOMPTES',nil,-1);
	QQ := OpenSql ('SELECT * FROM ACOMPTES WHERE GAC_NATUREPIECEG="'+Cledoc.NaturePiece+'" AND'+
  							 ' GAC_SOUCHE="'+Cledoc.Souche+'" AND GAC_NUMERO='+InttoStr(cledoc.NumeroPiece)+
                 ' AND GAC_INDICEG='+IntToStr(cledoc.Indice)+' AND GAC_JALECR="'+JOurnal+'"'+
                 ' AND GAC_NUMECR='+InttoStr(Numecr),true,-1,'',true);
  if not QQ.eof then
  begin
  	TOBA.SelectDB ('',QQ);
    TOBA.PutValue('GAC_MONTANTDEV',TOBA.GetValue('GAC_MONTANTDEV')+MontantDEv);
    TOBA.PutValue('GAC_MONTANT',TOBA.GetValue('GAC_MONTANT')+Montant);
    TOBA.UpdateDB (false);
  end;
  TOBA.Free;
end;

procedure DefiniPvLIgneModifiable (TOBpiece,TOBPOrcs: TOB);
var indice : integer;
begin
	if (TOBPiece.GetValue('GP_NATUREPIECEG')='DBT') and
//  	 (RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'))='AVA') and
     (IsExisteFraisChantier (TOBPiece) or (GetTauxFg(Tobporcs)<>0)) and
		 (TOBPiece.GetValue('ETATDOC')='ACP') then
  begin
  	for Indice := 0 to TOBPiece.detail.count -1 do
    begin
    	TOBPiece.Detail[Indice].putValue('GL_BLOQUETARIF','X');
    end;
  end;
end;


procedure RecalculeCoefFrais (TOBPiece,TOBporcs : TOB);
var ExistFG : boolean;
		TotalFrais,TauxFG,TauxFC : double;
    i : integer;
begin

  TotalFrais := GetMontantFraisDetail (TOBPiece, ExistFg);
  if (ExistFg) and (TotalFrais = 0) then
  begin
    BTPSupprimePieceFrais (TOBpiece);
    ExistFg := false;
  end;

  if (TOBPiece.getValue('GP_MONTANTPAFC') <> 0) and (TotalFrais <> 0) then
  begin
    //FV1 - 19/06/2017 : FS#2610 - TEAM - Le calcul du coefficient de frais de chantier doit se faire avant les frais généreaux
    if GetParamSocSecur('SO_BTFGINFC', False) then
      TauxFC := arrondi(TotalFrais / (TOBPiece.getValue('GP_MONTANTPAFC')),9)
    else
  	TauxFC := arrondi(TotalFrais / (TOBPiece.getValue('GP_MONTANTPAFC')+TOBPiece.getValue('GP_MONTANTFG')),9);
  end
  else
  begin
  	TauxFC := 0;
  end;
  TOBPIece.putValue('GP_COEFFC',TauxFC);


// Calcul du pourcentage total de frais répartis
// => Dans TOBPorcs cumul des pourcentage de frais ayant le code FRAISREPARTIS à TRUE
//    puis calcul du coefficient global que réprésentent ces frais sur le total HT du document
  TOBPIece.putValue('GP_COEFFR',Arrondi(getTauxFG(TOBporcs)/100,6));

end;


function GetCoefFC (TOBpiece : TOB) : double;
var TotalFrais : double;
		ExistFg : boolean;
begin
  TotalFrais := GetMontantFraisDetail (TOBPiece, ExistFg);
  if (TOBPiece.getValue('GP_MONTANTPAFC') <> 0) and (TotalFrais <> 0) then
  begin
  	result := arrondi(TotalFrais / (TOBPiece.getValue('GP_MONTANTPAFC')+TOBPiece.getValue('GP_MONTANTFG')),9);
  end else
  begin
  	result := 0;
  end;
end;

function GetCoefFC (MontantAppliquable,MontantFraisChantier : double) : double;
begin
  if (MontantAppliquable <> 0 ) and (MontantFraisChantier <> 0) then
  begin
  	result := arrondi(MontantFraisChantier / MontantAppliquable,9);
  end else
  begin
  	result := 0;
  end;
end;

Function GetContrat (Affaire : string) : string;
var QQ : TQUery;
begin
  result := '';
  if Affaire = '' then exit;
  QQ := OpenSql ('SELECT AFF_AFFAIREINIT FROM AFFAIRE WHERE AFF_AFFAIRE="'+Affaire+'"',True,-1,'',true);
  if not QQ.eof then
  begin
    Result := QQ.findField('AFF_AFFAIREINIT').AsString;
  end;
  ferme (QQ);
end;

procedure  SetLigneFacture (TOBL,TOBD : TOB);
begin
  TOBD.putValue('BLF_NATUREPIECEG', TOBL.GetValue('GL_NATUREPIECEG'));
  TOBD.putValue('BLF_SOUCHE',       TOBL.GetValue('GL_SOUCHE'));
  TOBD.putValue('BLF_DATEPIECE',    TOBL.GetValue('GL_DATEPIECE'));
  TOBD.putValue('BLF_NUMERO',       TOBL.GetValue('GL_NUMERO'));
  TOBD.putValue('BLF_INDICEG',      TOBL.GetValue('GL_INDICEG'));
  TOBD.putValue('BLF_NUMORDRE',     TOBL.GetValue('GL_NUMORDRE'));
  TOBD.putValue('BLF_MTMARCHE',     TOBL.GetValue('BLF_MTMARCHE'));
  TOBD.putValue('BLF_MTDEJAFACT',   TOBL.GetValue('BLF_MTDEJAFACT'));
  TOBD.putValue('BLF_MTCUMULEFACT', TOBL.GetValue('BLF_MTCUMULEFACT'));
  TOBD.putValue('BLF_MTSITUATION',  TOBL.GetValue('BLF_MTSITUATION'));
  TOBD.putValue('BLF_QTEMARCHE',    TOBL.GetValue('BLF_QTEMARCHE'));
  TOBD.putValue('BLF_QTEDEJAFACT',  TOBL.GetValue('BLF_QTEDEJAFACT'));
  TOBD.putValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTECUMULEFACT'));
  TOBD.putValue('BLF_QTESITUATION', TOBL.GetValue('BLF_QTESITUATION'));
  TOBD.putValue('BLF_POURCENTAVANC',TOBL.GetValue('BLF_POURCENTAVANC'));
  TOBD.putValue('BLF_NATURETRAVAIL',TOBL.Getvalue('GLC_NATURETRAVAIL'));
  TOBD.putValue('BLF_FOURNISSEUR',  TOBL.Getvalue('GL_FOURNISSEUR'));
  TOBD.putValue('BLF_TOTALTTCDEV',  TOBL.Getvalue('BLF_TOTALTTCDEV'));
  TOBD.putValue('BLF_MTPRODUCTION', TOBL.Getvalue('BLF_MTPRODUCTION'));
  TOBD.SetAllModifie (true);
end;

procedure AddChampsSupTable (TOBCur : TOB;PrefixeTable : string);
var ITable,Indice,IndiceTOB : integer;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  iTable := PrefixeToNum(PrefixeTable);  if iTable = 0 then exit;
  //
  Table := Mcd.getTable(Mcd.PrefixeToTable(PrefixeTable));
  FieldList := Table.Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
    if not TOBCur.FieldExists ((FieldList.Current as IFieldCOM).name) then
    begin
      TOBCur.AddChampSup ((FieldList.Current as IFieldCOM).name,false);
      IndiceTOB := TOBCur.GetNumChamp ((FieldList.Current as IFieldCOM).name);
      InitChampsSup (TOBCur,(FieldList.Current as IFieldCOM).name,(FieldList.Current as IFieldCOM).tipe);
    end;
  end;
end;

procedure InitChampsSup(TOBL : TOB; Champs,tipe : string);
var typeC : string;
    IndiceChamps : integer;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  if (tipe = 'INTEGER') or (tipe = 'SMALLINT') then TOBL.putValue(Champs,0)
  else if (tipe = 'DOUBLE') or (tipe = 'EXTENDED') or (typeC = 'DOUBLE') then TOBL.putValue(Champs,0)
  else if (tipe = 'DATE') then TOBL.putValue(Champs,iDate1900)
  else if (tipe = 'BLOB') or (typeC = 'DATA') then TOBL.putValue(Champs,'')
  else if (tipe = 'BOOLEAN') then TOBL.putValue(Champs,'-')
  else TOBL.putValue(Champs,'')
end;

function getInfoPuissance (NbLig : integer) : integer;
var termine : boolean;
    Iloc : integer;
begin
  result := 1;
  termine := false;
  Iloc := 0;
  if NbLig = 0 then exit;
  repeat
    if nbLig < Power(10, Iloc) then
    begin
      termine := true;
    end else
    begin
      Inc(Iloc);
    end;
  until termine;
  result := Iloc;
end;


procedure SetItemHvalComboBox (Combo : THValComboBox; Valeur : string);
var Indice : integer;
begin
	for Indice := 0 to Combo.Values.Count -1 do
  begin
  	if Combo.Values[Indice]=Valeur then
    begin
    	Combo.ItemIndex := Indice;
      break;
    end;
  end;
end;

//FV1 : function de recherche des informations bancaire à partir de la table BANQUECP et envoie dans TOB
Procedure LectBanque(TOBBQE : TOB; NumeroRIB : String);
var StSQL : string;
    QQ    : TQuery;
begin

  //Chargement suite info bancaire...
  StSQL := 'SELECT * FROM BANQUECP WHERE BQ_GENERAL="' + NumeroRib + '"';

	QQ := OpenSql (StSQL, true, -1, '', true);
  if not QQ.eof then TOBBQE.SelectDB ('',QQ);

end;

Procedure LectRIB(TOBBQE : TOB; Auxiliaire : string; NumeroRIB : Integer);
var StSQL : string;
    QQ    : TQuery;
begin

  //Chargement suite info bancaire...
  StSQL := 'SELECT * FROM RIB WHERE R_AUXILIAIRE="' +  Auxiliaire + '" AND R_NUMERORIB=' + IntToStr(NumeroRib);

	QQ := OpenSql (StSQL, true, -1, '', true);
  if not QQ.eof then TOBBQE.SelectDB ('',QQ);

end;

procedure TGenereLivraison.TransformeAchatEnVente(TOBL: TOB);

	function GetUniteVte (Article : string) : string;
  var QQ : TQuery;
  begin
  	result := '';
    QQ := openSql ('SELECT GA_QUALIFUNITEVTE FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"',true,1,'',true);
    if not QQ.eof then
    begin
    	result := QQ.findField('GA_QUALIFUNITEVTE').AsString;
    end;
    Ferme (QQ);
  end;
	function GetUniteSTO (Article : string) : string;
  var QQ : TQuery;
  begin
  	result := '';
    QQ := openSql ('SELECT GA_QUALIFUNITESTO FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"',true,1,'',true);
    if not QQ.eof then
    begin
    	result := QQ.findField('GA_QUALIFUNITESTO').AsString;
    end;
    Ferme (QQ);
  end;

var FUA,FUV,FUS : double;
begin
	if TOBL.getvalue('GL_COEFCONVQTE') <> 0 then
  begin
    FUV := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
    FUS := RatioMesure('PIE', GetUniteSTO(TOBL.GetValue('GL_ARTICLE'))); if FUS = 0 then FUS := 1;
//    TOBL.PutValue('GL_PUHTNET', Arrondi(TOBL.GetValue('GL_PUHTNET') / TOBL.GetValue('GL_COEFCONVQTE')/FUS*FUV,V_PGI.okdecP));
//    TOBL.PutValue('GL_QUALIFQTEACH', TOBL.GetValue('GL_QUALIFQTEVTE'));
    TOBL.PutValue('GL_QUALIFQTEVTE', GetUniteVte(TOBL.GetValue('GL_ARTICLE')));
  end else
  begin
//    TOBL.PutValue('GL_QUALIFQTEACH', TOBL.GetValue('GL_QUALIFQTEVTE'));
    TOBL.PutValue('GL_QUALIFQTEVTE', GetUniteVte(TOBL.GetValue('GL_ARTICLE')));
    FUV := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEVTE')); if FUV = 0 then FUV := 1;
    FUA := RatioMesure('PIE', TobL.GetValue('GL_QUALIFQTEACH')); if FUA = 0 then FUA := 1;
    //
//    TOBL.PutValue('GL_PUHTNET', Arrondi(TOBL.GetValue('GL_PUHTNET') * FUV / FUA,V_PGI.okdecP));
  end;
end;

function BTGetSectionStock(Depot : string) : string;
var QQ : TQuery;
begin
  Result := '';
  if (Depot <> '') then
  begin
		QQ := OpenSQL('SELECT GDE_SECTION FROM DEPOTS WHERE GDE_DEPOT="'+Depot+'"',True,1,'',true);
    if not QQ.Eof then Result := QQ.findField('GDE_SECTION').AsString;
    ferme (QQ);
  end;
  if (Result = '') then
  begin
    Result := GetParamSocSecur('SO_BTCPTANALSTOCK','') ;
  end;
end;

function IsParamGeneDepot : Boolean;
begin
	Result := (GetParamSocSecur('SO_BTCPTANALSTOCK','')<>'') ;
  if not Result then
  begin
		Result := ExisteSQL('SELECT GDE_SECTION FROM DEPOTS WHERE GDE_SECTION <> ""');
  end;
end;

function GetShare (NomTable : string) : string;
var Q : TQuery;
begin
  result := '';
  Q := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="'+NomTable+'"', True);
  if not Q.EOF then
  begin
  	result := Q.FindField('DS_NOMBASE').AsString+'.dbo.';
  end;
	Ferme (Q);
end;

//FV : 14/05/2012 - Contrôle par date de saisie et date parametre saisie de la pièce (Revisé le 17/12/2014)
Function CtrlDateLimiteSaisie(DateSaisie : TDateTime) : Boolean;
var DateLimiteAvant : TDateTime;
    DateLimiteApres : TDateTime;
    DateUserAvant   : TDateTime;
    DateUserApres   : TDateTime;
    UserID          : String;
begin

  Result := False;

  DateLimiteAvant := GetParamSocSecur('SO_LIMITEAVANT',iDate1900);
  DateLimiteApres := GetParamSocSecur('SO_LIMITEAPRES',iDate2099);
  //
  DateUserAvant   := GetParamSocSecur('SO_USERAVANT', iDate1900);
  DateUserApres   := GetParamSocSecur('SO_USERAPRES', iDate2099);
  //
  UserID          := GetParamsocsecur('SO_USERSAISIE',  'LSE');

  //recherche de l'utilisateur dans la liste des users autorisé
  if pos(V_PGI.User, UserID) > 0 then
  begin
    if (DateUserAvant > DateSaisie) OR (DateUserApres < DateSaisie) then
    begin
      exit;
    end;
  end
  else
  begin
    if (DateLimiteAvant > DateSaisie) OR (DateLimiteApres < DateSaisie) then
    begin
      exit;
    end;
  end;

  Result := True;

end;

function CtrlOkReliquat(TOBL : TOB; Prefixe : String) : Boolean;
Var StSQL : string;
    QQ    : TQuery;
begin

  Result := False;

  if TOBL = nil then Exit;

  if TOBL.GetString(Prefixe + '_TYPELIGNE') <> 'ART' then Exit;

  if TOBL.GetString(Prefixe + '_TYPEARTICLE') <> 'PRE' then Exit;
  if TOBL.FieldExists('GESTRELIQUAT') then
  begin
    Result := TOBL.GetBoolean ('GESTRELIQUAT');
  end else
  begin
  StSQL := 'SELECT GA_RELIQUATMT FROM ARTICLE WHERE GA_ARTICLE="' + TOBL.GetString(Prefixe + '_ARTICLE') + '" ';
  StSQL := StSql + ' AND GA_TYPEARTICLE="' + TOBL.GetString(Prefixe + '_TYPEARTICLE') + '" ';

  QQ := OpenSQL(StSQL, False, -1,'',true);
  if Not QQ.eof then
    if QQ.FindField('GA_RELIQUATMT').AsString = 'X' then Result := True else Result := False;

  ferme(QQ);
  end;
end;

Function FormatMultiValComboforSQL(Champ : String) : String;
Var Critere : String;
begin

  Critere := Uppercase(Trim(ReadTokenSt(Champ)));

  while Critere <> '' do
  begin
     Result  := Result + '"' + Critere + '",';
     Critere:= uppercase(Trim(ReadTokenSt(Champ)));
  end;

  Result := Copy(Result, 1, Length(Result)-1);

end;

{ Publipost }

class function Publipost.GetPrefixFromTableName(TableName : string) : string;
begin
  case Tools.CaseFromString(TableName, ['AFFAIRE', 'TIERS', 'PIECEADRESSE', 'RESSOURCE', 'UTILISAT', 'SUSPECTS', 'CONTACT']) of
    0 : Result := 'Contrat';     {AFFAIRE}
    1 : Result := 'CliFac';      {TIERS}
    2 : Result := 'AdrInt';      {PIECEADRESSE}
    3 : Result := 'Ressource';   {RESSOURCE}
    4 : Result := 'Utilisateur'; {UTILISAT}
    5 : Result := 'Suspect';     {SUSPECTS}
    6 : Result := 'Contact';     {CONTACT}
  else
    Result := '';
  end;
end;

class function Publipost.GetTranslateData(FieldName, Code : string) : string;
var
  Sql     : string;
  Suffixe : string;
  Qry     : TQuery;
begin
  Suffixe := ExtractSuffixe(FieldName);
  Sql     := 'SELECT DO_COMBO FROM DECOMBOS WHERE DO_NOMCHAMP LIKE "%' +Suffixe + '%"';
  Qry     := OpenSQL(Sql,True,1,'',true);
  try
    Result := iif(not Qry.Eof, Qry.Fields[0].AsString, '');
  finally
    ferme (Qry);
  end;
  if Result = '' then
    Result := Code
  else
    Result := RechDom(Result, Code, False);
end;

class procedure Publipost.NewModel(InputFilePath : string; TobFields : TOB; FieldsList : string);
begin
  if (InputFilePath <> '') and (Assigned(TobFields)) and (FieldsList <> '') then
    LancePublipostage('NEW', InputFilePath, '', TobFields, FieldsList, nil, False);
end;

class procedure Publipost.UpdateModel(InputFilePath: string; TobFields: TOB; FieldsList: string);
begin
  if (InputFilePath <> '') and (Assigned(TobFields)) and (FieldsList <> '') then
    LancePublipostage('OPEN', InputFilePath, '', TobFields, FieldsList, nil, False);
end;

class function Publipost.DeleteModel(InputFilePath, Title : string): Boolean;
begin
  Result := False;
  if not FileExists(InputFilePath) then
  	PGIError(Format('Le modèle %s n''existe pas', [InputFilePath]))
  else
  if PGIAsk(Format('Désirez-vous réellement supprimer le modèle %s ?', [InputFilePath]), Title) = mrYes then
    Result := (DeleteFile(PAnsiChar(InputFilePath)));
end;

class procedure Publipost.SetFieldsList(Prefix, TableName : string; var FieldsList, SqlFieldsList : string);
var
  Qry : TQuery;
	Sql : string;
begin
  SqlFieldsList := '';
  FieldsList    := '';
  Sql := 'SELECT DH_NOMCHAMP,DH_LIBELLE'
       + ' FROM DECHAMPS'
       + ' WHERE DH_PREFIXE IN (SELECT DT_PREFIXE FROM DETABLES WHERE DT_NOMTABLE IN ("' + TableName + '")) AND DH_CONTROLE LIKE "%W%"'
       ;
	Qry := OpenSql(Sql, True, 0, '', True);
  try
    if Not Qry.Eof then
    begin
      Qry.first;
      While not Qry.eof do
      begin
        SqlFieldsList := SqlFieldsList + ',' + Qry.Fields[0].AsString;
        FieldsList    := FieldsList    + ';' + Prefix + '_' + Qry.Fields[1].AsString;
        Qry.Next;
      end;
      SqlFieldsList := Copy(SqlFieldsList, 2, Length(SqlFieldsList));
      FieldsList    := Copy(FieldsList, 2, Length(FieldsList));
    end;
  finally
    Ferme (Qry);
  end;
end;

class procedure Publipost.SetData(TobData: TOB; FF: TField);
var
  FieldName : string;
  II        : Integer;
begin
  FieldName := FF.FullName;
  if  Copy(FieldName,1,7 )= 'T_TABLE' then // cas Particulier des tables libres tiers
  begin
    II := StrToInt(Copy(FieldName, 8, 1)) + 1;
    TobData.AddChampSupvaleur(FieldName, RechDom('GCZONELIBRE' + InttoStr(II), FF.AsString, False));
  end else if  Copy(FieldName, 1, 12) = 'AFF_LIBREAFF' then // cas Particulier des tables libres affaires
    TobData.AddChampSupvaleur(FieldName, GetTranslateData(FieldName, FF.AsString))
  else
  begin
    case GetFieldType(FieldName) of
      ttfInt     : TobData.AddChampSupvaleur(FieldName, IntToStr(FF.AsInteger));
      ttfNumeric : TobData.AddChampSupvaleur(FieldName, StrF00(FF.AsFloat, 2));
      ttfDate    : TobData.AddChampSupvaleur(FieldName, DateToStr(FF.AsDateTime));
      ttfMemo    : TobData.AddChampSupvaleur(FieldName, FF.AsString);
      ttfCombo   : TobData.AddChampSupvaleur(FieldName, GetTranslateData(FieldName, FF.AsString));
    else
      TobData.AddChampSupvaleur(FieldName, FF.AsString);
    end;
  end;
end;

class procedure Publipost.AddFieldsInTob(TobData : TOB; TableName, SqlFieldsList, Where : string; WithData: Boolean);
var
  Sql  : string;
  Qry  : TQuery;
  Cpt  : Integer;
begin
  if     (assigned(TobData))
     and (TableName <> '')
     and (SqlFieldsList <> '')
     and (Where <> '')
  then
  begin
    Sql := 'SELECT '  + SqlFieldsList
          + ' FROM '  + TableName
          + ' WHERE ' + Where;
    Qry := OpenSQL(Sql, True, 1, '', True);
    try
      for Cpt := 0 to Qry.Fields.Count -1 do
      begin
        if WithData then
          SetData(TobData, Qry.Fields[Cpt])
        else
          TobData.AddChampSupvaleur(Qry.Fields[Cpt].FullName, '');
      end;
    finally
      Ferme(Qry);
    end;
  end;
end;

class procedure Publipost.MergeExecute(TableName, Where, TemplateDoc, OtputDoc : string);
var
  TobFields     : TOB;
  FieldsList    : string;
  FieldsListSql : string;
begin
  if (TableName <> '') and (Where <> '') and (TemplateDoc <> '') and (OtputDoc <> '') then
  begin
    TobFields := TOB.Create ('CONTRAT',nil,-1);
    try
      Publipost.SetFieldsList(Publipost.GetPrefixFromTableName(TableName), TableName, FieldsList, FieldsListSql);
      Publipost.AddFieldsInTob(TobFields, TableName, FieldsListSql, Where, True);
      LancePublipostage('TOB', TemplateDoc, OtputDoc, TobFields, FieldsList, nil, True);
    finally
      FreeAndNil(TobFields);
    end;
  end;
end;

end.

