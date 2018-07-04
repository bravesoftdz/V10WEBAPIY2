unit UtilArticle;

interface
uses {$IFDEF VER150} variants,{$ENDIF}  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox, Hdimension, UTOB, UTOM, AGLInit,
{$IFDEF EAGLCLIENT}
      emul, Maineagl,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids, db,Fe_Main,
{$ENDIF}
      wCommuns, { GPAO1 }
      LookUp,EntGC,ED_TOOLS,UtilPgi,UtofBTChangeCodeArt;


Type TChangeCodeArt = class
     private
     CodeArticle : string;
     NewCode : string;
     TOBArticle: TOB;
     procedure creeLesTobs;
     procedure detruitLesTobs;
     procedure LoadLesTobs;
     procedure LoadLesArticles;
     procedure TraitelesTobs;
    procedure TraiteLesArt;
    procedure TraiteLesComplementsArt;
    procedure TraiteLesConsommations;
    procedure TraiteLesArtLies;
    procedure TraiteLesArtTiers;
    procedure TraiteLesCatalogues;
    procedure TraiteLesArtPiece;
    procedure TraiteLesArtComm;
    procedure TraiteLesConditArt;
    procedure TraiteLesTradArt;
    procedure TraiteLesNomenEnt;
    procedure TraiteLesNomenLig;
    procedure TraiteLesLigNomen;
    procedure TraiteLesLigOuv;
    procedure TraiteLesLignes;
    // FQ 12048
    procedure TraiteLesOuvP;
    procedure TraiteLesStock;
    // ---
    procedure TraiteLesDecision;
    procedure TraiteActivitePaie;
    procedure TraiteLesTarifs;
    procedure TraiteLesRessources;
     public
     destructor Destroy; override;
     constructor Create;
     procedure TraiteChange;
     end;

{$IFDEF GPAO}
  { Initialisaton des unités dans la fiche article }
  TWInitUniteArticle = Class(TObject)
  Public
    TobArticle: Tob;
    ListIndex: Integer;
    Valid: Boolean;
    constructor Create;
    Destructor Destroy; Override;
    Procedure MakeTobFromDS(DS: TDataSet);
    Procedure PutTobInDS(DS: TDataSet);
    Procedure OpenForm;
  end;

  TwProfil = (paMajArticle, paMajArticleFromProfil, paCreeArticle);
{$ENDIF}

function LibelleArticleGenerique( CodeArticle :string) : string;
procedure VerifCodeArticleUnique;
function CodeArticleUnique ( codeArticle, Dim1, Dim2,Dim3,Dim4, Dim5 : string):string;
function CodeArticleUnique2 ( codeArticle, Dimensions : string):string;
{$IFDEF HDIM}
Procedure ShowDimension (Masque,TypeMasqueDim : String; Dim : THDimension; TobEtab : TOB; DimDepot : String; Valeurs : THDimensionITemList;
                        var AffDimDet : array of boolean;
                        var AffDimTotFin : array of boolean) ;
{$ELSE}
Procedure ShowDimension (Masque,TypeMasqueDim : String; Dim : THDimension; TobEtab : TOB; DimDepot : String; Valeurs : THDimensionITemList;
                        var DimPlus : array of string) ;
{$ENDIF}
function AffecteDim(QMasque: Tquery; i:integer; ItemDim : THDimensionItem): string;
function ControlCodeBarre(var CodeBarre: string; TypeCodeBarre : string ) : Integer ;
function CalculCleCodeBarre(CodeBarre,TypeCodeBarre : string ) : String ;
function ControleCAB(Article,StatutArt,OldCodeCAB:string;var NewCodeCAB:string;QualifCAB:string) : integer ;
function ControleTobCAB(TobArt:TOB) : integer ;
Function CalculMarge( ModeCalcul :Integer; PA,PR,PV : double) : double;
function ChampTobToSQL (NomChamp : string; TOTO : Tob ) :string ;
function CodeArticleGenerique ( var Article, Dim1, Dim2,Dim3,Dim4, Dim5 : string):string;
function CodeArticleGenerique2 ( var Article, Dimensions : string):string;
function IsArticleSpecif (TypeArticle : string) : String;
function IsCodeArticleUnique (codeArticle: string): boolean;

procedure GetArticleMul_Disp (G_CodeArticle: THCritMaskEdit; stWhere, stRange, stMul : string);
procedure GetArticleLookUp_Disp (G_CodeArticle: THCritMaskEdit; stWhere : String; bCodeArticle : Boolean = False);
function  GetArticleRecherche_Disp (G_CodeArticle : THCritMaskEdit;stWhere, stRange, stMul : string; bCodeArticle : Boolean = False) : string;
function  GetRechArticle (G_CodeArticle : THCritMaskEdit;stWhere, stRange, stMul : string) : string;
Function GetChampsArticle (Article, NomChamp : String) : String;
Function PrixModifiable (NomChamp : string; PrixUnique, Defaut : boolean) : boolean ;
Function  TrouverArticle (G_CodeArticle: THCritMaskEdit; ToBArt : TOB) : T_RechArt ; overload;
Function  TrouverArticle (G_CodeArticle: string; ToBArt : TOB) : T_RechArt ; overload;
Function  ControlArticle(Article,CodeArticle: String;Dim:Boolean=False;TypeArticle:string='NOR'): String ;
function NaturePieceGeree (naturePiece : string) : boolean ;
//
procedure FiltreComboTypeArticle(TYPEARTICLE : THValComboBox); overload;
procedure FiltreComboTypeArticle(TYPEARTICLE : THMultiValComboBox); overload;
function  PrepareValeurFiltreCombo : string;
//
function nbDecimales(NomChamp,Typ,DimDecQte,DimDecPrix : string) : string ;
function reinit_combo (cb : THValComboBox) : boolean ;
// Modif BTP
function ArticleOKInPOUR (TypeArt,Chaine : string) : boolean;
function RecupTypeGraph (TOBExamin : Tob) : Integer;
{$IFDEF BTP}
function GetTypeArticleBTP : string;
function GetTypeArticlePARC : string;
Function  RenvoieTypeArt (CodArt : string) : String ;
procedure chargeTOBs (TOBRef,TOBCatalog,TOBTiers,TOBTarif : TOB;QteRef : double = 1; DateRef : TdateTime =0);
//FV1 : 18/02/2015 - Prix d'achat sur fiche Article
function PassageUAUV (TOBRef,TOBcatalogue : TOB;MTPAF: double) : double; OverLoad;
function PassageUAUV (TobCatalogue : TOB; MTPAF, PrixPourQte, CoefUniteAUniteS, CoefUniteSUniteV : Double; UniteA, UniteS, UniteV : String) : double; OverLoad;
Function CalculUaUv  (CoefUaUs,CoefUSUv,MTPAF,PrixPourQte : Double; UA, UV, US : String) : Double;
{$ENDIF}
function ChangeCodificationArticle (CodeArticle,NewCode : String) : boolean;
Procedure ChangementCodeTarif;
Function ChangeCodificationTarif (CodeTarif,NewCode,TraiteCodeArt : String) : Boolean ;
Function CreationIndex (TraiteCodeArt : String) : Boolean;
// --
function GetPresentation : integer ;
function DispatchArtMode (Num : integer; Range,Lequel,Argument : String) : variant ;
procedure chargeTobPrestation (Prestation : String;TOBPrestation:TOB);
procedure GetDetailPrixPose(CodeArticle : string;TOBPrestation,TOBEDetail,TOBDetail: TOB);
{$IFDEF BTP}
procedure RecalculPrPV (TOBRef,TOBCatalog : TOB;VenteAchat : string='VEN');
procedure RecalculPr (TOBRef,TOBCatalog : TOB;VenteAchat : string='VEN');
procedure RecalculPvHt (TOBRef,TOBCatalog : TOB;VenteAchat : string='VEN');
procedure RecalculPvTTC (TOBRef,TOBCatalog : TOB;VenteAchat : string='VEN');
function GetTypeRessourceFromNature (Naturepres : string) : string;
procedure GetInfoFromCatalog(Fournisseur, Article: string; var UA : string; Var PQQ : double; var CoefuaUs: double);
procedure GetInfoFromArticle(Article: string; var UA : string; Var PQQ : double; var CoefuaUs: double);
{$ENDIF}
procedure InitValoArtNomen (TOBRef : TOB;VenteAchat : string='VEN');
{$IFDEF STK}
Function ConvUnite(Qte : Double; sUdepart, sUArrive, sArticle, sTypeDepart, sTypeArrive : String) : Double; { GPAO STKMOUV }
{$ENDIF}
function wIsSuffixeCodeArticle(Const FieldName: String): Boolean;

{$IFDEF BTP}
procedure CreationEnteteOuv(CodeArt,Libelle,Domaine : String);
procedure BTPDuplicOuvrage(Source, Destination : string);
Function isExistsArticle (Article : string) : boolean;
{$ENDIF}
procedure InitValoArtLigBord (TOBRef : TOB;VenteAchat : string='VEN');
procedure InitChampsSupArticle (TOBA : TOB);
function GetStatutArt(Const Article: string): string;

procedure  recalculCoefMarg (TOBRef : TOB);
procedure  recalculCoef(TOBRef : TOB);
Function wCreateArticleFromTob(Const Article: String; TheTob: Tob): Boolean;
function wUpDateArticleFromTob(Article: String; TheTob: Tob): Boolean;
function GetprixRemise (Article : string; TypeInfo : string) : double;

Procedure DuplicationArticleMetre(TypeArticle, CodeArticle, Duplicart : String);


{ ------------------------------------------------------------------------------- }
{ ----------------------------- UTILITAIRES GPAO -------------------------------- }
{ ------------------------------------------------------------------------------- }
  { Appel de la fiche article }
  procedure LanceFicheArticle(Lequel: String; Action: TActionFiche = taModif; Params: String = ''; Range : String = '');
  { Appel du mul article }
  procedure LanceMulArticle(Range: String = ''; Lequel: String = ''; Params: String = '');

  function  whereGA(Article: string): string;
  { Retourne le premier article }
  function  wGetFirstGA: String;
  { Retourne le dernier article }
  function  wGetLastGA: String;

  function wExistGA(Article: string; WithAlert: Boolean = false): Boolean;
  function  wGetArticleFromCodeArticle(CodeArticle : String; CodeDim1: string = ''; CodeDim2: string = ''; CodeDim3: string = ''; CodeDim4: string = ''; CodeDim5: string = '') : String;
  function  wGetCodeArticleFromArticle(Article: string): string;
  function  wGetArticleGENFromArticleDIM(Article: string): string;
  function  wGetTobGA(sChamp, sWhere: String; TobGA: Tob; SelectDB:Boolean = False): Boolean;

  function  wGetQualifUniteArticle(CodeChamp: String): String;
  { Retourne le Code Article pour une Référence client }
  function GetArticleFromReferenceClient(RefArticleTiers: String): String;
  function GetArticleTiers(Article, RefTiers: String): String;
  function CreateArticleTiers(TobGAT: Tob): boolean;
  function ExistArticleTiers(Article, Tiers: String): Boolean;
  function WhereRefArticleTiers(sTiers, sRefArticle: String): String;
  function WhereArticleTiers(sTiers, sArticle: String): String;

{$IFDEF GPAO}
  {---- Divers ----}
  function uGetStatutArticle (Article:string) :string; // Statut article (UNI, GEN ou DIM)
  function  wExistArticleInGPF(Article: string; WithAlert: Boolean = false): Boolean;
  function  wExistArticleProfilInGA(Article: string; WithAlert: Boolean = false): Boolean;
  procedure wCallWAN(Article: String; Action: TActionFiche; Params: String);
  function  wGetFieldsFromGA(Fields: Array of string; Article: string): MyArrayValue;
	function  wIsSuffixeCodeArticle(NomChamp: String): Boolean;
	function  wIsSuffixeArticle(NomChamp: String): Boolean;
	function  xSetCodeArticle(Var SaisArticle: String): Boolean;
  {---- Configurateur ----}
  function  wDuplicGA(Article, ArticleProfil: String): String;
  function  wUpDateArticleFromTob(Article: String; TheTob: Tob): String;
  {---- Gestion des profils avancés ----}
  function  wSetAdvancedProfil(From: TwProfil; DS: TDataSet): Integer;
  function  wSetAdvancedProfil2(From: TwProfil; TobArticle: Tob): Integer;
  function  wSetAdvancedProfilToArticle(DSArticle: TDataSet; QProfil: TQuery): Boolean;
  function  wSetAdvancedProfilToArticleByTob(TobArticle, TobProfil: TOb): Boolean;
  procedure wGetProfilFieldList(TobProfil: Tob; TS: tStrings);
  procedure wGetAdvancedProfilFieldsList(TobProfil, TobUnAutorizedFields: Tob; TS: tStrings);
  procedure wLoadTobUnAutorizedFieldsInProfil(TheTob: Tob);
  { ---- }
  function GetWInitArticle(iObject: Integer): TWInitUniteArticle;
  { ---- }
  function GetMesureFromFlux(Article, Flux: string): string;
  { **** }
  function GetArticleFromCB(cb: String): String;
  function GetFieldFromArticleTiers(Field, sTiers, sArticle: string; ByRefArticle: Boolean = False): Variant;
{$ENDIF}
  function  wGetFieldFromGA(Field, Article: string): Variant;
  procedure wCallGA(Range: String);
  function wGetListeNaturePiecG(Compteur: string): string;


Const
    ART_S3            = 1 ;
    ART_ORLI          = 2 ;
    TailleArticle     = 35;
    TailleCodeArticle = 18;
    TailleCodeDim     = 3;


implementation
Uses
{$IFDEF MODE}
     ArtPrestation_Tof,
{$ENDIF}
     UtilTarif,
     FactUtil,FactArticle, UDimArticle,UTomArticle,ParamSoc,
     UtilDispGC, AglInitGC
{$IFDEF BTP}
     ,NomenUtil
     //,UmetreArticle
{$ENDIF}
{$IFDEF GPAO}
     ,wArtNat
     ,ed_tools
     ,M3FP
{$ENDIF}
			,uTofListeInv
      ,UtilsMetresXLS
     ;
{$IFDEF GPAO}
Const
	{ Libellés des messages lors de l'application d'un profil }
	TexteMessageProfil: array[1..10] of string 	= (
          {1}      'Paramètre d''appel incorrect'
          {2}     ,'Aucun article n''utilise ce profil'
          {3}     ,'Code profil inconnu'
          {4}     ,'Le profil n''est pas un profil avancé'
          {5}     ,'Mode de copie du profil incorrect'
          {6}     ,'Mise à jour abandonnée'
          {7}     ,'Les articles utilisant ce profil ne seront pas mis à jour.' +#13 + 'Le profil de cet article n''est pas paramètré pour être mis à jour depuis la fiche article.'
          {8}     ,'Mode de remplacement du profil incorrect'
          {9}     ,'L''article de référence du profil est inconnu'
          {10}    ,'L''article de référence du profil n''est pas renseigné'
          );
{$ENDIF}
var
{$IFDEF GPAO}
  wListInitArticle: TList;
{$ENDIF}
  MetreArticle: TMetreArt;


function GetprixRemise (Article : string; TypeInfo : string) : double;
var TOBART : TOB;
begin
  TOBART := TOB.Create('ARTICLE',nil,-1);
  TOBART.SelectDB('"'+Article+'"',nil);
  InitValoArtNomen (TOBART);
  if TypeInfo = 'ACHAT' then Result := TOBART.GetDouble('GA_PAHT')
  else if TypeInfo = 'VENTE' then Result := TOBART.GetDouble('GA_PVHT');
  TOBART.Free;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 19/01/2000
Modifié le ... :   /  /
Description .. : retourne le libellé d'un article à partir du code générique
Mots clefs ... : ARTICLE;DIMENSION;GENERIQUE
*****************************************************************}
function LibelleArticleGenerique( CodeArticle : string ) : string;
var  qq : TQuery ;
begin
Result:='' ;
QQ:=OpenSQL('SELECT GA_LIBELLE FROM ARTICLE WHERE GA_CODEARTICLE="'+ CodeArticle +'" and (GA_STATUTART="GEN" OR GA_STATUTART="UNI")',true,-1,'',true);
if not QQ.EOF then
    Result:=QQ.findField('GA_LIBELLE').asString ;
ferme(QQ);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Pierre LENORMAND
Créé le ...... : 31/05/2000
Modifié le ... : 31/05/2000
Description .. : Retourne le champs de l'article passé en argument
Mots clefs ... : ARTICLE
*****************************************************************}
Function GetChampsArticle (Article, NomChamp : String) : String;
Var Q : TQuery ;
BEGIN
Result := '';
Q := OpenSQL ('SELECT '+NomChamp + ' From ARTICLE WHERE GA_ARTICLE="' + Article +'"',True,-1,'',true);
Try
 if Not Q.EOF then Result := Q.Findfield(NomChamp).AsString;
 Finally
 Ferme(Q);
 END;
END;

{ Charge les combos des dimensions avec :
  - soit Plus et DataType si renseignés (dimension classique),
  - soit les établissements de TobEtab (dimension dépôt).
}
procedure ChargeComboDim(var CB : THValComboBox ; Plus,DataType : string ; TobEtab : TOB) ;
var iTob : integer ;
begin
if DataType<>'' then
    BEGIN
    CB.Plus:=Plus ;
    CB.DataType:=DataType ;
    END else
    BEGIN
    CB.Items.Clear ; CB.Values.Clear ;
    if TobEtab<>nil then for iTob:=0 to TobEtab.Detail.Count-1 do
        BEGIN
        CB.Values.Add(TobEtab.Detail[iTob].GetValue('GUD_ETABLISSEMENT')) ;
        CB.Items.Add(TobEtab.Detail[iTob].GetValue(Plus)) ; // Plus contient le nom du champ à afficher dans le THDIM
        END ;
    END ;
end ;

{$IFDEF HDIM}
Procedure ShowDimension (Masque,TypeMasqueDim : String; Dim : THDimension; TobEtab : TOB; DimDepot : String; Valeurs : THDimensionITemList;
                        var AffDimDet : array of boolean;
                        var AffDimTotFin : array of boolean) ;
{$ELSE}
Procedure ShowDimension (Masque,TypeMasqueDim : String; Dim : THDimension; TobEtab : TOB; DimDepot : String; Valeurs : THDimensionITemList;
                        var DimPlus : array of string) ;
{$ENDIF}
Var Q : TQuery ;
    St,St2,Plus,DataType,Titre : String ;
    i, MaxDim : Integer ;
    TitreOngl,TitreLig1,TitreLig2,TitreCol1,TitreCol2 : string ;
    DiOn,DiL1,DiL2,DiC1,DiC2 : THValComboBox ;
    Par : TForm ;
{$IFDEF HDIM}
    TotOnglFin,TotLig1Fin,TotLig2Fin,TotCol1Fin,TotCol2Fin : boolean ;
    DetOngl,DetLig1,DetLig2,DetCol1,DetCol2 : boolean ;
    iVal : integer ;
    stFormat, stF : string ;
{$ENDIF}
    iDim : integer ;
    lgMaxColWidth,lgMaxCol1Width,lgMaxCol2Width,lgMaxColValWidth,lg : integer ;
    Position : array[1..MaxDimension+1] of string ;
BEGIN
if Dim=Nil then exit ;
Dim.FreeCombo:=TRUE ;
Dim.Active:=FALSE ;
Dim.FreeCombo:=FALSE ;
//Dim.TypeDonnee:=otReel ;
Par:=TForm(Dim.OWner) ;
DiOn:=nil ; DiL1:=nil ; DiL2:=nil ; DiC1 :=nil ; DiC2:=nil ;
// Doublage de la requête : prise en compte du type de masque en contexte mode
if GetParamSoc('SO_GCARTTYPEMASQUE') then
   BEGIN
   Q:=OpenSQL('SELECT * FROM DIMMASQUE WHERE GDM_MASQUE="'+Masque+'" and GDM_TYPEMASQUE="'+TypeMasqueDim+'"',TRUE,-1,'',true) ;
   if (Q.EOF) and (TypeMasqueDim<>VH_GC.BOTypeMasque_Defaut) then
       BEGIN
       Ferme(Q) ;
       Q:=OpenSQL('select * from DIMMASQUE where GDM_MASQUE="'+Masque+'" and GDM_TYPEMASQUE="'+VH_GC.BOTypeMasque_Defaut+'"',TRUE,-1,'',true) ;
       END ;
   MaxDim:=MaxDimension+1 ;
   END else
   BEGIN
   Q:=OpenSQL('SELECT * FROM DIMMASQUE WHERE GDM_MASQUE="'+Masque+'"',TRUE,-1,'',true) ;
   MaxDim:=MaxDimension ;
   END ;
if Q.EOF then BEGIN Ferme(Q) ; Exit ; END ;
iDim:=1 ;
for i:=1 to MaxDim do
   BEGIN
   St:=Q.FindField('GDM_POSITION'+IntToStr(i)).AsString ;
   if Not ((GetParamSoc('SO_GCARTTYPEMASQUE')) and (i=MaxDim)) then St2:=Q.FindField('GDM_TYPE'+IntToStr(i)).AsString ;
   // Prise en compte de la position de l'établissement en contexte mode : GDM_POSITION6
   if (GetParamSoc('SO_GCARTTYPEMASQUE')) and (i=MaxDim) then
      BEGIN
      Plus:=DimDepot ; DataType:='' ;
      Titre:=TraduireMemoire('Etablissement') ;
      Position[iDim]:=St ; inc(iDim) ;
      END else
      BEGIN
      // DimPlus[i-1] car indice tableau passé en paramètre commence à 0 !!
      Plus:='GDI_TYPEDIM="DI'+IntToStr(i)+'" AND GDI_GRILLEDIM="'+St2+'" '+DimPlus[i-1]+'ORDER BY GDI_RANG' ;
      //Plus:='GDI_TYPEDIM="DI'+InttoStr(i)+'" AND GDI_GRILLEDIM="'+St2+'" ORDER BY GDI_RANG' ;
      DataType:='GCDIMENSION' ;
      Titre:=RechDom('GCCATEGORIEDIM','DI'+IntToStr(i),FALSE) ;
      Position[iDim]:=St ; inc(iDim) ;
      END ;
   if St='LI1' then
      BEGIN
      DiL1:=THValComboBox.Create(Par) ; DiL1.Parent:=Par ; DiL1.Visible:=FALSE ;
      ChargeComboDim(DiL1,Plus,DataType,TobEtab) ;
      TitreLig1:=Titre ;
      END
   else if St='LI2' then
      BEGIN
      DiL2:=THValComboBox.Create(Par) ; DiL2.Parent:=Par ; DiL2.Visible:=FALSE ;
      ChargeComboDim(DiL2,Plus,DataType,TobEtab) ;
      TitreLig2:=Titre ;
      END
   else if St='CO1' then
      BEGIN
      DiC1:=THValComboBox.Create(Par) ; DiC1.Parent:=Par ; DiC1.Visible:=FALSE ;
      ChargeComboDim(DiC1,Plus,DataType,TobEtab) ;
      TitreCol1:=Titre ;
      END
   else if St='CO2' then
      BEGIN
      DiC2:=THValComboBox.Create(Par) ; DiC2.Parent:=Par ; DiC2.Visible:=FALSE ;
      ChargeComboDim(DiC2,Plus,DataType,TobEtab) ;
      TitreCol2:=Titre ;
      END
   else if St='ON1' then
      BEGIN
      DiOn:=THValComboBox.Create(Par) ; DiOn.Parent:=Par ; DiOn.Visible:=FALSE ;
      ChargeComboDim(DiOn,Plus,DataType,TobEtab) ;
      TitreOngl:=Titre ;
      END ;
   END ;
Ferme(Q) ;


// Ajustement de la largeur des colonnes
// =====================================

// Recherche des longueurs maxi des titres des colonnes (en nb de caractères)
lgMaxColWidth:=0 ; lgMaxCol1Width:=0 ; lgMaxCol2Width:=0 ; lgMaxColValWidth:=0 ;
if Dim.NbValeurs>1 then
  BEGIN
  St:=Dim.TitreValeur ;
  for i:=1 to Dim.NbValeurs do
    BEGIN
    lg:=length(ReadTokenSt(St)) ;
    if lg>lgMaxColValWidth then lgMaxColValWidth:=lg ;
    END ;
  END ;
if DiL1<>Nil then
  for i:=0 to DiL1.Items.Count-1 do
    if length(DiL1.Items[i])>lgMaxCol1Width then lgMaxCol1Width:=length(DiL1.Items[i]) ;
if DiL2<>Nil then
  for i:=0 to DiL2.Items.Count-1 do
    if length(DiL2.Items[i])>lgMaxCol2Width then lgMaxCol2Width:=length(DiL2.Items[i]) ;
if DiC2<>Nil then
  BEGIN
  for i:=0 to DiC2.Items.Count-1 do
    if length(DiC2.Items[i])>lgMaxColWidth then lgMaxColWidth:=length(DiC2.Items[i]) ;
  END
else if DiC1<>Nil then
  for i:=0 to DiC1.Items.Count-1 do
    if length(DiC1.Items[i])>lgMaxColWidth then lgMaxColWidth:=length(DiC1.Items[i]) ;

// Maj des tailles de colonnes par défaut
lgMaxColWidth:=lgMaxColWidth*6 ; lgMaxCol1Width:=lgMaxCol1Width*6 ;
lgMaxCol2Width:=lgMaxCol2Width*6 ; lgMaxColValWidth:=lgMaxColValWidth*6 ;
if lgMaxColWidth>Dim.DefaultColWidth then Dim.DefaultColWidth:=lgMaxColWidth ;
if lgMaxCol1Width>Dim.DefaultCol1Width then Dim.DefaultCol1Width:=lgMaxCol1Width ;
if lgMaxCol2Width>Dim.DefaultCol2Width then Dim.DefaultCol2Width:=lgMaxCol2Width ;
if lgMaxColValWidth>Dim.DefaultColValWidth then Dim.DefaultColValWidth:=lgMaxColValWidth ;

{$IFDEF HDIM}
// Affichage (ou non) des totaux/détails
TotOnglFin:=False ; TotLig1Fin:=False ; TotLig2Fin:=False ; TotCol1Fin:=False ; TotCol2Fin:=False ;
DetOngl:=True ; DetLig1:=True ; DetLig2:=True ; DetCol1:=True ; DetCol2:=True ;
for iDim:=1 to MaxDimension do
     BEGIN
     if Position[iDim]='ON1' then
         BEGIN
         TotOnglFin:=AffDimTotFin[iDim-1] ;
         DetOngl:=AffDimDet[iDim-1] ;
         END
     else if Position[iDim]='LI1' then
         BEGIN
         TotLig1Fin:=AffDimTotFin[iDim-1] ;
         DetLig1:=AffDimDet[iDim-1] ;
         END
     else if Position[iDim]='LI2' then
         BEGIN
         TotLig2Fin:=AffDimTotFin[iDim-1] ;
         DetLig2:=AffDimDet[iDim-1] ;
         END
     else if Position[iDim]='CO1' then
         BEGIN
         TotCol1Fin:=AffDimTotFin[iDim-1] ;
         DetCol1:=AffDimDet[iDim-1] ;
         END
     else if Position[iDim]='CO2' then
         BEGIN
         TotCol2Fin:=AffDimTotFin[iDim-1] ;
         DetCol2:=AffDimDet[iDim-1] ;
         END
     END ;
// Gestion affichage ou non du total des 1..10 valeurs
// Pour l'instant, pas de cumuls sur les champs décimaux (à priori = prix unitaire donc non cumulable)
stFormat:=Dim.DisplayFormat ;
stF:=Trim(ReadTokenSt(stFormat)) ; Dim.AffTotVal1 :=(pos('#',stF)>0) and (pos('.',stF)=0) ;
stF:=Trim(ReadTokenSt(stFormat)) ; Dim.AffTotVal2 :=(pos('#',stF)>0) and (pos('.',stF)=0) ;
stF:=Trim(ReadTokenSt(stFormat)) ; Dim.AffTotVal3 :=(pos('#',stF)>0) and (pos('.',stF)=0) ;
stF:=Trim(ReadTokenSt(stFormat)) ; Dim.AffTotVal4 :=(pos('#',stF)>0) and (pos('.',stF)=0) ;
stF:=Trim(ReadTokenSt(stFormat)) ; Dim.AffTotVal5 :=(pos('#',stF)>0) and (pos('.',stF)=0) ;
stF:=Trim(ReadTokenSt(stFormat)) ; Dim.AffTotVal6 :=(pos('#',stF)>0) and (pos('.',stF)=0) ;
stF:=Trim(ReadTokenSt(stFormat)) ; Dim.AffTotVal7 :=(pos('#',stF)>0) and (pos('.',stF)=0) ;
stF:=Trim(ReadTokenSt(stFormat)) ; Dim.AffTotVal8 :=(pos('#',stF)>0) and (pos('.',stF)=0) ;
stF:=Trim(ReadTokenSt(stFormat)) ; Dim.AffTotVal9 :=(pos('#',stF)>0) and (pos('.',stF)=0) ;
stF:=Trim(ReadTokenSt(stFormat)) ; Dim.AffTotVal10:=(pos('#',stF)>0) and (pos('.',stF)=0) ;
// Affichage THDIM
Dim.InitDimension(TitreOngl,TitreLig1,TitreLig2,TitreCol1,TitreCol2,DiOn,DiL1,DiL2,DiC1,DiC2,Valeurs,
                  False,False,False,False,False,
                  TotOnglFin,TotLig1Fin,TotLig2Fin,TotCol1Fin,TotCol2Fin,
                  DetOngl,DetLig1,DetLig2,DetCol1,DetCol2,
                  'Total','Total','Total','Total','Total') ;
{$ELSE}
Dim.InitDimension(TitreOngl,TitreLig1,TitreLig2,TitreCol1,TitreCol2,DiOn,DiL1,DiL2,DiC1,DiC2,Valeurs) ;
{$ENDIF}
END ;

procedure VerifCodeArticleUnique;
var QQ:Tquery ;
    SQL:string ;
begin
QQ:=OpenSQL('Select GA_ARTICLE,GA_CODEARTICLE,GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,GA_STATUTART from ARTICLE',TRUE,-1,'',true);
while Not QQ.EOF do
    BEGIN
    if (QQ.Findfield('GA_STATUTART').AsString='DIM') then
       begin
       SQL:='UPDATE ARTICLE set GA_ARTICLE="' ;
       SQL:=SQL+CodeArticleUnique (QQ.Findfield('GA_CODEARTICLE').AsString
                                    ,QQ.Findfield('GA_CODEDIM1').AsString
                                    ,QQ.Findfield('GA_CODEDIM2').AsString
                                    ,QQ.Findfield('GA_CODEDIM3').AsString
                                    ,QQ.Findfield('GA_CODEDIM4').AsString
                                    ,QQ.Findfield('GA_CODEDIM5').AsString);
       SQL:=SQL+'" Where GA_ARTICLE="'+QQ.Findfield('GA_ARTICLE').AsString+'"' ;
       ExecuteSQL(SQL);
       end;
       QQ.Next;
    end;
ferme(QQ);

end;
function AffecteDim(QMasque: Tquery; i:integer; ItemDim : THDimensionItem): string;
var MasquePosition : string ;
begin
    MasquePosition:=QMasque.Findfield('GDM_POSITION'+IntToStr(I)).AsString ;
    if MasquePosition='ON1' then Result:=ItemDim.Ong
    else if  MasquePosition='LI1' then Result:=ItemDim.Lig1
    else if  MasquePosition='LI2' then Result:=ItemDim.Lig2
    else if  MasquePosition='CO1' then Result:=ItemDim.Col1
    else if  MasquePosition='CO2' then Result:=ItemDim.Col2
    else Result:='';
end;

Function CalculMarge( ModeCalcul :Integer; PA,PR,PV : double) : double;
begin
Result := 0 ;
if PV <> 0 then
  begin
    case ModeCalcul of
        0:  if PA <> 0 then Result := PV/PA ;
        1:  if PR <> 0 then Result := PV/PR ;
        2:  if PA <> 0 then Result := ((PV - PA) / PA) * 100 ;
        3:  Result := ((PV - PA) / PV) * 100 ;
        4:  if (PR<> 0) then Result := ((PV - PR) / PR) * 100 ;
        5:  Result := ((PV - PR) / PV) * 100 ;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 19/01/2000
Modifié le ... : 25/01/2000
Description .. : Controle d'un code barre en fonction de son type
Mots clefs ... : ARTICLE;CODE BARRE;
*****************************************************************}
{--> type de codes controlés :
EAN8    : TypeCodeBarre=EA8
EAN13   : TypeCodeBarre=E13
Code 39 : TypeCodeBarre=39 ou TypeCodeBarre=39C (code 39 controlé avec clé)
}
function ControlCodeBarre(var CodeBarre: string; TypeCodeBarre : string ) : Integer ;
var  Sum,MaxLength,Index, I  : integer;
const
     CarCode39 : String ='0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%' ;
Begin
  Sum := 0;
  Result :=0;

  /////  Code EAN8 et EAN 13 ////////////
  if (TypeCodeBarre ='EA8') or (TypeCodeBarre ='E13') then
     Begin
     if (TypeCodeBarre ='EA8') then MaxLength := 8 else MaxLength := 13;
     if (not IsNumeric(CodeBarre))then
         begin
         Result :=4;  // code doit être numérique
         exit;
         end;
     if (length(CodeBarre)=(MaxLength-1)) then
         begin
         // Ajout automatique de la clé de contrôle
         CodeBarre := CodeBarre + CalculCleCodeBarre(CodeBarre+'0', TypeCodeBarre );
         end;
     if (length(CodeBarre)<>MaxLength) then
         begin
         Result :=1;  // manque ou trop de caractères
         exit;
         end;
     if (CodeBarre[length(CodeBarre)] <> CalculCleCodeBarre(CodeBarre, TypeCodeBarre)) then Result := 2;  // Clé incorrecte
     End;

   ///////  Code 39   /////////////////
   if (TypeCodeBarre ='39') or (TypeCodeBarre ='39C')then
      begin
      for index := length(CodeBarre)-1 downto 1 do
          begin
          I := Pos (CodeBarre[index], CarCode39);
          if (I>0) then
              Sum := Sum + I
          else
               begin
               Result := 3;
               exit;
               end;
          end;
      if (TypeCodeBarre ='39C') then
         begin
           if (CodeBarre[length(CodeBarre)] <> CarCode39[Sum mod 43]) then Result := 2;  // Clé incorrecte
         end;
      end;

  /////  Code 2/5 Entrelacé ou 2/5 avec contrôle ////////////
  if (TypeCodeBarre ='ITC') or (TypeCodeBarre ='ITF') then
     Begin
     if (not IsNumeric(CodeBarre))then
         begin
         Result :=4;  // code doit être numérique
         exit;
         end;
     if (length(CodeBarre) mod 2)<>0 then
         begin
         // Le code à barres doit être de longueur paire, ajout automatique de la clé de contrôle
         CodeBarre := CodeBarre + CalculCleCodeBarre(CodeBarre+'0', TypeCodeBarre ) ;
         end;
     if (TypeCodeBarre ='ITC') and (CodeBarre[length(CodeBarre)] <> CalculCleCodeBarre(CodeBarre, TypeCodeBarre)) then Result := 2;  // Clé incorrecte
     end;
  /////  Code EAN 18 ////////////
  if (TypeCodeBarre ='E18') then
  begin
    MaxLength := 18;
    if (not IsNumeric(CodeBarre))then
    begin
      Result := 4;  // code doit être numérique
      Exit;
    end;
    if (length(CodeBarre) = (MaxLength - 1)) then
    begin
      // Ajout automatique de la clé de contrôle
      CodeBarre := CodeBarre + CalculCleCodeBarre(CodeBarre+'0', TypeCodeBarre );
    end;
    if length(CodeBarre) <> 18 then
    begin
      Result := 1;  // manque ou trop de caractères
      Exit;
    end;
    if (CodeBarre[length(CodeBarre)] <> CalculCleCodeBarre(CodeBarre, TypeCodeBarre)) then
      Result := 2;  // Clé incorrecte
  end;
  /////  Code EAN Presse (idem EAN 13, avec 5 chiffres aprés la clé.) /////
  if (TypeCodeBarre ='EPF') then
  begin
    {MaxLength := 18; JLS Conseil}
    if (not IsNumeric(CodeBarre))then
    begin
      Result := 4;  // code doit être numérique
      Exit;
    end;
    if length(CodeBarre) <> 18 then
    begin
      Result := 1;  // manque ou trop de caractères
      Exit;
    end;
    if (CodeBarre[13] <> CalculCleCodeBarre(Copy(CodeBarre, 0, 13), 'E13')) then
      Result := 2;  // Clé incorrecte
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : AV
Créé le ...... : 09/01/2001
Modifié le ... :
Description .. : Calcul de la clé d'un code barre en fonction de son type
Mots clefs ... : ARTICLE;CODE BARRE;CLE;
*****************************************************************}
function CalculCleCodeBarre(CodeBarre : string; TypeCodeBarre : string ) : String ;
var  Factor,Sum,Index, I, chiffre  : integer;
const
     CarCode39 : String ='0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%' ;
Begin
  factor := 3;
  Sum := 0;
  Result :='';

  /////  Code EAN8 et EAN 13 ////////////
  if (TypeCodeBarre ='EA8') or (TypeCodeBarre ='E13') then
     Begin
     for index := length(CodeBarre)-1 downto 1 do
         begin
         Sum := Sum + strtoint(CodeBarre[index]) * factor;
         factor := 4 - factor;
         end;
     Result := inttostr((1000 - sum) mod 10);
     End;

   ///////  Code 39C   /////////////////
   if (TypeCodeBarre ='39C')then
      begin
      for index := length(CodeBarre)-1 downto 1 do
          begin
          I := Pos (CodeBarre[index], CarCode39);
          if (I>0) then
              Sum := Sum + I
          else exit;
          end;
      Result := CarCode39[Sum mod 43];
      end;

  /////  Code 2/5 Entrelacé ////////////
  if (TypeCodeBarre ='ITC') or (TypeCodeBarre ='ITF') then
     Begin
     for index := 1 to length(CodeBarre)-1 do
         begin
         if (CodeBarre[index]>='0') and (CodeBarre[index]<='9') then
            begin
            chiffre:=strtoint(CodeBarre[index]) ;
            if (index mod 2)<>0 then
               begin
               if (chiffre*2)<10 then Sum:=Sum+(chiffre*2)
                                 else Sum:=Sum+(chiffre*2)-9 ;
               end else
               Sum:=Sum+chiffre ;
            end;
         end;
     chiffre:=Sum mod 10 ;
     if chiffre=0 then chiffre:=10 ;
     Result := inttostr(10-Chiffre);
     End;
end;

// Fonction de contrôle du format d'un code à barres et vérification de l'unicité de celui-ci
function ControleCAB (Article,StatutArt,OldCodeCAB:string;var NewCodeCAB:string;QualifCAB:string) : integer ;
var CodeCAB,Sql : String ;
begin
Result:=0 ;
if (NewCodeCAB<>'') and (QualifCAB='') then
   BEGIN
   {si article GEN pas une erreur mais pas de controle}
   if (StatutArt<>'GEN') then BEGIN Result:=5 ; exit END ;
   END else
   BEGIN
   if (NewCodeCAB<>'') and (QualifCAB<>'') then
       BEGIN
       CodeCAB:=NewCodeCAB ;
       case ControlCodeBarre(CodeCAB,QualifCAB) of
            1: Result:=6 ;
            2: Result:=7 ;
            3: Result:=8 ;
            4: Result:=9 ;
            END ;
       if Result<>0 then exit ;
       NewCodeCAB:=CodeCAB ;

       // Verification si le code à barres n'existent pas déjà
       if (GetParamsoc('SO_GCCABARTICLE')) and (OldCodeCAB<>NewCodeCAB) then
          BEGIN
          Sql:='select GA_CODEBARRE from ARTICLE where GA_CODEBARRE="'+NewCodeCAB+'"' ;
          if Article<>'' then Sql:=Sql+' and GA_ARTICLE<>"'+Article+'"' ;
          if (ExisteSQL(Sql))
             then BEGIN Result:=39 ; exit END ;
          END ;
       END ;
   END ;
end ;

// Contrôle du CAB d'une Tob Article (dimensionnée donc incomplète=ne contient pas tous les champs de la fiche article)
function ControleTobCAB(TobArt:TOB) : integer ;
var CodeCAB : string ;
begin
//if TobArt.IsFieldModified('GA_CODEBARRE') then
if TobArt.GetValue('GA_CODEBARRE')<>TobArt.GetValue('ZGA_CODEBARRE') then
    BEGIN
    CodeCAB:=TobArt.GetValue('GA_CODEBARRE') ;
    Result:=ControleCAB(TobArt.GetValue('GA_ARTICLE'),TobArt.GetValue('GA_STATUTART'),'',CodeCAB,TobArt.GetValue('GA_QUALIFCODEBARRE')) ;
    if (Result=0) and (CodeCAB<>TobArt.GetValue('GA_CODEBARRE')) then TobArt.PutValue('GA_CODEBARRE',CodeCAB) ;
    if (Result=0) and (CodeCAB<>TobArt.GetValue('ZGA_CODEBARRE')) then TobArt.PutValue('ZGA_CODEBARRE',CodeCAB) ;
    END else Result:=0 ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 21/01/2000
Modifié le ... : 12/09/2003
Description .. : retourne la valeur d'un champ d'une TOB sous la forme d'un
Suite ........ : string pouvant être utilisé dans un ordre SQL
Mots clefs ... : SQL;TOB;CONVERSION
*****************************************************************}
function ChampTobToSQL (NomChamp : string; TOTO : Tob ) :string ;
var stTmp,Typ : string ;
    posLgrChamp,lgrChamp : integer ;
begin
typ:=ChampToType( NomChamp);
if TOTO.GetValue(NomChamp)<>null then
    BEGIN
    if (Typ='INTEGER') or (Typ='SMALLINT') then result:=IntToStr(VarAsType(TOTO.GetValue(NomChamp),VArInteger))  else
    if (Typ='DOUBLE') or (Typ='RATE')  then result:=STRFPOINT(VarAsType(TOTO.GetValue(NomChamp),VArDouble))  Else
    if Typ='DATE'    then result:='"'+USDateTime(VarAsType(TOTO.GetValue(NomChamp),VArDate))+'"' Else
    if Typ='BOOLEAN' then result:='"'+VarAsType(TOTO.GetValue(NomChamp),VArString)+'"'  else
        BEGIN // Limite la longueur de la donnée par rapport à la longueur du champ dans la base
        stTmp:=VarAsType(TOTO.GetValue(NomChamp),VarString) ;
        posLgrChamp:=pos('(',typ)+1 ;
        if posLgrChamp>1 then lgrChamp:=strToInt(copy(typ,posLgrChamp,pos(')',typ)-posLgrChamp)) else lgrChamp:=9999 ;
        if length(stTmp)>lgrChamp then result:='"'+copy(stTmp,1,lgrChamp)+'"'
        else result:='"'+stTmp+'"' ;
        END ;
    END else
    BEGIN // Correctif pb base Oracle chez QUIKSILVER
    if (Typ='INTEGER') or (Typ='SMALLINT') then result:='0'  else
    if (Typ='DOUBLE') or (Typ='RATE')  then result:='0.0' else
    if Typ='DATE'    then result:='"01/01/1900"' else
    if Typ='BOOLEAN' then result:='"-"' else
        result:='""' ;
    END ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 19/01/2000
Modifié le ... : 02/02/2000
Description .. : retourne le code article unique à partir du code générique
Mots clefs ... : CODE;ARTICLE;GENERIQUE;UNIQUE;DIMENSION
*****************************************************************}
function CodeArticleUnique ( codeArticle, Dim1, Dim2,Dim3,Dim4, Dim5 : string):string;
begin
CodeArticleUnique:=format('%-18.18s%3.3s%3.3s%3.3s%3.3s%3.3sX',[CodeArticle,Dim1,Dim2,Dim3,Dim4,Dim5]);
end;

function CodeArticleUnique2 ( codeArticle, Dimensions : string):string;
var iDim : integer;
begin
CodeArticle:=format('%-18.18s',[CodeArticle]);
for iDim := 1 to MaxDimension do
    CodeArticle:=CodeArticle+format('%3.3s',[ReadTokenSt(Dimensions)]);
iDim := 18 + (MaxDimension * 3);
CodeArticleUnique2:=format('%-'+IntToStr(iDim)+'.'+IntToStr(iDim)+'sX',[CodeArticle]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 15/02/2000
Modifié le ... :   /  /
Description .. : Verifie si le code article à la structure d'un code article unique
Mots clefs ... : ARTICLE;UNIQUE;CODE;GENERIQUE;DIMENSION
*****************************************************************}
function IsCodeArticleUnique ( codeArticle: string): boolean;
begin
result:=True ;
if length(CodeArticle) < 34 then result:= false
    else if copy(CodeArticle,34,1)<>'X' then result:= false ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 19/01/2000
Modifié le ... : 02/02/2000
Description .. : retourne le code generique de l'article et les codes dimensions à partir du code unique
Mots clefs ... : CODE;ARTICLE;GENERIQUE;UNIQUE;DIMENSION
*****************************************************************}
function CodeArticleGenerique ( var Article, Dim1, Dim2,Dim3,Dim4, Dim5 : string):string;
begin
CodeArticleGenerique:=copy(Article,1,18);
Dim1:=copy(Article,19,3);
Dim2:=copy(Article,22,3);
Dim3:=copy(Article,25,3);
Dim4:=copy(Article,28,3);
Dim5:=copy(Article,31,3);
end;

function CodeArticleGenerique2 ( var Article, Dimensions : string):string;
var
    iDim : integer;
begin
CodeArticleGenerique2:=Trim(copy(Article,1,18));
Dimensions := '';
for iDim := 1 to MaxDimension do
    begin
    if iDim > 1 then Dimensions := Dimensions + ';';
    Dimensions := Dimensions + copy(Article,19 + ((iDim - 1) * 3),3);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 11/02/2000
Modifié le ... : 11/02/2000
Description .. : Cet Article utilise-t-il l'interface affaire ou mode de saisie d'articles
Mots clefs ... : ARTICLE;AFFAIRE;MODE
*****************************************************************}
function IsArticleSpecif (TypeArticle : string) : String;
Begin
result := 'Gescom';
{$IFDEF GESCOM}
if (TypeArticle='FI') then Result:='FicheGescomFi' ;
{$ENDIF}
{$IFDEF GPAO} { GPAO1 }
Result := 'FicheGpao';
{$ENDIF}
If not((ctxAffaire in V_PGI.PGIContexte) or (ctxMode in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte))  then Exit
Else
    Begin
{$IFDEF BTP}
   if (TypeArticle = 'MAR') or (TypeArticle = 'NOM') or (TypeArticle = 'PRE') or
      (TypeArticle = 'OUV') or (TypeArticle = 'POU') or (TypeArticle = 'FRA') or
      (TypeArticle ='ARP') or (copy(typearticle,1,2)='PA') then
       begin
       Result := 'FicheBtp';
       exit;
       end;
{$ENDIF}
    // Pour scot tous les articles sont de type affaire
    If (ctxScot in V_PGI.PGIContexte) or (ctxAffaire in V_PGI.PGIContexte) then  Begin Result:= 'FicheAffaire'; Exit; End;
    If ctxGCAFF in V_PGI.PGIContexte then    //mcd 26/09/03 ajout POU
       if (TypeArticle = 'PRE') Or (TypeArticle = 'FRA') Or (TypeArticle = 'POU') Or (TypeArticle = 'CTR') Then Result := 'FicheAffaire';
    If ctxMode in V_PGI.PGIContexte then
        Begin
        if (TypeArticle='PRE') then Result:='FicheModePre'
        else if (TypeArticle='FI') then Result:='FicheModeFi'
        else Result:='FicheModeArt' ;
        End;
    End;
End;


{***********************Recherche d'article ******************************************}

procedure GetArticleMul_Disp (G_CodeArticle: THCritMaskEdit; stWhere, stRange, stMul : string);
Var CodeArt, Domaine : string;
BEGIN
CodeArt := '';
if ctxAffaire in V_PGI.PGIContexte then
    Begin
{$IFDEF BTP}
    Domaine := 'BTP';
{$ELSE}
    Domaine := 'AFF';
{$ENDIF}
    end
else
begin
{$IFDEF GPAO}
    Domaine := NatureMulArticle;
{$ELSE}
    Domaine := 'GC';  // PA le 5/06/2000
{$ENDIF}
end;
if stMul <> '' then
    BEGIN
    if ctxMode in V_PGI.PGIContexte then CodeArt:=DispatchArtMode(1,stWhere,'',stRange)
    else CodeArt := AGLLanceFiche (Domaine, stMul, stWhere, '', stRange) ;
    END
else CodeArt := AGLLanceFiche ('GC', 'GCARTICLE_RECH', stWhere, '', stRange);
if CodeArt <> '' then // Codeart pour laisser la valeur initiale si pas de selection
    BEGIN
    G_CodeArticle.Text := CodeArt;
    END ;
END ;

procedure GetArticleLookUp_Disp (G_CodeArticle: THCritMaskEdit; stWhere : string; bCodeArticle : Boolean = False);
Var G_article : THCritMaskEdit;
BEGIN
G_Article := G_CodeArticle;
if stWhere <> '' then stWhere := 'GA_STATUTART <> "DIM" AND' + stWhere
else stWhere := 'GA_STATUTART <> "DIM"';
if bCodeArticle then
   LookupList (G_Article, 'Articles', 'ARTICLE', 'GA_CODEARTICLE', 'GA_LIBELLE', stWhere,'GA_CODEARTICLE', True, 7)
else
   LookupList (G_Article, 'Articles', 'ARTICLE', 'GA_ARTICLE', 'GA_LIBELLE', stWhere,'GA_ARTICLE', True, 7) ;
if (G_Article.Text <> '') and (G_Article.Text <> G_CodeArticle.Text) then G_CodeArticle.Text := G_Article.Text;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 27/03/2000
Modifié le ... :   /  /
Description .. : Rercherche d'un Article
Mots clefs ... : RECHERCHE;ARTICLE
*****************************************************************}

Function GetArticleRecherche_Disp (G_CodeArticle: THCritMaskEdit;
                              stWhere, stRange, stMul : string; bCodeArticle : Boolean = False) : string ;
var G_Article : THCritMaskEdit;
BEGIN
Result:='' ;
G_Article:=Nil;
if G_CodeArticle = nil then
    begin
    G_Article := THCritMaskEdit.create (nil);
    G_Article.Text := '';
    end;
if (GetParamSocSecur('SO_GCRECHARTAV', False)) or (G_CodeArticle = nil) then
    BEGIN
    if bCodeArticle then stRange := stRange + ';RETOUR_CODEARTICLE=X';
    if G_CodeArticle = nil then
        begin
        GetArticleMul_Disp (G_Article, stWhere, stRange, stMul);
        Result := G_Article.Text;
        end else
        begin
        GetArticleMul_Disp (G_CodeArticle, stWhere, stRange, stMul);
        Result := G_CodeArticle.Text;
        end;
    END
    else
    BEGIN
    GetArticleLookUp_Disp (G_CodeArticle, stWhere,bCodeArticle);
    Result := G_CodeArticle.Text;
    END;
if G_CodeArticle = nil then G_Article.Free;
END ;

Function GetRechArticle (G_CodeArticle: THCritMaskEdit;
                         stWhere, stRange, stMul : string ) : string ;
var G_Article : THCritMaskEdit;
BEGIN
Result:='' ;
G_Article :=Nil;
if G_CodeArticle = nil then
    begin
    G_Article := THCritMaskEdit.create (nil);
    G_Article.Text := '';
    end;
if (GetParamSocSecur('SO_GCRECHARTAV', False)) or (G_CodeArticle = nil) then
    BEGIN
    if G_CodeArticle = nil then
        begin
        GetArticleMul_Disp (G_Article, stWhere, stRange, stMul);
        Result := G_Article.Text;
        end else
        begin
        GetArticleMul_Disp (G_CodeArticle, stWhere, stRange, stMul);
        Result := G_CodeArticle.Text;
        end;
    END
    else
    BEGIN
    LookUpCombo (G_CodeArticle);
    Result := G_CodeArticle.Text;
    END;
if G_CodeArticle = nil then G_Article.Free;
END ;

(***********A.G.L.***********************************************
Auteur  ...... : Didier Carret
Créé le ...... : 21/08/2000
Modifié le ... : 17/09/2002 Gestion PMAP, PMRP
Modifié le ... : 19/09/2003 PVHT,PVTTC,PAHT,PRHT modifibles si prix multiples
Description .. : Détermine si un champ prix de la fiche article est modifiable
Suite ........ : en fonction du paramètre prix unique (True/False).
Suite ........ : Defaut précise la valeur renvoyée si le champ n'est pas un
Suite ........ : prix
Mots clefs ... : CHAMP;PRIX;ARTICLE;DIM
*****************************************************************)
Function PrixModifiable (NomChamp : string; PrixUnique, Defaut : boolean) : boolean ;
begin
// DCA - FQ MODE 10813
if ((nomChamp='GA_PVHT') or (nomChamp='GA_PVTTC') or
    (nomChamp='GA_PAHT') or (nomChamp='GA_PRHT'))
    then result:=Not PrixUnique
    else result:=Defaut ;
end ;

Function TrouverArticle (G_CodeArticle: THCritMaskEdit; ToBArt : TOB) : T_RechArt ;
BEGIN
result:=TrouverArticle (G_CodeArticle.text, ToBArt ) ;
END;

Function TrouverArticle (G_CodeArticle: string; ToBArt : TOB) : T_RechArt ;
var Q    :TQuery;
    Etat : string ;
BEGIN
Result := traAucun;
if G_CodeArticle = '' Then exit;
Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
               'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
               'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_CODEARTICLE="'+
               G_CodeArticle + '" AND GA_STATUTART <> "DIM" ',true,-1,'',true);
if Not Q.EOF then
    BEGIN
    TOBArt.SelectDB('',Q);
    InitChampsSupArticle (TOBART);
    Etat:=TOBArt.GetValue('GA_STATUTART') ;
    if Etat='UNI' then Result:=traOk else
        if Etat = 'GEN' then Result:=traGrille else Result:=traOk ;
    END ;
Ferme(Q) ;
END;

//////////////// Control Article ///////////////////////////
Function  ControlArticle(Article,CodeArticle: String;Dim:Boolean;TypeArticle:string): String ;
var NumError, MessError, sInfo : string ;
TabErreur: array [1..24] of variant ;
i,j: integer ;
Begin
i:=1 ;
j:=1 ;
if Dim then
begin // Article contient l'identifiant article d'un article dimensionné
  //if ExisteSQL('SELECT GL_ARTICLE FROM LIGNE WHERE GL_ARTICLE="'+Article+'"') then
  sInfo := GetColonneSQL ('LIGNE', 'GL_NATUREPIECEG||";"||GL_SOUCHE||";"||TRIM(STR(GL_NUMERO))', 'GL_ARTICLE="' + Article + '"' );
  if sInfo <> '' then
     begin
        NumError:='1' ;
        MessError:=TraduireMemoire('Utilisé dans la pièce de nature ') + ReadTokenSt(sInfo) +
          TraduireMemoire(' souche ') + ReadTokenSt(sInfo) +
          TraduireMemoire(' numéro ') + ReadTokenSt(sInfo) ;
        TabErreur[i]:= NumError+';'+MessError ;
        inc(i) ;
     end ;
  //if ExisteSQL('SELECT GQ_ARTICLE FROM DISPO WHERE GQ_CLOTURE="-" AND GQ_ARTICLE = "'+Article+'" AND GQ_PHYSIQUE>0') then
  sInfo := GetColonneSQL ('DISPO', 'GQ_DEPOT', 'GQ_CLOTURE="-" AND GQ_ARTICLE = "'+Article+'" AND GQ_PHYSIQUE>0' );
  if sInfo <> '' then
     begin
        NumError:='8' ;
        MessError:=TraduireMemoire('Le stock n''est pas nul dans le dépôt ') + sInfo;
        TabErreur[i]:= NumError+';'+MessError ;
        inc(i) ;
     end ;
end else
begin
  // Contrôle existence en passant par la table ARTICLE sur la clé statut + code article.
  // Le champ GL_REFARTSAISIE n'est pas indexé + contient éventuellement le code barre !
  if ExisteSQL('SELECT GL_ARTICLE FROM LIGNE WHERE GL_ARTICLE IN '+
               '(SELECT GA_ARTICLE FROM ARTICLE WHERE GA_STATUTART="DIM" AND GA_CODEARTICLE="'+CodeArticle+'")')
  or ExisteSQL('SELECT GL_ARTICLE FROM LIGNE WHERE GL_ARTICLE IN '+
               '(SELECT GA_ARTICLE FROM ARTICLE WHERE GA_STATUTART="UNI" AND GA_CODEARTICLE="'+CodeArticle+'")') then
     begin
        NumError:='1' ;
        MessError:=TraduireMemoire('Utilisé dans une pièce');
        TabErreur[i]:= NumError+';'+MessError ;
        inc(i) ;
     end ;
  //if ExisteSQL('SELECT GQ_ARTICLE FROM DISPO WHERE GQ_CLOTURE="-" AND GQ_ARTICLE LIKE "'+Copy(Article,1,18)+'%" AND GQ_PHYSIQUE>0') then
  sInfo := GetColonneSQL ('DISPO', 'GQ_DEPOT', 'GQ_CLOTURE="-" AND GQ_ARTICLE LIKE "'+Copy(Article,1,18)+'%" AND GQ_PHYSIQUE>0' );
  if sInfo <> '' then
     begin
        NumError:='8' ;
        MessError:=TraduireMemoire('Le stock n''est pas nul dans le dépôt ') + sInfo;
        TabErreur[i]:= NumError+';'+MessError ;
        inc(i) ;
     end ;
end ;

  //if ExisteSQL('SELECT GIL_ARTICLE FROM LISTEINVLIG WHERE GIL_ARTICLE="'+Article+'"') then
  sInfo := GetColonneSQL ('LISTEINVLIG', 'GIL_CODELISTE', 'GIL_ARTICLE="' + Article + '"' );
  if sInfo <> '' then
   begin
      NumError:='2' ;
      MessError:=TraduireMemoire('Utilisé en inventaire dans la liste ') + sInfo;
      TabErreur[i]:= NumError+';'+MessError ;
      inc(i) ;
   end ;
  //

  if TypeArticle <> 'ARP' then // pour un article prix pose ..normal
  begin
  //if ExisteSQL('SELECT GNL_ARTICLE FROM NOMENLIG WHERE GNL_ARTICLE="'+Article+'"')  then
  sInfo := GetColonneSQL ('NOMENLIG', 'GNL_NOMENCLATURE', 'GNL_ARTICLE="' + Article + '"' );
  if sInfo <> '' then
     begin
        NumError:='3' ;
  {$IFDEF BTP}
        MessError:=TraduireMemoire('Utilisé dans l''ouvrage ') + sInfo;
  {$ELSE}
        MessError:=TraduireMemoire('Utilisé dans la nomenclature ') + sInfo;
  {$ENDIF}
        TabErreur[i]:= NumError+';'+MessError ;
        inc(i) ;
     end ;
  end;

   //if ExisteSQL('SELECT ACT_ARTICLE FROM ACTIVITE WHERE ACT_ARTICLE="'+Article+'"')  then
   sInfo := GetColonneSQL ('ACTIVITE', 'ACT_AFFAIRE', 'ACT_ARTICLE="' + Article + '"' );
   if sInfo <> '' then
   begin
      NumError:='4' ;
      MessError:=TraduireMemoire('Utilisé en activité dans l''affaire ') + sInfo;
      TabErreur[i]:= NumError+';'+MessError ;
      inc(i) ;
   end ;

   //if ExisteSQL('SELECT ARS_ARTICLE FROM RESSOURCE WHERE ARS_ARTICLE="'+Article+'"') then
   sInfo := GetColonneSQL ('RESSOURCE','ARS_RESSOURCE', 'ARS_ARTICLE="' + Article + '"' );
   if sInfo <> '' then
   begin
      NumError:='5' ;
      MessError:=TraduireMemoire('Utilisé dans la ressource ') + sInfo;
      TabErreur[i]:= NumError+';'+ MessError ;
      inc(i) ;
   end ;

  //if ExisteSQL('SELECT ATA_ARTICLE FROM TACHE WHERE ATA_ARTICLE="'+Article+'"')  then
  sInfo := GetColonneSQL ('TACHE', 'ATA_AFFAIRE||";"||TRIM(STR(ATA_NUMEROTACHE))', 'ATA_ARTICLE="' + Article + '"' );
  if sInfo <> '' then
   begin
      NumError:='6' ;
      MessError:=TraduireMemoire('Utilisé dans l''affaire ') + ReadTokenSt(sInfo) +
        TraduireMemoire(' tache numéro ') + ReadTokenSt(sInfo) ;
    TabErreur[i]:= NumError+';'+ MessError ;
      inc(i) ;
   end ;

  // DCA - FQ MODE 10809 - Amélioration message
  //if ExisteSQL('SELECT GAL_ARTICLELIE FROM ARTICLELIE WHERE GAL_ARTICLELIE="'+Article+'"') then
  sInfo := GetColonneSQL ('ARTICLELIE', 'GAL_ARTICLE', 'GAL_ARTICLELIE="' + Article + '"' );
  if sInfo <> '' then
   begin
      NumError:='7' ;
      MessError:=TraduireMemoire('Utilisé comme article lié de l''article ') + sInfo;
    TabErreur[i]:= NumError+';'+ MessError ;
      inc(i) ;
   end ;

  if ExisteSQL('SELECT BLO_ARTICLE FROM LIGNEOUV WHERE BLO_ARTICLE="'+Article+'"')  then
   begin
  //sInfo := GetColonneSQL ('LIGNEOUV', 'BLO_NATUREPIECEG||";"||BLO_SOUCHE||";"||TRIM(STR(BLO_NUMERO))', 'BLO_ARTICLE="' + Article + '"' );
  //if sInfo <> '' then
  begin
      NumError:='9' ;
      MessError:=TraduireMemoire('Utilisé dans un sous-détail de l''ouvrage de nature ') + ReadTokenSt(sInfo) +
      TraduireMemoire(' souche ') + ReadTokenSt(sInfo) +
      TraduireMemoire(' numéro ') + ReadTokenSt(sInfo) ;
    TabErreur[i]:= NumError+';'+ MessError ;
      inc(i) ;
   end ;

end;

Result:=TabErreur[j] ;
while j<i-1 do
begin
inc(j) ;
Result:=result+';'+TabErreur[j] ;
end ;
END ;

function NaturePieceGeree (naturePiece : string) : boolean ;
var QQ : TQuery ;
begin
Result:=True ;
if ( ctxFO in V_PGI.PGIContexte ) and  ( naturePiece = NAT_VENTESFFO ) then Result := False
else if naturePiece=NATPIE_ARTICLE then Result:=GetParamSoc('SO_GCPARDIMART')
else if naturePiece=NATPIE_STOCK then Result:=GetParamSoc('SO_GCPARDIMSTO')
else if naturePiece=NAT_PROPSAI then Result:=GetParamSoc('SO_GCPARDIMPSA')
else if naturePiece=NAT_PROPDEP then Result:=GetParamSoc('SO_GCPARDIMPSA')
else if naturePiece=NAT_STOVTE then Result:=GetParamSoc('SO_GCPARDIMSTV')
else if naturePiece=NAT_TOBVIEW then Result:=GetParamSoc('SO_GCPARDIMSTV')
else if naturePiece=NAT_ARRETESTO then Result:=False
else // Paramétrage défini dans le PARPIECE
    BEGIN
    QQ:=OpenSql('select GPP_PARAMDIM from PARPIECE where GPP_NATUREPIECEG="'+naturePiece+'"',True,-1,'',true) ;
    if not QQ.eof then Result:=(QQ.FindField('GPP_PARAMDIM').AsString='X') ;
    Ferme(QQ) ;
    END ;
end ;

// DCA - FQ MODE 11033 - ParCombien Obsolète -> DecPrix et DecQte
function nbDecimales(NomChamp,Typ,DimDecQte,DimDecPrix : string) : string ;
begin

  if (Typ='DOUBLE') or (Typ='RATE') then
    Result := DimDecPrix
  else if (Typ='INTEGER') or (Typ='SMALLINT') then
    Result := DimDecQte
  else
    Result := '';

  if (NomChamp='GQ_STOCKMIN')     or (NomChamp='GQ_STOCKMAX')       or (NomChamp='GQ_PHYSIQUE')     or
   (NomChamp='GQ_RESERVECLI') or (NomChamp='GQ_RESERVEFOU') or (NomChamp='GQ_PREPACLI') or
   (NomChamp='GQ_LIVRECLIENT') or (NomChamp='GQ_LIVREFOU') or (NomChamp='GQ_TRANSFERT') or
   (NomChamp='GQ_STOCKINITIAL') or (NomChamp='GQ_CUMULSORTIES') or (NomChamp='GQ_CUMULENTREES') or
     (NomChamp='GQ_STOCKINV')     or
   (NomChamp='GQ_VENTEFFO') or (NomChamp='GQ_ENTREESORTIES') or (NomChamp='GQ_ECARTINV') then
    BEGIN
    //if GereParCombien<1.0 then nbdec:=2 else nbdec:=0 ;
    Result := DimDecQte ;
    END ;

end ;
//

function  PrepareValeurFiltreCombo : string;
var SepOR : string;
begin
  Result := '';
  SepOR := '';
  if not((ctxAffaire in V_PGI.PGIContexte) and
         (ctxScot in V_PGI.PGIContexte))  then
  begin
    Result := Result + SepOR + '(CO_LIBRE like "%GC%")';
    SepOR := ' or ';
  end;
  if ctxScot in V_PGI.PGIContexte then
  begin
    Result := Result + SepOR + '(CO_LIBRE like "%GI%")';
    SepOR := ' or ';
  end;
      //mcd 04/10/04 ctxaffaire existe aussi pour la GI ...if ctxAffaire in V_PGI.PGIContexte then
  if not (ctxScot in V_PGI.PGIContexte) and (ctxaffaire in V_PGI.PGIContexte) then
  begin
    Result := Result + SepOR + '(CO_LIBRE like "%GA%")';
    SepOR := ' or ';
  end;
  if ctxGcAff in V_PGI.PGIContexte then
  begin
    Result := Result + SepOR + '(CO_LIBRE like "%GA%")';
    SepOR := ' or ';
  end;
  if ctxGPAO in V_PGI.PGIContexte then
  begin
    Result := Result + SepOR + '(CO_LIBRE like "%GP%")';
    SepOR := ' or ';
  end;
  if ctxBTP in V_PGI.PGIContexte then
  begin
    Result := Result + SepOR + '(CO_LIBRE like "%BTP%")';
    SepOR := ' or ';
  end;
end;

procedure FiltreComboTypeArticle(TYPEARTICLE : THValComboBox); overload;
var Valeur : string;
begin
  Valeur := PrepareValeurFiltreCombo;
  if Valeur <> '' then
  begin
    Valeur := ' and (' + Valeur + ')';
    TYPEARTICLE.Plus := TYPEARTICLE.Plus + Valeur;
  end;
end;

procedure FiltreComboTypeArticle(TYPEARTICLE : THMultiValComboBox); overload
var Valeur : string;
begin
  Valeur := PrepareValeurFiltreCombo;
  if Valeur <> '' then
  begin
    Valeur := ' and (' + Valeur + ')';
    TYPEARTICLE.Plus := TYPEARTICLE.Plus + Valeur;
  end;
end;
//

function reinit_combo(cb : THValComboBox) : boolean ;
begin
if (cb.text<>'') and (cb.text<>'<<Tous>>') then
    BEGIN
    cb.value:='' ;
    result:=True ;
    END
    else result:=False ;
end ;

function GetPresentation : integer ;
begin
if (V_PGI.LaSerie>S3) and GetParamSoc('SO_ARTLOOKORLI') then Result:=ART_ORLI
else Result:=ART_S3 ;
end ;

// Sélection mul article Mode
Function DispatchArtMode (Num : integer; Range,Lequel,Argument : String) : variant ;
BEGIN
Case Num of
  1 : if GetPresentation=ART_ORLI  // Mul article mode
          then Result:=AGLLanceFiche('MBO','ARTICLE_MODES5',Range,Lequel,Argument)
          else Result:=AGLLanceFiche('MBO','ARTICLE_MODE',Range,Lequel,Argument) ;
  2 : if GetPresentation=ART_ORLI  // Etiquettes articles sur stock
          then Result:=AGLLanceFiche('GC','GCETIARTSTK_MODS5',Range,Lequel,Argument)
          else Result:=AGLLanceFiche('GC','GCETIARTSTK_MODE',Range,Lequel,Argument) ;
  3 : if GetPresentation=ART_ORLI  // Stock disponible article
          then Result:=AGLLanceFiche('MBO','DISPODIMS5',Range,Lequel,Argument)
          else Result:=AGLLanceFiche('MBO','DISPODIM',Range,Lequel,Argument) ;
  4 : if GetPresentation=ART_ORLI  // Tarif détail - Maj tarifs mode
          then AGLLanceFiche('MBO','TARIFMAJ_MULS5',Range,Lequel,Argument)
          else AGLLanceFiche('MBO','TARIFMAJ_MUL',Range,Lequel,Argument) ;
  5 : if GetPresentation=ART_ORLI  // Edition - Etat du stock par taille
          then AGLLanceFiche('MBO','ARTDISPODIMS5',Range,Lequel,Argument)
          else AGLLanceFiche('MBO','ARTDISPODIM',Range,Lequel,Argument) ;
  6 : if GetPresentation=ART_ORLI  // Article - Liste des articles
          then AGLLanceFiche('MBO','EDTART_MODES5',Range,Lequel,Argument)
          else AGLLanceFiche('MBO','EDTART_MODE',Range,Lequel,Argument) ;
  7 : if GetPresentation=ART_ORLI  // Tarifs détail - Edition - tarifs + comparatif
          then AGLLanceFiche('GC','GCETATTARIFTTCS5',Range,Lequel,Argument)
          else AGLLanceFiche('GC','GCETATTARIFTTC',Range,Lequel,Argument) ;
{  8 : if GetPresentation=ART_ORLI  // Stocks - Consultation - Statistique stock disponible
          then AGLLanceFiche('MBO','DISPO_VUES5',Range,Lequel,Argument)
          else AGLLanceFiche('MBO','DISPO_VUE',Range,Lequel,Argument) ;
  9 : if GetPresentation=ART_ORLI  // Stocks - Consultation - Cube des stocks
          then AGLLanceFiche('GC','GCDISPO_CUBES5',Range,Lequel,Argument)
          else AGLLanceFiche('GC','GCDISPO_CUBE',Range,Lequel,Argument) ;  }
 10 : if GetPresentation=ART_ORLI  // Données de base - Articles - Prestations
          then AGLLanceFiche('MBO','MBOPRESTS5_MUL',Range,Lequel,Argument)
          else AGLLanceFiche('MBO','MBOPREST_MUL',Range,Lequel,Argument) ;
 11 : if GetPresentation=ART_ORLI  // Mul Prestations - DbleClick
          then AGLLanceFiche('MBO','MBOPRESTATIONS5',Range,Lequel,Argument)
          else AGLLanceFiche('MBO','MBOPRESTATION',Range,Lequel,Argument) ;
{ 12 : if GetPresentation=ART_ORLI  // OT : Cube à la dimension
          then AGLLanceFiche('MBO','CUBEDIMS5',Range,Lequel,Argument)
          else AGLLanceFiche('MBO','CUBEDIM',Range,Lequel,Argument) ;     }
 13 : if GetPresentation=ART_ORLI  // OT : Tableaux de bord du disponible article avec statistiques sur les ventes et les achats
          then AGLLanceFiche('MBO','STATDIMS5',Range,Lequel,Argument)
          else AGLLanceFiche('MBO','STATDIM',Range,Lequel,Argument) ;
  else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
  END ;
END ;

// Modif BTP
function RecupTypeGraph (TOBExamin : Tob) : Integer;
var Provenance,NiveauImbric,IndiceNomen,TypePresent : integer;
    TypeArticle,TypeLigne : String;
begin

Result := -1;

if TOBExamin.FieldExists ('GL_ARTICLE') Then Provenance := 0 // Ligne document
else if TOBExamin.FieldExists ('BLO_ARTICLE') Then Provenance := 1 // Ouvrage
else if TOBExamin.fieldexists ('GA_ARTICLE') then Provenance := 4 // Article
else exit;  // provenance Non Gérées

TypePresent := DOU_AUCUN ;
IndiceNomen := 0 ;
 NiveauImbric := 0 ;
if Provenance = 0 then
   begin
   TypeArticle := TOBExamin.getValue('GL_TYPEARTICLE');
   TypePresent := TOBExamin.getValue('GL_TYPEPRESENT');
   IndiceNomen := TOBExamin.getValue('GL_INDICENOMEN');
   TypeLigne := TOBExamin.getValue('GL_TYPELIGNE');
   NiveauImbric := TOBExamin.getValue('GL_NIVEAUIMBRIC');
   end else if Provenance = 1 then
   begin
   TypeArticle := TOBExamin.getValue('BLO_TYPEARTICLE');
   TypeLigne := '';
   end else if Provenance = 4 then
   begin
   TypeArticle := TOBExamin.getValue('GA_TYPEARTICLE');
   TypeLigne := '';
   end;

if TypeArticle = 'MAR' Then Result := 7
else if (TypeArticle = 'OUV') and (TypePresent > DOU_AUCUN) then Result := 0
else if (TypeArticle = 'OUV') and (TypePresent = DOU_AUCUN) then Result := 1
else if (TypeArticle = 'POU') then Result := 6
else if (TypeArticle = 'COM') and (IndiceNomen = 0) then Result := 2
// VARIANTE
(*else if (copy(TypeLigne,1,2)='DP') and (NiveauImbric > 1) then Result := 4
else if (copy(TypeLigne,1,2)='DP') and (NiveauImbric = 1 ) then Result := 3 *)
else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='DV')) and (NiveauImbric > 1) then Result := 4
else if ((copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='DV')) and (NiveauImbric = 1 ) then Result := 3
// --
else if TypeArticle ='PRE' then Result := 5
else if TypeArticle ='FRA' then Result := 8
else if TypeArticle = 'ARP' then Result := 9
else Result := -1;
end;

// Modif BTP
function ArticleOKInPOUR (TypeArt,Chaine : string) : boolean;
var chainecmp ,EltCur : string;
begin
if Chaine='' then
   begin
   result := true;
   exit;
   end;
ChaineCmp := Chaine;
result := false;
repeat
EltCur := TRIM(readtokenSt(chaineCmp));
if EltCur <> '' then
   begin
   if TypeArt = Eltcur then
      begin
      result:=true;
      break;
      end;
   end;
until Eltcur = ''
end;


procedure chargeTobPrestation (Prestation : String;TOBPrestation:TOB);
var REq : string;
    QQ : TQuery;
begin
if TOBPrestation = nil then exit;
TOBPrestation.InitValeurs;
if prestation = '' then exit;
Req := 'SELECT GA_ARTICLE,GA_CODEARTICLE,GA_LIBELLE,GA_QUALIFUNITEVTE,'
       + 'GA_PAHT,GA_DPR,GA_PVHT,GA_PVTTC,GA_PRIXPOURQTE FROM ARTICLE '
       + 'WHERE GA_ARTICLE="'+Prestation+'"';
QQ := OpenSql (Req,true,-1,'',true);
TOBPrestation.SelectDB ('',QQ);
ferme (QQ);
end;

procedure GetDetailPrixPose(CodeArticle : string;TOBPrestation,TOBEDetail,TOBDetail: TOB);
var Req : string;
    QQ : Tquery;
    Prestation : string;
begin
Req := 'SELECT * FROM NOMENENT WHERE GNE_NOMENCLATURE="'+CodeArticle+'"';
QQ := OpenSql (Req,true,-1,'',true);
if not QQ.eof then TOBEDetail.SelectDB ('',QQ);
ferme (QQ);
Req := 'SELECT * FROM NOMENLIG WHERE GNL_NOMENCLATURE="'+CodeArticle+'"';
QQ := OpenSql (Req,true,-1,'',true);
if not QQ.eof then TOBDetail.LoadDetailDB ('NOMENLIG','','',QQ,false);
ferme (QQ);
if TobDetail.detail.count = 2 then Prestation := TOBDetail.detail[1].getValue('GNL_ARTICLE');
chargeTobPrestation (Prestation,TOBPrestation);
end;

{$IFDEF BTP}
function GetTypeRessourceFromNature (Naturepres : string) : string;
var Req : string;
    QQ : Tquery;
begin
result := '';
REq := 'SELECT BNP_TYPERESSOURCE FROM NATUREPREST WHERE BNP_NATUREPRES="'+Naturepres+'"';
QQ := OpenSql (Req,true,-1,'',true);
if not QQ.eof Then Result := QQ.findfield('BNP_TYPERESSOURCE').AsString ;
ferme (QQ);
end;

// fonction retournant le type de l'article dans la biblio
Function  RenvoieTypeArt (CodArt : string) : String ;
Var Q : TQuery ;
    Req, TypArt : String;
BEGIN

Req:='SELECT GA_TYPEARTICLE FROM ARTICLE WHERE GA_ARTICLE="'+CodArt+'"' ;
Q:=OpenSQL(Req,TRUE,-1,'',true) ;
if Not Q.EOF then TypArt:=Q.Fields[0].AsString
else TypArt := '';
Ferme(Q) ;

result := TypArt;

end;

function GetTypeArticleBTP : string;
begin
result := '(GA_TYPEARTICLE="ARP" OR GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="PRE" OR GA_TYPEARTICLE="OUV" OR GA_TYPEARTICLE="POU")';
end;

function GetTypeArticlePARC : string;
begin
result := '(GA_TYPEARTICLE="PA1" OR GA_TYPEARTICLE="PA2")';
end;

procedure RecalculPrPV (TOBRef,TOBCatalog : TOB;VenteAchat : string='VEN');
begin
RecalculPr (TOBRef,TobCatalog,VenteAchat);
RecalculPvHt (TOBRef,TobCatalog,VenteAchat);
RecalculPvTTC (TOBRef,TobCatalog,VenteAchat);
end;

procedure RecalculPr (TOBRef,TOBCatalog : TOB;VenteAchat : string='VEN');
var CoefcalcPR       : Double;
    PrixRef          : Double;
    QuotiteAchat     : Double;
    QuotiteVente     : double;
    TarifFournisseur : Double;
    CodeBasePR       : string;
    QQUALIFACH       : TQuery;
    QQUALIFVTE       : TQuery;
begin

  if TOBRef.GetValue('GA_DPRAUTO')<>'X' then exit;

  if VenteAchat = 'VEN' Then
  begin
  CoefCalcPR:=TOBRef.GetValue('GA_COEFFG');
  end else
  begin
  CoefCalcPR:=1;
  end;
  CodeBasePR:=TOBRef.GetValue('GA_CALCPRIXPR') ;

  if (CoefCalcPR<>0) and (CodeBasePR<>'') then
  BEGIN
    if CodeBasePR='DPA' then PrixRef:=TOBRef.GetValue('GA_DPA')
    else if CodeBasePR='PAA' then PrixRef:=TobRef.GetValue('GA_PAHT')
    else if CodeBasePR='PRA' then PrixRef:=TobRef.GetValue('GA_PRHT')
    else if CodeBasePR='PMA' then PrixRef:=TobRef.GetValue('GA_PMAP')
    else if CodeBasePR='FOU' then
    Begin
      if tobcatalog <> nil then
      begin
        QQUALIFACH:=OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="'+TobCatalog.GetValue('GCA_QUALIFUNITEACH')+'"',True,-1,'',true) ;
        QQUALIFVTE:=OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="'+TobRef.GetValue('GA_QUALIFUNITEVTE')+'"',True,-1,'',true) ;
        If not QQUALIFACH.Eof then QuotiteAchat:=Valeur(QQUALIFACH.FindField('GME_QUOTITE').AsString) else QuotiteAchat:=1;
        If not QQUALIFVTE.Eof then QuotiteVente:=Valeur(QQUALIFVTE.FindField('GME_QUOTITE').AsString) else QuotiteVente:=1;
        //TarifFournisseur:=TobRef.GetValue('GCA_DPA');
        //Modif FV du 31/10/2006
        TarifFournisseur:=tobcatalog.GetValue('GCA_DPA');
        if QuotiteAchat <> 0 then
          PrixRef:=((TarifFournisseur * QuotiteVente)/QuotiteAchat)*TobRef.GetValue('GA_PRIXPOURQTE');
        Ferme(QQUALIFACH) ;
        Ferme(QQUALIFVTE) ;
      end
      else PrixRef := TobRef.GetValue('GA_PAHT');
    End;
    TOBRef.putValue('GA_DPR',arrondi(PrixRef*CoefCalcPR,V_PGI.okDEcP));
  END;

end;

procedure RecalculPvHt (TOBRef,TOBCatalog : TOB;VenteAchat : string='VEN');
var CoefcalcPV,PrixRef,TarifFournisseur,QuotiteAchat,QuotiteVente : double;
    CodeBasePV : string;
    QQUALIFACH,QQUALIFVTE : TQuery;
begin

  if (TOBRef.GetValue('GA_CALCAUTOHT')<>'X') then exit;

  if VenteAchat = 'VEN' Then
    CoefCalcPV:=TOBRef.GetValue('GA_COEFCALCHT')
  else
    CoefCalcPV:=1;

  CodeBasePV:=TOBRef.GetValue('GA_CALCPRIXHT') ;

  if (CoefCalcPV<>0) and (CodeBasePV<>'') then
  BEGIN
    if CodeBasePV='DPA' then PrixRef:=TOBRef.GetValue('GA_DPA')
    else if CodeBasePV='DPR' then PrixRef:=TOBRef.GetValue('GA_DPR')
    else if CodeBasePV='PAA' then PrixRef:=TobRef.GetValue('GA_PAHT')
    else if CodeBasePV='PRA' then PrixRef:=TobRef.GetValue('GA_PRHT')
    else if CodeBasePV='PMA' then PrixRef:=TobRef.GetValue('GA_PMAP')
    else if CodeBasePV='PMR' then PrixRef:=TobRef.GetValue('GA_PMRP')
    else if CodeBasePV='FOU' then
    Begin
      if tobcatalog <> nil then
      begin
        QQUALIFACH:=OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="'+TobCatalog.GetValue('GCA_QUALIFUNITEACH')+'"',True,-1,'',true) ;
        QQUALIFVTE:=OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="'+TobRef.GetValue('GA_QUALIFUNITEVTE')+'"',True,-1,'',true) ;
        If not QQUALIFACH.Eof then QuotiteAchat:=Valeur(QQUALIFACH.FindField('GME_QUOTITE').AsString) else QuotiteAchat:=1;
        If not QQUALIFVTE.Eof then QuotiteVente:=Valeur(QQUALIFVTE.FindField('GME_QUOTITE').AsString) else QuotiteVente:=1;
        TarifFournisseur:=tobcatalog.GetValue('GCA_DPA');
        PrixRef:=((TarifFournisseur * QuotiteVente/QuotiteAchat)*TobRef.GetValue('GA_PRIXPOURQTE')) ;
        Ferme(QQUALIFACH) ;
        Ferme(QQUALIFVTE) ;
      end
      else
        PrixRef := TobRef.GetValue('GA_DPA');
    End;
    TOBRef.putValue('GA_PVHT',arrondi(PrixRef*CoefCalcPV,V_PGI.okDEcP));
  END;

end;


procedure RecalculPvTTC (TOBRef,TOBCatalog : TOB;VenteAchat : string='VEN');
var CoefcalcPV,PrixRef,TarifFournisseur,QuotiteAchat,QuotiteVente : double;
    CodeBasePV : string;
    QQUALIFACH,QQUALIFVTE : TQuery;
begin
if TOBRef.GetValue('GA_CALCAUTOTTC')<>'X' then exit;

if VenteAchat = 'VEN' Then
begin
CoefCalcPV:=TOBRef.GetValue('GA_COEFCALCTTC');
end else
begin
CoefCalcPV:=1;
end;
CodeBasePV:=TOBRef.GetValue('GA_CALCPRIXTTC') ;

if (CoefCalcPV<>0) and (CodeBasePV<>'') then
   BEGIN
   if CodeBasePV='DPA' then PrixRef:=TOBRef.GetValue('GA_DPA')
   else if CodeBasePV='DPR' then PrixRef:=TOBRef.GetValue('GA_DPR')
   else if CodeBasePV='PAA' then PrixRef:=TobRef.GetValue('GA_PAHT')
   else if CodeBasePV='PRA' then PrixRef:=TobRef.GetValue('GA_PRHT')
   else if CodeBasePV='PMA' then PrixRef:=TobRef.GetValue('GA_PMAP')
   else if CodeBasePV='PMR' then PrixRef:=TobRef.GetValue('GA_PMRP')
   else if CodeBasePV='HT' then PrixRef:=TobRef.GetValue('GA_PVHT')
   else if CodeBasePV='FOU' then
      Begin
      if TobCatalog <> nil then
         begin
         QQUALIFACH:=OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="'+TobCatalog.GetValue('GCA_QUALIFUNITEACH')+'"',True,-1,'',true) ;
         QQUALIFVTE:=OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="'+TobRef.GetValue('GA_QUALIFUNITEVTE')+'"',True,-1,'',true) ;
         If not QQUALIFACH.Eof then QuotiteAchat:=Valeur(QQUALIFACH.FindField('GME_QUOTITE').AsString) else QuotiteAchat:=1;
         If not QQUALIFVTE.Eof then QuotiteVente:=Valeur(QQUALIFVTE.FindField('GME_QUOTITE').AsString) else QuotiteVente:=1;
         TarifFournisseur:=tobcatalog.GetValue('GCA_DPA');
         PrixRef:=((TarifFournisseur * QuotiteVente/QuotiteAchat)*TobRef.GetValue('GA_PRIXPOURQTE')) ;
         Ferme(QQUALIFACH) ;
         Ferme(QQUALIFVTE) ;
         end else PrixRef := TobRef.GetValue('GA_DPA');
      End;
  TOBRef.putValue('GA_PVTTC',arrondi(PrixRef*CoefCalcPV,V_PGI.okDEcP));
  END;
end;

{$ENDIF}

Function ChangeCodificationTarif (CodeTarif,NewCode,TraiteCodeArt : String) : Boolean ;
var ChangeCodifArt : TChangeCodeArt;
    Req, NomChoixCod, AncCodArt, NewCodArt : String;
    Nb : Integer;
    Q : TQuery ;
begin
  Result := True;
  // Récupération nom de base principale en cas de partage
  NomChoixcod := 'CHOIXCOD';
  Q := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="GCTARIFARTICLE"', True);
  if not Q.EOF then
  begin
  	NomChoixCod := Q.FindField('DS_NOMBASE').AsString+'.DBO.CHOIXCOD';
  end;
	Ferme (Q);

  // Contrôle existence nouveau code tarif avant traitement dans tablette tarif article (GCTARIFARTICLE)
  //Req := 'SELECT CC_CODE FROM '+NomChoixCod+' WHERE CC_TYPE="TAR" AND CC_CODE="'+NewCode+'"';
  //FV1 : 21/01/2016 - FS#1854 - En import catalogue, les tables BTFAMILLETARIF et BTSOUSFAMILLETARIF ne sont pas alimentées
  //Je vois pas comment cette fonction peut marcher on pas les bonnes zones ni les bonnes tables (???)
  //Req := 'SELECT CC_CODE FROM FAMILLETARIF WHERE BFT_FAMILLETARIF="'+NewCode+'"';
  Req := 'SELECT BFT_FAMILLETARIF FROM BTFAMILLETARIF WHERE BFT_FAMILLETARIF="'+ NewCode + '"';
  if Not ExisteSql(Req) then
  begin
    // Maj Code tarif dans tablette tarif article (GCTARIFARTICLE)
    //Req := 'UPDATE '+NomChoixCod+' SET CC_CODE="'+ NewCode +'" WHERE CC_TYPE="TAR" AND CC_CODE="'+CodeTarif+'"';
    //Req := 'UPDATE FAMILLETARIF SET BFT_FAMILLETARIF="'+ NewCode +'" WHERE BFT_FAMILLETARIF="'+CodeTarif+'"';
    Req := 'UPDATE BTFAMILLETARIF SET BFT_FAMILLETARIF="'+ NewCode +'" WHERE BFT_FAMILLETARIF="'+CodeTarif+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN Result:=false ; Exit ; END ;
    AvertirTable ('GCTARIFARTICLE'); // rechargement tablette
  end ;

  // Maj Code tarif dans tablette sous-famille tarif article (BTSOUSFAMTARART)
  //Req := 'UPDATE '+NomChoixCod+' SET CC_LIBRE="'+ NewCode +'" WHERE CC_TYPE="BFT" AND CC_LIBRE="'+CodeTarif+'"';
  Req := 'UPDATE BTSOUSFAMILLETARIF SET BSF_FAMILLETARIF="'+ NewCode +'" WHERE BSF_FAMILLETARIF="'+CodeTarif+'"';
  Nb := ExecuteSql (Req);
  if Nb < 0 then BEGIN Result:=false ; Exit ; END ;
  AvertirTable ('BTSOUSFAMTARART'); // rechargement tablette

  // Maj Code tarif dans table TARIF
  Req := 'UPDATE TARIF SET GF_TARIFARTICLE="'+ NewCode +'" WHERE GF_TARIFARTICLE="'+CodeTarif+'"';
  Nb := ExecuteSql (Req);
  if Nb < 0 then BEGIN Result:=false ; Exit ; END ;

  // Maj Code tarif dans table ARTICLE
  Req := 'UPDATE ARTICLE SET GA_TARIFARTICLE="'+ NewCode +'" WHERE GA_TARIFARTICLE="'+CodeTarif+'"';
  Nb := ExecuteSql (Req);
  if Nb < 0 then BEGIN Result:=false ; Exit ; END ;

  // Traitement des articles si le code comporte le code tarif comme préfixe :
  if TraiteCodeArt = 'X' then
  begin
    Req:='SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_CODEARTICLE LIKE "'+CodeTarif+'%"' ;
    Q:=OpenSQL(Req,TRUE,-1,'',true) ;
    InitMoveProgressForm(nil,TraduireMemoire('Modification des codes articles en cours ...'), '', Q.RecordCount, false, true);

    while Not Q.EOF do
    begin
      AncCodArt:=Q.Fields[0].AsString;
      NewCodArt:=NewCode+Copy(AncCodArt,4,Length(AncCodArt)-3);

      MoveCurProgressForm(TraduireMemoire('Article : ') + AncCodArt);
      ChangeCodificationArticle(AncCodArt, NewCodArt);

      Q.Next;
    end;

    FiniMoveProgressForm;
    Ferme(Q) ;
  end;
end;

Function CreationIndex (TraiteCodeArt : String) : Boolean;
begin
  result := True;
  if TraiteCodeArt <> 'X' then Exit;

  InitMoveProgressForm(nil,TraduireMemoire('Création des index en cours ...'), '', 8, false, true);
  // CREATION DES INDEX
  Try
    // Table LIGNE
    MoveCurProgressForm(TraduireMemoire('Table LIGNE : Index 1'));
    ExecuteSql ('CREATE INDEX GL_ZIND1 ON LIGNE (GL_ARTICLE)');
    MoveCurProgressForm(TraduireMemoire('Table LIGNE : Index 2'));
    ExecuteSql ('CREATE INDEX GL_ZIND2 ON LIGNE (GL_CODEARTICLE)');
    MoveCurProgressForm(TraduireMemoire('Table LIGNE : Index 3'));
    ExecuteSql ('CREATE INDEX GL_ZIND3 ON LIGNE (GL_REFARTSAISIE)');
    // Table LIGNEOUV
    MoveCurProgressForm(TraduireMemoire('Table LIGNEOUV : Index 1'));
    ExecuteSql ('CREATE INDEX BLO_ZIND1 ON LIGNEOUV (BLO_ARTICLE)');
    MoveCurProgressForm(TraduireMemoire('Table LIGNEOUV : Index 2'));
    ExecuteSql ('CREATE INDEX BLO_ZIND2 ON LIGNEOUV (BLO_CODEARTICLE)');
    MoveCurProgressForm(TraduireMemoire('Table LIGNEOUV : Index 3'));
    ExecuteSql ('CREATE INDEX BLO_ZIND3 ON LIGNEOUV (BLO_COMPOSE)');
    MoveCurProgressForm(TraduireMemoire('Table LIGNEOUV : Index 4'));
    ExecuteSql ('CREATE INDEX BLO_ZIND4 ON LIGNEOUV (BLO_REFARTSAISIE)');
  Except
    PGIBox ('Erreur en création des index. Traitement impossible','Attention');
    result := False;
  End;
  FiniMoveProgressForm;
end;

Procedure ChangementCodeTarif;
Var CodeTarif,NewCode,TraiteCodeArt : string;
    IndexCree : Boolean;
begin
	AGLLanceFiche('BTP','BTVIDEINSIDE','','','') ;
  if SaisieNewCodeTar (CodeTarif,NewCode,TraiteCodeArt) then
  begin
    if (PGIAsk ('Confirmez-vous le changement de codification ?', '')=mrYes) then
    begin
      IndexCree := CreationIndex(TraiteCodeArt);
      if Not IndexCree then Exit;

      BeginTrans;
      if Not ChangeCodificationTarif(CodeTarif,NewCode,TraiteCodeArt) then
      begin
        Rollback;
        PGIBox ('Erreur de mise à jour','Changement code tarif');
      end else
      begin
        CommitTrans;
        PgiBox ('Traitement terminé','Changement code tarif');
      end;

      //SUPPRESSION DES INDEX
      if (IndexCree) and (TraiteCodeArt = 'X') then
      begin
        // Table LIGNE
        ExecuteSql ('DROP INDEX LIGNE.GL_ZIND1');
        ExecuteSql ('DROP INDEX LIGNE.GL_ZIND2');
        ExecuteSql ('DROP INDEX LIGNE.GL_ZIND3');
        // Table LIGNEOUV
        ExecuteSql ('DROP INDEX LIGNEOUV.BLO_ZIND1');
        ExecuteSql ('DROP INDEX LIGNEOUV.BLO_ZIND2');
        ExecuteSql ('DROP INDEX LIGNEOUV.BLO_ZIND3');
        ExecuteSql ('DROP INDEX LIGNEOUV.BLO_ZIND4');
      end;
    end;
  end;
end;

function ChangeCodificationArticle (CodeArticle,NewCode : String) : boolean;
var ChangeCodifArt : TChangeCodeArt;
begin
ChangeCodifArt := TChangeCodeArt.Create;
TRY
  changeCodifArt.CodeArticle := CodeArticle;
  changeCodifArt.newCode := NewCode;
  ChangeCodifArt.TraiteChange;
FINALLY
  ChangeCodifArt.Free;
END;
Result:=True;
end;

{ TChangeCodeArt }

constructor TChangeCodeArt.Create;
begin
creeLesTobs;
end;

destructor TChangeCodeArt.Destroy;
begin
  inherited;
detruitLesTobs;
end;

procedure TChangeCodeArt.creeLesTobs;
begin
TOBArticle := TOB.Create ('ARTICLES',nil,-1); (* Articles prefixe GA *)
end;


procedure TChangeCodeArt.detruitLesTobs;
begin
TOBArticle.free;
end;

procedure TChangeCodeArt.LoadLesArticles;
var QQ : TQuery;
    TOBSup : TOB;
    Indice : integer;
    TOBAdd : TOB;
begin
Indice := 0;
TOBAdd := TOB.Create ('AAA',nil,-1);
QQ := OpenSql ('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="'+CodeArticle+'"',true,-1,'',true);
if not QQ.eof then
   BEGIN
   TOBAdd.LoadDetailDB ('BBB','','',QQ,false,true);
   Repeat
       TOBSUp := TOBADD.detail[Indice];
       TOBSUp.addChampSupValeur('PRECEDENT',TOBSup.getValue('GA_ARTICLE'));
       TOBSup.ChangeParent (TOBArticle,-1);
   until TOBAdd.detail.count = 0;
   TobAdd.clearDetail;
   END;
ferme (QQ);
TobAdd.free;
end;

procedure TChangeCodeArt.LoadLesTobs;
begin
// Articles
LoadLesArticles;
end;

procedure TChangeCodeArt.TraiteChange;
begin
LoadLesTobs;
if transactions (TraitelesTobs,1) <> OeOk then PGIBox ('Erreur de mise à jour','Changement de Code');
end;

procedure TChangeCodeArt.TraiteLesArt;
var Indice,NB : integer;
    TOBArt : TOB;
    AncArt,Req : string;
    Dim1,Dim2,Dim3,DIm4,Dim5 : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];
    AncArt := TOBArt.getValue('GA_ARTICLE');
    AncArt := CodeArticleGenerique (AncArt,Dim1,Dim2,Dim3,Dim4,Dim5);
    TOBArt.PutValue('GA_ARTICLE',CodeArticleUnique(NewCode,Dim1,Dim2,DIm3,DIm4,Dim5));

    Req := 'SELECT GA_ARTICLE FROM ARTICLE WHERE GA_ARTICLE="'+TOBArt.GetValue('GA_ARTICLE')+'"';
    if ExisteSql(Req) then Exit;

    Req := 'UPDATE ARTICLE SET GA_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE GA_ARTICLE="'+TOBART.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
Req := 'UPDATE ARTICLE SET ';
Req := req + 'GA_CODEARTICLE="'+ NewCode+'"';
Req := Req + ' WHERE ';
REq := REq + 'GA_CODEARTICLE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
end;

procedure TChangeCodeArt.TraiteLesComplementsArt;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];

    Req := 'SELECT GA2_ARTICLE FROM ARTICLECOMPL WHERE GA2_ARTICLE="'+TOBArt.GetValue('GA_ARTICLE')+'"';
    if ExisteSql(Req) then Exit;

    Req := 'UPDATE ARTICLECOMPL SET ';
    Req := req + 'GA2_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'GA2_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
Req := 'UPDATE ARTICLECOMPL SET ';
Req := req + 'GA2_CODEARTICLE="'+newCode+'" WHERE ';
REq := REq + 'GA2_CODEARTICLE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
end;

procedure TChangeCodeArt.TraiteLesArtLies;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];

    Req := 'SELECT GAL_ARTICLE FROM ARTICLELIE WHERE GAL_ARTICLE="'+TOBArt.GetValue('GA_ARTICLE')+'"';
    if ExisteSql(Req) then Exit;

    Req := 'UPDATE ARTICLELIE SET ';
    Req := req + 'GAL_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'"';
    Req := Req + ' WHERE ';
    REq := REq + 'GAL_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    //
    Req := 'UPDATE ARTICLELIE SET ';
    Req := req + 'GAL_ARTICLELIE="'+ TOBArt.GetValue('GA_ARTICLE')+'"';
    Req := Req + ' WHERE ';
    REq := REq + 'GAL_ARTICLELIE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
Req := 'UPDATE ARTICLELIE SET ';
Req := req + 'GAL_ARTICLE="'+ NewCode+'"';
Req := Req + ' WHERE ';
REq := REq + 'GAL_ARTICLE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
//
Req := 'UPDATE ARTICLELIE SET ';
Req := req + 'GAL_ARTICLELIE="'+ NewCode +'"';
Req := Req + ' WHERE ';
REq := REq + 'GAL_ARTICLELIE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
end;

procedure TChangeCodeArt.TraiteLesConsommations;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];
    Req := 'UPDATE CONSOMMATIONS SET ';
    Req := req + 'BCO_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'BCO_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
Req := 'UPDATE CONSOMMATIONS SET ';
Req := req + 'BCO_CODEARTICLE="'+newCode+'" WHERE ';
REq := REq + 'BCO_CODEARTICLE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
end;

procedure TChangeCodeArt.TraiteLesArtTiers;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];
    Req := 'UPDATE ARTICLETIERS SET ';
    Req := req + 'GAT_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'"';
    Req := Req + ' WHERE ';
    REq := REq + 'GAT_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
end;

procedure TChangeCodeArt.TraiteLesCatalogues;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];

    Req := 'SELECT GCA_ARTICLE FROM CATALOGU WHERE GCA_ARTICLE="'+TOBArt.GetValue('GA_ARTICLE')+'"';
    if ExisteSql(Req) then Exit;

    Req := 'UPDATE CATALOGU SET ';
    Req := req + 'GCA_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'"';
    Req := Req + ' WHERE ';
    REq := REq + 'GCA_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
end;

procedure TChangeCodeArt.TraiteLesArtPiece;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];

    Req := 'SELECT GAP_ARTICLE FROM ARTICLEPIECE WHERE GAP_ARTICLE="'+TOBArt.GetValue('GA_ARTICLE')+'"';
    if ExisteSql(Req) then Exit;

    Req := 'UPDATE ARTICLEPIECE SET ';
    Req := req + 'GAP_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'"';
    Req := Req + ' WHERE ';
    REq := REq + 'GAP_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
end;

procedure TChangeCodeArt.TraiteLesArtComm;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];
    Req := 'UPDATE COMMISSION SET ';
    Req := req + 'GCM_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'GCM_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
Req := 'UPDATE COMMISSION SET ';
Req := req + 'GCM_CODEARTICLE="'+ newCode+'" WHERE ';
REq := REq + 'GCM_CODEARTICLE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
end;

procedure TChangeCodeArt.TraiteLesConditArt;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];

    Req := 'SELECT GCO_ARTICLE FROM CONDITIONNEMENT WHERE GCO_ARTICLE="'+TOBArt.GetValue('GA_ARTICLE')+'"';
    if ExisteSql(Req) then Exit;

    Req := 'UPDATE CONDITIONNEMENT SET ';
    Req := req + 'GCO_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'"';
    Req := Req + ' WHERE ';
    REq := REq + 'GCO_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
end;

procedure TChangeCodeArt.TraiteLesTradArt;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];

    Req := 'SELECT GTA_ARTICLE FROM GTRADARTICLE WHERE GTA_ARTICLE="'+TOBArt.GetValue('GA_ARTICLE')+'"';
    if ExisteSql(Req) then Exit;

    Req := 'UPDATE GTRADARTICLE SET ';
    Req := req + 'GTA_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'"';
    Req := Req + ' WHERE ';
    REq := REq + 'GTA_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
end;

procedure TChangeCodeArt.TraiteLesNomenEnt;
var Indice,NB : integer;
    TOBArt : TOB;
    Req: string;
begin
  for Indice := 0 to TOBArticle.detail.count -1 do
  begin
    TOBArt := TOBArticle.detail[Indice];

    Req := 'SELECT GNE_ARTICLE FROM NOMENENT WHERE GNE_ARTICLE="'+TOBArt.GetValue('GA_ARTICLE')+'"';
    if ExisteSql(Req) then Exit;

    Req := 'UPDATE NOMENENT SET ';
    Req := req + 'GNE_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'GNE_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;

    Req := 'UPDATE NOMENENT SET ';
    if RenvoieTypeArt(TOBArt.GetValue('GA_ARTICLE')) = 'OUV' then
    begin
      Req := req + 'GNE_NOMENCLATURE="'+ NewCode+'" WHERE ';
      REq := REq + 'GNE_NOMENCLATURE="'+CodeArticle+'"';
    end else
    begin
      Req := req + 'GNE_NOMENCLATURE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
      REq := REq + 'GNE_NOMENCLATURE="'+TOBArt.GetValue('PRECEDENT')+'"';
    end;
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
  end;
end;

procedure TChangeCodeArt.TraiteLesNomenLig;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];

    Req := 'SELECT GNL_NOMENCLATURE FROM NOMENLIG WHERE GNL_NOMENCLATURE="'+TOBArt.GetValue('GA_ARTICLE')+'"';
    if ExisteSql(Req) then Exit;

    Req := 'UPDATE NOMENLIG SET ';
    Req := req + 'GNL_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'GNL_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;

    Req := 'UPDATE NOMENLIG SET ';
    if RenvoieTypeArt(TOBArt.GetValue('GA_ARTICLE')) = 'OUV' then
    begin
      Req := req + 'GNL_NOMENCLATURE="'+ NewCode+'" WHERE ';
      REq := REq + 'GNL_NOMENCLATURE="'+CodeArticle+'"';
    end else
    begin
      Req := req + 'GNL_NOMENCLATURE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
      REq := REq + 'GNL_NOMENCLATURE="'+TOBArt.GetValue('PRECEDENT')+'"';
    end;
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;

    Req := 'UPDATE NOMENLIG SET ';
    Req := req + 'GNL_SOUSNOMEN="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'GNL_SOUSNOMEN="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
Req := 'UPDATE NOMENLIG SET ';
Req := req + 'GNL_CODEARTICLE="'+ NewCode+'" WHERE ';
REq := REq + 'GNL_CODEARTICLE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
end;

procedure TChangeCodeArt.TraiteLesLigNomen;
var Indice,NB : integer;
    TOBArt : TOB;
    Req  : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];
    //
    Req := 'UPDATE LIGNENOMEN SET ';
    Req := req + 'GLN_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq +  'GLN_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    //
    Req := 'UPDATE LIGNENOMEN SET ';
    Req := req + 'GLN_COMPOSE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq +  'GLN_COMPOSE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
Req := 'UPDATE LIGNENOMEN SET ';
Req := req + 'GLN_CODEARTICLE="'+ NewCode+'" WHERE ';
REq := REq +  'GLN_CODEARTICLE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
end;

procedure TChangeCodeArt.TraiteLesLigOuv;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin

for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];
    Req := 'UPDATE LIGNEOUV SET ';
    Req := req + 'BLO_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'BLO_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    Req := 'UPDATE LIGNEOUV SET ';
    Req := req + 'BLO_COMPOSE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'BLO_COMPOSE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
Req := 'UPDATE LIGNEOUV SET ';
Req := req + 'BLO_CODEARTICLE="'+ NewCode+'" WHERE ';
REq := REq + 'BLO_CODEARTICLE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
Req := 'UPDATE LIGNEOUV SET ';
Req := req + 'BLO_REFARTSAISIE="'+ NewCode+'" WHERE ';
REq := REq + 'BLO_REFARTSAISIE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
end;

procedure TChangeCodeArt.TraiteLesLignes;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];
    Req := 'UPDATE LIGNE SET ';
    Req := req + 'GL_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'GL_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
Req := 'UPDATE LIGNE SET ';
Req := req + 'GL_CODEARTICLE="'+ NewCode+'" WHERE ';
REq := REq + 'GL_CODEARTICLE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
//
Req := 'UPDATE LIGNE SET ';
Req := req + 'GL_REFARTSAISIE="'+ NewCode+'" WHERE ';
REq := REq + 'GL_REFARTSAISIE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
end;

procedure TChangeCodeArt.TraitelesTobs;
begin
// Articles
TraiteLesArt;
// Complements Articles
if V_PGI.IOError = OeOk Then TraiteLesComplementsArt;
// Articles Liees
if V_PGI.IOError = OeOk Then TraiteLesArtLies;
// Codification Article Chez tiers
if V_PGI.IOError = OeOk Then TraiteLesArtTiers;
// Catalogue Article
if V_PGI.IOError = OeOk Then TraiteLesCatalogues;
// Exceptions Article Piece prfixe GAP
if V_PGI.IOError = OeOk Then TraiteLesArtPiece;
// commisonnement representant
if V_PGI.IOError = OeOk Then TraiteLesArtComm;
// conditionnements articles
if V_PGI.IOError = OeOk Then TraiteLesConditArt;
// traduction articles
if V_PGI.IOError = OeOk Then TraiteLesTradArt;
// entete de nomenclature
if V_PGI.IOError = OeOk Then TraiteLesNomenEnt;
// Ligne de nomenclature
if V_PGI.IOError = OeOk Then TraiteLesNomenLig;
// Ligne nomenclature dans document
if V_PGI.IOError = OeOk Then TraiteLesLigNomen;
// ligne d'ouvrage
if V_PGI.IOError = OeOk Then TraiteLesLigOuv;
// Ligne de consommatons
if V_PGI.IOError = OeOk Then TraiteLesConsommations;
// lignes detail
if V_PGI.IOError = OeOk Then TraiteLesLignes;
// lignes detail ouvrage a plat
if V_PGI.IOError = OeOk Then TraiteLesOuvP;
// gestion du stock
if V_PGI.IOError = OeOk Then TraiteLesStock;
// lignes de décisionnel
if V_PGI.IOError = OeOk Then TraiteLesDecision;
// Paramétrage liaison Paie
if V_PGI.IOError = OeOk Then TraiteActivitePaie;
// Tarifs par article
if V_PGI.IOError = OeOk Then TraiteLesTarifs;
//Traitement des ressources
if V_PGI.IOError = OeOk Then TraiteLesRessources;
end;

procedure InitValoArtNomen (TOBRef : TOB;VenteAchat : string='VEN');
{$IFDEF BTP}
var  TOBCatalog,TOBTarif,TOBTiers : TOB;
     QQ : TQuery;
     MTPAF : double;
     TypeArticle :string;
     Fournisseur : string;
     PrixPourQte : double;
     IsUniteAchat : boolean;
{$ENDIF}
begin
{$IFDEF BTP}
  IsUniteAchat := true; // par défaut
  TypeArticle:= TOBRef.GetValue('GA_TYPEARTICLE');
  Fournisseur := TOBRef.GetValue('GA_FOURNPRINC');
(* OPTIMIZATION *)
  if not TOBRef.FieldExists ('QUALIFQTEACH') then TOBREF.AddChampSupValeur ('QUALIFQTEACH','',false);
  if not TOBRef.FieldExists ('PRXACHBASE') then TOBREF.AddChampSupValeur ('PRXACHBASE',0,false);
  if not TOBRef.FieldExists ('REMISES') then TOBREF.AddChampSupValeur ('REMISES','',false);
  if not TOBRef.FieldExists ('DEJA CALCULE') then TOBREF.AddChampSupValeur ('DEJA CALCULE','-',false);
// --
  if (TypeArticle<> 'MAR') and (TypeArticle <> 'ARP') and (TypeArticle <> 'PRE') then exit;
  if TOBRef.GetValue('DEJA CALCULE')='X' then exit;
  TOBTarif := TOB.Create ('TARIF',nil,-1);
  TOBCatalog := TOB.Create ('CATALOGU',nil,-1);
  TOBTiers := TOB.Create ('TIERS',nil,-1);
  TRY
     if Fournisseur <> '' then
     begin
       QQ := OpenSql ('SELECT * FROM TIERS WHERE T_TIERS="'+fournisseur+'"',true,-1,'',true);
       TOBTiers.SelectDB ('',QQ);
       ferme (QQ);
       chargeTOBs  (TOBRef,TOBCatalog,TOBTiers,TOBTarif);
       if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
       begin
        MTPAF :=TOBTarif.GetValue('GF_PRIXUNITAIRE');
        TOBREF.PutValue ('PRXACHBASE',MTPAF);
       end else
       begin
        PrixPourQte := TOBCatalog.GetValue('GCA_PRIXPOURQTEAC');
        if PrixPourQte = 0 then PrixPourQte := 1;
        if TOBCatalog.GetValue('GCA_PRIXBASE') <> 0 then MTPAF :=TOBCatalog.GetValue('GCA_PRIXBASE')/PrixPourQte
                                                    else MTPAF :=TOBCatalog.GetValue('GCA_PRIXVENTE')/PrixPourQte;
        TOBREF.PutValue ('PRXACHBASE',MTPAF);
       end;
       if MTPAF = 0 then
       begin
         MTPAF := TobRef.getValue('GA_PAHT');
        	TOBREF.PutValue ('PRXACHBASE',MTPAF);
         IsUniteAchat := false;
       end;
       MTPAF := arrondi(TOBREF.GetValue ('PRXACHBASE') * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecP );
       // voila voila voila ..le seul hic c'est que ce prix est en UA..donc passage de l'UA en UV
       if IsUniteAchat then
       begin
       		MTPAF := PassageUAUV (TOBRef,TOBCatalog,MTPAF);
       end;
       //
       (* OPTIMIZATION *)
       TOBREF.PutValue ('QUALIFQTEACH',TOBCatalog.GetValue('GCA_QUALIFUNITEACH'));
       TOBREF.putValue ('REMISES',TOBTarif.GetValue('GF_CALCULREMISE'));
       (* ------------- *)
       TOBREF.putValue('GA_PAHT',MTPAF);
       TOBREF.PutValue ('DEJA CALCULE', 'X');
       RecalculPrPV (TOBRef,TOBCatalog,VenteAchat);
     end;

  FINALLY
     TOBTarif.free;
     TOBCatalog.free;
     TOBTiers.free;
  end;
{$ENDIF}
end;


procedure InitValoArtLigBord (TOBRef : TOB;VenteAchat : string='VEN');
{$IFDEF BTP}
var  TOBCatalog,TOBTarif,TOBTiers : TOB;
     QQ : TQuery;
     MTPAF : double;
     TypeArticle :string;
     Fournisseur : string;
     PrixPourQte : double;
     IsUniteAchat : boolean;
{$ENDIF}
begin
{$IFDEF BTP}
  IsUniteAchat := true; // par défaut
  TypeArticle:= TOBRef.GetValue('GA_TYPEARTICLE');
  Fournisseur := TOBRef.GetValue('GA_FOURNPRINC');
(* OPTIMIZATION *)
  if not TOBRef.FieldExists ('QUALIFQTEACH') then TOBREF.AddChampSupValeur ('QUALIFQTEACH','',false);
  if not TOBRef.FieldExists ('PRXACHBASE') then TOBREF.AddChampSupValeur ('PRXACHBASE',0,false);
  if not TOBRef.FieldExists ('REMISES') then TOBREF.AddChampSupValeur ('REMISES','',false);
  if not TOBRef.FieldExists ('DEJA CALCULE') then TOBREF.AddChampSupValeur ('DEJA CALCULE','-',false);
// --
  if (TypeArticle<> 'MAR') and (TypeArticle <> 'ARP') and (TypeArticle <> 'PRE') then exit;
  TOBTarif := TOB.Create ('TARIF',nil,-1);
  TOBCatalog := TOB.Create ('CATALOGU',nil,-1);
  TOBTiers := TOB.Create ('TIERS',nil,-1);
  TRY
     if Fournisseur <> '' then
     begin
       QQ := OpenSql ('SELECT * FROM TIERS WHERE T_TIERS="'+fournisseur+'"',true,-1,'',true);
       TOBTiers.SelectDB ('',QQ);
       ferme (QQ);
       chargeTOBs  (TOBRef,TOBCatalog,TOBTiers,TOBTarif);
       if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
       begin
        MTPAF :=TOBTarif.GetValue('GF_PRIXUNITAIRE');
       end else
       begin
        PrixPourQte := TOBCatalog.GetValue('GCA_PRIXPOURQTEAC');
        if PrixPourQte = 0 then PrixPourQte := 1;
        if TOBCatalog.GetValue('GCA_PRIXBASE') <> 0 then MTPAF :=TOBCatalog.GetValue('GCA_PRIXBASE')/PrixPourQte
                                                    else MTPAF :=TOBCatalog.GetValue('GCA_PRIXVENTE')/PrixPourQte;
       end;
       if MTPAF = 0 then
       begin
         MTPAF := TobRef.getValue('GA_PAHT');
         IsUniteAchat := false;
       end;
       if (TOBREF.GetValue ('DEJA CALCULE') <> 'X') and (VenteAchat='VEN') then
       begin
            MTPAF := arrondi(MTPAF * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecP );
       end;
       // voila voila voila ..le seul hic c'est que ce prix est en UA..donc passage de l'UA en UV
       if IsUniteAchat then MTPAF := PassageUAUV (TOBRef,TOBCatalog,MTPAF);
       //
       (* OPTIMIZATION *)
       TOBREF.PutValue ('QUALIFQTEACH',TOBCatalog.GetValue('GCA_QUALIFUNITEACH'));
       TOBREF.PutValue ('PRXACHBASE',MTPAF);
       TOBREF.putValue ('REMISES',TOBTarif.GetValue('GF_CALCULREMISE'));
       (* ------------- *)
       TOBREF.putValue('GA_PAHT',MTPAF);
       TOBREF.PutValue ('DEJA CALCULE', 'X');
       RecalculPrPV (TOBRef,TOBCatalog,VenteAchat);
     end;

  FINALLY
     TOBTarif.free;
     TOBCatalog.free;
     TOBTiers.free;
  end;
{$ENDIF}
end;

{$IFDEF BTP}
procedure chargeTOBs (TOBRef,TOBCatalog,TOBTiers,TOBTarif : TOB;QteRef : double = 1; Dateref : TDateTime=0);
var REq : String;
    QQ : TQuery;
    TheTiers : string;
    WithRemiseFouPrinc : boolean;
    TOBTiersloc,TOBTARifLoc : TOB;
    Remise : double;
    StCascade : string;
begin
  WithRemiseFouPrinc := GetParamSocSecur ('SO_BTAPLICREMFOUPRINC',false);
	// verif de base //
  if (VarIsNull(TOBRef.getValue('GA2_SOUSFAMTARART'))) or
  	 (VarAsType(TOBRef.getValue('GA2_SOUSFAMTARART'),VarString)=#0) then
  begin
  	TOBRef.putValue('GA2_SOUSFAMTARART',''); // comme ca on est trankilleeeeee
  end;
  // ------------- //
  TOBCatalog.InitValeurs;
  TheTiers := TOBTiers.GetValue('T_TIERS');
  Req := 'SELECT * FROM CATALOGU WHERE GCA_ARTICLE="'+TOBRef.GetValue('GA_ARTICLE')
       +'" AND GCA_TIERS="'+TheTiers+'" and GCA_DATESUP>="' + USDateTime(V_PGI.DateEntree)+'" ORDER BY GCA_DATEREFERENCE DESC';
  QQ := OpenSql (Req,true,-1,'',true);
  TOBCatalog.SelectDB ('',QQ);
  ferme (QQ);
  if (TOBCatalog.GetValue('GCA_PRIXBASE')=0) and (TOBCatalog.GetValue('GCA_PRIXVENTE')=0) and (TOBRef.GetValue('GA_FOURNPRINC')<>'') then
  begin
    // le fournisseur n'a pas de catalogue produit--> on prend celui du fournisseur principal
    Req := 'SELECT * FROM CATALOGU WHERE GCA_ARTICLE="'+TOBRef.GetValue('GA_ARTICLE')
        +'" AND GCA_TIERS="'+TOBRef.GetValue('GA_FOURNPRINC')+'" and GCA_DATESUP>="' + USDateTime(V_PGI.DateEntree)+'" ORDER BY GCA_DATEREFERENCE DESC';
    QQ := OpenSql (Req,true,-1,'',true);
    if not QQ.eof then TOBCatalog.SelectDB ('',QQ);
    ferme (QQ);
  end;

  Req := 'SELECT * FROM TIERS WHERE T_TIERS="'+TheTiers+'"';
  QQ := OpenSql (Req,true,-1,'',true);
  TOBTiers.SelectDB ('',QQ);
  ferme (QQ);
  GetTarifGlobal (TOBRef.getValue('GA_ARTICLE'),TOBRef.getValue('GA_TARIFARTICLE'),TOBRef.GetValue('GA2_SOUSFAMTARART'),'ACH',TOBRef,TOBTiers,TOBTarif,true,QteRef,DateRef);

  if (WithRemiseFouPrinc) and (TOBRef.GetValue('GA_FOURNPRINC') <> '' ) and (TOBRef.GetValue('GA_FOURNPRINC') <> TOBTIers.getValue('T_TIERS') ) then
  begin
    TOBTiersloc := TOB.create ('TIERS',nil,-1);
    TOBTARifLoc := TOB.Create ('TARIF',nil,-1);
    Req := 'SELECT * FROM TIERS WHERE T_TIERS="'+TOBRef.GetValue('GA_FOURNPRINC')+'"';
    QQ := OpenSql (Req,true);
    TOBTiersloc.SelectDB ('',QQ);
    ferme (QQ);
    GetTarifGlobal (TOBRef.getValue('GA_ARTICLE'),TOBRef.getValue('GA_TARIFARTICLE'),TOBRef.GetValue('GA2_SOUSFAMTARART'),'ACH',TOBRef,TOBTiersLoc,TOBTARifLoc,true);
    //
    if (TOBtarif.getValue('GF_CASCADEREMISE')='CAS') or (TOBtarif.getValue('GF_CASCADEREMISE')='') then
    begin
      Remise := TOBTarif.getValue('GF_REMISE'); // remise pour le founisseur
      stCasCade := TOBTarif.getValue('GF_CALCULREMISE');
      Remise := 100.0 * (1.0 - (1.0 - Remise / 100.0) * (1.0 - TOBTarifloc.getValue('GF_REMISE') / 100));
      if stcascade <> '' then StCascade := TOBTarifloc.getValue('GF_CALCULREMISE')+ '+' + stcascade
      else stCascade := TOBTarifloc.getValue('GF_CALCULREMISE')+ '%';
      TOBTarif.PutValue('GF_REMISE', Remise);
      TOBTarif.PutValue('GF_CALCULREMISE', StCascade);
    end else if (TOBtarif.getValue('GF_CASCADEREMISE')='MIE') and (TOBTARifLoc.getValue('GF_CASCADEREMISE')='MIE') then
    begin
      Remise := TOBTarif.getValue('GF_REMISE'); // remise pour le founisseur
      if Remise < TOBTarifLoc.getValue('GF_REMISE') then
      begin
        TOBTarif.dupliquer (TOBTarifLoc,false,true);
    end;
    end;
    //
    TOBTiersloc.free;
    TOBTARifLoc.free;
  end;
end;

function PassageUAUV (TOBRef,TOBcatalogue : TOB; MTPAF: double) : double; OverLoad;
var UA            : String;
    UV            : String;
    US            : String;
    PrixPourQte   : Double;
    CoefUaUs      : Double;
    CoefUSUV      : double;
begin

  result := MTPAF;

  //FV1 : 18/02/20015 - Gestion de l'unité d'achat dans la fiche article
  CoefUaUS := TOBCatalogue.GetValue('GCA_COEFCONVQTEACH');

  UV := TOBREf.GetValue('GA_QUALIFUNITEVTE');
  US := TOBRef.GetValue('GA_QUALIFUNITESTO');

  if (UV = '') then exit;

  UA := TOBCatalogue.GetValue('GCA_QUALIFUNITEACH');
  if UA = '' then
  begin
    UA := TOBRef.GetValue('GA_QUALIFUNITEACH');
    CoefUaUs := TobRef.GetValue('GA_COEFCONVQTEACH');
  end;

  if (UA = '') then exit;

  PrixPourQte := TobRef.GetValue('GA_PRIXPOURQTE');
  if PrixPourQte = 0 then PrixPourQte := 1;

  CoefUsUV := TOBRef.GetValue('GA_COEFCONVQTEVTE');

  Result := CalculUaUV(CoefUaUs,CoefUSUv,MTPAF,PrixPourQte, UA, UV, US);

end;

function PassageUAUV(TobCatalogue : TOB; MTPAF, PrixPourQte, CoefUniteAUniteS, CoefUniteSUniteV : Double; UniteA, UniteS, UniteV : String) : double; OverLoad;
Var UA        : string;
    UV        : string;
    US        : string;
    //
    CoefUaUs  : Double;
    CoefUSUv  : Double;
begin

  result := MTPAF;
  CoefUSUv := CoefUniteSUniteV;

  //FV1 : 18/02/20015 - Gestion de l'unité d'achat dans la fiche article
  if TobCatalogue <> nil then
  begin
    CoefUaUS := TOBCatalogue.GetValue('GCA_COEFCONVQTEACH');
    //
    UA := TOBCatalogue.GetValue('GCA_QUALIFUNITEACH');
  end;

  if (UA = '') Or (UA <> UniteA) then
  begin
    UA := UniteA;
    CoefUaUs := CoefUniteAUniteS;
  end;

  if (UA = '') then exit;

  if (UniteV = '') then exit;
  UV := UniteV;
  US := UniteS;

  if PrixPourQte = 0 then PrixPourQte := 1;

  Result := CalculUaUv(CoefUaUs,CoefUSUv, MTPAF,PrixPourQte, UA, UV, US);

end;

Function CalculUaUv(CoefUaUs,CoefUSUv,MTPAF,PrixPourQte : Double; UA, UV, US : String) : Double;
Var QuotiteAchat  : Double;
    QuotiteVte    : Double;
    QuotiteSto    : Double;
begin

  QuotiteAchat := RatioMesure('PIE', UA);
  if QuotiteAchat = 0 then QuotiteAchat := 1;
  //
  QuotiteVte := RatioMesure('PIE',UV);
  if QuotiteVte = 0 then QuotiteVte := 1;
  //
  QuotiteSto := RatioMesure('PIE',US);
  if QuotiteSto = 0 then QuotiteSto := 1;

  if CoefUaUs <> 0 then
  begin
    if CoefUSUv <> 0 then
    begin
      result := MTPAF / CoefUAUS / CoefUSUv;
    end else
    begin
  		Result:=MTPAF / CoefuaUs / QuotiteSto * QuotiteVte;
    end;
  end else
  begin
  	Result:=MTPAF * QuotiteVte/QuotiteAchat ;
  end;

  result := result / PrixPourQte;

end;

procedure GetInfoFromCatalog(Fournisseur, Article: string; var UA : string; Var PQQ : double; var CoefuaUs: double);
var QQ : Tquery;
begin

  UA := '';
  PQQ := 1;
  CoefUaUs := 0;

  QQ := OpenSql ('SELECT GCA_QUALIFUNITEACH,GCA_PRIXPOURQTEAC,GCA_COEFCONVQTEACH FROM CATALOGU WHERE GCA_TIERS="'+Fournisseur+'" AND GCA_ARTICLE="'+Article+'"',true,-1,'',true);

  if not QQ.eof Then
  begin
    UA := QQ.findField('GCA_QUALIFUNITEACH').asString;
    PQQ := QQ.findField('GCA_PRIXPOURQTEAC').Index;
    CoefUaUs := QQ.findField('GCA_COEFCONVQTEACH').AsFloat;
  end;

  Ferme (QQ);

end;

procedure GetInfoFromArticle(Article: string; var UA : string; Var PQQ : double; var CoefuaUs: double);
var QQ : Tquery;
begin

  UA        := '';
  PQQ       := 1;
  CoefUaUs  := 0;

  QQ := OpenSql ('SELECT GA_QUALIFUNITEACH,GA_PRIXPOURQTEAC,GA_COEFCONVQTEACH FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"',true,-1,'',true);

  if not QQ.eof Then
  begin
    UA        := QQ.findField('GA_QUALIFUNITEACH').asString;
    PQQ       := QQ.findField('GA_PRIXPOURQTEAC').Index;
    CoefUaUs  := QQ.findField('GA_COEFCONVQTEACH').AsFloat;
  end;

  Ferme (QQ);

end;

{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 13/06/2002
Description .. : Ouverture de la fiche article
*****************************************************************}
Procedure LanceFicheArticle(Lequel: String; Action: TActionFiche = taModif; Params: String = ''; Range : String = '');
begin
  DispatchTTArticle(Action, Lequel, Params, Range);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 13/06/2002
Description .. : Ouverture du Mul article
*****************************************************************}
Procedure LanceMulArticle(Range: String = ''; Lequel: String = ''; Params: String = '');
begin
  AglLanceFiche(NatureMulArticle, NomMulArticle, Range, Lequel, Params);
end;

procedure wCallGA(Range: string);

  Function GetArticle: string;
  begin
    Result := GetArgumentValue(Range, 'GA_ARTICLE');
  end;

var
	ActionFiche : TActionFiche;
  Lequel      : string;
begin
  if GetArticle = '' then exit;
  Lequel := GetArgumentValue(Range, 'GA_ARTICLE');

  if JAiLeDroitGestion('GA') then
    ActionFiche := taModif
  else
    ActionFiche := taConsult;

  LanceFicheArticle(Lequel, ActionFiche, '', Range);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Florence DURANTET
Créé le ...... : 24/10/2002
Description .. : Retourne le statut de l'article (DIM,GEN ou UNI)
Mots clefs ... : STATUT ARTICLE
*****************************************************************}
{$IFDEF GPAO}
function uGetStatutArticle (Article:string) :string;
begin
	Result := wGetFieldFromGA('GA_STATUTART',Article);
end;
{$ENDIF GPAO}

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 21/02/2002
Description .. : Renvoie un where pour le code générique
*****************************************************************}
function whereGA(Article: string): string;
begin
	Result := 'GA_ARTICLE = "' + Article + '"';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 10/06/2002
Modifié le ... :   /  /
Description .. : Recherche d'un article
Suite ........ : sur le code article sur 18 caractères
Mots clefs ... :
*****************************************************************}
function wExistGA(Article: string; WithAlert: Boolean = false): Boolean;
var
	sSql  : string;
begin
	sSql := ' SELECT GA_ARTICLE'
			 + ' FROM ARTICLE'
			 + ' WHERE ' + WhereGA(Article)
         ;
	Result := ExisteSQL(sSql);

  if WithAlert and (not Result) then
  begin
    PgiError(TraduireMemoire('L''article ') + wGetCodeArticleFromArticle(Article) + TraduireMemoire(' n''existe pas.'), 'Article');
  end;
end;

{$IFDEF GPAO} 
{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 25/03/2003
Description .. : Teste l'existence d'un article dans les profils
Suite ........ : avancés (GPF)
*****************************************************************}
function wExistArticleInGPF(Article: string; WithAlert: Boolean = false): Boolean;
begin
  Result := ExisteSQL('SELECT GPF_ARTICLE FROM PROFILART WHERE GPF_ARTICLE="' + Article + '"');
  if WithAlert and (not Result) then
  begin
    PgiError(TraduireMemoire('L''article ') + wGetCodeArticleFromArticle(Article) + TraduireMemoire(' est utilisé dans les profils avancés.'), 'Article');
  end;
end;
{$ENDIF}

{$IFDEF GPAO} 
function wExistArticleProfilInGA(Article: string; WithAlert: Boolean = false): Boolean;
begin
  Result := ExisteSQL('SELECT GA_ARTICLEPROFIL FROM ARTICLE WHERE GA_ARTICLEPROFIL ="' + Article + '"');
  if WithAlert and (not Result) then
  begin
    PgiError(TraduireMemoire('L''article profil ') + wGetCodeArticleFromArticle(Article) + TraduireMemoire(' est utilisé dans les articles.'), 'Article');
  end;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 24/09/2002
Description .. : Appel de la fiche Wartnat
*****************************************************************}
procedure wCallWAN(Article: String; Action: TActionFiche; Params: String);
begin
  if Pos('WAN_ARTICLE=', Params) = 0 then
  begin
    if Params <> '' then Params := Params + ';';
    Params := Params + 'WAN_ARTICLE=' + Article;
  end;
  if Pos('ACTION=', Params) = 0 then
  begin
    if Params <> '' then Params := Params + ';';
    Params := Params + ActionToString(Action);
  end;
  AglLanceFiche('W','WARTNAT_FIC', Article, '', Params);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 21/02/2002
Description .. : renvoie field de GA
*****************************************************************}
function wGetFieldFromGA(Field, Article: string): Variant;
begin
	Result := wGetSqlFieldValue(Field, 'ARTICLE', WhereGA(Article));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 04/03/2002
Description .. : renvoie une liste de champ depuis la fiche article
*****************************************************************}
{$IFDEF GPAO} 
function wGetFieldsFromGA(Fields: Array of string; Article: string): MyArrayValue;
begin
	Result := wGetSqlFieldsValues(Fields, 'ARTICLE', WhereGA(Article));
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 06/12/2002
Description .. : Récupère le 1er article
*****************************************************************}
function wGetFirstGA: String;
var
  Q: TQuery;
begin
  Q := OpenSQL('SELECT GA_ARTICLE FROM ARTICLE ORDER BY GA_ARTICLE', True, 1,'',true);
  try
    if not Q.eof then
      Result := Q.Fields[0].AsString
    else
      Result := '';
  finally
    Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 06/12/2002
Description .. : Récupère le dernier article
*****************************************************************}
function wGetLastGA: String;
var
  Q: TQuery;
begin
  Q := OpenSQL('SELECT GA_ARTICLE FROM ARTICLE ORDER BY GA_ARTICLE DESC', True, 1,'',true);
  try
    if not Q.eof then
      Result := Q.Fields[0].AsString
    else
      Result := '';
  finally
    Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 04/07/2003
Modifié le ... :   /  /
Description .. : Charge la TOBGA avec un ou plusieurs articles suivant le where
Suite ........ : LoadDetailDB ou SelectDB en fonction du paramètre SelectDB
Mots clefs ... :
*****************************************************************}
function wGetTobGA(sChamp, sWhere: String; TobGA: Tob; SelectDB: Boolean = False): Boolean;
var
  sRequete: String;
  Q: TQuery;
begin
  Result := False;
  if (TobGA <> nil) then
  begin
    sRequete := 'SELECT '+sChamp+' FROM ARTICLE LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GA_ARTICLE WHERE '+sWhere;
    if ExisteSQL(sRequete) then
    begin
      Q := OpenSQL(sRequete, True,-1,'',true);
      try
        if not SelectDB then
        begin
          TobGA.LoadDetailDB('ARTICLE', '', '', Q, True, True);
          Result := (TobGA.Detail.Count > 0);
        end
        else
        begin
          TobGA.SelectDB('', Q);
          Result := True;
        end;
      finally
        Ferme(Q);
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 30/01/2002
Description .. : Rajoute un 'X' au 34 ème caractère du code Article.
*****************************************************************}
function wGetArticleFromCodeArticle(CodeArticle : String; CodeDim1: string = ''; CodeDim2: string = ''; CodeDim3: string = ''; CodeDim4: string = ''; CodeDim5: string = '') : String;
begin
  // A VOIR : Il faut utiliser : CodeArticleUnique
  CodeDim1    := wPadRight(CodeDim1, TailleCodeDim);
  CodeDim2    := wPadRight(CodeDim2, TailleCodeDim);
  CodeDim3    := wPadRight(CodeDim3, TailleCodeDim);
  CodeDim4    := wPadRight(CodeDim4, TailleCodeDim);
  CodeDim5    := wPadRight(CodeDim5, TailleCodeDim);
  CodeArticle := wPadRight(CodeArticle, TailleCodeArticle);
  Result := iif(Trim(CodeArticle) <> '', Concat(CodeArticle, CodeDim1, CodeDim2, CodeDim3, CodeDim4, CodeDim5, 'X'), '');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 07/08/2002
Description .. : Ramène le Code article sans dimension
*****************************************************************}
function wGetCodeArticleFromArticle(Article: string): string;
begin
	Result := Trim(Copy(Article, 1, TailleCodeArticle));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 15/01/2003
Description .. : Renvoie l'Article Générique à partir d'un article dimensioné
*****************************************************************}
function wGetArticleGENFromArticleDIM(Article: string): string;
begin
  Result := wGetArticleFromCodeArticle(wGetCodeArticleFromArticle(Article));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 26/09/2002
Description .. : Retourne vrai si le nom de champ est un code article
*****************************************************************}
{$IFDEF GPAO} 
function wIsSuffixeCodeArticle(NomChamp: String): Boolean;
begin
  Result := Pos(wGetSuffixe(NomChamp), 'CODEARTICLE;CODECOMPOSANT') <> 0;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 26/09/2002
Description .. : Retourne vrai si le nom de champ est un article
*****************************************************************}
{$IFDEF GPAO} 
function wIsSuffixeArticle(NomChamp: String): Boolean;
var
  i: Integer;
  s: String;
begin
  { Traitement des suffixes xxx_ARTICLE1, xxx_COMPOSANT2 }
  Result := False; i := -1;
  while (i < 9) and not Result do
  begin
    Inc(i);
    s := StringReplace(wGetSuffixe(NomChamp), IntToStr(i), '', [rfReplaceAll]);
    Result := (Pos(s, 'ARTICLE;COMPOSANT') <> 0);
  end;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 27/09/2002
Description .. : Vérifie si le CodeArticle est correct et ajoute un X à la fin si besoin
*****************************************************************}
Function xSetCodeArticle(Var SaisArticle: String): Boolean;
begin
	Result := True;
  if Copy(SaisArticle, TailleArticle-1, 1) <> 'X' then
  begin
    if Length(SaisArticle) > TailleCodeArticle then
      result := false
    else
      SaisArticle := wGetArticleFromCodeArticle(SaisArticle);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Retourne le qualifiant de la table MEA pour un nom de
Suite ........ : champ de la table article
Mots clefs ... :
*****************************************************************}
function wGetQualifUniteArticle(CodeChamp: String): String;
begin
  if CodeChamp = 'GA_QUALIFPOIDS' then Result := 'POI'
  else if CodeChamp = 'GA_QUALIFLINEAIRE' then Result := 'LIN'
  else if CodeChamp = 'GA_QUALIFSURFACE' then Result := 'SUR'
  else if CodeChamp = 'GA_QUALIFVOLUME' then Result := 'VOL'
  else if CodeChamp = 'GA_QUALIFHEURE' then Result := 'TEM'
  else Result := '';
end;

{$IFDEF GPAO}
{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetein
Créé le ...... : 20/11/2002
Description .. : Création d'un article à partir d'un autre article.
                 L'initialisation des champs de l'article à créer
                 est piloté par les profils.
                 Le code profil à utiliser est recherché dans le
                 champ GP_PROFILARTICLE de l'article de référence.
*****************************************************************}
Function wDuplicGA(Article, ArticleProfil: String): String;
var
  TobArticle, TobArticleProfil, TobProfil: Tob;
  Sql, CodeProfil: String;
  OkProfil: Boolean;
  Q: TQuery;
begin
	Result := '';
  if Article <> '' then
  begin
   	{ Code article => article }
    if not xSetCodeArticle(Article) then
    begin
		  Result := TraduireMemoire('Code article incorrect') + ' : ' + Article;
      EXIT;
    end;
   	{ Code Profil => article }
    if not xSetCodeArticle(ArticleProfil) then
    begin
			Result := TraduireMemoire('Code profil incorrect') + ' : ' + ArticleProfil;
      EXIT;
    end;
    if not wExistGA(Article) then             { non existence de l'article à crée }
    begin
      if wExistGA(ArticleProfil) then        { Existence de l'article de référence }
      begin
        { Tob contenant l'article de référence }
        TobArticleProfil := Tob.Create('ARTICLE', nil, -1);
        try
          wGetTobGA('*', whereGA(ArticleProfil), TobArticleProfil, True);
          { Tob du profil servant à l'initialisation de l'article }
          CodeProfil := TobArticleProfil.GetValue('GA_PROFILARTICLE');
          if CodeProfil = '' then
            CodeProfil := GetParamSoc('SO_GCPROFILART');
          if CodeProfil <> '' then
          begin
            TobProfil := Tob.Create('PROFILART', nil, -1);
            try
              Sql := 'SELECT * FROM PROFILART WHERE GPF_PROFILARTICLE="' + CodeProfil + '"';
              OkProfil := False;
              try
                Q := OpenSQL(Sql, True,-1,'',true);
                if not Q.Eof then
                begin
                  TobProfil.Selectdb('', Q);
                  OkProfil := True;
                end;
              finally
                Ferme(Q);
              end;
              if OkProfil then
              begin
                { Tob de l'article à créer }
                TobArticle := Tob.Create('ARTICLE', nil, -1);
                try
                  { Initialise le code article }
                  TobArticle.PutValue('GA_ARTICLE', Article);
                  TobArticle.PutValue('GA_CODEARTICLE', wGetCodeArticleFromArticle(Article));
                  TobArticle.PutValue('GA_PROFILARTICLE', CodeProfil);
                  TobArticle.PutValue('GA_ARTICLEPROFIL', ArticleProfil);
                  { Initialisation des champs de la TobArticle }
                  if wSetAdvancedProfilToArticleByTob(TobArticle, TobProfil) then
                  begin
                    { Création de l'article }
                    if TobArticle.InsertDB(nil) then
                    begin
                       { Création des natures de travail associées }
                       wDuplicWAN(ArticleProfil, Article);
                    end
                    else
                       Result := TraduireMemoire('La création de l''article a échouée');
                  end
                  else
                    Result := TraduireMemoire('Impossible d''initialiser l''article')
                            + ' : ' + wGetCodeArticleFromArticle(Article) + ' '
                            + TraduireMemoire('à partir du profil') + ' : ' + wGetCodeArticleFromArticle(ArticleProfil);
                finally
                  TobArticle.free;
                end;
              end
              else
                Result := TraduireMemoire('Le code profil') + ' : ' + CodeProfil + ' ' + TraduireMemoire('contenu dans l''article de référence') + ' : ' + wGetCodeArticleFromArticle(ArticleProfil) + ' ' + TraduireMemoire('n''existe pas');
            finally
              TobProfil.Free;
            end;
          end
          else
            Result := TraduireMemoire('L''article') + ' : ' + wGetCodeArticleFromArticle(ArticleProfil) + ' ' + TraduireMemoire('ne contient pas de profil de référence');
        finally
          TobArticleProfil.Free;
        end;
      end
      else
        Result := TraduireMemoire('Le profil') + ' : ' + wGetCodeArticleFromArticle(ArticleProfil) + ' ' + TraduireMemoire('est inconnu');
    end
    else
      Result := TraduireMemoire('L'' article') + ' : ' + wGetCodeArticleFromArticle(Article) + ' ' + TraduireMemoire('existe déjà');
  end
  else
    Result := TraduireMemoire('Le code de l''article a créer n''est pas renseigné');
end;
{$ENDIF}

{$IFDEF GPAO}
Function wUpDateArticleFromTob(Article: String; TheTob: Tob): String;
var
  iChamp: Integer;
  FieldName: String;
  Exist, IsValid: Boolean;
  TobArticle: Tob;
  TomArticle: Tom;
begin
	Result := '';
  { X Dans le code article }
  if not xSetCodeArticle(Article) then
  begin
    Result := TraduireMemoire('Code article incorrect') + ' : ' + Article;
    EXIT;
  end;
  { Contrôle existence de l'article }
	if wExistGA(Article) then
  begin
   	if TheTob <> nil then
    begin
      { Crée une TobArticle et contrôle l'existence des champs de la tob passée en paramêtre }
      TobArticle := Tob.Create('ARTICLE', nil, -1);
      try
        { Charge la tob avec les données de l''article }
        wGetTobGA('*', whereGA(Article), TobArticle, True);
        TobArticle.SetAllModifie(False);
        { Recopie la tob à mettre à jour dans la tob article }
        iChamp := 999; Exist := True;
        while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) and Exist do
        begin
          Inc(iChamp);
          FieldName := TheTob.GetNomChamp(iChamp);
          { Vérifie si le champ fait partie de la table article }
          Exist := TobArticle.FieldExists(FieldName);
          if Exist then TobArticle.PutValue(FieldName, TheTob.GetValue(FieldName));
        end;
        if Exist then
        begin
          { Vérifie les données en passant par une TomArticle }
          TomArticle := CreateTOM('ARTICLE', nil, false, true);
          try
            IsValid := TomArticle.VerifTOB(TobArticle);
          finally
            TomArticle.Free;
          end;
          if IsValid then
            TobArticle.UpdateDB(False) { Enregistre la TobArticle }
          else
            Result := TraduireMemoire('Les données de l''article : ') + Article + TraduireMemoire(' ne sont pas valides');
        end
        else
          Result := TraduireMemoire('Le champ : ') + FieldName + TraduireMemoire(' n''existe pas dans la table : ') + 'ARTICLE';
       finally
        TobArticle.Free;
       end;
    end
    else
      Result := TraduireMemoire('Paramètres d''appel incorrect (Tob non renseignée)');
  end
  else
    Result := TraduireMemoire('L''article : ') + Article + TraduireMemoire(' est inconnu');
end;
{$ENDIF}

{$IFDEF GPAO}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 04/04/2002
Description .. : Applique un profil avancé aux articles associés
*****************************************************************}
function wSetAdvancedProfil(From: TwProfil; DS: TDataSet): Integer;
var
	TheTob: Tob;
begin
  TheTob := Tob.Create('ARTICLE', nil, -1);
  try
    { Crée une tob à partir du DataSet }
    wMakeTobFromDS(DS, TheTob);
    { Appel la procedure de mise à jour }
    Result := WSetAdvancedProfil2(From, TheTob);
  finally
    TheTob.Free;
  end;
end;
{$ENDIF}

{$IFDEF GPAO} 
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 04/04/2002
Description .. : Application d'un profil avancé aux articles
*****************************************************************}
function wSetAdvancedProfil2(From: TwProfil; TobArticle: Tob): Integer;
{
	Application d'un profil aux articles
   ------------------------------------

   Paramètres :

      From				 : Pour identifier l'origine de l'appel (la pelle ?)
      					  		paMajArticle 					=> Appel depuis la TomArticle / Modification d'un article profil => Remise à jour des articles associés
							  		paMajArticleFromProfil 		=> Appel depuis la TomProfilArt / Modification d'un profil => Remise à jour des articles associés
   	TobArticle		 : Tob contenant l'article de référence du profil

   La TobArticle contient les données de l'article de référence du profil.
   Ces données n'ont pas encore été mise à jour dans la table article (Appel depuis OnUpdateRecord)
   On peut donc appliquer le filtre en mode 'protégé' c'est à dire que seuls les champs des articles
   contenant une valeur identique à celle de l'article de référence seront mises à jour...

   Cette function renvoi un numéro d'erreur qui fait référence au TexteMessage ci-dessus
   ou si tout est Ok le nombre d'article mis à jour en négatif
}
var
  Q: TQuery;
  ModeRemplace, Requete, SetSQL, Msg, ArticleDuProfil, CodeProfil: String;
  TobUnAutorizedFields, TobProfil, TobArticlesProfil, OldValuesProfil: Tob;
  ListFields: TStrings;
  i, NbrArticle: Integer;
  WithArticleInProfil: Boolean;

  Function CanCopy(sFieldName: String): Boolean;
  begin
    Result := (TobUnAutorizedFields.FindFirst(['WCA_NOMCHAMP'], [sFieldName], True) = nil);
  end;

  Function MakeColumnSet(iTobArticle : Integer): String;
  var
    sValue: String;
  begin
    Result := '';
    if (iTobArticle <> -1) then
    begin
      sValue := WFieldValueToStringWithCote(TobArticle.GetNomChamp(iTobArticle), TobArticle.GetValeur(iTobArticle));
      if sValue <> '' then Result := TobArticle.GetNomChamp(iTobArticle) + '=' + sValue;
    end;
  end;

  Procedure AddNewSet(NewSet: String);
  begin
    if NewSet <> '' then
    begin
      if SetSQL <> '' then SetSQL := SetSQL + ',';
      SetSQL := SetSQL + NewSet;
    end;
  end;

  Procedure PutValueInTob(sFieldName: String);
  var
    i: Integer;
  begin
    for i := 0 to TobArticlesProfil.Detail.Count - 1 do
    begin
      if TobArticlesProfil.Detail[i].GetValue(sFieldName) = OldValuesProfil.GetValue(sFieldName) then
        TobArticlesProfil.Detail[i].PutValue(sFieldName, TobArticle.GetValue(sFieldName));
    end;
  end;

begin
	Result := 0;
  if Assigned(TobArticle) then
  begin
		if From in [paMajArticle, paMajArticleFromProfil] then
    begin
         { Récupère le profil associé à l'article }
         Requete := 'SELECT * FROM PROFILART WHERE GPF_ARTICLE="' + TobArticle.GetValue('GA_ARTICLE') + '"';
         if not ExisteSQL(Requete) then
         begin
            { Aucun profil n'utilise cet article, on utilise le premier profil avancé sans code article }
            Requete := 'SELECT * FROM PROFILART WHERE GPF_ARTICLE="" AND GPF_TYPEPROFILART="ADV"';
            if not ExisteSQL(Requete) then
               Result := 3;
         end;
    end;
    if Result = 0 then
    begin
      WithArticleInProfil := False;
      TobProfil := Tob.Create('PROFILART', nil, -1);
      try
        Q := OpenSQL(Requete, True,-1,'',true);
        try
          TobProfil.SelectDB('', Q);
        finally
          Ferme(Q);
        end;
        { Contrôle si le profil est un profil avancé }
        if (TobProfil.GetValue('GPF_TYPEPROFILART') = 'ADV' ) then
        begin
          if ((From in [paMajArticle, paCreeArticle]) and (TobProfil.GetValue('GPF_MAJFICHEARTICLE') = wTrue)) or (From = paMajArticleFromProfil) then
          begin
            { Code profil }
            CodeProfil := TobProfil.GetValue('GPF_PROFILARTICLE');
            { Code article de référence }
            if WithArticleInProfil then
              ArticleDuProfil := TobProfil.GetValue('GPF_PROFILARTICLE')
            else
              ArticleDuProfil := TobArticle.GetValue('GA_ARTICLE');
            { Détermine le nombre d'article à traiter }
            if WithArticleInProfil and (ArticleDuProfil <> '') then
            begin
              { Il y a un code article dans le profil = Vérifie si des articles utilise le profil (Même code profil) }
              Requete := 'SELECT COUNT(GA_ARTICLE) FROM ARTICLE WHERE GA_PROFILARTICLE="' + CodeProfil + '"';
              Q := OpenSQL(Requete, True,-1,'',true);
              try
                if not Q.EOF then
                  NbrArticle := Q.Fields[0].AsInteger
                else
                  NbrArticle := 0;
              finally
                Ferme(Q);
              end;
            end
            else
            begin
              { Pas de profil dans le code article = Vérifie si des articles ont été crée à partir de l'article profil }
              Requete := 'SELECT COUNT(GA_ARTICLE) FROM ARTICLE WHERE GA_ARTICLEPROFIL="' + ArticleDuProfil + '"';
              Q := OpenSQL(Requete, True,-1,'',true);
              try
                if not Q.EOF then
                  NbrArticle := Q.Fields[0].AsInteger
                else
                  NbrArticle := 0;
              finally
                Ferme(Q);
              end;
            end;
            if NbrArticle > 0 then
            begin
              { Préparation du message de confirmation }
              if WithArticleInProfil then
                Msg := TraduireMemoire('Mettre à jour les articles associés au profil') + ' : ' + CodeProfil + ' ?'
              else
                Msg := TraduireMemoire('Mettre à jour les articles associés à l''article') + ' : ' + wGetCodeArticleFromArticle(ArticleDuProfil) + ' ?';
              Msg := Msg + #13 + '(' + TraduireMemoire('Il y a') + ' ' + IntToStr(NbrArticle) + ' ' + TraduireMemoire('article(s) à mettre à jour') + ')';
              { Message de confirmation }
              if PGIAsk(Msg, TraduireMemoire('Profil avancé')) = MrYes then
              begin
                ModeRemplace := TobProfil.GetValue('GPF_MODEREMPLACE');
                if (ModeRemplace = 'SAF') or (ModeRemplace = 'ECR') then
                begin
                  ListFields := TStringList.Create;
                  WGetProfilFieldList(TobProfil, ListFields);
                  TobUnAutorizedFields := Tob.Create('WCHAMP', nil, -1);
                  OldValuesProfil := Tob.Create('ARTICLE', nil, -1);
                  TobArticlesProfil := Tob.Create('_ARTICLE_', nil, -1);
                  try
                    { Récupère la tob des champs systématiquement exclu de la copie }
                    wLoadTobUnAutorizedFieldsInProfil(TobUnAutorizedFields);
                    { Mode 'protégé' récupère les champs à mettre à jour }
                    if TobProfil.GetValue('GPF_MODEREMPLACE') = 'SAF' then
                    begin
                      { Récupère les champs de l'article de référence du profil dans la table article }
                      if wGetTobGA('*', 'GA_ARTICLE="' + ArticleDuProfil + '"', OldValuesProfil, True) then
                      begin
                        { Récupère les articles associés au profil qu'il va falloir mettre à jour }
                        if WithArticleInProfil then
                          wGetTobGA('*', 'GA_PROFILARTICLE="' + CodeProfil + '" AND GA_ESTPROFIL="-"', TobArticlesProfil)
                        else
                          wGetTobGA('*', 'GA_ARTICLEPROFIL="' + ArticleDuProfil + '" AND GA_ESTPROFIL="-"', TobArticlesProfil);
                        { Flag aucun champ modifié dans la Tob des articles associés au profil }
                        for i := 0 to TobArticlesProfil.Detail.Count -1 do
                          TobArticlesProfil.Detail[i].SetAllModifie(False);
                      end
                      else
                        Result := 9;
                    end;
                    if Result = 0 then
                    begin
                      ModeRemplace := TobProfil.GetValue('GPF_MODEREMPLACE');
                      { Mode 'brutal' => Les articles sont systématiquement remis à jour à partir de l'article profil }
                      { Forme le Set du Update }
                      SetSQL := '';
                      if TobProfil.GetValue('GPF_MODECOPIE') = 'INC' then { Ne copie que les champs figurant dans la liste }
                      begin
                          for i := 0 to ListFields.Count - 1 do
                          begin
                             { Le champ n'est pas en liste rouge }
                             if CanCopy(ListFields[i]) then
                             begin
                                if ModeRemplace = 'ECR' then
                                   AddNewSet(MakeColumnSet(TobArticle.GetNumChamp(ListFields[i])))
                                else
                                   PutValueInTob(ListFields[i]);
                             end;
                          end;
                      end
                      else if TobProfil.GetValue('GPF_MODECOPIE') = 'EXC' then  { Ne copie pas les champs figurant dans la liste }
                      begin
                          for i := 1 to TobArticle.NbChamps do
                          begin
                             { Le champ ne figure pas dans la liste et il n'est pas sur la liste rouge }
                             if (ListFields.IndexOf(TobArticle.GetNomChamp(i)) = -1) and (CanCopy(TobArticle.GetNomChamp(i))) then
                             begin
                                if ModeRemplace = 'ECR' then
                                   AddNewSet(MakeColumnSet(i))
                                else
                                   PutValueInTob(TobArticle.GetNomChamp(i));
                             end;
                          end;
                      end
                      else
                        Result := 5; { Mode de copie incorrect dans le profil }
                      if (Result = 0) and (ModeRemplace = 'ECR') and (SetSQL <> '') then
                      begin
                          InitMoveProgressForm(nil, TraduireMemoire('Mise à jour des articles'), '', 1, false, true);
                          { Mise à jour des articles }
                          Requete := 'UPDATE ARTICLE SET ' + SetSQL;
                          if WithArticleInProfil then
                             Requete := Requete + ' WHERE GA_PROFILARTICLE="' + CodeProfil + '"'
                          else
                             Requete := Requete + ' WHERE GA_ARTICLEPROFIL="' + ArticleDuProfil + '"';
                          { Nombre d'article modifié }
                          Result := -1 * ExecuteSql(Requete);
                          FiniMoveProgressForm;
                      end
                      else if (Result = 0) and (ModeRemplace = 'SAF') then
                      begin
                        InitMoveProgressForm(nil,TraduireMemoire('Mise à jour des articles'), '', TobArticlesProfil.Detail.Count, false, true);
                        Result := 0;
                        if TobArticlesProfil.IsOneModifie then
                        begin
                          for i := 0 to TobArticlesProfil.Detail.Count - 1 do
                          begin
                            if TobArticlesProfil.detail[i].IsOneModifie(False) then
                            begin
                              TobArticlesProfil.detail[i].UpdateDB;
                              Inc(Result);
                            end;
                            MoveCurProgressForm(TraduireMemoire('Article : ') + TobArticlesProfil.Detail[i].GetValue('GA_ARTICLE'));
                          end;
                          Result := -1 * Result;
                        end;
                        FiniMoveProgressForm;
                      end;
                    end;
                  finally
                    ListFields.Free;
                    TobUnAutorizedFields.Free;
                    OldValuesProfil.Free;
                    TobArticlesProfil.Free;
                  end;
                end
                else
                 Result := 8; { Mode de remplacement des profils incorrect }
              end
              else
                Result := 6; { Mise à jour abandonnée }
            end
            else
               Result := 2; { pas d'article pour le profil }
          end
          else
            Result := 7; { Mise à jour non autorisée }
        end
        else
           Result := 4; { Pas profil avancé }
      finally
        TobProfil.Free;
      end;
    end
    else
      Result := 3; { Profil inconnu }
  end
  else
   	Result := 1; { TobArticle nil }

  if (Result > 0) then
  begin
    if Result = 7 then
      PGIInfo('Attention : ' + #13 + TexteMessageProfil[Result] , 'Mise à jour des profils articles')
    else
      if (Result <> 6) then PGIError('Erreur lors de la mise à jour des articles associés au profil' + #13 + #10 + TexteMessageProfil[Result] , 'Profil Article');
  end;
end;
{$ENDIF}

{$IFDEF GPAO}
function wSetAdvancedProfilToArticle(DSArticle: TDataSet; QProfil: TQuery): Boolean;
{
  Lors de la création d'un article, applique les données du profil contenu dans QProfil à l'article contenu dans DSArticle
  DSArticle: DataSet contenant les données de l'article qui va être crée
  QProfil: Query contenant le profil à appliquer
}
var
  TobArticle, TobProfil: Tob;
begin
  { Création des tobs }
  TobArticle := Tob.Create('ARTICLE', nil, -1);
  TobProfil := Tob.Create('PROFILART', nil, -1);
  try
    { Charge la TobArticle à mettre à jour à partir du DataSet }
    wMakeTobFromDS(DSArticle, TobArticle);
    { Charge la TobProfil à partir du Query }
    TobProfil.SelectDB('', QProfil);
    Result := wSetAdvancedProfilToArticleByTob(TobArticle, TobProfil);
    if Result then
    begin
      { Mise à jour du DATASET Article }
      if (DSArticle.State = dsBrowse) then DSArticle.Edit;
{$IFDEF EAGLCLIENT}
      wSetDSFromTob(TobArticle, DSArticle);
{$ELSE}
      TobArticle.PutDataSet(DSArticle);
{$ENDIF}
    end;
  finally
    TobArticle.Free;
    TobProfil.Free;
  end;
end;
{$ENDIF}

{$IFDEF GPAO}
Function wSetAdvancedProfilToArticleByTob(TobArticle, TobProfil: TOb): Boolean;
{
  Lors de la création d'un article,
  applique les données du profil contenu dans TobProfil à l'article contenu dans TobProfil

  TobArticle: Tob contenant les données de l'article qui va être crée
  TobProfil: Tob contenant le profil à appliquer

  Cette procédure ne s'occupe pas de déterminer le profil à appliquer
}
var
  TobProfilArticle, TobUnAutorizedFields: Tob;
  CodeProfil, Article, ProfilArticle, ModeRemplace, NomChamp: String;
  i: Integer;
  ListFields: TStrings;

  function CanCopy(sFieldName: String): Boolean;
  begin
    Result := (TobUnAutorizedFields.FindFirst(['WCA_NOMCHAMP'], [sFieldName], True) = nil);
  end;

  Procedure PutValueInTobArticle(NomChamp: String);
  begin
    TobArticle.PutValue(NomChamp, TobProfilArticle.GetValue(Nomchamp));
  end;

begin
	Result := False;
  { Création des tobs }
  TobProfilArticle := Tob.Create('ARTICLE', nil, -1);
  try
    { Article à traiter }
    Article := TobArticle.GetValue('GA_ARTICLE');
    { Récupère le code profil }
    CodeProfil := TobProfil.GetValue('GPF_PROFILARTICLE');
    if (CodeProfil <> '') then
    begin
      { Vérifie que le profil est un profil avancé }
      if TobProfil.GetValue('GPF_TYPEPROFILART') = 'ADV' then
      begin
        { Récupère l'article de référence }
        ProfilArticle := TobProfil.GetValue('GPF_ARTICLE');
        if ProfilArticle = '' then ProfilArticle := TobArticle.GetValue('GA_ARTICLEPROFIL');
        if ProfilArticle = '' then
        begin
          { Demande à l'utilisatrice de sélectionner l'article profil }
          ProfilArticle := GetArticleRecherche_Disp(nil, '', 'RECHERCHEARTICLE;ONLYPROFIL=X', NomMulArticle);
        end;
        if ProfilArticle <> '' then
        begin
          { Charge la tob de l'article profil }
          if wGetTobGA('*', 'GA_ARTICLE="' + ProfilArticle + '"', TobProfilArticle, True) then
          begin
            ModeRemplace := TobProfil.GetValue('GPF_MODEREMPLACE');
            if (ModeRemplace = 'SAF') or (ModeRemplace = 'ECR') then
            begin
              ListFields := TStringList.Create;
              try
                { Récupère la liste des champs gérés dans le profil }
                wGetProfilFieldList(TobProfil, ListFields);
                { Récupère la tob des champs systématiquement exclu de la copie }
                TobUnAutorizedFields := Tob.Create('WCHAMP', nil, -1);
                wLoadTobUnAutorizedFieldsInProfil(TobUnAutorizedFields);
                { Copie des champs }
                if TobProfil.GetValue('GPF_MODECOPIE') = 'INC' then { Ne copie que les champs figurant dans la liste }
                begin
                  for i := 0 to ListFields.Count - 1 do
                  begin
                    { Le champ n'est pas en liste rouge }
                    if CanCopy(ListFields[i]) then
                      PutValueInTobArticle(ListFields[i]);
                  end;
                end
                else if TobProfil.GetValue('GPF_MODECOPIE') = 'EXC' then { Ne copie pas les champs figurant dans la liste }
                begin
                  for i := 1 to TobProfilArticle.NbChamps do
                  begin
                    NomChamp := TobProfilArticle.GetNomChamp(i);
                    { Le champ ne figure pas dans la liste et il n'est pas sur la liste rouge }
                    if (ListFields.IndexOf(NomChamp) = -1) and (CanCopy(NomChamp)) then
                      PutValueInTobArticle(NomChamp);
                  end;
                end;
                { Dans tous les cas mise à jour du code article profil }
                TobArticle.PutValue('GA_ARTICLEPROFIL', ProfilArticle);
                Result := True;
              finally
                ListFields.Free;
                TobUnAutorizedFields.Free;
              end;
              Result := True;
            end
            else
              PgiError(TraduireMemoire('Mode de remplacement incorrect, vérifiez le profil') + ' : ' + CodeProfil, '');
          end
          else
            PgiError(TraduireMemoire('Article profil inconnu, impossible d''appliquer le profil') + #13 + '(' + ProfilArticle + ')', '');
        end
        else
          PgiError(TraduireMemoire('Article profil non renseigné, impossible d''appliquer le profil'), '');
      end
      else
        PgiError(TraduireMemoire('Ce profil n''est pas un profil ''avancé'', impossible de l''appliquer'), '');
    end
    else
       PgiError(TraduireMemoire('Code profil non renseigné, impossible d''appliquer le profil'), '');
  finally
    TobProfilArticle.Free;
  end;
end;
{$ENDIF}

{$IFDEF GPAO} 
Procedure wGetProfilFieldList(TobProfil: Tob; TS: tStrings);

var
  s, s8: String;
  i: Integer;
begin
	if (TobProfil <> nil) and (TS <> nil) and (TobProfil.FieldExists('GPF_LISTECHAMPS1'))
													  and (TobProfil.FieldExists('GPF_LISTECHAMPS2'))
													  and (TobProfil.FieldExists('GPF_LISTECHAMPS3'))
													  and (TobProfil.FieldExists('GPF_LISTECHAMPS4'))
													  and (TobProfil.FieldExists('GPF_LISTECHAMPS5')) then
  begin
    s := '';
    for i := 1 to 5 do
      s := s + TobProfil.GetValue('GPF_LISTECHAMPS' + IntToStr(i)) + ';';
    while (s <> '') do
    begin
      s8 := ReadTokenSt(s);
      if (WGetLibChamp(s8) <> '') then TS.Add(s8);
    end;
  end;
end;
{$ENDIF}

{$IFDEF GPAO}
Procedure wGetAdvancedProfilFieldsList(TobProfil, TobUnAutorizedFields: Tob; TS: tStrings);
var
  TSTmp: TStrings;
  TobArticle: Tob;
  i: Integer;

  function CanCopy(sFieldName: String): Boolean;
  begin
    Result := (TobUnAutorizedFields.FindFirst(['WCA_NOMCHAMP'], [sFieldName], True) = nil);
  end;

begin
	if (TS <> nil) and (TobProfil <> nil) and (TobProfil.GetValue('GPF_TYPEPROFILART') = 'ADV') and (TobUnAutorizedFields <> nil) then
  begin
    TSTmp := TStringList.Create;
    try
      { Charge la liste des champs }
      WGetProfilFieldList(TobProfil, TSTmp);
      if TobProfil.GetValue('GPF_MODECOPIE') = 'INC' then { Champs inclus }
      begin
        for i := 0 to TSTmp.Count -1 do
        begin
          if CanCopy(TSTmp[i]) then TS.Add(TsTmp[i]);
        end;
      end
      else if TobProfil.GetValue('GPF_MODECOPIE') = 'EXC' then { Champs exclus }
      begin
        TobArticle := Tob.Create('ARTICLE', nil, -1);
        try
          for i := 1 to TobArticle.NbChamps do
          begin
            if (TSTmp.IndexOf(TobArticle.GetNomChamp(i)) = -1) and CanCopy(TobArticle.GetNomChamp(i)) then TS.Add(TobArticle.GetNomChamp(i));
          end;
        finally
          TobArticle.Free;
        end;
      end;
    finally
      TSTmp.Free;
    end;
  end;
end;
{$ENDIF}

{$IFDEF GPAO}
Procedure wLoadTobUnAutorizedFieldsInProfil(TheTob: Tob);
var
  Q: tQuery;
  sRequete: String;
begin
	if TheTob <> nil then
  begin
    sRequete := 'SELECT * FROM WCHAMP WHERE WCA_CONTEXTEPROFIL="NOC"';
    if ExisteSQL(sRequete) then
    begin
      Q := OpenSQL(sRequete, True,-1,'',true);
      try
        TheTob.LoadDetailDB('', '', '', Q, False);
      finally
        Ferme(Q);
      end;
    end;
  end;
end;
{$ENDIF}


procedure TChangeCodeArt.TraiteLesStock;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
  for Indice := 0 to TOBArticle.detail.count -1 do
  begin
    TOBArt := TOBArticle.detail[Indice];

    Req := 'SELECT GQ_ARTICLE FROM DISPO WHERE GQ_ARTICLE="'+TOBArt.GetValue('GA_ARTICLE')+'"';
    if ExisteSql(Req) then Exit;

    Req := 'UPDATE DISPO SET ';
    Req := req + 'GQ_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'GQ_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
  end;
end;

procedure TChangeCodeArt.TraiteLesOuvP;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];
    Req := 'UPDATE LIGNEOUVPLAT SET ';
    Req := req + 'BOP_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'BOP_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
Req := 'UPDATE LIGNEOUVPLAT SET ';
Req := req + 'BOP_CODEARTICLE="'+ NewCode+'" WHERE ';
REq := REq + 'BOP_CODEARTICLE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
//
Req := 'UPDATE LIGNEOUVPLAT SET ';
Req := req + 'BOP_REFARTSAISIE="'+ NewCode+'" WHERE ';
REq := REq + 'BOP_REFARTSAISIE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
end;

procedure TChangeCodeArt.TraiteLesDecision;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];
    Req := 'UPDATE DECISIONACHLIG SET ';
    Req := req + 'BAD_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'BAD_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    Req := 'UPDATE DECISIONACHLFOU SET ';
    Req := req + 'BDF_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'BDF_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
Req := 'UPDATE DECISIONACHLIG SET ';
Req := req + 'BAD_CODEARTICLE="'+ NewCode+'" WHERE ';
REq := REq + 'BAD_CODEARTICLE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
Req := 'UPDATE DECISIONACHLFOU SET ';
Req := req + 'BDF_REFARTTIERS="'+ NewCode+'" WHERE ';
REq := REq + 'BDF_REFARTTIERS="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
end;

procedure TChangeCodeArt.TraiteActivitePaie;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];
    Req := 'UPDATE ACTIVITEPAIE SET ';
    Req := req + 'ACP_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'ACP_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
Req := 'UPDATE ACTIVITEPAIE SET ';
Req := req + 'ACP_CODEARTICLE="'+ NewCode+'" WHERE ';
REq := REq + 'ACP_CODEARTICLE="'+CodeArticle+'"';
Nb := ExecuteSql (Req);
if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
end;

procedure TChangeCodeArt.TraiteLesTarifs;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin
  for Indice := 0 to TOBArticle.detail.count -1 do
    begin
    TOBArt := TOBArticle.detail[Indice];
    Req := 'UPDATE TARIF SET ';
    Req := req + 'GF_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'" WHERE ';
    REq := REq + 'GF_ARTICLE="'+TOBArt.GetValue('PRECEDENT')+'"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
    end;
end;

//FV1 - 23/07/2015 - FS#1481 - LEVM : Problème avec la modif code article et la fiche Ressource
procedure TChangeCodeArt.TraiteLesRessources;
var Indice,NB : integer;
    TOBArt : TOB;
    Req : string;
begin

  for Indice := 0 to TOBArticle.detail.count -1 do
  begin
    TOBArt := TOBArticle.detail[Indice];
    Req := 'UPDATE RESSOURCE ';
    Req := Req + '  SET ARS_ARTICLE ="' + TOBArt.GetValue('GA_ARTICLE') + '"';
    REq := Req + 'WHERE ARS_ARTICLE ="' + TOBArt.GetValue('PRECEDENT')  + '"';
    Nb := ExecuteSql (Req);
    if Nb < 0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
  end;

end;


{$IFDEF GPAO}
{ TWInitUniteArticle }
constructor TWInitUniteArticle.Create;
begin
	Inherited Create;
  TobArticle := Tob.Create('ARTICLE', nil, -1);
  Valid := False;
  { Pour passer l'objet à la Tof de visu }
  wListInitArticle.Add( Self );
  ListIndex := wListInitArticle.Count - 1;
end;

destructor TWInitUniteArticle.Destroy;
begin
  inherited;
  if (ListIndex <> -1) and (wListInitArticle.Count > 0) and (ListIndex <= wListInitArticle.Count -1) then wListInitArticle.Delete(ListIndex);
  TobArticle.Free;
end;

procedure TWInitUniteArticle.MakeTobFromDS(DS: tDataSet);
begin
	if DS <> nil then wMakeTobFromDS(DS, TobArticle);
end;

procedure TWInitUniteArticle.OpenForm;
begin
	AglLanceFiche('W', 'WARTICLEUNITE_VIE', '', '', 'OBJECTINDEX=' + IntToStr(ListIndex));
end;

function GetWInitArticle(iObject: Integer): TWInitUniteArticle;
begin
  if (iObject >= 0) and (iObject <= wListInitArticle.Count -1) then
    Result := tWInitUniteArticle(wListInitArticle[iObject])
  else
    Result := nil;
end;

procedure tWInitUniteArticle.PutTobInDS(DS: tDataSet);
{ Copie les champs modifiés dans le DS }
begin
  if (DS <> nil) then
  begin
    if (TobArticle.IsOneModifie(False)) then
    begin
      if not(DS.State in [dsInsert,dsEdit]) then DS.edit;
{$IFDEF EAGLCLIENT}
      WSetDSFromTob( TobArticle, DS); {PMJEAGL}
{$ELSE}
      TobArticle.PutDataSet(DS);
{$ENDIF}
    end;
  end;
end;
{$ENDIF }

{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 04/12/2002
Description .. : Renvoie une liste des NaturePieceG qui joue sur le
Suite ........ : compteur passé en paramètre
*****************************************************************}
function wGetListeNaturePiecG(Compteur: string): string;
var
  iTob  : integer;
  sql   : string;
  TobGPP: tob;
begin
  sql := 'SELECT GPP_NATUREPIECEG, GPP_QTEPLUS'
      + ' FROM PARPIECE'
      ;
  TobGPP := Tob.Create('GPP', nil, -1);
  try
    if TobGPP.LoadDetailDBFromSql('PARPIECE', sql) then
    begin
      for iTob := 0 to TobGPP.Detail.Count - 1 do
      begin
        if Pos(Compteur, TobGPP.detail[iTob].getValue('GPP_QTEPLUS')) > 0 then
        begin
          Result := Result + iif(Result <> '', '~', '') + TobGPP.detail[iTob].getValue('GPP_NATUREPIECEG');
        end;
      end;
    end;
  finally
    TobGPP.free;
  end;
end;

{$IFDEF GPAO}    
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 28/08/2003
Modifié le ... :   /  /
Description .. : Renvoi le code unité de la fiche article en fonction du flux
Suite ........ : demandé
Mots clefs ... :
*****************************************************************}
function GetMesureFromFlux(Article, Flux: string): string;
begin
  if      Flux = 'STO' then Result := wGetFieldFromGA('GA_QUALIFUNITESTO', Article)
  else if Flux = 'VTE' then Result := wGetFieldFromGA('GA_UNITEQTEVTE'   , Article)
  else if Flux = 'ACH' then Result := wGetFieldFromGA('GA_UNITEQTEACH'   , Article)
  else if Flux = 'PRO' then Result := wGetFieldFromGA('GA_UNITEPROD'     , Article)
  else if Flux = 'CON' then Result := wGetFieldFromGA('GA_UNITECONSO'    , Article)
end;
{$ENDIF}


{$IFDEF GPAO}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 20/02/2003
Description .. : Test existence d'un CODE ARTICLE depuis le script
*****************************************************************}
function AglExisteCodeArticle(Parms : array of variant; nb : integer): Variant;
begin
  Result := wExistGA(Parms[0], Parms[1]);
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 03/04/2003
Description .. : Retourne le Code Article pour une Référence client
*****************************************************************}
function GetArticleFromReferenceClient(RefArticleTiers: String): String;
var
  Q: TQuery;
begin
  Result := '';
  Q := OpenSQL('SELECT GAT_ARTICLE FROM ARTICLETIERS WHERE GAT_REFARTICLE="'+ RefArticleTiers + '"', true,-1,'',true);
  try
    if not Q.Eof then
      Result := Q.FindField('GAT_ARTICLE').asString;
  finally
    Ferme(Q);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 03/04/2003
Description .. : Retourne le Code ArticleTiers à partir d'un article et d'un Tiers
*****************************************************************}
function GetArticleTiers(Article, RefTiers: String): String;
var
  Q: TQuery;
begin
  Result := '';
  Q := OpenSQL('SELECT GAT_REFARTICLE FROM ARTICLETIERS WHERE (GAT_ARTICLE="'+ Article + '" AND GAT_REFTIERS="' + RefTiers + '")', true,-1,'',true);
  try
    if not Q.Eof then
      Result := Q.FindField('GAT_REFARTICLE').asString;
  finally
    Ferme(Q);
  end;
end;

function CreateArticleTiers(TobGAT: Tob): boolean;
var
  TomGAT: Tom;
begin
  Result := False;
  if Assigned(TobGAT) then
  begin
    TomGAT := CreateTOM('ARTICLETIERS', nil, false, True);
    try
      TobGAT.AddChampSupValeur('IKC', 'C', false);
      TomGAT.VerifTOB(TobGAT);
      Result := (TobGAT.GetValue('Error') <> '');
      if Result then
        TobGAT.InsertDB(nil);
    finally
      TomGAT.Free;
    end;
  end;
end;

{$IFDEF GPAO}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 08/04/2003
Description .. : Retourne le code article correspondant
Suite ........ : à un code barre article
*****************************************************************}
function GetArticleFromCB(cb: String): String;
var
  TobArticle: Tob;
begin
  Result := '';
  if cb <> '' then
  begin
    TobArticle := Tob.Create('ARTICLE', nil, -1);
    try
      if wGetTobGA('GA_ARTICLE', 'GA_CODEBARRE="' + cb + '"', TobArticle, True) then
        Result := TobArticle.GetValue('GA_ARTICLE');
    finally
      TobArticle.Free;
    end;
  end;
end;
{$ENDIF}

function ExistArticleTiers(Article, Tiers: String): Boolean; 
var
  Q: TQuery;
begin
  Result := False;
  Q := OpenSQL('SELECT GAT_REFARTICLE FROM ARTICLETIERS WHERE (GAT_ARTICLE="'+ Article + '" AND GAT_REFTIERS="' + Tiers + '")', true,-1,'',true);
  try
    Result := not Q.Eof;
  finally
    Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 22/04/2003
Description .. : Fabrique un where sur la table ARTICLETIERS en utilisant
Suite ........ : la référence ArticleTiers
*****************************************************************}
function WhereRefArticleTiers(sTiers, sRefArticle: String): String;
begin
  Result := 'GAT_REFARTICLE="' + sRefArticle + '" AND GAT_REFTIERS="' + sTiers + '"';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 22/04/2003
Description .. : Forme un Where pour la table ARTICLETIERS
*****************************************************************}
function WhereArticleTiers(sTiers, sArticle: String): String;
begin
  Result := 'GAT_ARTICLE="' + sArticle + '" AND GAT_REFTIERS="' + sTiers + '"';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 01/04/2003
Description .. : Récupère un champ dans la table ARTICLETIERS
*****************************************************************}
{$IFDEF GPAO}
function GetFieldFromArticleTiers(Field, sTiers, sArticle: string; ByRefArticle: Boolean = False): Variant;
begin
  if not ByRefArticle then
    Result := wGetSqlFieldValue(Field, 'ARTICLETIERS', WhereArticleTiers(sTiers, sArticle))
  else
    Result := wGetSqlFieldValue(Field, 'ARTICLETIERS', WhereRefArticleTiers(sTiers, sArticle));
end;
{$ENDIF}

{$IFDEF GPAO}
procedure InitAglFunc;
begin
  RegisterAglFunc('ExisteCodeArticle', False , 2, AglExisteCodeArticle);
end;
{$ENDIF}


{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 13/08/2003
Modifié le ... : 13/08/2003
Description .. : Conversion d'une quantité d'une unité à une autre, soit
Suite ........ : entre deux unitées quelconques de la table MEA, soit entre
Suite ........ : deux unitées prédéfinies d'un article (Flux), par exemple, de
Suite ........ : l'unité de stock en unité d'achat... (valeurs possibles : 'STO',
Suite ........ : 'VTE', 'ACH', 'PRO' et 'CON')
Mots clefs ... : UNITÉ, CONVERSION, MEA, QUANTITÉ, FLUX
*****************************************************************}
function ConvUnite(Qte : Double; sUdepart, sUArrive, sArticle, sTypeDepart, sTypeArrive : String) : Double;
var Newvalue : Double;
    QuotiteDepart, QuotiteArrive : Double;

Const sFlux = 'STO_VTE_ACH_PRO_CON';

begin      //Voir pour utiliser des fonctions de wConversionUnite.
  Newvalue := Qte;
  if (sTypedepart <> '') and (sTypeArrive <> '') and (sArticle <> '') then   // On travaille avec des Flux
  begin
    // on récupére les unités correspondantes aux flux
    if Pos(sTypeDepart, sFlux) > 0 then
      sUDepart := GetMesureFromFlux(sArticle, sTypeDepart);
    if Pos(sTypeArrive, sFlux) > 0 then
      sUArrive := GetMesureFromFlux(sArticle, sTypeArrive);
  end;

  if (sUdepart <> '') and (sUArrive <> '') then                       // On travaille avec des codes d'unités.
  begin
    if (sUDepart <> sUarrive) then // Si on est sur la même unité, pas de conversions nécessaires. Newvalue = Qte.
    begin
      QuotiteDepart := Double(wGetSqlFieldValue('GME_QUOTITE', 'MEA', 'GME_MESURE = "'+sUdepart+'"'));
      QuotiteArrive := Double(wGetSqlFieldValue('GME_QUOTITE', 'MEA', 'GME_MESURE = "'+sUArrive+'"'));
      If QuotiteDepart = 0 then QuotiteDepart := 1;
      If QuotiteArrive = 0 then QuotiteArrive := 1;
      NewValue := Qte * QuotiteDepart / QuotiteArrive;
    end;
  end;
  Result := NewValue;
end;
{$ENDIF}

{$IFDEF BTP}
procedure CreationEnteteOuv(CodeArt,Libelle,Domaine : String);
var
   TOBEnt : Tob;
   QQ : Tquery;
   Req : String;
begin
  TOBEnt := TOB.create ('NOMENENT',nil,-1);
  Req := 'SELECT GNE_NOMENCLATURE,GNE_QTEDUDETAIL FROM NOMENENT WHERE GNE_ARTICLE="'+ CodeArt +'"';
  QQ := Opensql (Req,true,-1,'',true);
  TRY
    if QQ.EOF then
      begin
      TOBEnt.PutValue ('GNE_ARTICLE',CodeArt);
      TOBEnt.PutValue ('GNE_LIBELLE',Libelle);
      TOBEnt.putValue ('GNE_NOMENCLATURE',copy(CodeArt,0,18));
      TOBEnt.putValue ('GNE_DOMAINE',Domaine);
      TOBENt.PutValue ('GNE_QTEDUDETAIL',strtoint('1'));
      TOBENt.PutValue ('GNE_DATEMODIF', V_PGI.DateEntree);
      TOBENt.PutValue ('GNE_DATECREATION', V_PGI.DateEntree);
      TOBEnt.InsertDB(nil,true);
      end
    else
      begin
      // Mise à jour du libellé de la déclinaison principale (même code)
      Req := 'UPDATE NOMENENT SET GNE_LIBELLE="'+Libelle+'",GNE_DOMAINE="'+Domaine+'" WHERE GNE_NOMENCLATURE="'+ copy(CodeArt,0,18) +'"';
      ExecuteSql (Req);
      end;
  FINALLY
    TOBEnt.Free;
    Ferme(QQ);
  END;
end;

procedure BTPDuplicOuvrage(Source, Destination : string);

  procedure ChargeOuvrageSource (Source : String;TOBNOMEN : TOB);
  var Req : String;
      QQ : TQuery;
      Indice : integer;
      ThisTOB : TOB;
  begin
    QQ := OpenSql ('SELECT * FROM NOMENENT WHERE GNE_ARTICLE="'+Source+'"',True,-1,'',true);
    TRY
      if not QQ.eof then
      begin
        TOBNOMEN.LoadDetailDB ('NOMENENT','','',QQ,false,true);
      end;
    FINALLY
      Ferme (QQ);
    END;
    if TOBNOMEN.detail.count > 0 then
    begin
      // Chargement du détail de chaque ouvrage
      for Indice := 0 to TOBNomen.detail.count -1 do
      begin
        ThisTOB := TOBNOMEN.detail[Indice];
        QQ := OpenSql ('SELECT * FROM NOMENLIG WHERE GNL_NOMENCLATURE="'+ThisTOB.GetValue('GNE_NOMENCLATURE')+'"',true,-1,'',true);
        TRY
          if not QQ.Eof then ThisTOB.LoadDetailDB ('NOMENLIG','','',QQ,false,true);
        FINALLY
          Ferme (QQ);
        END;
      end;
    end;
  end;

var TOBNOMEN,ThisTOB : TOB;
    Indice,IndFille : integer;
    OldCode : string;
    NewCode : string;
    AncienCode  : String;
    WithError : boolean;
begin
  WithError := false;
  TOBNOMEN := TOB.Create ('LES NOMENENT',nil,-1);
  TRY
    ChargeOuvrageSource (Source,TOBNOMEN);
    if TOBNomen.detail.count > 0 then
    begin
      for Indice := 0 to TOBNomen.detail.count -1 do
      begin
        ThisTOB := TOBNOMEN.detail[Indice];
        OldCode := ThisTOB.GetValue('GNE_NOMENCLATURE');
// Ahhhhhhhhhhhhhhh quelle horreur
(*
        if Indice = 0 then // Sur de l'unicite bicose le code article est une clef unique
          NewCode := copy(Destination,1,pos('X',Destination)-1)
        else
          NewCode := AffecteCodeNomen (Destination,copy(Destination,1,pos('X',Destination)-1),TOBNomen);
*)
// d'ou la correction suivante
        if Indice = 0 then // Sur de l'unicite bicose le code article est une clef unique
          NewCode := Trim(copy(Destination,1,18))
        else
          NewCode := AffecteCodeNomen (Destination,Trim(copy(Destination,1,18)),TOBNomen);
//          
        if NewCode <> ''  then
        BEGIN
          AncienCode := ThisTob.GetValue('GNE_ARTICLE');
          ThisTOB.putValue('GNE_ARTICLE',Destination);
          ThisTOB.putValue('GNE_NOMENCLATURE',NewCode);
          for IndFille := 0 to ThisTOB.detail.count -1 do
          begin
            ThisTOB.detail[IndFille].putvalue ('GNL_NOMENCLATURE',newCode);
          end;
          ThisTOB.addChampSupValeur ('OKOK','X');
{$IFDEF BTP}
          //Gestion de la copie du métré et des variables
          DuplicationArticleMetre('', AncienCode, Destination);
          {$ENDIF}
        END else
        begin
          ThisTOB.addChampSupValeur ('OKOK','-');
          WithError := True;
        end;
      end;
      IndFille := 0;
      repeat
        if TOBNomen.detail[Indfille].GetValue('OKOK') <> 'X' then TOBNomen.free else inc(indFille);
      until IndFille > TOBNomen.detail.count -1;
      TOBNomen.InsertDB (nil,true);
    end;
  FINALLY
    FreeAndNil (TOBNOMEN);
    if WithError then PGIBox (traduirememoire('Des erreurs se sont produites durant la duplication des déclinaisons'));
  END;
end;

Function isExistsArticle (Article : string) : boolean;
begin
	result := ExisteSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE ="' + Article + '"');
end;

procedure InitChampsSupArticle (TOBA : TOB);
begin
	if not TOBA.FieldExists ('GA2_SOUSFAMTARART') then
  begin
  	TOBA.AddChampSupValeur ('GA2_SOUSFAMTARART','');
  end else
  begin
  	if (VarIsNull(TOBA.GetValue ('GA2_SOUSFAMTARART'))) or
    	 (VarAsType(TOBA.GetValue ('GA2_SOUSFAMTARART'),VarString)=#0) then
    begin
      TOBA.PutValue ('GA2_SOUSFAMTARART','');
    end;
  end;
	if not TOBA.FieldExists ('BNP_TYPERESSOURCE') then
  begin
  	TOBA.AddChampSupValeur ('BNP_TYPERESSOURCE','');
  end else
  begin
  	if (VarIsNull(TOBA.GetValue ('BNP_TYPERESSOURCE'))) or
    	 (VarAsType(TOBA.GetValue ('BNP_TYPERESSOURCE'),VarString)=#0) then
    begin
      TOBA.PutValue ('BNP_TYPERESSOURCE','');
    end;
  end;
	if not TOBA.FieldExists ('BNP_LIBELLE') then
  begin
  	TOBA.AddChampSupValeur ('BNP_LIBELLE','');
  end else
  begin
  	if (VarIsNull(TOBA.GetValue ('BNP_LIBELLE'))) or
    	 (VarAsType(TOBA.GetValue ('BNP_LIBELLE'),VarString)=#0) then
    begin
      TOBA.PutValue ('BNP_LIBELLE','');
    end;
  end;
  //
  TOBA.AddChampSupValeur('REFARTSAISIE', '', False);
  TOBA.AddChampSupvaleur('REFARTBARRE', '', False);
  TOBA.AddChampSupValeur('REFARTTIERS', '', False);
  TOBA.AddChampSupValeur('SUPPRIME', '-', False);
  TOBA.AddChampSupValeur('UTILISE', '-', False);
  //
end;

{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 26/09/2002
Description .. : Retourne vrai si le nom de champ est un code article
*****************************************************************}
function wIsSuffixeCodeArticle(Const FieldName: String): Boolean;
begin
  Result := (Pos('CODEARTICLE', FieldName) > 0) or (Pos('CODECOMPOSANT', FieldName) > 0);
end;

procedure  recalculCoefMarg (TOBRef : TOB);
var CoefcalcPV,PrixRef,TarifFournisseur,QuotiteAchat,QuotiteVente : double;
    CodeBasePV : string;
    QQUALIFACH,QQUALIFVTE : TQuery;
    PV : double;
begin

if TOBRef.GetValue('GA_CALCAUTOHT')<>'X' then exit;
PV := TOBRef.GetValue('GA_PVHT') ;

CoefCalcPV:=TOBRef.GetValue('GA_COEFCALCHT');
CodeBasePV:=TOBRef.GetValue('GA_CALCPRIXHT') ;
if (CoefCalcPV<>0) and (CodeBasePV<>'') then
   BEGIN
   if CodeBasePV='DPA' then PrixRef:=TOBRef.GetValue('GA_DPA')
   else if CodeBasePV='DPR' then PrixRef:=TOBRef.GetValue('GA_DPR')
   else if CodeBasePV='PAA' then PrixRef:=TobRef.GetValue('GA_PAHT')
   else if CodeBasePV='PRA' then PrixRef:=TobRef.GetValue('GA_PRHT')
   else if CodeBasePV='PMA' then PrixRef:=TobRef.GetValue('GA_PMAP')
   else if CodeBasePV='PMR' then PrixRef:=TobRef.GetValue('GA_PMRP')
   else if CodeBasePV='FOU' then PrixRef := TobRef.GetValue('GA_DPA');
  if PrixRef <> 0 then TOBRef.putValue('GA_COEFCALCHT',arrondi(PV/PrixRef,4))
  								else TOBRef.putValue('GA_COEFCALCHT',0);
  END;
end;

procedure  recalculCoef(TOBRef : TOB);
var CoefcalcPV : Double;
    PrixRef    : Double;
    PV         : double;
    //
    CodeBasePV : string;
begin

  if TOBRef.GetValue('GA_CALCAUTOTTC')<>'X' then exit;

  PV := TOBRef.GetValue('GA_PVTTC') ;

  CoefCalcPV:=TOBRef.GetValue('GA_COEFCALCTTC');
  CodeBasePV:=TOBRef.GetValue('GA_CALCPRIXTTC') ;

  if (CoefCalcPV<>0) and (CodeBasePV<>'') then
  BEGIN
    if CodeBasePV='DPA'      then PrixRef:=TOBRef.GetValue('GA_DPA')
    else if CodeBasePV='DPR' then PrixRef:=TOBRef.GetValue('GA_DPR')
    else if CodeBasePV='PAA' then PrixRef:=TobRef.GetValue('GA_PAHT')
    else if CodeBasePV='PRA' then PrixRef:=TobRef.GetValue('GA_PRHT')
    else if CodeBasePV='PMA' then PrixRef:=TobRef.GetValue('GA_PMAP')
    else if CodeBasePV='PMR' then PrixRef:=TobRef.GetValue('GA_PMRP')
    else if CodeBasePV='HT'  then PrixRef:=TobRef.GetValue('GA_PVHT')
    else if CodeBasePV='FOU' then PrixRef:=TobRef.GetValue('GA_DPA');
    //
    if PrixRef <> 0 then TOBRef.putValue('GA_COEFCALCTTC',arrondi(PV/PrixRef,4))
    else TOBRef.putValue('GA_COEFCALCTTC',0);
  END;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 02/03/2004
Modifié le ... :   /  /
Description .. : Retourne le statut de l'article demandé
Mots clefs ... :
*****************************************************************}
function GetStatutArt(Const Article: string): string;
const
  LastArticle: string = '';
  LastResult : String = '';
begin
  Result := LastResult;
  try
    if Article <> LastArticle then
      Result := wGetFieldFromGA('GA_STATUTART', Article)
  finally
    LastArticle := Article;
    LastResult  := Result;
  end;
end;

Function wCreateArticleFromTob(Const Article: String; TheTob: Tob): Boolean;
var
  iChamp: Integer;
  FieldName: String;
  Exist, IsValid: Boolean;
  TobArticle, TobZonelibre: Tob;
  TomArticle: Tom;
//  BBI Correction de bug
  sError, stArgs, stPref: String;
//  BBI Fin Correction de bug
begin
  sError := '';
  { Contrôle existence de l'article }
  if not wExistGA(Article) then
  begin
   	if TheTob <> nil then
    begin
      { Crée une TobArticle et contrôle l'existence des champs de la tob passée en paramêtre }
      TobArticle := Tob.Create('ARTICLE', nil, -1);
      TobZonelibre := TOB.Create('ARTICLECOMPL', Nil, -1);
      stPref := TableToPrefixe('ARTICLE');
      try
        { Recopie la tob à mettre à jour dans la tob article }
        iChamp := 999; Exist := True;
        while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) and Exist do
        begin
          Inc(iChamp);
//  BBI Correction de bug
          FieldName := TheTob.GetNomChamp(iChamp);
          { Vérifie si le champ fait partie de la table article }
          if Copy(FieldName, 0, Pos('_', FieldName) - 1) = stPref then
            Exist := TobArticle.FieldExists(FieldName)
          else
            TobZoneLibre.PutValue(TheTob.GetNomChamp(iChamp), TheTob.GetValue(TheTob.GetNomChamp(iChamp)));
//  BBI Fin Correction de bug
        end;
        if Exist then
        begin
          {$IFDEF PGISIDE}
          stArgs := 'ORIGINE=PGISIDE;ARTICLE=' + Article;
          if TheTob.FieldExists('GA_TYPEARTICLE') then
            stArgs := stArgs + ';TYPEARTICLE=' + TheTob.GetString('GA_TYPEARTICLE');
          if TheTob.FieldExists('GA_TYPENOMENC') then
            stArgs := stArgs + ';TYPENOMENC=' + TheTob.GetString('GA_TYPENOMENC');
          {$ELSE PGISIDE}
          stArgs := '';
          {$ENDIF PGISIDE}
          { Vérifie les données en passant par une TomArticle }
          TomArticle := CreateTOM('ARTICLE', nil, false, true);
          TomArticle.Argument(stArgs);
          TomArticle.InitTOB(TobArticle);
          { Recopie la tob à mettre à jour dans la tob article }
          iChamp := 999;
          while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) do
          begin
            Inc(iChamp);
            FieldName := TheTob.GetNomChamp(iChamp);
//  BBI Correction de bug
            if (TobArticle.FieldExists(FieldName) and (Copy(FieldName, 0, Pos('_', FieldName) - 1) = stPref)) then
              TobArticle.PutValue(FieldName, TheTob.GetValue(FieldName));
//  BBI Fin Correction de bug
          end;
          try
            TobArticle.AddChampSupValeur('IKC', 'C', False);
            IsValid := TomArticle.VerifTOB(TobArticle);
            sError := TomArticle.LastErrorMsg;
          finally
            TomArticle.Free;
          end;
          if IsValid then
          begin
            try
              TobArticle.InsertDB(nil, False); { Enregistre la TobArticle }
            except
              on E: Exception do
              begin
              {$IFDEF PGISIDE}
                MAJJnalEvent('PSE', 'ERR', TraduireMemoire('Mise à jour de la table ARTICLE'), E.Message);
              {$ENDIF PGISIDE}
                sError := E.Message;
              end;
            end;
            try
              TobZoneLibre.PutValue('GA2_ARTICLE', TobArticle.GetValue('GA_ARTICLE'));
              TobZoneLibre.InsertDB(nil, False);
            except
              on E: Exception do
              begin
                {$IFDEF PGISIDE}
                MAJJnalEvent('PSE', 'ERR', TraduireMemoire('Mise à jour de la table ARTICLECOMPL'), E.Message);
                {$ENDIF PGISIDE}
                sError := E.Message;
              end;
            end;
            TheTob.Dupliquer(TobArticle, False, True);
          end;
        end
        else
          sError := TraduireMemoire('Le champ : ') + FieldName + TraduireMemoire(' n''existe pas dans la table : ') + 'ARTICLE';
      finally
        TobArticle.Free;
      end;
    end
    else
      sError := TraduireMemoire('Paramètres d''appel incorrect (Tob non renseignée)');
  end
  else if not wUpDateArticleFromTob(Article, TheTob) and TheTob.FieldExists('Error') then
    sError := TheTob.GetString('Error');

  Result := sError = '';
  if sError <> '' then
    TheTob.AddChampSupValeur('Error', sError);
end;

Function wUpDateArticleFromTob(Article: String; TheTob: Tob): Boolean;
var
  iChamp: Integer;
  FieldName: String;
  Exist, IsValid: Boolean;
  TobArticle: Tob;
  TomArticle: Tom;
//  BBI Correction de bug
  sError, stPref : String;
//  BBI Fin Correction de bug
begin
	sError := '';
  Result := False;
  { X Dans le code article }
  if not xSetCodeArticle(Article) then
  begin
    sError := TraduireMemoire('Code article incorrect') + ' : ' + Article;
    EXIT;
  end;
  { Contrôle existence de l'article }
	if wExistGA(Article) then
  begin
   	if TheTob <> nil then
    begin
      { Crée une TobArticle et contrôle l'existence des champs de la tob passée en paramêtre }
      TobArticle := Tob.Create('ARTICLE', nil, -1);
//  BBI Correction de bug
      stPref := TableToPrefixe('ARTICLE');
//  BBI Fin Correction de bug
      try
        { Charge la tob avec les données de l''article }
        wGetTobGA('*', whereGA(Article), TobArticle, True);
        TobArticle.SetAllModifie(False);
        { Recopie la tob à mettre à jour dans la tob article }
        iChamp := 999; Exist := True;
        while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) and Exist do
        begin
          Inc(iChamp);
//  BBI Correction de bug
          FieldName := TheTob.GetNomChamp(iChamp);
          { Vérifie si le champ fait partie de la table article }
          Exist := (TobArticle.FieldExists(FieldName) and (Copy(FieldName, 0, Pos('_', FieldName) - 1) = stPref));
//  BBI Fin Correction de bug
          if Exist then TobArticle.PutValue(FieldName, TheTob.GetValue(FieldName));
        end;
        if Exist then
        begin
          { Vérifie les données en passant par une TomArticle }
          TomArticle := CreateTOM('ARTICLE', nil, false, true);
          try
            TomArticle.LoadBufferAvantModif(TobArticle);
            TobArticle.AddChampSupValeur('IKC', 'M', False);
            IsValid := TomArticle.VerifTOB(TobArticle);
          finally
            TomArticle.Free;
          end;
          if IsValid then
          begin
            TobArticle.UpdateDB(False) { Enregistre la TobArticle }
          end
          else
            sError := TraduireMemoire('Les données de l''article : ') + Article + TraduireMemoire(' ne sont pas valides');
        end
        else
          sError := TraduireMemoire('Le champ : ') + FieldName + TraduireMemoire(' n''existe pas dans la table : ') + 'ARTICLE';
       finally
        TobArticle.Free;
       end;
    end
    else
      sError := TraduireMemoire('Paramètres d''appel incorrect (Tob non renseignée)');
  end
  else
    sError := TraduireMemoire('L''article : ') + Article + TraduireMemoire(' est inconnu');

  Result := sError = '';
  if sError <> '' then
    TheTob.AddChampSupValeur('Error', sError);
end;

Procedure DuplicationArticleMetre(TypeArticle, CodeArticle, Duplicart : String);
Var FileOrigine     : string;
    FileDestination : string;
    ToblV           : TOB;
    TobDLV          : TOB;
    I               : Integer;
begin

  MetreArticle    := TMetreArt.CreateArt(TypeArticle, CodeArticle);
  FileOrigine     := MetreArticle.fRepMetre + '\Bibliotheque\' + MetreArticle.FormatArticle(CodeArticle) + '.xlsx';
  FileDestination := MetreArticle.fRepMetre + '\Bibliotheque\' + MetreArticle.FormatArticle(DuplicArt) + '.xlsx';

  //Si l'article dupliqué a un fichier Xlsx on demande siu l'on veut le dupliquer...
  if FileExists(FileOrigine) then
  begin
    if PgiAsk('L''article dupliqué est associé à un fichier métré. Voulez-vous le dupliquer également ?', 'Duplication métrés')=mrYes then
    begin
      MetreArticle.CopieLocaltoServeur(FileOrigine, FileDestination);
    end;
  end;

  //Si l'article dupliqué dispose de variables spécifique on demande si on veut le dupliquer...
  MetreArticle.fArticle := CodeArticle;
  MetreArticle.ChargementVariablesArt(CodeArticle);

  if MetreArticle.TTobVariables.detail.count > 0 then
  begin
    if PgiAsk('L''article dupliqué est associé à des variables. Voulez-vous les dupliquer également ?', 'Duplication variables')=mrYes then
    begin
      for I:=0 to MetreArticle.TTobVariables.Detail.count - 1 do
      begin
        TobLV := MetreArticle.TTobVariables.Detail[I];
        TobDLV:= Tob.Create('BVARIABLES',MetreArticle.TTobVariables, -1);
        TobDLV.dupliquer(ToblV, True,True);
        TobDLV.PutValue('BVA_ARTICLE', DuplicArt);
      end;
      MetreArticle.TTobVariables.SetAllModifie(true);
      MetreArticle.TTobVariables.InsertOrUpdateDB(true);
    end;
  end;

  //On vérifie si l'article dupliqué ne comporte pas de variables.
  FreeAndNil(MetreArticle);

end;

{$IFDEF GPAO}
Initialization
  wListInitArticle := TList.Create;
  InitAglFunc;
{$ENDIF}

{$IFDEF GPAO}
Finalization
	wListInitArticle.Free;
{$ENDIF}


end.
