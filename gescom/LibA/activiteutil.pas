unit ActiviteUtil;

interface

Uses HEnt1, HCtrls, UTOB, Ent1,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     DB,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
{$ENDIF}

     LookUp,
     Controls,
     ComCtrls,
     StdCtrls,
     ExtCtrls,
     SysUtils,
     Dialogs,
     SaisUtil,
     UtilPGI,
     AGLInit,
     EntGC,

{$IFDEF AFFAIRE}
     affaireutil,
{$ENDIF}

{$IFDEF BTP}
	 CalcOleGenericBTP,
{$ENDIF}

     FactUtil,uafo_ressource,HMsgBox,
     UtilArticle, AFUtilArticle, UtilRessource,
     paramsoc, confidentaffaire, HPanel,
     UtofafTiers_Rech,UtofAfArticle_mul,GereTobInterne,
     Classes, FactArticle, FactComm;


Const SG_Curseur   : integer = 0 ;
      SG_Jour      : integer = 1 ;
      SG_Tiers     : integer = 3 ;
      SG_Affaire   : integer = 4 ;
      SG_Aff0      : integer = 9 ;
      SG_Aff1      : integer = 9 ;
      SG_Aff2      : integer = 9 ;
      SG_Aff3      : integer = 9 ;
      SG_Avenant   : integer = 9 ;
      SG_Ressource : integer = 5 ;
      SG_Article   : integer = 6 ;
      SG_Unite     : integer = 7 ;
      SG_Qte       : integer = 8 ;
      SG_PUCH      : integer = 9 ;
      SG_PV        : integer = 9 ;
      SG_TOTPRCH   : integer = 9 ;
      SG_Desc      : integer = 9 ;
      SG_Lib       : integer = 9 ;
      SG_TotV      : integer = 9 ;
      SG_TypA      : integer = 9 ;
      SG_ActRep    : integer = 9 ;
      SG_TypeHeure : integer = 9 ;
      SG_MntRemise : integer = 9 ;
      SG_MontantTVA: integer = 9 ;

Const NbRowsInit = 55 ;
      NbRowsPlus = 25 ;
//  0 :  1ere semaine = contient le 1er janvier
//  1 :  1ere semaine = contient 1er jeudi
//  2 :  1ere semaine = contient 1er lundi
      PremiereSemaine = 2;

Type T_Sai = (tsaRess, tsaClient) ;
Type T_Acc = (tacTemps, tacFrais, tacFourn, tacGlobal) ;
Type T_TypeBlocageActivite = (tbaDatesAct, tbaAssistant, tbaAffaire, tbaModeConsult) ;
Type T_EnsBlocagesActivite = Set of  T_TypeBlocageActivite;

Type T_TypeBlocage35H = (tbhAnBloc, tbhAnAlert, tbhSemLim1, tbhSemConsLim1, tbhSemLim2, tbhSemConsLim2);
Type T_EnsBlocages35H = Set of  T_TypeBlocage35H;

Type R_CLEACT = RECORD
                TypeSaisie : T_Sai ;   // Saisie par ressource ou par Client/Affaire
                TypeAcces : T_Acc ;    // Accès global, Temps, Frais, Fournitures
                TypeActivite : string ;// Prévisionnel, Réalisé ou Simulation
                Tiers : string ;       // Code du client en cours
                Affaire : string ;     // Code de l'affaire/Mission en cours
                Aff0 : string ;        // Code de l'affaire/Mission en cours 0
                Aff1 : string ;        // Code de l'affaire/Mission en cours 1
                Aff2 : string ;        // Code de l'affaire/Mission en cours 2
                Aff3 : string ;        // Code de l'affaire/Mission en cours 3
                Avenant : string;      // Code avenant de l'affaire
                TypeRess : string ;    // type de la Ressource en cours
                Ressource : string ;   // Code de la Ressource en cours
                Annee : integer;       // Annee courante
                Mois : integer;        // Mois courant
                Semaine : integer;     // Semaine courante
                Jour : integer;        // Jour courant
                Folio : integer;       // folio courant
                Fonction : string;     // fonction de la ressource courante
                END ;

// Fonctions liees à la valorisation
function MajTOBValo(AcdDate:TDateTime; AcTypeAcces:T_Acc; AcsAffaire:string; AcsRessource:string; AcsCodeArticle:string;
                    AcsTypeHeure:string; TOBArticles, TOBAffaires, TobZoom:TOB; AFOAssistants:TAFO_Ressources;
                    bAvecZoom:boolean; AcsValoActPR,AcsValoActPV : string; var AvbPROK, AvbPVOK : boolean):TOB;
function ValoriseLesActivites(TobActs, TOBArticles : TOB; AFOAssistants:TAFO_Ressources; TOBAffaires : TOB; bPR, bPV : boolean):boolean;

// Article
function TrouverCodeArticle(CodeArticle, sTypeArt:string; TOBArticles:TOB):TOB;
Function FindTOBArtSaisAff (TOBArticles:TOB; RefSais : String; AbRef:boolean ) : TOB ;
Procedure ZoomArticle ( RefUnique : String; AcsTypeArticle: String ) ;
Procedure ZoomArticlesD1Affaire ( AcsCodeAffaire: String; AcsCodeArticle: String; AcsTypeArticle: String ; AcbSansPrix:boolean ) ;
function TOBArticlesD1Affaire (AcsCodeAffaire:string; AcsCodeArticle:string; AcsTypeArticle: String):TOB;
function NbArticlesD1Affaire(AcsCodeAffaire:string; var AvsCodeArticle:string;
                                 var AvsLibArticle:string; var AvdPrHT, AvdPvHT:double; var numordre : integer;
                                 var FormuleVar : string; var MntRemise : double; var NumPieceOrig : String):integer; // PL le 06/08/03 : pour CB ajout NumPieceOrig

function TOBArticlesD1AffaireOnyx (AcsCodeAffaire, AcsCodeArticle, AcsTypeArticle : String) : TOB;

// Ressource/assistant
Procedure ZoomRessource ( AcsRessource: String ) ;
procedure ZoomClient (AcsTiers:string);

//Parc/Matériel : Nouveauté 2015
procedure ZoomFamMateriel(FamMat : string);

// Lignes d'activite
Procedure InitialiseLigneAct ( TOBL : TOB; ARow : integer; CleAct : R_CLEACT; Action : TActionFiche; CodeDev : string);
Procedure InitLesColsAct ;
Function EstRempliGCActivite ( GS : THGrid ; Lig : integer; bAvecJour:boolean=true ) : boolean ;
Function FabricWhereTypArt ( TypeArticle : String ) : String ;
function MaxNumLigneActivite(TobLigneAct:TOB):integer;
function TestBlocageDates(AcdDate:TDateTime; AciInitResult:integer):integer;
procedure IntervalleDatesActivite(var AvdDateDebut, AvdDateFin : TDateTime);

// Remplir les TOBs
Function RemplirTOBAssistant ( TOBAssistant:TOB;AFOAssistants:TAFO_Ressources;CodeAssistant : String ) : integer ;
Function RemplirTOBUnites ( TOBListeConv:TOB; CodeUnite: String; TOBUnites : TOB ) : boolean ;
Function RemplirTOBTiersP ( TOBTiers,TOBTierss : TOB ; CodeTiers : String ) : boolean ;
Function RemplirTOBTiersLibP ( TOBTiers, TOBTierss : TOB ; LibTiers : String ) : boolean ;

// Recherches, muls, ...
Function AFGetArticleLookUp ( Control:TControl ; TypeArticle : String;TitreSel : String ) : boolean ;
Function AFGetArticleRecherche ( Control:TControl ; TypeArticle : String; TitreSel : String ) : boolean ;
Function AFGetArticleMul ( Control:TControl ; TypeArticle,TitreSel : String ) : boolean ;
Function GetAssistantRecherche ( Control:TControl ; TitreSel : String ; bInit : boolean) : boolean ;
Function GetAssistantLookUp ( TH : TControl; AcsTitre:string ) : boolean ;
Function GetUnite ( TH : TControl; AcsTitre:string; AcType:T_Acc; AcsTypeArt:string) : boolean ;
Function GetAssistantMul ( Control:TControl ; TitreSel : String ; bInit : boolean) : boolean ;
Function GetTiersLookUp ( Control:TControl ; TitreSel : String  ) : boolean ;
Function GetTiersMul ( Control:TControl ; TitreSel : String ; AcbRechParCode:boolean; var AvsCode:string ) : boolean ;
Function GetTiersRechercheAF ( Control:TControl ; TitreSel : String ; AcbRechParCode:boolean; var AvsCode:string) : boolean ;
Function GetAffaire ( TH : THCritMaskEdit; AcsTitre:string; AcsCode:string ) : boolean ;
Function GetTypeHeure ( TH : TControl; AcsTitre:string) : boolean ;

// Fonctions de conversions d'unités
function ListeUnitesConversion(AcsCodeAConvertir:string):TOB;
function IsConvertible(AcsCodeSource, AcsCodeCible:string):boolean;
// Heures supplémentaires
Function  GetTauxHSup (stTaux : String; Var Tx1,Tx2 : Double):Boolean;
Function  TauxHSupFromCode (codeTT : string; Var Tx1,Tx2 : Double): Boolean;

Function TraiteObjActivite ( Num : integer; PRien : THPanel): Boolean;


// PL le 23/05/02 : réparation de la fonction
// Temporaire
//function PremierJourSemaineAGL540i(AciSemaine:integer; AciAnnee:integer):TDateTime;
//function PremierJourSemaineTempo(AciSemaine:integer; AciAnnee:integer):TDateTime;
//function NumSemaine1(today: Tdatetime): word;
//Function  NumSemaine2( D : TDateTime) : Integer ;
////////////////////////
{$IFDEF AGLINF545}
function PremierJourSemaineTempo(AciSemaine:integer; AciAnnee:integer):TDateTime;
{$ELSE}
{$ENDIF}
//Function  MyNumSemaine( D : TDateTime ; WithSemaine53 : boolean = true) : Integer ;

// Fonctions de gestion de la ligne unique dans l'activite
function LaTobDeLaLigneUnique (sTypeActivite, sAffaire, sNumLigneUnique : string; bRestreint : boolean = true) : TOB;
function MaxNumLigneUniqueActivite (sTypeActivite, sAffaire : string) : integer;
function ProchainPlusNumLigneUniqueActivite (sTypeActivite, sAffaire : string) : integer;
function MinNumLigneUniqueActivite (sTypeActivite, sAffaire : string) : integer;
function ProchainMoinsNumLigneUniqueActivite (sTypeActivite, sAffaire : string) : integer;


function ProchainIndiceAffaires (sTypeActivite, sAffaire : string; ListeAffaires : TStringList; DansLesNegatifs : boolean = false; AvecPreinit : boolean = true) : integer;


implementation

uses UtofRessource_Mul, UtofArticles_Aff;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Function qui lance un lookup sur le tiers à partir d'un control
Suite ........ : donné en entrée (THCritMaskEdit ou THGrid)
Suite ........ : Le lookup affecte le control automatiquement avec le code choisi
Suite ........ : Elle renvoie true s'il y a eu modification et false sinon
Mots clefs ... : TIERS; LOOKUP;  GIGA
*****************************************************************}
Function GetTiersLookUp ( Control:TControl ; TitreSel : String  ) : boolean ;
Var St : String ;
BEGIN
if (Control is THGrid) then
   with THGrid(Control) do
      st := Cells[SG_Tiers, Row]
else
   St:=THCritMaskEdit(Control).Text ;

if (GetParamSoc('SO_AFRECHTIERSPARNOM')=true) then
        LookupList(Control, 'Liste des clients','TIERS','T_LIBELLE','T_TIERS','T_NATUREAUXI="CLI"','T_LIBELLE',False,8)
   else
        LookupList(Control, 'Liste des clients','TIERS','T_TIERS','T_LIBELLE','T_NATUREAUXI="CLI"','T_TIERS',False,8) ;


if (Control is THGrid) then
   with THGrid(Control) do
      Result:=(St<>Cells[SG_Tiers, Row])
else
   Result:=(St<>THCritMaskEdit(Control).Text) ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Fonction qui lance le MUL GA de recherche du tiers à partir
Suite ........ : d'un control (THGrid ou THCritMaskEdit)
Suite ........ : Le code tiers choisi est rangé dans le control
Suite ........ :
Suite ........ : la fonction renvoie true si un code a été sélectionné
Mots clefs ... : TIERS; MUL; GA ; GIGA
*****************************************************************}
Function GetTiersMul ( Control:TControl ; TitreSel : String ; AcbRechParCode:boolean; var AvsCode:string ) : boolean ;
Var CodeTiers, RepTiers : String ;
sTmp, sLibelle: string;
x : integer;
BEGIN
result := true;
RepTiers:='';  AvsCode:=''; sLibelle:='';
if (Control is THGrid) then
   with THGrid(Control) do
      CodeTiers := Cells[Col,Row]
else ;
// PL le 16/12/02 : Pour FIDEA : sur le choix du client en saisie par client/mission, on veut rappeler tous les clients
//if (Control is THCritMaskEdit) then
//   CodeTiers := THCritMaskEdit(Control).Text;

if Not AcbRechParCode then
        // mcd 12/0602 CodeTiers:=AGLLanceFiche('AFF','AFTIERS_RECH','T_NATUREAUXI=CLI; T_LIBELLE=' + CodeTiers,'','NOFILTRE;PARLIBELLE')
        RepTiers:=AFLanceFiche_REch_tiers ('T_NATUREAUXI=CLI; T_LIBELLE=' + CodeTiers,'NOFILTRE;PARLIBELLE')
   else
        // mcd 12/06/02 CodeTiers:=AGLLanceFiche('AFF','AFTIERS_RECH','T_NATUREAUXI=CLI; T_TIERS=' + CodeTiers,'','NOFILTRE');
        RepTiers:=AFLanceFiche_REch_tiers ('T_NATUREAUXI=CLI; T_TIERS=' + CodeTiers,'NOFILTRE;PARLIBELLE');

//PL le 13/09/02 : dans les deux cas on récupère le code et le libellé
sTmp:=(Trim(ReadTokenSt(RepTiers)));
x := Pos('CODE:',sTmp) ;
AvsCode := copy( sTmp, x+5, length(sTmp));
sTmp:=(Trim(ReadTokenSt(RepTiers)));
x := Pos('LIBELLE:',sTmp) ;
sLibelle := copy( sTmp, x+8, length(sTmp));

if Not AcbRechParCode then
    CodeTiers := sLibelle
else
    CodeTiers := AvsCode;

if CodeTiers<>'' then
   BEGIN
   Result:=True ;
   if (Control is THGrid) then
      with THGrid(Control) do
         Cells[Col,Row] := CodeTiers
   else if (Control is THCritMaskEdit) then
         THCritMaskEdit(Control).Text := CodeTiers;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Fonction de séléction du tiers en GA/GI suivant le
Suite ........ : type de recherche avancée ou non
Suite ........ :
Suite ........ : renvoie true s'il y a eu sélection d'un nouveau tiers, false
Suite ........ : sinon
Suite ........ :
Mots clefs ... : TIERS; CLIENT; RECHERCHE ; GIGA
*****************************************************************}
Function GetTiersRechercheAF ( Control:TControl ; TitreSel : String; AcbRechParCode:boolean; var AvsCode:string) : boolean ;
BEGIN
  if GetParamSocSecur('SO_GCRECHTIERSAV', False) then
    Result := GetTiersMul(Control, TitreSel, AcbRechParCode, AvsCode)
  else
    Result := GetTiersLookUp(Control, TitreSel);
END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Fonction de séléction de l'assistant en GA/GI suivant le
Suite ........ : type de recherche avancée ou non
Suite ........ :
Suite ........ : renvoie true s'il y a eu sélection d'un nouvel assistant, false
Suite ........ : sinon
Suite ........ :
Mots clefs ... : RESSOURCE; ASSISTANT; RECHERCHE ; GIGA
*****************************************************************}
Function GetAssistantRecherche (Control:TControl; TitreSel : String; bInit : boolean) : boolean ;
BEGIN
if VH_GC.AFRechResAv then Result:=GetAssistantMul(Control, TitreSel, bInit) else
   Result:=GetAssistantLookUp(Control,TitreSel) ;

END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Function qui lance un lookup sur la ressource à partir d'un control
Suite ........ : donné en entrée (THCritMaskEdit ou THGrid)
Suite ........ : Le lookup affecte le control automatiquement avec le code choisi
Suite ........ : Elle renvoie true s'il y a eu modification et false sinon
Mots clefs ... : RESSOURCE; LOOKUP ; GIGA
*****************************************************************}
Function GetAssistantLookUp ( TH : TControl; AcsTitre:string ) : boolean ;
Var St : String ;
BEGIN
if (TH is THGrid) then
   with THGrid(TH) do
      st := Cells[SG_Ressource, Row]
else
   St:=THCritMaskEdit(TH).Text ;

Result:=LookupList(TH,AcsTitre,'RESSOURCE','ARS_RESSOURCE','ARS_LIBELLE','','ARS_RESSOURCE',False,6) ;
(*
if (TH is THGrid) then
   with THGrid(TH) do
      Result:=(St<>Cells[SG_Ressource, Row])
else
   Result:=(St<>THCritMaskEdit(TH).Text) ;
   *)
END ;


{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Function qui lance un lookup sur l'affaire à partir d'un THCritMaskEdit
Suite ........ : Le lookup affecte le THCritMaskEdit automatiquement avec le code choisi
Suite ........ : Elle renvoie true s'il y a eu modification et false sinon
Mots clefs ... : AFFAIRE; LOOKUP ; GIGA
*****************************************************************}
Function GetAffaire ( TH : THCritMaskEdit; AcsTitre:string; AcsCode:string) : boolean ;
Var St : String ;
BEGIN
St:=TH.Text ;
LookupList(TH,AcsTitre,'AFFAIRE','AFF_AFFAIRE','AFF_LIBELLE','','AFF_AFFAIRE',False,5) ;
Result:=(St<>TH.Text) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Function qui lance un lookup sur l'unité à partir d'un control
Suite ........ : donné en entrée (THCritMaskEdit ou THGrid) et suivant le type de saisie
Suite ........ : dans l'activité passé en paramètre AcType et le type d'article
Suite ........ : Le lookup affecte le control automatiquement avec le code choisi
Suite ........ : Elle renvoie true s'il y a eu modification et false sinon
Mots clefs ... : UNITE; LOOKUP ; GIGA
*****************************************************************}
Function GetUnite ( TH : TControl; AcsTitre:string; AcType:T_Acc; AcsTypeArt:string) : boolean ;
Var St : String ;
where:string;
BEGIN
if (TH is THGrid) then
   with THGrid(TH) do
      st := Cells[SG_Unite, Row]
else
   St:=THCritMaskEdit(TH).Text ;

if (AcType=tacTemps) then
   where := 'GME_QUALIFMESURE="TEM" or GME_QUALIFMESURE="AUC"'
else
if (AcType=tacGlobal) then
   begin
   if (AcsTypeArt='') or (AcsTypeArt='PRE') then
      where := 'GME_QUALIFMESURE="TEM" or GME_QUALIFMESURE="AUC"'
   else
      where := 'GME_QUALIFMESURE="AUC"'
   end
else
   where := 'GME_QUALIFMESURE="AUC"';
//Result:=LookupList(TH,AcsTitre,'MEA','GME_MESURE','GME_LIBELLE','CC_TYPE="DEP"','CC_CODE',False,0,'',Letl) ;
// limiter à TEM et AUC pour les prestations,
// limiter à AUC pour frais et fournitures
Result:=LookupList(TH,AcsTitre,'MEA','GME_MESURE','GME_LIBELLE',where,'GME_MESURE',False,-1) ;


if (TH is THGrid) then
//   with THGrid(TH) do
//      Result:=Result AND (St<>Cells[SG_Unite, Row])
else
   Result:=Result AND (St<>THCritMaskEdit(TH).Text) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Function qui lance un lookup sur le type d'heure à partir d'un control
Suite ........ : donné en entrée (THCritMaskEdit ou THGrid)
Suite ........ : Le lookup affecte le control automatiquement avec le code choisi
Suite ........ : Elle renvoie true s'il y a eu modification et false sinon
Mots clefs ... : TYPEHEURE; LOOKUP ; GIGA
*****************************************************************}
Function GetTypeHeure ( TH : TControl; AcsTitre:string) : boolean ;
Var St : String ;
BEGIN
if (TH is THGrid) then
   with THGrid(TH) do
      st := Cells[SG_Unite, Row]
else
   St:=THCritMaskEdit(TH).Text ;

Result:=LookupList(TH,AcsTitre,'CHOIXCOD','CC_CODE','CC_LIBELLE','CC_TYPE="ATH"','CC_CODE',False,-1) ;


if (TH is THGrid) then
   with THGrid(TH) do
      Result:=Result AND (St<>Cells[SG_TypeHeure, Row])
else
   Result:=Result AND (St<>THCritMaskEdit(TH).Text) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Fonction de sélection de l'article en GA/GI suivant le
Suite ........ : type de recherche avancée ou non
Suite ........ :
Suite ........ : renvoie true s'il y a eu sélection d'un nouvel article, false
Suite ........ : sinon
Suite ........ :
Mots clefs ... : ARTICLE; RECHERCHE ; GIGA
*****************************************************************}
Function AFGetArticleRecherche ( Control:TControl ; TypeArticle : String; TitreSel : String ) : boolean ;
BEGIN
if  GetParamSocSecur('SO_GCRECHARTAV', False) then Result:=AFGetArticleMul(Control,TypeArticle,TitreSel)
                    												  else Result:=AFGetArticleLookUp(Control,TypeArticle,TitreSel) ;
END;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Function qui lance un lookup sur l'article à partir d'un control
Suite ........ : donné en entrée (THCritMaskEdit ou THGrid) en réduisant la liste
Suite ........ : en fonction du type d'article fourni en entrée
Suite ........ : Le lookup affecte le control automatiquement avec le code choisi
Suite ........ : Elle renvoie true s'il y a eu modification et false sinon
Mots clefs ... : ARTICLE; LOOKUP ; GIGA
*****************************************************************}
Function AFGetArticleLookUp ( Control:TControl ; TypeArticle : String; TitreSel : String ) : boolean ;
var
sWhere:string;
BEGIN
sWhere:=FabricWhereTypArt(TypeArticle) ;
Result:=LookupList(Control,TitreSel,'ARTICLE','GA_ARTICLE','GA_LIBELLE',sWhere,'GA_ARTICLE',False,7) ;

END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Fonction qui lance le MUL GA de recherche de la ressource à partir
Suite ........ : d'un control (THGrid ou THCritMaskEdit)
Suite ........ : Le code ressource choisi est rangé dans le control
Suite ........ :
Suite ........ : la fonction renvoie true si un code a été sélectionné
Mots clefs ... : RESSOURCE; MUL; GA ; GIGA
*****************************************************************}
Function GetAssistantMul ( Control:TControl ; TitreSel : String ; bInit : boolean) : boolean ;
Var CodeAss: String ;
BEGIN
  result := false;
  CodeAss := '';
  if (Control is THGrid) then
     with THGrid(Control) do
        CodeAss := Cells[Col,Row]
  else if (Control is THCritMaskEdit) and bInit then
     CodeAss := THCritMaskEdit(Control).Text;

  // mcd 12/06/02 CodeAss:=AGLLanceFiche('AFF','RESSOURCERECH_MUL','','','ARS_RESSOURCE='+CodeAss);
   Codeass:=AFLanceFiche_Rech_Ressource ('','ARS_RESSOURCE='+CodeAss);

  if CodeAss<>'' then
     BEGIN
     Result:=True ;
     if (Control is THGrid) then
        with THGrid(Control) do
           Cells[Col,Row] := CodeAss
     else if (Control is THCritMaskEdit) then
           THCritMaskEdit(Control).Text := CodeAss;
     END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Fonction qui lance le MUL GA de recherche de l'article à partir
Suite ........ : d'un control (THGrid ou THCritMaskEdit) en réduisant le MUL
Suite ........ : au type d'article choisi en entrée et en prépositionnant la sélection
Suite ........ : sur le code article saisi dans le control
Suite ........ : Le code article choisi est rangé dans le control
Suite ........ :
Suite ........ : la fonction renvoie true si un code a été sélectionné
Mots clefs ... : ARTICLE; MUL; GA ; GIGA
*****************************************************************}
Function AFGetArticleMul ( Control:TControl ; TypeArticle,TitreSel : String ) : boolean ;
  Var
  StChamps, sWhere, CodeArt, sTypeArticleEnabled, sSansLesFermes : String;
BEGIN
  sWhere := '';
  result := true;
  sTypeArticleEnabled := 'TypeArticleEnabled=';

  if (Control is THGrid) then
     with THGrid(Control) do
        CodeArt := Cells[Col,Row]
  else if (Control is THCritMaskEdit) then
     CodeArt := THCritMaskEdit(Control).Text;

  StChamps:='GA_CODEARTICLE='+Trim(Copy(CodeArt,1,18));


  if (TypeArticle='') then
     sTypeArticleEnabled := sTypeArticleEnabled + 'true'
  else
     begin
     sTypeArticleEnabled := sTypeArticleEnabled + 'false';
     sWhere:= 'GA_TYPEARTICLE='+TypeArticle;
     end;

   sSansLesFermes := 'SansArticleFermes=true';   // PL le 10/07/03 : on veut pouvoir interdire les article fermés

  // mcd 12/06/02 CodeArt:=AGLLanceFiche('AFF','AFARTICLE_RECH','','',StChamps+';'+sTypeArticleEnabled+';'+'NOFILTRE'+';'+sWhere) ;
  CodeArt:=AFLanceFiche_REch_Article (StChamps+';'+sTypeArticleEnabled+';'+sSansLesFermes+';'+'NOFILTRE'+';'+sWhere) ;

  if CodeArt<>'' then
     BEGIN
     Result:=True ;
     if (Control is THGrid) then
        with THGrid(Control) do
           Cells[Col,Row] := CodeArt
     else if (Control is THCritMaskEdit) then
           THCritMaskEdit(Control).Text := CodeArt;
     END ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Recherche d'une TOB article suivant le code en entrée et
Suite ........ : le type de code dans la TOBArticles fournie en entrée :
Suite ........ : AbRefDim = true => ARTICLE
Suite ........ : AbRefDim = false =>CODEARTICLE
Mots clefs ... : ARTICLE; TOB ; GIGA
*****************************************************************}
Function FindTOBArtSaisAff (TOBArticles:TOB; RefSais : String; AbRef:boolean ) : TOB ;
BEGIN
Result:=Nil ; if RefSais='' then Exit ;
if (TOBArticles=nil) then exit;

if (AbRef=true) then
   Result:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefSais],False)
else
   Result:=TOBArticles.FindFirst(['GA_CODEARTICLE'],[RefSais],False) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : zoom sur le code Article unique en monofiche et
Suite ........ : consultation
Mots clefs ... : ARTICLE; ZOOM ; GIGA
*****************************************************************}
Procedure ZoomArticle ( RefUnique : String; AcsTypeArticle: String ) ;
BEGIN
V_PGI.DispatchTT( 7, taConsult, RefUnique, '', 'MONOFICHE') ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : zoom sur le code ressource unique en monofiche et
Suite ........ : consultation
Mots clefs ... : RESSOURCE; ZOOM ; GIGA
*****************************************************************}
procedure ZoomRessource(AcsRessource:string);
begin
V_PGI.DispatchTT( 6, taConsult, AcsRessource, '', 'MONOFICHE') ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : zoom sur le code tiers unique en monofiche et
Suite ........ : consultation
Mots clefs ... : TIERS; ZOOM ; GIGA
*****************************************************************}
procedure ZoomClient(AcsTiers:string);
begin
V_PGI.DispatchTT( 8, taConsult, AcsTiers, '', 'MONOFICHE') ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : FV
Créé le ...... : 21/07/2015
Modifié le ... :   /  /
Description .. : zoom sur la sous-famille Matériel/Parc
Suite ........ : consultation
Mots clefs ... : FAMMAT; ZOOM ; GIGA
*****************************************************************}
procedure ZoomFamMateriel(FamMat : string);
begin
V_PGI.DispatchTT( 9, taConsult, FamMat, '', 'MONOFICHE') ;
end;   

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /    
Description .. : Fonction d'affichage d'un zoom contenant tous les articles 
Suite ........ : saisis sur une affaire, suivant le code affaire en entrée, le 
Suite ........ : type et le code article déjà saisi (partiellement).
Mots clefs ... : ARTICLES; AFFAIRE; ZOOM ; GIGA
*****************************************************************}
Procedure ZoomArticlesD1Affaire ( AcsCodeAffaire: String; AcsCodeArticle: String; AcsTypeArticle: String ; AcbSansPrix:boolean  ) ;
var
TobArticles:TOB;
sSansPrix:string;
BEGIN
sSansPrix:='A';
if AcbSansPrix then sSansPrix:='S';

//TobArticles:=nil;
if (AcsCodeAffaire='') then exit;

{$IFDEF AFFAIRE}
if GetParamSoc('SO_AFVARIABLES') then
  TobArticles := TOBArticlesD1AffaireOnyx (AcsCodeAffaire, AcsCodeArticle, AcsTypeArticle)
else
{$ENDIF}
  TobArticles := TOBArticlesD1Affaire(AcsCodeAffaire, AcsCodeArticle, AcsTypeArticle);



if (TobArticles<>nil) and (TobArticles.Detail.Count<>0) then
   begin
   if (TheTOB=nil) then TheTOB:=TOB.Create('la TOB',nil,-1);

   TheTOB.ClearDetail;
   TobArticles.ChangeParent(TheTOB, -1);

   // mcd 12/06/02 AGLLanceFiche('AFF','AFARTICLESAFFAIRE','','',AcsCodeAffaire+';'+AcsCodeArticle+';'+AcsTypeArticle+';'+sSansPrix) ;
   AFLanceFiche_ArticleParAff(AcsCodeAffaire+';'+AcsCodeArticle+';'+AcsTypeArticle+';'+sSansPrix) ;
   end;

END ;




{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Fonction y-a-t-il pour une affaire, plusieurs ligne
Suite ........ : de pièce contenant le même article fourni en entrée ?
Mots clefs ... : ARTICLES; AFFAIRE; NOMBRE ; GIGA
*****************************************************************}
function NbArticlesD1Affaire(AcsCodeAffaire:string; var AvsCodeArticle:string;
                                 var AvsLibArticle:string; var AvdPrHT, AvdPvHT:double; var numordre : integer;
                                 var FormuleVar : string; var MntRemise : double; var NumPieceOrig : String):integer; // PL le 06/08/03 : pour CB ajout NumPieceOrig
Var i, index : Integer;
    TOBPiece, TOBLigne:TOB;
Begin
Result := 0;
index := 0;
AvsLibArticle:='';
AvdPrHT:=0;
AvdPvHT:=0;
if (AcsCodeAffaire='') then exit;
TobPiece:= TOB.Create('PIECE', nil, -1);
TobPiece.PutValue('GP_AFFAIRE',AcsCodeAffaire);

try
{$IFDEF AFFAIRE}
LectureLignesAffaire(TobPiece,False);
{$endif}
for i:=0 to (TobPiece.Detail.Count-1) do
   begin
   if (TobPiece.Detail[i].GetValue('GL_NATUREPIECEG')=VH_GC.AFNatAffaire) and
      ((TobPiece.Detail[i].GetValue('GL_CODEARTICLE')=AvsCodeArticle) or (AvsCodeArticle='')) then
         begin
         Result:=Result+1;
         index:=i;
         end;
   end;

if (Result = 1) then
   begin
   TOBLigne := TobPiece.Detail[index];
   if (AvsCodeArticle='') then AvsCodeArticle:=TOBLigne.GetValue('GL_CODEARTICLE');
   AvsLibArticle:=TOBLigne.GetValue('GL_LIBELLE');
   //AvdPrHT:=TOBLigne.GetValue('GL_DPR'); // PL le 08/09/03 : faudra voir à se mettre d'accord entre le DPR et le PMRP
   AvdPrHT:=TOBLigne.GetValue('GL_PMRP');
   //AvdPvHT:=TOBLigne.GetValue('GL_MONTANTHT'); // PL le 08/09/03 : ça va pas non ?!!! c'est le PVHT qu'il faut mettre la dedans !
   AvdPvHT:=TOBLigne.GetValue('GL_PUHT');
   numordre := TOBLigne.GetValue('GL_NUMORDRE');
   FormuleVar := TOBLigne.GetValue('GL_FORMULEVAR');
   MntRemise := TOBLigne.GetValue('GL_VALEURREMDEV');
   NumPieceOrig := EncodeRefPiece (TOBLigne); // PL le 06/08/03 : pour CB ajout NumPieceOrig
   end;

finally
TOBPiece.Free;
end;

End;


{=============================== Entete ================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Initialisation des variables de suivi des colonnes du Grid
Suite ........ : de saisie de l'activité
Mots clefs ... : ACTIVITE; INIT ; GIGA
*****************************************************************}
Procedure InitLesColsAct ;
BEGIN
SG_Curseur:=-1; SG_Jour:=-1 ; SG_Affaire:=-1; SG_Aff0:=-1; SG_Aff1:=-1; SG_Aff2:=-1; SG_Aff3:=-1; SG_Tiers:=-1 ; SG_Ressource:=-1 ;
SG_Article:=-1 ; SG_Unite:=-1 ; SG_Qte:=-1 ; SG_PUCH:=-1 ; SG_PV:=-1 ;SG_TOTPRCH:=-1 ;SG_Desc:=-1 ;SG_Lib:=-1 ;
SG_TotV:=-1 ; SG_TypA:=-1 ; SG_ActRep:=-1 ;  SG_TypeHeure := -1; SG_MntRemise := -1; SG_MontantTVA := -1;
SG_Avenant:=-1;

END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Initialisation de la TOB associée à la ligne d'activité suivant
Suite ........ : le contexte de saisie : n° de ligne, toute la clé de l'activité,
Suite ........ : l'action et le code devise
Mots clefs ... : ACTIVITE; TOB; INIT ; GIGA
*****************************************************************}
Procedure InitialiseLigneAct ( TOBL : TOB; ARow : integer; CleAct : R_CLEACT; Action : TActionFiche;
                                CodeDev : string);
  var
  dateCourante, PremJourSem : TDateTime;
begin
  TOBL.PutValue ('ACT_NUMLIGNE', ARow) ;
  TOBL.PutValue ('ACT_NUMLIGNEUNIQUE', 0) ;
  TOBL.PutValue ('ACT_TYPEACTIVITE', CleAct.TypeActivite) ;
  TOBL.PutValue ('ACT_FOLIO', CleAct.Folio) ;
  TobL.PutValue ('ACT_DATECREATION', Date) ;
  TOBL.PutValue ('ACT_DEVISE', CodeDev) ;
  TOBL.PutValue ('ACT_CREATEUR', V_PGI.User) ;
  TOBL.PutValue ('ACT_ACTORIGINE', 'SAI') ;
  if (VH_GC.AFTYPESAISIEACT = 'SEM') then
    begin
      dateCourante := EncodeDate (CleAct.Annee, CleAct.Mois, CleAct.Jour);

      {$IFDEF AGLINF545}
      PremJourSem:= PremierJourSemaineTempo (CleAct.Semaine, CleAct.Annee);
      {$ELSE}
      // PL le 23/05/02 : réparation de la fonction , CleAct.Mois  AGL550
      PremJourSem := PremierJourSemaine (CleAct.Semaine, CleAct.Annee);
      //////////////////////////// AGL550
      {$ENDIF}

      if (dateCourante < PremJourSem) or (dateCourante > PremJourSem + 6) then
        dateCourante := PremJourSem;
    end
  else
    dateCourante := EncodeDate (CleAct.Annee, CleAct.Mois, 1);

  TOBL.PutValue ('ACT_DATEACTIVITE', dateCourante) ;
  TOBL.PutValue ('ACT_PERIODE', GetPeriode(DateCourante));
  TOBL.PutValue ('ACT_SEMAINE', NumSemaine(DateCourante));

  if (CleAct.TypeSaisie = tsaRess) then
    begin
      TOBL.PutValue ('ACT_RESSOURCE', CleAct.Ressource);
      TOBL.PutValue ('ACT_TYPERESSOURCE', CleAct.TypeRess);
      TOBL.PutValue ('ACT_FONCTIONRES', CleAct.Fonction);
      if Not (VH_GC.AFProposAct) then
        begin
          TOBL.PutValue ('ACT_AFFAIRE0' , 'A' );
          CleAct.Aff0 := 'A';
        end;
    end
  else
    begin
      {$IFDEF BTP}
    	 BTPCodeAffaireDecoupe (CleAct.Affaire, CleAct.Aff0, CleAct.Aff1, CleAct.Aff2, CleAct.Aff3, CleAct.Avenant, Action, false);
	  {$ELSE}
      CodeAffaireDecoupe(CleAct.Affaire, CleAct.Aff0, CleAct.Aff1, CleAct.Aff2, CleAct.Aff3, CleAct.Avenant, Action, false);
      {$ENDIF}
      TOBL.PutValue ('ACT_AFFAIRE'  , CleAct.Affaire);
      TOBL.PutValue ('ACT_AFFAIRE0' , CleAct.Aff0   );
      TOBL.PutValue ('ACT_AFFAIRE1' , CleAct.Aff1   );
      TOBL.PutValue ('ACT_AFFAIRE2' , CleAct.Aff2   );
      TOBL.PutValue ('ACT_AFFAIRE3' , CleAct.Aff3   );
      TOBL.PutValue ('ACT_AVENANT'  , CleAct.Avenant);
      TOBL.PutValue ('ACT_TIERS'    , CleAct.Tiers  );
    end;

  // configuration de la nouvelle TOB detail suivant le type d'acces à la saisie, par temps, frais, fournitures, global
  case CleAct.TypeAcces of
    tacTemps :
              begin
                TOBL.PutValue ('ACT_TYPEARTICLE', 'PRE');
                TOBL.PutValue ('ACT_ACTIVITEEFFECT', 'X') ;
              end;
    tacFrais :
              begin
                TOBL.PutValue ('ACT_TYPEARTICLE', 'FRA');
                TOBL.PutValue ('ACT_ACTIVITEEFFECT', '-') ;
              end;
    tacFourn :
              begin
                TOBL.PutValue ('ACT_TYPEARTICLE', 'MAR');
                TOBL.PutValue ('ACT_ACTIVITEEFFECT', '-') ;
              end;
    tacGlobal :
              begin

              end;
  end;

  if TOBL.FieldExists ('ARTSLIES') then // PL le 11/09/03 : articles liés
    TOBL.PutValue('ARTSLIES', '');   // Ce champs boolean se doit d'être initialisé à vide


(*if TOBL.GetValue('GL_ARTICLE')<>'' then
   BEGIN
   TOBL.PutValue('GL_QTESTOCK',1) ;
   TOBL.PutValue('GL_QTEFACT',1) ;
   END ;
   *)
end;

Function FabricWhereTypArt ( TypeArticle : String ) : String ;
BEGIN
Result:='' ; if TypeArticle='' then Exit ;

Result:='(GA_TYPEARTICLE="'+TypeArticle+'")' ;

END ;

{=========================== Opérations sur le Grid ===============================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /    
Description .. : fonction Test du remplissage significatif d'une ligne d'activité 
Suite ........ : sur le grid
Suite ........ :
Suite ........ : renvoie true si la ligne est considérée comme remplie sur le
Suite ........ : grid de saisie
Mots clefs ... : ACTIVITE; GRID; SAISIE ; GIGA 
*****************************************************************}
Function EstRempliGCActivite (GS : THGrid; Lig : integer; bAvecJour : boolean = true ) : boolean;
  Var
  LaCol : integer ;
BEGIN
  Result := False;
  for LaCol := GS.FixedCols to GS.ColCount - 1 do
    BEGIN
      if (LaCol <> SG_Curseur) and (GS.ColWidths [LaCol] <> 0) then
        Result := (GS.Cells [LaCol, Lig] <> '');

      // Si le jour n'est pas considéré comme significatif, le résultat n'est pas pris en compte
      if (bAvecJour = False) and (LaCol = SG_Jour) then
        Result := false;

      // Si la partie 0 du code affaire est remplie et qu'aucune des trois autres parties n'est remplie, on ne
      // tient pas compte de cette saisie
      if (Result and (LaCol = SG_Aff0) and (GS.Cells [SG_Aff1, Lig] = '') and (GS.Cells [SG_Aff1, Lig] = '')
          and (GS.Cells [SG_Aff1, Lig] = '')) then
        Result := false;

      if Result then
        Break;
    END;

END;


{***********A.G.L.***********************************************
Auteur  ...... : Pierre Lenormand
Créé le ...... : 31/01/2000
Modifié le ... :   /  /
Description .. : Pour une unité de conversion en entrée, renvoie une
Suite ........ : TOB avec comme TOBs détailles les unités de même type (y compris elle même)
Mots clefs ... : UNITE; CONVERSION; QUOTITE; MEA ; GIGA
*****************************************************************}
function ListeUnitesConversion(AcsCodeAConvertir:string):TOB;
// ATTENTION : la libération de la TOB est à la charge de l'appelant
// exemple : ListeUnitesConversion('H')
// en sortie la TOB entête contient les données de l'unité 'H' dans la table MEA
// et les TOB détailles contiennent, 'H' et données associées, 'J' et données associées etc...
// utilisation type : il me faut la liste de toutes les unités susceptible d'être utilisées pour la saisie des
// temps et leur coefficient, afin de ramener tout type de saisie (en jour, en semaine...) en heure.
var
TOBAConvertir, TOBUnit :TOB;
Q : TQuery ;
sType:string;
i:integer;

   procedure  AjouteTOBMEA(T : TOB) ;
   var
   TOBU:TOB;
   begin
   TOBU := Tob.Create('MEA',TOBAConvertir,-1);
   TOBU.Dupliquer(T, true, true, false);
   end;

begin
Result := nil;
TOBAConvertir := Tob.Create('MEA',nil,-1);

if (VH_GC.MTOBMEA = nil) then
   begin
   Q := nil;
   try
    // SELECT * possible car très peu de champs, très peu de lignes
      Q:=OpenSQL('SELECT * FROM MEA WHERE GME_MESURE="'+AcsCodeAConvertir+'"',True,-1,'',true) ;
      if Not TOBAConvertir.SelectDB('', Q) then exit;
   finally
   ferme(Q);
   end;

   sType := TOBAConvertir.GetValue('GME_QUALIFMESURE');
   if TOBAConvertir.LoadDetailDB('MEA','"'+sType+'"','GME_MESURE',Nil,FALSE) then
      Result := TOBAConvertir
   else
        begin
        TOBAConvertir.Free;
        TOBAConvertir:=nil;
        end;
   end
else
   begin
   TOBUnit := VH_GC.MTOBMEA.FindFirst(['GME_MESURE'], [AcsCodeAConvertir], true);
   if (TOBUnit <> nil) then
      begin
//      VH_GC.TOBMEA.ParcoursTraitement(['GME_QUALIFMESURE'],[TOBUnit.GetValue('GME_QUALIFMESURE')], false, AjouteTOBMEA);
      for i:=0 to VH_GC.MTOBMEA.Detail.Count-1 do
         if (VH_GC.MTOBMEA.Detail[i].GetValue('GME_QUALIFMESURE')=TOBUnit.GetValue('GME_QUALIFMESURE')) then
            AjouteTOBMEA( VH_GC.MTOBMEA.Detail[i] ) ;

      if (TOBAConvertir.Detail.Count<>0) then
         Result := TOBAConvertir;
      end
   else
        begin
        TOBAConvertir.Free;
        TOBAConvertir:=nil;
        end;
   end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Pierre Lenormand
Créé le ...... : 17/04/2000
Modifié le ... :   /  /
Description .. : Pour une unité source, une unité cible en entrées, renvoie vrai si l'unité source est convertible en unité cible
Mots clefs ... : CONVERSION; UNITE; MEA ; GIGA
*****************************************************************}
function IsConvertible(AcsCodeSource, AcsCodeCible:string):boolean;
var
TOBT:TOB;
sFamSource, sFamCible : string;
dQuotS, dQuotC : double;
begin
//dQuotS:=0;
//dQuotC:=0;
Result:=false;

TOBT := VH_GC.MTOBMEA.FindFirst(['GME_MESURE'], [AcsCodeSource], true);
if (TOBT=nil) then exit;
sFamSource:=TOBT.GetValue('GME_QUALIFMESURE');
dQuotS:=TOBT.GetValue('GME_QUOTITE');

TOBT := VH_GC.MTOBMEA.FindFirst(['GME_MESURE'], [AcsCodeCible], true);
if (TOBT=nil) then exit;
sFamCible:=TOBT.GetValue('GME_QUALIFMESURE');
dQuotC:=TOBT.GetValue('GME_QUOTITE');

if (sFamSource<>sFamCible) or (dQuotS=0) or (dQuotC=0) then exit;

Result:=true;

end;

{=========================== Remplir les TOBs ===============================}

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /    
Description .. : Remplit une TOB assistant suivant le code assistant fourni 
Suite ........ : en entrée (recherche par l'objet RESSOURCE)
Suite ........ : renvoie true si la tob assistant est remplie ou si le code en 
Suite ........ : entrée est vide, false si le code n'existe pas
Mots clefs ... : RESSOURCE; TOB ; GIGA
*****************************************************************}
Function RemplirTOBAssistant ( TOBAssistant:TOB; AFOAssistants:TAFO_Ressources;CodeAssistant : String ) : integer ;
BEGIN
Result:=-1;
if (AFOAssistants=nil) then exit;


Result:=AFOAssistants.AddRessource(CodeAssistant); 

if (CodeAssistant<>TOBAssistant.GetValue('ARS_RESSOURCE')) then
   BEGIN

   if (Result<>-1) and (Result<>-2) then
      // -1 le code ressource n'existe pas
      // -2 le code ressource en entrée est vide
      begin
      TOBAssistant.ClearDetail ;
      TOBAssistant.InitValeurs(False) ;
      TOBAssistant.Dupliquer(TAFO_Ressource(AFOAssistants.Objects[Result]).tob_Champs,TRUE,TRUE,TRUE);
      end
   else
      if (Result=-2) then
         begin
         TOBAssistant.ClearDetail ;
         TOBAssistant.InitValeurs(False) ;
         end
      else
         Result:=-1;
   END ;

END ;


{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Remplit une TOB unités suivant le code unité fourni
Suite ........ : en entrée
Suite ........ : renvoie true si la tob unité est remplie
Mots clefs ... : UNITE; TOB ; GIGA
*****************************************************************}
Function RemplirTOBUnites ( TOBListeConv:TOB; CodeUnite: String; TOBUnites : TOB ) : boolean ;
Var TOBU : TOB ;
BEGIN
Result:=false ; 
if (TOBUnites=nil) then exit;
TOBUnites.ClearDetail ; TOBUnites.InitValeurs(False) ;

TOBU := TOBListeConv.FindFirst(['GME_MESURE'],[CodeUnite],TRUE) ;
if (TOBU<>nil) then
   begin
   TOBUnites.Dupliquer(TOBU,TRUE,TRUE,TRUE) ;
   Result:=true;
   end;

END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Remplit une TOB Tiers suivant le code tiers fourni
Suite ........ : en entrée
Suite ........ : renvoie true si la tob unité est remplie ou vide (code vide en entrée)
Mots clefs ... : TIERS; TOB ; GIGA
*****************************************************************}
Function RemplirTOBTiersP ( TOBTiers, TOBTierss : TOB ; CodeTiers : String ) : boolean ;
Var Q : TQuery ;
T:TOB;
BEGIN
Result:=true ;
if CodeTiers='' then
   BEGIN
   TOBTiers.ClearDetail ;
   TOBTiers.InitValeurs(False) ;
   END
else
   if (CodeTiers<>TOBTiers.GetValue('T_TIERS')) then
      begin
      Result:=false ;
      T:=nil;
      if (TOBTierss<>nil) then
            T:=TOBTierss.FindFirst(['T_TIERS','T_NATUREAUXI'], [CodeTiers, 'CLI'], false);

      if (T<>nil) then
            begin
            TOBTiers.ClearDetail ;
            TOBTiers.InitValeurs(False) ;
            TOBTiers.dupliquer(T, true, true, false);
            Result:=true;
            end
      else
        begin
          Q := nil;
          try
          // PL le 06/03/02 : INDEX 5
          Q:=OpenSQL('SELECT T_TIERS, T_NATUREAUXI, T_LIBELLE, T_MOISCLOTURE, T_FERME FROM TIERS WHERE T_NATUREAUXI="CLI" AND T_TIERS="' +CodeTiers+ '"',True,-1,'',true) ;
          if Not Q.EOF then
           begin
           TOBTiers.SelectDB('',Q);
           if (TOBTierss<>nil) then
              begin
              T:=TOB.Create('',TOBTierss,-1) ;
              T.dupliquer(TOBTiers, true, true, false);
              end;
           Result:=true;
           end;
          finally
          Ferme(Q) ;
          end;
        end;
      end;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 01/03/2002
Modifié le ... :   /  /
Description .. : Remplit une TOB Tiers suivant le libellé tiers fourni
Suite ........ : en entrée
Suite ........ : renvoie true si la tob unité est remplie ou vide (code vide en entrée)
Mots clefs ... : TIERS; TOB ; GIGA
*****************************************************************}
Function RemplirTOBTiersLibP ( TOBTiers, TOBTierss : TOB ; LibTiers : String ) : boolean ;
Var Q : TQuery ;
T:TOB;
BEGIN
Result:=true ;
if LibTiers='' then
   BEGIN
   TOBTiers.ClearDetail ;
   TOBTiers.InitValeurs(False) ;
   END
else
   if (LibTiers<>TOBTiers.GetValue('T_LIBELLE')) then
      begin
      Result:=false ;
      T:=nil;
      if (TOBTierss<>nil) then
            T:=TOBTierss.FindFirst(['T_LIBELLE','T_NATUREAUXI'], [LibTiers, 'CLI'], false);

      if (T<>nil) then
            begin
            TOBTiers.ClearDetail ;
            TOBTiers.InitValeurs(False) ;
            TOBTiers.dupliquer(T, true, true, false);
            Result:=true;
            end
      else
        begin
          Q := nil;
          try
          Q:=OpenSQL('SELECT T_TIERS, T_NATUREAUXI, T_LIBELLE, T_MOISCLOTURE FROM TIERS WHERE T_LIBELLE="' +LibTiers+ '" AND T_NATUREAUXI="CLI"',True,-1,'',true) ;
          if Not Q.EOF then
           begin
           TOBTiers.SelectDB('',Q);
           if (TOBTierss<>nil) then
              begin
              T:=TOB.Create('',TOBTierss,-1) ;
              T.dupliquer(TOBTiers, true, true, false);
              end;
           Result:=true;
           end;
          finally
          Ferme(Q) ;
          end;
        end;
      end;
END ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Valorisation d'une ligne d'activité suivant le contexte de
Suite ........ : saisie (paramétrage, type d'accès, type de valorisation...)
Mots clefs ... : VALO ; GIGA
*****************************************************************}
function MajTOBValo(AcdDate:TDateTime; AcTypeAcces:T_Acc; AcsAffaire:string; AcsRessource:string; AcsCodeArticle:string;
                    AcsTypeHeure:string; TOBArticles, TOBAffaires, TobZoom:TOB; AFOAssistants:TAFO_Ressources;
                    bAvecZoom:boolean; AcsValoActPR,AcsValoActPV : string; var AvbPROK, AvbPVOK : boolean):TOB;
var
LibArticle, Article, codeArticle, sTypeArt, sUniteRessourcePR, sUniteRessourcePV:string;
dPrHT, dPvHT, dPrHTRes, dPvHTRes, MntRemise:double;
FormuleVar : string;
//rechart : T_RechArt;
TOBA,TOBRep:TOB;
NbArticles, IndexRess, NumOrdre : integer;
NumPieceOrig : String; // PL le 06/08/03 : pour CB ajout NumPieceOrig
bForcerValoArtPR, bForcerValoArtPV, bUneValoAff, bValoAffPR, bValoAffPV, bUneValoRes, bUneValoArt, bValoResPR, bValoResPV, bFraisFourn, bValoArtPR, bValoArtPV:boolean;

   procedure MajPrixValorises(TOBResult:TOB);
      var
      dCoeffConvert, TxHSupPR, TxHSupPV:double;
      begin
      // Dans tous les cas de valorisation, on récupère le code et le libellé de la ligne de piece
      TOBResult.PutValue('GA_CODEARTICLE', codeArticle);
      TOBResult.PutValue('GA_LIBELLE', LibArticle);
      if not TOBResult.FieldExists('NUMORDRE') then  //ONYX-AFORMULEVAR
        TOBResult.AddChampSup('NUMORDRE', false);
      TOBResult.PutValue('NUMORDRE', NumOrdre);
      if not TOBResult.FieldExists('NUMPIECEORIG') then  // PL le 06/08/03 : pour CB ajout NumPieceOrig
        TOBResult.AddChampSup('NUMPIECEORIG', false);
      TOBResult.PutValue('NUMPIECEORIG', NumPieceOrig);
      if not TOBResult.FieldExists('MNTREMISE') then
        TOBResult.AddChampSup('MNTREMISE', false);
      TOBResult.PutValue('MNTREMISE', MntRemise);
      TOBResult.PutValue('GA_FORMULEVAR', FormuleVar);

      if bValoArtPR then  dPrHT := TOBResult.GetValue('GA_PMRP');
      if bValoArtPV then  dPvHT := TOBResult.GetValue('GA_PVHT');

      // Les prix fournis sont en unité de facturation, il faut les convertir en unité d'activité
      dCoeffConvert := ConversionUnite(TOBResult.GetValue('GA_QUALIFUNITEVTE'), TOBResult.GetValue('GA_QUALIFUNITEACT'), 1);
      if dCoeffConvert=0 then dCoeffConvert:=1;
      dPrHT := dPrHT / dCoeffConvert;
      dPvHT := dPvHT / dCoeffConvert;

      // Majoration heures sup
      if GetParamSoc('SO_AFTYPEHEUREACT')= True then
      if TauxHSupFromCode (AcsTypeHeure, TxHSupPR, TxHSupPV) then
         begin
         dPrHT := dPrHT * TxHSupPR;
         dPvHT := dPvHT * TxHSupPV;

         dPrHTRes := dPrHTRes * TxHSupPR;
         dPvHTRes := dPvHTRes * TxHSupPV;
         end;


      // Dans le cas de la valo par article, on modifie les prix avec les valeurs converties en unité d'activité
      if bValoArtPR then
         TOBResult.PutValue('GA_PMRP', dPrHT);

      if bValoArtPV then
         TOBResult.PutValue('GA_PVHT', dPvHT);

      // Dans le cas de la valo par affaire, on modifie les prix avec les valeurs à partir de la ligne de piece
      if (AcsValoActPR='AFF') then
         TOBResult.PutValue('GA_PMRP', dPrHT);

      if (AcsValoActPV='AFF') then
         TOBResult.PutValue('GA_PVHT', dPvHT);

      // Dans le cas de la valo par ressource, on modifie les prix avec les valeurs de la ressource
      // Les prix fournis sont déjà en unité d'activité
      if bValoResPR then
         TOBResult.PutValue('GA_PMRP', dPrHTRes);

      if bValoResPV then
         TOBResult.PutValue('GA_PVHT', dPvHTRes);
      end;

begin
NumOrdre := 0;
MntRemise := 0;
FormuleVar := '';
NumPieceOrig := '';  // PL le 06/08/03 : pour CB ajout NumPieceOrig
AvbPROK:=true; AvbPVOK:=true;
sUniteRessourcePR:='';
sUniteRessourcePV:='';
codeArticle:=AcsCodeArticle;
LibArticle:=''; Article:='';
dPrHT:=0;dPvHT:=0;dPrHTRes:=0;dPvHTRes:=0;
bUneValoAff:=false; bValoAffPR:=false; bValoAffPV:=false; bUneValoRes:=false; bUneValoArt:=false; bFraisFourn:=false;
bValoResPR:=false; bValoResPV:=false; bValoArtPR:=false; bValoArtPV:=false;bForcerValoArtPR:=false; bForcerValoArtPV:=false;
NbArticles:=-1;
IndexRess:=-1;
Result := nil;
TOBA:=nil;

if (AcTypeAcces=tacFrais) or (AcTypeAcces=tacFourn) then bFraisFourn:=true;
if (AcsValoActPR='ART') then bValoArtPR:=true;
if (AcsValoActPV='ART') then bValoArtPV:=true;
if (AcsValoActPR='RES') then bValoResPR:=true;
if (AcsValoActPV='RES') then bValoResPV:=true;

// en saisie de frais ou de fourniture, il n'y a pas de valorisation par assistant
// on passe alors en valorisation par article
// SAUF s'il existe une personnalisation par rapport à l'article pour la ressource courante (pas par rapport à la famille d'article)
// PL le 03/10/02
if (bFraisFourn and bValoResPR) then
   begin
   bForcerValoArtPR:=true;
   end;
if (bFraisFourn and bValoResPV) then
   begin
   bForcerValoArtPV:=true;
   end;

// Conditions d'impossibilite de valorisation :
if ((AcsCodeArticle='') and (AcsAffaire='') and (AcsValoActPR='AFF') and (AcsValoActPV='AFF')) or
   ((AcsCodeArticle='') and (AcsRessource='') and bValoResPR and bValoResPV) or
   ((AcsCodeArticle='') and bValoArtPR and bValoArtPV)
   then exit;

if (AcsValoActPR='AFF') then bValoAffPR:=true;
if (AcsValoActPV='AFF') then bValoAffPV:=true;
if (bValoAffPR or bValoAffPV) then bUneValoAff := true;
if (bValoResPR or bValoResPV) then bUneValoRes:=true;
if (bValoArtPR or bValoArtPV) then bUneValoArt:=true;

if (bUneValoAff and bUneValoRes and (AcsAffaire='') and (AcsRessource='') and (AcsCodeArticle='')) or
   (bUneValoAff and bUneValoArt and (AcsAffaire='') and  (AcsCodeArticle='')) or
   (bUneValoRes and bUneValoArt and (AcsRessource='') and  (AcsCodeArticle=''))
   then exit;

sTypeArt:='';
case AcTypeAcces of
  tacTemps :  begin
              sTypeArt := 'PRE';
              end;
  tacFrais :  begin
              sTypeArt := 'FRA';
              end;
  tacFourn :  begin
              sTypeArt := 'MAR';
              end;
end;

// PL le 25/09/03 : on déplace ce bloc pour généraliser à tous les cas de valorisation
// on remplit les zones dans tous les cas à partir du moment où l'on a choisi un article dans le zoom
// de choix des articles de l'affaire
if (TobZoom <> nil) then
  begin
    NbArticles := 1;
    CodeArticle := TobZoom.GetValue('GA_CODEARTICLE');
    LibArticle := TobZoom.GetValue('GA_LIBELLE');
    dPrHT := TobZoom.GetValue('GA_PMRP');
    dPvHT := TobZoom.GetValue('GA_PVHT');
    FormuleVar := TobZoom.GetValue('GA_FORMULEVAR');
    if TobZoom.FieldExists('NUMORDRE') then  //ONYX-AFORMULEVAR
      numordre := TobZoom.GetValue('NUMORDRE');
    if TobZoom.FieldExists('NUMPIECEORIG') then  // PL le 06/08/03 : pour CB ajout NumPieceOrig
      NumPieceOrig := TobZoom.GetValue('NUMPIECEORIG');
    if TobZoom.FieldExists('MNTREMISE') then
      MntRemise := TobZoom.GetValue('MNTREMISE');
    // PL le 17/06/03 : si on a deja choisi dans le zoom et qu'il n'y a qu'un article
    // inutile de choisir à nouveau...
    bAvecZoom := false;
  end;

// Combien de fois cet article est present sur les lignes de pieces de l'affaire
if (bUneValoAff = true) then
  begin
 (*
   if (TobZoom <> nil) then
      begin
        NbArticles := 1;
        CodeArticle := TobZoom.GetValue('GA_CODEARTICLE');
        LibArticle := TobZoom.GetValue('GA_LIBELLE');
        dPrHT := TobZoom.GetValue('GA_PMRP');
        dPvHT := TobZoom.GetValue('GA_PVHT');
        FormuleVar := TobZoom.GetValue('GA_FORMULEVAR');
        if TobZoom.FieldExists('NUMORDRE') then  //ONYX-AFORMULEVAR
          numordre := TobZoom.GetValue('NUMORDRE');
        if TobZoom.FieldExists('NUMPIECEORIG') then  // PL le 06/08/03 : pour CB ajout NumPieceOrig
          NumPieceOrig := TobZoom.GetValue('NUMPIECEORIG');
        if TobZoom.FieldExists('MNTREMISE') then
          MntRemise := TobZoom.GetValue('MNTREMISE');
        // PL le 17/06/03 : si on a deja choisi dans le zoom et qu'il n'y a qu'un article
        // inutile de choisir à nouveau...
        bAvecZoom := false;
      end
   else*)
  if (TobZoom = nil) then
    NbArticles := NbArticlesD1Affaire (AcsAffaire, codeArticle, LibArticle, dPrHT, dPvHT, numordre, FormuleVar, MntRemise, NumPieceOrig); // PL le 06/08/03 : pour CB ajout NumPieceOrig

    // Si plusieurs articles sont possibles en valorisation par affaire, et que le zoom de choix n'est pas activé
    // On doit avertir l'appelant que le prix valorisé par affaire n'est pas significatif
    if (NbArticles > 1) then
      begin
        if bValoAffPR then AvbPROK := false;
        if bValoAffPV then AvbPVOK := false;
        // Si le zoom est activé, il se déclenchera et le choix de l'article rectifiera les variables AvbPROK et AvbPVOK
        // Sinon, il faut valoriser sur l'autre Prix, celui qui n'est pas valorisé sur l'affaire (s'il y en a un)
        if (bAvecZoom = false) and ((bValoAffPR = false) or (bValoAffPV = false)) then NbArticles := 1;
      end;
 end;

if (codeArticle<>'') then
   begin
   TOBA:=TrouverCodeArticle(codeArticle, sTypeArt, TOBArticles);
   if (TOBA<>nil) then
        codeArticle:=TOBA.GetValue('GA_CODEARTICLE');

   if (TOBA<>nil) and (NbArticles<1) then
      begin
      NbArticles:=1;
      // en saisie globale (tous types d'article) si l'article selectionné est une fourniture ou un frais, on ne valorise
      // pas par ressource... si c'était une valorisation par ressource qui était prévue, mais on valorise par article
      // SAUF s'il existe une personnalisation par rapport à l'article pour la ressource courante (pas par rapport à la famille d'article)
      // PL le 03/10/02
      if (AcTypeAcces=tacGlobal) and ((TOBA.GetValue('GA_TYPEARTICLE')='FRA') or (TOBA.GetValue('GA_TYPEARTICLE')='MAR')) then
         begin
         if (bValoResPR) then
            begin
            bForcerValoArtPR:=true;
            end;
         if (bValoResPV) then
            begin
            bForcerValoArtPV:=true;
            end;
         end;
      end;
   end;


// En valorisation par ressource pour une fourniture ou un frais, on ne valorise
// pas par ressource... mais on valorise par article
// SAUF s'il existe une personnalisation par rapport à l'article pour la ressource courante (pas par rapport à la famille d'article)
// PL le 03/10/02
if (bValoResPR or bValoResPV) then
    begin
    IndexRess := AFOAssistants.AddRessource(AcsRessource);
    if (IndexRess<>-1) and (IndexRess<>-2) then
        if (TOBA<>nil) then
            begin
            if (bForcerValoArtPR=true) then
                begin
                if Not TAFO_Ressource(AFOAssistants.Objects[IndexRess]).ExisteUnArticleDeValo('R',TOBA.GetValue('GA_ARTICLE')) then
                    begin
                    bValoResPR := false;
                    bValoArtPR := true;
                    end;
                end;
            if (bForcerValoArtPV=true) then
                begin
                if Not TAFO_Ressource(AFOAssistants.Objects[IndexRess]).ExisteUnArticleDeValo('V',TOBA.GetValue('GA_ARTICLE')) then
                    begin
                    bValoResPV := false;
                    bValoArtPV := true;
                    end;
                end;
            end;
    end;

bUneValoRes:=(bValoResPR or bValoResPV);
//bUneValoArt:=(bValoArtPR or bValoArtPV);

// l'assistant en entree existe-t-il dans la TOBAssistants
if (bUneValoRes=true) then
   begin
   if (IndexRess<>-1) and (IndexRess<>-2) then
      begin
      if ((codeArticle='') and (NbArticles<1)) then
         // Si on avait encore rien trouvé, ni par l'article en entree, ni par les articles de l'affaire
         // on prend l'article par defaut de la ressource comme article de reference
         if TAFO_Ressource(AFOAssistants.Objects[IndexRess]).AssocieArticle('') then
            begin
            TOBA:=CreerTOBArt(TOBArticles) ;
            TOBA.InitValeurs(FALSE) ;
            TOBA.Dupliquer(TAFO_Ressource(AFOAssistants.Objects[IndexRess]).Article,true,true,false);
            NbArticles:=1;
            end;

      // détermine les prix de revient et de vente de la ressource exprimés dans l'unité d'ACTIVITE de l'article
      // en cours (passé en entrée ou lié à la ressource)
      if (TOBA<>nil) then
         TAFO_Ressource(AFOAssistants.Objects[IndexRess]).gsUniteValo := TOBA.GetValue('GA_QUALIFUNITEACT');
      dPrHTRes:=TAFO_Ressource(AFOAssistants.Objects[IndexRess]).PRRessource(AcdDate, TOBA, TOBAffaires, '', AcsAffaire);

      if (AcsValoActPV='RES') then
         dPvHTRes:=TAFO_Ressource(AFOAssistants.Objects[IndexRess]).PVRessource(AcdDate, true, TOBA, TOBAffaires, '', AcsAffaire); // true pour le HT
      end;
   end;

// si on n'a pas d'article à ce moment
// on ne peut valoriser sur rien, on sort
if (TOBA=nil) and (NbArticles<=1) then exit;

// On arrive ici avec une TOB article TOBA, initialisée avec les données de la table ARTICLE
// ou avec la TOBA à nil et un choix à faire dans la liste d'articles

// on peut deja creer notre TOB valo
Result := TOB.Create('ARTICLE',Nil,-1) ;
if (TOBA<>nil) then
   // on initialise avec l'article courant
   Result.Dupliquer(TOBA, true, true, false);

try
// Il faut valoriser suivant le contexte et generer la TOB valo à partir de cette TOBA
Case NbArticles of
        1 : BEGIN
            if (LibArticle='') then
               LibArticle := Result.GetValue('GA_LIBELLE');

            if (FormuleVar='') then
              FormuleVar := Result.GetValue('GA_FORMULEVAR');

            // S'il n'y a qu'un article pour l'affaire
            MajPrixValorises(Result);
            END ;

        else   BEGIN
               // S'il existe plusieurs fois le même article pour l'affaire saisie, il faut choisir dans la liste
               if bAvecZoom then
                    ZoomArticlesD1Affaire( AcsAffaire, codeArticle, sTypeArt, ((Not AffichageValorisation) and ((AcTypeAcces=tacTemps) or (AcTypeAcces=tacGlobal))) ) ;

               if bAvecZoom and (TheTob <> nil) and (TheTOB.TOBNameExist('ARTICLE')) then
                  begin
                  TOBRep := TOB.Create('ARTICLE',Nil,-1) ;
                  TOBRep.Dupliquer(TheTOB.Detail[0], true, true, false);
                  end
               else
                  begin
                  Result.Free;
                  Result:=nil;
                  exit;
                  end;

               try
               CodeArticle := TOBRep.GetValue('GA_CODEARTICLE');
               LibArticle := TOBRep.GetValue('GA_LIBELLE');
               dPrHT := TOBRep.GetValue('GA_PMRP');
               dPvHT := TOBRep.GetValue('GA_PVHT');
               if TOBRep.FieldExists('NUMORDRE') then  //ONYX-AFORMULEVAR
                  numordre := TOBRep.GetValue('NUMORDRE');
               if TOBRep.FieldExists('NUMPIECEORIG') then  // PL le 06/08/03 : pour CB ajout NumPieceOrig
                  NumPieceOrig := TOBRep.GetValue('NUMPIECEORIG');
               if TOBRep.FieldExists('MNTREMISE') then
                  MntRemise := TOBRep.GetValue('MNTREMISE');
               FormuleVar := TOBRep.GetValue('GA_FORMULEVAR');
               if bValoAffPR then AvbPROK:=true;
               if bValoAffPV then AvbPVOK:=true;

               // il n'y a plus qu'un article pour l'affaire
               MajPrixValorises(Result);

               finally
               TOBRep.Free;
               //TOBRep:=nil;
               end;
               END ;
      End;
finally
// Si on ne gère pas la liste des articles, on peut libérer la TOBA créée temporairement
  if (TOBArticles=nil) then TOBA.Free;
end;
end;

function ValoriseLesActivites (TobActs, TOBArticles : TOB; AFOAssistants : TAFO_Ressources; TOBAffaires : TOB; bPR, bPV : boolean):boolean;
  var
  TOBValo, TOBActCourante : TOB;
  UniteValo : string;
  i : integer;
  dCoeffConvert : double;
  bPROK, bPVOK : boolean;
  begin
    TobValo := nil;
    try
    // Sur toutes les lignes d'activite
    for i := 0 to TobActs.detail.Count - 1 do
      begin
        Result := false;
        TOBActCourante := TobActs.detail[i];
        //
        // On valorise
        //
        // On vide la tob valo
        if (TOBValo <> nil) then TOBValo.Free;
        bPROK := true; bPVOK := true;

        // Valorisation
        TOBValo := MajTOBValo (TOBActCourante.GetValue('ACT_DATEACTIVITE'), tacGlobal,
                                TOBActCourante.GetValue('ACT_AFFAIRE'),
                                TOBActCourante.GetValue('ACT_RESSOURCE'),
                                TOBActCourante.GetValue('ACT_CODEARTICLE'),
                                TOBActCourante.GetValue('ACT_TYPEHEURE'),
                                TOBArticles, TOBAffaires, nil, AFOAssistants, false,
                                VH_GC.AFValoActPR, VH_GC.AFValoActPV, bPROK, bPVOK);

        if (TOBValo <> nil) then
          begin
            // L'unité de valorisation est l'unité dans laquelle on a valorisé les prix dans la fonction MajTOBValo
            UniteValo := TOBValo.GetValue ('GA_QUALIFUNITEACT');
            if (TOBActCourante.GetValue ('ACT_UNITE') = '') then
              TOBActCourante.PutValue ('ACT_UNITE', UniteValo);

            // Si ce n'est pas convertible, on ne touche pas aux données saisies
            if IsConvertible (TOBActCourante.GetValue ('ACT_UNITE'), UniteValo) or (UniteValo = '') then
              begin
                dCoeffConvert := 1;
                // On calcule le coefficient de conversion pour passer de l'unité de saisie en unité de valorisation
                if (UniteValo <> TOBActCourante.GetValue ('ACT_UNITE')) and (UniteValo <> '') then
                  begin
                    // il faut convertir les prix exprimés en unité de valorisation, en unité de saisie
                    dCoeffConvert := ConversionUnite (UniteValo, TOBActCourante.GetValue ('ACT_UNITE'), 1);
                    if dCoeffConvert = 0 then dCoeffConvert := 1;
                  end;

                // PUPRCHARGE
                if bPR and bPROK then
                  begin
                    TOBActCourante.PutValue ('ACT_PUPRCHARGE', TOBValo.GetValue ('GA_PMRP') / dCoeffConvert) ;
                    TOBActCourante.PutValue ('ACT_PUPR', TOBValo.GetValue ('GA_PMRP') / dCoeffConvert) ;
                    TOBActCourante.PutValue ('ACT_PUPRCHINDIRECT', TOBValo.GetValue ('GA_PMRP') / dCoeffConvert) ;
                    TOBActCourante.PutValue ('ACT_TOTPRCHARGE', TOBActCourante.GetValue ('ACT_QTE') * TOBActCourante.GetValue ('ACT_PUPRCHARGE'));
                    TOBActCourante.PutValue ('ACT_TOTPR', TOBActCourante.GetValue ('ACT_QTE')*TOBActCourante.GetValue ('ACT_PUPR'));
                    TOBActCourante.PutValue ('ACT_TOTPRCHINDI', TOBActCourante.GetValue ('ACT_QTE') * TOBActCourante.GetValue ('ACT_PUPRCHINDIRECT'));
                  end;

                // PUVENTE
                if bPV and bPVOK then
                  begin
                    TOBActCourante.PutValue ('ACT_PUVENTE', TOBValo.GetValue ('GA_PVHT') / dCoeffConvert);
                    TOBActCourante.PutValue ('ACT_PUVENTEDEV', TOBValo.GetValue ('GA_PVHT') / dCoeffConvert);
                    TOBActCourante.PutValue ('ACT_TOTVENTE', TOBActCourante.GetValue ('ACT_QTE') * TOBActCourante.GetValue ('ACT_PUVENTE'));
                    TOBActCourante.PutValue ('ACT_TOTVENTEDEV', TOBActCourante.GetValue ('ACT_QTE') * TOBActCourante.GetValue ('ACT_PUVENTEDEV'));
                  end;
              end;
          end;

      end;  // Boucle sur les lignes d'Activité


    // On enregistre les lignes d'activité modifiées
    // PL le 31/01/02 : on optimise, on sait dans ce cas qu'on n'aura que des update et pas d'insertion...
    //TobActs.InsertOrUpdateDB(false);
    TobActs.UpdateDB ();
    ////////////////////////
    Result:=true;


    finally
      TOBValo.Free;
    end;

  end;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /    
Description .. : recherche d'un article suivant le type et le code donnés en
Suite ........ : entrée sans savoir si c'est le code unique ou pas qui est
Suite ........ : fourni. 
Suite ........ : TOB article trouvea en sortie 
Mots clefs ... : ARTICLE; CODEARTICLE; RECHERCHE ; GIGA
*****************************************************************}
function TrouverCodeArticle(CodeArticle, sTypeArt:string; TOBArticles:TOB):TOB;
var
Article:string;
rechart : T_RechArt;
TOBA:TOB;
begin
rechart := traOK;
Article := CodeArticleUnique(codeArticle,'','','','','');
TOBA := FindTOBArtSaisAff( TOBArticles, Article, true );
if (TOBA=nil) then
        begin
        TOBA:=CreerTOBArt(TOBArticles);
        TOBA.InitValeurs(FALSE);
        rechart := TrouverArticleSQL_GI(true, Article, TOBA, sTypeArt);
        end;
if (TOBA=nil) then
        begin
        TOBA := FindTOBArtSaisAff(TOBArticles, codeArticle, false );
        if (TOBA=nil) then
            begin
            TOBA:=CreerTOBArt(TOBArticles) ;
            TOBA.InitValeurs(FALSE) ;
            rechart := TrouverArticleSQL_GI(false, codeArticle, TOBA, sTypeArt);
            end;
        end;
if (rechart=traAucun) then
        begin
        TOBA.Free;
        TOBA:=nil;
        end;

Result:=TOBA;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Fonctions recup sur affaire depuis la piece
Suite ........ : renvoie une TOB Articles
Mots clefs ... : ARTICLES; AFFAIRE; TOB ; GIGA
*****************************************************************}
function TOBArticlesD1Affaire(AcsCodeAffaire:string; AcsCodeArticle:string; AcsTypeArticle: String):TOB;
Var i : Integer;
    TOBPiece, TOBArt, TOBLigne : TOB;
    rechart : T_RechArt;
Begin
Result := Tob.Create('Liste Articles affaire',nil,-1);
TobPiece:= TOB.Create('PIECE', nil, -1);
TobPiece.PutValue('GP_AFFAIRE',AcsCodeAffaire);

try
{$IFDEF AFFAIRE}
LectureLignesAffaire(TobPiece,False);
{$endif}
for i:=0 to (TobPiece.Detail.Count-1) do
   begin
   if ( TobPiece.Detail[i].GetValue('GL_NATUREPIECEG')=VH_GC.AFNatAffaire) then
      if ((AcsTypeArticle='') or ((AcsTypeArticle<>'') and (AcsTypeArticle=TobPiece.Detail[i].GetValue('GL_TYPEARTICLE')))) then
      if ((AcsCodeArticle='') or ((AcsCodeArticle<>'') and (AcsCodeArticle=TobPiece.Detail[i].GetValue('GL_CODEARTICLE')))) then
      if ((Result.FindFirst(['GA_CODEARTICLE','GA_LIBELLE','GA_PMRP','GA_PVHT'],
                                 [TobPiece.Detail[i].GetValue('GL_CODEARTICLE'),
                                 TobPiece.Detail[i].GetValue('GL_LIBELLE'),
                                 TobPiece.Detail[i].GetValue('GL_PMRP'),
                                 TobPiece.Detail[i].GetValue('GL_PUHT')],
                                 False))=nil) then
         begin
         TOBArt:=CreerTOBArt(Result) ;
         TOBArt.InitValeurs(FALSE) ;
         rechart := TrouverArticleSQL_GI(false, TobPiece.Detail[i].GetValue('GL_CODEARTICLE'), TOBArt, '');
         if (rechart=traAucun) then TOBArt.Free
         else
            begin
            TOBLigne := TobPiece.Detail[i];
            TOBArt.PutValue('GA_LIBELLE', TOBLigne.GetValue('GL_LIBELLE'));
            TOBArt.PutValue('GA_PMRP', TOBLigne.GetValue('GL_PMRP'));
            TOBArt.PutValue('GA_PVHT', TOBLigne.GetValue('GL_PUHT'));
            TOBArt.PutValue ('GA_FORMULEVAR', TOBLigne.GetValue ('GL_FORMULEVAR')); //ONYX-AFORMULEVAR
            TOBArt.AddChampSupValeur('NUMORDRE', TOBLigne.GetValue ('GL_NUMORDRE'));
            TOBArt.AddChampSupValeur('MNTREMISE', TOBLigne.GetValue ('GL_VALEURREMDEV'));
            TOBArt.AddChampSupValeur('NUMPIECEORIG', EncodeRefPiece (TOBLigne)); // PL le 06/08/03 : pour CB ajout NumPieceOrig
            end;
         end;
   end;

finally
TOBPiece.Free;
end;
End;



{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 26/05/2003
Modifié le ... :   /  /
Description .. : Fonctions recup sur affaire depuis la piece sans les lignes liées
Suite ........ : renvoie une TOB Articles
Mots clefs ... : ARTICLES; AFFAIRE; TOB ; GIGA
*****************************************************************}
function TOBArticlesD1AffaireOnyx (AcsCodeAffaire, AcsCodeArticle, AcsTypeArticle : string) : TOB;
Var i : Integer;
    TOBPiece, TOBArt, TOBLigne : TOB;
    rechart : T_RechArt;
Begin
  Result := Tob.Create ('Liste Articles affaire', nil, -1);
  TobPiece := TOB.Create ('PIECE', nil, -1);
  TobPiece.PutValue ('GP_AFFAIRE', AcsCodeAffaire);

  try
  {$IFDEF AFFAIRE}
  LectureLignesAffaire (TobPiece, False);
  {$endif}
  for i := 0 to (TobPiece.Detail.Count - 1) do
    begin
      if (TobPiece.Detail[i].GetValue ('GL_NATUREPIECEG') = VH_GC.AFNatAffaire) then
      if ((AcsTypeArticle='') or ((AcsTypeArticle<>'') and (AcsTypeArticle=TobPiece.Detail[i].GetValue('GL_TYPEARTICLE')))) then
      if ((AcsCodeArticle='') or ((AcsCodeArticle<>'') and (AcsCodeArticle=TobPiece.Detail[i].GetValue('GL_CODEARTICLE')))) then
        if (TobPiece.Detail[i].GetValue ('GL_LIGNELIEE') = 0) then
          begin
            TOBArt := CreerTOBArt (Result);
            TOBArt.InitValeurs (FALSE);
            rechart := TrouverArticleSQL_GI (false, TobPiece.Detail[i].GetValue ('GL_CODEARTICLE'), TOBArt, '');
            if (rechart = traAucun) then TOBArt.Free
            else
              begin
                TOBLigne := TobPiece.Detail[i];
                TOBArt.PutValue ('GA_LIBELLE', TOBLigne.GetValue ('GL_LIBELLE'));
                TOBArt.PutValue ('GA_PMRP', TOBLigne.GetValue ('GL_PMRP'));
                TOBArt.PutValue ('GA_PVHT', TOBLigne.GetValue ('GL_PUHT'));
                TOBArt.PutValue ('GA_FORMULEVAR', TOBLigne.GetValue ('GL_FORMULEVAR')); //ONYX-AFORMULEVAR
                TOBArt.AddChampSupValeur('NUMORDRE', TOBLigne.GetValue ('GL_NUMORDRE'));
                TOBArt.AddChampSupValeur('MNTREMISE', TOBLigne.GetValue ('GL_VALEURREMDEV'));
                TOBArt.AddChampSupValeur('NUMPIECEORIG', EncodeRefPiece (TOBLigne)); // PL le 06/08/03 : pour CB ajout NumPieceOrig
              end;
          end;
    end;

  finally
    TOBPiece.Free;
  end;
End;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : renvoie le dernier numéro d'une ligne d'activite pour une clé  (ancienne clé)
Suite ........ : donnée en entrée (dans la tob activite)
Mots clefs ... : ACTIVITE; MAX ; GIGA
*****************************************************************}
function MaxNumLigneActivite(TobLigneAct:TOB):integer;
var
sReq:string;
Q:TQuery;
begin
Result:=0;

// On recherche le max pour cette clé
sReq := 'SELECT MAX(ACT_NUMLIGNE) FROM ACTIVITE WHERE ACT_TYPEACTIVITE="'+TobLigneAct.GetValue('ACT_TYPEACTIVITE')
        + '" AND ACT_AFFAIRE="'+TobLigneAct.GetValue('ACT_AFFAIRE')
        + '" AND ACT_RESSOURCE="'+TobLigneAct.GetValue('ACT_RESSOURCE')
        + '" AND ACT_DATEACTIVITE="'+UsDateTime(TobLigneAct.GetValue('ACT_DATEACTIVITE'))
        + '" AND ACT_TYPEARTICLE="'+TobLigneAct.GetValue('ACT_TYPEARTICLE')
        + '"';
Q := nil;
try
Q:=OpenSQL(sReq, True,-1,'',true);
if Not Q.EOF then
    begin
    Result := Q.Fields[0].AsInteger;
    end;
finally
Ferme(Q);
end;
end;


// *************  Heures supplémentaires ACT_TYPEHEURE *************************
Function GetTauxHSup (stTaux : String; Var Tx1,Tx2 : Double): Boolean;
Var unTaux : String;
BEGIN
Result := True;
Tx1 := 0; Tx2 := 0;
if stTaux <> '' then
   BEGIN
   unTaux := ReadTokenSt (stTaux);
   if unTaux <> '' then
      BEGIN
      if (isFloat(unTaux)) or (isNumeric(UnTaux)) then Tx1 := StrToFloat(unTaux) else Result := False;
      END else Result := False;
   // taux de vente facultatif
   unTaux := ReadTokenSt (stTaux);
   if unTaux <> '' then
      BEGIN
      if (isFloat(unTaux)) or (isNumeric(UnTaux)) then Tx2 := StrToFloat(unTaux) else Result := False;
      END
   else Tx2 := TX1;  // taux de vente = taux de revient par défaut
   END;
END;

Function  TauxHSupFromCode (codeTT : string; Var Tx1,Tx2 : Double): Boolean;
Var stTaux : string;
BEGIN
Tx1 := 0; Tx2 := 0;
Result := False;
if CodeTT = '' then Exit;
stTaux := RechDom ('AFTTYPEHEURE',CodeTT,True);
Result := GetTauxHSup (stTaux, Tx1, Tx2);
// Passage de pourcentage en coefficient
Tx1 := 1 + Tx1/100;
Tx2 := 1 + Tx2/100;
END;


{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /    
Description .. : Calcul les dates de début et de fin de saisie d'activité 
Suite ........ : suivant le contexte de saisie (paramètres)
Mots clefs ... : DATE; ACTIVITE ; GIGA
*****************************************************************}

procedure IntervalleDatesActivite(var AvdDateDebut, AvdDateFin : TDateTime);
var
iAnneeDateFin, iMoisDateFin, iJourDateFin:word;
begin
AvdDateDebut:=iDate1900;
AvdDateFin:=iDate2099;
if (GetParamSoc('SO_AFDateDebutAct')<>0) and (GetParamSoc('SO_AFDateDebutAct')<>iDate1900) and (GetParamSoc('SO_AFDateDebutAct')<>iDate2099) then
    AvdDateDebut:=GetParamSoc('SO_AFDateDebutAct');

if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='DAF') then
    begin
    AvdDateFin:=VH_GC.AFDateFinAct;
    end
else
    begin
    if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='FMC') then
        begin
        // Fin du mois courant
        AvdDateFin:=FINDEMOIS(V_PGI.DateEntree);
        end
    else
    if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='FM1') then
        begin
        // Fin du mois courant + 1
        DecodeDate( V_PGI.DateEntree, iAnneeDateFin, iMoisDateFin, iJourDateFin);
        iJourDateFin:=1;
        iMoisDateFin:=iMoisDateFin+1;
        if  (iMoisDateFin>12) then begin iMoisDateFin:=1; iAnneeDateFin:=iAnneeDateFin+1; end;


        AvdDateFin:=FINDEMOIS(EncodeDate(iAnneeDateFin, iMoisDateFin, iJourDateFin));
        end
    else
    if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='FAC') then
        begin
        // Fin d'annee courante
        DecodeDate( V_PGI.DateEntree, iAnneeDateFin, iMoisDateFin, iJourDateFin);
        AvdDateFin:=FINDEMOIS(EncodeDate(iAnneeDateFin, 12, 1));
        end;
    end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /    
Description .. : fonction test du blocage ou non de la saisie d'activité 
Suite ........ : suivant le contexte (paramétrage, et différents blocages sur 
Suite ........ : les dates)
Mots clefs ... : BLOCAGE; DATE; ACTIVITE ; GIGA
*****************************************************************}
function TestBlocageDates(AcdDate:TDateTime; AciInitResult:integer):integer;
var
iMoisCourant, iMoisDateFin, iAnneeCourante, iAnneeDateFin, iJourCourant, iJourDateFin:word;
iTestCourant, iTestFin, iMoisPlus1, iAnneMoisPlus1 : word;
begin
Result:=AciInitResult;
//
// Gestion du blocage sur l'intervalle de dates de la saisie d'activité
//
if (GetParamSoc('SO_AFDateDebutAct')<>0) and (GetParamSoc('SO_AFDateDebutAct')<>iDate1900) and (GetParamSoc('So_AFDateDebutAct')<>iDate2099) then
    begin
    if (AcdDate<GetParamSoc('SO_AFDateDebutAct')) then
         begin
         Result := 2;
         end;
    end;

if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='DAF') then
    begin
    if (VH_GC.AFDateFinAct<>0) and (VH_GC.AFDateFinAct<>iDate1900) and (VH_GC.AFDateFinAct<>iDate2099) then
        if (AcdDate>VH_GC.AFDateFinAct) then
            begin
            Result := 3;
            end;
    end
else
    begin
    DecodeDate( V_PGI.DateEntree, iAnneeDateFin, iMoisDateFin, iJourDateFin);
    DecodeDate( AcdDate, iAnneeCourante, iMoisCourant, iJourCourant);
    if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='FMC') then
        begin
        iTestCourant := (iAnneeCourante * 100) +  iMoisCourant;
        iTestFin := (iAnneeDateFin * 100) +  iMoisDateFin;
        if (iTestCourant>iTestFin) then
            begin
            Result := 4;
            end;
        end
    else
    if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='FM1') then
        begin
        iMoisPlus1 := iMoisDateFin+1;
        iAnneMoisPlus1 := iAnneeDateFin;
        if (iMoisPlus1>12) then begin iMoisPlus1:=1; iAnneMoisPlus1:=iAnneMoisPlus1+1; end;

        iTestCourant := (iAnneeCourante * 100) +  iMoisCourant;
        iTestFin := (iAnneMoisPlus1 * 100) +  iMoisPlus1;
        if (iTestCourant>iTestFin) then
            begin
            Result := 5;
            end;
        end
    else
    if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='FAC') then
        begin
        if (iAnneeCourante>iAnneeDateFin) then
            begin
            Result := 6;
            end;
        end;
    end;



end;




{***********A.G.L.***********************************************
Auteur  ...... : LENORMAND Pierre
Créé le ...... : 11/09/2001
Modifié le ... :   /  /    
Description .. : Fonction de traitement des objets activité gérées d'un écran 
Suite ........ : à l'autre
Mots clefs ... : ACTIVITE ; GIGA
*****************************************************************}
Function TraiteObjActivite ( Num : integer; PRien : THPanel): Boolean;
begin
  Result := True;                                            
  if Not ((Num - 72000) in [102..105]) and Not ((Num - 72000) in [111..114])
      and Not (Num = 76170)  // PL le 16/04/03 : les menu saisie activité du bureau GED (MD)  // PL le 29/07/03 : suppression du Not (Num = 188210) à la demande de Mister Marco
      then
      // On n'est pas en saisie d'activité
    begin
      // Il faut libérer VH_GC.AFTobInterne qui est utilisée en saisie d'activité pour passer les copier/coller
      // d'un type de saisie à un autre. On ne delete ici que la sous-TOB comprenant de l'ACTIVITE
      if (VH_GC.AFTobInterne <> nil) then
        DetruitMaTobInterne ('Mon Activite');
    end
  else
      // On est en saisie d'activité
    begin
      // Dans le cas où l'on est déjà en saisie d'activité et qu'on cherche à réouvrir un écran de saisie
      // d'activité par clique droit, on sort
      if (PRien.InsideForm  <> nil) then
          if (PRien.insideform.name = 'FActivite') then Result := False;
    end;
end;

(*
// PL le 23/05/02 : réparation de la fonction
function PremierJourSemaineAGL540i(AciSemaine:integer; AciAnnee:integer):TDateTime;
var
dDateCourante:TDateTime;
iInc:integer;
begin
Result := 0;
if (AciSemaine<1) or (AciSemaine>53) then exit;
if (AciAnnee<1900) or (AciAnnee>2100) then exit;

dDateCourante := EncodeDate(AciAnnee, 1,1);
iInc:=1;
if Not V_PGI.Semaine53 then iInc:=-1;

while DayOfWeek(dDateCourante)<>2 do dDateCourante := dDateCourante + iInc;

Result := dDateCourante + (AciSemaine-1) * 7;
end;


////////////////////
// PL le 23/05/02 : réparation de la fonction
function PremierJourSemaineTempo1(AciSemaine:integer; AciMois:integer; AciAnnee:integer):TDateTime;
var
dPremierJourAnnee,dPremierJourSemaine1:TDateTime;
AnneeReelle:integer;
dPremierJeudi:TDateTime;
begin
Result := 0;
if (AciSemaine<1) or (AciSemaine>53) then exit;
if (AciMois<1) or (AciMois>12) then exit;
if (AciAnnee<1900) or (AciAnnee>2100) then exit;

AnneeReelle := AciAnnee;
if (AciMois=12) and (AciSemaine=1) then AnneeReelle := AnneeReelle + 1;

dPremierJourAnnee := EncodeDate(AnneeReelle, 1,1);
dPremierJeudi := dPremierJourAnnee;
while DayOfWeek(dPremierJeudi)<>5 do dPremierJeudi := dPremierJeudi + 1 ;
dPremierJourSemaine1 := dPremierJeudi - 3;

if V_PGI.Semaine53 and (dPremierJourSemaine1<dPremierJourAnnee) and (AciSemaine=1) then
    dPremierJourSemaine1:=dPremierJourAnnee;

Result := dPremierJourSemaine1 + (AciSemaine-1) * 7;
end;
////////////////////


{
NumSemaine retourne une valeur de 1 à 7 (1 = Dimanche 7 = Samedi)
les semaines ISO 8601 commencent avec le lundi
et la première semaine est celle qui inclue le premier jeudi
le tableau Fiddle résume tout ca.
}
function NumSemaine1(today: Tdatetime ): word;
const Fiddle : array[1..7] of Byte = (6,7,8,9,10,4,5);
var
 present, startOfYear: Tdatetime;
 firstDayOfYear, weekNumber, numberOfDays: integer;
 year, month, day: word;
 YearNumber: string;
begin
 present:= trunc(today); // enleve la notion d'heure
 decodeDate(present, year, month, day); // decode la date JJ MM AAAA
 startOfYear:= encodeDate(year, 1, 1); // encode date 01/01/AAAA

 // trouve quel jour de la semaine est le 01/01 façon ISO
 firstDayOfYear:= Fiddle[dayOfWeek(startOfYear)];

  if (month=1) and ( (dayOfWeek(startOfYear)>3) and ((dayOfWeek(startOfYear)<7-(day-1))) ) then
    begin
        // jours avant la semaine 1 de l'année courante ont le meme num de semaine
        // que le dernier jour de l'année précédente
        present     := startOfYear - 1;
        startOfYear := encodeDate(year-1, 1, 1);
       firstDayOfYear:= Fiddle[dayOfWeek(startOfYear)];
        month       := 12;
        day         := 31;
    end;

// calcul nombre de jour depuis le debut de l'année vs date passée en param
numberOfDays:= trunc(present - startOfYear) + firstDayOfYear;

  // calcul nbumero de semaine
  weekNumber:= trunc(numberOfDays / 7);
  result := weekNumber;

  if weekNumber = 0 then // récursivité
    // voir si année précédente se termine en 52 ou 53 eme semaine
    weekNumber:= NumSemaine(encodeDate(year - 1, 12, 31))
  else if weekNumber = 53 then
    // si 31/12 inféerieur au jeudi : semaine 1 de l'année suivante
    if (not V_PGI.Semaine53) and (dayOfWeek(encodeDate(year, 12, 31)) < 5) then
      weekNumber := 01;

  result := weekNumber;
end;


Function  NumSemaine2( D : TDateTime) : Integer ;
var AYear, AWeek, ADay : integer;
begin
case PremiereSemaine of
//case V_PGI.PremiereSemaine of
  0 : // 1ere semaine = commence le 1er janvier
    DefaultWeekDef := WeekDefMon1stJan;
  1 : // 1ere semaine = 1ere semaine de 4 jours
    DefaultWeekDef := WeekDefISO8601;
  2 : // 1ere semaine = 1ere semaine entiere
    DefaultWeekDef := WeekDefMon7thJan;
end;
DateToWeek(D,AYear, AWeek, ADay, DefaultWeekDef);
if AWeek = 0 then // récursivité // voir si année précédente se termine en 52 ou 53 eme semaine
  AWeek:= NumSemaine2(encodeDate(AYear - 1, 12, 31))
// si 31/12 inféerieur au jeudi : semaine 1 de l'année suivante
else if (AWeek = 53) and (not V_PGI.Semaine53) and (dayOfWeek(encodeDate(AYear, 12, 31)) < 5) then
  AWeek := 01;
result := AWeek;
end;




////////////////////
// PL le 23/05/02 : réparation de la fonction
//  0 :  1ere semaine = contient le 1er janvier
//  1 :  1ere semaine = contient 1er jeudi
//  2 :  1ere semaine = contient 1er lundi
//function PremierJourSemaineTempo(AciSemaine:integer; AciMois:integer; AciAnnee:integer):TDateTime;
function PremierJourSemaineTempo(AciSemaine:integer; AciAnnee:integer):TDateTime;
const
FDW = 2; // le premier jour de la semaine est le lundi
var
dPremierJourAnnee,dPremierJourSemaine1:TDateTime;
AnneeReelle:integer;
dPivot:TDateTime;
begin
Result := 0;
if (AciSemaine<1) or (AciSemaine>53) then exit;
//if (AciMois<1) or (AciMois>12) then exit;
if (AciAnnee<1900) or (AciAnnee>2100) then exit;

//AnneeReelle := AciAnnee;
//if (AciMois=12) and (AciSemaine=1) then AnneeReelle := AnneeReelle + 1;
//if (AciMois=1) and ((AciSemaine=52) or (AciSemaine=53)) then AnneeReelle := AnneeReelle - 1;
//dPremierJourAnnee := EncodeDate(AnneeReelle, 1,1);

dPremierJourAnnee := EncodeDate(AciAnnee, 1,1);

case PremiereSemaine of
//case V_PGI.PremiereSemaine of
  0 : //  1ere semaine = contient le 1er janvier
        begin
        dPivot := dPremierJourAnnee;
        while DayOfWeek(dPivot)<>FDW do dPivot := dPivot - 1 ;
        dPremierJourSemaine1 := dPivot;
        end;
  1 : // 1ere semaine = contient 1er jeudi
        begin
        dPivot := dPremierJourAnnee;
        while DayOfWeek(dPivot)<>(FDW+3) do dPivot := dPivot + 1 ;
        dPremierJourSemaine1 := dPivot - 3;  // lundi par rapport au jeudi => -3 ou début de première semaine de 4 jours
        end;
  2 : // 1ere semaine = contient 1er lundi
        begin
        dPivot := dPremierJourAnnee;
        while DayOfWeek(dPivot)<>FDW do dPivot := dPivot + 1 ;
        dPremierJourSemaine1 := dPivot;
        end;
end;


Result := dPremierJourSemaine1 + (AciSemaine-1) * 7;
end;
////////////////////

*)

{$IFDEF AGLINF545}
 // PL le 23/05/02 : réparation de la fonction
function PremierJourSemaineTempo(AciSemaine:integer; AciAnnee:integer):TDateTime;
var
dDateCourante:TDateTime;
iInc:integer;
begin
Result := 0;
if (AciSemaine<1) or (AciSemaine>53) then exit;
if (AciAnnee<1900) or (AciAnnee>2100) then exit;

dDateCourante := EncodeDate(AciAnnee, 1,1);
iInc:=1;
if Not V_PGI.Semaine53 then iInc:=-1;

while DayOfWeek(dDateCourante)<>2 do dDateCourante := dDateCourante + iInc;

Result := dDateCourante + (AciSemaine-1) * 7;
end;
////////////////////
{$ELSE}
{$ENDIF}
(* dernière version de la fonction NumSemaine, à compiler avec NumSemaine dans les uses
Function  MyNumSemaine( D : TDateTime ; WithSemaine53 : boolean = true) : Integer ;
var AYear, AWeek, ADay : integer;
begin
case V_PGI.Semaine53 of
  0 : // 1ere semaine = commence le 1er janvier
//    DefaultWeekDef := WeekDefMon1stJan;
    DefaultWeekDef := WeekDefMon1stJan;
  1 : // 1ere semaine = 1ere semaine de 4 jours
//    DefaultWeekDef := WeekDefISO8601;
    DefaultWeekDef := WeekDefISO8601;
  2 : // 1ere semaine = 1ere semaine entiere
//    DefaultWeekDef := WeekDefMon7thJan;
    DefaultWeekDef := WeekDefMon7thJan;
end;
DateToWeek(D,AYear, AWeek, ADay, DefaultWeekDef);
if AWeek = 0 then // récursivité // voir si année précédente se termine en 52 ou 53 eme semaine
  AWeek:= HEnt1.NumSemaine(encodeDate(AYear - 1, 12, 31))
// si 31/12 inféerieur au jeudi : semaine 1 de l'année suivante
else if (AWeek = 53) and (not WithSemaine53) and (dayOfWeek(encodeDate(AYear, 12, 31)) < 5) then
  AWeek := 01;
result := AWeek;
end;
*)

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 11/03/2003
Modifié le ... :   /  /
Description .. : renvoie la tob d'une ligne d'activite à partir de son numéro unique
Suite ........ : et nil si rien
Mots clefs ... : ACTIVITE; GIGA
*****************************************************************}
function LaTobDeLaLigneUnique (sTypeActivite, sAffaire, sNumLigneUnique : string; bRestreint : boolean = true) : TOB;
  var
  sReq : string;
  Q : TQuery;
begin

  if bRestreint then
    sReq := 'SELECT ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_RESSOURCE,ACT_DATEACTIVITE,ACT_FOLIO'
            + ',ACT_TYPEARTICLE,ACT_NUMLIGNE,ACT_NUMLIGNEUNIQUE,ACT_DATEMODIF,ACT_DATEVISA'
            + ',ACT_DATEVISAFAC,ACT_FORMULEVAR,ACT_UNITE,ACT_ACTIVITEEFFECT,ACT_ACTIVITEREPRIS,ACT_QTE'
            + ' FROM ACTIVITE WHERE ACT_TYPEACTIVITE="'
  else
    sReq := 'SELECT * FROM ACTIVITE WHERE ACT_TYPEACTIVITE="';

  // On recherche la ligne pour cette clé
  sReq := sReq + sTypeActivite
          + '" AND ACT_AFFAIRE="' + sAffaire
          + '" AND ACT_NUMLIGNEUNIQUE=' + sNumLigneUnique;
  Q := nil;

  Result := Tob.Create ('ACTIVITE', nil, -1);

  try
  Q := OpenSQL (sReq, True,-1,'',true);
  if Not Result.SelectDB ('', Q) then
    begin
      Result.free;
      Result := nil;
    end;

  finally
    Ferme (Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 11/03/2003
Modifié le ... :   /  /
Description .. : renvoie le plus grand numéro unique d'une ligne d'activite pour une clé
Suite ........ : donnée en entrée et 0 si rien
Mots clefs ... : ACTIVITE; MAX ; GIGA
*****************************************************************}
function MaxNumLigneUniqueActivite (sTypeActivite, sAffaire : string) : integer;
  var
  sReq : string;
  Q : TQuery;
begin
  Result := 0;

  // On recherche le max pour cette clé
  sReq := 'SELECT MAX(ACT_NUMLIGNEUNIQUE) FROM ACTIVITE WHERE ACT_TYPEACTIVITE="' + sTypeActivite
        + '" AND ACT_AFFAIRE="' + sAffaire
        + '"';
  Q := nil;
  try
  Q := OpenSQL (sReq, True,-1,'',true);
  if Not Q.EOF then
    begin
      Result := Q.Fields[0].AsInteger;
    end;
  finally
    Ferme (Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : renvoie le prochain numéro utilisable unique et positif d'une ligne d'activite pour une clé
Suite ........ : donnée en entrée et 1 si rien
Mots clefs ... : ACTIVITE; MAX ; GIGA
*****************************************************************}
function ProchainPlusNumLigneUniqueActivite (sTypeActivite, sAffaire : string) : integer;
begin

  Result := MaxNumLigneUniqueActivite (sTypeActivite, sAffaire) + 1;
  if (Result <= 0) then
    // S'il n'y a pas de max positif on renvoit 1
    Result := 1;

end;


{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : renvoie le plus petit numéro unique d'une ligne d'activite pour une clé
Suite ........ : donnée en entrée et 0 si rien
Mots clefs ... : ACTIVITE; MIN ; GIGA
*****************************************************************}
function MinNumLigneUniqueActivite (sTypeActivite, sAffaire : string) : integer;
  var
  sReq : string;
  Q : TQuery;
begin
  Result := 0;

  // On recherche le max pour cette clé
  sReq := 'SELECT MIN(ACT_NUMLIGNEUNIQUE) FROM ACTIVITE WHERE ACT_TYPEACTIVITE="' + sTypeActivite
        + '" AND ACT_AFFAIRE="' + sAffaire
        + '"';
  Q := nil;
  try
  Q := OpenSQL (sReq, True,-1,'',true);
  if Not Q.EOF then
    begin
      Result := Q.Fields[0].AsInteger;
    end;
  finally
    Ferme (Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : renvoie le prochain numéro utilisable unique et négatif d'une ligne d'activite pour une clé
Suite ........ : donnée en entrée et -1 si rien
Mots clefs ... : ACTIVITE; MIN ; GIGA
*****************************************************************}
function ProchainMoinsNumLigneUniqueActivite (sTypeActivite, sAffaire : string) : integer;
begin

  Result := MinNumLigneUniqueActivite (sTypeActivite, sAffaire) - 1;
  if (Result >= 0) then
    // Si il n'y a pas de min négatif on renvoit -1
    Result := -1;

end;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. : renvoie le prochain numéro unique d'une ligne d'activite pour une clé TYPEACTIVITE , AFFAIRE
Suite ........ : donnée en entrée et 0 si rien.
                  DansLesNegatifs indique si on fait progresser l'indice dans les nombres négatifs. On est dans les positifs par defaut.
                  AvecPreinit indique si on veut préinitialiser avec les indices préexistants pour chaque affaire.
Mots clefs ... : NUMLIGNEUNIQUE; ACTIVITE; GIGA;
*****************************************************************}
function ProchainIndiceAffaires (sTypeActivite, sAffaire : string; ListeAffaires : TStringList; DansLesNegatifs : boolean = false; AvecPreinit : boolean = true) : integer;
  var
  IndexAffaire, Indice : integer;
begin
  Result := 0;
  Indice := 1;
  if DansLesNegatifs then Indice := -1;
  if ListeAffaires = nil then exit;


  IndexAffaire := ListeAffaires.IndexOf (sTypeActivite + sAffaire);
  if (IndexAffaire = -1) then
    // L'affaire n'existe pas encore dans la liste
    begin
      // Si on gere une préinitialisation de l'indice suivant l'affaire
      if AvecPreinit then
        begin
          if not DansLesNegatifs then
            // Si on est dans les positifs
              // recherche du prochain indice positif pour cette affaire
              Indice := ProchainPlusNumLigneUniqueActivite (sTypeActivite, sAffaire)
          else
            // Si on est dans les négatifs
              // recherche du prochain indice négatif pour cette affaire
              Indice := ProchainMoinsNumLigneUniqueActivite (sTypeActivite, sAffaire);
        end;

      ListeAffaires.AddObject (sTypeActivite + sAffaire, TObject(Indice));

      Result := Indice;
    end
  else
    // L'affaire existe, on incremente le nombre d'affaires associé
    begin
      Result := integer(ListeAffaires.Objects [IndexAffaire]) + Indice;

      ListeAffaires.Objects [IndexAffaire] := TObject(Result);
    end;

end;

end.
