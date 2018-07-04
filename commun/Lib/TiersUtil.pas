unit TiersUtil;

interface

uses
  HEnt1,
  sysutils,
  UTOB,
  ParamSoc,
  Ent1,
  StdCtrls,
  Controls,
  Classes,
  forms,
  ComCtrls,
  HCtrls,
  HMsgBox,
  Hdimension,
  UTOM,
  {$IFDEF EAGLCLIENT}
    MaineAGL,
  {$ELSE EAGLCLIENT}
    db,
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
    DBGrids,
    {$IFNDEF EAGLSERVER}
      {$IFNDEF ERADIO}
        Fe_Main,
      {$ENDIF !ERADIO}
    {$ENDIF EAGLSERVER}
  {$ENDIF EAGLCLIENT}
  wCommuns,
  variants,
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
    LookUp,
    AGLInit,
  {$ENDIF !ERADIO}
  Web,
{$ENDIF EAGLSERVER}
  M3FP,
  {$IFDEF PGISIDE}
  UtilGC,
  {$ENDIF PGISIDE}
  UtilPGI
{$IFDEF GCGC}
  ,EnvironnementUtil
{$ENDIF GCGC}
{$IFDEF MODENT1}
  , CPTypeCons
{$ENDIF MODENT1}
,uEntCommun
,UFonctionsCBP
  ;

Type Info_Tiers = RECORD
                  Libelle : String35 ;
                  Adresse1,Adresse2,Adresse3,Ville,Telephone : String35 ;
                  CodePostal : String[9] ;
                  Pays : String3 ;
                  END ;
    TypeAdresse = (taUnknow, taLivr, taFact, taRegl);

Function  CreationTiers ( InfTiers : Info_Tiers; var LeRapport : string ; TypAux : String =''; LeNumero : String=''; bParle : Boolean=True) : String ;

Function TiersAuxiliaire (Code: String ; AuxiliaireVersTiers : boolean = FALSE; NatureAuxi : string=''): String;
Function RechZoneTiers (ChampRetourne, Nomchamp, Code, NatureAuxi : String) : String;
Function  VerifEstSeparateur(const Texte : string; const Pos : integer) : Boolean ;
Function  VerifSupDoubleSep(const Texte : string) : string ;
Function  verifSupDoubleEspace(const Texte : string) : string ;
Function  MetTexteEnForme(const Texte, Format : string; SupEsp : boolean) : string;
Function  MetNomEnForme(const Texte, Format : string; Formatpart : string; const TailleMaxPart : integer; SupEsp : boolean) : string;
Function  VerifAppliqueFormat( NomChamp, ValeurChamp : string; clipart : boolean ) : variant ;
Function  VerifFormateChamp( parms: array of variant; nb: integer ) : variant ;
procedure MajPhonetiqueTiers ;
// mcd 15/07/02 passage de affaireUtil dans TiersUtil
Function GetChampsComplementTiers (Codetiers,NomChamp : String) : String;

{$IFDEF GCGC}
Function  GetEtatRisqueClient (CodeTiers : string): string; overload;
Function  GetEtatRisqueClient (TobTiers : Tob): string; overload;
// GC_20080606_DM_GC16079 DEBUT
//function RisqueTiersGC(TOBTiers : TOB ; sSocietes : string = ''; TobTiersParam : TOB=nil) : Double ;
function RisqueTiersGC(TOBTiers : TOB ; sSocietes : string = ''; TobTiersParam : TOB=nil; Typenattiers : string = ''; QuelTiers : string = '') : Double ;
// GC_20080606_DM_GC16079 FIN
{$ENDIF}
function RisqueTiersCPTA(TOBTiers : TOB ; DateB : TDateTime ; sSocietes : string=''; TobTiersParam : TOB=nil) : Double ;
//CEGID-CCMX le 15/09/2006 : ajout iNbJoursEncours  {Evolution}
//  function  EnCoursRegleNonEchu(Aux : String ; DateButoir : tDateTime ; sSocietes : string = '') : Double ;
function  EnCoursRegleNonEchu(Aux : String ; DateButoir : tDateTime ; iNbJoursEnCours : Integer; sSocietes : string = '') : Double ;
//  Function  EnCoursEchuNonRegle (Aux : String ; DateButoir : tDateTime ; sSocietes : string = ''): double;
Function  EnCoursEchuNonRegle (Aux : String ; DateButoir : tDateTime ; iNbJoursEncours : Integer; sSocietes : string = ''): double;
//CEGID-CCMX le 15/09/2006 : ajout iNbJoursEncours  {Evolution}
function CalculeEnCours(TobTiers : TOB; bEntreeMultiSoc : boolean; var RisqueGC, RegleNonEchu, EchuNonRegle, RisqueCPTA : double) : double;
function TiersDebitMoinsCredit (TobTiers : Tob; TQ : TQuery; sSocGroupe : string = '') : double;
function TiersEncoursMultiSoc : boolean;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
  procedure GetTiersLookUp (G_CodeTiers: THCritMaskEdit; stWhere : String);
  function  GetTiersRecherche (G_CodeTiers : THCritMaskEdit; Const stWhere, stRange, stMul: string) : string;
  procedure AfficheRisqueClient (CodeTiers : string ); overload ;
  procedure AfficheRisqueClient (TobTiers : Tob); overload ;
  procedure AfficheRisqueClientDetail (CodeTiers : string ); overload ;
  procedure AfficheRisqueClientDetail (TobTiers : Tob); overload ;
  { Appel de la fiche Tiers }
  procedure CallTiers(Lequel: string; Action:TActionFiche = taModif; Params: String = ''; Range : String = '');
  { Ouverture de la fiche des Adresses des Tiers }
  procedure CallAdressesTiers(Range, Lequel, Params: string);
  { Ouverture de la fiche des adresses en eAGL }
  procedure AglFicheAdressesTiers(parms: array of variant; nb: integer);
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}


procedure CalculSoldesAuxi ( Auxi : String ) ;


Function  TiersAvecAffaireActive ( CodeTiers : String):Boolean;

// gestion des adresses du tiers
function ModifierAdrFillesParMere( AcsAdrMere:string ):boolean;
function ModifierAdrMereParFille( AcsAdrFille:string ):boolean;

{ Test l'existence d'un tiers }
function ExistTiers(sNatureAuxi, sCodeTiers: string; sWhereCompl: string = ''; WithAlert: Boolean = false): Boolean;
function ExistT(Const Tiers: String): Boolean;
function ExistTiersMulti(sNatureAuxi, sCodeTiers: string; sWhereCompl: string = ''; WithAlert: Boolean = false): Boolean;
function ExisteCodeTiers (NatureAuxi,Tiers : string) : boolean;

{ Construction du WHERE pour la fiche TIERS }
function  WhereTiers(sNatureAuxi, sCodeTiers : string): string;
function WhereT(Const Tiers : string): string;
function WhereNature(sNatureAuxi : string): string;
{ Construction du WHERE pour la fiche TIERSCOMPL }
function  WhereTiersCompl(sNatureAuxi, sCodeTiers : string): string;
function WhereYTC(Const Tiers : string): string;
{ Retourne le dernier N° d'adresse pour un tiers }
function  GetTiersLastAdresseNum(CodeTiers, NatureAuxi: String): Integer;

{ Renvoie une liste de champ depuis la fiche Tiers }
function GetFieldsFromTiers(Fields: Array of string; sNatureAux, sTiers: string): MyArrayValue;
{ Charge de la Tob avec un ou plusieurs tiers suivant le where }
function GetTobTiers(sChamp, sWhere: String; TobTiers: Tob; SelectDB:Boolean = False; sOrderBy: string=''): Boolean;
{ Test l'existence d'un représentant - Commercial }
function ExistCommercial(sCodeCommercial : string; WithAlert: Boolean = false):Boolean;

{ Renvoie field de TIERS ou de TIERSCOMPL }
function GetFieldFromTiers(Field, sNatureAux, sTiers: string): String;
function GetFieldFromT(Const FieldName, Tiers: string): Variant;
{ Retourne le numéro de l'adresse par défaut d'un tiers }
function GetNumAdresseFromTiers(NatureAuxi, Tiers: String; TheType: TypeAdresse): String;
{ Recherche du pays d'un code postal ou d'une ville }
function GetFieldFromCodePostal(Field, sCodePostal, sVille: string): Variant;
function GetNomsFromContact(Fields: Array of string; Auxiliaire: string; NumeroContact: integer): MyArrayValue;
function GetFieldFromContact(Field, sAuxiliaire, where: string): Variant;
function GetTiersFromCB(const NatureAuxi, CodeBarre: String): String;

function GetNomsFromTypeContact(Fields: Array of string; Auxiliaire: string; NumeroContact: integer; TypeContact :String): MyArrayValue; // FS N° GC14693

function RechercheTiersDepot(Depot: string): string;
procedure InitTiersCompl(Auxiliaire, CodeTiers: String);

{ Affectation de stock }
function GetRefAffectationT(Const Tiers: string): string;
function GetTiersFromRefAffectation(RefAffectation: string): string;

{ Creation de tiers via une tob externe }
function CreateTiersFromTob(Tiers: String; TheTob: Tob): String;
function CreateOrUpdateTiersFromTob(Tiers: String; TheTob: Tob): String;
function CreateOrUpdateTiers(NewTiers : boolean; Tiers: String; TheTob: Tob): String;
{$IFDEF DP}
Function CreeUnAnnuaire(CodeTiers : string;VientDuTiers:boolean=false):string;
{$endif}

function RechercheTiersFromAuxi(LeAuxi, LeAuxiOrig, LeTiers : string) : string;
function AutoriseCreationTiers (NatureAuxi : string) : boolean;
function GetNomChampTiersParam(NatTiers, PrefixeWith_ : string) : string;

{$IFNDEF EAGLSERVER}
Procedure LanceGoogleMaps( Adresse, Ville : string; BPlan : boolean);
{$ENDIF EAGLSERVER}

procedure VerifDoublonSiretPays (TheTom : TOM ; Siret : String ; Pays : String ; NatureAuxi : String ; Auxiliaire : String ; GuidPer : String ; ControlDoublon : String ; var LastError : Integer ; var DoublonValue : String ) ;
procedure MajProspectClient ( TOBGenere : Tob ; TOBTiers : Tob);
procedure TransformeProspectClient ( TOBGenere : Tob ; NomChamp : String);
function AFMajProspectClient(CodeTiers: string): Boolean;
Procedure RTMajNatureContacts(StAuxi,StNat : String);
Procedure RTMajNatureAdresses(StAuxi,StNat : String);
Procedure RTMajNatureFrais(StAuxi,StNat : String);
procedure RTMajNatureIntervenant(StAuxi, StNat: String);
function GetNatureTiers(CodeTiers : string; IsAuxi : boolean=false) : string;
function LookupListeEnseigne(Enseigne, Siret: THEdit): hString;
//FV1 : 23/08/2011
function GetLibTiers(CliFou : string; Code : string; var Libelle : string) : boolean;
function GetTiersFerme(CliFou : string; Code : string; var Libelle : string) : boolean;
function FormatNomLib(Nomtiers : String) : String;
function FindTiersFromIDPSIGAO (IDSPIGAo : Integer) : string;
procedure AddIDPSIGAOToTiers (CodeTiers : string; IDSPIGAo : Integer);
//
function ControleFournisseurOk (Fourn : string) : boolean;


Const
   villeLyon : String = 'LYON';
   villeMarseille : String = 'MARSEILLE';
   villeParis : String = 'PARIS';

implementation

{$IFNDEF GCGC}
  {$IFDEF DP}
  uses
    Annoutils;
  {$ENDIF DP}
{$ELSE GCGC}
  uses
   EntGC
  {$IFDEF DP}
  ,DpJurOutils
  ,AnnOutils
  {$ENDIF DP}
  , AGLInitGC { GC_20080729_PFA_010;16327 }
  ;
  {$ENDIF GCGC}

Const Liste_delimiters : String = ',.;:-_/ ' ;    // liste des separateurs
      TexteMessage: array [1..2] of string = (
        {1} 'Liste des clients '
        {2},'Liste des fournisseurs '
        );


function ControleFournisseurOk (Fourn : string) : boolean;
begin
  result := ExisteSQL('SELECT 1 FROM TIERS WHERE T_TIERS="'+Fourn+'" AND T_NATUREAUXI="FOU" AND T_FERME="-"');
end;

function FindTiersFromIDPSIGAO (IDSPIGAo : Integer) : string;
var QQ : TQuery;
begin
  Result := '';
  QQ := OpenSQL('SELECT T_TIERS FROM TIERS LEFT JOIN BTIERS ON BT1_AUXILIAIRE=T_AUXILIAIRE WHERE BT1_IDSPIGAO='+InttoStr(IDSPIGAO)+' AND T_NATUREAUXI IN ("CLI","PRO")',True,1,'',true);
  if Not QQ.eof then
  begin
		Result := QQ.Fields [0].AsString;
  end;
  ferme (QQ);
end;

procedure AddIDPSIGAOToTiers (CodeTiers : string; IDSPIGAo : Integer);
var QQ : TQuery;
begin
  ExecuteSQL('UPDATE BTIERS SET BT1_IDSPIGAO='' WHERE BT1_IDSPIGAO="'+InttOstr(IDSPIGAO)+'" AND (SELECT T_NATUREAUXI FROM TIERS WHERE T_AUXILIAIRE=BT1_AUXILIAIRE) IN ("CLI","PRO")');
  ExecuteSQL('UPDATE BTIERS SET BT1_IDSPIGAO='+InttOstr(IDSPIGAO)+' WHERE BT1_AUXILIAIRE=(SELECT ##TOP 1## T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+CodeTiers+'" AND T_NATUREAUXI IN ("CLI","PRO"))');
end;


{$IFDEF DP}
Function CreeUnAnnuaire(CodeTiers : string;VientDuTiers:boolean=false):string;
var
  TobAnn: TOB;
  QQ: TQuery;
	GuidpER: string;
begin
  if (GetParamsocSecur ('SO_AFANNUAIREIMPETIERS', '-') = false) and not (VientDuTiers) then exit;
  if (GetParamsocSecur ('SO_AFLIENDP', '-') = false) then  exit;
  if (TableSurAutreBase ('ANNUAIRE')) then exit;
  TobAnn := TOB.Create ('ANNUAIRE', nil, -1);
	// fct qui permet de créer la fiche annauire correspodnant au tiers
	// uniquement si paramsoc activé
  try
    QQ := OPENSQL ('SELECT ANN_GUIDPER From ANNUAIRE WHERE ANN_TIERS="' + CodeTiers + '"', True,-1,'',true);
    if QQ.Eof then
    begin
      TobAnn.PutValue ('ANN_TIERS', CodeTiers);
      GUidPer := AglGetGuid();
      TobAnn.PutValue ('ANN_GUIDPER', GuidPer);
      TobAnn.PutValue ('ANN_CODEPER', -2); // $$$ JP 21/04/06 - il faut -2 depuis l'apparition de ANN_GUIDPER
      TobAnn.PutValue ('ANN_DATECREATION', V_PGI.DateEntree);
      TobAnn.PutValue ('ANN_UTILISATEUR', V_PGI.User);
      TobAnn.InsertDB (nil);
    end
    else
      Guidper := QQ.fields [0] .AsString;
  finally
    TobAnn.Free;
  end;
 result:=GuidPer; //mcd 04/03/2006
 SynchroniseTiers (False, GuidPer, CodeTiers);
end;
{$ENDIF}

function TiersEncoursMultiSoc : boolean;
begin
  Result := EstBaseMultiSoc and GetParamSocSecur ('SO_ENCOURSMS', false) and
            TablePartagee('TIERS');
end;

procedure CalculSoldesAuxi ( Auxi : String ) ;
Const icDeb=1 ; icCre=2 ; icEnc=1 ; icSui=2 ; icNor=1 ; icAno=2 ;
Var Q : TQuery ;
    SQL,sExo,sSet : String ;
    iExo,iAno : integer ;
    TT : Array[icDeb..icCre,icEnc..icSui,icNor..icAno] of Double ; {Sens,Exo,Ano}
BEGIN
if Auxi='' then Exit ;
FillChar(TT,Sizeof(TT),#0) ;
if VH^.Suivant.Code<>'' then sExo:='AND (E_EXERCICE="'+VH^.Encours.Code+'" OR E_EXERCICE="'+VH^.Suivant.Code+'") '
                        else sExo:='AND E_EXERCICE="'+VH^.Encours.Code+'" ' ;
SQL:='SELECT E_EXERCICE, E_ECRANOUVEAU, SUM(E_DEBIT), SUM(E_CREDIT) FROM ECRITURE '
    +'WHERE E_AUXILIAIRE="'+Auxi+'" AND E_QUALIFPIECE="N" ' ;
SQL:=SQL+sExo ;
SQL:=SQL+'GROUP BY E_EXERCICE, E_ECRANOUVEAU' ;
Q:=OpenSQL(SQL,True,-1,'',true) ;
While Not Q.EOF do
   BEGIN
   if Q.Fields[0].AsString=VH^.Encours.Code then iExo:=1 else iExo:=2 ;
   if Q.Fields[1].AsString='N' then iAno:=1 else iAno:=2 ;
   TT[icDeb,iExo,iAno]:=Q.Fields[2].AsFloat ; TT[icCre,iExo,iAno]:=Q.Fields[3].AsFloat ;
   Q.Next ;
   END ;
Ferme(Q) ;
sSet:='T_TOTALDEBIT='+StrfPoint(Arrondi(TT[icDeb,icEnc,icNor]+TT[icDeb,icSui,icNor]+TT[icDeb,icEnc,icAno],V_PGI.OkDecV))+', '
     +'T_TOTALCREDIT='+StrfPoint(Arrondi(TT[icCre,icEnc,icNor]+TT[icCre,icSui,icNor]+TT[icCre,icEnc,icAno],V_PGI.OkDecV))+', '
     +'T_TOTDEBE='+StrfPoint(Arrondi(TT[icDeb,icEnc,icNor]+TT[icDeb,icEnc,icAno],V_PGI.OkDecV))+', '
     +'T_TOTCREE='+StrfPoint(Arrondi(TT[icCre,icEnc,icNor]+TT[icCre,icEnc,icAno],V_PGI.OkDecV))+', '
     +'T_TOTDEBS='+StrfPoint(Arrondi(TT[icDeb,icSui,icNor]+TT[icDeb,icSui,icAno],V_PGI.OkDecV))+', '
     +'T_TOTCRES='+StrfPoint(Arrondi(TT[icCre,icSui,icNor]+TT[icCre,icSui,icAno],V_PGI.OkDecV))+', '
     +'T_TOTDEBANO='+StrfPoint(Arrondi(TT[icDeb,icEnc,icAno],V_PGI.OkDecV))+', '
     +'T_TOTCREANO='+StrfPoint(Arrondi(TT[icCre,icEnc,icAno],V_PGI.OkDecV))+', '
     +'T_TOTDEBANON1='+StrfPoint(Arrondi(TT[icDeb,icSui,icAno],V_PGI.OkDecV))+', '
     +'T_TOTCREANON1='+StrfPoint(Arrondi(TT[icCre,icSui,icAno],V_PGI.OkDecV))+' ' ;
ExecuteSQL('UPDATE TIERS SET '+sSet+' WHERE T_AUXILIAIRE="'+Auxi+'"') ;
END ;

{$IFDEF GCGC}

// GC_20080606_DM_GC16079 DEBUT
// Ajout de l'origine du calcul (Typenattiers et QuelTiers)
//function RisqueTiersGC(TOBTiers : TOB ; sSocietes : string = ''; TobTiersParam : TOB=nil) : Double ;
function RisqueTiersGC(TOBTiers : TOB ; sSocietes : string = ''; TobTiersParam : TOB=nil;
                      Typenattiers : string = ''; QuelTiers : string = '') : Double ;
// GC_20080606_DM_GC16079 FIN
Var StNat,SQL,CodeTiers : String ;
    i     : integer ;
    TOBPP : TOB ;
    Q     : TQuery ;
    TobParPieceMS  : Tob;
    sSocGroupe, sBase : string;
    TobParamTiersImpact : TOB;
    ChampL, ChampT, LeCodeTiers : string;


  function GetBase (UneBase, LaTable : string) : String;
  begin
    if isMssql then result := UneBase + '.DBO.' + LaTable
    else result := UneBase + '.' + LaTable;
  end;

BEGIN
  Result:=0 ;
  if not assigned(TOBTiers) then exit ;
  ChampL := '';
  ChampT := '';
  // GC_20080606_DM_GC16079 DEBUT
  if Typenattiers = '' then
  begin
  // GC_20080606_DM_GC16079 FIN
  { Il faut tester si le tiers courant à un autre tiers paramétré pour l'encours }
  if not assigned(TobTiersParam) then
  begin
    TobParamTiersImpact := ImpactPiece_ChargeTiers('VEN', '').FindFirst(['GTI_ELEMENTFC'], ['ENC'], true);
    if assigned(TobParamTiersImpact) then
    begin
      ChampL := GetNomChampTiersParam(TobParamTiersImpact.GetString('GTI_TYPENATTIERS'), 'GL_');
      ChampT := GetNomChampTiersParam(TobParamTiersImpact.GetString('GTI_TYPENATTIERS'), 'T_');
      { Si Principal, on garde la valeur de la tob tiers
        sinon, il faut chercher le code (la valeur de la tob tiers est l'auxiliaire)}
      if TobParamTiersImpact.GetString('GTI_TYPENATTIERS') = 'PRI' then
        LeCodeTiers := TobTiers.GetString(ChampT)
        else
        LeCodeTiers := TiersAuxiliaire(TobTiers.GetString(ChampT), True);
    end else
    begin
      ChampL := 'GL_TIERS';
      ChampT := 'T_TIERS';
      LeCodeTiers := TobTiers.GetString(ChampT);
    end;
  end else
  if (TobTiersParam.FieldExists('NATURETIERS')) and (TobTiersParam.FieldExists('CODE')) then
  begin
    ChampL := GetNomChampTiersParam(TobTiersParam.GetString('NATURETIERS'), 'GL_');
    ChampT := TobTiersParam.GetString('CODE');
    LeCodeTiers := ChampT;
  end;
  // GC_20080606_DM_GC16079 DEBUT
  end
  else
  begin
    if Typenattiers = 'PRI' then
    begin
      ChampL := 'GL_TIERS';
      LeCodeTiers := TobTiers.GetString('T_TIERS');
    end
    else
    begin
      ChampL := 'GL_TIERSFACTURE';
      if QuelTiers = 'PRI' then
        LeCodeTiers := TiersAuxiliaire(TobTiers.GetString('T_AUXILIAIRE'), True)
      else
        LeCodeTiers := TiersAuxiliaire(TobTiers.GetString('T_FACTURE'), True);
    end;
  end;
  // GC_20080606_DM_GC16079 FIN
  if (ChampL = '') or (LeCodeTiers = '') then exit;
    CodeTiers := ' ' + ChampL + ' = "' + LeCodeTiers + '"';
  if CodeTiers='' then exit ;
  { Multi-sociétés }
  sSocGroupe := sSocietes;
  sBase := ReadTokenSt(sSocGroupe) ;
  repeat
     stNat:='' ;
    if (sBase <> '') and not TablePartagee('PARPIECE') then
    begin
      TobParPieceMS := Tob.Create('',nil,-1);
      try
        // GC_20071120_GM_GC15580 DEBUT
        // TobParPieceMS.LoadDetailFromSQL('SELECT GPP_NATUREPIECEG FROM ' + GetBase (sBase, 'PARPIECE') + ' WHERE GPP_ENCOURS="X" AND GPP_TYPEECRCPTA<>"NOR" AND GPP_VENTEACHAT="VEN"');
        TobParPieceMS.LoadDetailFromSQL('SELECT GPP_NATUREPIECEG FROM ' + GetBase (sBase, 'PARPIECE') + ' WHERE GPP_ENCOURS="X" AND ((GPP_TYPEECRCPTA<>"NOR") OR (GPP_TYPEECRCPTA="NOR" AND (GPP_TYPEPASSCPTA = "" OR GPP_TYPEPASSCPTA="AUC"))) AND GPP_VENTEACHAT="VEN"');
        // GC_20071120_GM_GC15580 FIN
        for i:=0 to TobParPieceMS.Detail.Count-1 do
          stNat := stNat + ' OR GL_NATUREPIECEG="' + TobParPieceMS.Detail[i].GetString('GPP_NATUREPIECEG')+'"' ;
      finally
        TobParPieceMS.Free;
      end;
    end else
    begin
      for i:=0 to VH_GC.TOBParPiece.Detail.Count-1 do
      begin
        TOBPP:=VH_GC.TOBParPiece.Detail[i] ;
        // GC_20071120_GM_GC15580 DEBUT
        // if ((TOBPP.GetValue('GPP_ENCOURS')='X') and (TOBPP.GetValue('GPP_TYPEECRCPTA')<>'NOR') and
        if ((TOBPP.GetValue('GPP_ENCOURS')='X') and
         ((TOBPP.GetValue('GPP_TYPEECRCPTA')<>'NOR')
        OR  ((TOBPP.GetValue('GPP_TYPEECRCPTA')='NOR') AND
        ((TOBPP.GetValue('GPP_TYPEPASSCPTA') = '') OR (TOBPP.GetValue('GPP_TYPEPASSCPTA')='AUC')))) and
        // GC_20071120_GM_GC15580 FIN
            (TOBPP.GetValue('GPP_VENTEACHAT')='VEN'))
           then StNat:=StNat+' OR GL_NATUREPIECEG="'+TOBPP.GetValue('GPP_NATUREPIECEG')+'"' ;
           //then StNat:=StNat+' OR GP_NATUREPIECEG="'+TOBPP.GetValue('GPP_NATUREPIECEG')+'"' ;
      end ;
    end;
    if StNat='' then Exit ;
    Delete(StNat,1,4) ; StNat:='('+StNat+')' ;
    // BDU - 14/05/07 - FQ : 14038, 14039. Ajout de "VALEUR" après la formule  
    SQL:='SELECT SUM(GL_TOTALTTC*ABS(GL_QTERESTE)/ABS(GL_QTEFACT)) VALEUR FROM LIGNE WHERE ' + CodeTiers
                +' AND '+StNat+' AND GL_VIVANTE="X" AND GL_ARTICLE<>"" AND GL_ETATSOLDE="ENC" AND GL_QTEFACT<>0 AND GL_TOTALTTC<>0' ;
    Q:=OpenSQL(SQL, True, -1, '', true, sBase) ;
    try
      if Not Q.EOF then result:=result + Arrondi(Q.Fields[0].AsFloat,V_PGI.OkDecV);
    finally
      Ferme(Q) ;
    end;
    sBase  := ReadTokenSt(sSocGroupe) ;
  until (sBase = '');
end;
{$ENDIF}
//CEGID-CCMX le 15/09/2006 : ajout iNbjoursEncours {Evolution}
//function EnCoursRegleNonEchu(Aux : String ; DateButoir : tDateTime ; sSocietes : string = '') : Double ;
function EnCoursRegleNonEchu(Aux : String ; DateButoir : tDateTime ; iNbJoursEncours : Integer ; sSocietes : string = '') : Double ;

Var St,St1 : String ;
    Q : TQuery ;
    sBase, sSocGroupe : string;
begin
  Result:=0 ;
  { Multi-sociétés }
  sSocGroupe := sSocietes;
  sBase := ReadTokenSt(sSocGroupe) ;
  repeat
    st:='SELECT SUM(E_DEBIT-E_CREDIT) FROM ECRITURE ' ;
    St:=St+' LEFT JOIN MODEPAIE ON E_MODEPAIE=MP_MODEPAIE ' ;
    St:=St+' LEFT JOIN GENERAUX ON E_CONTREPARTIEGEN=G_GENERAL ' ;
    St:=St+' WHERE E_AUXILIAIRE="'+Aux+'" ' ;
    St:=St+' AND (MP_CATEGORIE="LCR") AND (G_NATUREGENE="BQE" OR G_NATUREGENE="CAI") '  ;
 {Evolution}
    // BDU - 04/06/07 - FQ : 14138. DB2 n'accepte pas E_DATEECHEANCE + 20
    // Je déplace l'ajout du nombre de jours et j'inverse le signe du nombre de jours
    // St:=st+' AND E_DATEECHEANCE+ ' + IntToStr(iNbJoursEncours) + '>"'+USDATETIME(DateButoir)+'" ' ;
    St := St + 'AND E_DATEECHEANCE > "' + UsDateTime(DateButoir + (iNbJoursEnCours * -1)) + '" ';
    St:=St+' AND E_QUALIFPIECE="N" AND E_MODEPAIE<>"" AND (E_ECRANOUVEAU="N" OR (E_ECRANOUVEAU="H" AND E_EXERCICE="'+VH^.ExoV8.Code+'"))' ;
    St1:=LWhereV8 ; If St1<>'' Then St:=St+' AND '+St1 ;
    Q:=OpenSQL(St, true, -1, '', true, sBase) ;
    try
      If Not Q.Eof Then Result := result + Q.Fields[0].AsFloat ;
    finally
      Ferme(Q) ;
    end;
    sBase  := ReadTokenSt(sSocGroupe) ;
  until (sBase = '');
end ;

//CEGID-CCMX le 15/09/2006 : ajout iNbjoursEncours
//function EnCoursEchuNonRegle (Aux : String ; DateButoir : tDateTime ; sSocietes : string = ''): double;
function EnCoursEchuNonRegle (Aux : String ; DateButoir : tDateTime ; iNbJoursEncours : Integer; sSocietes : string = ''): double;
Var St,St1 : String ;
    Q : TQuery ;
    sBase, sSocGroupe : string;
begin
result:=0;
  { Multi-sociétés }
  sSocGroupe := sSocietes;
  sBase := ReadTokenSt(sSocGroupe) ;
  repeat
    st:='Select sum(E_DEBIT-E_COUVERTURE)';
    st:=st+' From ECRITURE ' ;
    st:=st+' Left Join GENERAUX on G_GENERAL=E_GENERAL';
    st:=st+' Where E_AUXILIAIRE="'+Aux+'" and E_DEBIT>0 ';
    st:=st+'   And (E_DATECOMPTABLE<="'+USDATETIME(DateButoir)+'" Or E_DATEPAQUETMIN<="'+USDATETIME(DateButoir)+'")';
    //CEGID-CCMX le 15/09/2006 : ajout iNbjoursEncours
    St:=st+'   AND E_DATEECHEANCE+ ' + IntToStr(iNbJoursEncours) + '<"'+USDATETIME(DateButoir)+'" ' ;
    st:=st+'   And E_QUALIFPIECE="N"  And E_ETATLETTRAGE<>"TL"';
    st:=st+'   And E_ETATLETTRAGE<>"RI" And E_ECRANOUVEAU<>"CLO" And E_ECRANOUVEAU<>"OAN" and E_QUALIFPIECE<>"C"';
    St1:=LWhereV8 ; If St1<>'' Then St:=St+' AND '+St1 ;
    Q:=OpenSQL(st, true, -1, '', true, sBase) ;
    try
      if Not Q.EOF then Result:=Result + Q.Fields[0].AsFloat ;
    finally
      Ferme(Q) ;
    end;
    sBase  := ReadTokenSt(sSocGroupe) ;
  until (sBase = '');
end;

function TiersDebitMoinsCredit (TobTiers : Tob; TQ : TQuery; sSocGroupe : string = '') : double;
var
  QMulti : TQuery;
  sBase : string;
  stTiers : string;
begin
  Result := 0;
  if not Assigned (TobTiers) and not Assigned (TQ) then exit;
  if EstTablePartagee ('TIERS') then
  begin
    if Assigned (TobTiers) then stTiers := TobTiers.GetString ('T_AUXILIAIRE')
    else stTiers := TQ.FindField('T_AUXILIAIRE').AsString;

    sBase := ReadTokenSt(sSocGroupe) ;
    repeat
      QMulti := OpenSql ('SELECT CU_DEBIT1-CU_CREDIT1 FROM CUMULS WHERE CU_TYPE="MST" AND CU_COMPTE1="' +
                          stTiers + '"', true, -1, '', true, sBase);
      try
      if not QMulti.Eof then
        Result := Result + QMulti.Fields[0].AsFloat;
      finally
        Ferme (QMulti);
      end;
      sBase  := ReadTokenSt(sSocGroupe) ;
    until (sBase = '');
  end else
  begin
    if Assigned (TobTiers) then
      Result := TOBTiers.GetDouble('T_TOTALDEBIT')-TOBTiers.GetDouble('T_TOTALCREDIT') //LM20070315 Result := TOBTiers.GetValue('T_TOTALDEBIT')-TOBTiers.GetValue('T_TOTALCREDIT')
    else
      Result := TQ.FindField('T_TOTALDEBIT').AsFloat - TQ.FindField('T_TOTALCREDIT').AsFloat;
  end;
end;

function RisqueTiersCPTA(TOBTiers : TOB ; DateB : TDateTime ; sSocietes : string=''; TobTiersParam : TOB=nil) : Double ;
Var Solde,RegleNonEchu : Double ;
 {Evolution}
iNbJoursEnCours : Integer;  //CEGID-CCMX le 15/09/2006
BEGIN
  Result := 0;
  //CEGID-CCMX LE 15/09/2006 Debut {Evolution}
  iNbJoursEncours := GetParamSocSecur('SO_NBJENCOURS', 0);
  //CEGID-CCMX LE 15/09/2006 Fin {Evolution}
  if not assigned(TOBTiers) then Exit ;
  { Si le tiers paramétré n'est pas celui de la pièce, il faut trouver les infos }
  if assigned(TobTiersParam) then
  begin
    Solde := TobTiersParam.GetDouble('SOLDE');
    //CEGID-CCMX le 15/09/2006 : ajout iNbJoursEncours {Evolution}
//    RegleNonEchu := EncoursRegleNonEchu(TobTiersParam.GetString('AUXI'), DateB, sSocietes) ;
    RegleNonEchu := EncoursRegleNonEchu(TobTiersParam.GetString('AUXI'), DateB, iNbJoursEnCours, sSocietes) ;

  end else
  begin
    Solde := TiersDebitMoinsCredit (TobTiers, nil, sSocietes);
    //CEGID-CCMX le 15/09/2006 : ajout iNbJoursEncours {Evolution}
//    RegleNonEchu := EncoursRegleNonEchu(TOBTiers.GetValue('T_AUXILIAIRE'),DateB,sSocietes) ;
    RegleNonEchu := EncoursRegleNonEchu(TOBTiers.GetValue('T_AUXILIAIRE'),DateB, iNbJoursEnCours,sSocietes) ;

  end;
  Result := Solde-RegleNonEchu ;
END ;

function CalculeEnCours(TobTiers : TOB; bEntreeMultiSoc : boolean; var RisqueGC, RegleNonEchu, EchuNonRegle, RisqueCPTA : double) : double;
var sSocGroupe : string;
    Solde : double;
    iNbJoursEnCours : Integer;  //CEGID-CCMX le 15/09/2006
begin
  { Multi-sociétés }
  if bEntreeMultiSoc then
    sSocGroupe := GetBasesMS
  else
    sSocGroupe := '';

  //CEGID-CCMX LE 15/09/2006 Debut {Evolution}
  iNbJoursEncours := GetParamSocSecur('SO_NBJENCOURS', 0);
  //CEGID-CCMX LE 15/09/2006 Fin {Evolution}

  Solde := TiersDebitMoinsCredit (TobTiers, nil, sSocGroupe);
{$IFDEF GCGC}
  RisqueGC := RisqueTiersGC(TobTiers, sSocGroupe);
{$ENDIF}
  RegleNonEchu := -EncoursRegleNonEchu(TobTiers.GetString('T_AUXILIAIRE'), V_PGI.DateEntree, iNbJoursEncours,sSocGroupe) ;
  EchuNonRegle := EnCoursEchuNonRegle(TobTiers.GetString('T_AUXILIAIRE'), V_PGI.DateEntree, iNbJoursEncours,sSocGroupe) ;
  RisqueCPTA := Solde + RegleNonEchu ;

  Result := RisqueGC + RisqueCPTA;
end;

Function  CreationTiers ( InfTiers : Info_Tiers; var LeRapport : string ; TypAux : String =''; LeNumero : String=''; bParle : Boolean=True) : String ;
Var TOBT : TOB ;
    Racine : String ;
    NumAuxi : integer ;
    CodeAuxi,Auxi,st : String ;
    L1,L     : integer ;
BEGIN
Result:='' ; CodeAuxi:='' ;
Auxi := '';
TOBT:=TOB.Create('TIERS',Nil,-1) ;
// Infos depuis ressource
TOBT.PutValue('T_NATUREAUXI','FOU') ;
if TypAux <> '' then TOBT.PutValue('T_NATUREAUXI',TypAux) ;

TOBT.PutValue('T_LIBELLE',InfTiers.Libelle) ; TOBT.PutValue('T_ABREGE',Copy(InfTiers.Libelle,1,17)) ;
TOBT.PutValue('T_ADRESSE1',InfTiers.Adresse1) ; TOBT.PutValue('T_ADRESSE2',InfTiers.Adresse2) ; TOBT.PutValue('T_ADRESSE3',InfTiers.Adresse3) ;
TOBT.PutValue('T_CODEPOSTAL',InfTiers.CodePostal) ; TOBT.PutValue('T_VILLE',InfTiers.Ville) ;
TOBT.PutValue('T_PAYS',InfTiers.Pays) ; TOBT.PutValue('T_TELEPHONE',InfTiers.Telephone) ;
// Code auxi

if TypAux = 'SAL' then
   begin
   Racine:=GetParamSoc('SO_PGRACINEAUXI') ;
   if LeNumero <> '' then
    begin
      if isnumeric(LeNumero) then  //deb Test pour création automatique du compte tiers lorsque le matricule est alpha-numérique
      begin
        NumAuxi:= StrToInt(LeNumero); // recup du numéro du salarie passé en paramètre
      end
      else
        Auxi := LeNumero;
    end //fin test pour création automatique du compte tiers lorsque le matricule est alpha-numérique
    else NumAuxi:=GetParamSoc('SO_PGNUMAUXI') ;
   end
 else begin Racine:=GetParamSoc('SO_AFFRACINEAUXI')  ; NumAuxi:=GetParamSoc('SO_AFFNUMAUXI') ;  end;
l1 := Length (racine);
L  := VH^.Cpta[fbAux].Lg;
If Auxi='' then 
  Auxi := IntToStr(NumAuxi);
{VG - On doit respecter la longueur des comptes auxiliaires des paramètres soc
While Length(Auxi)<L-L1 do Auxi:='0'+Auxi ;
CodeAuxi:=Racine+Auxi ;
}
While Length(Auxi)<L-L1 do
      Auxi:=VH^.Cpta[fbAux].Cb+Auxi ;
if (Length(Auxi)>L-L1) then
   CodeAuxi:=Racine+Copy(Auxi, l1+Length(Auxi)-L+1,L-l1)
else
   CodeAuxi:=Racine+Auxi ;
//FIN VG

//BourreLaDonc(CodeAuxi,fbAux) ;
  if TypAux = 'SAL' then
  begin
    TOBT.PutValue('T_MODEREGLE', '008') ; // Mettre en fait VH_PAIE.VirementDefautCompta
    if ((LeNumero <> '') and (isnumeric(LeNumero))) then SetParamSoc('SO_PGNUMAUXI',NumAuxi+1);
  end else
  begin
{$IFDEF GCGC}
    TOBT.PutValue('T_MODEREGLE',VH_GC.GCModeRegleDefaut) ;
{$ELSE}
    TOBT.PutValue('T_MODEREGLE',GetParamSoc('SO_GCMODEREGLEDEFAUT')) ;
{$ENDIF}
    SetParamSoc('SO_AFFNUMAUXI',NumAuxi+1);
  end;
TOBT.PutValue('T_AUXILIAIRE',CodeAuxi) ;
TOBT.PutValue('T_TIERS',CodeAuxi) ;
// Zones obligatoires
TOBT.PutValue('T_REGIMETVA',VH^.RegimeDefaut) ;

if TypAux = 'SAL' then TOBT.PutValue('T_COLLECTIF',VH^.DefautSal)
 else TOBT.PutValue('T_COLLECTIF',VH^.DefautFou) ;
// Zones diverses
TOBT.PutValue('T_LETTRABLE','X') ; TOBT.PutValue('T_MULTIDEVISE','X') ;
TOBT.PutValue('T_FACTUREHT','X') ; TOBT.PutValue('T_DEVISE',V_PGI.DevisePivot) ;
// Enregistrement
if TypAux = 'SAL' then St:='Création du tiers salarié'
                  else St:='Création du tiers';
if Not ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+CodeAuxi+'"') then
   BEGIN
   if TOBT.InsertDB(Nil) then Result:=CodeAuxi
                         else begin
                              LeRapport :='Création du tiers ' + InfTiers.Libelle+ ' impossible' ;
                              if bParle then PGIBox (LeRapport, st) ;
                              end;
   END else
   BEGIN
   LeRapport := 'Le tiers ' + InfTiers.Libelle+ ' existe déjà' ;
   if bParle then PGIBox (LeRapport, st) ;
   END ;
FreeAndNil (TOBT);       //VG - 28/02/2005 Perte de mémoire
END ;



{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Function GetTiersMul(G_CodeTiers: THCritMaskEdit; Const stWhere, stRange, stMul : string) : String ;
{$IFDEF BTP}
function ClientArechercher (G_CodeTiers : THCritMaskEdit;IndPro : integer) : boolean;
begin
  result := false;
  if G_CodeTiers.Text = '' then BEGIN Result := true; exit; END;
  if (IndPro = 0) and (not ExistTiers ('CLI',G_CodeTiers.text)) then BEGIN result := true; Exit; END;
  if (IndPro > 0) and (not ExistTiers ('CLI',G_CodeTiers.text)) and (not ExistTiers ('PRO',G_CodeTiers.text)) then BEGIN result := true; Exit; END;
  end;
{$ENDIF}
Var
  CodeTiers,stRangeLib : string;
  ind: integer;
{$IFDEF BTP}
    IndPro : integer; // JTR enlève conseils et avertissements
{$ENDIF}
  stTemp : string; { Modif CHR - LRJ 22/09/2008 FQ 031;11269 }
begin
  Result:='' ;
  CodeTiers := '';
  { Modif CHR - LRJ 22/09/2008 FQ 031;11269 }
  // le stTemp contient un texte quelconque en plus du stWhere
  // Ainsi, dans le cas où  Pos('T_NATUREAUXI="FOU"',stWhere) renvoie 0
  // par ex. s'il n'y a que T_NATUREAUXI="FOU" dans le stWhere
  // on peux tester le Pos avec le stTemp
  stTemp := 'FQ11269:' + StringReplace(stWhere,' ','',[rfReplaceAll,rfIgnoreCase]);
  { Fin Modif CHR - LRJ 22/09/2008 FQ 031;11269 }
  if stMul <> '' then
    CodeTiers := AGLLanceFiche ('GC', stMul,'T_TIERS='+G_CodeTiers.text , '', stWhere)
  else
  begin
    ind :=Pos('T_NATUREAUXI="FOU"',stWhere) ;

    {Recherche par le Ctrl F5 : Nom du tiers}
    if (pos('T_LIBELLE', stRange)<>0) then
      stRangeLib:=stRange
    else
    {Recherche par le F5 : par le Code tiers alors on complète le Range avec la valeur du control G_CodeTiers}
      stRangeLib:='T_TIERS='+G_CodeTiers.text+';'+stRange;

{$IFDEF BTP}
    if ind > 0 then
    begin
      if not ExistTiers ('FOU',G_CodeTiers.text) then
        CodeTiers := AGLLanceFiche('GC','GCFOURNISSEUR_MUL','T_TIERS='+G_CodeTiers.text+';T_NATUREAUXI=FOU','','SELECTION')
      else
         CodeTiers := G_CodeTiers.text;
    end
    else
    begin
       IndPro := Pos('PRO',stWhere);
       if ClientArechercher (G_CodeTiers,IndPro) then
          //FV1 - 13/02/2018 : FS#2934 - TEST VD Saisie devis-création article-Création Fournisseur-fenêtre de création Clients prosp arrive
          CodeTiers := AGLLanceFiche ('GC', 'GCTIERS_RECH','T_TIERS='+G_CodeTiers.text, '', stWhere)
          //FV1 - 08/01/2018 : FS#2831 - JPG - Erreur en création de tiers via un devis
          //CodeTiers := AGLLanceFiche ('GC', 'GCTIERS_RECH','T_TIERS='+G_CodeTiers.text+';NATUREAUXI=CLI' , '', stWhere)
          //CodeTiers := AGLLanceFiche ('GC', 'GCTIERS_MUL','T_TIERS='+G_CodeTiers.text+';NATUREAUXI=CLI' , '', stWhere)
       else
          CodeTiers := G_CodeTiers.text;
    end;
{$ELSE}
    if (ctxAffaire in V_PGI.PGIContexte) And (ind = 0) then
    begin
        //mcd 15/03/2005 si GRC,appel tjrs même fiche	if (ctxGRC in V_PGI.PGIContexte) And (V_PGI.menucourant=92) then
    	if (ctxGRC in V_PGI.PGIContexte)then
      begin
        CodeTiers := AGLLanceFiche ('RT', 'RTTIERS_RECH',stRangeLib , '', stWhere);
      end
      else
        CodeTiers := AGLLanceFiche ('AFF', 'AFTIERS_RECH',stRangeLib , '', stWhere)
    end
    {$IFDEF CHR}
    else
    if ctxChr in V_PGI.PGIContexte then
    begin
      { Modif CHR - LRJ 22/09/2008 FQ 031;11269 }
      // Si Pos('T_NATUREAUXI="FOU"',stTemp) renvoie 0 on lance HRCLIENTS_MUL
      // sinon on lance la fiche tiers
      if (Pos('T_NATUREAUXI="FOU"',stTemp) = 0) then
      begin
      { Fin Modif CHR - LRJ 22/09/2008 FQ 031;11269 }
        if stWhere = '' then
          CodeTiers := AGLLanceFiche ('H', 'HRCLIENTS_MUL','T_TIERS=' + G_CodeTiers.Text, '', 'SELECTION')
        else
          CodeTiers := AGLLanceFiche ('H', 'HRCLIENTS_MUL','T_TIERS=' + G_CodeTiers.Text, '', 'SELECTION;'+stWhere);
      { Modif CHR - LRJ 22/09/2008 FQ 031;11269 }
      end
      else
        CodeTiers := AGLLanceFiche ('GC', 'GCTIERS_RECH', stRangeLib , '', stWhere);
      { Fin Modif CHR - LRJ 22/09/2008 FQ 031;11269 }
    end
    {$ENDIF}
    else
    if (ctxGRC in V_PGI.PGIContexte)
      {$IFDEF GCGC}
      or VH_GC.GRFSeria then
    begin
      if (Pos('T_NATUREAUXI="CON"',stWhere)>0) then
        CodeTiers := AGLLanceFiche ('GC', 'GCTIERS_RECH',stRangeLib , '', stWhere)
      else
      begin
        // GRC
        CodeTiers := AGLLanceFiche ('RT', 'RTTIERS_RECH',stRangeLib , '', stWhere);
      end;
    end
    else if ctxAffaire in V_PGI.PGIContexte then  // mcd 14/01/2003 pour accès fiche founissuer si besoin
    begin
      //mcd 10/03/03 ajout test CLI : si apppel depusi création/duplic cliet, ilf aut tout, si appel depsui rech facture fourni uniquement fournisseur ..
      //mcd 01/07/2005 ajout test longueur.. 12215.. même si FOU, il peut y avoir d'autres antures.;donc il faut passer dans fiche commune
      if (ind > 0) and (Pos('CLI',stWhere)=0) and (StrLen(Pchar(stWhere))<=18) then
          { GC_20080729_PFA_010;16327 }
//        CodeTiers := AGLLanceFiche ('GC', 'GCFOURNISSEUR_MUL','T_TIERS='+G_CodeTiers.text , '', 'SELECTION')
          CodeTiers := DispatchRecherche(nil, 2, 'T_NATUREAUXI="FOU"', 'T_TIERS='+G_CodeTiers.text, '')
      else
        CodeTiers := AGLLanceFiche ('GC', 'GCTIERS_RECH',stRangeLib , '', stWhere);
    end
    else  // fin mcd 14/01/2003
       CodeTiers := AGLLanceFiche ('GC', 'GCTIERS_RECH', stRangeLib , '', stWhere);
    {$ELSE}
    then begin end ; // temporaire, pour corriger un pb de compil, a finaliser...
    {$ENDIF GCGC}
{$ENDIF BTP}
  END;
  if CodeTiers <> '' then // Codeart pour laisser la valeur initiale si pas de selection
  begin
    G_CodeTiers.Text := CodeTiers;
  end;
  Result := CodeTiers;
END ;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure GetTiersLookUp (G_CodeTiers: THCritMaskEdit; stWhere : string);
BEGIN
if ctxaffaire in V_PGI.PGIContexte then //mcd 06/02/2006 11519 GIGA
begin
  if pos ('FOU' ,stWhere) > 0 then  {2,'Liste des fournisseurs '} //AB-200701-FQ13078
    LookupList(G_CodeTiers, TraduireMemoire(TexteMessage[2]), 'TIERS', 'T_TIERS', 'T_LIBELLE', stWhere, 'T_TIERS', False, 8)
  else                      {1 'Liste des clients ',}
    LookupList(G_CodeTiers, TraduireMemoire(TexteMessage[1]), 'TIERS', 'T_TIERS', 'T_LIBELLE', stWhere, 'T_TIERS', False, 8)
end else
  LookupCombo( G_CodeTiers ) ;
END ;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 28/03/2000
Modifié le ... :   /  /
Description .. : Recherche de tiers par Mul ou Lookup
Mots clefs ... : RECHERCHE; TIERS
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Function GetTiersRecherche (G_CodeTiers: THCritMaskEdit; Const stWhere, stRange, stMul : string ) : string ;
var
  G_Tiers : THCritMaskEdit;
BEGIN
  Result:='' ;
  G_Tiers:=Nil;
  if G_CodeTiers = nil then
  begin
    G_Tiers := THCritMaskEdit.create (nil);
    G_Tiers.Text := '';
  end;
  {$IFDEF GCGC}
  if (GetParamSocSecur('SO_GCRECHTIERSAV', False)) or (G_CodeTiers = nil) {$IFDEF GRC} or (GetParamsoc('SO_RTCONFIDENTIALITE')) {$ENDIF} then
  begin
    if G_CodeTiers = nil then
    begin
      if GetTiersMul (G_Tiers, stWhere, stRange, stMul) <> '' then
        Result := G_Tiers.Text;
    end
    else
    begin
      if GetTiersMul (G_CodeTiers, stWhere, stRange, stMul) <> '' then
        Result := G_CodeTiers.Text;
    end;
  end
  else
  {$ENDIF GCGC}
  begin
    GetTiersLookUp (G_CodeTiers, stWhere) ;
    Result:=G_CodeTiers.Text;
  end;

  if G_CodeTiers = nil then
    G_Tiers.Free;
END ;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Paul CHAPUIS
Créé le ...... : 21/04/2000
Modifié le ... :   /  /
Description .. : retourne le code auxiliaire du tiers à partir du code Tiers
Mots clefs ... : AUXILIAIRE;TIERS
*****************************************************************}

Function TiersAuxiliaire (Code: String ; AuxiliaireVersTiers : boolean = FALSE; NatureAuxi : string=''): String;
Var Q : TQuery;
    NomChamp,ChampRetourne,PlusAuxi : String;
Begin

  Result:='';

  If AuxiliaireVersTiers Then
  Begin
    NomChamp := 'T_AUXILIAIRE';
    ChampRetourne := 'T_TIERS';
  End
  Else
  Begin
    NomChamp := 'T_TIERS';
    ChampRetourne :='T_AUXILIAIRE'
  End;

  if NatureAuxi<>'' then PlusAuxi := ' AND T_NATUREAUXI="'+NatureAuxi+'"' else PlusAuxi := '';

  Q:=OPENSQL('SELECT '+ ChampRetourne +' From TIERS WHERE '+ NomChamp +'="'+Code+'"'+PlusAuxi,True,-1,'',true);

  If Not Q.EOF Then result:=Q.Fields[0].AsString;

  Ferme(Q);

end;

Function RechZoneTiers (ChampRetourne, Nomchamp, Code, NatureAuxi : String) : String;
Var StSQL     : String;
    PlusAuxi  : String;
    QQ        : TQuery;
begin

  Result := '';

  if NatureAuxi <>'' then
    PlusAuxi := ' AND T_NATUREAUXI="'+ NatureAuxi + '"'
  else
    PlusAuxi := '';

  StSQL := 'SELECT '+ ChampRetourne +' From TIERS WHERE '+ NomChamp + '="' + Code + '"' + PlusAuxi;

  QQ := OpenSQL(StSql,True,-1,'',true);

  if Not QQ.Eof then Result :=QQ.Findfield(ChampRetourne).AsString;

  ferme(QQ);

end;

//************************************************************************
//*  HABILLEMENT : fonction de formatage des zones de la fiche client    *
//************************************************************************

///////////////////////////////////////////////////////////////////////////////////////
//  VerifEstSeparateur : Vérifie si un caractère est un séparateur
///////////////////////////////////////////////////////////////////////////////////////
Function VerifEstSeparateur(const Texte : string; const Pos : integer) : Boolean ;
BEGIN
Result := IsDelimiter(Liste_delimiters, Texte, Pos) ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifSupDoubleSep : Supprime les séparateurs redondants dans un texte.
///////////////////////////////////////////////////////////////////////////////////////
Function VerifSupDoubleSep(const Texte : string) : string ;
Var i, j : Integer ;
    SupEspaceSuivant : Boolean ;
BEGIN
Result := Texte;
SupEspaceSuivant := False;
for i := 1 to Length(Result) do
    BEGIN
    if (SupEspaceSuivant) and (VerifEstSeparateur(Result, i)) then
        BEGIN
        j := i ;
        while (j <= Length(Result)) and (VerifEstSeparateur(Result, j)) do inc(j) ;
        Delete (Result, i, (j - i)) ;
        END;
    SupEspaceSuivant := (VerifEstSeparateur(Result, i)) ;
    END;
END;

///////////////////////////////////////////////////////////////////////////////////////
//  verifSupDoubleEspace : Supprime les espaces redondants dans un texte.
///////////////////////////////////////////////////////////////////////////////////////
Function verifSupDoubleEspace(const Texte : string) : string ;
Var i, j : Integer ;
    SupEspaceSuivant : Boolean ;
BEGIN
Result := Texte;
SupEspaceSuivant := False;
for i := 1 to Length(Result) do
    BEGIN
    if (SupEspaceSuivant) and (Result[i] = ' ') then
        BEGIN
        j := i ;
        while (j <= Length(Result)) and (Result[j] = ' ') do inc(j) ;
        Delete (Result, i, (j - i)) ;
        END;
    SupEspaceSuivant := (Result[i] = ' ') ;
    END;
END;


///////////////////////////////////////////////////////////////////////////////////////
//  MetTexteEnForme : Met en forme un texte.
///////////////////////////////////////////////////////////////////////////////////////
Function MetTexteEnForme(const Texte, Format : string; SupEsp : boolean) : string;
BEGIN
Result := Texte ;
if SupEsp then result := VerifSupDoubleSep(Result) ;
if Format = 'MAJ' then Result := AnsiUpperCase(Result) else // JTR - eQualité 10389
if Format = 'MIN' then Result := AnsiLowerCase (Result) else // JTR - eQualité 10389
if Format = 'PRE' then Result := FirstMajuscule (LowerCase (Result));
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  MetNomEnForme : Met en forme un texte.
///////////////////////////////////////////////////////////////////////////////////////
Function MetNomEnForme(const Texte, Format : string; Formatpart : string; const TailleMaxPart : integer; SupEsp : boolean) : string;
Var i, lg : Integer ;
    buf, trav, buffer, buffersv : string;
BEGIN
  Result := Texte ;
  buffer := Texte ;
  lg := Length(buffer);
  trav := '';
  if SupEsp then
    buffer := VerifSupDoubleSep(buffer) ;
  buffersv := buffer ;
  i := 1 ;
  while i <= lg do
  begin
    if (VerifEstSeparateur(buffer, i)) then  buffer[i] := ';';
    inc (i) ;
  end;
  while buffer <> '' do
  BEGIN
    buf:=Trim(ReadTokenSt(buffer));
    if (buf <> ' ') then
    BEGIN
      if ((Length(buf)) > TailleMaxPart) then
      begin
        if Format = 'MAJ' then buf := AnsiUpperCase (buf) else
        if Format = 'MIN' then buf := AnsiLowerCase (buf) else
        if Format = 'PRE' then buf := FirstMajuscule (LowerCase (buf)) else
      end else
      begin
        if Formatpart = 'MAJ' then buf := AnsiUpperCase (buf) else
        if Formatpart = 'MIN' then buf := AnsiLowerCase (buf) else
        if Formatpart = 'PRE' then buf := FirstMajuscule (LowerCase (buf)) else
      end;
      trav := Concat(trav, buf);
      trav := Concat(trav, ' ');
    end;
  END ;
  i := 1;
  while i <= lg do
  begin
    if (VerifEstSeparateur(buffersv, i)) then  trav[i] := buffersv[i];
    inc (i) ;
  end;
  Result := trav ;
END ;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifAppliqueFormat : Applique à un champ le format défini dans les paramètres société.
///////////////////////////////////////////////////////////////////////////////////////
Function VerifAppliqueFormat( NomChamp, ValeurChamp : string; clipart : boolean ) : variant ;
Var  chp, fmt, fmtpart : string ;
     sup : boolean ;
     max : integer ;
BEGIN
Result := ValeurChamp ;

  if clipart then
  begin
    if NomChamp = 'T_LIBELLE'   then chp := 'SO_FMTTIERSLIBELLE'  else
    if NomChamp = 'T_PRENOM'    then chp := 'SO_FMTTIERSPRENOM'   else
    if NomChamp = 'T_ADRESSE1'  then chp := 'SO_FMTTIERSADR1'     else
    if NomChamp = 'T_ADRESSE2'  then chp := 'SO_FMTTIERSADR2'     else
    if NomChamp = 'T_ADRESSE3'  then chp := 'SO_FMTTIERSADR3'     else
    if NomChamp = 'T_VILLE'     then chp := 'SO_FMTTIERSVILLE'    else
    exit;
  end
  else
  begin
    if NomChamp = 'T_LIBELLE'   then chp := 'SO_FMTSOCLIBELLE'    else
    if NomChamp = 'T_PRENOM'    then chp := 'SO_FMTSOCPRENOM'     else
    if NomChamp = 'T_ADRESSE1'  then chp := 'SO_FMTTIERSADR1'     else
    if NomChamp = 'T_ADRESSE2'  then chp := 'SO_FMTTIERSADR2'     else
    if NomChamp = 'T_ADRESSE3'  then chp := 'SO_FMTTIERSADR3'     else
    if NomChamp = 'T_VILLE'     then chp := 'SO_FMTTIERSVILLE'    else
    exit;
  end;


  fmt := varAsType(GetParamSoc(chp), varString);
  sup := varAsType(GetParamSoc('SO_FMTTIERSSUPESPACE'), varBoolean);

  if (clipart = true) then
  begin
    if (NomChamp = 'T_LIBELLE') then
  	begin
    	fmtpart := varAsType(GetParamSoc('SO_FMTTIERSPARTIC'), varString);
    	max     := varAsType(GetParamSoc('SO_FMTTIERSTAILLEMOT'), varInteger);
    	Result  := MetNomEnForme (Result, fmt, fmtpart, max, sup) ;
    end
    else
        //   if (NomChamp = 'T_PRENOM') then
  	begin
	    fmtpart := ' ';
	    max := 0;
	    Result := MetNomEnForme (Result, fmt, fmtpart, max, sup) ;
    end;
  end
  else
  begin
    //FV1 : 29/12/20014 - FS#1183 - Bati Alpes 74 En création adresse client le 1er caractère de chaque mot ne passe pas en majuscule
	  //Result := MetTexteEnForme (Result, fmt, sup) ;
    fmtpart := ' ';
    max := 0;
    Result := MetNomEnForme(Result, fmt, fmtpart, max, sup) ;
  end;

END ;

{  JCF : Plus utilisé en GESCOM
Function FormateCodeCliNum( St : string ) : string ;
var Lg,ll,i : Integer ;
    Bourre  : Char ;
    buf :string;
BEGIN
   Lg:=VH^.Cpta[fbAux].Lg ;
   Bourre:=VH^.Cpta[fbAux].Cb ;
   Result:=St ;
   ll:=Length(Result) ;
   If ll<Lg then
   begin
      for i:=1 to (Lg - ll) do buf:=buf+Bourre ;
	   Result:= Concat(buf, St) ;
   end;
END;
}


{$IFDEF GCGC}
{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 13/02/2001
Modifié le ... : 13/02/2001
Description .. : Retourne l'état d'encours du client ( Rouge, Orange, Vert)
Mots clefs ... : ENCOURS;RISQUE
*****************************************************************}
// Modif P Chapuis appel avec Code tiers ou Tob
Function GetEtatRisqueClient (CodeTiers : string): string;
var TobTiers: TOB;
    Q : TQuery;
begin
Result:='V' ;
TOBTiers:=Tob.create('TIERS',Nil,-1);
Q:=OPENSQL('SELECT * From TIERS WHERE T_TIERS="'+CodeTiers+'"',True,-1,'',true);
TOBTiers.selectDB ('',Q);
Ferme(Q);
Result:=GetEtatRisqueClient(TobTiers);
TobTiers.free;
end;

Function GetEtatRisqueClient (TobTiers : Tob): string;
var RisqueTiers,P1,P2 : double;
begin
Result:='V' ;
// Si le risque est forcé sur le tiers celui-ci est prioritaire
if TOBTiers.GetValue('T_ETATRISQUE') <> '' then
   Result:=TOBTiers.GetValue('T_ETATRISQUE')
else
   begin
   // Recalcul du risque
   RisqueTiers:= RisqueTiersGC(TOBTiers) + RisqueTiersCPTA(TOBTiers,V_PGI.DateEntree) ;
   P1:=TOBTiers.GetValue('T_CREDITACCORDE') ; P2:=TOBTiers.GetValue('T_CREDITPLAFOND') ;
   if ((P1>0 ) and (RisqueTiers > P1)) then Result:='O' ;
   if ((P2>0 ) and (RisqueTiers > P2)) then Result:='R' ;
   end;
end;
/////////////////////////////////////////
{$ENDIF}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure AfficheRisqueClient (CodeTiers : string );
var TobTiers: TOB;
    Q : TQuery;
begin
TOBTiers:=Tob.create('TIERS',Nil,-1);
Q:=OPENSQL('SELECT * From TIERS WHERE T_TIERS="'+CodeTiers+'"',True,-1,'',true);
TOBTiers.selectDB ('',Q);
Ferme(Q);
AfficheRisqueClient(TobTiers);
TobTiers.free;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure AfficheRisqueClient (TobTiers : tob );
begin
TheTob:=TobTiers;
AglLanceFiche('GC', 'GCENCOURS','','',ActionToSTring(taConsult));
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure AfficheRisqueClientDetail (CodeTiers : string );
var TobTiers: TOB;
    Q : TQuery;
begin
TOBTiers:=Tob.create('TIERS',Nil,-1);
Q:=OPENSQL('SELECT * From TIERS WHERE T_TIERS="'+CodeTiers+'"',True,-1,'',true);
TOBTiers.selectDB ('',Q);
Ferme(Q);
AfficheRisqueClientDetail (TobTiers);
TobTiers.free;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure AfficheRisqueClientDetail (TobTiers : tob );
begin
TheTob:=TobTiers;
AGLLanceFiche ('GC','GCENCOURSGC','','',ActionToString(taConsult)) ;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

/////////////////////////////////


procedure MajPhonetiqueTiers ;
Var Q : TQuery ;
BEGIN
  if PGIAsk('Voulez-vous mettre à jour le champ phonétique ?', 'MAJ champs phonétique') <> mrYes then exit;
  Q:=OpenSQL('SELECT T_LIBELLE,T_PHONETIQUE FROM TIERS WHERE T_PHONETIQUE=""',FALSE) ;
  while not Q.EOF do
    BEGIN
    Q.Edit ;
    Q.Fields[1].AsString:=Phoneme(Q.Fields[0].AsString);
    Q.Post ;
    Q.Next ;
    END ;
  Ferme(Q) ;
  PGIInfo('Traitement de mise à jour terminé', 'MAJ champs phonétique');
END ;

//******************************************************************************
//************************* gestion des adresses du tiers **********************
//******************************************************************************


// req de répercution des modifs de l'adresse de l'affaire sur l'adresse du tiers
function ModifierAdrMereParFille( AcsAdrFille:string ):boolean;
var
TobMere, TobFille : TOB;
Q : TQuery ;
begin
result:=false;
TobMere := Tob.Create('ADRESSES',nil,-1);
TobFille := Tob.Create('ADRESSES',nil,-1);

Q := OpenSQL( 'SELECT * FROM ADRESSES WHERE ADR_NUMEROADRESSE="' + AcsAdrFille + '"', True,-1,'',true );
if Not TobFille.SelectDB( '', Q ) then exit;
ferme(Q);

Q := OpenSQL( 'SELECT * FROM ADRESSES WHERE ADR_NUMEROADRESSE="' + inttostr(TobFille.GetValue('ADR_NUMADRESSEREF')) + '"', True,-1,'',true );
if Not TobMere.SelectDB( '', Q ) then exit;
ferme(Q);


TobMere.PutValue('ADR_JURIDIQUE', TobFille.GetValue('ADR_JURIDIQUE'));
TobMere.PutValue('ADR_LIBELLE', TobFille.GetValue('ADR_LIBELLE'));
TobMere.PutValue('ADR_LIBELLE2', TobFille.GetValue('ADR_LIBELLE2'));
TobMere.PutValue('ADR_ADRESSE1', TobFille.GetValue('ADR_ADRESSE1'));
TobMere.PutValue('ADR_ADRESSE2', TobFille.GetValue('ADR_ADRESSE2'));
TobMere.PutValue('ADR_ADRESSE3', TobFille.GetValue('ADR_ADRESSE3'));
TobMere.PutValue('ADR_CODEPOSTAL', TobFille.GetValue('ADR_CODEPOSTAL'));
TobMere.PutValue('ADR_VILLE', TobFille.GetValue('ADR_VILLE'));
TobMere.PutValue('ADR_PAYS', TobFille.GetValue('ADR_PAYS'));
TobMere.PutValue('ADR_TELEPHONE', TobFille.GetValue('ADR_TELEPHONE'));
TobMere.PutValue('ADR_NUMEROCONTACT', TobFille.GetValue('ADR_NUMEROCONTACT'));
TobMere.PutValue('ADR_CONTACT', TobFille.GetValue('ADR_CONTACT'));
TobMere.PutValue('ADR_BLOCNOTE', TobFille.GetValue('ADR_BLOCNOTE'));

TobMere.InsertOrUpdateDB(false);
result:=true;

TobMere.Free;
TobFille.Free;
end;

// Req de répercution des modifs de l'adresse du tiers sur toutes les adresses liées
function ModifierAdrFillesParMere( AcsAdrMere:string ):boolean;
var
  TobMere, TobFilles : TOB;
  Q : TQuery ;
  i:integer;
begin
  result:=false;
  TobMere := Tob.Create('ADRESSES',nil,-1);
  TobFilles := Tob.Create('Filles Adresses',nil,-1);
  try
    Q := OpenSQL( 'SELECT * FROM ADRESSES WHERE ADR_NUMADRESSEREF="' + AcsAdrMere + '"', True,-1,'',true );
    try
      if Not Q.EOF then
        TobFilles.LoadDetailDB('ADRESSES', '', '', Q, False)
      else
        Exit;
    finally
      Ferme(Q);
    end;

    Q := OpenSQL( 'SELECT * FROM ADRESSES WHERE ADR_NUMEROADRESSE="' + AcsAdrMere + '"', True,-1,'',true );
    try
      if Not TobMere.SelectDB( '', Q ) then
        exit;
    finally
      ferme(Q);
    end;

i:=0;
while (i<TobFilles.Detail.Count) do
    begin
    TobFilles.Detail[i].PutValue('ADR_JURIDIQUE', TobMere.GetValue('ADR_JURIDIQUE'));
	TobFilles.Detail[i].PutValue('ADR_LIBELLE', TobMere.GetValue('ADR_LIBELLE'));
	TobFilles.Detail[i].PutValue('ADR_LIBELLE2', TobMere.GetValue('ADR_LIBELLE2'));
	TobFilles.Detail[i].PutValue('ADR_ADRESSE1', TobMere.GetValue('ADR_ADRESSE1'));
	TobFilles.Detail[i].PutValue('ADR_ADRESSE2', TobMere.GetValue('ADR_ADRESSE2'));
	TobFilles.Detail[i].PutValue('ADR_ADRESSE3', TobMere.GetValue('ADR_ADRESSE3'));
	TobFilles.Detail[i].PutValue('ADR_CODEPOSTAL', TobMere.GetValue('ADR_CODEPOSTAL'));
	TobFilles.Detail[i].PutValue('ADR_VILLE', TobMere.GetValue('ADR_VILLE'));
	TobFilles.Detail[i].PutValue('ADR_PAYS', TobMere.GetValue('ADR_PAYS'));
	TobFilles.Detail[i].PutValue('ADR_TELEPHONE', TobMere.GetValue('ADR_TELEPHONE'));
	TobFilles.Detail[i].PutValue('ADR_CONTACT', TobMere.GetValue('ADR_CONTACT'));
	TobFilles.Detail[i].PutValue('ADR_BLOCNOTE', TobMere.GetValue('ADR_BLOCNOTE'));

    Inc(i);
    end;

TobFilles.InsertOrUpdateDB(false);
result:=true;
finally
TobMere.Free;
TobFilles.Free;
end;
end;

Function GetChampsComplementTiers (Codetiers,NomChamp : String) : String;
Var Q : TQuery ;
BEGIN
Result := '';
Try
 Q:= OPENSQL('SELECT '+ NomChamp + ' From TIERSCOMPL WHERE YTC_TIERS="' + CodeTiers +'"',True,-1,'',true);
 if Not Q.EOF then Result := Q.Findfield(NomChamp).AsString;
 Finally
 Ferme(Q);
 END;
END;

Function  TiersAvecAffaireActive ( CodeTiers : String):Boolean;
var stSQL : string;
begin
Result:=False;
If CodeTiers = '' then Exit;
stSQL := 'SELECT AFF_AFFAIRE FROM AFFAIRE WHERE '
         + ' AFF_TIERS="' + CodeTiers + '"'
         + ' AND AFF_STATUTAFFAIRE="AFF"'    // GA_20080408_GME_15018 on ne teste que sur "Affaire" en cours de facturation
         + ' AND AFF_DATEDEBGENER <> "'+ UsDateTime(iDate1900) +'"'
         + ' AND AFF_GENERAUTO <> "MAN" '
         + ' AND AFF_DATEFINGENER > "' + UsDateTime(V_PGI.DateEntree) + '"';

Result := ExisteSQL(stSQL);
end;


// =====  pour Script AGL =========================================

////////////////// Appel fiche Tiers ////////////////
Function AGLTiersAuxiliaire( parms: array of variant; nb: integer ): variant;
begin
    Result:=TiersAuxiliaire(string(Parms[0]),boolean(Parms[1])) ;
End;
///////////////////////////////////////////////////////////////////////////////////////
//  VerifFormateChamp : Applique à un champ le format défini dans les paramètres société.
///////////////////////////////////////////////////////////////////////////////////////
Function VerifFormateChamp( parms: array of variant; nb: integer ) : variant ;
begin
if (GetParamsoc('SO_GCFORMATEZONECLI') = TRUE )then
	result := VerifAppliqueFormat (Parms[0], Parms[1], Parms[2])
else
	result := Parms[1];
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Marc Morretton
Créé le ...... : 11/02/2003
Description .. : Teste existence d'un TIERS
*****************************************************************}
function ExistTiers(sNatureAuxi, sCodeTiers: string; sWhereCompl: string = ''; WithAlert: Boolean = false): Boolean;
var
	sSql  : string;
begin
	sSql := ' SELECT T_TIERS'
			+ ' FROM TIERS'
      + ' LEFT JOIN TIERSCOMPL ON (YTC_AUXILIAIRE=T_AUXILIAIRE)'
			+ ' WHERE ' + WhereTiers(sNatureAuxi,sCodeTiers)
      + iif(sWhereCompl <> '', ' AND ' + sWhereCompl, '');
         ;
	Result := ExisteSQL(sSql);

  if WithAlert and (not Result) then
  begin
    PgiError(TraduireMemoire('Le tiers ') + sCodeTiers + TraduireMemoire(' de nature ') + sNatureAuxi + TraduireMemoire(' n''existe pas.'), 'Tiers');
  end;
end;

function ExistT(Const Tiers: String): Boolean;
var
	Sql: string;
begin
	Sql := 'SELECT 1'
		  + ' FROM TIERS'
		  + ' WHERE ' + WhereT(Tiers)
      ;
	Result := ExisteSQL(Sql);
end;

function ExisteCodeTiers (NatureAuxi,Tiers : string) : boolean;
var	sSql  : string;
begin
	sSql := 'SELECT T_TIERS '
			+ 'FROM TIERS '
			+ 'WHERE '
      + 'T_TIERS="'+Tiers+'" AND T_NATUREAUXI="'+NatureAuxi+'"';
         ;
	Result := ExisteSQL(sSql);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MN Garnier
Créé le ...... : 17/08/2005
Description .. : Teste existence d'un TIERS sur différentes natures
*****************************************************************}
function ExistTiersMulti(sNatureAuxi, sCodeTiers: string; sWhereCompl: string = ''; WithAlert: Boolean = false): Boolean;
var sSql,stNat,Critere,wNat : string;
begin
  stNat:= sNatureAuxi;
  sSql := ' SELECT T_TIERS FROM TIERS'
  + ' LEFT JOIN TIERSCOMPL ON (YTC_AUXILIAIRE=T_AUXILIAIRE) WHERE (';
  wNat:='';
  repeat
    Critere := ReadTokenSt(stNat);
    if Critere <> '' then
      begin
      if wNat <> '' then wNat:=wNat+' OR ';
      wNat := wNat + WhereNature(Critere);
      end;
  until Critere = '';

  sSql := sSql + wNat + ') AND ' + WhereT(sCodeTiers) + iif(sWhereCompl <> '', ' AND ' + sWhereCompl, '');

  Result := ExisteSQL(sSql);

  if WithAlert and (not Result) then
  begin
    PgiError(TraduireMemoire('Le tiers ') + sCodeTiers + TraduireMemoire(' de nature ') + sNatureAuxi + TraduireMemoire(' n''existe pas.'), 'Tiers');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Marc Morretton
Créé le ...... : 11/02/2003
Description .. : Construction du WHERE pour la fiche TIERS
*****************************************************************}
function WhereTiers(sNatureAuxi, sCodeTiers : string): string;
begin
   if sNatureAuxi <> '' then
	   Result := 'T_NATUREAUXI = "' + sNatureAuxi + '" and '+'T_TIERS = "' + sCodeTiers + '"'
   else
	   Result := 'T_TIERS = "' + sCodeTiers + '"';
end;

function WhereT(Const Tiers : string): string;
begin
  Result := 'T_TIERS="' + Tiers + '"';
end;

function WhereNature(sNatureAuxi : string): string;
begin
   if sNatureAuxi <> '' then
     Result := 'T_NATUREAUXI = "' + sNatureAuxi + '"' ;
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : Marc Morretton
Créé le ...... : 11/02/2003
Description .. : Test l' existence d'un représentant - Commercial
*****************************************************************}
function ExistCommercial(sCodeCommercial : string; WithAlert: Boolean = false):Boolean;
var
	sSql: string;
begin
  sSql := ' SELECT GCL_COMMERCIAL'
        + ' FROM COMMERCIAL'
        + ' WHERE GCL_COMMERCIAL="'+sCodeCommercial+'"'
     ;
  Result := ExisteSQL(sSql);

  if WithAlert and (not Result) then
    PgiError(TraduireMemoire('Le représentant - commercial ') + sCodeCommercial + TraduireMemoire(' n''existe pas.'), 'Tiers');

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : KOZA Denis
Créé le ...... : 11/02/2003
Description .. : Construction du WHERE pour la fiche TIERSCOMPL
*****************************************************************}
function WhereTiersCompl(sNatureAuxi, sCodeTiers : string): string;
begin
  if sNatureAuxi <> '' then
	  Result := 'YTC_NATUREAUXI = "' + sNatureAuxi + '" and '+'YTC_TIERS = "' + sCodeTiers + '"'
  else
	  Result := 'YTC_TIERS = "' + sCodeTiers + '"';
end;

function WhereYTC(Const Tiers : string): string;
begin
  Result := 'YTC_TIERS = "' + Tiers + '"'
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 11/02/2003
Description .. : Appel de la fiche Tiers
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure CallTiers(Lequel:string; Action:TActionFiche = taModif; Params: String = ''; Range : String = '');
Var
  NatureAuxi: string;
begin
  if Pos('T_NATUREAUXI', Params) = 0 then
  begin
    NatureAuxi:=wGetSQLFieldValue('T_NATUREAUXI', 'TIERS', 'T_AUXILIAIRE="' + Lequel + '"');
    Params := Params + iif(Params <> '', ';', '') + 'T_NATUREAUXI=' + NatureAuxi;
  end
  else
    NatureAuxi:=GetArgumentString(Params, 'T_NATUREAUXI');

  if NatureAuxi='FOU' then
    AGLLanceFiche('GC','GCFOURNISSEUR', '', lequel, ActionToString(Action) + ';T_NATUREAUXI=FOU')
  else
    AGLLanceFiche('GC', 'GCTIERS', Range, Lequel, ActionToString(Action) + ';MONOFICHE;' + Params);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : KOZA Denis
Créé le ...... : 11/09/2002
Modifié le ... : 11/02/2003
Description .. : Construction du WHERE pour la fiche TIERSCOMPL
*****************************************************************}
function GetFieldFromTiers(Field, sNatureAux, sTiers: string): String;
begin
   if wGetPrefixe(Field) = 'T' then
	   Result := wGetSqlFieldValue(Field, 'TIERS', WhereTiers(sNatureAux, sTiers))
   else if wGetPrefixe(Field) = 'YTC' then
      Result := wGetSqlFieldValue(Field, 'TIERSCOMPL', WhereTiersCompl(sNatureAux, sTiers))
   else
      Result := '';
end;

function GetFieldFromT(Const FieldName, Tiers: string): Variant;
begin
  if wGetPrefixe(FieldName) = 'T' then
    Result := wGetSqlFieldValue(FieldName, 'TIERS', WhereT(Tiers))
  else if wGetPrefixe(FieldName) = 'YTC' then
    Result := wGetSqlFieldValue(FieldName, 'TIERSCOMPL', WhereYTC(Tiers))
  else
    Result := null;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : KOZA Denis
Créé le ...... : 11/09/2002
Description .. : Renvoie une liste de champ depuis la fiche Tiers
*****************************************************************}
function GetFieldsFromTiers(Fields: Array of string; sNatureAux, sTiers: string): MyArrayValue;
begin
	Result := wGetSqlFieldsValues(Fields, 'TIERS', WhereTiers(sNatureAux, sTiers));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Marc Morretton
Créé le ...... : 11/02/2003
Description .. : Charge de la Tob avec un ou plusieurs tiers suivant le
Suite ........ : where
*****************************************************************}
function GetTobTiers(sChamp, sWhere: String; TobTiers: Tob; SelectDB:Boolean = False; sOrderBy: string=''): Boolean;
var
  sRequete: String;
  Q: tQuery;
begin
  Result := False;
  if (TobTiers <> nil) then
  begin
    sRequete :=  'SELECT '+sChamp+' FROM TIERS';
    if (pos('YTC_', sChamp)>0) then
      sRequete := sRequete + ' LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS';
    sRequete := sRequete + ' WHERE '+sWhere;

    if (sOrderBy<>'') then
     sRequete := sRequete + 'ORDER BY '+sOrderBy;

    if ExisteSQL(sRequete) then
    begin
      Q := OpenSQL(sRequete, True,-1,'',true);
      try
        if not SelectDB then
        begin
          TobTiers.LoadDetailDB('TIERS', '', '', Q, True, True);
          Result := (TobTiers.Detail.Count > 0);
        end
        else
        begin
          TobTiers.SelectDB('', Q);
          Result := True;
        end;
      finally
        Ferme(Q);
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 31/01/2003
Description .. : Retourne le numéro de l'adresse par défaut d'un tiers
*****************************************************************}
function GetNumAdresseFromTiers(NatureAuxi, Tiers: String; TheType: TypeAdresse): String;
var
  FieldName: String;
begin
  case TheType of
    taLivr: FieldName := 'YTC_NADRESSELIV';
    taFact: FieldName := 'YTC_NADRESSEFAC';
  else FieldName := '';
  end;
  if FieldName <> '' then
    Result := GetFieldFromTiers(FieldName, NatureAuxi, Tiers)
  else
    Result := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 03/02/2003
Description .. : Recherche d'un champs dans la table des codes postaux
*****************************************************************}
function GetFieldFromCodePostal(Field, sCodePostal, sVille: string): Variant;
var
  sWhere : string;
begin
  if (sCodePostal<>'')                   then sWhere := '(O_CODEPOSTAL="'+sCodePostal+'")';
  if (sWhere     <>'') and (sVille <>'') then sWhere := sWhere + ' and ';
  if (sVille     <>'')                   then sWhere := '(O_VILLE="'+sVille+'")';
  Result := wGetSqlFieldValue(Field, 'CODEPOST', sWhere);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 22/01/2003
Description .. : Retourne le dernier N° d'adresse pour un tiers
*****************************************************************}
function GetTiersLastAdresseNum(CodeTiers, NatureAuxi: String): Integer;
var
   Q: TQuery;
begin
   Result := 0;
   Q := OpenSQL('SELECT MAX(ADR_NADRESSE)'
                 + ' FROM ADRESSES'
                 + ' WHERE ADR_NATUREAUXI="' + NatureAuxi + '"'
                 + '   AND ADR_TYPEADRESSE="TIE" '
                 + '   AND ADR_REFCODE="' + CodeTiers + '"'
                   ,True,-1,'',true) ;
   try
      if not Q.EOF then
         Result := Q.Fields[0].AsInteger;
   finally
      Ferme(Q);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 24/01/2003
Description .. : Ouverture de la fiche des Adresses des Tiers
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure CallAdressesTiers(Range, Lequel, Params: string);
begin
   AglLanceFiche('GC', 'GCADRESSES', Range, Lequel, Params);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry PETETIN
Créé le ...... : 31/01/2003
Description .. : Ouverture de la fiche des adresses en eAGL
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
Procedure AglFicheAdressesTiers(parms: array of variant; nb: integer);
begin
   {
   Params[0] = La fiche
   Params[1] = Range
   Params[2] = Lequel
   Params[3] = Paramètres
   }
   CallAdressesTiers(Parms[1], Parms[2], Parms[3]);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

function GetNomsFromContact(Fields: Array of string; Auxiliaire: string; NumeroContact: integer): MyArrayValue;
begin
  Result := wGetSqlFieldsValues(Fields, 'CONTACT', 'C_TYPECONTACT="T" AND C_AUXILIAIRE = "' + Auxiliaire + '" AND C_NUMEROCONTACT = ' + intToStr(NumeroContact));
end;

// DEBUT CCMX-CEGID FQ N° GC14693
function GetNomsFromTypeContact(Fields: Array of string; Auxiliaire: string; NumeroContact: integer; TypeContact :String): MyArrayValue;
begin
  Result := wGetSqlFieldsValues(Fields, 'CONTACT', 'C_TYPECONTACT="' + TypeContact + '" AND C_AUXILIAIRE = "' + Auxiliaire + '" AND C_NUMEROCONTACT = ' + intToStr(NumeroContact));
end;
// FIN CCMX-CEGID FQ N° GC14693
{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Renvoie un champ de CONTACT pour un client
Suite ........ : et en fonction du WHERE particulier
Mots clefs ... :
*****************************************************************}
function GetFieldFromContact(Field, sAuxiliaire, where: string): Variant;
begin
  Result := wGetSqlFieldValue(Field, 'CONTACT', 'C_TYPECONTACT = "T" AND C_AUXILIAIRE = "' + sAuxiliaire + iif(where<>'', ' " AND ' + where, '"'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 09/09/2004
Modifié le ... :   /  /
Description .. : Ramène un tiers par code barre (EAN)
Mots clefs ... :
*****************************************************************}
function GetTiersFromCb(const NatureAuxi, CodeBarre: String): String;
var
  Q_Tiers: TQuery;
begin
  Result := '';
  if CodeBarre <> '' then
  begin
    Q_Tiers := OpenSQL('SELECT T_TIERS'
                     + ' FROM TIERS'
                     + ' WHERE T_NATUREAUXI="' + NatureAuxi + '"'
                     + ' AND T_EAN="' + CodeBarre + '"', True, 1,'',true);
    try
      if not Q_Tiers.Eof then
        Result := Q_Tiers.FindField('T_TIERS').AsString
    finally
      Ferme(Q_Tiers)
    end
  end
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Denis KOZA
Créé le ...... : 11/06/2003
Modifié le ... : 11/06/2003
Description .. : > Recherche du code tiers pour un dépôt
Suite ........ :
Suite ........ : Renvoie le tiers associé à la fiche dépôt
Suite ........ : si celui-ci est renseigné
Suite ........ : sinon le tiers par défaut des paramères société GC
Mots clefs ... :
*****************************************************************}
function RechercheTiersDepot(Depot: string): string;
var
  TiersDuDepot, TiersDefaut  : string;
begin
  TiersDefaut := GetParamSoc('SO_GCTIERSDEFAUT');
  TiersDuDepot := NullToVide(wGetSqlFieldValue('GDE_TIERS' ,'DEPOTS', 'GDE_DEPOT="' +Depot+ '"'));

  if TiersDuDepot <> '' then
    Result := TiersDuDepot
  else
    Result := TiersDefaut;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 28/08/2003
Description .. : Création d'un enregistrement dans TiersCompl si besoin
*****************************************************************}
procedure InitTiersCompl(Auxiliaire, CodeTiers: String);
var
  TobTiersCompl: Tob;
begin
  if (Auxiliaire <> '') and (CodeTiers <> '') and (not ExisteSQL('SELECT YTC_TIERS FROM TIERSCOMPL WHERE ' + WhereTiersCompl('', CodeTiers))) then
  begin
    TobTiersCompl := Tob.Create('TIERSCOMPL', nil, -1);
    try
       TobTiersCompl.P('YTC_AUXILIAIRE', Auxiliaire);
       TobTiersCompl.P('YTC_TIERS', CodeTiers);
       TobTiersCompl.InsertDB(nil);
    finally
      TobTiersCompl.Free;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 03/06/2004
Modifié le ... :   /  /    
Description .. : Renvoie la référence affectatio pour l'auxiliaure demandé
Mots clefs ... : 
*****************************************************************}
function GetRefAffectationT(Const Tiers: string): string;
begin
  if Tiers <> '' then
    Result := 'T'
            + '~' + Tiers
  else
    Result := ''
end;

function GetTiersFromRefAffectation(RefAffectation: string): string;
begin
  if wleft(RefAffectation, 2) = 'T~' then
  begin
    ReadTokenPipe(RefAffectation, '~');
    Result := ReadTokenPipe(RefAffectation, '~')
  end
  else
    Result := '';
end;

//==============================================================================
//  fonction de creation d'un tiers à partir d'une tob.
//  elle applique les regles de validation definies dans la tom tiers
//  et retourne un message d'erreur ainsi qu'un champ Error eventuel.
//  les champs de donnees à renseigner sont fournis en tant que champs sup
//  dans TheTob.
//==============================================================================
function CreateTiersFromTob(Tiers: String; TheTob: Tob): String;
begin
  Result := '';
  { Contrôle existence du tiers }
  {GC_NFO_20080721_FQ;PGISIDE;10187_DEB}
  if ExisteSQL('Select T_TIERS from TIERS where T_TIERS="' + Tiers + '" OR T_CODEIMPORT="'
             + Tiers +'"') then
  {GC_NFO_20080721_FQ;PGISIDE;10187_FIN}
  begin
    Result := CreateOrUpdateTiersFromTob(Tiers, TheTob);
    TheTob.AddChampSupValeur('Error', Result);
    Exit;
  end;

  Result := CreateOrUpdateTiers(True, Tiers, TheTob);
  if Result <> '' then
    TheTob.AddChampSupValeur('Error', Result);
end;

//==============================================================================
//  fonction de creation et/ou modif d'un tiers à partir d'une tob.
//  elle applique les regles de validation definies dans la tom tiers
//  et retourne un message d'erreur ainsi qu'un champ Error eventuel.
//  les champs de donnees à renseigner sont fournis en tant que champs sup
//  dans TheTob.
//==============================================================================
function CreateOrUpdateTiersFromTob(Tiers: String; TheTob: Tob): String;
begin
  Result := '';
  Result := CreateOrUpdateTiers(False, Tiers, TheTob);
  if Result <> '' then
    TheTob.AddChampSupValeur('Error', Result);
end;

function CreateOrUpdateTiers(NewTiers : boolean; Tiers: String; TheTob: Tob): String;
var Q : TQuery;
    iChamp: Integer;
    NatAuxi, FieldName: String;
    // TiersExist
    TiersFerme, Exist, IsValid: Boolean;
    TobTiers, TobZonelibre: Tob;
    TomTiers: Tom;
    {$ifdef DP}
      Guid :string;
      Q1   :TQuery;
    {$endif}
    stAuxiliaire : string;
begin
  Result := '';
  { Contrôle existence du tiers }

  {GC_NFO_20080721_FQ;PGISIDE;10187_DEB}
  TiersFerme := ExisteSQL('Select T_TIERS from TIERS where (T_TIERS="' + Tiers + '" OR T_CODEIMPORT="'
                        + Tiers +'") AND T_FERME="X"');
  {GC_NFO_20080721_FQ;PGISIDE;10187_FIN}
  if TheTob <> nil then
  begin
    if not TiersFerme then
    begin
      { Crée une TobTiers et contrôle l'existence des champs de la tob passée en paramêtre }
      TobTiers := Tob.Create('TIERS', nil, -1);
      TobZonelibre := TOB.Create('TIERSCOMPL', Nil, -1);
      stAuxiliaire := '';
      if not NewTiers then
      begin
        {GC_NFO_20080721_FQ;PGISIDE;10187_DEB}
        Q := OpenSQL('Select * from TIERS where T_TIERS="' + Tiers + '" OR T_CODEIMPORT="'
                   + Tiers +'"', True,-1,'',true);
        {GC_NFO_20080721_FQ;PGISIDE;10187_FIN}           
        if not Q.EOF then
        begin
          TobTiers.SelectDB('"' + Q.FindField('T_AUXILIAIRE').AsString + '"', nil);
          stAuxiliaire := Q.FindField('T_AUXILIAIRE').AsString;
          TobZonelibre.SelectDB('"' + Q.FindField('T_AUXILIAIRE').AsString + '"', nil);
        end
        else
        begin
          NewTiers := True;
          Ferme(Q);
        end;
      end;
      try
        { Recopie la tob à mettre à jour dans la tob tiers }
        iChamp := 999; Exist := True;
        while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) and Exist do
        begin
          Inc(iChamp);
          { Vérifie si le champ fait partie de la table tiers }
          if Copy(TheTob.GetNomChamp(iChamp), 0, Pos('_', TheTob.GetNomChamp(iChamp)) - 1) = 'T' then
            Exist := TobTiers.FieldExists(TheTob.GetNomChamp(iChamp))
          else
            TobZoneLibre.PutValue(TheTob.GetNomChamp(iChamp), TheTob.GetValue(TheTob.GetNomChamp(iChamp)));
          if TheTob.GetNomChamp(iChamp) = 'T_NATUREAUXI' then
            NatAuxi := TheTob.GetValue(TheTob.GetNomChamp(iChamp));
        end;
        if Exist then
        begin
          { Vérifie les données en passant par une TomTiers }
          if not NewTiers then
          begin
            TomTiers := TOM.Create(Q, nil, False, 'TIERS');
            Ferme (Q);
          end else
            TomTiers := CreateTOM('TIERS', nil, false, true);
          TomTiers.Argument('ORIGINE=PGISIDE;T_NATUREAUXI=' + NatAuxi + ';T_TIERS=' + Tiers);
          {GC_NFO_20080721_FQ;PGISIDE;10187_DEB}
          if not NewTiers then
          begin
            TomTiers.LoadRecord;
            TheTob.DelChampSup('T_TIERS', true);
          end
          else
            TomTiers.InitTOB(Tobtiers);
          {GC_NFO_20080721_FQ;PGISIDE;10187_FIN}  
          { Recopie la tob à mettre à jour dans la tob tiers }
          iChamp := 999;
          while (iChamp < (1000 + (TheTob.ChampsSup.Count - 1))) do
          begin
            Inc(iChamp);
            FieldName := TheTob.GetNomChamp(iChamp);
            if Copy(FieldName, 0, Pos('_', FieldName) - 1) = 'T' then
              TobTiers.PutValue(FieldName, TheTob.GetValue(FieldName));
          end;
          try
            TobTiers.AddChampSup ('MODIFICATIONPGISIDE', false);
            TobTiers.AddChampSupValeur('IKC', 'M', false);
            IsValid := TomTiers.VerifTOB(TobTiers);
            Result := TomTiers.LastErrorMsg;
          finally
            TomTiers.Free;
          end;
          if IsValid then
          begin
            if NewTiers then
            begin
              try
                TobTiers.InsertDB(nil, False); { Enregistre la TobTiers }
                {$ifdef DP}
                  CreeUnAnnuaire (TobTiers.getvalue('T_TIERS'));
                {$endif}
                TheTob.Dupliquer(TobTiers, False, True);
                try
                  TobZoneLibre.PutValue('YTC_AUXILIAIRE', TobTiers.GetValue('T_AUXILIAIRE'));
                  TobZoneLibre.PutValue('YTC_TIERS', TobTiers.GetValue('T_TIERS'));
                  TobZoneLibre.InsertDB(nil, False);
                except
                  on E: Exception do
                  begin
                    {$IFDEF PGISIDE}
                    MAJJnalEvent('PSE', 'ERR', TraduireMemoire('Mise à jour de la table TIERSCOMPL'), E.Message);
                    {$ENDIF PGISIDE}
                    Result := E.Message;
                  end;
                end;
              except
                on E: Exception do
                begin
                  {$IFDEF PGISIDE}
                  MAJJnalEvent('PSE', 'ERR', TraduireMemoire('Mise à jour de la table TIERS'), E.Message);
                  {$ENDIF PGISIDE}
                  Result := E.Message;
                end;
              end;
            end
            else
            begin
              try
                TobTiers.UpdateDB; { Enregistre la TobTiers }
                {$ifdef DP}
                  Guid:='';
                  Q1 := OpenSQL('select ANN_GUIDPER from ANNUAIRE where ANN_TIERS = '
                    + '"'+TobTiers.getvalue('T_TIERS')+'"', TRUE,-1,'',true);
                  if not Q1.eof then Guid := Q1.FindField('ANN_GuidPER').AsString;
                  Ferme (Q1);
                  SynchroniseTiers (False,Guid,TobTiers.getvalue('T_TIERS'));
                {$endif}
                TheTob.Dupliquer(TobTiers, False, True);
                try
                  if stAuxiliaire <> '' then
                    TobZoneLibre.SetString ('YTC_AUXILIAIRE', stAuxiliaire)
                  else TobZoneLibre.SetString ('YTC_AUXILIAIRE', TobTiers.GetString ('T_AUXILIAIRE'));
                  TobZoneLibre.PutValue('YTC_TIERS', TobTiers.GetValue('T_TIERS'));
                  TobZoneLibre.InsertOrUpdateDB;
                except
                  on E: Exception do
                  begin
                    {$IFDEF PGISIDE}
                    MAJJnalEvent('PSE', 'ERR', TraduireMemoire('Mise à jour de la table TIERSCOMPL'), E.Message);
                    {$ENDIF PGISIDE}
                    Result := E.Message;
                  end;
                end;
              except
                on E: Exception do
                begin
                  {$IFDEF PGISIDE}
                  MAJJnalEvent('PSE', 'ERR', TraduireMemoire('Mise à jour de la table TIERS'), E.Message);
                  {$ENDIF PGISIDE}
                  Result := E.Message;
                end;
              end;
            end;
          end;
  //        else
  //          Result := TraduireMemoire('Les données du tiers : ') + Tiers + TraduireMemoire(' ne sont pas valides');
        end
        else
          Result := TraduireMemoire('Le champ : ') + FieldName + TraduireMemoire(' n''existe pas dans la table : ') + 'TIERS';
      finally
        TobTiers.Free;
        TobZoneLibre.Free;
      end;
    end
    else
    begin
      Result := Format(TraduireMemoire('Le tiers %s est fermé et ne peut être modifié.'), [Tiers]);
    end;
  end
  else
    Result := TraduireMemoire('Paramètres d''appel incorrect (Tob non renseignée)');
  if Result <> '' then
    TheTob.AddChampSupValeur('Error', Result);
end;


function RechercheTiersFromAuxi(LeAuxi, LeAuxiOrig, LeTiers : string) : string;
var Qry : TQuery;
begin
  // L'auxiliaire est celui du tiers courant
  if LeAuxi = LeAuxiOrig then
    Result := LeTiers
  else if LeAuxi <> '' then
  begin
    Qry := OpenSql('SELECT T_TIERS FROM TIERS WHERE T_AUXILIAIRE="' + LeAuxi + '"', true,-1,'',true);
    try
      if not Qry.Eof then
        Result := Qry.Findfield('T_TIERS').AsString
        else
        Result := '';
    finally
      Ferme(Qry);
    end;
  end else
    Result := '';
end;

function AutoriseCreationTiers (NatureAuxi : string) : boolean;
begin
 Result :=
  ((NatureAuxi = 'CLI') and ExJaiLeDroitConcept(TConcept(gcCLICreat),False) and ExJaiLeDroitConcept(TConcept(gcCLIModif),False))
  or
  ((NatureAuxi = 'FOU') and ExJaiLeDroitConcept(TConcept(gcFOUCreat),False) and ExJaiLeDroitConcept(TConcept(gcFOUModif),False))
    or
  ((NatureAuxi = 'PRO') and ExJaiLeDroitConcept(TConcept(gcProspectCreat),False) and ExJaiLeDroitConcept(TConcept(gcProspectModif),False));
end;

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. : Renvoie le nom du champ par rapport au paramétrage impact tiers
Mots clefs ... :
*****************************************************************}
function GetNomChampTiersParam(NatTiers, PrefixeWith_ : string) : string;
var FromTiers : boolean;
begin
  Result := '';
  if (NatTiers = '') or (PrefixeWith_ = '') then exit;
  FromTiers := (PrefixeWith_ = 'T_');
  if NatTiers = 'PRI' then
    Result := PrefixeWith_ + 'TIERS'
  else if NatTiers = 'LIV' then
  begin
    if FromTiers then
      Result := 'YTC_TIERSLIVRE'
      else
      Result := PrefixeWith_ + 'TIERSLIVRE';
  end else
  if NatTiers = 'FAC' then
  begin
    if FromTiers then
      Result := 'T_FACTURE'
      else
      Result := PrefixeWith_ + 'TIERSFACTURE';
  end else
  if NatTiers = 'PAY' then
  begin
    if FromTiers then
      Result := 'T_PAYEUR'
      else
      Result := PrefixeWith_ + 'TIERSPAYEUR';
  end else
  if NatTiers = 'GPR' then
  begin
    if FromTiers then
      Result := 'T_SOCIETEGROUPE'
      else
      Result := 'GPRIMPACTPCE';
  end else
  if NatTiers = 'GPA' then
    Result := 'GPAIMPACTPCE'

  else if NatTiers = 'GLI' then
    Result := 'GLIIMPACTPCE'

  else if NatTiers = 'GFA' then
    Result := 'GFAIMPACTPCE'
  else
    Result := '';
end;


{$IFNDEF EAGLSERVER}
Procedure LanceGoogleMaps( Adresse, Ville : string; BPlan : boolean);
var sql,sHttp,adr1,vil1 : String;
    Q : TQuery;
begin
  if Uppercase(copy(Ville, 1, 4)) = villeLyon then Ville:=villeLyon
    else if Uppercase(copy(Ville, 1, 5)) = villeParis then Ville:=villeParis
    else if Uppercase(copy(Ville, 1, 9)) = villeMarseille then Ville:=villeMarseille
  ;
  if BPlan then
      sHttp := 'http://maps.google.fr/maps?f=q&hl=fr&q='+
    StringReplace(trim(Adresse), ' ', '+',[rfReplaceAll])+',+'+Ville
  else
    begin
    adr1:='';
    vil1:='';
    if VH^.ProfilUserC[prEtablissement].Etablissement <> '' then
      begin
      sql:='select ET_ADRESSE1,ET_VILLE from etabliss where ET_ETABLISSEMENT = "'+VH^.ProfilUserC[prEtablissement].Etablissement+'"';
      Q:=OpenSql(sql,true,-1,'',true);
      if (not Q.eof) then
        begin
        adr1:=trim(Q.FindField('ET_ADRESSE1').AsString);
        vil1:=Q.FindField('ET_VILLE').AsString;
        end;
      Ferme(Q);
      end;
    if (vil1 = '' ) or (adr1 = '' ) then
        begin
        adr1:=trim(GetParamSocSecur('SO_ADRESSE1',''));
        vil1:=GetParamSocSecur('SO_VILLE','');
        end;

    sHttp := 'http://maps.google.fr/maps?f=q&hl=fr&saddr='+
    StringReplace(adr1, ' ', '+',[rfReplaceAll])
      +',+'+vil1+
      '&daddr='+StringReplace(trim(Adresse), ' ', '+',[rfReplaceAll])+',+'+Ville;
    end;

  LanceWeb(sHttp,False);
end;
{$ENDIF EAGLSERVER}
// MB : 11/05/2007
// Décalage de la gestion des doublon tiers pour réutilisation sur la fiche annuaire.
//
procedure VerifDoublonSiretPays (TheTom : TOM ; Siret : String ; Pays : String ; NatureAuxi : String ; Auxiliaire : String ; GuidPer : String ; ControlDoublon : String ; var LastError : Integer; var DoublonValue : String ) ;
var
  stNat,SO_GCCONTROLESIREN, stField, stLib,
  SO_GCPAYSSIRET, stSQL, stPrefixe, stSiren, stCleSiret : String ;
  {$IFNDEF EAGLSERVER}
  ret : string;
  {$ENDIF !EAGLSERVER}
  Qr : TQuery ;
  procedure SetLastError(ierr : integer ; serr : string) ;
  begin
     if serr<>'' then TheTom.SetFocusControl(serr);
     LastError:=ierr;
     //LastErrorMsg:=MSG_Tiers[LastError];
  end;
begin
   // LastError > 0 --> break le OnUpdateRecord.
   // LastError = 1 --> break pour ressaisie.
   SetLastError(0,'');
   DoublonValue := '' ;
   SO_GCPAYSSIRET := GetParamSoc('SO_GCPAYSSIRET') ;
   Siret := Trim(Siret);
   if ( Siret <> '' ) and
      ( Pos(Pays,SO_GCPAYSSIRET) <> 0 ) then
   begin
     if not VerifSiret( Siret) then
     begin
        SetLastError(29, 'T_SIRET'); exit ;
     end;
     if Auxiliaire <> '' then // GI
     begin
        stSiren := '' ;
        stSQL := 'SELECT T_TIERS FROM TIERS WHERE T_SIRET LIKE "'+Siret+
         '%" and T_AUXILIAIRE<>"'+Auxiliaire+'"' ;
        stPrefixe := 'T' ;
        stField := 'T_TIERS' ;
        stLib := 'tiers' ;
     end
     else // BUREAU
     begin
        stSiren := Copy(Siret,1,9);
        stCleSiret := Copy(Siret,10,5);
        stSQL := 'Select ANN_NOMPER from ANNUAIRE where ANN_SIREN="'+stSiren+'" AND ANN_CLESIRET="'+stCleSiret+'" AND ANN_GUIDPER <> "'+GuidPer+'"' ;
        stPrefixe := 'ANN' ;
        stField := 'ANN_NOMPER' ;
        stLib := 'annuaire' ;
     end;

     { mng : ajout condition sur la nature : exemple SIC même tiers en client et fournisseur }
     if ( NatureAuxi='CLI' ) or ( NatureAuxi='PRO' ) then
        stNat:=' and ( ('+stPrefixe+'_NATUREAUXI="PRO") or ('+stPrefixe+'_NATUREAUXI="CLI"))';
     if ( NatureAuxi='CON' ) then
        stNat:=' and ('+stPrefixe+'_NATUREAUXI="CON")';
     if ( NatureAuxi='FOU' ) then
        stNat:=' and ('+stPrefixe+'_NATUREAUXI="FOU")';

     Qr := OpenSql(stSQL+stNat, TRUE,-1,'',true);
     try
       if not Qr.Eof then
       begin
        // mng : si Siret, refus doublon, si Siren, suivant ParamSoc
        // MB sauf si bureau ...
        SO_GCCONTROLESIREN := GetParamsoc('SO_GCCONTROLESIREN') ;

        DoublonValue := Qr.FindField(stField).AsString ;
        if  ( Auxiliaire <> '' )  // MB : Cas non bureau pour iso-foncitonnalité TIERS.
        and ( length(Siret)<>9 )
        and ( SO_GCCONTROLESIREN <> 'DOU' ) then
        begin
          LastError:=30;
          exit;
        end;
        if SO_GCCONTROLESIREN = 'BLO' then
        begin
          LastError:=34;
          exit;
        end;
        if ( Auxiliaire = '' ) // MB : Cas bureau.
        and ( SO_GCCONTROLESIREN <> 'DOU' ) then
        begin
            if (PGIAsk('Ce code Siret existe déjà pour la fiche ' + stLib +' : ' +
                DoublonValue + chr(13) +
                ' Confirmez-vous ce code ?')<>mrYes) then
               LastError:=1
             else
               LastError:=0;
          exit;
        end;

        {$IFNDEF EAGLSERVER}
        if SO_GCCONTROLESIREN = 'DOU' then
        begin
          if ( TheTom.GetControl(ControlDoublon) <> nil )
          and TheTom.GetControl(ControlDoublon).Visible
          and not TheTom.GetControl(ControlDoublon).Enabled
          and ( TheTom.GetControlText(ControlDoublon) = '' ) then
          begin // il faut que je puisse saisir la justification du doublon.
             if (PGIAsk('Ce code Siret existe déjà pour la fiche ' + stLib +' : '+
                DoublonValue + chr(13) +
                ' Confirmez-vous ce code ?')<>mrYes) then
             begin
               LastError:=1;
               //LastErrorMsg:='';
               exit;
             end
             else
             begin
               TheTom.GetControl(ControlDoublon).Enabled := True ;
               if TheTom.GetControl(ControlDoublon) is ThValcombobox then
                  if (not ThValCombobox(TheTom.GetControl(ControlDoublon)).DataTypeParametrable) and
                     (JaiLeDroitTag(74348)) then              // FQ 14289  verif accès tablette   TJA 21/06/2007
                     ThValCombobox(TheTom.GetControl(ControlDoublon)).DataTypeParametrable := True ;
               // GetControl('YTC_DOUBLON').SetFocus ;
               PGIBox('Veuillez renseigner la justification du doublon sur ce Siret ! '+#10+#13+'Saisie obligatoire !');
               TheTom.SetFocusControl(ControlDoublon);
               LastError:=1;
               //LastErrorMsg:='';
               exit ;
             end;
          end;
          if ( TheTom.GetControl(ControlDoublon) <> nil )
          and TheTom.GetControl(ControlDoublon).Visible
          and TheTom.GetControl(ControlDoublon).Enabled
          and ( TheTom.GetControlText(ControlDoublon) = '' ) then
          begin
             if (PGIAsk('Ce code Siret existe déjà pour la fiche ' + stLib +' : ' +
                DoublonValue + chr(13) +
                ' Confirmez-vous ce code ?')<>mrYes) then
             begin
               LastError:=1;
               exit;
             end
             else
             begin
               if TheTom.GetControlText(ControlDoublon) = '' then
               begin
                  PGIBox('Veuillez renseigner la justification du doublon sur ce Siret ! '+#10+#13+'Saisie obligatoire !');
                  TheTom.SetFocusControl(ControlDoublon);
                  LastError:=1;
                  exit ;
               end;
             end;
          end;
        end
        else if SO_GCCONTROLESIREN = 'AVE' then
        begin
          if (PGIAsk('Ce code Siret existe déjà pour la fiche ' + stLib +' : '+
             DoublonValue + chr(13) +
             ' Confirmez-vous ce code ?')<>mrYes) then
          begin
            LastError:=1;
            exit;
          end;
        end
        else if SO_GCCONTROLESIREN = 'ALD' then // MB : 
        begin
          if (PGIAsk('Ce code Siret est déjà utilisé' + chr(13) +
             ' Voulez-vous consulter la liste des doublons ?')=mrYes) then
          begin
            ret:=AGLLanceFiche('RT','RTDOUBLONSSIREN','T_SIRET='+Siret,'','');
            if ret<>'ok' then
            begin
              LastError:=1;
              exit;
            end;
          end;
        end
        else
        begin
          LastError:=34;
          exit;
        end;
        {$ENDIF EAGLSERVER}
      end;
     finally
       Ferme(Qr);
     end;
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Garnier Marie-Noëlle
Créé le ...... : 12/09/2001
Modifié le ... : 17/09/2001
Description .. : Procédure de transformation d'un Tiers de Nature Prospect
Suite ........ : à la Nature Client
Mots clefs ... : NATURE TIERS
*****************************************************************}
procedure MajProspectClient ( TOBGenere : Tob ; TOBTiers : Tob);
begin
(*
if GetInfoParPiece(TOBGenere.GetValue('GP_NATUREPIECEG'),'GPP_PROCLI') = 'X' then
    begin
    if (TOBTiers.GetValue('T_NATUREAUXI')='PRO') then
       begin
       // transformer le tiers de PRO en CLI
       TOBTiers.putValue('T_NATUREAUXI','CLI');
       TOBTiers.putValue('T_DATEPROCLI',V_PGI.DateEntree);
       RTMajNatureContacts(TOBTiers.GetValue('T_AUXILIAIRE'),'CLI');
       RTMajNatureAdresses(TOBTiers.GetValue('T_TIERS'),'CLI');
       RTMajNatureFrais(TOBTiers.GetValue('T_TIERS'),'CLI');
       RTMajNatureIntervenant(TOBTiers.GetValue('T_TIERS'), 'CLI'); //GA_200809_AB-GA
       end;

    TransformeProspectClient(TOBGenere,'GP_TIERSLIVRE');
    TransformeProspectClient(TOBGenere,'GP_TIERSFACTURE');
    TransformeProspectClient(TOBGenere,'GP_TIERSPAYEUR');
    end;
*)
end;

procedure TransformeProspectClient ( TOBGenere : Tob ; NomChamp : String);
var Requete : String;
begin
// transformer le tiers de PRO en CLI
if ( TOBGenere.GetValue(NomChamp) <> '' ) and
   ( TOBGenere.GetValue('GP_TIERS')<>TOBGenere.GetValue(NomChamp) ) then
   begin
   Requete:='UPDATE TIERS SET T_NATUREAUXI="CLI",T_DATEPROCLI="'+USDATETIME (V_PGI.DateEntree)+
     '" WHERE T_TIERS="'+TOBGenere.GetValue(NomChamp)+'" and T_NATUREAUXI="PRO"';
   ExecuteSQL (Requete);
   // mng 27-12-2002 : maj nature contacts
   RTMajNatureContacts(TiersAuxiliaire(TOBGenere.GetValue(NomChamp),false),'CLI');
   RTMajNatureAdresses(TOBGenere.GetValue(NomChamp),'CLI');
   RTMajNatureFrais(TOBGenere.GetValue(NomChamp),'CLI');   
   if (GetParamsocSecur('SO_AFLIENDP',False) =true)  then  //mcd 06/10/2005 il faut changer dans annuaire
       ExecuteSQL ('Update ANNUAIRE set ann_natureauxi="CLI" where ann_tiers="'+TOBGenere.GetValue(NomChamp)+'"');
   // BDU - 13/13/07 - FQ : 13797
   //RTMajNatureIntervenant(TOBGenere.GetValue(NomChamp), 'CLI'); //GA_200809_AB-GA
   end;
end;

function AFMajProspectClient(CodeTiers: string): Boolean;
var StAuxi: string;
begin //voir fct MajProspectClient de UtilRt
  Result := True;
  if existeSql('SELECT T_TIERS FROM TIERS where T_NATUREAUXI="CLI" AND T_TIERS="' + CodeTiers + '"') then exit; //mcd 02/03/05 inutile si déjà client
  ExecuteSQL('UPDATE TIERS SET T_NATUREAUXI="CLI", T_DATEPROCLI ="' + UsDateTime(V_PGI.DateEntree) +
    '" WHERE T_NATUREAUXI="PRO" AND T_TIERS="' + CodeTiers + '"');
  StAuxi := TiersAuxiliaire(CodeTiers, False);
  ExecuteSQL('UPDATE CONTACT SET C_NATUREAUXI="CLI" WHERE C_TYPECONTACT="T" AND ' +
    'C_AUXILIAIRE="' + StAuxi + '"');
  //mcd 02/03/2005 manque adresse et tiersfrais
    //mcd 30/08/07 mise code tiers au lieu auxi dans refcode.. sinon pas OK + ajout TIE pour plus rapide
  ExecuteSQL('Update ADRESSES set adr_natureauxi="CLI" where ADR_TYPEADRESSE="TIE" and ADR_REFCODE="' + Codetiers + '"');
  ExecuteSQL('Update TIERSFRAIS set gtf_natureauxi="CLI" where gtf_tiers="' + StAuxi + '"');
  if (GetParamsocSecur('SO_AFLIENDP',False) = true) then //mcd 06/10/2005 il faut changer dans annuaire
    ExecuteSQL('Update ANNUAIRE set ann_natureauxi="CLI" where ann_tiers="' + CodeTiers + '"');
  // BDU - 13/13/07 - FQ : 13797
  ExecuteSQL('UPDATE AFFTIERS SET AFT_TYPEINTERV = "CLI" WHERE AFT_TYPEINTERV = "PRO" ' +
    'AND AFT_TIERS = "' + CodeTiers + '"');
end;

Procedure RTMajNatureContacts(StAuxi,StNat : String);
begin
ExecuteSQL ('Update contact set c_natureauxi="'+StNat+'" where c_typecontact="T" and c_auxiliaire="'+StAuxi+
'" and c_natureauxi<> "'+StNat+'"');
end;
Procedure RTMajNatureAdresses(StAuxi,StNat : String);
begin
//  BBI Optimisation : ajout de 'ADR_TYPEADRESSE="TIE" and' en debut de where
ExecuteSQL ('Update ADRESSES set adr_natureauxi="'+StNat+'" where ADR_TYPEADRESSE="TIE" and ADR_REFCODE="'+StAuxi+
'" and adr_natureauxi<> "'+StNat+'"');
end;
Procedure RTMajNatureFrais(StAuxi,StNat : String);
begin
ExecuteSQL ('Update TIERSFRAIS set gtf_natureauxi="'+StNat+'" where gtf_tiers="'+StAuxi+
'" and gtf_natureauxi<> "'+StNat+'"');
end;

// BDU - 13/13/07 - FQ : 13797
procedure RTMajNatureIntervenant(StAuxi, StNat: String);
begin
  ExecuteSQL('UPDATE AFFTIERS SET AFT_TYPEINTERV = "' + StNat + '" WHERE AFT_TIERS = "' + StAuxi +
    '" AND AFT_TYPEINTERV  = "PRO"');  //GA_200809_AB-GA
end;

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 25/06/2007
Modifié le ... :
Description .. : Renvoie la nature d'un tiers
Mots clefs ... : NATURE TIERS
*****************************************************************}
function GetNatureTiers(CodeTiers : string; IsAuxi : boolean=false) : string;
var Qry : TQuery;
    LeChamp : string;
begin
  if CodeTiers <> '' then
  begin
    if not IsAuxi then
      LeChamp := 'T_TIERS'
      else
      LeChamp := 'T_AUXILIAIRE';
    Qry := OpenSQL('SELECT T_NATUREAUXI FROM TIERS WHERE ' + LeChamp + ' = "' + CodeTiers + '"', True,-1,'',true);
    if not Qry.Eof then
      Result := Qry.Fields[0].AsString
      else
      Result := '';
    Ferme(Qry);
  end else
    Result := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 23/08/2011
Modifié le ... :   /  /    
Description .. : récupération du nom d'un tiers en fonction de son code et 
Suite ........ : de sa nature...
Mots clefs ... : 
*****************************************************************}
function LookupListeEnseigne(Enseigne, Siret: THEdit): hString;
{$IFNDEF EAGLSERVER}
var
  stEnseigne: THEdit;
  stWhere: string;
  {$IFNDEF ERADIO}
  NumDisp: integer;
  {$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
begin
  Result := '';
{$IFNDEF EAGLSERVER}
  stWhere := '';
  stEnseigne := THEdit.Create(Nil);
  try
    if GetParamSoc('SO_GCENSEIGNETAB') then
    begin
      stEnseigne.Parent := Enseigne.Parent;
      stEnseigne.visible := False;
      if Assigned(Siret) then
        stEnseigne.Top := Siret.top;
      stEnseigne.Left := Enseigne.left + Enseigne.Width;
      {$IFNDEF ERADIO}
      if GetParamSoc('SO_GCENSEIGNECREAT') then NumDisp := 43 else NumDisp := 0;
      {$ENDIF !ERADIO}
      {$IFNDEF EAGLCLIENT}
        stEnseigne.Text := Enseigne.Text;
      {$ENDIF}
      {$IFNDEF ERADIO}
      if (LookUpList(stEnseigne, 'Enseignes', 'CHOIXCOD', 'CC_CODE', 'CC_LIBELLE', 'CC_TYPE="REN"', 'CC_LIBELLE', True, NumDisp, '', tlDefault )) then
        Result := RechDom('RTENSEIGNECODE', stEnseigne.Text, FALSE);
      {$ENDIF ERADIO}
    end
    else
    begin
      stEnseigne.Parent := Enseigne.Parent;
      stEnseigne.visible := False;
      if Assigned(Siret) then
        stEnseigne.Top := Siret.Top;
      stEnseigne.Left := Enseigne.Left + Enseigne.Width;
      stWhere := 'T_ENSEIGNE<>"" AND T_ENSEIGNE LIKE "' + Enseigne.Text + '%"';
      {$IFNDEF EAGLCLIENT}
        stEnseigne.text := Enseigne.Text;
      {$ENDIF}
      {$IFNDEF ERADIO}
        {$IFNDEF EAGLCLIENT}
          if (LookupList(stEnseigne, 'Enseigne', 'TIERS','DISTINCT T_ENSEIGNE','',stWhere,'',True,0,'',tlLocate )) then
        {$ELSE}
          if (LookupList(stEnseigne, 'Enseigne', 'TIERS','DISTINCT T_ENSEIGNE','',stWhere,'',True,0,'',tlDefault )) then
        {$ENDIF}
          Result := stEnseigne.Text;
      {$ENDIF ERADIO}
    end;
  finally
    stEnseigne.Destroy;
  end;
{$ENDIF !EAGLSERVER}
end;

function GetLibTiers(CliFou : string; Code : string; var Libelle : string) : boolean;
var QQ : TQuery;
begin

	result := false;

  QQ := OpenSql ('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+Code+'" AND T_NATUREAUXI="'+Clifou+'"',True,1,'',true);

  if not QQ.eof then
  begin
    Libelle := QQ.findField('T_LIBELLE').AsString;
    result := true;
  end;

  ferme (QQ);

end;

function GetTiersFerme(CliFou : string; Code : string; var Libelle : string) : boolean;
var QQ : TQuery;
begin

	result := false;

  QQ := OpenSql ('SELECT T_FERME FROM TIERS WHERE T_TIERS="'+Code+'" AND T_NATUREAUXI="'+Clifou+'"',True,1,'',true);

  if not QQ.eof then
  begin
    if QQ.findField('T_FERME').AsString = 'X' then result := True;
  end;

  ferme (QQ);

end;

function FormatNomLib(Nomtiers : String) : String;
Var i:integer;
begin

  Result := '';

  for i:=1 to Length(Nomtiers) do
  if NomTiers[i]='/' then NomTiers[i]:='_'        //Remplace le '/' par le '_'
  else if NomTiers[i]='\' then NomTiers[i]:='_'   //Remplace le '\' par le '_'
  else if NomTiers[i]='"' then NomTiers[i]:='_'   //Remplace le '"' par le '_'
  else if NomTiers[i]='&' then NomTiers[i]:='_'   //Remplace le '&' par le '_'
  else if NomTiers[i]='~' then NomTiers[i]:='_'   //Remplace le '~' par le '_'
  else if NomTiers[i]='.' then NomTiers[i]:='_'   //Remplace le '.' par le '_'
  else if NomTiers[i]='%' then NomTiers[i]:='_'   //Remplace le '%' par le '_'
  else if NomTiers[i]='*' then NomTiers[i]:='_'   //Remplace le '*' par le '_'
  else if NomTiers[i]='^' then NomTiers[i]:='_'   //Remplace le '^' par le '_'
  else if NomTiers[i]='¨' then NomTiers[i]:='_'   //Remplace le '¨' par le '_'
  else if NomTiers[i]=';' then NomTiers[i]:='_'   //Remplace le ';' par le '_'
  else if NomTiers[i]=',' then NomTiers[i]:='_'   //Remplace le ',' par le '_'
  else if NomTiers[i]=':' then NomTiers[i]:='_'   //Remplace le ':' par le '_'
  else if NomTiers[i]='?' then NomTiers[i]:='_'   //Remplace le '?' par le '_'
  else if NomTiers[i]='''' then NomTiers[i]:='_'; //Remplace le '''' par le '_'

  Result :=Nomtiers; //Affiche le résultat

end;


Initialization
//=== pour Script AGL  =========
RegisterAglFunc ( 'TiersAuxiliaire', False , 2, AGLTiersAuxiliaire);
RegisterAglFunc('VerifFormateChamp', FALSE, 3, VerifFormateChamp) ;
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
  RegisterAglProc('AglFicheAdressesTiers', True, 3, AglFicheAdressesTiers);
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

end.
