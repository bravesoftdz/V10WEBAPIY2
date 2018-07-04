// FQ 17215 - TGA 21/12/2005 - GetParamSoc => GetParamSocSecur
// MVG - 28/02/2006 - Report des corrections de C. Ayel
// BTY - 03/06 - Remonter ici la class du contexte d'édition de AmEdServer
// MVG - 13/04/2006 - ImBourreetless, ajout du paramètre optionnel ForceLg
// MVG - 13/04/2006 - Ajout des fonctions de SERIE1 getDateClotureImmo et getExoCloImmo
// MVG - 13/04/2006 - Ajout pour compilation BL700
// BTY - 05/06 - Positionner l'indicateur de modif de compta en eaglclient aussi
// XVI - 06/06/2006 - FQ (BL) 12422 Vérif conformité FD
// BTY - 07/06 - Ajout IMRempliExoDate pour compilation SERIE1
// MVG 12/07/2006 pour correction SERIE1 => XVI 18/04/2006 FD 3978 avant 5;
// MVG 06/06/2006 - uses que pour SERIE1 et ajout des nouvelles TOF
// MVG 11/09/2006 - Pour Delphi 7
// MVG 13/09/2006 - FQ 18791
// BTY 09/06 Restrictions utilisateur sur les établissements cf FQ 16149
// BTY 09/06 Applicable uniquement à des COMBOS THMultiValComboBox
// MVG 05/10/2006 - FQ 18920 - Impression en Web Access
// MVG 30/10/2006 - FQ 19046 - Création des motifs de sortie par défaut
// MVG 01/02/2007 Motif 999 Remplacement d'un composant de 2ème catégorie
// XVI 28/02/2007 S/FQ Nouveaux types des variants pour D7 (préconisation CBP)
// MVG 03/07/2007 Modification du libellé du modif 999
// MVG 17/08/2007 FQ 21253
// BTY 15/11/07 Date passage au régime réel (Agricoles)


unit ImEnt;

interface

uses Dialogs, Classes, StdCtrls, ComCtrls, Forms, SysUtils,
     UTob, Hent1, HCtrls, Paramsoc, Controls, ImPlanInfo,hmsgbox,HTB97,
{$IFDEF EAGLCLIENT}

{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}

{$IFDEF EAGLSERVER}
  {$IFNDEF SERIE1}
  uLibCpContexte,  // BTY 03/06
  {$ENDIF}
  eSession,        // BTY 03/06
{$ENDIF}

{$IFDEF VER150}
  Variants,
{$ENDIF}

     {$IFDEF SERIE1}
      {$IFDEF eAGLServer}
      {$ELSE}
      UtTX_,
      {$ENDIF !eAGLServer}
      S1Util, Ut
     {$ELSE}
     uLibExercice
      {$IFDEF EAGLSERVER}
      {$ELSE}
      ,UtilPgi
      {$ENDIF}
     , Ent1
     {$IFDEF MODENT1}
     , CPTypeCons
     {$ENDIF}
     {$ENDIF}
     {$IFDEF A_SUPPRIMER?}
     , Forms
     , Classes
     {$ENDIF !A_SUPPRIMER?}
    {$IFDEF MODENT1}
    , CPProcGen
    , CPProcMetier
    {$ENDIF MODENT1}
  ,UentCommun    
     ;


const
{$IFDEF SERIE1}
  ImMaxAxe = 3 ; //XVI 18/04/2006 FD 3978 avant 5;
{$ELSE}
  ImMaxAxe = MaxAxe;
{$ENDIF}

{ne pas changer l'ordre des 5 premiers (axes)}
Type
  TImExoDate = TExoDate;

  TInfoLog = record
    TVARecuperable: double;
    TVARecuperee: double;
  end;

type
  TBaseImmo = record
    CompteImmo: string;
    Fournisseur: string;
    Reference: string;
    CodeEtab: string;
    DateAchat: TDateTime;
    MontantHT: double;
    MontantTVA: double;
    Quantite: double;
    Libelle: string;
    {$IFDEF SERIE1}
    Ventilable : Boolean ; //XVI 18/04/2006 FD 3978
    {$ELSE}
    {$ENDIF !SERIE1}
  end;


 //YCP 25/08/05  fin

Type LaVariableImmo = RECORD
     TenueEuro:  boolean ;
     EnCours,Suivant,Precedent : TExoDate ;
     Cpta : Array[fbAxe1..fbAux] of TInfoCpta ;
     EtablisDefaut : string[3];
     EtablisCpta : boolean;
     //YCP 31-07-01 OBImmo   : TOB ;
     PlanInfo : TPlanInfo ;
     // Immobilisations
     Exercices : Array [1..20] of TExoDate ;
     CpteCBInf,CpteCBSup : string[17];
     CpteLocInf,CpteLocSup : string[17];
     CpteDepotInf,CpteDepotSup : string[17];
     CpteImmoInf,CpteImmoSup : string[17];
     CpteFinInf,CpteFinSup : string[17];
     CpteAmortInf,CpteAmortSup : string[17];
     CpteDotInf,CpteDotSup : string[17];
     CpteExploitInf,CpteExploitSup : string[17];
     CpteDotExcInf,CpteDotExcSup : string[17];
     CpteRepExcInf,CpteRepExcSup : string[17];
     CpteVaCedeeInf,CpteVaCedeeSup : string[17];
     CpteDerogInf,CpteDerogSup : string[17];
     CpteProvDerInf,CpteProvDerSup : string[17];
     CpteRepDerInf,CpteRepDerSup : string[17];
     Specif : Array[0..9] Of Boolean ;
     ChargeOBImmo : boolean;
     SpeedCum : Boolean ;
     AttribRibAuto : boolean;
     JalATP,JalVTP : String[3] ;
     CoeffDegressif : TOB;
     InsertMotifcesssion : boolean;
    end;

// BTY 03/06
{$IFDEF EAGLSERVER}
// MVG 12/07/2006
{$IFDEF SERIE1}
type
   TCPContexte = class(TObject)
  protected { Déclarations protégées}
     constructor Create( MySession : TISession ) ; virtual;
  public
     destructor  Destroy ; override ;
     class function  GetCurrent  : TCPContexte ; // retourne un pointeur sur le contexte de l'utilisateur, ou le cree s'il n'existe pas
     class procedure Release ; // detruit le contexte de l'utilisateur

   end ;
{$ENDIF}

type
  TAmortissementContext = class (TCPContexte)

  private { Déclarations privées}
    FPlanInfo: TPlanInfo;
    // MVG 05/10/06 FQ 18920
    SavCalcEdt: TProcCalcEdt;
    procedure ChargeInfoExo;
    // BTY 02/06 Rattacher les paramètres de TVA au contexte d'édition
    procedure ChargeInfoTVA;
    // BTY 03/06 Rattacher les coefficients dégressifs au contexte d'édition
    procedure ChargeCoeffDegressif;

  protected { Déclarations protégées}
    // BTY 02/06 FQ 17547   Montants faux en Web Access
    // => Démarrer à l'indice 0 pour pouvoir passer le tableau en paramètre
    Exercices: array[0..20] of TExoDate; //array[1..20]
    // BTY 02/06
    TobTva: TOB;
    // BTY 03/06
    CoeffDegressif : TOB;
    constructor Create( MySession : TISession ) ; override;
    //YPC suppr. conseil destructor Destroy; override; MVG 12/07/2006

  public { Déclarations publiques}
    destructor Destroy; override;
    property PlanInfo : TPlanInfo read FPlanInfo;
  end;
{$ENDIF}


var VHImmo : ^LaVariableImmo;

procedure ChargeVHImmo;
procedure ChargeImoExo ;


// Fonction communes S1-S5 Immos.
{$IFNDEF EAGLSERVER}  //YCP 25/08/05
Procedure ImLibellesTableLibre ( PZ : TTabSheet ; NamL,NamH,Pref : String ) ;
{$ENDIF EAGLSERVER}
function  ImAxeToFb(Axe: String ): TFichierBase ;
function ImBourreLaDoncSurLesComptes(st : string ; c : string = '') : string;
function ImBourreEtLess ( St : String ; LeType : TFichierBase ; ForceLg : Integer = 0 ) : string ;
function ImGeneTofb : TFichierBase;
function ImAuxTofb : TFichierBase;
function ImJalTofb : TFichierBase;
function ImImmoTofb : TFichierBase;
{$IFDEF EAGLSERVER}
{$ELSE}
function  ImTVA2TAUX(ModeTVA,Tva : String3 ; Achat : Boolean) : Real ;
{$ENDIF}
function ImCHARGEMAGHALLEY : Boolean ;
procedure ImDirDefault(Sauve : TSaveDialog ; FileName : String) ;
{$IFDEF EAGLSERVER}
{$ELSE}
Procedure ImMarquerPublifi ( Flag : boolean ) ;
function ImBlocage ( Niveau : Array of String ; SaufMoi : boolean ; Quoi : String  ; QueCloture : Boolean = FALSE) : boolean ;
Function  ImBlocageMonoPoste ( Totale : boolean ; TypeBlocage : String = 'NRTOUTSEUL' ; Shunte : Boolean = FALSE) : boolean ;
procedure ImDeblocageMonoPoste ( Totale : boolean ; TypeBlocage : String = 'NRTOUTSEUL' ; Shunte : Boolean = FALSE) ;
function  ImExisteCarInter ( St : String ) : boolean ;
{$ENDIF}
function ImccSecModif : integer;
Function ImGetPeriode ( LaDate : TDateTime ) : integer ;
{$IFDEF EAGLSERVER}
{$ELSE}
procedure ImExoToDates ( Exo : String3 ; ED1,ED2 : TControl ) ;
{$ENDIF}
procedure ImRempliExoDate (Var Exo : TExoDate ) ;  // BTY 07/06
function ImDupliqueVentil(CodeImmo,CodeDest: string): boolean ;
procedure SommeChampsTob(T1,T2: Tob;FieldNamez: array of string;SaufFieldNamez: array of string) ;
function ImQuelExoDt(DD : TDateTime) : string;
function ImQuelDateDeExo(Exo : string ; var ExoDt : TExoDate) : boolean;
procedure ImNOMBREPEREXO (ExoDt : TExoDate;  var PremMois,PremAnnee,nMois : Word);
procedure AmChargeCoefficientDegressif ( T : TOB );
function ImCQuelExercice( Date : TDateTime ; var Exo : TExoDate) : boolean;
(*
procedure ImZoomEdtEtatImmo(Quoi : String) ;
Procedure ImZoomEdtEtat(Quoi : String) ;
function ImCalcOLEEtat(sf,sp : string) : variant ;
*)
{$IFDEF EAGLSERVER}
{$ELSE}
FUNCTION ImCompte2Taux(Compte : String; Achat:boolean=false) : Double ;
{$ENDIF}
function GetModeReglementTiers(Auxiliaire: string): string ;

{$IFNDEF EAGLSERVER}  //YCP 25/08/05
procedure AjouteImmoS1(pE_:Tob ) ;
{$ENDIF EAGLSERVER}

{$IFDEF SERIE1}
function getDateClotureImmo(): tDateTime ;
function getExoCloImmo(): string ;

procedure initialisationImEnt ;
procedure finalisationImEnt ;

// BTY 09/06 FQ 16149
{$ELSE}
procedure ImGereEtablissement (ComboEtab : TControl);
procedure ImControlEtab (ComboEtab : TControl);
{$ENDIF}

// BTY 15/11/07 Date passage au régime réel (Agricoles)
function AMDatePassageREGFR (Nature, Code: string; var stMsg : string): boolean;



implementation
{$IFNDEF EAGLSERVER}
  {$IFDEF SERIE1}
  uses Immo_TOm, IMMOCPTE_TOM, AMUNITEOEUVRE_TOF, AMREGROUPEMENT_TOF,
  AMLISTE_TOF, AMCLOTURE_TOF, AMINTEG_TOF, TIMOTIFCESSION_TOT,
  AMCHGTMETHODE_TOF, AMDEPRECIATION_TOF, AMDEPRECGROUPE_TOF,
  AMSREGRO_TOF, AMMODIFSERVICE_TOF, AMANALYSE_TOF, AMEDITION_TOF,
  AMNDEDUCTION_TOF, AMSUIVIDPI_TOF, AMSYNTHESEDPI_TOF,
  //MVG 06/09/2006
  AMBASEECO_TOF, AMDPIANTERIEURS_TOF, AMIMMODPI_TOF,
  // MVG 12/10/2006
  AMPRIME_TOF, AMREDUCPRIME_TOF,
  // MVG 13/10/2006
  AMSUBVENTION_TOF
  {$ELSE}
  uses UtilSoc
  {$ENDIF}
  ;
{$ENDIF EAGLSERVER}

// ----------------------------------------------------------
// BTY 15/11/07 Date passage au régime réel (Agricoles)
function AMDatePassageREGFR (Nature, Code : string; var stMsg : string): boolean;
const
  REGFR_VIDE = 'Veuillez définir la date du passage au régime réel.';
  REGFR_ERR = ' n''est pas une année correcte pour dater le passage au régime réel.';
  REGFR_EXO = 'La date du passage au régime réel ne coïncide pas avec le début de l''exercice en cours.';
var
  stAnnee, stDeb : string;
  AnneeREGFR : integer;
  DateREGFR :TDateTime;
  a1900, a2099, m,j : word;
begin
  result := True;
  stMSg := '';
  if (Nature <>'IFR') or (Code <> 'IFR') then
      exit;
  stAnnee := TRIM (GetParamSocSecur('SO_DATEREGREEL',''));
  stDeb := Copy(stAnnee, 1, 1);

  if (stAnnee = '') then
  begin
      stMsg := REGFR_VIDE;
      result := False;
      exit;
  end else
  if (not IsNumeric(stAnnee)) or (Length (stAnnee) <> 4)
        or ((stDeb <> '1') and (stDeb <> '2')) then
  begin
      stMsg := stAnnee + REGFR_ERR;
      result := False;
      exit;
  end;
  DecodeDate(iDate1900, a1900, m, j);
  DecodeDate(iDate2099, a2099, m, j);
  AnneeREGFR := StrToInt(stAnnee);
  if (AnneeREGFR < a1900) or (AnneeREGFR > a2099) then
  begin
     stMsg := REGFR_VIDE;
     result := False;
     exit;
  end else
  begin
     DateREGFR := EncodeDate (AnneeREGFR, 1, 1);
     if (DateREGFR <> VHImmo^.Encours.Deb) then
        begin
        stMsg := REGFR_EXO;
        result := False;
        exit;
      end;
  end;
end;
// ------------------------------

Procedure InsertLesMotifs ( Nombase : string = '');
var stTable : string;
begin
  if NomBase = '' then stTable := 'CHOIXCOD'
  else stTable := NomBase+'.dbo.CHOIXCOD';
ExecuteSQL('DELETE FROM '+stTable+' WHERE CC_TYPE="MDC"') ;
// MVG 17/08/07 ajout de TraduireMemoire FQ 21253
ExecuteSQL('INSERT INTO '+stTable+' (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("MDC","001","'+ TraduireMemoire('Cession') +'","","")') ;
ExecuteSQL('INSERT INTO '+stTable+' (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("MDC","002","'+ TraduireMemoire('Rebut') +'","","")') ;
ExecuteSQL('INSERT INTO '+stTable+' (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("MDC","003","'+ TraduireMemoire('Vol') +'","","")') ;
end ;

Procedure InsertLeMotif999 ( Nombase : string = '');
var stTable : string;
begin
  if NomBase = '' then stTable := 'CHOIXCOD'
  else stTable := NomBase+'.dbo.CHOIXCOD';
// MVG 17/08/07 ajout de TraduireMemoire FQ 21253
ExecuteSQL('INSERT INTO '+stTable+' (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("MDC","999","'+ TraduireMemoire('Rempl. composant 2ème cat.') +'","","")') ;
end ;


procedure InitLaVariableIMO ;
begin
  New(VHImmo) ;
  FillChar(VHImmo^,Sizeof(VHImmo^),#0) ;
  //YCP 31-07-01 VHImmo^.OBImmo:=TOB.Create('',Nil,-1) ;
end ;

procedure ChargeImoExo ;
var Premier: boolean ; DDeb,DFin : TDateTime ;  SCode,SEtat : String ; i: integer ; Q: TQuery ;
begin
  i:=1 ; Premier:=true ;
  Q:=OpenSQL('SELECT EX_DATEDEBUT,EX_DATEFIN,EX_EXERCICE,EX_ETATCPTA FROM EXERCICE ORDER BY EX_DATEDEBUT',TRUE) ;
  try
    if not Q.Eof then
    begin
      Q.First ;
      while ((not Q.Eof) and (i<=20)) do
      begin
        DDeb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
        DFin:=Q.FindField('EX_DATEFIN').AsDateTime ;
        SCode:=Q.FindField('EX_EXERCICE').AsString ;
        SEtat:=Q.FindField('EX_ETATCPTA').AsString ;
        VHImmo^.Exercices[i].Code := SCode ;
        VHImmo^.Exercices[i].Deb := DDeb ;
        VHImmo^.Exercices[i].Fin := DFin ;
        if SEtat='CDE' then
        begin
        VHImmo^.Precedent.Deb:=DDeb ;
        VHImmo^.Precedent.Fin:=DFin ;
        VHImmo^.Precedent.Code:=SCode ;
        end ;
        if (SEtat='OUV') Or (SEtat='CPR') then
        begin
          if Premier then
          begin
            VHImmo^.Encours.Deb:=DDeb ;
            VHImmo^.Encours.Fin:=DFin ;
            VHImmo^.Encours.Code:=SCode ;
          end
          else
          begin
            VHImmo^.Suivant.Deb:=DDeb ;
            VHImmo^.Suivant.Fin:=DFin ;
            VHImmo^.Suivant.Code:=SCode ;
            Inc(i) ; // CA - 19/04/2004 - Sinon, on écrase le derneir exercice
            Break ;
          end ;
          Premier:=False ;
        end ;
        Inc(i) ;
        Q.Next; //YCP OK
      end;
    end
  finally
    Ferme(Q) ;

    VHImmo^.Exercices[i].Code := '';
    VHImmo^.Exercices[i].Deb := iDate1900;
    VHImmo^.Exercices[i].Fin := iDate1900;
  end ;
end ;

procedure ChargeVHImmoCommun;
begin
  ChargeImoExo ;
  VHImmo^.CpteCBInf     :=GetParamSocSecur('SO_CPTECBINF','');      VHImmo^.CpteCBSup     :=GetParamSocSecur('SO_CPTECBSUP','');
  VHImmo^.CpteLocInf    :=GetParamSocSecur('SO_CPTELOCINF','');     VHImmo^.CpteLocSup    :=GetParamSocSecur('SO_CPTELOCSUP','');
  VHImmo^.CpteDepotInf  :=GetParamSocSecur('SO_CPTEDEPOTINF','');   VHImmo^.CpteDepotSup  :=GetParamSocSecur('SO_CPTEDEPOTSUP','');

  VHImmo^.CpteImmoInf   :=GetParamSocSecur('SO_CPTEIMMOINF','');    VHImmo^.CpteImmoSup   :=GetParamSocSecur('SO_CPTEIMMOSUP','');
  VHImmo^.CpteFinInf    :=GetParamSocSecur('SO_CPTEFININF','');     VHImmo^.CpteFinSup    :=GetParamSocSecur('SO_CPTEFINSUP','');
  VHImmo^.CpteAmortInf  :=GetParamSocSecur('SO_CPTEAMORTINF','');   VHImmo^.CpteAmortSup  :=GetParamSocSecur('SO_CPTEAMORTSUP','');
  VHImmo^.CpteDotInf    :=GetParamSocSecur('SO_CPTEDOTINF','');     VHImmo^.CpteDotSup    :=GetParamSocSecur('SO_CPTEDOTSUP','');
  VHImmo^.CpteExploitInf:=GetParamSocSecur('SO_CPTEEXPLOITINF',''); VHImmo^.CpteExploitSup:=GetParamSocSecur('SO_CPTEEXPLOITSUP','');
  VHImmo^.CpteDotExcInf :=GetParamSocSecur('SO_CPTEDOTEXCINF','');  VHImmo^.CpteDotExcSup :=GetParamSocSecur('SO_CPTEDOTEXCSUP','');
  VHImmo^.CpteRepExcInf :=GetParamSocSecur('SO_CPTEREPEXCINF','');  VHImmo^.CpteRepExcSup :=GetParamSocSecur('SO_CPTEREPEXCSUP','');
  VHImmo^.CpteVaCedeeInf:=GetParamSocSecur('SO_CPTEVACEDEEINF',''); VHImmo^.CpteVaCedeeSup:=GetParamSocSecur('SO_CPTEVACEDEESUP','');
  VHImmo^.CpteDerogInf  :=GetParamSocSecur('SO_CPTEDEROGINF','');   VHImmo^.CpteDerogSup  :=GetParamSocSecur('SO_CPTEDEROGSUP','');
  VHImmo^.CpteProvDerInf:=GetParamSocSecur('SO_CPTEPROVDERINF',''); VHImmo^.CpteProvDerSup:=GetParamSocSecur('SO_CPTEPROVDERSUP','');
  VHImmo^.CpteRepDerInf :=GetParamSocSecur('SO_CPTEREPDERINF','');  VHImmo^.CpteRepDerSup :=GetParamSocSecur('SO_CPTEREPDERSUP','');
  { Chargement des coefficients dégressifs }
  if VHImmo^.CoeffDegressif= nil then
  begin
    VHImmo^.CoeffDegressif := TOB.Create ('', nil, -1);
    AmChargeCoefficientDegressif ( VHImmo^.CoeffDegressif );
  end;
  // Création des motifs de sortie par défaut
  if VHImmo^.InsertMotifcesssion= false then
  begin
    if not ExisteSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="MDC"') then
    InsertLesMotifs ;
    VHImmo^.InsertMotifcesssion:=true;
    if not ExisteSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="MDC" AND CC_CODE="999"') then
    InsertLeMotif999 ;

  end;
end ;

procedure ChargeVHImmo;
{$IFDEF SERIE1}
begin
//  InitLaVariableIMO ;
  VHImmo^.EtablisDefaut:='001';
  VHImmo^.TenueEuro       :=v_pgi.TenueEuro ;
  VHImmo^.Cpta[fbGene].Lg :=VS1Cpta.LEN_GENE ;
  VHImmo^.Cpta[fbGene].Cb :='0' ;
  VHImmo^.Cpta[fbAux].Lg  :=VS1Cpta.LEN_GENE;
  VHImmo^.Cpta[fbAux].Cb  :='0' ;
  VHImmo^.JalATP:='' ;
  VHImmo^.JalVTP:='' ;

  //Immo Commun S1/S5
  ChargeVHImmoCommun ;
end ;
{$ELSE}
var fb : TFichierBase ;
begin
//  InitLaVariableIMO ; fait dans l'initialization en bas !
  VHImmo^.EtablisDefaut:=GetParamSocSecur('SO_ETABLISDEFAUT','');
  VHImmo^.EtablisCpta:= GetParamSocSecur('SO_ETABLISCPTA','');
  VHImmo^.TenueEuro      :=GetParamSocSecur('SO_TENUEEURO','');
{$IFDEF EAGLSERVER}
  for fb:=fbAxe1 to fbAxe5 do VHImmo^.Cpta[fb] :=TAmortissementContext.GetCurrent.InfoCpta.Cpta[fb] ;
  VHImmo^.Cpta[fbGene]   := TAmortissementContext.GetCurrent.InfoCpta.Cpta[fbGene];
  VHImmo^.Cpta[fbAux]    := TAmortissementContext.GetCurrent.InfoCpta.Cpta[fbAux];
{$ELSE}
  for fb:=fbAxe1 to fbAxe5 do VHImmo^.Cpta[fb] :=VH^.Cpta[fb] ;
  VHImmo^.Cpta[fbGene]   :=VH^.Cpta[fbGene] ;
  VHImmo^.Cpta[fbAux]    :=VH^.Cpta[fbAux] ;
{$ENDIF}
  VHImmo^.AttribRibAuto  := GetParamSocSecur('SO_ATTRIBRIBAUTO','');
  VHImmo^.JalATP:=GetParamSocSecur('SO_JALATP','');
  VHImmo^.JalVTP:=GetParamSocSecur('SO_JALVTP','');

  //Immo Commun S1/S5
  ChargeVHImmoCommun ;
end;
{$endIF}

{$IFNDEF EAGLSERVER}  //YCP 25/08/05 
Procedure ImLibellesTableLibre ( PZ : TTabSheet ; NamL,NamH,Pref : String ) ;
{$IFDEF SERIE1}
begin
if length(Pref)=1 then Pref:=Pref+'0' ;
  LibelleTableLibre (PZ , NamL,Pref ) ;
{$ELSE}
begin
  LibellesTableLibre(PZ , NamL,NamH,Pref ) ;
{$endIF}
end ;
{$ENDIF EAGLSERVER}

function ImBourreLaDoncSurLesComptes(st : string ; c : string = '') : string;
{$IFDEF SERIE1}
var st1 : string ;
    l : integer;
    cb : string;
begin
  if c = '' then cb := '0' else cb := c;
  st1 := Trim(st) ;
  l := VS1Cpta.LEN_GENE ;
  if l<= 0 then {!!!!????} else
  if l<Length(st1) then st1 := Copy(st1,1,l);
  while Length(st1)<l do st1 := st1 + cb ;
  Result := st1 ;
{$ELSE}
begin
  result:=BourreLaDoncSurLesComptes(st , c );
{$endIF}
end;

function  ImAxeToFb(Axe: String ): TFichierBase ;
{$IFDEF SERIE1}
begin
  Result:=fbAxe1 ;
  if Length(Axe)>=2 then Result:=TFichierBase(Ord(Axe[2])-49) ;
{$ELSE}
begin
  result:=AxeToFb(Axe) ;
{$endIF}
end ;


function ImBourreEtLess ( St : String ; LeType : TFichierBase ; ForceLg : Integer = 0 ) : string ;
{$IFDEF SERIE1}
var Lg,ll,i : Integer ;
    Bourre  : Char ;
begin
If LeType In [fbAxe1..fbAux,fbNatCpt] Then
   begin
   Lg:=VHImmo^.Cpta[LeType].Lg ;
   // MVG 13/09/06 FQ 18791
   If ForceLg>0 then Lg:=ForceLg;
   // Fin MVG
   Bourre:='0' ;
   If Length(St)>Lg Then St:=Trim(Copy(St,1,Lg)) ;
   Result:=St ; ll:=Length(Result) ;
   If ll<Lg then
      begin
      for i:=ll+1 to Lg do Result:=Result+Bourre ;
      end Else Result:=Copy(Result,1,Lg) ;
   end Else Result:=St ;
{$ELSE}
begin
Result := BourreEtLess ( St ,LeType, ForceLg);
{$endIF}
end ;

function ImGeneTofb : TFichierBase;
begin
  Result := fbGene;
end;

function ImAuxTofb : TFichierBase;
begin
  Result := fbAux;
end;

function ImJalTofb : TFichierBase;
begin
  Result := fbJal;
end;

function ImImmoTofb : TFichierBase;
begin
  Result := fbimmo;
end;

function ImccSecModif : integer;
begin
{$IFDEF SERIE1}
result:=0 ;
{$ELSE}
Result:=ccSecModif;
{$ENDIF}
end;

{$IFDEF EAGLSERVER}
{$ELSE}
function  ImTVA2TAUX(ModeTVA,Tva : String3 ; Achat : Boolean) : Real ;
begin
  {$IFDEF SERIE1}
  Result:=TVA2TAUX(Tva,ModeTVA,Achat);
  {$ELSE}
  Result:=TVA2TAUX(ModeTVA,Tva ,Achat);
  {$ENDIF}
end;
{$ENDIF}

function ImCHARGEMAGHALLEY : Boolean ;
begin
{$IFDEF SERIE1}
  result:=ChargeSerie1 (False) ;
{$ELSE}
//  Result:=ChargeMagHalley ;
  result := True;
{$endIF}
end;

procedure ImDirDefault(Sauve : TSaveDialog ; FileName : String) ;
{$IFDEF SERIE1}
var j,i : integer ;
begin
j:=Length(FileName);
for i:=Length(FileName) downto 1 do if FileName[i]='\' then begin j:=i ; Break ; end ;
Sauve.InitialDir:=Copy(FileName,1,j) ;
{$ELSE}
begin
DirDefault(Sauve ,FileName) ;
{$endIF}
end ;

{$IFDEF EAGLSERVER}
{$ELSE}
Procedure ImMarquerPublifi ( Flag : boolean ) ;
begin
{$IFDEF SERIE1}
if Flag then SetParamSoc('SO_ACHARGERPUBLIFI','X') else SetParamSoc('SO_ACHARGERPUBLIFI','-') ;
{$ELSE}
  // BTY 05/06 MarquerPublifi en eaglclient
  //{$IFDEF EAGLCLIENT}
  //
  //{$ELSE}
    MarquerPublifi(Flag) ;
  //{$ENDIF}
{$ENDIF}
end ;
{$ENDIF}

Function ImGetPeriode ( LaDate : TDateTime ) : integer ;
{$IFDEF SERIE1}
Var YY,MM,DD : Word ;
BEGIN
YY:=0 ; MM:=0 ;
try DecodeDate(LaDate,YY,MM,DD) ; finally Result:=100*YY+MM ; end ;
{$ELSE}
begin
Result:= GetPeriode (LaDate);
{$ENDIF}
END ;

{$IFDEF EAGLSERVER}
{$ELSE}
procedure ImExoToDates ( Exo : String3 ; ED1,ED2 : TControl ) ;
{$IFDEF SERIE1}
Var D1,D2 : TDateTime ;
    Q     : TQuery;
    Okok  : boolean ;
BEGIN
if EXO='' then Exit ;
Okok:=True ; D1:=Date ; D2:=Date ;
If EXO=VHImmo^.Precedent.Code Then BEGIN D1:=VHImmo^.Precedent.Deb ; D2:=VHImmo^.Precedent.Fin ; END Else
If EXO=VHImmo^.EnCours.Code Then   BEGIN D1:=VHImmo^.Encours.Deb ;   D2:=VHImmo^.Encours.Fin ; END Else
If EXO=VHImmo^.Suivant.Code Then   BEGIN D1:=VHImmo^.Suivant.Deb ;   D2:=VHImmo^.Suivant.Fin ; END Else
   BEGIN
   Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exo+'"' ,TRUE) ;
   if Not Q.EOF then
      BEGIN
      D1:=Q.FindField('EX_DATEDEBUT').asDateTime ; D2:=Q.FindField('EX_DATEFIN').asDateTime ;
      END else Okok:=False ;
   Ferme(Q) ;
   END;
if Okok then BEGIN TEdit(ED1).Text:=DateToStr(D1) ; TEdit(ED2).Text:=DateToStr(D2) ; END ;
{$ELSE}
begin
ExoToDates ( Exo,ED1,ED2  );
{$ENDIF}
END ;
{$ENDIF}

function imGetPrecedent: TExoDate ;
{$IFDEF SERIE1}
begin
  Result := VHImmo^.Precedent;
{$ELSE}
begin
  result:=GetPrecedent ;
{$ENDIF}
end ;

function imGetEnCours : TExoDate ;
{$IFDEF SERIE1}
begin
  Result := VHImmo^.Encours;
{$ELSE}
begin
  result:=GetEncours ;
{$ENDIF}
end ;

function imGetSuivant : TExoDate ;
{$IFDEF SERIE1}
begin
  Result := VHImmo^.Suivant;
{$ELSE}
begin
  result:=GetSuivant ;
{$ENDIF}
end ;

// BTY 07/06
Procedure IMRempliExoDate (Var Exo : TExoDate ) ;
{$IFDEF SERIE1}
begin
if Exo.Code=imGetPrecedent.Code then
   BEGIN
   Exo.Deb:=imGetPrecedent.Deb ; Exo.Fin:=imGetPrecedent.Fin ;
   Exo.NombrePeriode:=imGetPrecedent.NombrePeriode ;
   END else
   if Exo.Code=imGetEnCours.Code then
      BEGIN
      Exo.Deb:=imGetEnCours.Deb ; Exo.Fin:=imGetEnCours.Fin ;
      Exo.NombrePeriode:=imGetEnCours.NombrePeriode ;
      END else
      if Exo.Code=imGetSuivant.Code then
         BEGIN
         Exo.Deb:=imGetSuivant.Deb ; Exo.Fin:=imGetSuivant.Fin ;
         Exo.NombrePeriode:=imGetSuivant.NombrePeriode ;
         END
         else // ajout ME 02/10/2003 pour les exercice n-2
              imQuelDateDeExo(Exo.Code,Exo) ;
{$ELSE}
begin
RempliExoDate (Exo);
{$ENDIF}
end;


{$IFDEF EAGLSERVER}
{$ELSE}
FUNCTION ImCompte2Taux(Compte : String; Achat:boolean=false) : Double ;
{$IFDEF SERIE1}
begin
result:=Compte2Taux(Compte,Achat) ;
end ;
{$ELSE}
var
{$IFDEF SPEC302}
  i: integer ;
{$ENDIF}
  TobT: tob ;
  t: string ;
begin
  {$IFDEF SPEC302}
  result:=0 ;
  For i:=1 to 100 do
    begin
    if Achat and (VH^.TabTVA[i].CpteACH=Compte) then begin result:=VS1.TabTVA[i].TauxACH ; break ; end ;
    if not Achat and (VH^.TabTVA[i].CpteVTE=Compte) then begin result:=VS1.TabTVA[i].TauxVTE ; break ; end ;
    end ;
  {$ELSE}
  result:=0 ;
  if Achat then t:='ACH' else t:='VTE' ;
  Tobt:=VH^.LaTOBTVA.FindFirst(['TV_CPTE'+t],[Compte],False) ;
  if TOBT<>Nil then Result:=TOBT.GetValue('TV_TAUX'+t) ;
  {$ENDIF}
end ;
{$ENDIF}
{$ENDIF}

{$IFDEF EAGLSERVER}
{$ELSE}
function ImBlocage ( Niveau : Array of String ; SaufMoi : boolean ; Quoi : String  ; QueCloture : Boolean = FALSE) : boolean ;
begin
{$IFDEF SERIE1}
result:=Blocage(Niveau,SaufMoi,Quoi) ;
{$ELSE}
result:=_Blocage(Niveau,SaufMoi,Quoi,QueCloture) ;
{$ENDIF}
end ;

function  ImBlocageMonoPoste ( Totale : boolean ; TypeBlocage : String = 'NRTOUTSEUL' ; Shunte : Boolean = FALSE) : boolean ;
begin
{$IFDEF SERIE1}
result:=BlocageMonoPoste(Totale,TypeBlocage) ;
{$ELSE}
result:=_BlocageMonoPoste(Totale,TypeBlocage, Shunte) ;
{$ENDIF}
end ;

procedure ImDeblocageMonoPoste ( Totale : boolean ; TypeBlocage : String = 'NRTOUTSEUL' ; Shunte : Boolean = FALSE) ;
BEGIN
{$IFDEF SERIE1}
DeBlocageMonoPoste(Totale,TypeBlocage) ;
{$ELSE}
_DeBlocageMonoPoste(Totale,TypeBlocage, Shunte) ;
{$ENDIF}
END ;
{$ENDIF}


{$IFDEF EAGLSERVER}
{$ELSE}
function  ImExisteCarInter ( St : String ) : boolean ;
  {$IFDEF SERIE1}
const TabInter: array[1..6] of string=('"','*','%','_','?','''') ;
var i: integer ;
begin
  Result:=False ;
  for i:=1 to 6 do if Pos('"',St)>0 then begin result:=true ; break ; end ;
 {$ELSE}
begin
  Result := ExisteCarInter(St) ;
  {$ENDIF}
end ;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 28/07/2003
Modifié le ... :   /  /
Description .. : Renvoie l'exercice associé à la date passée en
Suite ........ : paramètre.
Mots clefs ... :
*****************************************************************}
function ImCQuelExercice( Date : TDateTime ; var Exo : TExoDate) : boolean;
{$IFDEF SERIE1}
var i    : integer;
{$ENDIF}
begin
{$IFDEF SERIE1}
  i := 1; Result := False;
  while (VHImmo^.Exercices[i].Code<>'') do
  begin
    if (Date >= VHImmo^.Exercices[i].Deb) and  (Date <= VHImmo^.Exercices[i].Fin) then
    begin
      Exo := VHImmo^.Exercices[i];
      Result := true;
      break;
    end;
    Inc (i,1);
  end;
{$ELSE}
  Result := CQuelExercice( Date , Exo );
{$ENDIF}
end;

function GetModeReglementTiers(Auxiliaire: string): string ;
var Q: TQuery ;
begin
  result:='' ;
  if Auxiliaire='' then exit ;
  Q:=OpenSQL('SELECT T_MODEREGLE FROM TIERS WHERE T_AUXILIAIRE="'+Auxiliaire+'"',false) ;
  if not Q.Eof then result:=Q.Fields[0].AsString ;
  ferme(Q) ;
end ;

function ImDupliqueVentil(CodeImmo,CodeDest: string): boolean ;
var T1: Tob ; Q: TQuery ; i: integer ;
begin
  result:=false ;
  if (CodeImmo<>CodeDest) then
  begin
    if not ExisteSql('SELECT * FROM VENTIL WHERE V_COMPTE="'+CodeDest+'"') then
    begin
      T1:=Tob.Create('VENTIL',nil,-1) ;
      try
        Q:=OpenSql('SELECT * FROM VENTIL WHERE V_COMPTE="'+CodeImmo+'"',true) ;
        try
          T1.LoadDetailDB('VENTIL','','',Q,false) ;
          For i:=0 to T1.Detail.Count-1 do T1.Detail[i].PutValue('V_COMPTE',CodeDest) ;
          T1.SetAllModifie(true) ;
          T1.InsertOrUpdateDB ;
          result:=true ;
        finally
          ferme(Q) ;
        end ;
      finally
        T1.Free ;
      end ;
    end ;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 29/08/2002
Modifié le ... :   /  /    
Description .. : Retourne le code exercice qui contient la date passée en
Suite ........ : paramètre
Mots clefs ... :
*****************************************************************}
function ImQuelExoDt(DD : TDateTime) : string ;
begin
Result:='' ;
If (dd>=VHImmo^.EnCours.Deb) and (dd<=VHImmo^.EnCours.Fin) then Result:=VHImmo^.EnCours.Code else
   If (dd>=VHImmo^.Suivant.Deb) and (dd<=VHImmo^.Suivant.Fin) then Result:=VHImmo^.Suivant.Code Else
      If (dd>=VHImmo^.Precedent.Deb) and (dd<=VHImmo^.Precedent.Fin) then Result:=VHImmo^.Precedent.Code;
end ;

function ImQuelDateDeExo(Exo : string ; var ExoDt : TExoDate) : boolean;
{$IFDEF SERIE1}
Var
  Q : TQuery ;
BEGIN
ExoDt:=VHImmo^.EnCours ; Result:=FALSE ;
if Exo=VHImmo^.Precedent.Code then BEGIN Result:=TRUE ; ExoDt:=VHImmo^.Precedent ; END else
   if Exo=VHImmo^.EnCours.Code then BEGIN Result:=TRUE ; ExoDt:=VHImmo^.EnCours ; END else
      if Exo=VHImmo^.Suivant.Code then BEGIN Result:=TRUE ; ExoDt:=VHImmo^.Suivant ; END Else
         BEGIN
         (*For i:=1 To 5 Do BEGIN If Exo=VHImmo^.ExoClo[i].Code then
            BEGIN Exo:=VHImmo^.ExoClo[i] ; Result:=TRUE ; Exit ; END ; END ;*)
         Q:=OpenSQL('SELECT * FROM EXERCICE WHERE EX_EXERCICE="'+Exo+'"',TRUE) ;
         If Not Q.Eof Then
            BEGIN
            ExoDt.Code:=Q.FindField('EX_EXERCICE').AsString ;
            ExoDt.Deb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
            ExoDt.Fin:=Q.FindField('EX_DATEFIN').AsDateTime ;
            Result:=TRUE ;
            END ;
         Ferme(Q) ;
         END ;
{$ELSE}
begin
  {$IFDEF EAGLSERVER}
  ExoDt := TAmortissementContext.GetCurrent.Exercice.QuelExoDate ( Exo );
  {$ELSE}
  QuelDateDeExo(Exo,ExoDt);
  {$ENDIF}
  Result := (ExoDt.Code<>'');
{$ENDIF}
end;

procedure ImNOMBREPEREXO (ExoDt : TExoDate; var PremMois,PremAnnee,nMois : Word) ;
{$IFDEF SERIE1}
begin
  if ExoDt.Deb>ExoDt.Fin then nMois:=0
     else NombreMois(ExoDt.Deb,ExoDt.Fin,PremMois,PremAnnee,nMois) ; //YCP 06/02/03
{$ELSE}
begin
  NOMBREPEREXO (ExoDt,  PremMois,PremAnnee,nMois) ;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 31/07/2001
Modifié le ... : 31/07/2001
Description .. : Somme les champs "sommeable" d'une TOB, possibilité
Suite ........ : d'interdir le calcul de certains champs avec un tableau des
Suite ........ : champs à ne pas sommer.
Mots clefs ... : SOMME TOB
*****************************************************************}
procedure SommeChampsTob(T1,T2: Tob;FieldNamez: array of string;SaufFieldNamez: array of string) ;
var i,j,k: integer ; NomChamps: string ; Okok: boolean ;
begin
  if not (T1 is tob) or (T1=nil) then exit ;
  if not (T2 is tob) or (T2=nil) then exit ;
  for i:=1 to T1.NbChamps do
    begin
    NomChamps:=T1.GetNomChamp(i) ;
    Okok:=false ;
    if (high(FieldNamez)<>0) then
      begin
      for k:=Low(FieldNamez) to High(FieldNamez) do
        if NomChamps=FieldNamez[k] then
          begin
          okok:=true ;
          break ;
          end ;
      end
      else Okok:=true ;
    if (Okok=true) and (high(SaufFieldNamez)<>0) then
      begin
      for k:=Low(SaufFieldNamez) to High(SaufFieldNamez) do
        if NomChamps=SaufFieldNamez[k] then
          begin
          okok:=false ;
          break ;
          end ;
      end ;
    if Okok and
       //(varType(T1.Valeurs[i]) in [varByte,varSmallint,varInteger,varSingle,varDouble,varCurrency]) then //XVI 28/02/2007 S/FQ Nouveaux types des variants pour D7 (préconisation CBP)
       (varType(T1.Valeurs[i]) in [varSmallint,varInteger,varByte,varShortInt,varWord,varLongWord,varInt64,varSingle,varDouble,varCurrency]) then
      begin
      j:=T2.GetNumChamp(NomChamps) ;
      //if (j<>0) and (varType(T2.Valeurs[j]) in [varByte,varSmallint,varInteger,varSingle,varDouble,varCurrency]) then //XVI 28/02/2007 S/FQ Nouveaux types des variants pour D7 (préconisation CBP)
      if (j<>0) and (varType(T2.Valeurs[j]) in [varSmallint,varInteger,varByte,varShortInt,varWord,varLongWord,varInt64,varSingle,varDouble,varCurrency]) then
        T1.PutValeur(i,T1.Valeurs[i]+VarAsType(T2.Valeurs[j],VarType(T1.Valeurs[i]))) ;
      end ;
    end ;
end ;

///////////////////////////////////////////////////////////////////////////////
{$IFNDEF EAGLSERVER}  //YCP 25/08/05 
procedure AjouteImmoS1(pE_:Tob ) ;
{$IFDEF SERIE1}
var RecImmo: TBaseImmo ;
begin
   //XVI 18/04/2006 FD 3978 début
   if PgiAsk('Voulez-vous créer une fiche d''immobilisation?')=mryes then
   begin
      FillChar(RecImmo,Sizeof(RecImmo),#0) ;
      RecImmo.CompteImmo:=vString(pE_.GetValue('E_GENERAL'));
      //RecImmo.Fournisseur:=LibAux ;
      RecImmo.Reference:=vString(pE_.GetValue('E_REFERENCE'));
      //RecImmo.Libelle:=g.cells[Gc_Libelle, ARow];
      RecImmo.DateAchat:=vDate(pE_.GetValue('E_DATECOMPTABLE'));
      RecImmo.MontantHT:=vDouble(pE_.GetValue('E_DEBITDEV'))+vDouble(pE_.GetValue('E_CREDITDEV'));
      RecImmo.Quantite:=vDouble(pE_.GetValue('E_QTE1')) ;
      if RecImmo.Quantite=0 then RecImmo.Quantite:=1 ;
      //RecImmo.Ventilable:=StrToBool_(pE_.GetValue('E_ANA')) ; //XVI 18/04/2006 FD 3978
      RecImmo.Ventilable:=existeSQL('select G_GENERAL from GENERAUX where G_GENERAL="'+RecImmo.CompteImmo+'" and G_VENTILABLE="X"') ; //XVI 06/06/2006 FQ 12422
      pE_.PutValue('IMMO',CreationSaisieCompta(RecImmo)) ;
   end ;
   //XVI 18/04/2006 FD 3978 fin
{$ELSE}
begin
{$ENDIF SERIE1}
end ;
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 03/07/2003
Modifié le ... :   /  /
Description .. : Chargement des coefficients dégressifs
Mots clefs ... :
*****************************************************************}
procedure AmChargeCoefficientDegressif ( T : TOB );
var Q : TQuery;
    i : integer;
    St : string;
begin
  Q := OpenSQL ( 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="ICD"', True );
  T.LoadDetailDB('CHOIXCOD','','',Q,False);
  Ferme (Q);
  for i:=0 to T.Detail.Count - 1 do
  begin
    St := T.Detail[i].GetValue('CC_LIBELLE');
    T.Detail[i].AddChampSupValeur('DATE',ReadTokenSt(St));
    T.Detail[i].AddChampSupValeur('COEFF1',Valeur(ReadTokenSt(St)));
    T.Detail[i].AddChampSupValeur('COEFF2',Valeur(ReadTokenSt(St)));
    T.Detail[i].AddChampSupValeur('COEFF3',Valeur(ReadTokenSt(St)));
  end;
end;



{ TAmortissementContext }

// BTY 03/06
{$IFDEF EAGLSERVER}
// MVG 12/07/2006
{$IFDEF SERIE1}
constructor TCPContexte.Create( MySession : TISession ) ;
begin
// FMySession     := MySession ;
// FZExercice     := TZExercice.Create ;
// FZInfoCpta     := TZInfoCpta.create;
// FZHalleyUser   := TZHalleyUser.Create ;
// FZc
end;

destructor TCPContexte.Destroy;
begin
// FreeAndNil(FZExercice);
// FreeAndNil(FZInfoCpta);
// FreeAndNil(FZHalleyUser);
 inherited ;
end;

class function TCPContexte.GetCurrent : TCPContexte ;
var
 MySession      : TISession ;
 lIndex         : integer ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','TCPContexte.GetCurrent') ;
{$ENDIF}

 // on recupere la session courante
 MySession := LookupCurrentSession ;

 // on recherche si cette classe a deja ete cree
 // on utilise le classname pour stocker/retrouver un classe
 lIndex    := MySession.UserObjects.IndexOf('ID_IMMO') ;

 if lIndex > -1 then
  result := TCPContexte(MySession.UserObjects.Objects[lIndex]) // on renvoie un pointeur sur la classe courante
   else
    begin
     {$IFDEF TTW}
      cWA.MessagesAuClient('COMSX.IMPORT','','TCPContexte.GetCurrent : creation de la classe ' + ClassName) ;
     {$ENDIF}   
     // la classe n'existe pas il faut la creer
     // on recherche la classe ds l'ensemble des classes recensées et on la creer ds la foulée
     result       := TAmortissementContext.Create(MySession) ;
     
     //MVG 05/10/06 FQ 18920
     TAmortissementContext(result).SavCalcEdt:=LookUpCurrentSession.UserCalcEdt;

     //Result.Count := Result.Count + 1 ;
     MySession.UserObjects.AddObject('ID_IMMO',result) ; // on ajoute la nouvelle instance de classe a la session
     if ClassParent.InheritsFrom(TCPContexte) then
      MySession.UserObjects.AddObject(ClassParent.ClassName,result) ; // on ajoute la classe ancetre a la session si elle est de type TCpContexte
    end ;


end;


{***********A.G.L.***********************************************
Auteur  ...... : Manou Entressangle & Laurent Gendreau
Créé le ...... : 21/04/2005
Modifié le ... : 21/04/2005
Description .. : Methode de destruction du contexte courant
Mots clefs ... : 
*****************************************************************}
class procedure TCPContexte.Release ;
var
 MySession : TISession ;
 lIndex    : integer ;
 MyContext : tObject ;
begin
 // on recupere le pointeur sur la session courante
 MySession := LookupCurrentSession ;
 // on recherche la classe ds la liste des objets de l'utilisateurs
 lIndex    := MySession.UserObjects.IndexOf(ClassName) ;
 if lIndex > -1 then
  begin
   // on detruit le contexte utilisateur
   MyContext := MySession.UserObjects.Objects[lIndex] ;
   MySession.UserObjects.Delete(lIndex) ;
   FreeAndNil(MyContext) ;
  end ;
 // on recherche si on avait aussi ajouter la classe ancetre a la liste des objets utilisateur
 lIndex := MySession.UserObjects.IndexOf(ClassParent.ClassName) ;
 if lIndex > -1 then
  MySession.UserObjects.Delete(lIndex) ;
end ;
{$ENDIF}
{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 06/03/2006
Modifié le ... :   /  /
Description .. : Déplacer ici les procédures privées de la class d'édition
Mots clefs ... : Commencer à l'indice 0 (tableau passé en paramètre)
*****************************************************************}
procedure TAmortissementContext.ChargeInfoExo;
var
  ll: integer;
  i, j, k: Word;
  Q: TQuery;
begin
  FillChar(Exercices, SizeOF(Exercices), #0);
  { Chargement de la liste des exercices }
  Q := OpenSQL('SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT', TRUE, -1,
    'EXERCICE');
  if not Q.Eof then
  begin
    Q.First;
    ll := 0; // 1;  BTY 02/06 FQ 17547
    while ((not Q.Eof) and (ll <= 20)) do
    begin
      Exercices[ll].Code := Q.FindField('EX_EXERCICE').AsString;
      Exercices[ll].Deb := Q.FindField('EX_DATEDEBUT').AsDateTime;
      Exercices[ll].Fin := Q.FindField('EX_DATEFIN').AsDateTime;
      Exercices[ll].DateButoir := Q.FindField('EX_DATECUM').AsDateTime;
      Exercices[ll].DateButoirRub := Q.FindField('EX_DATECUMRUB').AsDateTime;
      Exercices[ll].DateButoirBud := Q.FindField('EX_DATECUMBUD').AsDateTime;
      Exercices[ll].DateButoirBudgete :=
        Q.FindField('EX_DATECUMBUDGET').AsDateTime;
      ImNOMBREPEREXO(Exercices[ll], i, j, k);
      Exercices[ll].NombrePeriode := k;
      Exercices[ll].EtatCpta := Q.FindField('EX_ETATCPTA').AsString;
      Inc(ll);
      Q.Next;
    end;
    Exercices[ll].Code := '';
    Exercices[ll].Deb := iDate1900;
    Exercices[ll].Fin := iDate1900;
  end
  else
  begin
    // BTY 02/06 FQ 17547
    //Exercices[1].Code := '';
    //Exercices[1].Deb := iDate1900;
    //Exercices[1].Fin := iDate1900;
    Exercices[0].Code := '';
    Exercices[0].Deb := iDate1900;
    Exercices[0].Fin := iDate1900;
  end;
  Ferme(Q);
end;

// BTY 02/06 Rattacher les paramètres de TVA au contexte d'édition
procedure TAmortissementContext.ChargeInfoTVA;
begin
 TobTva := TOB.Create('',Nil,-1);
 TobTva.LoadDetailDB('TXCPTTVA','','',Nil,True)
end;

// BTY 03/06 Rattacher les coefficients dégressifs au contexte d'édition
procedure TAmortissementContext.ChargeCoeffDegressif;
begin
 if CoeffDegressif= nil then
    begin
    CoeffDegressif := TOB.Create('',Nil,-1);
    AmChargeCoefficientDegressif ( CoeffDegressif );
    end;
end;

constructor TAmortissementContext.Create ( MySession : TISession );
begin
  inherited  Create(MySession);
  FPlanInfo := TPlanInfo.Create('');
  ChargeInfoExo;
  FPlanInfo.SetExercicesServer(Exercices);
  // BTY 02/06 Charger la TVA dans le contexte d'édition
  ChargeInfoTVA;
  // BTY 02/06 Recopier la TVA dans PlanInfo qui est visible des formules de ImEdCalc
  FPlanInfo.SetTauxTvaServer(TobTva);
  // BTY 03/06 Charger les coefficients dégressifs
  ChargeCoeffDegressif;
  // BTY 03/06 Recopier la TOB dans PlanInfo qui est visible des formules de ImEdCalc
  FPlanInfo.SetCoeffDegressifServer(CoeffDegressif);
end;

destructor TAmortissementContext.Destroy;
begin
  FPlanInfo.Free;
  // BTY 02/06
  TobTVA.Free;
  // BTY 03/06
  CoeffDegressif.Free;
  
  // MVG 05/10s/06 FQ 18920
  LookUpCurrentSession.UserCalcEdt:=SavCalcEdt;

  inherited;
end;
{$ENDIF}



////////////////////////////////////////////////////////////////////////////////
{$IFDEF SERIE1}

function getDateClotureImmo(): tDateTime ;
begin
result:=GetParamSocSecur('SO_DATECLOTUREIMMO',idate1900) ;
end ;

function getExoCloImmo(): string ;
begin
result:=GetParamSocSecur('SO_EXOCLOIMMO','') ;
end ;

procedure initialisationImEnt ;
begin
  New(VHImmo) ;
  FillChar(VHImmo^,Sizeof(VHImmo^),#0) ;

//libi MVG 12/07/2006 suite demande Y. PELUD
  {$IFNDEF EAGLSERVER}
  registerclasses ( [ TOM_IMMO]);
  registerclasses ( [ TOM_IMMOCPTE ] ) ;
  registerclasses ( [ TOF_AMUNITEOEUVRE ] ) ;
  registerclasses ( [ TOF_AMREGROUPEMENT ] ) ;
  registerclasses ( [ TOF_AMLISTE]);
  registerclasses ( [ TOF_AMCLOTURE ] ) ;
  registerclasses ( [ TOF_AMINTEG ] ) ;
  registerclasses ( [ TOT_TIMOTIFCESSION ] ) ;
  registerclasses ( [ TOF_AMCHGTMETHODE ] ) ;
  registerclasses ( [ TOF_AMDEPRECIATION ] ) ;
  registerclasses ( [ TOF_AMDEPRECGROUPE ] ) ;
  registerclasses ( [ TOF_AMSREGRO ] ) ;
  registerclasses ( [ TOF_AMMODIFSERVICE ] ) ;
  registerclasses ( [ TOF_AMANALYSE ] ) ;
  registerclasses ( [ TOF_AMEDITION ] ) ;
  registerclasses ( [ TOF_AMNDEDUCTION ] ) ;
  registerclasses ( [ TOF_AMSUIVIDPI ] ) ;
  registerclasses ( [ TOF_AMSYNTHESEDPI ] ) ;
  registerclasses ( [ TOF_AMBASEECO ] ) ;
  registerclasses ( [ TOF_AMDPIANTERIEURS ] ) ;
  registerclasses ( [ TOF_AMIMMODPI ] ) ;
  registerclasses ( [ TOF_AMPRIME ] ) ;
  registerclasses ( [ TOF_AMREDUCPRIME ] ) ;
  registerclasses ( [ TOF_AMSUBVENTION ] ) ;
  {$ENDIF EAGLSERVER}

//registerclasses ([TOF_QRIMMOAMC,TOF_QRIMMOAMP,TOF_QRIMMOAMO,TOF_QRIMMOAML,TOF_QRIMMOAMD
//                 ,TOF_AMANALYSE, TOF_AMEDITION]) ;
end ;

procedure finalisationImEnt ;
begin
  Dispose(VHImmo) ;
end ;
{$ELSE}

// BTY 09/06 FQ 16149
// Alimenter la Combo Etablissements selon les restrictions utilisateur
{---------------------------------------------------------------------------------------}
procedure ImGereEtablissement (ComboEtab : TControl);
{---------------------------------------------------------------------------------------}
begin
  if ComboEtab=nil then exit;

  if Assigned(ComboEtab) then
    begin

    {Si l'on ne gère pas les établissements ...}
    {... on affiche l'établissement par défaut}
    if not VHImmo^.EtablisCpta  then
      begin
      {if ComboEtab is ThValComboBox then
         begin
         ThValCOmboBox(ComboEtab).Text:= VHImmo^.EtablisDefaut;
         ThValCOmboBox(ComboEtab).Enabled:= False;
         end
      else
      }
      if ComboEtab is ThMultiValComboBox then
         begin
         ThMultiValCOmboBox(ComboEtab).Text:= VHImmo^.EtablisDefaut;
         ThMultiValCOmboBox(ComboEtab).Enabled:= False;
         end;
      end
      {On gère l'établisement, donc ...}
      else
      begin
      {$IFNDEF EAGLSERVER}
      {... On commence par regarder les restrictions utilisateur}
      PositionneEtabUser(ComboEtab);
      {$ENDIF}
      {... s'il n'y a pas de restrictions, on reprend le paramSoc}
      {if ComboEtab is ThValComboBox then
      begin
           if ThValCOmboBox(ComboEtab).Value = '' then
              begin
              ThValCOmboBox(ComboEtab).Value:= VHImmo^.EtablisDefaut;
              ThValCOmboBox(ComboEtab).Enabled:= True;
              end;
      end else
      }
      if ComboEtab is ThMultiValComboBox then
      begin
           if ThMultiValCOmboBox(ComboEtab).Text = '' then
              begin
              ThMultiValCOmboBox(ComboEtab).Text:= VHImmo^.EtablisDefaut;
              ThMultiValCOmboBox(ComboEtab).Enabled:= True;
              end;
      end;
      end;
    end;
end;

// On s'assure que le filtre ne va pas à l'encontre des restrictions utlisateur
// sur la Combo Etablissements, sinon appliquer les restrictions
{---------------------------------------------------------------------------------------}
procedure ImControlEtab (ComboEtab : TControl);
{---------------------------------------------------------------------------------------}
var
  Eta : string;
begin
  if ComboEtab=nil then exit;

  if Assigned(ComboEtab) then
     begin

     {S'il n'y a pas de gestion des établissement, logiquement, on ne force pas l'établissement !!!}
     if not VHImmo^.EtablisCpta then
        // appliquer l'établissement par défaut car la restauration du filtre a pu l'écraser
        //Exit;
        begin
        {if ComboEtab is ThValComboBox then
           begin
           ThValCOmboBox(ComboEtab).Text:= VHImmo^.EtablisDefaut;
           ThValCOmboBox(ComboEtab).Enabled:= False;
           end
        else
        }
        if ComboEtab is ThMultiValComboBox then
           begin
           ThMultiValCOmboBox(ComboEtab).Text:= VHImmo^.EtablisDefaut;
           ThMultiValCOmboBox(ComboEtab).Enabled:= False;
           end;
        end
        else
        begin
           // Récup restrictions utilisateur
           {$IFNDEF EAGLSERVER}
           Eta := EtabForce;
           {$ENDIF}

           {Si restriction utilisateur et différent du contenu de la combo
            => y mettre l'établissement restriction de l'utilisateur}
           {
           if ComboEtab is ThValComboBox then
              begin
              if (Eta <> '') and (Eta <> ThValCOmboBox(ComboEtab).Value) then
                 begin
                 //... on affiche l'établissement des restrictions
                 ThValCOmboBox(ComboEtab).Text := Eta;
                 //... on désactive la zone
                 ThValCOmboBox(ComboEtab).Enabled:= False;
                 end;
              end
           else
           }
           if ComboEtab is ThMultiValComboBox then
              begin
              if (Eta <> '') and (Eta <> ThMultiValCOmboBox(ComboEtab).Text) then
                 begin
                 {... on affiche l'établissement des restrictions}
                 ThMultiValCOmboBox(ComboEtab).Text := Eta;
                 {... on désactive la zone}
                 ThMultiValCOmboBox(ComboEtab).Enabled:= False;
                 end;
              end;

        end;  // EtablisCpta
     end;  // assigned
end;
//

initialization
  InitLaVariableIMO ;
//  VHImmo^.OBImmo:=TOB.Create('',Nil,-1) ;
finalization
  if VHImmo^.CoeffDegressif <> nil then VHImmo^.CoeffDegressif.Free;
  //YCP 31-07-01 if VHImmo^.OBImmo<>nil then VHImmo^.OBImmo.Free ;
  //YCP 31-07-01 if VHImmo^.PlanInfo<>nil then VHImmo^.PlanInfo.Free ;
{$ENDIF}
end.






