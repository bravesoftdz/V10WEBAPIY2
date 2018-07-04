unit AnnOutils;

// DWI 310108 Rajout d'une commande pre-*processeur "annuaire_seul"
//     qui permet d'isoler les unites pour ne gerer que les fiches "annuaire".

//mcd 12/2005 : suppression de la fct TextContatcPer ==> les contacts associés
// à une fiche annauire, sont supprimé dans TOUS les cas

//mcd 16/02/2006 diverses modif pour passer de codeper à GUID

interface

uses SysUtils, Classes, Controls, Forms, HEnt1, HCtrls, Windows,
{$IFDEF EAGLCLIENT}
  MaineAGL, eMul, eFiche, eFichList, uHttp, uHttpCs,
{$ELSE}
  Db,
{$IFNDEF DBXPRESS} dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
  Fe_main,
  Mul,
  Fiche,
  FichList,
  Sql7Util,
{$ENDIF EAGLCLIENT}
  HDB, UTOB, HQry,
{$IFDEF DP}
  EntDP,
{$ENDIF}
  hmsgbox, ParamSoc, hpanel, //LMO
{$IFDEF BUREAU}
  DpTofSaisie_Tiers,
{$ENDIF}

{$IFDEF VER150}
  Variants,
{$ENDIF}

  HStatus,
  M3FP,
  MajTable,
  Utom,
{$IFNDEF Annuaire_seul}
  TiersUtil,
  utilPGI,
{$ENDIF}
  UTof;

type ArrayOfString = array of string;

{$IFDEF BUREAU}
procedure DpCreerTiers(GuidPer: string; NatureAuxi, Nom, NomPer: string);
{$ENDIF}
{$IFNDEF Annuaire_seul}
// dwi ceic 070208
procedure SynchroniseTiers(Annuaire: Boolean; var GuidPer : string; CodeTiers : string; MajAdresseManuelle: Boolean=False);
{$ENDIF}
function DP_SupprimerEspace(Chaine: string; Mode: Integer): string;
procedure TraiteChoixPersonne(GuidPer: string; var ZoneLibelle, ZoneNomPer, ZoneTypePer: string);
procedure CreationLienPerDosVirtuel(GuidPerDos, TypeDos, CodeFnc, GuidPer: string;
  NoOrdre: Integer; Forme, CodeDos: string);
function TraduitStrCle(Val, Len: integer): string;
function SupprimeInfoDp(GuidPer: string; bForceSuppr: Boolean = False): Boolean;
function TestSupprimePersonne(GuidPer: string; EstBavard: Boolean = True): boolean;
procedure SupprimeResteAnnuaire(GuidPer: string);
procedure SupprimeEnregAnnu(GuidPer: string);
procedure AGLSupprimeListAnnu(parms: array of variant; nb: integer);
procedure AGLSupprimeEnregAnnu(parms: array of variant; nb: integer);
function CoherenceSiret(Chaine: string; Taille: Integer): Boolean;
function CoherenceFrp(Chaine: string): Boolean;

function CtrlCodeFRP(AffMsg: Boolean; CodeFRP: string): Boolean;
function CalculerCodeFRP(CodeFRP: string; var sCleCalculee_p: string): Boolean;
function CtrlEtCalculeCodeFRP(AffMsg, bCalcul_p: Boolean; CodeFRP: string; var sCleCalculee_p: string): Boolean;
function GetDosjuri(GuidPer: string): string;
function GetPerJuri(sCodeDos_p: string): string;
function DupliquerAnnuaire(sOldCle_p: string): string;
procedure SelectPersonne(var sGuidPer_p, sTannNom_p: string);
procedure FermeEtOuvre(sNature_p, sFiche_p, sRange_p, sLequel_p, sArgument_p: string);
function VerifierDoublon(TypeDoublon, ChSiren, ChCleSiret, ChPmPp, ChNom1, ChNom2, ChNomPer: string): string;
function FormateCroise(sValNom1_p, sValNom2_p, sValNomPer_p: string): string;
function TraiteChoixInterlocuteur(sGuidPer_p: string): string;
function TraiteSelectInterlocuteur(sContact_l: string): string;
function StrLAdd(sChaine_p, sCar_p: string; nNbCar_p: integer): string;
function StrRAdd(sChaine_p, sCar_p: string; nNbCar_p: integer): string;
function JaiLeDroitConceptBureau(NumTag: Integer): Boolean;
procedure GereDroitsConceptsAnnuaire(LaTom: TOM; Lecran: TForm); overload;
procedure GereDroitsConceptsAnnuaire(LaTof: TOF; Lecran: TForm); overload;

{$IFNDEF Annuaire_seul}
function AccesContact(sGuidPer_p: string; inside: ThPanel = nil): string; //LMO  //mcd 12/2005
{$ENDIF}

function JaiLeDroitAdminDossier: boolean;
function ArgumentToTabsheet(Argument: string): string; //LM20070528
function AnnuaireLanceeDepuisSynthese: Boolean;
procedure SetFicheLectureSeule(LaTom: TOM; Lecran: TForm; NoDoss: string);

const
  SUPPDEBUT = 1;
  SUPPFIN = 2;
  SUPPMIXTE = 3;
  SUPPPARTOUT = 4;

// Concepts annuaire (déplacés dans UDroitsAcces de Commun\Lib)
{const
  ccCreatAnnuaireOuLiens = 187001;
  ccSupprAnnuaireOuLiens = 187002;
  ccModifAnnuaireOuLiens = 187003;
  ccVoirLesLiens = 187004;
  ccParamListeAnnuaire = 187005;
  ccModifAnnuaireDossier = 187007; // FQ 11371 nouveau concept pour autoriser les modifs annuaire du dossier en cours
  ccAccesGlobalCompta = 187100;
  ccGroupeTravailAffect = 26052;}

type
  TSupprDoss = class(TObject)
    GuidPer: string;
    NoDossier: string;
    ForceSuppr: Boolean;
  private
    ListChpDPSTD: TStringList;
    function SuppressionBasePhysique: Boolean;
    procedure SuppressionEnregDPSTD;
  public
    function SuppressionDossier: Boolean;
    constructor Create;
    destructor Destroy; override;
  end;

type
  TPerSup = class
    GuidPer: string;
    NomPer: string;
    Bavard: Boolean;
  private
    SupprEvents: Boolean; // Evènements à supprimer ?
    constructor Create;
    function ExistePlaquettes(nomchamp, nomtable, chpGuidPer, declaration, texte: string;
      var Msg: string): Boolean;
  public
    function TestCompteTiers(var Msg: string): boolean;
    function TestJuriDosPer(var Msg: string): Boolean;
    function TestJuriLienPer(var Msg: string): Boolean;
    function TestJuriDosInfoPer(var Msg: string): boolean;
    function TestJuriGroupeSocPer(var Msg: string): boolean;
    function TestJuriEvenementPer(var Msg: string): boolean;
    procedure SupprimeEvents;
    function TestLienPer(var Msg: string): Boolean;
    function TestLienPlaq(var Msg: string): Boolean;
    function TestDossierParti(var Msg: string): Boolean;
    function TestFisPer(var Msg: string): Boolean;
    function TestEnregPer: Boolean;
    procedure SuppressionEnregPer;
  end;


////////////////// IMPLEMENTATION ///////////////////
implementation

uses
  galOutil, DpJurOutils, UtilGed, galSystem,
  UDossierSelect, uSatUtil,
{$IFDEF BUREAU}
  AfDossier_tiers,
{$ENDIF}
{$IFDEF EWS}
     // UtileWS,
{$ENDIF}
  UtilMessages,
  PwdDossier,
{$IFNDEF Annuaire_seul}
  wCommuns,
{$ENDIF}
  ReformateChamp,
  uiUtil, // $$$ JP 02/05/06
  UDroitsAcces // dans Commun\Lib (gestion en mémoire des droits d'accès)
 ,CbpMCD
 ,CbpEnumerator
  ;

// $$$ JP 02/05/06 - pour reformatage ANN_NOMPER
const
  csCarDebut_g = '- _';
  csCarMiste_g = '';
  csCarFin_g = '';
  csCarTous_g = '.';
  csLesSigles_g = 'ASS;DR;EURL;EARL;SA;SARL;SRL;SC;SCI;SCM;SNC;M;MME;MELLE;MLLE;MR;ME;MM;ETS;A;à;AU;AUX;L'';LE;LA;LES;D'';DU;DE;DES';

(* mcd 13/07/07 n'a plus d'utilité.. supprimé
// PL le 06/02/07 : KPMG ajout parser d'adresse
procedure ParseAddress (const sAdresse : String;
                        var sNumvoie : String;
                        var sBisTer : String;
                        var sTypeVoie : String;
                        var sNomvoie : String);
var

  L, TSLVoieType, TSLVoieNoCompl : TStringList ;
  w : integer ;
  inparseerror : boolean ;
  code : string;
  sCodeBisTer : String;
  sCodeTypeVoie : String;
  IndVoieNoCompl, IndVoieType : integer;

  function SupprimeVirgule(S : String) : String ;
  begin
     Result := S ;
     if copy(Result,Length(Result),1) = ',' then
        Result := copy(Result,1,Length(Result)-1);
  end;

  function IsNumber (S : String) : boolean ;
  var
    i : integer ;
  begin
    Result := False;
    if S = '' then exit;

    if inparseerror then exit;

    Result := True ;

    for i:=1 to Length(S) do
    begin
       // MB le 03/07/07 : Ajout de la virgule 12,14 --> 5 carctères OK donc "," , "-" , "/" , " "
       Result := (( Ord(S[i])>=48 ) and ( Ord(S[i])<=57 )) or ( ( (i > 1) and (i < Length(S))) and ( (Ord(S[i])=44) or (Ord(S[i])=45) or (Ord(S[i])=47) or (Ord(S[i])=32) ) ) ;
       if not Result then break ;
    end;
  end;

  function IsBisTer (S : String; var code : string) : boolean ;
  var
    iposLib, j : integer;
    chaine : string;
  begin
    Result := not inparseerror ;

    if inparseerror then Exit ;

    S := LowerCase(S) ;

    // Result := (S = 'bis') or (S = 'ter') or (S = 'quater') ;
    chaine := TSLVoieNoCompl.CommaText;
    iposLib := pos('='+S, chaine);
    code := '' ;

    if (iposLib <> 0) then
      // la chaine existe dans la liste, on retrouve le code du libellé trouvé
      begin
        for j := iposLib downto 1 do // recherche du séparateur de code précédent
          if (chaine[j] = ',') then break;

        if chaine[j] = ',' then inc(j); // on est positionné sur le code associé au libellé trouvé
        code := copy (chaine, j, iposLib-j);

        if pos('=',code) > 0 then
           code :=  copy(code,1,pos('=',code)-1);

//        if (pos('=', code) = -1) then // on a trouvé le libellé mais dans l'abrégé de la tablette
//          code := '';
      end;

    if (code <> '') then
      Result := true
    else
      Result := false;

  end;

  function IsTypeVoie (S : String; var code : string) : boolean ;
  var
    j, iposLib : integer;
    chaine : string;
  begin
    Result := not inparseerror ;

    if inparseerror then exit ;

    S := LowerCase(S) ;

    // Result := (S = 'allée') or (S = 'allee') or (S = 'avenue') or (S = 'boulevard') or (S = 'chemin') or (S = 'cours') or (S = 'mas') or (S = 'place') or (S = 'quai') or (S = 'route') or (S = 'rue') or (S = 'square') or (S = 'voie')  ;
    chaine := TSLVoieType.CommaText;
    // Recherche =?=
    iposLib := pos('='+S+'=', chaine);
    if iposLib = 0 then
    begin
       // Recherche =?,
       iposLib := pos('='+S+',', chaine);
    end;
    code := '' ;

    if (iposLib <> 0) then
      // la chaine existe dans la liste, on retrouve le code du libellé trouvé
      begin
        for j := iposLib downto 1 do // recherche du séparateur de code précédent
          if (chaine[j] = ',') then break;

        if chaine[j] = ',' then inc(j); // on est positionné sur le code associé au libellé trouvé
        code := copy (chaine, j, iposLib-j);

        if pos('=',code) > 0 then
           code :=  copy(code,1,pos('=',code)-1);

//        if (pos('=', code) = -1) then // on a trouvé le libellé mais dans l'abrégé de la tablette
//          code := '';
      end;

    if (code <> '') then
      Result := true
    else
      Result := false;

  end;

  procedure AddNomVoie(S : String; inerror : boolean);
  begin
     if sNomvoie <> '' then
        sNomvoie := sNomvoie + ' ';

     sNomvoie := sNomvoie + S;

     inparseerror := inerror;
  end;

begin
  // PL le 06/02/07 : KPMG Juste pour charger les tablettes en mémoire si elles ne le sont pas
  code := RechDom('YYVOIENOCOMPL', 'B', false);
  code := RechDom('YYVOIETYPE', 'AVE', false);
  code := 'YYVOIENOCOMPL';
  IndVoieNoCompl := TTToNum(code);
  code := 'YYVOIETYPE';
  IndVoieType := TTToNum(code);

//  V_PGI.DECombos[IndVoieNoCompl].valeurs.count
//  V_PGI.DECombos[IndVoieNoCompl].valeurs[1]
  try
    TSLVoieNoCompl := TStringList.Create;
    code := StringReplace(hTstrings(V_PGI.DECombos[IndVoieNoCompl].valeurs).text,
                          chr(VK_TAB) + chr(VK_TAB), ',', [rfReplaceAll]);
    code := StringReplace(code, chr(VK_TAB), '=', [rfReplaceAll]);
    TSLVoieNoCompl.CommaText := LowerCase(code);

    TSLVoieType := TStringList.Create;
    code := StringReplace(hTstrings(V_PGI.DECombos[IndVoieType].valeurs).text,
                          chr(VK_TAB) + chr(VK_TAB), ',', [rfReplaceAll]);
    code := StringReplace(code, chr(VK_TAB), '=', [rfReplaceAll]);
    TSLVoieType.CommaText := LowerCase(code);

  //TSLVoieType
  //  code := hTstrings(V_PGI.DECombos[IndVoieNoCompl].valeurs).commatext;

  //  V_PGI.DECombos[IndVoieType].valeurs.count

    inparseerror := false ;
    L := TStringList.Create ;
    L.Delimiter := ' ' ;
    // MB 05/07/07
    L.DelimitedText := sAdresse;
    //StringReplace(sAdresse, ',', ' ', [rfReplaceAll])

    // Commence par mettre a vide.
    sNumvoie := '' ;
    sBisTer := '' ;
    sTypeVoie := '' ;
    sNomvoie := '' ;

    for w:=0 to L.Count -1 do
    begin
       // 0 : Numéro de Voie ou Type Voie
       // 1 : bis/ter ou Type voie
       // 2 : Type voie ou nomvoie
       // 3 : a partir de la nomvoie
       case w of
       0:
       begin
          if IsNumber(SupprimeVirgule(L.Strings[w])) then
             sNumvoie := L.Strings[w]
          else
             if IsTypeVoie(L.Strings[w], sCodeTypeVoie) then
                sTypeVoie := sCodeTypeVoie
             else
                AddNomVoie(L.Strings[w], true) ;
       end;

       1:
       begin
          if IsNumber(SupprimeVirgule(L.Strings[w])) then
             sNumvoie := sNumvoie +' '+L.Strings[w]
          else if IsBisTer(L.Strings[w], sCodeBisTer) then
             sBisTer := sCodeBisTer
          else if IsTypeVoie(L.Strings[w], sCodeTypeVoie) then
             sTypeVoie := sCodeTypeVoie
          else
             AddNomVoie(L.Strings[w], true) ;
       end;

       2:
       begin
          if IsBisTer(L.Strings[w], sCodeBisTer) then
             sBisTer := sCodeBisTer
          else if IsTypeVoie(L.Strings[w], sCodeTypeVoie) then
             sTypeVoie := sCodeTypeVoie
          else
             AddNomVoie(L.Strings[w], inparseerror) ;
       end;

       3:
       begin
          if IsTypeVoie(L.Strings[w], sCodeTypeVoie) then
             sTypeVoie := sCodeTypeVoie
          else
             AddNomVoie(L.Strings[w], inparseerror) ;
       end;
       else
           AddNomVoie(L.Strings[w], inparseerror) ;
       end;

    end;


  finally
    TSLVoieNoCompl.Free;
    TSLVoieType.Free;
    L.Free ;
  end;

end;
 *)

{$IFNDEF Annuaire_seul}
// dwi ceic 070208
procedure SynchroniseTiers (Annuaire:boolean; var GuidPer:String; CodeTiers:string; MajAdresseManuelle: Boolean=False);
var
  FTob, UneTobFiscal: TOB;
  Q1, Q2, QT: TQuery;
  T_AUXILIAIRE, Auxi, Form: string;
  strDossier: string;
  iYear, iMonth, iDay: word;
  strNomPer: string; // $$$ JP 02/05/06
   // sNumvoie, sBisTer, sTypeVoie, sNomvoie : string;
begin
  // $$$ JP 21/07/06: idem que pour SynchroniseParamSoc & MajContact: que si on est bien connecté à la base commune
  if (V_PGI.DefaultSectionName <> '') and (V_PGI.DbName <> V_PGI.DefaultSectionDbName) then
    exit;

  // MD 25/09/00 - exit déplacé, select * inutile, test codetiers
  if GuidPer = '' then
  begin
    if CodeTiers = '' then exit;
    Q1 := OpenSQL('select ANN_GUIDPER from ANNUAIRE where ANN_TIERS = '
      + '"' + CodeTiers + '"', TRUE);
    if not Q1.eof then GuidPer := Q1.FindField('ANN_GuidPER').AsString;
    Ferme(Q1);
    exit;
  end;

  // Annuaire vers Tiers
  if Annuaire = TRUE then
  begin
    Q1 := OpenSQL('select * from TIERS where T_TIERS ="' + CodeTiers + '"', TRUE);
    // MD 19/06/02 - Ne pas créer le Tiers sinon sera incomplet
    if Q1.eof then begin Ferme(Q1); exit; end; // SORTIE
    Auxi := Q1.FindField('T_AUXILIAIRE').AsString;
    T_AUXILIAIRE := '"' + Auxi + '"';

    FTob := TOB.Create('TIERS', nil, -1);
    FTob.SelectDB(T_AUXILIAIRE, Q1, TRUE);
    Ferme(Q1);

    Q1 := OpenSQL('select * from ANNUAIRE where ANN_GUIDPER ="' + GuidPer + '"', TRUE);
    if not Q1.eof then
    begin
      FTob.PutValue('T_NATIONALITE', Q1.FindField('ANN_NATIONALITE').Asstring);
      FTob.PutValue('T_LANGUE', Q1.FindField('ANN_LANGUE').Asstring);

      // MD 27/12/02
      // MD 13/08/04 test InBaseCommune était erroné
      if ((V_PGI.DefaultSectionName = '') or (V_PGI.DefaultSectionDbName = V_PGI.DbName)) and
        (FTob.GetValue('T_DEVISE') <> Q1.FindField('ANN_DEVISE').AsString) then
      begin
           // $$$ JP 12/08/2004 - gestion si on est pas dans le bureau
        if V_PGI.RunFromLanceur then
          strDossier := V_PGI.NoDossier
        else
          strDossier := VH_DOSS.NoDossier;
        if ExisteSQL('SELECT DAP_NOMEXEC FROM DOSSAPPLI WHERE DAP_NOMEXEC="CGIS5.EXE" AND DAP_NODOSSIER="' + strDossier + '"') = TRUE then
        begin
          if PGIAsk(' Attention, la fiche tiers ' + CodeTiers + ' va être actualisée à partir de la fiche annuaire,' + #13#10
            + ' ce qui va entrainer une modification de la devise en ' + Q1.FindField('ANN_DEVISE').AsString + '.' + #13#10
            + ' Vous n''allez plus être cohérent avec ce que vous aviez antérieurement dans la GI (' + FTob.GetValue('T_DEVISE') + ').' + #13#10
            + ' Voulez-vous quand même mettre à jour la fiche Tiers ?', TitreHalley) = mrNo then
          begin
            Ferme(Q1);
            FTob.Free;
            exit; // SORTIE
          end;
        end;
      end;
        //mcd 15/10/2008 GIGA 15049 il faut faire le lien, car plus modifiable en GI
      if (Q1.FindField('ANN_PPPM').Asstring = 'PP') then  FTob.PutValue('T_PARTICULIER','X')
       else  FTob.PutValue('T_PARTICULIER','-');
      FTob.PutValue('T_DEVISE', Q1.FindField('ANN_DEVISE').Asstring);
      FTob.PutValue('T_FAX', Q1.FindField('ANN_FAX').Asstring);
      if Q1.FindField('ANN_MOISCLOTURE').Asstring <> '' then
        FTob.PutValue('T_MOISCLOTURE', StrToInt(Q1.FindField('ANN_MOISCLOTURE').Asstring));
      FTob.PutValue('T_ABREGE', copy(Q1.FindField('ANN_NOMPER').Asstring, 1, 17));

      // ME 22/11/00 harmonisation forme juridique
      Q2 := OpenSQL('select JFJ_CODEDP from JUFORMEJUR where JFJ_FORME = '
        + '"' + Q1.FindField('ANN_FORME').Asstring + '"', TRUE);
      if not Q2.eof then
        FTob.PutValue('T_FORMEJURIDIQUE', Q2.FindField('JFJ_CODEDP').Asstring);
        // FTob.PutValue('T_JURIDIQUE', Q2.FindField('JFJ_CODEDP').Asstring);
      Ferme(Q2);

      //--- CAT : Alimentation de T_NIF par DFI_NOINTRACOMM
      Q2 := OpenSQL('select DFI_NOINTRACOMM from DPFISCAL where DFI_GUIDPER="' + GuidPer + '"', TRUE);
      if not Q2.eof then
      begin
        if (Q2.FindField('DFI_NOINTRACOMM').Asstring <> '') then
          FTob.PutValue('T_NIF', Copy(Q2.FindField('DFI_NOINTRACOMM').Asstring, 1, 17));
      end;
      Ferme(Q2);

      FTob.PutValue('T_RVA', Q1.FindField('ANN_SITEWEB').Asstring);
      FTob.PutValue('T_TELEPHONE', Q1.FindField('ANN_TEL1').Asstring);
     //mcd 16/10/2008 GIGA 15714 FTob.PutValue('T_TELEX', Q1.FindField('ANN_MINITEL').Asstring);
      FTob.PutValue('T_TELEX', Q1.FindField('ANN_TEL2').Asstring); //mcd 16/10/2008 15714
      FTob.PutValue('T_APE', Q1.FindField('ANN_CODENAF').Asstring);
      FTob.PutValue('T_SIRET', Q1.FindField('ANN_SIREN').Asstring + Q1.FindField('ANN_CLESIRET').Asstring);
      FTob.PutValue('T_ADRESSE1', Q1.FindField('ANN_ALRUE1').Asstring);
      FTob.PutValue('T_ADRESSE2', Q1.FindField('ANN_ALRUE2').Asstring);
      FTob.PutValue('T_ADRESSE3', Q1.FindField('ANN_ALRUE3').Asstring);
      FTob.PutValue('T_CODEPOSTAL', Q1.FindField('ANN_ALCP').Asstring);
      FTob.PutValue('T_VILLE', Q1.FindField('ANN_ALVILLE').Asstring);
      FTob.PutValue('T_PAYS', Q1.FindField('ANN_PAYS').Asstring);
      FTob.PutValue('T_DATEMODIF', Q1.FindField('ANN_DATEMODIF').AsDateTime);
      FTob.PutValue('T_LIBELLE', copy(Q1.FindField('ANN_NOM1').Asstring, 1, 35));
      FTob.PutValue('T_PRENOM', copy(Q1.FindField('ANN_NOM2').Asstring, 1, 35));
      FTob.PutValue('T_TIERS', CodeTiers);
      //mcd 02/03/2005 ajout sexe et date naissance
      FTob.PutValue('T_ENSEIGNE', Q1.FindField('ANN_ENSEIGNE').Asstring);
      FTob.PutValue('T_SEXE', Q1.FindField('ANN_SEXE').Asstring);
      DecodeDate(Q1.FindField('ANN_DATENAIS').AsDateTime, iYear, iMonth, iDay);
      FTob.PutValue('T_ANNEENAISSANCE', iYear);
      FTob.PutValue('T_MOISNAISSANCE', iMonth);
      FTob.PutValue('T_JOURNAISSANCE', iDay);
      FTob.PutValue('T_EMAIL', Q1.FindField('ANN_EMAIL').Asstring);
      FTob.PutValue('T_RVA', Q1.FindField('ANN_SITEWEB').Asstring); //mcd 24/05/05
      FTob.InsertOrUpdateDB(true);
      // MB : Ajout de la gestion du doublon
      // Obligation de synchroniser ANNUBIS et TIERSCOMPL
      freeandnil(FTob);
      FTob := TOB.Create('TIERSCOMPL', nil, -1);
      FTob.SelectDB(T_AUXILIAIRE, nil, TRUE);
      QT := OpenSQL('select ANB_DOUBLON from ANNUBIS where ANB_GUIDPER = "' + GuidPer + '"', TRUE);
      if not QT.eof then
      begin
        if FTob.GetValue('YTC_TIERS')='' then FTob.PutValue('YTC_TIERS', CodeTiers);
        FTob.PutValue('YTC_AUXILIAIRE', Auxi);
        FTob.PutValue('YTC_DOUBLON', QT.FindField('ANB_DOUBLON').Asstring);
        FTob.InsertOrUpdateDB(true);
      end;
      freeandnil(FTob);
      Ferme(QT);
    end;
  end
  else
    // Tiers vers Annuaire
  begin
    FTob := TOB.Create('ANNUAIRE', nil, -1);
    FTob.SelectDB('"' + GuidPer + '"', nil, TRUE);
    Q1 := OpenSQL('select * from TIERS where T_TIERS = "' + CodeTiers + '"', TRUE);
    if not Q1.eof then
    begin
      FTob.PutValue('ANN_NATUREAUXI', Q1.FindField('T_NATUREAUXI').Asstring); //mcd 27/05/05 il faut le lien avec la nature associé
      FTob.PutValue('ANN_NATIONALITE', Q1.FindField('T_NATIONALITE').Asstring);
      FTob.PutValue('ANN_LANGUE', Q1.FindField('T_LANGUE').Asstring);
      FTob.PutValue('ANN_DEVISE', Q1.FindField('T_DEVISE').Asstring);
      if Q1.FindField('T_MOISCLOTURE').AsInteger <> 0 then
        FTob.PutValue('ANN_MOISCLOTURE', Format('%2.2d', [Q1.FindField('T_MOISCLOTURE').Asinteger]));

      // ME 22/11/00 harmonisation forme juridique
      Q2 := OpenSQL('select JFJ_FORME,JFJ_FORMEGEN  from JUFORMEJUR where JFJ_CODEDP = '
        + '"' + Q1.FindField('T_FORMEJURIDIQUE').Asstring + '"', TRUE);
      //mcd 25/04/01   + '"'+Q1.FindField('T_JURIDIQUE').Asstring+'"', TRUE);
      if not Q2.eof then
      begin
        Form := Q2.FindField('JFJ_FORME').Asstring;
        FTob.PutValue('ANN_FORME', Form);
        if Q2.FindField('JFJ_FORMEGEN').Asstring <> '' then
          FTob.PutValue('ANN_FORMEGEN', Q2.FindField('JFJ_FORMEGEN').Asstring);
      end;
      Ferme(Q2);

      //--- CAT : Alimentation de DFI_NOINTRACOMM par T_NIF
      UneTobFiscal := TOB.Create('DPFISCAL', nil, -1);
      if UneTobFiscal.SelectDB('"' + GuidPer + '"', nil, TRUE) then
      begin
        if (Q1.FindField('T_NIF').Asstring <> '') then
        begin
          UneTobFiscal.PutValue('DFI_NOINTRACOMM', Q1.FindField('T_NIF').Asstring);
          UneTobFiscal.UpdateDB;
        end;
      end;
      UneTobFiscal.Free;

      FTob.PutValue('ANN_TEL2', Q1.FindField('T_TELEX').Asstring); //mcd 16/10/2008 giga 15714
      FTob.PutValue('ANN_TEL1', Q1.FindField('T_TELEPHONE').Asstring);
//mcd 23/08/07 pas fait comem doc 14403     if Q1.FindField('T_PARTICULIER').AsString='X' then
      if Q1.FindField('T_PARTICULIER').AsString = '-' then
      begin
        //mcd 16/10/2008 giga 15714 FTob.PutValue('ANN_MINITEL', Q1.FindField('T_TELEX').Asstring);
        FTob.PutValue('ANN_FAX', Q1.FindField('T_FAX').Asstring);
        if Ftob.getValue('ANN_PPPM') = '' then
          FTob.PutValue('ANN_PPPM', 'PM');//mcd 15/10/2008 15049
      end
      else
        if Ftob.getValue('ANN_PPPM') = '' then
          FTob.PutValue('ANN_PPPM', 'PP');//mcd 15/10/2008 15049
      FTob.PutValue('ANN_SITEWEB', copy(Q1.FindField('T_RVA').Asstring, 1, 35)); //mcd 24/05/2005 fait dans tous les cas (et pas si particulier uniquement)
      FTob.PutValue('ANN_EMAIL', Q1.FindField('T_EMAIL').Asstring); //mcd 24/05/2005
      //pour JB       FTob.PutValue('ANN_NOMPER', Q1.FindField('T_ABREGE').Asstring);
      if wGetTypeField('ANN_CODENAF')='VARCHAR(4)' then
        FTob.PutValue('ANN_CODENAF', copy(Q1.FindField('T_APE').Asstring, 1, 4))
      else
        FTob.PutValue('ANN_CODENAF', Q1.FindField('T_APE').Asstring);
      // MD 16/02/01
      FTob.PutValue('ANN_SIREN', copy(Q1.FindField('T_SIRET').Asstring, 1, 9));
      FTob.PutValue('ANN_CLESIRET', copy(Q1.FindField('T_SIRET').Asstring, 10, 5));
      //FTob.PutValue('ANN_SIREN', Q1.FindField('T_SIRET').Asstring);
      if Not MajAdresseManuelle then //mcd 24/07/07 on ne le fait pas si fiche gep et annuaire prioritaire
      begin
        FTob.PutValue('ANN_ALRUE1', Q1.FindField('T_ADRESSE1').Asstring);
        FTob.PutValue('ANN_ALRUE2', Q1.FindField('T_ADRESSE2').Asstring);
        FTob.PutValue('ANN_ALRUE3', Q1.FindField('T_ADRESSE3').Asstring);
        FTob.PutValue('ANN_ALCP', Q1.FindField('T_CODEPOSTAL').Asstring);
        FTob.PutValue('ANN_ALVILLE', Q1.FindField('T_VILLE').Asstring);
      end;
      FTob.PutValue('ANN_PAYS', Q1.FindField('T_PAYS').Asstring);
      FTob.PutValue('ANN_DATEMODIF', Q1.FindField('T_DATEMODIF').AsDateTime);
      //mcd 17/05/05 champ supprimé FTob.PutValue('ANN_AUXILIAIRE', Q1.FindField('T_AUXILIAIRE').Asstring);


    (* mcd 13/07/07 n'a plus d'utilité avec fiche pero.. supprimé
      // PL le 06/02/07 : KPMG on met à jour intelligemment les champs voie, type de voie, nom de voie, etc...
      ParseAddress (Q1.FindField('T_ADRESSE1').Asstring, sNumvoie, sBisTer, sTypeVoie, sNomvoie); *)

      // MD le 17/10/07 : il fallait aussi enlever ceci qui vidait le détail des rues dans l'annuaire !!
    { FTob.PutValue('ANN_VOIENO', sNumvoie);
      FTob.PutValue('ANN_VOIENOCOMPL', sBisTer);
      FTob.PutValue('ANN_VOIETYPE', sTypeVoie);
      FTob.PutValue('ANN_VOIENOM', sNomvoie); }

      FTob.PutValue('ANN_GUIDPER', GuidPer); // $$$ JP 21/04/06 - GuidPER modifié en GUIDPER (majuscule)
      FTob.PutValue('ANN_CODEPER', -2); // $$$ JP 21/04/06 - -2 nécessaire depuis qu'on a ANN_GUIDPER

      FTob.PutValue('ANN_NOM1', Q1.FindField('T_LIBELLE').Asstring);
      FTob.PutValue('ANN_NOM2', Q1.FindField('T_PRENOM').Asstring);
      FTob.PutValue('ANN_TIERS', CodeTiers);
      //mcd 02/03/2005 ajout sexe et date naissance
      FTob.PutValue('ANN_ENSEIGNE', Q1.FindField('T_ENSEIGNE').Asstring);
      if Q1.FindField('T_PARTICULIER').Asstring = 'X' then
      begin
        FTob.PutValue('ANN_SEXE', Q1.FindField('T_SEXE').Asstring);
        if (Q1.FindField('T_ANNEENAISSANCE').AsInteger > 1900) and (Q1.FindField('T_MOISNAISSANCE').AsInteger > 0) and (Q1.FindField('T_JOURNAISSANCE').AsInteger > 0) then
          FTob.PutValue('ANN_DATENAIS', EncodeDate(Q1.FindField('T_ANNEENAISSANCE').AsInteger, Q1.FindField('T_MOISNAISSANCE').AsInteger, Q1.FindField('T_JOURNAISSANCE').AsInteger));
      end;

      // $$$ JP 02/05/06 - il faut renseigner le nom de recherche, en filtrant les caractères indésirables (comme dans le dédoublonnage)
      if Trim(FTob.GetString('ANN_NOMPER')) = '' then
      begin
        InitFormats(csCarDebut_g, csCarMiste_g, csCarFin_g, csCarTous_g, csLesSigles_g);
        strNomPer := Trim(Q1.FindField('T_LIBELLE').AsString + ' ' + Q1.FindField('T_PRENOM').Asstring);
        strNomPer := ElimineCaracZarbi(strNomPer, TRUE);
        strNomPer := Copy(ReformateChaine(strNomPer), 1, 17);
        FTob.PutValue('ANN_NOMPER', strNomPer); // 17 car., si on synchronise plus tard avec T_ABREGE...
        RazFormats;
      end;

      FTob.InsertOrUpdateDB(true);
      freeandnil(FTob);
      // MB : Ajout de la gestion du doublon
      // Obligation de synchroniser ANNUBIS et TIERSCOMPL
      FTob := TOB.Create('ANNUBIS', nil, -1);
      FTob.SelectDB('"' + GuidPer + '"', nil, TRUE);
      QT := OpenSQL('select YTC_DOUBLON from TIERSCOMPL where YTC_AUXILIAIRE = "' + Q1.FindField('T_AUXILIAIRE').Asstring + '"', TRUE);
      if not QT.eof then
      begin
        FTob.PutValue('ANB_GUIDPER', GuidPer);
        FTob.PutValue('ANB_DOUBLON', QT.FindField('YTC_DOUBLON').Asstring);
        FTob.InsertOrUpdateDB(true);
      end;
      freeandnil(FTob);
      Ferme(QT);
    end;
  end;
  Ferme(Q1);
  if Assigned(Ftob) then
    Ftob.Free;
end;
{$ENDIF}


{----------------------------------------------------------------
 --- Nom       : FONCTION DP_SupprimerEspace
 --- Objet     : Supprime les espaces en fin de chaine.
 --- Paramètre : Chaine : Contient la chaine à traiter
 --- Retour    : Une chaine de caractere sans espace à la fin.
 ----------------------------------------------------------------}

function DP_SupprimerEspace(Chaine: string; Mode: Integer): string;
var Indice, Longueur: Integer;
  ChResultat: string;
  St: string;
begin
  longueur := Length(Chaine);
  case Mode of
    SUPPDEBUT: begin
        indice := 1;
        repeat
          if (Chaine[Indice] <> ' ') then
            break;
          inc(Indice);
        until (Indice = Longueur + 1);
        ChResultat := Copy(Chaine, Indice, Longueur + 1);
      end;
    SUPPFIN: begin
        indice := Longueur;
        repeat
          if (Chaine[Indice] <> ' ') then
            break;
          dec(Indice);
        until (Indice = 0);
        ChResultat := Copy(Chaine, 0, Indice);
      end;
    SUPPMIXTE: begin
        indice := 1;
        repeat
          if (Chaine[Indice] <> ' ') then
            break;
          inc(Indice);
        until (Indice = Longueur + 1);
        Chaine := Copy(Chaine, Indice, Longueur + 1);
        Longueur := Length(Chaine);
        indice := longueur;
        repeat
          if (Chaine[Indice] <> ' ') then
            break;
          dec(Indice);
        until (Indice = 0);
        ChResultat := Copy(Chaine, 0, Indice);
      end;
    SUPPPARTOUT:
      begin
        indice := 1;
        repeat
          if (Chaine[Indice] <> ' ') then
            St := St + Chaine[Indice];
          inc(Indice);
        until (Indice = Longueur + 1);
        ChResultat := St;
      end;
  end;
  DP_SupprimerEspace := ChResultat;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/01/2000
Modifié le ... : 07/11/2005
Description .. : Détermine les libellés/nomper/typeper à partir du guidper
Mots clefs ... :
*****************************************************************}

procedure TraiteChoixPersonne(GuidPer: string; var ZoneLibelle, ZoneNomPer, ZoneTypePer: string);
var
  Q: TQuery;
  sNom1_l, sNom2_l, sNom3_l: string;
begin
  ZoneLibelle := '';
  ZoneNomPer := '';
  if (GuidPer <> '') then
  begin
    Q := OpenSQL('select ANN_NOMPER,ANN_NOM1,ANN_NOM2,ANN_NOM3,ANN_TYPEPER from ANNUAIRE where ANN_GUIDPER="' + GuidPer + '"', true);
    if not Q.EOF then
    begin
      sNom1_l := Q.FindField('ANN_NOM1').AsString;
      sNom2_l := Q.FindField('ANN_NOM2').AsString;
      sNom3_l := Q.FindField('ANN_NOM3').AsString;
      if sNom1_l <> '' then
        ZoneLibelle := sNom1_l;

      if sNom2_l <> '' then
      begin
        if ZoneLibelle <> '' then
          ZoneLibelle := ZoneLibelle + ' ';
        ZoneLibelle := ZoneLibelle + sNom2_l;
      end;

      if sNom3_l <> '' then
      begin
        if ZoneLibelle <> '' then
          ZoneLibelle := ZoneLibelle + ' ';
        ZoneLibelle := ZoneLibelle + sNom3_l;
      end;

      ZoneNomPer := Q.FindField('ANN_NOMPER').AsString;
      ZoneTypePer := Q.FindField('ANN_TYPEPER').AsString;
    end;
    Ferme(Q);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function TraiteChoixInterlocuteur(sGuidPer_p: string): string;
var
  QRYReq_l: TQuery;
begin
  Result := '';
    (* //mcd 12/2005 QRYReq_l := OpenSQL('select ANI_NOM, ANI_PRENOM ' +
                       'from ANNUINTERLOC ' +
                       'where ANI_GUIDPER = ' + sGuidPer_p +
                       '  and ANI_PRINCIPAL = "X"', TRUE); *)
  QRYReq_l := OpenSQL('select C_NOM, C_PRENOM ' +
    'from CONTACT ' +
    'where C_GUIDPER = "' + sGuidPer_p + '"' +
    '  and C_PRINCIPAL = "X"', TRUE);
  if not QRYReq_l.Eof then
    Result := QRYReq_l.Fields[0].AsString + ' ' + QRYReq_l.Fields[1].AsString;
  Ferme(QRYReq_l);
end;


function TraiteSelectInterlocuteur(sContact_l: string): string;
var
  sType, sAux, sNum: string;
  QRYReq_l: TQuery;
begin
  Result := '';
  sType := ReadTokenSt(sContact_l);
  sAux := ReadTokenSt(sContact_l);
  sNum := ReadTokenSt(sContact_l);

  if (sType <> '') and (sAux <> '') and (sNum <> '') then
  begin
    QRYReq_l := OpenSQL('select C_NOM, C_PRENOM ' +
      'from CONTACT ' +
      'where C_TYPECONTACT = "' + sType + '"' +
      ' and C_AUXILIAIRE = "' + sAux + '"' +
      ' and C_NUMEROCONTACT = ' + sNum, TRUE);

    if not QRYReq_l.Eof then
      Result := QRYReq_l.Fields[0].AsString + ' ' + QRYReq_l.Fields[1].AsString;

    Ferme(QRYReq_l);
  end;
end;


{$IFNDEF Annuaire_seul}
// dwi ceic 310108

function AccesContact(sGuidPer_p: string; inside: ThPanel = nil): string; //LMO
var
  sttmp, stTmp1, stTmp2, Tiers, Nat, Titre, Auxi: string;
  QQ: Tquery;
begin
  //mcd  12/2005 fait une fct commune qui recoit guidper
  // et appele la saisie contact
  //car utiliser dans plein d'endoit ou on a le bouton contact
  QQ := OpenSQL('SELECT ANN_TIERS,ANN_NATUREAUXI,ANN_NOM1 FROM ANNUAIRE WHERE ANN_GUIDPER="' + sGuidPer_p + '"', True);
  Tiers := '';
  Nat := '';
  Result := '';
  if not QQ.Eof then
  begin
    tiers := QQ.FindField('ANN_TIERS').AsString;
    Nat := QQ.FindField('ANN_NATUREAUXI').AsString;
    Titre := QQ.FindField('ANN_NOM1').AsString;
  end;
  ferme(QQ);
  stTmp2 := ';TITRE=' + titre + ';ACTION=MODIFICATION';
  if tiers <> '' then
  begin
    Auxi := TiersAuxiliaire(tiers, False); //mcd 16/06/06
    if (Auxi <> '')
      then
    begin
      stTmp := 'TYPE=T;TIERS=' + tiers + ';TYPE2=' + Nat;
      stTmp1 := 'T;' + Auxi;
    end
{$IFDEF DP}
    else
    // MBrun 20/12/2007 - KPMG
    // si le tiers est renseigné mais qu'on ne trouve pas l'auxiliaire (table tiers vide),
    // il faut quand même passer une clé auxiliaire (= au code tiers chez kpmg)
      if VH_DP.SeriaKPMG then
      begin
        stTmp := 'TYPE=T;TIERS=' + tiers + ';TYPE2=' + Nat;
        stTmp1 := 'T;' + tiers;
      end
{$ENDIF}
      else
      begin
        stTmp := 'TYPE=ANN';
        stTmp1 := 'ANN;' + sGuidPer_p; //mcd 16/06/06 pour ne pas avoir tous les tiers,
      end; //si onpassse T + blanc. essaie avec ann et guidper, mais il y a de grande chance que cela n'existe pas
  end
  else
  begin
    stTmp := 'TYPE=ANN';
    stTmp1 := 'ANN;' + sGuidPer_p;
  end;
  Result := AglLanceFicheInside('YY', 'YYCONTACT', stTmp1, '', stTmp + ';FROMSAISIE;GUIDPER=' + sGuidPer_p + ';ALLCONTACT' + stTmp2, Inside); //LMO
end;

{$ENDIF}



function DPRecupChaine(var Chaine: string): string;
var Longueur, indice: Integer;
  ChResultat, St: string;
begin
  longueur := Length(Chaine);
  for indice := 1 to longueur do
    if (Chaine[Indice] = ' ') then break;
  ChResultat := copy(Chaine, 1, indice);
  St := copy(Chaine, indice + 1, longueur + 1 - indice);
  Chaine := St;
  DPRecupChaine := ChResultat;
end;

procedure CreationLienPerDosVirtuel(GuidPerDos, TypeDos, CodeFnc, GuidPer: string;
  NoOrdre: Integer; Forme, CodeDos: string);
var
  sWhere, sSelect: string;
  Q: TQuery;
  T: TOB;
begin
  sSelect := 'select * from ANNULIEN ';
  sWhere := 'where (ANL_GUIDPERDOS="' + GuidPerDos + '" ' +
    'AND ANL_NOORDRE=' + InttoStr(NoOrdre) + ' ' +
    'AND ANL_TYPEDOS="' + TypeDos + '" ' +
    'AND ANL_FONCTION="' + CodeFnc + '" ' +
    'AND ANL_GUIDPER="' + GuidPer + '")';
  if not ExisteSQL(sSelect + sWhere) then
  begin
    T := TOB.Create('ANNULIEN', nil, -1);
    T.InitValeurs;
    T.PutValue('ANL_GUIDPERDOS', GuidPerDos);
    T.PutValue('ANL_NOORDRE', NoOrdre);
    T.PutValue('ANL_CODEDOS', CodeDos);
    T.PutValue('ANL_GUIDPER', GuidPer);
    T.PutValue('ANL_FONCTION', CodeFnc);
    T.PutValue('ANL_AFFICHE', RechDom('JUTYPEFONCTAFF', CodeFnc, FALSE));
    T.PutValue('ANL_NOMPER', RechDom('ANNUAIRE', GuidPer, FALSE));
    // valeurs par défaut de JUFONCTION
    Q := OpenSQL('select JFT_DEFAUT, JFT_RACINE, JFT_TYPEDOS, JFT_TRI, JFT_TIERS'
      + ' from JUFONCTION where JFT_FONCTION="' + CodeFnc + '"'
      + ' and JFT_FORME="' + Forme + '"', True);
    if Q.Eof then // ne devrait jamais se produire
      T.PutValue('ANL_TYPEDOS', TypeDos)
    else
    begin
      if (Q.FindField('JFT_DEFAUT').AsString = 'X') and (CodeFnc <> 'STE') then
        T.PutValue('ANL_FORME', Forme);
      T.PutValue('ANL_RACINE', Q.FindField('JFT_RACINE').AsString);
      T.PutValue('ANL_TYPEDOS', Q.FindField('JFT_TYPEDOS').AsString);
      T.PutValue('ANL_TRI', Q.FindField('JFT_TRI').AsInteger);
      T.PutValue('ANL_TIERS', Q.FindField('JFT_TIERS').AsString);
    end;
    Ferme(Q);
    // maj
    T.InsertDB(nil);
    T.Free;
  end;
end;

function TraduitStrCle(Val, Len: integer): string;
var
  Ret: string;
begin
  Ret := IntToStr(Val);
  while Length(Ret) < Len do Ret := '0' + Ret;
  result := Ret;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 13/06/2003
Modifié le ... :   /  /
Description .. : Ajoute à gauche de chaînesChaine_p, la caractère sCar_p
Suite ........ : jusqu'à ce que la longueur de la chaîne soit de nNbCar_p
Mots clefs ... :
*****************************************************************}

function StrLAdd(sChaine_p, sCar_p: string; nNbCar_p: integer): string;
begin
  while Length(sChaine_p) < nNbCar_p do
    sChaine_p := sCar_p + sChaine_p;
  result := sChaine_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 02/07/2003
Modifié le ... :   /  /
Description .. : Ajoute à droite de chaînesChaine_p, la caractère sCar_p
Suite ........ : jusqu'à ce que la longueur de la chaîne soit de nNbCar_p
Mots clefs ... :
*****************************************************************}

function StrRAdd(sChaine_p, sCar_p: string; nNbCar_p: integer): string;
begin
  while Length(sChaine_p) < nNbCar_p do
    sChaine_p := sChaine_p + sCar_p;
  result := sChaine_p;
end;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 28/11/2003
Modifié le ... :   /  /
Description .. : Examine les droits sur le n° de tag
Suite ........ : de la branche 187 (Concepts bureau).
Suite ........ : Le test se fait dans la base commune.
Suite ........ : Retourne True si droit ok, False sinon.
Mots clefs ... : CONCEPT;BUREAU;DROITS
*****************************************************************}

function JaiLeDroitConceptBureau(NumTag: Integer): Boolean;
begin
  Result := JaiLeDroitAccesDB0(NumTag);
end;

procedure GereDroitsConceptsAnnuaire(LaTom: TOM; Lecran: TForm);
// Gère les boutons et le mode d'ouverture de la tom
// en fonction des droits sur les concepts annuaire
var cstDroitModif: Integer;
begin
  // FQ 11371 : la fiche annuaire du dossier présente des droits de modification spécifiques
  if AnnuaireLanceeDepuisSynthese then
  begin
    cstDroitModif := ccModifAnnuaireDossier;
    {+GHA - 12/2007}
    if (TForm(globalfichePile[1]).Name = 'ANNUAIRE') and (globalFichePile.Count = 2) then
    begin
      LaTom.setControlVisible('bInsert', FALSE);
      LaTom.setControlVisible('bDelete', FALSE);
    end;
    {-GHA - 12/2007}
  end
  else
    cstDroitModif := ccModifAnnuaireOuLiens;

  { FQ 11763 (KPMG) : l'absence de droit modif n'enlève plus le droit d'insert
   donc mise en commentaire =>
  // Pas le droit de modifier
  if Not JaiLeDroitConceptBureau(cstDroitModif) then
    begin
    // pas de création/suppression
    LaTom.SetControlVisible('BINSERT', False);
    LaTom.SetControlVisible('BDELETE', False);
    LaTom.SetControlVisible('BTLIENANNU', False);
    // passe en lecture seule (voir TraiteArguments dans UtilTom pour les autres fiches)
    if Lecran is TFFiche then TFFiche(Lecran).FTypeAction:=taConsult;
    if Lecran is TFFicheListe then TFFicheListe(Lecran).FTypeAction:=taConsult ;
    end

  // ou bien droits plus détaillés
  else  }

  if ((AnnuaireLanceeDepuisSynthese) and (not JaiLeDroitConceptBureau(ccModifAnnuaireDossier))) or
    (not AnnuaireLanceeDepuisSynthese) then
  begin
    // Pas de création
    if not JaiLeDroitConceptBureau(ccCreatAnnuaireOuLiens) then
    begin
      LaTom.SetControlVisible('BINSERT', False);
      LaTom.SetControlVisible('BTLIENANNU', False);
    end;
    // Pas de suppression
    if not JaiLeDroitConceptBureau(ccSupprAnnuaireOuLiens) then
      LaTom.SetControlVisible('BDELETE', False);
    // FQ 11763 : Pas de modification
    if not JaiLeDroitConceptBureau(cstDroitModif) then
    begin
      // passe en lecture seule (voir TraiteArguments dans UtilTom pour les autres fiches)
      if (Lecran is TFFiche)
        and (TFFiche(Lecran).FTypeAction = taModif) then // car il ne faut pas bloquer la création
        TFFiche(Lecran).FTypeAction := taConsult;
      if (Lecran is TFFicheListe)
        and (TFFicheListe(Lecran).FTypeAction = taModif) then // car il ne faut pas bloquer la création
        TFFicheListe(Lecran).FTypeAction := taConsult;
    end;
  end;
end;

procedure GereDroitsConceptsAnnuaire(LaTof: TOF; Lecran: TForm);
// Gère les boutons et le mode d'ouverture de la tof
// en fonction des droits sur les concepts annuaire
begin

  { FQ 11763 (KPMG) : l'absence de droit modif n'enlève plus le droit d'insert
   donc mise en commentaire =>
  // Pas le droit de modifier
  if Not ExExJaiLeDroitConceptBureau(ccModifAnnuaireOuLiens) then
    begin
    // pas de création/suppression
    LaTof.SetControlVisible('BINSERT', False);
    LaTof.SetControlVisible('BDELETE', False);
    LaTof.SetControlVisible('BTLIENANNU', False);
    end

  // ou bien droits plus détaillés
  else }
  if ((AnnuaireLanceeDepuisSynthese) and (not JaiLeDroitConceptBureau(ccModifAnnuaireDossier))) or
    (not AnnuaireLanceeDepuisSynthese) then
  begin
    // Pas de création
    if not JaiLeDroitConceptBureau(ccCreatAnnuaireOuLiens) then
    begin
      LaTof.SetControlVisible('BINSERT', False);
      LaTof.SetControlVisible('BTLIENANNU', False);
    end;
    // Pas de suppression
    if not JaiLeDroitConceptBureau(ccSupprAnnuaireOuLiens) then
      LaTof.SetControlVisible('BDELETE', False);
  end;
end;


function SupprimeInfoDp(GuidPer: string; bForceSuppr: Boolean = False): Boolean;
// comprend la suppression de la base dossier !
var
  QQ: TQuery;
  nodoss, Texte, etat: string;
  bDossierAsp: Boolean;
  Suppr: TSupprDoss;
begin
  Result := False;
  nodoss := '';
  bDossierAsp := False;
  QQ := OpenSQL('select DOS_NODOSSIER, DOS_VERROU, DOS_EWSCREE, DOS_NETEXPERT from DOSSIER where DOS_GUIDPER="' + GuidPer + '"', TRUE);
  if not QQ.eof then
  begin
    nodoss := QQ.FindField('DOS_NODOSSIER').AsString;
    etat := QQ.FindField('DOS_VERROU').AsString;
    bDossierAsp := (GetParamsocDPSecur('SO_NE_EWSACTIF', False) and (QQ.FindField('DOS_EWSCREE').AsString = 'X'))
      or ((GetParamsocDPSecur('SO_NECWASURL', '') <> '') and (QQ.FindField('DOS_NETEXPERT').AsString = 'X'));
  end;
  Ferme(QQ);
  if nodoss <> '' then
  begin
    if (etat <> '') and (etat <> 'ENL') then
    begin
      PGIInfo('Le dossier ' + nodoss + ' n''est pas accessible. Suppression impossible.', TitreHalley);
      exit;
    end;
    // la suppression vérifie le password
    if not VerifPwdDossier(nodoss) then exit;
    // la suppression vérifie l'existence ASP
    if not bForceSuppr then
    begin
      if bDossierAsp then
      begin
        Texte := 'Ce dossier est géré en ASP, il ne peut pas être supprimé.'
          + '#10 La suppression sera possible lorsque le dossier ASP aura été réinitialisé.'
          + '#10 (menu Outils / Outils système / Outils d''assistance)';
        PGIInfo(Texte);
        exit;
      end;
      Texte := 'Vous allez supprimer définitivement les liens DP et le dossier ' + nodoss + ', ainsi que son répertoire.#13#10';
      if ExisteSQL('SELECT 1 FROM DOSSAPPLI WHERE DAP_NODOSSIER="' + nodoss + '" AND DAP_NOMEXEC="CPS5.EXE"') then
        texte := texte + ' Attention : les informations propres au dossier de paie (paramétrages prédéfinis dossier, etc...)#13'
          + ' vont être définitivement supprimées.#13';
      texte := texte + ' Confirmez-vous?#13';
      if PGIAsk(Texte, TitreHalley) <> mrYes then exit;
    end;
    // suppression physique du dossier
    Suppr := TSupprDoss.Create;
    Suppr.NoDossier := nodoss;
    Suppr.GuidPer := GuidPer;
    Suppr.ForceSuppr := bForceSuppr;
    if not Suppr.SuppressionDossier then
    begin
      PgiInfo('Suppression impossible.', TitreHalley);
      Suppr.free; exit;
    end;
    Suppr.free;
  end;
  Result := TRUE;
end;


function TestSupprimePersonne(GuidPer: string; EstBavard: Boolean = True): boolean;
// prépare suppression d'une fiche annuaire (contrôle des dépendances)
// + suppression de quelques données annexes (contacts/évènements)
var PerSup: TPerSup;
  Q: TQuery;
begin
  Result := False;
  PerSup := TPerSup.Create;
  PerSup.GuidPer := GuidPer;
  PerSup.Nomper := '';
  if GuidPer <> '' then
  begin
    Q := OpenSQL('select ANN_NOMPER from ANNUAIRE where ANN_GUIDPER="' + GuidPer + '"', True);
    if not Q.Eof then PerSup.Nomper := Q.FindField('ANN_NOMPER').AsString;
    Ferme(Q);
  end;
  PerSup.Bavard := EstBavard;
  // test complet sur les tables dépendantes de la personne
  if PerSup.TestEnregPer then
  begin
    // suppression contacts/évènements
    if Transactions(PerSup.SuppressionEnregPer, 0) = oeOk then
      Result := True;
  end;
  PerSup.Free;
end;


procedure SupprimeResteAnnuaire(GuidPer: string);
// Suppression définitive sur une fiche annuaire et ses dépendances.
begin
  BeginTrans;
  try
    ExecuteSQL('delete from ANNUAIRE where ANN_GUIDPER="' + GuidPer + '"');
    ExecuteSQL('delete from ANNUBIS where ANB_GUIDPER="' + GuidPer + '"');
    ExecuteSQL('delete from LIENSOLE where (LO_TABLEBLOB="ANN" OR LO_TABLEBLOB="JUR") AND LO_IDENTIFIANT="' + GuidPer + '"');
    // mcd - efface zone guidper dans table contact sur tiers
    ExecuteSQL('delete from CONTACT where C_TYPECONTACT="ANN" AND C_GUIDPER="' + GuidPer + '"');
    ExecuteSql('update CONTACT set C_GUIDPER="" from contact where c_typecontact="T" and C_GUIDPER="' + GuidPer + '"');
  except
    V_PGI.IoError := oeUnknown;
  end;
  if V_PGI.IoError <> oeUnknown then CommitTrans else Rollback;
end;


procedure SupprimeEnregAnnu(GuidPer: string);
var Texte: string;
// suppression complète d'une personne dans l'annuaire
begin
  // suppression personne, avec vérifications
  if not TestSupprimePersonne(GuidPer) then exit;
  Texte := 'Vous allez supprimer définitivement cette personne.#13#10Confirmez-vous l''opération ?';
  if PGIAsk(Texte, TitreHalley) <> mrYes then
    exit; // SORTIE
  // suppression DP
  if SupprimeInfoDp(GuidPer) then
    SupprimeResteAnnuaire(GuidPer);
end;


{***********A.G.L.***********************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 18/07/2000
Modifié le ... :   /  /
Description .. : Fonction permettant de supprimer une personne dans
Suite ........ : la liste de l'annuaire
Mots clefs ... : SUPPRESSION;ANNUAIRE
*****************************************************************}

procedure AGLSupprimeListAnnu(parms: array of variant; nb: integer);
var F: TForm;
  L: THDBGrid;
  Q: THQuery;
  i: Integer;
  GuidPer, Texte: string;
  bavard: Boolean;
begin
  F := TForm(Longint(Parms[0]));
  if F = nil then exit;
  L := THDBGrid(F.FindComponent('FListe'));
  if L = nil then exit;
  Q := THQuery(F.FindComponent('Q'));
  if (Q = nil) then exit;
  // sélection
  if (L.NbSelected = 0) and (not L.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné', TitreHalley);
    exit;
  end;
  Texte := 'Vous allez supprimer définitivement la sélection (sauf les fiches faisant l''objet de liens avec des dossiers).#13#10Confirmez-vous l''opération ?';
  if PGIAsk(Texte, TitreHalley) = mrYes then
  begin
    if L.AllSelected then
    begin
      InitMove(10, '');
{$IFDEF EAGLCLIENT}
      if not TFMul(F).FetchLesTous then
        PGIInfo('Impossible de récupérer tous les enregistrements')
      else
{$ENDIF}
      begin
        while not Q.EOF do
        begin
          MoveCur(False);
          GuidPer := Q.FindField('ANN_GUIDPER').AsString;
          // MD 07/05/02 => test non bavard, et message global
          if not TestSupprimePersonne(GuidPer, False) then
          begin
            Q.Next;
            Continue;
          end;
          // supression
          if SupprimeInfoDp(GuidPer) then
            SupprimeResteAnnuaire(GuidPer);
          Q.Next;
        end;
      end;
      // Déselectionne
      TFMul(F).bSelectAllClick(nil);
    end
    else
    begin
      InitMove(L.NbSelected, '');
      // au delà de 10, on cause plus
      bavard := (L.NbSelected < 10);
      for i := 0 to L.NbSelected - 1 do
      begin
        MoveCur(False);
        L.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
        Q.TQ.Seek(L.Row - 1);
{$ENDIF}
        GuidPer := Q.FindField('ANN_GUIDPER').AsString;
        // MD 07/05/02 => test bavard + rajout message
        if not TestSupprimePersonne(GuidPer, bavard) then Continue;
        // suppression
        if SupprimeInfoDp(GuidPer) then
          SupprimeResteAnnuaire(GuidPer);
      end;
      L.ClearSelected;
    end;
    FiniMove;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 18/07/2000
Modifié le ... :   /  /
Description .. : Fonction permettant de supprimer une personne dans
Suite ........ : annuaire
Mots clefs ... : SUPPRESSION;ANNUAIRE
*****************************************************************}

procedure AGLSupprimeEnregAnnu(parms: array of variant; nb: integer);
var GuidPer: string;
begin
  GuidPer := Parms[0];
  SupprimeEnregAnnu(GuidPer);
end;

{***********A.G.L.***********************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 30/10/2002
Modifié le ... :   /  /
Description .. : Ferme et rouvre une fiche avec le nouveau paramètre
Suite ........ :
Mots clefs ... :
*****************************************************************}

procedure AGLFermeEtOuvre(parms: array of variant; nb: integer);
begin
  FermeEtOuvre(string(parms[1]), string(parms[2]),
    string(parms[3]), string(parms[4]), string(parms[5]));
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 31/08/2000
Modifié le ... :   /  /
Description .. : Teste la validité du Siret
Suite ........ : Chaine : Chaine Siret à analyser
 ---             Taille : Taille de la chaine à analyser
 --- Retour    : True si l'opération a réussie
 ---           : False si l'opération a échouée
Mots clefs ... : SIRET
*****************************************************************}

function CoherenceSiret(Chaine: string; Taille: Integer): Boolean;
var Resultat: Boolean;
  Indice, Indice2, Element, Total, Coeff: Integer;
  ChElement, ChTotal: string;
begin
  Resultat := True;

 //--- On Vérifie que la longueur de la chaine est égale à Taille
  if (Length(Chaine) <> Taille) then
    Resultat := False;

 //--- On vérifie que la chaine comporte que des caractères numériques
  if (Resultat) then
    for Indice := 1 to Taille do
      if (not (Chaine[Indice] in ['0'..'9'])) then
      begin
        Resultat := False;
        break;
      end;

 //--- Controle de validité
  if (Resultat) then
  begin
    Coeff := 1; Total := 0; //Element:=0;
    for Indice := Taille downto 1 do
    begin
      Element := StrToInt(Chaine[Indice]) * Coeff;
      ChElement := (IntToStr(Element));
      for Indice2 := 1 to length(ChElement) do
        Total := Total + StrToInt(ChElement[Indice2]);
      if (Coeff = 1) then
        Coeff := 2
      else
        Coeff := 1;
    end;
    ChTotal := IntToStr(Total);
    if (length(ChTotal) > 1) and (ChTotal[2] <> '0') then
      Resultat := False;
  end;
  CoherenceSiret := Resultat;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 31/08/2000
Modifié le ... :   /  /
Description .. : Teste la validité Code frp
Suite ........ : Chaine : Chaine Siret à analyser
 --- Retour    : True si l'opération a réussie
 ---           : False si l'opération a échouée
Mots clefs ... : FRP
*****************************************************************}

function CoherenceFrp(Chaine: string): Boolean;
var Resultat: Boolean;
  St, Cle: string;
  indice: integer;
  Clecalcule: integer;
begin
  Resultat := True;

  for indice := 1 to Length(Chaine) do
  begin
    if (Chaine[Indice] in ['0'..'9']) then continue;
    if (Indice = 0) and (Chaine[indice] = 'B') then continue;
    if (Indice = 1) and ((Chaine[indice] = 'A') or (Chaine[indice] = 'B')) then continue
    else begin Resultat := FALSE; break; end;
  end;

  St := Copy(Chaine, 1, 3);
  if St = '2A0' then St := '201';
  if St = '2B0' then St := '202';
  if St = 'B31' then St := '981';
  St := St + Copy(Chaine, 4, 4);
  Clecalcule := (Strtoint(St)) mod 97;
  Clecalcule := Clecalcule * 10 + 1;
  Clecalcule := Clecalcule mod 97;
  St := Copy(Chaine, 8, 6);
  Clecalcule := (Clecalcule * 1000000) + (Strtoint(St));
  Clecalcule := Clecalcule mod 97;
  Clecalcule := Clecalcule * 1000;
  Clecalcule := Clecalcule mod 97;
  Cle := Copy(Chaine, 14, 2);
  if (Cle = '') or (Clecalcule <> StrToInt(Cle)) then
    Resultat := False;
  CoherenceFrp := Resultat;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function CtrlCodeFRP(AffMsg: Boolean; CodeFRP: string): Boolean;
var
  sCleCalculee_l: string;
begin
  Result := CtrlEtCalculeCodeFRP(AffMsg, false, CodeFRP, sCleCalculee_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function CalculerCodeFRP(CodeFRP: string; var sCleCalculee_p: string): Boolean;
begin
  Result := CtrlEtCalculeCodeFRP(false, true, CodeFRP, sCleCalculee_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function CtrlEtCalculeCodeFRP(AffMsg, bCalcul_p: Boolean; CodeFRP: string; var sCleCalculee_p: string): Boolean;
var
  sRecette, sOrdre, sCle: string;
  sDir, sCentre, sBase: string;
  iRecette, iOrdre, iCle, X: Integer;
begin
  Result := (Length(CodeFRP) = 0);
  if Result then
    Exit; // Sasie Vide

  if not bCalcul_p then
    Result := (Length(CodeFRP) in [0, 15])
  else
    Result := (Length(CodeFRP) in [0, 13]);
  if not Result then
    Exit;

  sRecette := Copy(CodeFRP, 1, 7);
  sDir := Copy(CodeFRP, 1, 3);
  if (sDir = '2A0') then
    sRecette := Concat('201', Copy(CodeFRP, 4, 4))
  else if (sDir = '2B0') then
    sRecette := Concat('202', Copy(CodeFRP, 4, 4))
  else if (sDir = 'B31') then
  begin
    sRecette := '981';
    sCentre := Copy(CodeFRP, 4, 2);
    if (sCentre = '4A') then
      sRecette := Concat(sRecette, '23')
    else
      sRecette := Concat(sRecette, Copy(CodeFRP, 4, 2));

    sBase := Copy(CodeFRP, 6, 2);
    if (sBase = '05') then
      sRecette := Concat(sRecette, '01')
    else
      sRecette := Concat(sRecette, Copy(CodeFRP, 6, 2));
  end
  else if (sDir = 'A45') then
  begin
    sRecette := '007';
    sCentre := Copy(CodeFRP, 4, 2);
    if (sCentre = '7V') then
      sRecette := Concat(sRecette, '01')
    else
      sRecette := Concat(sRecette, Copy(CodeFRP, 4, 2));

    sBase := Copy(CodeFRP, 6, 2);
    if (sBase = '10') then
      sRecette := Concat(sRecette, '00')
    else
      sRecette := Concat(sRecette, Copy(CodeFRP, 6, 2));
  end;

  sOrdre := Copy(CodeFRP, 8, 6);
  if bCalcul_p then
    sCle := '0'
  else
    sCle := Copy(CodeFRP, 14, 2);

  try
    iRecette := StrToInt(sRecette);
    iOrdre := StrToInt(sOrdre);
    iCle := StrToInt(sCle);
  except
    if AffMsg then
      PGIInfo('Le code FRP saisi n''est pas valide.', TitreHalley);
    Result := False; Exit;
  end;

  X := iRecette mod 97;
  X := ((X * 10) + 1);
  X := X mod 97;
  X := (X * 1000000) + iOrdre;
  X := X mod 97;
  X := (X * 1000);
  X := X mod 97;
  if not bCalcul_p then
    Result := (X = iCle)
  else
  begin
    Result := true;
    sCleCalculee_p := IntToStr(X);
  end;
  if not Result and AffMsg then
    PGIInfo('Le clé du code FRP saisi n''est pas valide.', TitreHalley);
end;


function GetDosjuri(GuidPer: string): string;
var
  QQ: TQuery;
begin
  QQ := OpenSQL('select JUR_CODEDOS from JURIDIQUE where JUR_GUIDPERDOS="' + GuidPer + '" AND '
    + 'JUR_TYPEDOS="STE" AND JUR_NOORDRE=1', TRUE);
  if not QQ.Eof then
  begin
    if (QQ.FindField('JUR_CODEDOS').AsString <> '') and
      (QQ.FindField('JUR_CODEDOS').AsString <> '&#@') then
      result := QQ.FindField('JUR_CODEDOS').AsString
    else
      result := '&#@';
  end
  else
    result := '&#@';
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. : Renvoie le code de la personne en fonction du code du
Suite ........ : dossier
Mots clefs ... :
*****************************************************************}

function GetPerJuri(sCodeDos_p: string): string;
var
  QQ: TQuery;
begin
  result := '';
  QQ := OpenSQL('select JUR_GUIDPERDOS ' +
    'from JURIDIQUE ' +
    'where JUR_CODEDOS = "' + sCodeDos_p + '" ' +
    '  AND JUR_TYPEDOS = "STE" ' +
    '  AND JUR_NOORDRE = 1', TRUE);
  if not QQ.Eof then
    result := QQ.FindField('JUR_GUIDPERDOS').AsString;
  Ferme(QQ);
end;

procedure AGLDupliquerElementAnn(parms: array of variant; nb: integer);
var
  Table, sFiche, sCle, sDom: string;
begin
  Table := string(parms[0]);
  if Table = 'ANNUAIRE' then
    if PGIAsk('Voulez-vous vraiment dupliquer cette personne ?', TitreHalley) = mrYes then
    begin
      sCle := DupliquerAnnuaire(string(parms[1]));
      sFiche := 'ANNUAIRE';
      sDom := 'YY';
      AGLLanceFiche(sDom, sFiche, sCle, sCle, '');
    end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 26/06/2006
Modifié le ... :   /  /
Description .. : Passage en TO pou meilleur compatibilité eAGL
Mots clefs ... :
*****************************************************************}

function DupliquerAnnuaire(sOldCle_p: string): string;
var
  OBOld_l, OBNew_l: TOB;
  sNewCle_l: string;
begin
  Result := '';

  OBOld_l := TOB.Create('ANNUAIRE', nil, -1);
  OBOld_l.LoadDetailDBFromSQL('ANNUAIRE',
    'SELECT * FROM ANNUAIRE ' +
    'WHERE ANN_GUIDPER = "' + sOldCle_p + '"');
  if OBOld_l.Detail.Count = 0 then
  begin
    OBOld_l.Free;
    exit;
  end;

  OBNew_l := TOB.Create('ANNUAIRE', nil, -1);
  OBNew_l.InitValeurs(true);
  OBNew_l.Dupliquer(OBOld_l, true, true, true);
  sNewCle_l := AglGetGuid;

  OBNew_l.Detail[0].PutValue('ANN_GUIDPER', sNewCle_l);
  OBNew_l.Detail[0].PutValue('ANN_CODEPER', -2);
  OBNew_l.Detail[0].PutValue('ANN_TIERS', '');
  //mcd 17/05/05 champ supprime QQ.FindField('ANN_AUXILIAIRE').AsString := '';

  OBNew_l.Detail[0].InsertDB(nil);
  OBOld_l.Free;
  OBNew_l.Free;

  Result := sNewCle_l;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 25/10/02
Fonction ..... : SelectPersonne
Description .. : Recherche puis sélection d'une personne
Paramètres ... : Id de la personne
                 Nom de recherche de la personne
*****************************************************************}

procedure SelectPersonne(var sGuidPer_p, sTannNom_p: string);
var
  SQL, sLib_l, sCleSel_l: string;
  vValCode_l: variant;
begin
  if sTannNom_p = '' then
    sCleSel_l := ''
  else
    sCleSel_l := 'ANN_NOMPER = ' + VarToStr(GetValChamp('ANNUAIRE', 'ANN_NOMPER', 'ANN_GUIDPER="' + sGuidPer_p + '"'));
  SQL := 'select ANN_NOM1, ANN_NOM2 from ANNUAIRE where ANN_GUIDPER=';
  sLib_l := OnElipsisClick('YY', 'ANNUAIRE_SEL', sCleSel_l, SQL, 'ANN_NOM1', 'ANN_NOM2', vValCode_l);

  if VarToStr(vValCode_l) = '' then
    sGuidPer_p := ''
  else
    sGuidPer_p := vValCode_l;
  sTannNom_p := sLib_l;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 29/10/02
Procédure .... : FermeEtOuvre
Description .. : Ouvre, ferme et re-ouvre une fiche afin de recharger
                 la fiche avec le doublon
Paramètres ... : Nature de la fiche
                 Nom de la fiche
                 Range
                 Lequel
                 Paramètres
*****************************************************************}

procedure FermeEtOuvre(sNature_p, sFiche_p, sRange_p, sLequel_p, sArgument_p: string);
var
  sCodeDoublon_l, sParametres_l, sAction_l: string;
begin
  sParametres_l := AGLLanceFiche(sNature_p, sFiche_p, sRange_p, sLequel_p, sArgument_p);
  READTOKENPipe(sParametres_l, 'ANN_GUIDPER=');
  sCodeDoublon_l := READTOKENST(sParametres_l);
  if sCodeDoublon_l <> '' then
  begin
    sAction_l := READTOKENST(sArgument_p);
    if sAction_l = 'ACTION=CREATION' then
      sAction_l := ';';
    sArgument_p := sAction_l + sArgument_p;
    AGLLanceFiche(sNature_p, sFiche_p, sRange_p, sCodeDoublon_l, sArgument_p);
  end;
end;

//--------------------------------------------------------------------
//--- Nom   : VerifierDoublon
//--- Objet : Vérifie les doublons sur le nom de recherche (DOUBLONS)
//---         Vérifie les doublons sur le siret ((DOUBLONSIRET)
//--------------------------------------------------------------------

function VerifierDoublon(TypeDoublon, ChSiren, ChCleSiret, ChPmPp, ChNom1, ChNom2, ChNomPer: string): string;
var CtrlDoublon: Boolean;
  ChSql: string;
begin
  Result := '-1';
 //--- Vérifie doublon sur ANN_NOM1+ ANN_NOM2+ ANN_NOMPER
  if (TypeDoublon = 'DOUBLONS') then
  begin
    CtrlDoublon := GetParamSocDpSecur('SO_JUCTRLDOUBLON', True); // ParamSoc du juridique
    if (ChPmPp = 'PP') and (CtrlDoublon) then
      ChSql := FormateCroise(ChNom1, ChNom2, ChNomPer);
    if (ChPmPp = 'PM') and (CtrlDoublon) then
      ChSql := FormateCroise(ChNom1, '', ChNomPer);
  end
  else
  //--- Vérifie doublon sur ANN_SIREN + ANN_CLESIRET
    if (TypeDoublon = 'DOUBLONSIRET') then
      ChSql := 'ANN_SIREN="' + ChSiren + '" AND ANN_CLESIRET="' + ChCleSiret + '"';

// if ExisteSQL('Select ANN_GUIDPER from ANNUAIRE where '+ChSql) then
//  Result:=AGLLanceFiche('YY','ANNUAIRE_SEL','','','ACTION=CONSULTATION;'+TypeDoublon+';'+ChSql)

  if ExisteSQL('Select ANN_GUIDPER from ANNUAIRE where ' + ChSql) then
  begin
{$IFDEF BUREAU}
    if JaiLeDroitConceptBureau(ccListesAnnSimplifiees) then
      Result := AGLLanceFiche('YY', 'ANNUAIRE_SELLITE', '', '', 'ACTION=CONSULTATION;' + TypeDoublon + ';' + ChSql)
    else
      Result := AGLLanceFiche('YY', 'ANNUAIRE_SEL', '', '', 'ACTION=CONSULTATION;' + TypeDoublon + ';' + ChSql);
{$ELSE}
    Result := AGLLanceFiche('YY', 'ANNUAIRE_SELLITE', '', '', 'ACTION=CONSULTATION;' + TypeDoublon + ';' + ChSql);
{$ENDIF}
  end;
end;


{*****************************************************************
Auteur ....... : BM
Date ......... : 30/10/02
Fonction ..... : FormateCroise
Description .. : Formatage croisé de requete entre NOM1, NOM2 et NOMPER
Paramètres ... : NOM1, NOM2 et NOMPER
*****************************************************************}

function FormateCroise(sValNom1_p, sValNom2_p, sValNomPer_p: string): string;
var
  sWhere_l, sWhere1_l, sWhere2_l, sWhere3_l, sWhere4_l: string;
begin
  if sValNom1_p <> '' then
  begin
    sWhere1_l := 'upper(ANN_NOM1) like "' + sValNom1_p + '%"';
    sWhere2_l := 'upper(ANN_NOMPER) like "' + sValNom1_p + '%';
  end;

  if sValNom2_p = '' then
  begin
    if sWhere2_l <> '' then
      sWhere2_l := sWhere2_l + '" ';
  end
  else
  begin
    if sWhere1_l <> '' then
      sWhere1_l := sWhere1_l + ' and ';
    sWhere1_l := sWhere1_l + 'upper(ANN_NOM2) like "' + sValNom2_p + '%"';

    if sWhere2_l = '' then
      sWhere2_l := 'upper(ANN_NOMPER) like "'
    else
      sWhere2_l := sWhere2_l + ' ';
    sWhere2_l := sWhere2_l + sValNom2_p + '%"';
  end;

  if sValNomPer_p <> '' then
  begin
    sWhere3_l := 'upper(ANN_NOMPER) like "' + sValNomPer_p + '%"';
    sWhere4_l := 'upper(ANN_NOM1) || " " || upper(ANN_NOM2) like "' + sValNomPer_p + '%" ';
  end;

  if sWhere1_l <> '' then
    sWhere_l := '( ' + sWhere1_l + ' ) ';
  if sWhere2_l <> '' then
  begin
    if sWhere_l <> '' then
      sWhere_l := sWhere_l + ' OR ';
    sWhere_l := sWhere_l + '( ' + sWhere2_l + ' )';
  end;
  if sWhere3_l <> '' then
  begin
    if sWhere_l <> '' then
      sWhere_l := sWhere_l + ' OR ';
    sWhere_l := sWhere_l + '( ' + sWhere3_l + ' )';
  end;
  if sWhere4_l <> '' then
  begin
    if sWhere_l <> '' then
      sWhere_l := sWhere_l + ' OR ';
    sWhere_l := sWhere_l + '( ' + sWhere4_l + ' )';
  end;

  result := sWhere_l;
end;

{------------- objet TSupprDoss -----------}

function TSupprDoss.SuppressionBasePhysique: Boolean;
{$IFNDEF EAGLCLIENT}
var
  StDatabase: string;
  supprok: Boolean;
{$ENDIF}
begin
  Result := FALSE;

{$IFDEF EAGLCLIENT}
{$IFDEF BUREAU}
  if VH_DP.ePlugin <> nil then
    Result := VH_DP.ePlugin.DeleteBase('DB' + VH_DOSS.NoDossier);
{$ELSE}
     // MD 03/02/06 : En Web Access, pas de suppression possible
     // de la base physique si pas mode DP
  PGIInfo('Vous ne pouvez pas faire la suppression physique du dossier.', TitreHalley);
  exit;
{$ENDIF}
{$ELSE}
  // si dossier invalide, on sort
  if (VH_Doss.NoDossier = '') or (VH_Doss.PathDos = '') then
  begin
    PgiInfo('Le dossier n''a pas d''informations de connexions suffisantes.' + #13
      + 'La base n''existe pas ou ne sera pas supprimée...', TitreHalley);
    exit;
  end;

  if (V_PGI.Driver = dbMsSQL) or (V_PGI.Driver = dbMSSQL2005) then
  begin
    // nom de la DB
    StDatabase := VH_Doss.DBSocName; // ou StDatabase := 'DB'+NoDossier
    supprok := False;
    try
      supprok := SupprimeBaseSql(StDatabase);
    except
    end;
    if not supprok then
    begin
      PgiInfo('La base de données ' + StDatabase + ' n''a pas pu être supprimée.', TitreHalley);
      exit;
    end;
  end
  // ####DRIVER NON SUPPORTE
  else
  begin
    // normalement, cas déjà éliminé lors de la création...
    PgiBox('Le driver "' + DriverToSt(V_PGI.Driver, True) + '" de la connexion en cours n''est pas supporté dans un environnement multi-dossier.', TitreHalley);
    exit;
  end;
  Result := True;
{$ENDIF}
end;


procedure TSupprDoss.SuppressionEnregDPSTD;
var i: Integer;
  nomtbl, prefix, sql: string;
Mcd : IMCDServiceCOM;
Table     : ITableCOM ;
TableList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  TableList := Mcd.Tables;
  TableList.Reset;
	While TableList.MoveNext do
  begin
		Table := TableList.Current as ITableCOM;
    // élimine tables non communes (= DT_COMMUN<>'D' et DT_COMMUN<>'S')
    if (Table.TypeDP = 0) then Continue;
    nomtbl := Table.Name;

    // contient un NODOSSIER char(5) qui est l'ancien no (gamme 2)
    // utilisé pour les transferts => rien à voir avec le DP !
    if nomtbl = 'TRFDOSSIER' then Continue;
    // table traitée en dehors car utilisée dans les muls
    if nomtbl = 'ANNUAIRE' then Continue;
    // traitée plus bas, car ListChp ne contient pas ses chps car topée X et non D ou S
    if nomtbl = 'DOSSAPPLI' then Continue;
    prefix := Table.Pref;
    sql := '';

    // seules les tables communes dépendantes de l'annuaire sont à traiter
    // c'est à dire ayant no dossier, guidperdos, ou guidper

    // Si double réf. (ex: DOSSIER contient NODOSSIER et GUIDPER), privilégie nodossier
    // (rq : si tdDPSTD on devrait rajouter and XX_PREDEFINI='DOS')
    if ListChpDPSTD.IndexOf(prefix + '_NODOSSIER') > -1 then
    begin
      if NoDossier <> '' then
      begin
        sql := 'DELETE FROM ' + nomtbl + ' WHERE ' + prefix + '_NODOSSIER="' + NoDossier + '"';
        // traitement particulier : limité aux enregdp ####
        // on ne touche pas aux enreg rattachés à un dossier juridique
        if nomtbl = 'JUEVENEMENT' then sql := sql + ' AND JEV_CODEDOS="&#@"';
      end;
    end
    // sinon code personne d'origine dossier
    else if ListChpDPSTD.IndexOf(prefix + '_GUIDPERDOS') > -1 then // 2 tables : ANNULIEN et JURIDIQUE
    begin
      sql := 'DELETE FROM ' + nomtbl + ' WHERE ' + prefix + '_GUIDPERDOS="' + GuidPer + '"';
      // traitement particulier : limité aux enregdp ####
      // (les autres étant supprimés par Juri lors de la suppression du dossier juridique ?)
      if nomtbl = 'ANNULIEN' then sql := sql + ' AND ANL_CODEDOS="&#@"'
      else if nomtbl = 'JURIDIQUE' then sql := sql + ' AND JUR_CODEDOS="&#@"';
    end;

    // TRAITEMENT
    if sql <> '' then ExecuteSQL(sql);
  end;

  // table DOSSAPPLI topée X
  if NoDossier <> '' then
    ExecuteSQL('DELETE FROM DOSSAPPLI WHERE DAP_NODOSSIER="' + NoDossier + '"');
end;


function TSupprDoss.SuppressionDossier: Boolean;
var
  DossEnCours: string;
  supprok: Boolean;
  Q: TQuery;
  lstGuid: TStringList; //LstDoc : TStringList; // $$$ JP 01/08/06: même gestion pour suppr' des messages
  i: Integer;
begin
     // Par défaut, dossier non supprimé
  Result := False;

     // MD 27/01/06 - pour le transport, on doit pouvoir forcer la suppression
  if not ForceSuppr then
       // MD 12/12/00 - sécurité
    if not JaiLeDroitTag(26051) then exit;

     // pointe sur le dossier à supprimer
  if VH_Doss = nil then
    VH_Doss := TDossSelect.Create;
  DossEnCours := VH_Doss.NoDossier;
  if NoDossier <> DossEnCours then
    VH_Doss.NoDossier := NoDossier;

     // 1. TESTS SYSTEMES/FONCTIONNELS
     // mauvais dossier
{$IFNDEF EAGLCLIENT}
  if (VH_Doss.NoDossier = '') or (VH_Doss.PathDos = '') then
{$ELSE}
  if VH_Doss.NoDossier = '' then
{$ENDIF}
    PgiInfo('Le dossier ' + NoDossier + ' n''a pas d''informations de connexions suffisantes.'#13#10' La base n''existe pas ou ne sera pas supprimée', TitreHalley)

  else if V_PGI.NoDossier = VH_Doss.NoDossier then // $$$ JP 28/12/05 V_PGI.DBName=VH_Doss.DBSocName then
    PgiInfo('Le dossier ' + NoDossier + ' est celui du cabinet. Vous ne pouvez pas le supprimer.', TitreHalley)

  else if ExisteSQL('SELECT 1 FROM YMESSAGES WHERE YMS_NODOSSIER="' + NoDossier + '" AND YMS_TRAITE="-"') then
    PgiInfo('Le dossier ' + NoDossier + ' comporte un ou plusieurs messages non clôturés.'#13#10' Vous ne pouvez pas le supprimer.', TitreHalley)
  else
  begin
          // 2. SUPPRESSION DU DOSSIER ASP
          // #### manque la suppression dans Business Line !!! #### if GetParamSocDPSecur('SO_NECWASURL', '')<>'' then ...
{$IFDEF EWS}
          // if GetParamsocDPSecur('SO_NE_EWSACTIF', False) and IsDossierEws(NoDossier) then EwsSupprimeClient(NoDossier);
{$ENDIF}
          // MD 28/10/05 - DONC ON NE SUPPRIME PLUS EN ASP, CAR TROP DE CONSEQUENCES !

          // 3. SUPPRESSION BASE PHYSIQUE
          // MD 20/02/01 - Considère suppression ok si la base n'existe plus
    supprok := not DBExists('DB' + VH_Doss.NoDossier);
    if supprok = FALSE then
    begin
               // essaie suppression physique
               // $$$ JP 28/12/05 - inutile: supprok := False;
      try
        supprok := SuppressionBasePhysique;
      except
        PgiInfo('Erreur pendant suppression de la base du dossier ' + NoDossier, TitreHalley);
      end;
    end;
{$IFNDEF EAGLCLIENT}
          // Suppression du log résiduel
    ExecuteMSSQL7('xp_cmdshell ''del ' + VH_Doss.PathLdf + '\' + VH_Doss.DBSocName + '.ldf''');
{$ENDIF}

          // 4. SUPPRESSION ELEMENTS LIES A LA GED
    if supprok = TRUE then
    begin
               // On détruit dans les tables les plus hautes, les toms font la suppression en cascade
               // MESSAGES
               // $$$ JP 01/08/06: même gestion pour suppr' des messages
      lstGuid := TStringList.Create; //LstDoc := TStringList.Create;
      Q := OpenSQL('SELECT YMS_MSGGUID FROM YMESSAGES WHERE YMS_NODOSSIER="' + NoDossier + '"', True, -1, '', True);
      while not Q.Eof do
      begin
        lstGuid.Add(Trim(Q.FindField('YMS_MSGGUID').AsString));
        Q.Next;
      end;
      Ferme(Q);
      for i := 0 to lstGuid.Count - 1 do
        if lstGuid[i] <> '' then
          DeleteMessage(lstGuid[i]);
      lstGuid.Clear;

               // DOCUMENTS
      Q := OpenSQL('SELECT DPD_DOCGUID FROM DPDOCUMENT WHERE DPD_NODOSSIER="' + NoDossier + '"', True, -1, '', True);
      while not Q.Eof do
      begin
        lstGuid.Add(Q.FindField('DPD_DOCGUID').AsString); // $$$ JP 01/08/06: même gestion pour suppr' des messages
        Q.Next;
      end;
      Ferme(Q);
      for i := 0 to lstGuid.Count - 1 do // $$$ JP 01/08/06: même gestion pour suppr' des messages
        SupprimeDocumentGed(lstGuid[i]); // $$$ JP 01/08/06: même gestion pour suppr' des messages
      lstGuid.Free; //LstDoc.Free;// $$$ JP 01/08/06: même gestion pour suppr' des messages
    end;

          // 5. SUPPRESSION ENREG DP/DOSSIER
    if supprok = TRUE then
    begin
      ListChpDPSTD := TStringList.Create;

               // charge liste de chps en dehors des transactions car select sur tables systèmes
      ChargeListChpDPSTD(ListChpDPSTD);
      if Transactions(SuppressionEnregDPSTD, 3) = oeOK then //if Transactions(SuppressionEnregDPSTD,1) = oeOK then
      begin
        Result := True; // suppression ok

                    // MD 22/02/01 - suppression répertoire dossier
{$IFNDEF EAGLCLIENT}
        RemoveInDir(VH_Doss.PathDos, True, True);
        ExecuteMSSQL7('xp_cmdshell ''rmdir ' + VH_Doss.PathLdf + '''');
{$ENDIF}
                    // si c'était le dossier en cours, il n'est plus valide
        if NoDossier = DossEnCours then
          DossEnCours := '';
      end;
    end;
  end;

     // se repositionne sur le dossier en cours (DossEnCours peut être ''
     // on vient de supprimer le dossier qui était en cours de sélection)
  if VH_Doss.NoDossier <> DossEnCours then
    VH_Doss.NoDossier := DossEnCours;
end;


constructor TSupprDoss.Create;
begin
  ListChpDPSTD := nil;
end;

destructor TSupprDoss.Destroy;
begin
  if ListChpDPSTD <> nil then ListChpDPSTD.Free;
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{----------- Objet TPerSup -----------}

procedure TPerSup.SuppressionEnregPer;
begin
  // suppose TestEnregPer fait...
  if (V_PGI.IoError = oeOK) and SupprEvents then SupprimeEvents;
end;


function TPerSup.TestEnregPer: Boolean;
// Indique si on peut effectuer la suppression de la personne
var Msg: string;
  oksuppr: Boolean;
begin
  Msg := '';

  // GI / Compta
  oksuppr := (not TestCompteTiers(Msg));

  // JURI
  // dossiers juridique (JURIDIQUE where guidperdos=...)
  oksuppr := oksuppr and (not TestJuriDosPer(Msg));
  // liens vers des dossiers juridiques (JURIDIQUE, ANNULIEN where ANL_GUIDPER=...)
  oksuppr := oksuppr and (not TestJuriLienPer(Msg));
  // infos dossiers (, JUDOSINFO where JDI_GUIDPER=...)
  oksuppr := oksuppr and (not TestJuriDosInfoPer(Msg));
  // groupes de sociétés (JUGRPSOC where JGR_GUIDPER=...)
  oksuppr := oksuppr and (not TestJuriGroupeSocPer(Msg));
  // évènements (JUEVENEMENT where JEV_GUIDPER=...)
  oksuppr := oksuppr and (not TestJuriEvenementPer(Msg));

  // LIENS DIVERS
  // liens dp...
  oksuppr := oksuppr and (not TestLienPer(Msg));
  // liens IFU
  oksuppr := oksuppr and (not TestLienPlaq(Msg));

  // NOMADE
  oksuppr := oksuppr and (not TestDossierParti(Msg));

  // FISCALITE PERSONNELLE
  oksuppr := oksuppr and (not TestFisPer(Msg));

  Result := oksuppr;
end;


function TPerSup.ExistePlaquettes(nomchamp, nomtable, chpGuidPer, declaration, texte: string;
  var Msg: string): Boolean;
// Retourne True si des liens existent dans telle plaquette
// Complète le message passé en paramètre
var Q: TQuery;
begin
  Result := False;
  // #### A FAIRE GUID : quand les tables liasses seront modifiées
  exit;
{  if declaration  = '' then
    Q := OpenSQL('select '+nomchamp+' from '+nomtable+' where '+chpGuidPer+'="'+GuidPer+'" group by '+nomchamp,true)
  else
    Q := OpenSQL('select '+nomchamp+' from '+nomtable+' where '+chpGuidPer+'="'+GuidPer+'" and L02_LIALOGICIEL="'+declaration+'" group by '+nomchamp,true);
  if Not Q.Eof then Result := True; }
  Msg := Msg + Texte;
  while not Q.Eof do
  begin
    Msg := Msg + ' ' + Q.FindField(nomchamp).AsString;
    Q.Next;
  end;
  Msg := Msg + #13 + #10;
  Ferme(Q);
end;


function TPerSup.TestCompteTiers(var Msg: string): boolean;
// retourne True si on ne peut pas supprimer
var Q1: TQuery;
  Tiers, txt: string;
begin
  Result := False;
  // pour le compte tiers
  Q1 := OpenSQL('select ANN_TIERS from ANNUAIRE where ANN_GUIDPER="' + GuidPer + '"', TRUE);
  Tiers := '';
  if not Q1.eof then Tiers := Q1.Fields[0].AsString;
  Ferme(Q1);
  if Tiers <> '' then
  begin
    if ExisteSQL('select T_TIERS from TIERS where T_TIERS = "' + Tiers + '"') then
    begin
      Result := True;
      txt := '=> Attention, cette fiche est liée au tiers ' + Tiers + ' dans la compta ou dans la gestion interne.' + #13 + #10;
      txt := txt + ' Supprimer d''abord le tiers pour pouvoir supprimer la fiche.';
      //if Bavard then PGIInfo(txt, 'Suppression de '+NomPer);
      // ON REPASSE A TRUE ET ON DONNE LE CHOIX A L'UTILISATEUR
      if Bavard then
        if PgiAsk('Voulez-vous supprimer cette fiche dont le tiers ' + Tiers + ' est lié à la compta ou à la gestion interne ?', TitreHalley) = mrYes then
          Result := False;

      Msg := Msg + txt + #13 + #10;
      Exit;
    end;
  end;
end;


function TPerSup.TestJuriDosPer(var Msg: string): boolean;
// retourne True si on ne peut pas supprimer
var Q: TQuery;
  txt: string;
begin
  Result := FALSE;
 // on ne prend pas en compte la partie juridique venant uniquement du DP
  Q := OpenSQL('select JUR_NOMDOS from JURIDIQUE ' +
    'where JUR_GUIDPERDOS = "' + GuidPer + '" AND JUR_CODEDOS <> "&#@"', true);
  if (Q = nil) or (Q.EOF) then
  begin
    Ferme(Q);
    exit;
  end;
  Result := TRUE;
  txt := txt + 'Cette personne fait l''objet du dossier juridique '
    + Q.FindField('JUR_NOMDOS').AsString + #13 + #10;
  Ferme(Q);
  txt := txt + '=> Supprimer d''abord le dossier juridique pour pouvoir supprimer la fiche.';
  if Bavard then PGIInfo(txt, 'Suppression de ' + NomPer);
  Msg := Msg + txt + #13 + #10;
end;


function TPerSup.TestJuriLienPer(var Msg: string): Boolean;
// retourne True si on ne peut pas supprimer
var Q: TQuery;
  Select, txt: string;
begin
  Result := False;
  Select := 'select JUR_CODEDOS, JUR_NOMDOS from JURIDIQUE ' //JL 15/03/2005 Modifs jointure pour oracle (suppression INNER JOIN)
    + ', ANNULIEN where ANL_CODEDOS=JUR_CODEDOS AND ANL_GUIDPER="' + GuidPer
    + '" AND ANL_GUIDPER<>"" ' // permet de supprimer personne ayant guidper faux
    + 'AND ANL_CODEDOS<>"&#@" ' // => ces liens viennent du DP (déjà pris en compte dans TestLienPer)
    + 'group by JUR_CODEDOS, JUR_NOMDOS'; //MD 17/01/02
  Q := OpenSQL(Select, true, -1, '', True);
  if (Q = nil) or (Q.EOF) then begin Ferme(Q); exit; end; // SORTIE
  Result := True;
  txt := 'Cette personne intervient dans les liens du(des) dossier(s) juridique(s) :' + #13 + #10;
  while not Q.EOF do
  begin
    txt := txt + '- ' + Q.FindField('JUR_NOMDOS').AsString
      + ' (' + Q.FindField('JUR_CODEDOS').AsString + ')' + #13#10;
    Q.Next;
  end;
  Ferme(Q);
  txt := txt + '=> Veuillez modifier auparavant les dossiers cités.';
  if Bavard then PGIInfo(txt, 'Suppression de ' + NomPer);
  Msg := Msg + txt + #13 + #10;
end;


function TPerSup.TestJuriDosInfoPer(var Msg: string): boolean;
// retourne True si on ne peut pas supprimer
var Q: TQuery;
  Select, txt: string;
begin
  Result := False;
  Select := 'select JUR_NOMDOS from JURIDIQUE '
  //NCX 19/02/01 permet de supprimer personne ayant guidper à 0
  + 'left join JUDOSINFO on JDI_CODEDOS=JUR_CODEDOS where JDI_GUIDPER="' + GuidPer + '" and JDI_GUIDPER<>"" '
    + 'group by JUR_NOMDOS'; //MD 17/01/02
  Q := OpenSQL(Select, true, -1, '', True);
  if (Q = nil) or (Q.EOF) then begin Ferme(Q); exit; end; // SORTIE
  Result := True;
  txt := 'Des informations liées à cette personne sont enregistrées dans le(s) dossier(s) juridique(s):' + #13 + #10;
  while not Q.EOF do
  begin
    txt := txt + '- ' + Q.FindField('JUR_NOMDOS').AsString + #13 + #10;
    Q.Next;
  end;
  Ferme(Q);
  txt := txt + '=> Veuillez modifier auparavant les dossiers cités.';
  if Bavard then PGIInfo(txt, 'Suppression de ' + NomPer);
  Msg := Msg + txt + #13 + #10;
end;


function TPerSup.TestJuriGroupeSocPer(var Msg: string): boolean;
// retourne True si on ne peut pas supprimer
var Q: TQuery;
  txt: string;
begin
  Result := False;
  //NCX 19/02/01 permet de supprimer personne ayant guidper à 0
  Q := OpenSQL('select JGR_GRPSOC from JUGRPSOC where JGR_GUIDPER="' + GuidPer + '" AND JGR_GUIDPER<>"" '
    + 'group by JGR_GRPSOC', true, -1, '', True);
  if (Q = nil) or (Q.EOF) then begin Ferme(Q); exit; end; // SORTIE
  Result := True;
  txt := 'Cette personne est désignée comme société principale des groupes sociétés :' + #13 + #10;
  while not Q.EOF do
  begin
    txt := txt + '- ' + Q.FindField('JGR_GRPSOC').AsString + #13 + #10;
    Q.Next;
  end;
  Ferme(Q);
  txt := txt + '=> Veuillez modifier auparavant ce(s) groupe(s) de sociétés.';
  if Bavard then PGIInfo(txt, 'Suppression de ' + NomPer);
  Msg := Msg + txt + #13 + #10;
end;


function TPerSup.TestJuriEvenementPer(var Msg: string): boolean;
// retourne True si on ne peut pas supprimer
// Attention : accepte automatiquement la suppression d'évènements si bavard = False.
var Q: TQuery;
  txt, quest: string;
  nb: Integer;
begin
  Result := False;
  SupprEvents := False;
  //NCX 19/02/01 permet de supprimer personne ayant guidper à 0
  Q := OpenSQL('select count(JEV_GUIDEVT) from JUEVENEMENT where JEV_GUIDPER="'
    + GuidPer + '" and JEV_GUIDPER<>""', True);
  if (Q = nil) or (Q.Eof) or (Q.Fields[0].AsInteger = 0) then begin Ferme(Q); exit; end; // SORTIE
  Result := True;
  nb := Q.Fields[0].AsInteger;
  if nb = 1 then txt := IntToStr(nb) + ' évènement a été rattaché à cette personne.' + #13 + #10
  else txt := IntToStr(nb) + ' évènements ont été rattachés à cette personne.' + #13 + #10;
  Ferme(Q);
  // mode silencieux => suppression autom
  if not Bavard then
  begin
    if nb = 1 then txt := txt + '=> Il sera supprimé.' + #13 + #10
    else txt := txt + '=> Ils seront supprimés.' + #13 + #10;
    SupprEvents := True;
    Result := False;
  end
  // sinon question
  else
  begin
    quest := txt;
    if nb = 1 then quest := quest + 'Il ne pourra plus être recherché par ce critère.' + #13 + #10
    else quest := quest + 'Ils ne pourront plus être recherchés par ce critère.' + #13 + #10;
    quest := quest + 'Voulez-vous détruire également ces évènements ?';
    if PGIAsk(quest, 'Suppression de ' + NomPer) = mrYes then
    begin
      SupprEvents := True;
      if nb = 1 then txt := txt + '=> Il sera supprimé.' + #13 + #10
      else txt := txt + '=> Ils seront supprimés.' + #13 + #10;
      Result := False;
    end
    else
    begin
      txt := txt + '=> Supprimer d''abord les évènements pour supprimer la personne.' + #13 + #10;
    end;
  end;
  Msg := Msg + txt;
end;


procedure TPerSup.SupprimeEvents;
begin
  BeginTrans;
  try
    ExecuteSQL('delete from JUEVENEMENT where JEV_GUIDPER="' + GuidPer + '"');
  except
    V_PGI.IoError := oeUnknown;
  end;
  if V_PGI.IoError = oeUnknown then Rollback else CommitTrans;
end;



function TPerSup.TestLienPer(var Msg: string): Boolean;
// retourne True si des liens trouvés
var Q: TQuery;
  i: Integer;
  txt: string;
begin
  Result := False;
  i := 0;
  Q := OpenSQL('select ANN_NOM1, ANN_NOM2, JTF_FONCTABREGE from ANNULIEN, ANNUAIRE, JUTYPEFONCT '
    + 'where ANL_GUIDPERDOS=ANN_GUIDPER and ANL_GUIDPER="' + GuidPer
    + '" and ANL_FONCTION=JTF_FONCTION and ANL_GUIDPER<>"" group by ANN_NOM1, ANN_NOM2, JTF_FONCTABREGE'
    , true, -1, '', True);
  if (Q = nil) or (Q.EOF) then begin Ferme(Q); exit; end; // SORTIE
  Result := True;
  txt := 'Cette personne fait l''objet de liens dans le(s) dossier(s) :' + #13 + #10;
  while not Q.Eof do
  begin
    txt := txt + '- ' + Q.FindField('ANN_NOM1').AsString + ' ' + Q.FindField('ANN_NOM2').AsString
      + '(' + Q.FindField('JTF_FONCTABREGE').AsString + ')' + #13 + #10;
    Q.Next;
    Inc(i);
    // boucle limitée à 20 sinon message box énorme !
    if i > 20 then break;
  end;
  Ferme(Q);
  txt := txt + '=> Détruire d''abord les liens pour pouvoir détruire la fiche.';
  if Bavard then PGIInfo(txt, 'Suppression de ' + NomPer);
  Msg := Msg + txt + #13 + #10;
end;


function TPerSup.TestLienPlaq(var Msg: string): Boolean;
// retourne True si des liens trouvés
var txt: string;
begin
  Result := False;
  txt := '';
  // --------------- IFU ---------------

{ // MD 26/08/04 - Tables vidées en v6
  if ExistePlaquettes('P61_MILLESIME', 'F2561_PAYEUR', 'P61_GUIDPER',
    '- payeur dans le logiciel IFU (2561) pour le(s) millésime(s)', txt) then Result := True;
  if ExistePlaquettes('B61_MILLESIME', 'F2561_BENEF', 'B61_GUIDPERBENEF',
    '- bénéficiaire dans le logiciel IFU (2561) pour le(s) millésime(s)', txt) then Result := True;
  // remplacées par : }
 (* mcd 16/02/2006 pour GUID
 if ExistePlaquettes('L02_LIAMILLESIME', 'L02_ENTDECLA', 'L02_LIAGUIDPERPERE','IFU',
    '- payeur dans le logiciel IFU (2561) pour le(s) millésime(s)', txt) then Result := True;

  if ExistePlaquettes('L02_LIAMILLESIME', 'L02_ENTDECLA', 'L02_LIAGUIDPERFILS','IFU',
    '- bénéficiaire dans le logiciel IFU (2561) pour le(s) millésime(s)', txt) then Result := True;

  if ExistePlaquettes('L02_LIAMILLESIME', 'L02_ENTDECLA', 'L02_GUIDPER','CFE',
    '- personne dans le logiciel CFE ', txt) then Result := True; *)
  if ExistePlaquettes('L02_LIAMILLESIME', 'L02_ENTDECLA', 'L02_GUIDPERPERE', 'IFU',
    '- payeur dans le logiciel IFU (2561) pour le(s) millésime(s)', txt) then Result := True;

  if ExistePlaquettes('L02_LIAMILLESIME', 'L02_ENTDECLA', 'L02_GUIDPERFILS', 'IFU',
    '- bénéficiaire dans le logiciel IFU (2561) pour le(s) millésime(s)', txt) then Result := True;

  if ExistePlaquettes('L02_LIAMILLESIME', 'L02_ENTDECLA', 'L02_GUIDPER', 'CFE',
    '- personne dans le logiciel CFE ', txt) then Result := True;

{ Test impossible sur les autres plaquettes, car tables non communes, donc à tester
 dans les bases dossiers !
  // --------------- SCM ---------------
  if ExistePlaquettes('A36_MILLESIME', 'F2036_ASSOCIE', 'A36_GUIDPER',
    '- associé dans le logiciel SCM (2036) pour le(s) millésime(s)', txt) then Result := True;
  // --------------- SCI ---------------
  if ExistePlaquettes('A72_MILLESIME', 'F2072_ASSOCIE', 'A72_GUIDPER',
    '- associé dans le logiciel SCI (2072) pour le(s) millésime(s)', txt) then Result := True;
  if ExistePlaquettes('G72_MILLESIME', 'F2072_GERANT', 'G72_GUIDPER',
    '- gérant dans le logiciel SCI (2072) pour le(s) millésime(s)', txt) then Result := True;
  if ExistePlaquettes('J72_MILLESIME', 'F2072_JOUISSEUR', 'J72_GUIDPER',
    '- jouisseur dans le logiciel SCI (2072) pour le(s) millésime(s)', txt) then Result := True;
  if ExistePlaquettes('L72_MILLESIME', 'F2072_LOCATAIRE', 'L72_GUIDPER',
    '- locataire dans le logiciel SCI (2072) pour le(s) millésime(s)', txt) then Result := True;
  if ExistePlaquettes('M72_MILLESIME', 'F2072_LOCBES', 'M72_GUIDPER',
    '- locataire Loi Besson dans le logiciel SCI (2072) pour le(s) millésime(s)', txt) then Result := True;
  // --------------- BIF ---------------
  if ExistePlaquettes('B95_MILLESIME', 'F9065_MEMBRE', 'B95_GUIDPER',
    '- membre dans le logiciel BIF (9065) pour le(s) millésime(s)', txt) then Result := True;
}
  if Result then
  begin
    txt := 'Cette personne est utilisée en tant que :' + #13 + #10 + txt
      + '=> Effectuer d''abord la suppression dans le(s) logiciel(s) concerné(s).';
    if Bavard then PGIInfo(txt, 'Suppression de ' + NomPer);
    Msg := Msg + txt + #13 + #10;
  end;
end;


function TPerSup.TestDossierParti(var Msg: string): Boolean;
// retourne True si on ne peut pas supprimer
var txt: string;
begin
  Result := False;
  // MD 14/11/00 - pour dossier emportés sur portable
  if ExisteSQL('select DOS_VERROU from DOSSIER where DOS_GUIDPER="' + GuidPer + '" and DOS_VERROU="PAR"') then
  begin
    Result := True; // dossier parti
    txt := '=> Il existe un dossier parti basé sur cette fiche, il doit d''abord être réintégré.';
    if Bavard then PGIInfo(txt, 'Suppression de ' + NomPer);
    Msg := Msg + txt + #13 + #10;
  end;
end;


function TPerSup.TestFisPer(var Msg: string): Boolean;
var txt: string;
begin
  Result := False;
  txt := '';
  // #### Tests sur des tables non communes, donc pertinence uniquement
  // #### par rapport à la fiscalité perso ou CRE effectués sur la DB0
  if ExisteSQL('select 1 from FPDOSSIER where FPO_GUIDPER="' + GuidPer + '"') then
  begin
    txt := txt + '- dans un dossier de fiscalité personnelle' + #13 + #10;
    Result := True;
  end;
  if ExisteSQL('select 1 from FEMPRUNT where EMP_GUIDORGPRETEUR="' + GuidPer + '"') then
  begin
    txt := txt + '- en tant qu''organisme prêteur lié à un emprunt dans le logiciel CRE' + #13 + #10;
    Result := True;
  end;
  if Result then
  begin
    txt := 'Cette fiche est utilisée :' + #13 + #10 + txt
      + '=> Effectuer d''abord la suppression dans le(s) logiciel(s) concerné(s).';
    if Bavard then PGIInfo(txt, 'Suppression de ' + NomPer);
    Msg := Msg + txt + #13 + #10;
  end;
end;


constructor TPerSup.Create;
begin
  SupprEvents := False; // par défaut, on ne supprimera pas les évènements rattachés
end;

{$IFDEF BUREAU}
// cette fct permet de créer automatiquement un tiers à partir
// d'un annuaire, en calculant le code tiers

procedure DpCreerTiers(GuidPer: string; NatureAuxi, Nom, Nomper: string);
var Retour, X: Integer;
  CodeAux, CodeClient, Text, ZonesRetour, Auxi, Critere, Champ, Valeur: string;
begin
  CodeAux := '';
  CodeCLient := '';
 //pour nature de tiers CLi et PRO (qui peuvent exister en dossier,
 //utilisation de la fct générique pour calculer le n° de dossier
 //sinon, mise nature auxi + n° aléatoire sur 6c
  if (NatureAuxi = 'CLI') or (NatureAuxi = 'PRO') then
    Codeclient := NouveauNoDossier(GetParamSocSecur('SO_MDINCREMNODOSS', 'ALP'), Nom);
  if codeclient = '' then
  begin
    Randomize;
    Codeclient := Format('%.6d', [Random(1000000)]);
    if NatureAuxi = 'NCP' then Codeclient := 'AN' + Codeclient
    else Codeclient := NatureAuxi + Codeclient;
  end;
  Retour := 1;
  while Retour <> 0 do
  begin
    Retour := AfDos_Tiers(CodeClient, NatureAuxi, GuidPer, CodeAux);
    if (Retour <> 0) then
    begin
      case (Retour) of
      //--- Le tier existe déjà
        1: begin
            Text := 'LE TIERS EXISTE DEJA';
          end;
      //--- L'auxiliaire existe déjà
        2: begin
            Text := 'L''AUXILIAIRE EXISTE DEJA';
            Auxi := ';AUXI:' + codeAux;
          end;
      //--- Erreur général
        3: begin
            PgiInfo('Création du code client impossible, vérifier les renseignements ...', TitreHalley);
            Retour := 0; // pour sortir.. le tiers ne sera pas crée
          end;
      end;
      if retour <> 0 then
      begin //il faut demander le nouveau code client
        Zonesretour := DPLanceFiche_SaisieTiers('GUIDPER:' + GuidPer
          + ';TIERS:' + CodeClient + ';ERR:' + text + Auxi
          + ';NATUREAUXI:' + NatureAuxi
          + ';NOMPER:' + Nomper
          + ';NOM:' + Nom);
        if ZonesRetour = '' then retour := 0 // il a été fait annuler.. client non créer
        else begin
          Critere := (Trim(ReadTokenSt(ZonesRetour)));
          while (Critere <> '') do
          begin
            if Critere <> '' then
            begin
              X := pos('=', Critere);
              if x <> 0 then
              begin
                Champ := copy(Critere, 1, X - 1);
                Valeur := Copy(Critere, X + 1, length(Critere) - X);
              end;
              if Champ = 'TIERS' then CodeCLient := Valeur
              else if Champ = 'AUXI' then CodeAux := Valeur
            end;
            Critere := (Trim(ReadTokenSt(ZonesRetour)));
          end;
        end;
      end;
    end;
  end; //fin retour 0
end;
{$ENDIF}

function JaiLeDroitAdminDossier: boolean;
const ccDroitAdminDossier: integer = 26057;
begin
  Result := JaiLeDroitConceptBureau(ccDroitAdminDossier);
end;

function ArgumentToTabsheet(Argument: string): string; //LM20070528
var p: integer;
  s: string;
begin
  p := 1;
  while p <= length(Argument) do
    if argument[p] = ' ' then delete(argument, p, 1) else inc(p);

  p := pos('TABSHEET=', upperCase(Argument));
  if p > 0 then
  begin
    s := copy(Argument, p + length('TabSheet='), 100);
    result := trim(gtfs(s, ';', 1));
  end;
end;

// Indique si on est sur la fiche annuaire du dossier
// (= fiche annuaire lancée depuis la fiche synthèse)

function AnnuaireLanceeDepuisSynthese: Boolean;
{var FLog : Text;
    i : Integer;
    chemlog, msg : String;}
var
  i : integer;
const
  ArrDecla : array[0..5] of String = ('ANNUAIRE','ANNUAIRE_1','FICHINTERVENTION',
                                      'LIENANNUAIRE','ANNUAIRE_SEL','FICHELIEN');
begin
  Result := False;
{ // MD pour Debug pb KPMG de fiche Identité grisée
  // alors qu'on a les droits sur le concept "Modif fiche annuaire du dossier"
 DeleteFile('c:\test.txt');
  chemlog := 'c:\test.txt';
  OuvreLog(chemlog, FLog);
  for i:=0 to globalFichePile.Count-1 do
    begin
    msg := TForm(globalfichePile[i]).Name;
    Writeln(FLog, msg);
    Flush(FLog);
    end;
  FermeLog(FLog); }

  // Dans la fiche synthèse delphi (FSynthese), on a un panel pnlInfoDossier
  if globalFichePile.Count >= 2 then
    if TForm(globalfichePile[0]).Name = 'FSynthese' then
    begin
      // Fiche ANNUAIRE du dossier en inside dans le panel (bouton Identité)
      case globalFichePile.Count of
        2 : begin
              if (TForm(globalfichePile[1]).Name = 'ANNUAIRE') or
              {+GUH - 12/2007 FQ 11858}
              (TForm(globalfichePile[1]).Name = 'LIENINTERFILIALE') or
                (TForm(globalfichePile[1]).Name = 'FICHFISCALE') or
                  (TForm(globalfichePile[1]).Name = 'LIENTETEGROUPE') then
               {-GUH - 12/2007 FQ 11858}
                Result := True;
            end;
        // FQ 11765 - Fiche annuaire chargée depuis un hyperlien de la SYNTHDOSSIER du panel
        3 : begin
              if (TForm(globalfichePile[1]).Name = 'SYNTHDOSSIER') and ((TForm(globalfichePile[2]).Name = 'ANNUAIRE') or (TForm(globalfichePile[2]).Name = 'LIENINTERFILIALE')) then
                Result := True
              {+GUH - 12/2007 FQ 11858}
              else if (TForm(globalfichePile[1]).Name = 'ANNUAIRE') and (TForm(globalfichePile[2]).Name = 'LIENINTERVENANT') then
                Result := True
              else if (TForm(globalfichePile[1]).Name = 'FICHFISCALE') and ((TForm(globalfichePile[2]).Name = 'LIENINTERVENANT') or (TForm(globalfichePile[2]).Name = 'LIENTETEGROUPE') or (TForm(globalfichePile[2]).Name = 'LIENINTERCGE') or (TForm(globalfichePile[2]).Name = 'ANNUAIRE_SELLITE')) then
                Result := True
              else if (TForm(globalfichePile[1]).Name = 'FICHSOCIALE') and (TForm(globalfichePile[2]).Name = 'LIENINTERVENANT') then
                Result := True
              else if ((TForm(globalfichePile[1]).Name = 'LIENTETEGROUPE') or (TForm(globalfichePile[1]).Name = 'LIENINTERFILIALE')) and (TForm(globalfichePile[2]).Name = 'FICHINTERVENTION') then
                Result := True
              else if (TForm(globalfichePile[2]).Name = 'ANNUAIRE') then
                Result := True;
              {-GUH - 12/2007 FQ 11858}
            end;
        4 : begin
              for i:=Low(ArrDecla) to High(ArrDecla) do
              begin
                if ArrDecla[i] = TForm(globalfichePile[3]).Name then
                begin
                  Result := True;
                  break;
                end;
              end;
            end;
        5 : begin
              if (TForm(globalfichePile[4]).Name = 'ANNUAIRE') then
                Result := True;
            end;
      end;
    end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : SL
Créé le ...... : 10/03/2008
Modifié le ... : 20/03/2008
Description .. : Empêche les modifications de la fiche dans le DP
Suite ........ : sur le serveur quand le dossier est transporté
Suite ........ : (pour KPMG)
Mots clefs ... :
*****************************************************************}
procedure SetFicheLectureSeule(LaTom: TOM; Lecran: TForm; NoDoss: string);
{$IFDEF DP}
var
  Usr,Login:      String;
  SansBase:       Boolean;
{$ENDIF DP}
begin
{$IFDEF DP}
  //SL 17/06/2008 KPMG veut utiliser les portables en tant que serveurs et les SO_ENVPATHDAT commencent
  //par \\
  {PathDat := GetParamsocSecur('SO_ENVPATHDAT','');
  if (Copy(PathDat, 1, 2)='\\') //si on est sur une connexion réseau (il faut que le DP soit accessible en local)}
  if not (GetParamSocSecur('SO_MDCONFIGNOMADE', FALSE)=TRUE) //si on est sur un serveur
  and (EtatMarqueDossier(NoDoss,Usr,Login,SansBase)='PAR') //Si le dossier est transporté
  and VH_DP.SeriaKPMG then //Pour l'instant, only KPPMG
  begin
    //LaTom.SetControlEnabled('BVALIDER',False);    //###Du coup le paramètre LaTom est inutile!
    TFFiche(Lecran).fTypeAction := taConsult;
    if Lecran.Name='YYCONTACT' then
    begin
      LaTom.SetControlEnabled('BACTION',False);
      LaTom.SetControlEnabled('BCOMPLGRC',False);
      LaTom.SetControlEnabled('BCONTACT',False);
      LaTom.SetControlEnabled('BALLCONTACT',False);
    end
    else if Lecran.Name='ANNUAIRE' then
    begin
      LaTOM.SetControlEnabled('BFORME',False);
    end;
  end;
{$ENDIF DP}
end;

initialization
  RegisterAglProc('SupprimeListAnnu', TRUE, 1, AGLSupprimeListAnnu);
  RegisterAglProc('SupprimeEnregAnnu', FALSE, 1, AGLSupprimeEnregAnnu);
  RegisterAglProc('DupliquerElementAnn', FALSE, 4, AGLDupliquerElementAnn);
  RegisterAglProc('FermeEtOuvre', TRUE, 5, AGLFermeEtOuvre);

finalization

end.

