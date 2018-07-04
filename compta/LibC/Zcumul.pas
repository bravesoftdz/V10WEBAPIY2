unit ZCumul;

interface

uses
  { Delphi }
  SysUtils
  , classes
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
{$ENDIF}
  { AGL }
  , uTob
  , hctrls
  , HEnt1
  , HashTable
  { Compta }
  , Ent1
  , uLibExercice
  , CalcOLE
  , CpteUtil
  {$IFDEF MODENT1}
  , CPTypeCons
  {$ENDIF MODENT1}
  ,UentCommun  
  ;

type
  TModeFoncCumul = ( mfMemory, mfQuery );

  TDomaineCumul = ( dcEcriture, dcAnalytique, dcBalSit );

  TTypeCorresp = (coGeneral, coCorresp1, coCorresp2 );

  TTypeAgregat = (taRubrique, taFourchette, taDivers);

  TZCumul = class
  private
    fRubrique     : TOB;
    fIndex        : integer;
    fCurRubrique  : TOB;
    fTypeCumul    : string;
    fCorresp      : TTypeCorresp;

    { Critères }
    fDevise: string;
    fEtablissement: string;
    fExercice: string;
    fDateDebut: TDateTime;
    fDateFin: TDateTime;
    fBalSit : string;
    fAxe: string;
    fSection1: string;
    fSection2: string;
    fComplementTL : string;
    fRegroupement : string;
    fTypeEcr: SetttTypePiece;
    fTypeSolde : string;
    fAvecANO : string;
    fDateCalculRubrique : TDateTime;
    fpStTypeEcr : string;

    { Propriétés }
    fLastError              : integer;        // Dernier code erreur renvoyé
    fLastErrorMsg           : string;         // Dernier message d'erreur renvoyé
//    fbRubrique              : boolean;        // Indique si on travaille sur des rubriques
    fTypeAgregat            : TTypeAgregat;   // Indique le type d'agrégat sur lequel on travaille
    fHashTable              : THashTable;     // Objet de stockage des cumuls en mémoire
    fRacineCleHash          : string;         // identificateur de stockage des cumuls
    fModeFonc               : TModeFoncCumul; // Mode de fonctionnement du calcul du cumul
    fbCompensation          : boolean;        // Compensation des auxiliaires
    fbAvecLibelle           : boolean;        // Renvoi du libellé dans la ligne détail
    fTLesComptes            : TOB;            // Liste des comptes du dossier chargée en mémoire
    fTLesComptesRub         : TOB;            // Liste des comptes de la rubrique

    function    GetRubrique(stRubrique: string): integer;
    function    FaitRequeteDonnees(stType: string; stSQLCompte: string; pbAuxiliaire : boolean ): string;
    procedure   QuelTypEcr( St : string );
    function    PositionneCumuls(pTotalDebit, pTotalCredit: double; stCompte:
                    string; var TResult: TabloExt; bIsCollectif : boolean): double;
    procedure   RecupereInfoChamps( stType: string; var stPrefDonnee : string; var stPrefCode : string; var stCompte1 : string; var stCompte2 : string);
    function    TableToCodeTable ( St : string ) : string;
    function    FaitRequeteParametres(pstTable, pstQuoi: string ): string;
    procedure   SetModeFonc(const Value: TModeFoncCumul);
    procedure   ChargeCumulMemoire(pstType: string);
    procedure   ChargeLesComptesMemoire ( pstSQL : string );
    procedure   AjouteLigneDetail(pLeDetail: TStringList; pT : TOB ; pTResult: TabloExt);
    function    GetValeurUnCompte( pT: TOB; pLeDetail : TStringList; var pTResult : TabloExt; bIsCollectif : boolean ): double;
    function    GetValeurUnCompteCollectif(pT : TOB; pLeDetail: TStringList ; var pTResult : TabloExt ): double;
    procedure   FaitListeDesComptesRubrique(pLInclus, pLExclus: TStringList);
    procedure   FourchetteVersRubrique(pstFourchette : string);
  public
    constructor Create;
    destructor  Destroy; override;
    function    InitCriteres(pstTypeEcr,pstEtablissement, pstDevise, pstDate : string;
                  pstBalsit: string = ''; pDateCalculRubrique : TDateTime = 0 ): boolean;
    function    InitAnalytique (pstAxe, pstSection1,  pstSection2 : string) : boolean;
    function    GetValeur(stType : string; stQuoi1 : string; LeDetail : TStringList) : variant;
    function    ReinitCumul ( bAll : boolean ) : boolean;
    { Critères }
    property    Regroupement  : string read fRegroupement write fRegroupement;
    { Gestion des erreurs }
    property    LastErrorMsg  : string read fLastErrorMsg;
    property    LastError     : integer read fLastError;
    { Mode de fonctionnement }
    property    ModeFonc      : TModeFoncCumul read fModeFonc write SetModeFonc;
    property    AvecLibelle   : boolean read fbAvecLibelle write fbAvecLibelle;
    property    Corresp       : TTypeCorresp read fCorresp write fCorresp;

    // $$$ JP 21/12/05 - j'ai besoin des dates et du type, et du code exo/code balance
    property    dDateDebut    : TDateTime read fDateDebut;
    property    dDateFin      : TDateTime read fDateFin;
    property    strExercice   : string    read fExercice;
    property    strSituation  : string    read fBalSit;
    property    ComplementTL  : string    read fComplementTL write fComplementTL;
  end;

function KPMGVersCegid ( pstFourchette : string ) : string;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  ParamSoc;


{ TZCumul }

const
  TAILLE_CACHE = 1000;
//  BOURRE_COMPTE_SUP = '99999999999999999';
//  BOURRE_COMPTE_INF = '00000000000000000';
  BOURRE_COMPTE_INF = '!!!!!!!!!!!!!!!!!';
  BOURRE_COMPTE_SUP = '}}}}}}}}}}}}}}}}}';

{***********A.G.L.***********************************************
Auteur  ...... : Christophe AYEL
Créé le ...... : 29/10/2007
Modifié le ... :   /  /
Description .. : Transforme les fourchettes au format (KPMG) en format
Suite ........ : lisible par les applications Cegid soit :
Suite ........ : - Remplacement des "-" séparateurs de tranches en ":"
Mots clefs ... :
*****************************************************************}
function KPMGVersCegid ( pstFourchette : string ) : string;
begin
  Result := StringReplace(pstFourchette, '-', ':', [rfReplaceAll,rfIgnoreCase]);
end;

constructor TZCumul.Create;
begin
  fModeFonc := mfQuery;
//  fbRubrique := False;
  fTypeAgregat := taDivers;
  fLastError := 0;
  fLastErrorMsg := '';
  fDevise := '';
  fAxe := '';
  fbCompensation := False;
  fDateCalculRubrique := iDate2099;
  fRubrique := TOB.Create('LESRUBRIQUES', nil, -1);
  fHashTable := THashTable.Create();
  fHashTable.useElfHash := true;
  fHashTable.singleThread := true;
  fTLesComptes := TOB.Create ('',nil,-1);
  fTLesComptesRub := TOB.Create ('',nil,-1);
  // Chargement du plan de correspondance
  if GetParamSocSecur('SO_CPLIASSESURCORRESP',False) then
  begin
    if GetParamSocSecur('SO_CPLIASSEPLANCORRESP','')='1' then Corresp := coCorresp1
    else if GetParamSocSecur('SO_CPLIASSEPLANCORRESP','')='2' then Corresp := coCorresp2
    else Corresp := coGeneral;
  end;
end;

destructor TZCumul.Destroy;
begin
  fTLesComptes.Free;
  fTLesComptesRub.Free;
  fHashTable.Clear(True);

  // $$$ JP 21/12/05 - faut virer c'te hache table!
  FreeAndNil (fHashTable);

  if fRubrique <> nil then
    FreeAndNil(fRubrique);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 19/12/2005
Modifié le ... :   /  /
Description .. : Construit la requête qui donne la liste des paramètres à
Suite ........ : traiter (général, auxiliaire, section ...)
Mots clefs ... :
*****************************************************************}
function TZCumul.FaitRequeteParametres ( pstTable : string;  pstQuoi : string ): string;
var
  lstSQLParametre,lstWhereCompte : string;
  lstExclu : string;
  Lefb1, LeFb2 : TFichierBase;
  bOnBudget : boolean;
  lstNomChamp, LOrder : string;
  lLesComptesInclus, lLesComptesExclus : TStringList;
  lstWhere  : string;
  stCompte : string;
begin
  bOnBudget := False;
  lstSQLParametre := '';
  lLesComptesInclus := nil;
  lLesComptesExclus := nil;
  if fModeFonc = mfMemory then
  begin
    lstNomChamp := 'COMPTE1';
    lLesComptesInclus := TStringList.Create;
    lLesComptesInclus.Duplicates := dupIgnore;
    lLesComptesInclus.Sorted := True;
    lLesComptesExclus := TStringList.Create;
    lLesComptesExclus.Duplicates := dupIgnore;
    lLesComptesExclus.Sorted := True;
  end
  else lstNomChamp := '';
  { Cumul sur la table RUBRIQUE }
//  if fbRubrique then
  if ((fTypeAgregat = taRubrique) or (fTypeAgregat = taFourchette)) then
  begin
    if (( GetRubrique(pstQuoi) >= 0 ) or (fTypeAgregat = taFourchette))  then
    begin
      // Dans le cas d'une fourchette, on calcule les informations de la rubrique équivalente
      if (fTypeAgregat = taFourchette) then
        FourchetteVersRubrique ( KPMGVersCegid (pstQuoi) );  // On applique la fonction KPMGVersCegid pour les fourchettes du type xxx-yyy
      if lLesComptesInclus <> nil then lLesComptesInclus.Clear;
      if lLesComptesExclus <> nil then lLesComptesExclus.Clear;
      { Liste des comptes concernés par la rubrique }
      if fModeFonc = mfMemory then
      begin
        lstSQLParametre := 'SELECT G_GENERAL '+lStNomChamp+', G_LIBELLE LIBELLE';
        if (fAxe = '' ) then lstSQLParametre := lstSQLParametre + ', G_COLLECTIF ';
      end else
      begin
        { CA - 17/04/2007 - En mode Query, on travaille sur un seul champs : WHERE GENERAL IN (SELECT GENERAL FROM GENERAUX WHERE) }
        lstSQLParametre := 'SELECT G_GENERAL '+lStNomChamp+' ';
      end;
      { Si on travaille avec les comptes de correspondances, on les ajoute dans la requête }
      if (Corresp = coCorresp1) then lstSQLParametre := lstSQLParametre + ', IIF(G_CORRESP1="",G_GENERAL,G_CORRESP1) CORRESP1'
      else if (Corresp = coCorresp2) then lstSQLParametre := lstSQLParametre + ', IIF(G_CORRESP2="",G_GENERAL,G_CORRESP2) CORRESP2 ';
      lstSQLParametre := lstSQLParametre +  ' FROM GENERAUX ';
      { Pour avoir le code section le cas échéant }
      if (( fModeFonc = mfMemory ) and ( fAxe <> '')) then
      begin
        if fComplementTL = '' then
          lstSQLParametre := 'SELECT Y_GENERAL COMPTE1,Y_SECTION,G_LIBELLE LIBELLE,G_COLLECTIF FROM ANALYTIQ '
            +'LEFT JOIN GENERAUX ON G_GENERAL=Y_GENERAL WHERE Y_GENERAL IN ('
            +' SELECT G_GENERAL FROM GENERAUX )  GROUP BY Y_GENERAL,Y_SECTION,G_LIBELLE,G_COLLECTIF'
            +' ORDER BY Y_GENERAL,Y_SECTION'
        else
          lstSQLParametre := 'SELECT Y_GENERAL COMPTE1,Y_SECTION,G_LIBELLE LIBELLE,G_COLLECTIF FROM ANALYTIQ '
            +' LEFT JOIN GENERAUX ON G_GENERAL=Y_GENERAL'
            +' LEFT JOIN SECTION ON S_SECTION=Y_SECTION'
            +' WHERE Y_GENERAL IN ('
            +' SELECT G_GENERAL FROM GENERAUX )'
            +' AND ' + fComplementTL
            +' GROUP BY Y_GENERAL,Y_SECTION,G_LIBELLE,G_COLLECTIF'
            +' ORDER BY Y_GENERAL,Y_SECTION';
      end
      else
      begin
        if (Corresp = coCorresp1) then LOrder := ' ORDER BY CORRESP1,G_GENERAL'
        else if (Corresp = coCorresp2) then LOrder := ' ORDER BY CORRESP2,G_GENERAL'
        else LOrder := ' ORDER BY G_GENERAL';
      end;
      { Chargement des comptes en mémoire }
      if fAxe <> '' then ChargeLesComptesMemoire (lstSQLParametre)
      else ChargeLesComptesMemoire (lstSQLParametre+LOrder);
      { Récupération des comptes de la rubrique }
      lstWhere := AnalyseCompte(fCurRubrique.GetValue('RB_COMPTE1'), fbGene, False, False, True, True, lLesComptesInclus);
      if (Corresp = coCorresp1) or (Corresp = coCorresp2) then
      begin
        if Corresp = coCorresp1 then stCompte := 'G_CORRESP1' else stCompte := 'G_CORRESP2';
        lstWhere := '(' + lstWhere + ' AND '+stCompte+'="") OR ('+StringReplace(lstWhere,'G_GENERAL',stCompte,[rfReplaceAll,rfIgnoreCase])+')';
      end;
      lstSQLParametre := lstSQLParametre + ' WHERE '+ lstWhere;
      lstExclu := AnalyseCompte(fCurRubrique.GetValue('RB_EXCLUSION1'), fbGene, True, False, True,True,lLesComptesExclus);
      if lstExclu <> '' then
      begin
        if (Corresp = coCorresp1) or (Corresp = coCorresp2) then
        begin
          if Corresp = coCorresp1 then stCompte := 'G_CORRESP1' else stCompte := 'G_CORRESP2';
          lstExclu := '(' + lstExclu + ' AND '+stCompte+'="") OR ('+StringReplace(lstExclu,'G_GENERAL',stCompte,[rfReplaceAll,rfIgnoreCase])+')';
        end;
        lstSQLParametre := lstSQLParametre + ' AND ' + lstExclu;
      end;
      { CA - 17/04/2007 - On n'ajoute pas le ORDER BY en mode QUERY car SELECT ... IN SELECT... ORDER BY ne fonctionne pas }
      if fModeFonc = mfMemory then
        lstSQLParametre := lstSQLParametre + LOrder;

      if ( fModeFonc = mfMemory ) then FaitListeDesComptesRubrique ( lLesComptesInclus, lLesComptesExclus);
    end else
    begin
      fLastError := -1;
      fLastErrorMsg := TraduireMemoire ('Rubrique')+' '+pstQuoi+' '+TraduireMemoire('introuvable');
    end;
  { Cumul sur une Table }
  end else begin
    { Récupération du fb }
    if (QuelFb(TableToCodeTable(pstTable),pstQuoi,'',Lefb1,Lefb2,fAxe,bOnBudget) = 0) then
    begin
      fLastError := -1;
      fLastErrorMsg := TraduireMemoire ('Problème de type');
    end
    else
    begin
      lstWhereCompte := AnalyseCompte(pstQuoi, LeFb1, False, False);
      { Cumul sur la table GENERAUX }
      if pstTable = 'GENERAUX' then lstSQLParametre := 'SELECT G_GENERAL '+lStNomChamp+{',G_LIBELLE LIBELLE}' FROM GENERAUX WHERE '+lstWhereCompte
      {YMO 20/06/2006 FQ17896 Cette expression est évaluée dans un 'IN (', elle ne peut ramener que le champ G_GENERAL, T_AUXILIAIRE ou S_SECTION
      { Cumul sur la table TIERS }
      else if pstTable = 'TIERS' then lstSQLParametre := 'SELECT T_AUXILIAIRE '+lStNomChamp+{',T_LIBELLE LIBELLE}' FROM TIERS WHERE '+lstWhereCompte
      { Cumul sur la table SECTION }
      else if pstTable = 'SECTION' then lstSQLParametre := 'SELECT S_SECTION '+lStNomChamp+{',S_LIBELLE LIBELLE}' FROM SECTION WHERE '+lstWhereCompte
      else
      begin
        fLastError := -1;
        fLastErrorMsg := TraduireMemoire ('Non défini');
      end;
    end;


  end;
  if (lLesComptesInclus <> nil) then FreeAndNil (lLesComptesInclus);
  if (lLesComptesExclus <> nil) then FreeAndNil (lLesComptesExclus);
  Result := lstSQLParametre;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/05/2005
Modifié le ... : 03/11/2005
Description .. : Constitution de la requête de calcul des cumuls sur la table
Suite ........ : écriture. Cette requête est directement inspirée du
Suite ........ : FabricReqCpt3 de cpteutil.pas
Suite ........ : - FQ 16989 - CA - 03/11/2005 - Prise en compte des
Suite ........ : caractères Joker sur la section
Mots clefs ... :
*****************************************************************}

function TZCumul.FaitRequeteDonnees(stType : string; stSQLCompte: string; pbAuxiliaire : boolean ): string;

  function _ChampJointureCode ( stPrefixe : string ) : string;
  begin
    if stPrefixe = 'G' then Result := 'G_GENERAL'
    else if stPrefixe = 'T' then Result := 'T_AUXILIAIRE'
    else if stPrefixe = 'S' then Result := 'S_SECTION'
    else Result := '';
  end;

var
  stSQL: string;
  stDev: string;
  stCompte1, stCompte2 : string;
  stPrefCode, stPrefDonnee : string;
begin
  { Récupération des champs à exploiter }
  RecupereInfoChamps ( stType, stPrefDonnee, stPrefCode, stCompte1, stCompte2 );

  { SELECT }
  stSQL := 'SELECT ' + stPrefDonnee+stCompte1 + ' COMPTE1 '+','+stPrefCode+'_LIBELLE LIBELLE ';

  if pbAuxiliaire then stSQL := stSQL + ' , '+stPrefDonnee+stCompte2 + ' COMPTE2 ';
(*
  if stCompte2 <> '' then stSQL := stSQL + ' , ' + stPrefDonnee + stCompte2 + ' COMPTE2 '
  else stSQL := stSQL + ' , "-" COMPTE2 ';  // indispensable pour chargement en mémoire
*)
  if stCompte1 = 'GENERAL' then stSQL := stSQL + ',G_COLLECTIF ';

  if (fDevise <> '') and (fDevise <> V_PGI.DevisePivot) then
    stDev := 'DEV'
  else
    stDev := '';
  stSQL := stSQL + ', SUM (' + stPrefDonnee + '_DEBIT' + stDev + ') TOTDEBIT, SUM (' +
    stPrefDonnee + '_CREDIT' + stDev + ') TOTCREDIT ';
  if fAxe <> '' then
  begin
    stSQL := stSQL + ' , Y_AXE ';
    if stType <> 'SECTION' then // FQ 16801 SBO 11/10/2005 : Pb requête si cumuls sur Section (y_section doublé)
      stSQL := stSQL + ' , Y_SECTION ';
  end;

  stSQL := stSQL + ' FROM ' + PrefixeToTable(stPrefDonnee);

  { JOINTURE sur la table des codes : TIERS, GENERAUX, SECTIONS }
  stSQL := stSQL + ' LEFT JOIN ' + PrefixeToTable(stPrefCode) + ' ON ('+ _ChampJointureCode ( stPrefCode ) +'=' + stPrefDonnee + stCompte1 + ') ';

  { WHERE }
  stSQL := stSQL + ' WHERE ';

  if fBalsit = '' then
  begin
    { Exercice }
    if fExercice <> '' then
      stSQL := stSQL + ' AND ' + stPrefDonnee + '_EXERCICE = "' + fExercice + '" ';

    { Date comptable }
    if fDateDebut <> iDate1900 then
      stSQL := stSQL + ' AND ' + stPrefDonnee + '_DATECOMPTABLE>="' +
        UsDateTime(fDateDebut) + '" ';

    if fDateFin <> iDate1900 then
      stSQL := stSQL + ' AND ' + stPrefDonnee + '_DATECOMPTABLE<="' +
        UsDateTime(fDateFin) + '" ';

    { Devise }
    if fDevise <> '' then
      stSQL := stSQL + ' AND ' + stPrefDonnee + '_DEVISE = "' + fDevise + '" ';

    { Etablissement }
    if fEtablissement <> '' then
      stSQL := stSQL + ' AND ' + stPrefDonnee + '_ETABLISSEMENT = "' + fEtablissement +
        '" ';

    { Prise en compte des à-nouveaux }
    if fAvecANO = 'SAN' then stSQL := stSQL + ' AND ' + stPrefDonnee + '_ECRANOUVEAU = "N"'
    else if fAvecANO = 'ANO' then stSQL := stSQL + ' AND (' + stPrefDonnee + '_ECRANOUVEAU = "OAN" OR '+
                                                         stPrefDonnee + '_ECRANOUVEAU = "H"';

    { Type d'écriture }
    stSQL := stSQL + WhereSupp(stPrefDonnee+'_', fTypeEcr);
  end else stSQL := stSQL + ' AND BSE_CODEBAL="'+fBalSit+'"';

  if fAxe <> '' then
  begin
    stSQL := stSQL + ' AND Y_AXE="' + fAxe + '" ';
    if EstJoker ( fSection1 ) then  // FQ 16989 - CA -03/11/2005
    begin
      stSQL := stSQL + ' AND Y_SECTION LIKE "' + TraduitJoker(fSection1) + '" '
    end else
    begin
      if fSection1 <> '' then
        stSQL := stSQL + ' AND Y_SECTION>="' + fSection1 + '" ';
      if fSection2 <> '' then
        stSQL := stSQL + ' AND Y_SECTION<="' + fSection2 + '" ';
    end;
  end;

  if fModeFonc = mfQuery then
    stSQL := stSQL + ' AND ' + stPrefDonnee + stCompte1 + ' IN (' + stSQLCompte + ')';

  { GROUP BY }
  stSQL := stSQL + ' GROUP BY ' + stPrefDonnee + stCompte1 ;
  stSQL := stSQL + ',' + stPrefCode + '_LIBELLE' ;
  if stCompte1 = 'GENERAL' then stSQL := stSQL + ',G_COLLECTIF ';
  // if  ( stCompte2 <> '' ) then stSQL := stSQL + ',' + stPrefDonnee + stCompte2 + ' ';
  if pbAuxiliaire then stSQL := stSQL + ' , '+stPrefDonnee+stCompte2;
  if fAxe <> '' then
  begin
    stSQL := stSQL + ' , Y_AXE';
    if stType <> 'SECTION' then // FQ 16801 SBO 11/10/2005 : Pb requête si cumuls sur Section (y_section doublé dans le group by)
      stSQL := stSQL + ' , Y_SECTION ';
  end;

  { ORDER BY }
  stSQL := stSQL + ' ORDER BY ' + stPrefDonnee + stCompte1;
  if stCompte1 = 'GENERAL' then stSQL := stSQL + ',G_COLLECTIF ';
//  if  ( stCompte2 <> '' ) then stSQL := stSQL + ',' + stPrefDonnee + stCompte2 + ' ';
  if pbAuxiliaire then stSQL := stSQL + ' , '+stPrefDonnee+stCompte2;
  if fAxe <> '' then
  begin
    stSQL := stSQL + ' , Y_AXE';
    if stType <> 'SECTION' then // FQ 16801 SBO 11/10/2005 : Pb requête si cumuls sur Section (y_section doublé dans le order by)
      stSQL := stSQL + ' , Y_SECTION ';
  end;

  Result := stSQL;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/05/2005
Modifié le ... :   /  /
Description .. : Chargement des informations d'une rubrique
Mots clefs ... :
*****************************************************************}

function TZCumul.GetRubrique(stRubrique: string): integer;
var
  i: integer;
  Q: TQuery;
  lStColonne: string;
  stWhereDateCalcul : string;
begin
  Result := -1;
  fIndex := -1;
  fCurRubrique := nil;

  { On commence par chercher la rubrique en mémoire }
  for i := 0 to fRubrique.Detail.Count - 1 do
  begin
    if (fRubrique.Detail[i].GetValue('RB_RUBRIQUE') = stRubrique) then
    begin
      Result := i;
      fIndex := i;
      fCurRubrique := fRubrique.Detail[i];
      break;
    end;
  end;
  { Rubrique non trouvée en mémoire }
  if (Result = -1) then
  begin
    stWhereDateCalcul := ' AND RB_DATEVALIDITE>="'+USDateTime(fDateCalculRubrique)+'"';
    lStColonne := ' RB_RUBRIQUE, RB_COMPTE1, RB_EXCLUSION1, RB_SIGNERUB, RB_PREDEFINI, RB_DATEVALIDITE,RB_TYPERUB ';
    Q := OpenSQL('SELECT '+lStColonne +
      ' FROM RUBRIQUE WHERE RB_RUBRIQUE>="' + stRubrique +
      '" AND (RB_PREDEFINI<>"DOS" OR (RB_PREDEFINI="DOS" AND RB_NODOSSIER="'+V_PGI.NoDossier+'"))'+
      stWhereDateCalcul+' ORDER BY RB_RUBRIQUE, RB_NODOSSIER DESC ,RB_PREDEFINI DESC,RB_DATEVALIDITE ASC ', True);
    try
      if not Q.Eof then
      begin
        if fRubrique.LoadDetailDB('RUBRIQUE', '', '', Q, False) then
        begin
          if (fRubrique.Detail[0].GetValue('RB_RUBRIQUE') = stRubrique) then
          begin
            fCurRubrique := fRubrique.Detail[0];
            Result := 0;
            fIndex := 0;
          end;
        end;
      end;
    finally
      Ferme(Q);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 20/12/2005
Modifié le ... :   /  /
Description .. : Mise à jour de la ligne de détail
Mots clefs ... :
*****************************************************************}
procedure TZCumul.AjouteLigneDetail ( pLeDetail : TStringList; pT : TOB;  pTResult : TabloExt );
var
  lSt : string;
  i : integer;
begin
  lSt := pT.GetString('COMPTE1');
  if (fAxe <> '') then
    lSt := lSt + ';' + pT.GetString( 'Y_SECTION');

  for i := 1 to 6 do
    lSt := lSt + ':' + FloatToStr(pTResult[i]);
  lSt := lSt + ':' + fCurRubrique.GetValue('RB_SIGNERUB');
  if fbAvecLibelle then lSt := lSt +#9+pT.GetString('LIBELLE');
  pLeDetail.Add(lSt);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 20/12/2005
Modifié le ... :   /  /
Description .. : Calcul du cumul pour un compte
Mots clefs ... :
*****************************************************************}
function TZCumul.GetValeurUnCompte ( pT : TOB ; pLeDetail : TStringList; var pTResult : TabloExt; bIsCollectif : boolean ) : double ;
var
  dValCompte : double;
  lT : TOB;
  i : integer;
  lStCle : string;
  lStCompte : string;
begin
  if Corresp=coCorresp1 then lStCompte := 'CORRESP1'
  else if Corresp=coCorresp2 then lStCompte := 'CORRESP2'
  else lStCompte := 'COMPTE1';

  if (fModeFonc = mfQuery) then
    dValCompte := PositionneCumuls(pT.GetDouble('TOTDEBIT'), pT.GetDouble('TOTCREDIT'),
              pT.GetString(lStCompte), pTResult, bIsCollectif)
  else
  begin
  { ECR - 26/12/2007 : Construction de la clé de hashage : le séparateur des différents élements
  est le caractère ":" pour pouvoir utiliser la fonction RemoveStack sur une
  racine }
//    lStCle := fRacineCleHash+';'+pT.GetString('COMPTE1');
    lStCle := fRacineCleHash+':'+pT.GetString('COMPTE1');
//    if (fAxe <> '') then lStCle := lStCle + ';' + pT.GetString('Y_SECTION');      {En ANAL2, le champ s'appelle tj Y_SECTION}
    if (fAxe <> '') then lStCle := lStCle + ':' + pT.GetString('Y_SECTION');      {En ANAL2, le champ s'appelle tj Y_SECTION}
    lT := TOB(fHashTable.Get(lStCle));
    if lT <> nil then
      dValCompte := PositionneCumuls(lT.GetValue('TOTDEBIT'), lT.GetValue('TOTCREDIT'),
                pT.GetString(lStCompte), pTResult, bIsCollectif)
    else
    begin
      for i := 1 to 6 do pTResult[i]:=0;
      dValCompte := 0;
    end;
  end;

  Result := dValCompte;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 20/12/2005
Modifié le ... :   /  /
Description .. : Calcul du cumul pour un compte collectif dans le cas de la
Suite ........ : non compensation
Mots clefs ... :
*****************************************************************}
function TZCumul.GetValeurUnCompteCollectif ( pT : TOB ; pLeDetail : TStringList ; var pTResult : TabloExt ) : double ;
var
  lModeFoncTmp  : TModeFoncCumul;
  lSt           : string;
  ldValCompte   : double;
  lTResult      : TabloExt;
  iDet, i       : integer;
  lTColl        : TOB;
begin
  Result := 0;
  { On fait une requête sur la base, donc on force le mode mfQuery }
  lModeFoncTmp := fModeFonc;
  fModeFonc := mfQuery;
  try
    lSt := FaitRequeteDonnees('RUBRIQUE', '"'+pT.GetString('COMPTE1')+'"', True);
    lTColl := TOB.Create ('', nil, -1);
    try
      for i := 1 to 6 do pTResult[i]:=0;
      lTColl.LoadDetailFromSQL (lSt);
      if (lTColl.Detail.Count  >  0) then
      begin
        for iDet := 0 to lTColl.Detail.Count - 1 do
        begin
          ldValCompte := GetValeurUnCompte ( lTColl.Detail[iDet], pLeDetail, lTResult, True );
          for i := 1 to 5 do pTResult[i]:=pTResult[i]+lTResult[i];
          Result := Arrondi( Result + ldValCompte , V_PGI.OkDecV);
        end;
        pTResult[6] := lTResult[6];
      end;
    finally
      lTColl.Free;
    end;
  finally
    fModeFonc := lModeFoncTmp;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/05/2005
Modifié le ... :   /  /
Description .. : Renvoie la valeur d'une rubrique.
Mots clefs ... :
*****************************************************************}

function TZCumul.GetValeur(stType : string; stQuoi1 : string; LeDetail : TStringList):
  variant;
var
  stSQL, stSQLCompte : string;
  Q : TQuery;
  stErreur : string;
  lT : TOB;
  i : integer;
  ldValCompte : double;
  lTResult : TabloExt;
begin
  stErreur      := '';
  Result        := 0;
  fTypeCumul    := stType;
  fLastError    := 0;
  fLastErrorMsg := '';

//  fbRubrique := (stType = 'RUBRIQUE');
  if (stType = 'RUBRIQUE') then fTypeAgregat := taRubrique
  else if (stType = 'FOURCHETTE') then fTypeAgregat := taFourchette
  else fTypeAgregat := taDivers;

  if ( fModeFonc = mfMemory ) then
  begin
    { Chargement des cumuls en mémoire }
    lT := TOB(fHashTable.Get(fRacineCleHash));
    if lT = nil then ChargeCumulMemoire (stType);
  end;

  { Récupération de la requête qui donne la liste des éléments à traiter : comptes généraux, sections ... }
  stSQLCompte := FaitRequeteParametres(stType, stQuoi1 );

  { si la rubrique n'est pas de type 'GEN', on utilise le GetCumul pour
    calculer la valeur de la rubrique }
  if (( fCurRubrique <> nil ) and (fCurRubrique.GetString('RB_TYPERUB')<>'GEN'))  then
  begin
    // FQ 19749 - Si la rubrique n'est pas de type 'GEN', il faut passer le regroupement au GetCumul
    if fRegroupement <> '' then
    Result := GetCumulMS('RUB',stQuoi1,'',fpStTypeEcr,fEtablissement,fDevise,fExercice,
        fDateDebut,fDateFin,False, LeDetail<>nil, LeDetail,lTResult,False,'','','',False,fBalsit,'','',fRegroupement)
    else Result := GetCumul('RUB',stQuoi1,'',fpStTypeEcr,fEtablissement,fDevise,fExercice,
        fDateDebut,fDateFin,False, LeDetail<>nil, LeDetail,lTResult,False,'','','',False,fBalsit);
    exit;
  end;

  if (LastError = 0) then
  begin
    { Constitution de la requête sur la table des données }
    if (fModeFonc = mfQuery) then stSQL := FaitRequeteDonnees(stType, stSQLCompte, False);

    {Calcul des cumuls }
    if (fModeFonc = mfQuery) then
    begin
      Q := OpenSQL(stSQL, True,-1,'',False,GetListeMultiDossier(fRegroupement));
      try
        fTLesComptesRub.LoadDetailDB('','','',Q,False);
      finally
        Ferme (Q);
      end;
    end;
    for i :=0 to fTLesComptesRub.Detail.Count - 1 do
    begin
      if ((fTLesComptesRub.Detail[i].GetValue('G_COLLECTIF')='X') and (not fbCompensation)) then
          ldValCompte := GetValeurUnCompteCollectif ( fTLesComptesRub.Detail[i], LeDetail, lTResult )
      else ldValCompte := GetValeurUnCompte ( fTLesComptesRub.Detail[i], LeDetail , lTResult, false );

      { Mémorisation du détail des comptes }
      if LeDetail <> nil then AjouteLigneDetail ( LeDetail, fTLesComptesRub.Detail[i], lTResult );

      { Totalisation }
      Result := Arrondi( Result + ldValCompte , V_PGI.OkDecV);
    end;
  end  else Result := fLastErrorMsg;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 30/06/2005
Modifié le ... :   /  /
Description .. : Initialsation des critères de calcul.
Mots clefs ... :
*****************************************************************}
function TZCumul.InitCriteres(pstTypeEcr,pstEtablissement, pstDevise, pstDate : string; pstBalsit: string = '' ; pDateCalculRubrique : TDateTime = 0 ): boolean;
var
  Err: integer;
  Exo: TExoDate;
begin
  fLastError := 0;
  fLastErrorMsg := '';
  fEtablissement := pstEtablissement;
  fDevise := pstDevise;
  fBalsit := pstBalSit;
  if pDateCalculRubrique = 0 then
    fDateCalculRubrique := iDate2099
  else fDateCalculRubrique := pDateCalculRubrique;
  fpStTypeEcr := pStTypeEcr;
  QuelTypEcr( pStTypeEcr );
  Result := (WhatDate(pstDate, fDateDebut, fDateFin, Err, Exo) or (fBalsit<>''));
  if Result then
    fExercice := Exo.Code;

  fRacineCleHash := pstTypeEcr + ';' + pstEtablissement + ';' + pstDevise + ';' + pstDate + ';' + pstBalsit + ';' ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 30/06/2005
Modifié le ... :   /  /
Description .. : Initialisation des critères ANALYTIQUE de calcul. Cette
Suite ........ : méthode n'est à utiliser que dans le cas d'un calcul sur
Suite ........ : RUBRIQUE pour récupérer les écritures analytiques
Suite ........ : associées à un compte général
Mots clefs ... :
*****************************************************************}
function TZCumul.InitAnalytique(pstAxe, pstSection1, pstSection2: string): boolean;
begin
  fAxe := pstAxe;
  fSection1 := pstSection1;
  fSection2 := pstSection2;

  fRacineCleHash := fRacineCleHash +';'+ pstAxe +';'+ pstSection1 + ';' + pstSection2 + ';';
  
  Result := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 30/06/2005
Modifié le ... : 30/06/2005
Description .. : Décode l'expression indiquant le type d'écritures à traiter.
Suite ........ : Cette expression contient 3 type d'indication :
Suite ........ : - le type d'écriture (_QUALIFPIECE)
Suite ........ : - le type de solde
Suite ........ : - la prise en compte des écritures d'à-nouveau
Mots clefs ... :
*****************************************************************}
procedure TZCumul.QuelTypEcr( St: string );
begin
  fTypeSolde := 'SM';
  fAvecANO := 'TOU';
  fTypeEcr := [];
  if Pos('N', St) > 0 then fTypeEcr := fTypeEcr + [tpReel];
  if Pos('S', St) > 0 then fTypeEcr := fTypeEcr + [tpSim];
  if Pos('P', St) > 0 then fTypeEcr := fTypeEcr + [tpPrev];
  if Pos('U', St) > 0 then fTypeEcr := fTypeEcr + [tpSitu];
  if Pos('R', St) > 0 then fTypeEcr := fTypeEcr + [tpRevi];
  if Pos('I', St) > 0 then fTypeEcr := fTypeEcr + [tpIfrs];
  if fTypeEcr = [] then fTypeEcr := [tpReel];
  if Pos('-',St)>0 then fAvecAno:='SAN' ;
  if Pos('#',St)>0 then fAvecAno:='ANO' ;
  if Pos('D',St)>0 then fTypeSolde := 'TD' ;
  if Pos('C',St)>0 then fTypeSolde := 'TC' ;
  if Pos('/',St)>0 then fTypeSolde := 'SD' ;
  if Pos('\',St)>0 then fTypeSolde := 'SC' ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 01/06/2005
Modifié le ... :   /  /
Description .. : Calcul des cumuls pour une ligne de détail de la rubrique
Mots clefs ... :
*****************************************************************}

function TZCumul.PositionneCumuls(pTotalDebit, pTotalCredit: double;
  stCompte: string; var TResult: TabloExt; bIsCollectif : boolean): double;
var
  stTypeSolde: string;
  stSigne : string;
  dValRub: double;
  bUneFois: boolean;
begin
  { Si calcul sur rubrique, on récupère le type de solde défini sur la rubrique }
  if fTypeCumul = 'RUBRIQUE' then
    stTypeSolde := QuelTypDeSolde(fCurRubrique.GetValue('RB_COMPTE1'), stCompte,
      fbGene, bUneFois)
  { Sinon, on récupère le type de solde calculé auparavant par QuelTypeEcr }
  else stTypeSolde := fTypeSolde;

  (* 1 pour SM ; 2 pour SC ; 3 pour SD ; 4 pour TC ; 5 pour TD ; 6 pour TotalRub ou QuelCpte *)
  (* CA - 13/09/2006 : si balance de situation, on suppose qu'il n' a pas eu compensation *)

  TResult[1] := Arrondi(pTotalDebit - pTotalCredit, V_PGI.OkDecV);
  if ((fBalSit <> '') and (bIsCollectif)) then
    TResult[2] := Arrondi(pTotalCredit, V_PGI.OkDecV)
  else if (pTotalCredit > pTotalDebit) then
    TResult[2] := Arrondi(pTotalCredit - pTotalDebit, V_PGI.OkDecV)
  else
    TResult[2] := 0;

  if ((fBalSit <> '') and (bIsCollectif)) then
    TResult[3] := Arrondi(pTotalDebit, V_PGI.OkDecV)
  else
  if (pTotalDebit > pTotalCredit) then
    TResult[3] := Arrondi(pTotalDebit - pTotalCredit, V_PGI.OkDecV)
  else
    TResult[3] := 0;

  TResult[4] := pTotalCredit;
  TResult[5] := pTotalDebit;

  dValRub := 0;

  if fTypeCumul = 'RUBRIQUE' then stSigne := fCurRubrique.GetValue('RB_SIGNERUB') else
  stSigne := 'POS';

  if ( stSigne = 'NEG') then
  begin
    if (stTypeSolde = 'SM') then
      dValRub := (-1) * TResult[1]
    else if (stTypeSolde = 'SC') then
      dValRub := TResult[2]
    else if (stTypeSolde = 'SD') then
      dValRub := (-1) * TResult[3]
    else if (stTypeSolde = 'TC') then
      dValRub := TResult[4]
    else if (stTypeSolde = 'TD') then
      dValRub := (-1) * TResult[5];
  end
  else
  begin
    if (stTypeSolde = 'SM') then
      dValRub := TResult[1]
    else if (stTypeSolde = 'SC') then
      dValRub := (-1) * TResult[2]
    else if (stTypeSolde = 'SD') then
      dValRub := TResult[3]
    else if (stTypeSolde = 'TC') then
      dValRub := (-1) * TResult[4]
    else if (stTypeSolde = 'TD') then
      dValRub := TResult[5];
  end;
  if (stTypeSolde = 'SM') then TResult[6] := 7
  else if (stTypeSolde = 'SC') then TResult[6] := 6
  else if (stTypeSolde = 'SD') then TResult[6] := 5
  else if (stTypeSolde = 'TC') then TResult[6] := 3
  else if (stTypeSolde = 'TD') then TResult[6] := 2;
  Result := dValRub;
end;

procedure TZCumul.RecupereInfoChamps(stType: string; var stPrefDonnee,
  stPrefCode, stCompte1, stCompte2: string);
begin
  if fBalsit <> '' then
  begin
    stPrefDonnee := 'BSE';
    stPrefCode := 'G';
    stCompte1 := '_COMPTE1';
    stCompte2 := '_COMPTE2';
  end else
  begin
    if ((stType = 'RUBRIQUE') or (stType = 'FOURCHETTE')) then
    begin
      if (fAxe = '')  then
      begin
        { Comptabilité générale }
        stPrefDonnee := 'E';
        stPrefCode := 'G';
        stCompte1 := '_GENERAL';
        stCompte2 := '_AUXILIAIRE';
      end else
      begin
        { Comptabilité analytique }
        stPrefDonnee := 'Y';
        stPrefCode := 'G';
        stCompte1 := '_GENERAL';
        //stCompte2 := '_AUXILIAIRE';
        stCompte2 := '';
      end;
    end else if stType = 'GENERAUX' then
    begin
      stPrefDonnee := 'E';
      stPrefCode := 'G';
      stCompte1 := '_GENERAL';
      stCompte2 := '';
    end else if stType = 'TIERS' then
    begin
      stPrefDonnee := 'E';
      stPrefCode := 'T';
      stCompte1 := '_AUXILIAIRE';
      stCompte2 := '';
    end else if stType = 'SECTION' then
    begin
      stPrefDonnee := 'Y';
      stPrefCode := 'S';
      stCompte1 := '_SECTION';
      stCompte2 := '';
    end else exit;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 30/06/2005
Modifié le ... :   /  /
Description .. : Transforme le nom de la table de travail en compte lisible
Suite ........ : par les anciennes fonctions du GetCumul (QuelFb ...)
Mots clefs ... :
*****************************************************************}
function TZCumul.TableToCodeTable(St: string): string;
begin
  if (St = 'GENERAUX') then Result := 'GEN'
  else if (St = 'TIERS') then Result := 'TIE'
  else if (St = 'SECTION') then Result := 'ANA'
  else if (St = 'RUBRIQUE') then Result := 'RUB'
  else Result := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 28/11/2005
Modifié le ... :   /  /
Description .. : Positionnement du mode de fontionnement du calcul des
Suite ........ : cumuls
Mots clefs ... :
*****************************************************************}
procedure TZCumul.SetModeFonc(const Value: TModeFoncCumul);
begin
  fModeFonc := Value;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 28/11/2005
Modifié le ... :   /  /
Description .. : Chargement de tous les cumuls en mémoire
Mots clefs ... :
*****************************************************************}
procedure TZCumul.ChargeCumulMemoire (pstType : string );
var
  lstSQL : string;
  lT : TOB;
  lQ : TQuery;
  lStCle : string;
begin
  lstSQL := FaitRequeteDonnees( pstType, '', False);

  { Pour indiquer que les cumuls sont présents en base }
  lT := TOB.Create ('',nil,-1);
  FHashTable.Put(fRacineCleHash, TObject(lT));

  { Chargement des cumuls }
  lQ := OpenSQL (lstSQL, True);
  try
    while not lQ.Eof do
    begin
      lT := TOB.Create ('',nil,-1);
      lT.SelectDB('',lQ);
  { ECR - 26/12/2007 : Construction de la clé de hashage : le séparateur des différents élements
  est le caractère ":" pour pouvoir utiliser la fonction RemoveStack sur une
  racine }
//      lStCle := fRacineCleHash+';'+lT.GetString('COMPTE1');
        lStCle := fRacineCleHash+':'+lT.GetString('COMPTE1');
//      if (fAxe <> '') then lStCle := lStCle + ';'+ lT.GetString('Y_SECTION');     {En ANAL2, le champ s'appelle tj Y_SECTION}
        if (fAxe <> '') then lStCle := lStCle + ':'+ lT.GetString('Y_SECTION');     {En ANAL2, le champ s'appelle tj Y_SECTION}
      FHashTable.Put( lStCle, TObject(lT));
      lQ.Next;
    end;
  finally
    Ferme(lQ);
  end;
end;

procedure TZCumul.ChargeLesComptesMemoire ( pstSQL : string );
begin
  if ( fTLesComptes.Detail.Count = 0 ) then fTLesComptes.LoadDetailFromSQL(pstSQL);
end;


// $$$ JP 04/07/06: optimisat&ion d'accès aux données, en utilisant la FList de la TOB, et recherche dichotomique
// $$$ JP 04/07/06: amélioration possible plus tard: avoir un tableau de "position de départ" des classes 1 à 8, cf g31calc.dll
procedure TZCumul.FaitListeDesComptesRubrique ( pLInclus, pLExclus : TStringList);
var
   i           :integer;
   St, C1, C2  :string;
   lT          :TOB;
   TOBList     :PPointerList;
   iLow, iHigh :integer;
   iPos        :integer;
   iTmp        :integer;
   strTOBC1    :string;
   OkExclure   : Boolean;
   stCompte    : string;  // Champ compte sur lequel l'appartenance à la rubrique va être testé

    function getCompteCorresp ( plT : TOB; pstCompte : string ) : string;
    begin
      if Corresp = coGeneral then Result := plT.GetString (pstCompte)
      else
      begin
        { Pas de comptes de correspondance, on prend le compte général }
        if (Trim (plT.GetString ( stCompte ))='') then Result := plT.GetString ('COMPTE1')
        else Result := BourreOuTronque(plT.GetString (stCompte),fbGene);
      end;
    end;

    function IlfautExclure ( CC : string ) : boolean;
    var ii : integer;
        C1Ex, C2Ex : string;
    begin
      Result := False;
      for ii := 0 to pLExclus.Count - 1 do
      begin
        St := pLExclus[ii];
        if Pos (':', St) > 0 then
        begin
          // Tranche de compte
          C1Ex := ReadTokenPipe(St,':');
          C2Ex := St;
        end else if Pos ('%', St) > 0 then
        begin
          // Like <=> tranche de compte XXX000000 à XXXZZZZZZ
          C1Ex := Copy (St , 1, Pos ('%', St) - 1);
          C2Ex := C1Ex + Copy (BOURRE_COMPTE_SUP, 1, GetInfoCpta (fbGene).Lg - Length (C1Ex));
          C1Ex := C1Ex + Copy (BOURRE_COMPTE_INF, 1, GetInfoCpta (fbGene).Lg - Length (C1Ex));
        end else
        begin
          // Valeur <=> tranche de compte XXX à XXX (C1Ex = C2Ex)
          C2Ex := St + Copy (BOURRE_COMPTE_SUP, 1, GetInfoCpta (fbGene).Lg - Length (St));
          C1Ex := St + Copy (BOURRE_COMPTE_INF, 1, GetInfoCpta (fbGene).Lg - Length (St));
        end;
        if ((CC >= C1Ex) and (CC <= C2Ex)) then begin Result := TRUE; break; end;
      end;
    end;
(*
CA - 03/05/2007 - Ancienne fonction qui ne marche pas du tout !!!

    Function IlfautExlure (CC : string) : Boolean;
    var
    ii : integer;
    begin
          Result := False;
          for ii := 0 to pLExclus.Count - 1 do
          begin
               St := pLExclus[ii];
               if (Pos ( ':', St ) > 0) then
               begin
                    C1 := ReadTokenPipe(St,':');
                    C2 := St;
                    if (Copy (CC, 1 , Length(C1))) = C1 then begin Result := TRUE; break; end
                    else
                    if (Copy (CC, 1 , Length(C2))) = C2 then begin Result := TRUE; break; end;
               end
               else
               begin
                     if (Copy (CC, 1 , Length(St))) = St then begin Result := TRUE; break; end;
               end;
          end;
    end;
*)
begin
     fTLesComptesRub.ClearDetail;
     OkExclure := FALSE;
     if pLExclus.Count > 0 then OkExclure := TRUE;

     { Affectation du champ de correspondance }
     if Corresp=coCorresp1 then stCompte := 'CORRESP1'
     else if Corresp=coCorresp2 then stCompte := 'CORRESP2'
     else stCompte := 'COMPTE1';

     for i := 0 to pLInclus.Count - 1 do
     begin
          St := pLInclus[i];
          if Pos (':', St) > 0 then
          begin
               // Tranche de compte
               C1 := ReadTokenPipe(St,':');
               C2 := St;
          end
          else if Pos ('%', St) > 0 then
          begin
               // Like <=> tranche de compte XXX000000 à XXXZZZZZZ
               C1 := Copy (St , 1, Pos ('%', St) - 1);
               C2 := C1 + Copy (BOURRE_COMPTE_SUP, 1, GetInfoCpta (fbGene).Lg - Length (C1));
               C1 := C1 + Copy (BOURRE_COMPTE_INF, 1, GetInfoCpta (fbGene).Lg - Length (C1));
          end
          else
          begin
               // Valeur <=> tranche de compte XXX à XXX (C1 = C2)
               C1 := St;
               C2 := St;
          end;

          // $$$ JP 04/07/06: FList du détail de la TOB
          TOBList := fTLesComptes.Detail.List;
          iLow    := 0;
          iHigh   := fTLesComptes.Detail.Count - 1;
          iPos    := 0;
          { Récupération du premier compte trouvé qui correspond aux critères }
          while (iLow <= iHigh) do
          begin
               // Milieu
               iPos := (iLow + iHigh) div 2;

               // Récupération de l'objet pivot pour le test
               lT := TOB (TOBList^ [iPos]);
               // strTOBC1 := Trim (lT.GetString ('COMPTE1'));
               strTOBC1 := getCompteCorresp (lT, stCompte);
               if strTOBC1 <> '' then
               begin
                    if strTOBC1 < C1 then
                         iLow := iPos + 1
                    else if strTOBC1 > C1 then
                         iHigh := iPos - 1
                    else
                        break;
               end;
          end;

          // Si la borne inférieure non trouvée, le 1er compte possible dans la fourchette est celui pointée par iLow
          if iLow > iHigh then
             iPos := iLow;

          // Dans le cas des comptes de correspondance, on peut avoir plusieurs comptes
          // généraux pour le même compte de correspondance. Il faut donc recalculer iPos
          if (iPos <= iHigh) then // CA - FQ 21511 - 25/09/2007 - Cas d'une rubrique 99 : iPos > iHigh
          begin
            if ((Corresp=coCorresp1) or (Corresp=coCorresp2)) then
            begin
             lT := TOB (TOBList^ [iPos]);
             strTOBC1 := getCompteCorresp (lT, stCompte);
             while ((iPos> 0) and (strTOBC1 = getCompteCorresp (TOB (TOBList^ [iPos-1]),stCompte))) do
              Dec (iPos);
            end;
          end;


          // Position du premier compte trouvé (qu'il existe ou pas dans la liste des comptes)
          iHigh := fTLesComptes.Detail.Count - 1;
          while iPos <= iHigh do
          begin
               // On recopie la TOB Fille du compte, s'il est inclus dans la fourchette
               lT := TOB (TOBList^ [iPos]);
               //strTOBC1 := Trim (lT.GetString ('COMPTE1'));
               strTOBC1 := getCompteCorresp (lT, stCompte);
               if (strTOBC1 <> '') and (strTOBC1 <= C2) then
               begin
                 // CA - 29/09/2006 - Gestion du cas analytique = plusieurs lignes sur le même compte
                 if fAxe <> '' then
                 begin
                  if (Not OkExclure) or  (OkExclure and (not IlfautExclure (strTOBC1))) then
                  begin
                    // On se place sur la première occurence du compte
                    iTmp := iPos;
                    //while (Trim (TOB (TOBList^ [iTmp]).GetString ('COMPTE1'))=strTOBC1) do
                    while (getCompteCorresp(TOB (TOBList^ [iTmp]),stCompte)=strTOBC1) do
                    begin
                      Dec (iTmp);
                      if (iTmp < 0) then break;
                    end;
                    // On récupère tous les couples compte/section qui conviennent
                    Inc (iTmp);
                    // while (Trim (TOB (TOBList^ [iTmp]).GetString ('COMPTE1'))=strTOBC1) do
                    while (getCompteCorresp(TOB (TOBList^ [iTmp]),stCompte)=strTOBC1) do
                    begin
                      lT := TOB (TOBList^ [iTmp]);
                      TOB.Create ('', fTLesComptesRub, -1).Dupliquer (lT, FALSE, TRUE);
                      Inc (iTmp);
                      if (iTmp > iHigh) then break;
                    end;
                    // On se place sur le compte suivant
                    iPos := iTmp;
                  end else Inc(iPos);
                 end else
                 begin
                   if (Not OkExclure) or  (OkExclure and (not IlfautExclure (strTOBC1))) then
                       TOB.Create ('', fTLesComptesRub, -1).Dupliquer (lT, FALSE, TRUE);
                    Inc (iPos);
                 end;
               end
               else
                   break;
          end;
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe AYEL
Créé le ...... : 29/10/2007
Modifié le ... : 29/10/2007
Description .. : Tranformation d'une fourchette de comptes en rubrique
Suite ........ : Le type de solde est extrait de la description des fourchettes
Mots clefs ... :
*****************************************************************}
procedure TZCumul.FourchetteVersRubrique ( pstFourchette : string );
var
  lstCompte1 : string;
  lstTranche : string;
  lstCpt1, lstCpt2 : string;
begin
  if fRubrique.Detail.Count > 0 then fRubrique.ClearDetail;
  fCurRubrique := TOB.Create('RUBRIQUE',fRubrique,-1);

  // On retraite les tranches pour bourrer la tranche supérieure avec des '9'
  lstTranche := ReadTokenSt (pstFourchette);
  lstCompte1 := '';
  while (lstTranche <> '' ) do
  begin
    lstCpt1 := ReadTokenPipe(lstTranche,':');
    // Ajout CA - 07/12/2007 - Prise en compte des racines
    if lstTranche = '' then lstTranche := lstCpt1;
    // Fin ajout CA
    lstCpt1 := lstCpt1 + Copy (BOURRE_COMPTE_INF, 1, GetInfoCpta (fbGene).Lg - Length (lstCpt1));
    lstCpt2 := lstTranche + Copy (BOURRE_COMPTE_SUP, 1, GetInfoCpta (fbGene).Lg - Length (lstTranche));
    lstTranche := ReadTokenSt (pstFourchette);
    lstCompte1 := lstCompte1 + lstCpt1 + ':' + lstCpt2 + ';';
  end;

  // On ajoute le type de solde à la fin de chaque tranche
  StringReplace(lstCompte1, ';', '('+fTypeSolde+');', [rfReplaceAll,rfIgnoreCase]);

  // Mise à jour des champs
  fCurRubrique.PutValue('RB_COMPTE1',lstCompte1);
  fCurRubrique.PutValue('RB_EXCLUSION1','');
  fCurRubrique.PutValue('RB_SIGNERUB','POS');
  fCurRubrique.PutValue('RB_TYPERUB','GEN');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe AYEL
Créé le ...... : 26/11/2007
Modifié le ... : 26/11/2007
Description .. : Méthode de réinitialisation des cumuls. Cette méthode
Suite ........ : supprime les cumuls en mémoire. Ces cumuls seront
Suite ........ : recalculés à la prochaine demande (appel de GetValeur)
Suite ........ : La paramètre bAll permet de remettre à jour tous les cumuls
Suite ........ : stockés dans l'objet.
Mots clefs ... :
*****************************************************************}
function TZCumul.ReinitCumul ( bAll : boolean ) : boolean;
var
  lT : TOB;
begin
  // ECR 21/12/2007 : Effacer la liste pour recontruction avec les comptes éventuellement
  // crées ou supprimés
  fTLesComptes.Detail.Clear;


  if ( fModeFonc = mfMemory ) then
  begin
    if bAll then fHashTable.Clear(True)
    else
    begin
      lT := TOB(fHashTable.Get(fRacineCleHash));
      if lT <> nil then
      begin
        lT.Free;
        fHashTable.remove(fRacineCleHash);
        // ECR 26/12/07 : Vider toutes les entrées correspondantes
        fHashTable.removeStack(fRacineCleHash, True);
      end;
    end;
  end;
  Result := True;
end;

end.

