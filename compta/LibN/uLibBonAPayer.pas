{ Unité : Outils / objets de gestion des bons à payer
--------------------------------------------------------------------------------------
    Version    |   Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001  08/02/06  JP   Création de l'unité
 8.01.001.001  07/12/06  JP   FQ 19164 : Pouvoir appliquer un "et" ou un "ou" entre les conditions
                               libres : ajout du champs CTI_AUCHOIX dans la table CPTYPEVISA
 8.00.001.018  29/05/07  JP   Mise en place des rôles à la place des groupes pour filtrer les utilisateurs
 8.01.001.021  20/06/07  JP   FQ 20778 : Les natures de pièces ne sont plus obligatoires
--------------------------------------------------------------------------------------}
unit uLibBonAPayer;

interface

uses
  {$IFDEF EAGLCLIENT}
  {$ELSE}
    {$IFNDEF DBXPRESS}
    dbtables,
    {$ELSE}
    uDbxDataSet,
    {$ENDIF DBXPRESS}
  {$ENDIF EAGLCLIENT}
  SysUtils, HEnt1, Classes, uTob, HCtrls;

const
  USERTACHE = 'U@T';
  {Différents critères libres des types de visas : Tablette CPTYPELIBRE}
  tyt_Devise       = 'DEV';
  tyt_RefInterne   = 'REF';
  tyt_TLAuxiliaire = 'TLA';
  tyt_TLEcriture   = 'TLE';
  tyt_TLGeneral    = 'TLG';
  tyt_TLSection    = 'TLS';

  {Statuts des BAP : tablette CPSTATUTBAP}
  sbap_Encours     = 'ENC';
  sbap_Valide      = 'VAL';
  sbap_Refuse      = 'REF';
  sbpa_bloque      = 'BLO';
  sbap_Analytique  = 'ANA';
  sbap_Definitif   = 'DEF';

  {Modèle de mails : tablette CPNATUREMAIL }
  nam_Alerte       = 'ALE';
  nam_Bloque       = 'BLO';
  nam_Refuse       = 'REF';
  nam_Suppleant    = 'SUP';
  nam_Relance      = 'REL';

  {Erreurs lors de la génération des BAP}
  ebap_None      = -1;
  ebap_Requete   = 0;
  ebap_MultiTTC  = 1;
  ebap_Entete    = 2;
  ebap_TTC       = 3;
  ebap_HT        = 4;
  ebap_Section   = 5;
  ebap_WhereVisa = 6;
  ebap_PasDeVisa = 7;
  ebap_CreateBap = 8;
  ebap_TiersPaye = 9;
  ebap_Circuit   = 10;
  ebap_CalDelais = 11;
  ebap_GenereBap = 12;

  ebap_ParamAna  = 100;
  ebap_ParamGen  = 200;
  ebap_ParamAux  = 300;
  ebap_ParamEcr  = 400;

  {Tâches du process server}
  psrv_Etape     = 'ETAPE';
  psrv_EnvoiMail = 'ENVOI';
  psrv_Alerte    = 'ALERTE';

type
  PClefPiece = ^TClefPiece;

  TClefPiece = record
    E_JOURNAL       : string;
    E_EXERCICE      : string;
    E_DATECOMPTABLE : TDateTime;
    E_NUMEROPIECE   : Integer;
    E_NUMLIGNE      : Integer;
    E_NUMECHE       : Integer;
    E_QUALIFPIECE   : string;
  end;

  {nma_Alerte   : Alerte Viseur 1
   nma_Relance1 : Relance Viseur 1 et Alerte Viseur 2
   nma_Relance2 : Relance Viseur 1 et Relance Viseur 2}
  TNiveauMail = (nma_Alerte, nma_Relance1, nma_Relance2);

  {Classe permettant le chargement des différentes lignes d'un circuit et ainsi
   de récupérer les informations nécessaires en fonction du numéro d'ordre (ligne)}
  TObjCircuit = class
  private
    TobCircuit : TOB;
  public
    constructor Create(Circuit : string);
    destructor  Destroy; override;

    function GetViseur1(Num : Integer) : string;
    function GetViseur2(Num : Integer) : string;
    function GetDelais (Num : Integer) : Integer;
  end;

  {Classe permettant de de générer un enregistrement dans CPBONSAPAYER}
  TBonAPayer = class
    TobBap   : TOB;
    TobMasse : TOB;
  private
    {Tob qui contiendra toutes les informations nécessaires au "calcul" du Bap}
    TobCalcul  : TOB;

    function  TesteRequete(T : TOB) : Boolean;
    function  GetClauseWhere        : string;

    procedure ChargeTobCalcul (Clef : TClefPiece);
    procedure SetEntete       (T : TOB);
    procedure SetTTC          (T : TOB);
    procedure SetHT           (T : TOB);
    procedure SetSection      (T : TOB);
    procedure TesteComptes    (T : TOB);
    procedure TesteZonesLibres(T : TOB);
    procedure SetTobBap       (T : TOB);
    procedure GetTypeVisa     ;
    {$IFNDEF EAGLSERVER}
    procedure GetTiersPayeur  (var T : TOB; TobEcr : TOB = nil);
    {$ENDIF EAGLSERVER}
    procedure CalculEcheBap   (Q : TQuery; DateFact : TDateTime; var T : TOB);
  public
    {Si les critères pour générer un BAP ne sont pas réunis => True}
    BapImpossible   : Boolean;
    CodeErreur      : Integer;
    TraitementMasse : Boolean;
    JournalTP       : string; {Journal achate des tiers payeurs}

    destructor  Destroy; override;
    {Constructeur pour la compta : Calcul de la première étape du Bap ou bien chargement d'une étape}
    constructor Create  (ClePiece : TClefPiece; CalculOk : Boolean = True; MasseOk : Boolean = False); overload;
    {Constructeur pour le process server lors de la création des étapes autre que la première}
    constructor Create  (TobEtape : Tob; Definitif : Boolean); overload;
    {Pour la ComSx}
    constructor Create  (TobEcr : TOB; TypeVisa : string); overload;
    procedure   CreerBap(ClePiece : TClefPiece); overload;
    procedure   CreerBap(TobEcr : TOB); overload;
    procedure   ForceBap(ClePiece : TClefPiece; TypeVisa, Circuit : string);
    procedure   GetCircuit(NoOrdre : Integer; var T : TOB);
    function    DupliquerClef(TobSource : TOB; var TobDest : TOB) : Boolean;
  end;

  {Classe permettant de de générer un enregistrement dans CPTACHEBAP}
  TObjMail = class
    TobMail  : TOB;
  private
    TobBap   : TOB;
    TobTypeM : TOB;
    TobUser  : TOB;

    function  RemplirTob(Niveau : TNiveauMail) : string;
    function  MajBap    (Niveau : TNiveauMail) : string;
  public
    destructor  Destroy; override;
    constructor Create   (ClePiece : TClefPiece);
    function    CreerMail(Niveau  : TNiveauMail) : string;
    function    NatMailFromNiveau(Niveau  : TNiveauMail; Viseur : Char) : string;
  end;

  {Classe qui génère et envoie les mails}
  TObjEnvoiMail = class
    TobMail : TOB;
  private

  public
    destructor  Destroy; override;
    constructor Create(Nat, User : string; DtEnvoi : TDateTime; T : TOB = nil);
    function    EnvoieMail : Boolean;
  end;

  {Classe qui génère les enregistrements dans la Table CPTACHEBAP}
  TObjTacheBap = class(TOB)
  private
    FNature       : string;
    FDestinataire : string;
    FDateEnvoi    : TDateTime;
    procedure SetNature      (Value : string);
    procedure SetDestinataire(Value : string);
    procedure SetDateEnvoi   (Value : TDateTime);
    procedure CreerTache     (TobInfo, TobMail : TOB);
    procedure MajTache       ;
  public
    ViseurPrincipal : Boolean;

    destructor  Destroy; override;
    class function CreateObj(Nat, User : string; DtEnvoi : TDateTime; LeParent : TOB = nil) : TObjTacheBap;

    procedure PrepareTache(TobInfo, TobMail : TOB);

    property Nature       : string    read FNature       write SetNature;
    property Destinataire : string    read FDestinataire write SetDestinataire;
    property DateEnvoi    : TDateTime read FDateEnvoi    write SetDateEnvoi;
  end;

procedure SetPlusCombo(Zone : THValComboBox; Pref : string);
{$IFNDEF EAGLSERVER}
procedure LookUpUtilisateur(Sender : TObject);
{$ENDIF EAGLSERVER}
function  CanDeleteTypeVisa(Code : string) : Boolean;
function  CanDeleteCircuit (Code : string) : Boolean;
function  GetMessageErreur (Code : Integer) : string;
function  EncodeClefTP     (ClefTP : TClefPiece) : string;
function  DecodeClefTP     (Clef : string; Token : Char = ';') : TClefPiece;
function  GetWhereBapT     (T : TOB; Prefix : string = 'BAP'; AvecNum : Boolean = True) : string;
function  GetWhereBapC     (C : TClefPiece; AvecNum : Boolean = True) : string;
function  GestionBap       (NaturePiece, Journal, Dossier : string) : Boolean;
function  IsJournalBap     (NaturePiece, Journal, Dossier : string) : Boolean;
function  ExisteTypeVisa   (Dossier : string) : Boolean;
procedure DetruitBap       (aTob : TOB; Dossier : string);
function  ClefFromPieceTP  (PieceTP : string) : TClefPiece;


implementation

uses
  HMsgBox, ParamSoc, ULibEcriture, UtilPGI, ULibWindows, Math, Controls,
  UProcGen, UObjGen, Registry, Windows, SaisUtil, uFichierLog 
  {$IFDEF EAGLSERVER}
  , eHttp
  {$ELSE}
  ,TiersPayeur, MailOl, LookUp
  {$ENDIF EAGLSERVER}
  ;

const
  REQUETEBAP =
    ' SELECT E_JOURNAL, E_DATECOMPTABLE, E_DATEECHEANCE, E_EXERCICE, E_NUMEROPIECE, E_TYPEMVT, ' +
         'E_NUMECHE, E_NATUREPIECE, E_PIECETP, E_DEVISE, E_SOCIETE, EC_DOCGUID, E_QUALIFPIECE, ' + // SBO 22/05/2007 il manquait le champ E_QUALIFPIECE à priori
         'ABS(E_DEBITDEV + E_CREDITDEV - E_COUVERTUREDEV) AS MONTANT, E_BLOCNOTE, ' + {JP 28/01/08 : Ajout de E_BLOCNOTE}
         'E_REFINTERNE, E_ETABLISSEMENT, E_TABLE0, E_TABLE1, E_TABLE2, E_TABLE3, ' +
         'ABS(Y_DEBITDEV + Y_CREDITDEV) AS VENTILATION, Y_AXE, Y_SECTION, ' +
         'G_GENERAL, G_NATUREGENE, G_TABLE0, G_TABLE1, G_TABLE2, G_TABLE3, ' +
         'G_TABLE4, G_TABLE5, G_TABLE6, G_TABLE7, G_TABLE8, G_TABLE9, ' +
         'S_TABLE0, S_TABLE1, S_TABLE2, S_TABLE3, S_TABLE4, ' +
         'S_TABLE5, S_TABLE6, S_TABLE7, S_TABLE8, S_TABLE9, ' +
         'T_TABLE0, T_TABLE1, T_TABLE2, T_TABLE3, T_TABLE4, ' +
         'T_TABLE5, T_TABLE6, T_TABLE7, T_TABLE8, T_TABLE9, T_AUXILIAIRE ' +
         'FROM ECRITURE ' +
         'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
         'LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE  ' +
         'LEFT JOIN ANALYTIQ ON Y_JOURNAL = E_JOURNAL AND Y_EXERCICE = E_EXERCICE AND ' +
                               'Y_DATECOMPTABLE = E_DATECOMPTABLE AND Y_NUMEROPIECE = E_NUMEROPIECE ' +
                               'AND Y_NUMLIGNE = E_NUMLIGNE AND Y_QUALIFPIECE = "N" ' +
         'LEFT JOIN ECRCOMPL ON EC_JOURNAL = E_JOURNAL AND EC_EXERCICE = E_EXERCICE AND ' +
                               'EC_DATECOMPTABLE = E_DATECOMPTABLE AND EC_NUMEROPIECE = E_NUMEROPIECE ' +
                               'AND EC_NUMLIGNE = E_NUMLIGNE AND EC_QUALIFPIECE = "N" ' +
         'LEFT JOIN SECTION ON S_SECTION = Y_SECTION ';


{$IFNDEF EAGLSERVER}
{JP 29/05/07 : Remplacement des Combos Utilisateur par des THEdit
{---------------------------------------------------------------------------------------}
procedure LookUpUtilisateur(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Wh : string;
  s  : string;
begin
  Wh := '(US_GRPSDELEGUES LIKE "%' + GetParamSocSecur('SO_CPGROUPEVISEUR', '') +
        ';%" OR US_GRPSDELEGUES = "" OR US_GRPSDELEGUES IS NULL OR US_GRPSDELEGUES LIKE "<<%>>") AND (US_EMAIL <> "" AND US_EMAIL IS NOT NULL)';

       if Sender is THEdit then s := (Sender as THEdit).Text
  else if Sender is THGrid then s := (Sender as THGrid).Cells[(Sender as THGrid).Col, (Sender as THGrid).Row];

  if s <> '' then
    wh := Wh + ' AND (US_UTILISATEUR LIKE "' + s + '%" OR US_LIBELLE LIKE "' + s + '%")';

  LookupList(TControl(Sender), TraduireMemoire('Liste des viseurs'), 'UTILISAT', 'US_UTILISATEUR', 'US_LIBELLE', Wh, 'US_UTILISATEUR', True, 3170);
end;
{$ENDIF EAGLSERVER}

{---------------------------------------------------------------------------------------}
procedure SetPlusCombo(Zone : THValComboBox; Pref : string);
{---------------------------------------------------------------------------------------}
begin
       {JP 29/05/07 : Mise en place des rôles
       if Pref = 'US' then Zone.Plus := 'US_GROUPE = "' + GetParamSocSecur('SO_CPGROUPEVISEUR', '') + '" AND (US_EMAIL <> "" AND US_EMAIL IS NOT NULL)'}
       if Pref = 'US' then Zone.Plus := 'US_GRPSDELEGUES LIKE "%' + GetParamSocSecur('SO_CPGROUPEVISEUR', '') + ';%" AND (US_EMAIL <> "" AND US_EMAIL IS NOT NULL)'
  else if Pref = 'J'  then Zone.Plus := '(J_NATUREJAL  = "ACH") OR (J_NATUREJAL  = "OD" AND J_JOURNAL = "' + GetParamSocSecur('SO_JALATP', '') + '")'
  ;
end;

{---------------------------------------------------------------------------------------}
function CanDeleteTypeVisa(Code : string) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := not ExisteSQL('SELECT BAP_CODEVISA FROM CPBONSAPAYER WHERE BAP_CODEVISA = "' + Code + '"');
end;

{---------------------------------------------------------------------------------------}
function CanDeleteCircuit(Code : string) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := not ExisteSQL('SELECT BAP_CIRCUITBAP FROM CPBONSAPAYER WHERE BAP_CIRCUITBAP = "' + Code + '"') and
            not ExisteSQL('SELECT CTI_CIRCUITBAP FROM CPTYPEVISA   WHERE CTI_CIRCUITBAP = "' + Code + '"');
end;

{---------------------------------------------------------------------------------------}
function GetMessageErreur(Code : Integer) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := '';

  case Code of
    ebap_Requete   : Result := TraduireMemoire('Impossible de récupérer la pièce dans son intégralité.');
    ebap_MultiTTC  : Result := TraduireMemoire('Il y a plusieurs lignes de TTC : il est impossible de générer un bon à payer.');
    ebap_Entete    : Result := TraduireMemoire('Impossible de récupérer l''entête de la pièce.');
    ebap_TTC       : Result := TraduireMemoire('Impossible de récupérer les informations TTC.');
    ebap_HT        : Result := TraduireMemoire('Impossible de récupérer les informations HT.');
    ebap_Section   : Result := TraduireMemoire('Impossible de récupérer les informations Section.');
    ebap_WhereVisa : Result := TraduireMemoire('Impossible de récupérer la clause where sur les visas.');
    ebap_PasDeVisa : Result := TraduireMemoire('Il n''y a pas de type de visa correspondant à la pièce.');
    ebap_CreateBap : Result := TraduireMemoire('Impossible de créer le bon à payer.');
    ebap_TiersPaye : Result := TraduireMemoire('Impossible de récupérer le tiers payeur.');
    ebap_Circuit   : Result := TraduireMemoire('Impossible de récupérer le circuit de validation.');
    ebap_CalDelais : Result := TraduireMemoire('Impossible de calculer la date d''échéance du bon à payer.');
    ebap_GenereBap : Result := TraduireMemoire('Impossible de générer le bon à payer.');

    ebap_ParamAna  : Result := TraduireMemoire('Certains paramétrages des types de visa sur les tables libres'#13 +
                                               'des sections sont incorrects.');
    ebap_ParamGen  : Result := TraduireMemoire('Certains paramétrages des types de visa sur les tables libres'#13 +
                                               'des généraux sont incorrects.');
    ebap_ParamAux  : Result := TraduireMemoire('Certains paramétrages des types de visa sur les tables libres'#13 +
                                               'des auxiliaires sont incorrects.');
    ebap_ParamEcr  : Result := TraduireMemoire('Certains paramétrages des types de visa sur les tables libres'#13 +
                                               'des écritures sont incorrects.');


    ebap_PasDeVisa +
    ebap_ParamAna  : Result := TraduireMemoire('Il n''y a pas de type de visa correspondant à la pièce.'#13#13 +
                                               'Veuillez cependant vérifier les paramétrages des types de visa'#13 +
                                               'sur les tables libres des sections car certains sont incorrects.' );
    ebap_PasDeVisa +
    ebap_ParamGen  : Result := TraduireMemoire('Il n''y a pas de type de visa correspondant à la pièce.'#13#13 +
                                               'Veuillez cependant vérifier les paramétrages des types de visa'#13 +
                                               'sur les tables libres des généraux car certains sont incorrects.' );
    ebap_PasDeVisa +
    ebap_ParamAux  : Result := TraduireMemoire('Il n''y a pas de type de visa correspondant à la pièce.'#13#13 +
                                               'Veuillez cependant vérifier les paramétrages des types de visa'#13 +
                                               'sur les tables libres des auxiliaires car certains sont incorrects.' );
    ebap_PasDeVisa +
    ebap_ParamEcr  : Result := TraduireMemoire('Il n''y a pas de type de visa correspondant à la pièce.'#13#13 +
                                               'Veuillez cependant vérifier les paramétrages des types de visa'#13 +
                                               'sur les tables libres des écritures car certains sont incorrects.' );
    ebap_None      : Result := '';
  end;
end;

{---------------------------------------------------------------------------------------}
function  GestionBap(NaturePiece, Journal, Dossier : string) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := IsJournalBap(NaturePiece, Journal, Dossier) and ExisteTypeVisa(Dossier);
end;

{---------------------------------------------------------------------------------------}
function IsJournalBap(NaturePiece, Journal, Dossier : string) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := ((NaturePiece = 'OD') and (Journal = GetParamsocDossierSecur('SO_JALATP', '', Dossier))) or
            (NaturePiece = 'AF') or (NaturePiece = 'FF');
end;

{---------------------------------------------------------------------------------------}
function ExisteTypeVisa(Dossier : string) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := ExisteSQL('SELECT CTI_CODEVISA FROM ' + GetTableDossier(Dossier, 'CPTYPEVISA'));
end;

{---------------------------------------------------------------------------------------}
procedure DetruitBap(aTob : TOB; Dossier : string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  T := aTob.Detail[0];
  if Assigned(T) then
    ExecuteSQL('DELETE FROM ' + GetTableDossier(Dossier, 'CPBONSAPAYER') + ' WHERE ' +
               'BAP_JOURNAL = "' + T.GetString('E_JOURNAL') + '" AND ' +
               'BAP_EXERCICE = "' + T.GetString('E_EXERCICE') + '" AND ' +
               'BAP_DATECOMPTABLE = "' + UsDateTime(T.GetDateTime('E_DATECOMPTABLE')) + '" AND ' +
               'BAP_NUMEROPIECE = ' + T.GetString('E_NUMEROPIECE'));
end;

{---------------------------------------------------------------------------------------}
function EncodeClefTP(ClefTP : TClefPiece) : string;
{---------------------------------------------------------------------------------------}
begin
  {Dans les BAP : JOURNAL;EXERCICE;DATECOMPTABLE;NUMEROPIECE;}
  Result := ClefTP.E_JOURNAL + ';' + ClefTP.E_EXERCICE + ';' +
            DateToStr(ClefTP.E_DATECOMPTABLE) + ';' + IntToStr(ClefTP.E_NUMEROPIECE) + ';';
end;

{---------------------------------------------------------------------------------------}
function DecodeClefTP(Clef : string; Token : Char = ';') : TClefPiece;
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  {Dans les BAP : JOURNAL;EXERCICE;DATECOMPTABLE;NUMEROPIECE;}
  Result.E_JOURNAL := ReadTokenPipe(Clef, Token);
  Result.E_EXERCICE := ReadTokenPipe(Clef, Token);

  ch := ReadTokenPipe(Clef, Token);;
  if IsValidDate(ch) then Result.E_DATECOMPTABLE := StrToDate(ch)
                     else Result.E_DATECOMPTABLE := iDate1900;

  ch := ReadTokenPipe(Clef, Token);
  if IsNumeric(ch) then Result.E_NUMEROPIECE := StrToInt(ch)
                   else Result.E_NUMEROPIECE := 0;

  ch := ReadTokenPipe(Clef, Token);
  if ch <> '' then begin
    if IsNumeric(ch) then Result.E_NUMLIGNE := StrToInt(ch)
                     else Result.E_NUMLIGNE := 0;

  end;
end;

{---------------------------------------------------------------------------------------}
function  ClefFromPieceTP(PieceTP : string) : TClefPiece;
{---------------------------------------------------------------------------------------}
var
  St, StC : string;
begin
  {Codé dans MajRefTP selon la structure suivante
  DateComptable+';'+IntToStr(Numpiece)+';'+IntToStr(NumLigne)+';'+IntToStr(NumEche)+';'+Journal+';'}
  St  := PieceTP;

  StC := ReadTokenSt(St);
  Result.E_DATECOMPTABLE := EncodeDate(ValeurI(Copy(StC, 1, 4)), ValeurI(Copy(StC, 5, 2)), ValeurI(Copy(StC, 7, 2)));
  StC := ReadTokenSt(St);
  Result.E_NUMEROPIECE := ValeurI(StC) ;
  StC := ReadTokenSt(St);
  Result.E_NUMLIGNE := ValeurI(StC) ;
  StC := ReadTokenSt(St);
  Result.E_NUMECHE := ValeurI(StC) ;
  StC := ReadTokenSt(St);
  Result.E_JOURNAL := StC;
  Result.E_EXERCICE := QuelExoDT(Result.E_DATECOMPTABLE);
end;

{---------------------------------------------------------------------------------------}
function GetWhereBapT(T : TOB; Prefix : string = 'BAP'; AvecNum : Boolean = True) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := 'WHERE BAP_JOURNAL = "' + T.GetString(Prefix + '_JOURNAL') + '" AND ' +
            'BAP_EXERCICE = "' + T.GetString(Prefix + '_EXERCICE') + '" AND ' +
            'BAP_DATECOMPTABLE = "' + UsDateTime(T.GetDateTime(Prefix + '_DATECOMPTABLE')) + '" AND ' +
            'BAP_NUMEROPIECE = ' + T.GetString(Prefix + '_NUMEROPIECE') + ' ';
  if AvecNum then
    Result := Result + 'AND BAP_NUMEROORDRE = ' + T.GetString(Prefix + '_NUMEROORDRE');
end;

{---------------------------------------------------------------------------------------}
function GetWhereBapC(C : TClefPiece; AvecNum : Boolean = True) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := 'WHERE BAP_JOURNAL = "' + C.E_JOURNAL + '" AND ' +
            'BAP_EXERCICE = "' + C.E_EXERCICE + '" AND ' +
            'BAP_DATECOMPTABLE = "' + UsDateTime(C.E_DATECOMPTABLE) + '" AND ' +
            'BAP_NUMEROPIECE = ' + IntToStr(C.E_NUMEROPIECE) + ' ';
  if AvecNum then
    Result := Result + 'AND BAP_NUMEROORDRE = ' + IntToStr(C.E_NUMLIGNE);
end;

{ TBONAPAYER }

{Pour la ComSx : la pièce n'est pas encore dans la base : on force le type de visa et génère
 le BAP à partir de la Tob Ecriture
{---------------------------------------------------------------------------------------}
constructor TBonAPayer.Create(TobEcr : TOB; TypeVisa : string); 
{---------------------------------------------------------------------------------------}
begin
  if TypeVisa = '' then Exit;
  JournalTP := GetParamsocSecur('SO_JALATP', '');
  
  {Création de ma tob qui va contenir les bons à payer}
  TobMasse := TOB.Create('$$CPBONSAPAYER', nil, -1);
  {On charge les infos du type de visa}
  TobBap := TOB.Create('****', nil, -1);
  TobBap.LoadDetailFromSQL('SELECT * FROM CPTYPEVISA WHERE CTI_CODEVISA = "' + TypeVisa + '"');
  {Création de l'enregsitrement dans CPBONSAPAYER}
  CreerBap(TobEcr);
end;

{---------------------------------------------------------------------------------------}
constructor TBonAPayer.Create(ClePiece : TClefPiece; CalculOk : Boolean = True; MasseOk : Boolean = False);
{---------------------------------------------------------------------------------------}
begin
  TraitementMasse := CalculOk and MasseOk;
  {Il s'agit ici de trouver un type de visa à partir d'une  pièce comptable}
  if CalculOk then begin
    BapImpossible := False;
    CodeErreur    := ebap_None;
    if TraitementMasse then
      TobMasse := TOB.Create('$$CPBONSAPAYER', nil, -1)
    else
      TobBap := TOB.Create('$$CPBONSAPAYER', nil, -1);
    CreerBap(ClePiece);
  end

  {Il s'agit ici de générer une Ligne dans CPBONSAPAYER, sans avoir à chercher un type de visa}
  else begin
    TobBap := TOB.Create('€€€', nil, -1);
    TobBap.LoadDetailFromSQL('SELECT * FROM CPBONSAPAYER WHERE BAP_JOURNAL = "' + ClePiece.E_JOURNAL + '" AND ' +
                             'BAP_EXERCICE = "' + ClePiece.E_EXERCICE + '" AND ' +
                             'BAP_DATECOMPTABLE = "' + USDateTime(ClePiece.E_DATECOMPTABLE) + '" AND ' +
                             'BAP_NUMEROPIECE = ' + IntToStr(ClePiece.E_NUMEROPIECE) + ' AND ' +
                             'BAP_NUMEROORDRE = ' + IntToStr(ClePiece.E_NUMLIGNE));
  end;
end;

{Constructeur pour le process server lors de la création des étapes autre que la première
{---------------------------------------------------------------------------------------}
constructor TBonAPayer.Create(TobEtape : Tob; Definitif : Boolean);
{---------------------------------------------------------------------------------------}
var
  F : TOB;
begin
  try
    F := TOB.Create('CPBONSAPAYER', nil, -1);
    try
      if Definitif then begin
        // FQ 22432 : Pb Visu des écritures contenant des étapes non "DEF" dans CCMP
        // ici, il est possible de mettre en place un execute SQL afin de mettre DEF dans toute les étapes
        // et de se passer de la modification sur la vue.
{       // ancien traitement
        F.SetString('BAP_STATUTBAP', sbap_Definitif);
        F.SetString('BAP_MODIFICATEUR'   , USERTACHE);
        F.UpdateDB;}
        ExecuteSQL('UPDATE CPBONSAPAYER SET BAP_STATUTBAP ="' + sbap_Definitif + '", '
                                         + 'BAP_MODIFICATEUR ="' + USERTACHE + '", '
                                         + 'BAP_DATEMODIF ="' + UsTime(Now) + '" '

                           + GetWhereBapT( TobEtape, 'BAP', False ) ) ;
      end
      else begin
        DupliquerClef(TobEtape, F);
        F.SetDateTime('BAP_DATECREATION' , Date);
        F.SetDateTime('BAP_DATEMODIF'    , Now);
        F.SetString('BAP_STATUTBAP', sbap_Encours);
        F.SetDateTime('BAP_DATEMAIL'     , iDate1900);
        F.SetDateTime('BAP_DATERELANCE1' , iDate1900);
        F.SetDateTime('BAP_DATERELANCE2' , iDate1900);
        F.SetString('BAP_RELANCE1'       , '-'); {Mail de relance au viseur 1}
        F.SetString('BAP_RELANCE2'       , '-'); {Mail de relance au viseur 2}
        F.SetString('BAP_ALERTE1'        , '-'); {Mail d'alerte au viseur 1}
        F.SetString('BAP_ALERTE2'        , '-'); {Mail d'alerte au viseur 2}
        F.SetString('BAP_CREATEUR'       , USERTACHE);
        F.SetString('BAP_MODIFICATEUR'   , USERTACHE);
        GetCircuit(F.GetInteger('BAP_NUMEROORDRE') + 1, F);  {Remplit les champs concernant le circuit de validation}
        {JP 15/02/07 : On met le SetAllModifie, pour forcer le Champ BLOCNOTE qui n'étant pas considérer modifié
                       depuis le LoadDB, n'est pas insérer dans la Base : FQ 22432}
        F.SetAllModifie(True);
        F.InsertDb(nil);
      end;
    finally
      FreeAndNil(F);
    end;
  except
    on E : Exception do
      raise E;
  end;
end;

{---------------------------------------------------------------------------------------}
function TBonAPayer.DupliquerClef(TobSource : TOB; var TobDest : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  try
    ObjLog.Add(TobSource.GetString('BAP_JOURNAL') + ' / ' +
               TobSource.GetString('BAP_EXERCICE') + ' / ' +
               TobSource.GetString('BAP_DATECOMPTABLE') + ' / ' +
               TobSource.GetString('BAP_NUMEROPIECE') + ' / ' +
               TobSource.GetString('BAP_NUMEROORDRE') + ' / ' );
    TobDest.SetString('BAP_JOURNAL'        , TobSource.GetString('BAP_JOURNAL'));
    TobDest.SetString('BAP_EXERCICE'       , TobSource.GetString('BAP_EXERCICE'));
    TobDest.SetDateTime('BAP_DATECOMPTABLE', TobSource.GetDateTime('BAP_DATECOMPTABLE'));
    TobDest.SetInteger ('BAP_NUMEROPIECE'  , TobSource.GetInteger('BAP_NUMEROPIECE'));
    TobDest.SetInteger ('BAP_NUMEROORDRE'  , TobSource.GetInteger('BAP_NUMEROORDRE'));
    TobDest.LoadDB;
    Result := TobDest.GetString('BAP_STATUTBAP') <> '';
  except
    on E : Exception do begin
      raise E;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TBonAPayer.CreerBap(ClePiece : TClefPiece);
{---------------------------------------------------------------------------------------}
begin
  BapImpossible := False;
  CodeErreur := ebap_None;
  {Constitution de la TobCalcul à partir de la pièce}
  ChargeTobCalcul(ClePiece);
  if not BapImpossible then
    {Recherche du type de visa correspondant aux critères de la TobCalcul}
    GetTypeVisa;
end;

{Pour la ComSx
{---------------------------------------------------------------------------------------}
procedure TBonAPayer.CreerBap(TobEcr : TOB);
{---------------------------------------------------------------------------------------}
var
  F : TOB;
  E : TOB;
  C : TOB;
  n : Integer;
  p : Integer;
  Sortir : Boolean;
begin
  E := TobEcr.Detail[0];
  C := nil;
  
  if (not Assigned(E)) or
     (not Assigned(TobBap)) or
     (TobBap.Detail.Count = 0) or
     (not Assigned(TobMasse)) or
     ((E.GetString('E_NATUREPIECE') <> 'OD') and
      (E.GetString('E_NATUREPIECE') <> 'AF') and
      (E.GetString('E_NATUREPIECE') <> 'FF')) then Exit;

  if ((E.GetString('E_NATUREPIECE') = 'OD') and
      (E.GetString('E_JOURNAL') <> JournalTP)) then Exit;

  Sortir := False;
  for n := 0 to TobEcr.Detail.Count - 1 do begin
    if Sortir then Break;
    for p := 0 to TobEcr.Detail[n].Detail.Count - 1 do
      if TobEcr.Detail[n].Detail[p].FieldExists('EC_CLEECR') then begin
        Sortir := True;
        C := TobEcr.Detail[n].Detail[p];
        Break;
      end;
  end;

  F := TOB.Create('CPBONSAPAYER', TobMasse, -1);
  F.SetString('BAP_CODEVISA'       , TobBap.Detail[0].GetString('CTI_CODEVISA'));
  F.SetString('BAP_CIRCUITBAP'     , TobBap.Detail[0].GetString('CTI_CIRCUITBAP'));
  F.SetDateTime('BAP_DATECREATION' , Now);
  F.SetDateTime('BAP_DATEMODIF'    , Now);
  F.SetString('BAP_STATUTBAP'      , sbap_Encours);
  F.SetString('BAP_CREATEUR'       , V_PGI.User);
  F.SetString('BAP_MODIFICATEUR'   , V_PGI.User);
  F.SetDateTime('BAP_DATEMAIL'     , iDate1900);
  F.SetDateTime('BAP_DATERELANCE1' , iDate1900);
  F.SetDateTime('BAP_DATERELANCE2' , iDate1900);
  F.SetString('BAP_RELANCE1'       , '-'); {Mail de relance au viseur 1}
  F.SetString('BAP_RELANCE2'       , '-'); {Mail de relance au viseur 2}
  F.SetString('BAP_ALERTE1'        , '-'); {Mail d'alerte au viseur 1}
  F.SetString('BAP_ALERTE2'        , '-'); {Mail d'alerte au viseur 2}
  F.SetString('BAP_SOCIETE'        , E.GetString('E_SOCIETE'));
  F.SetString('BAP_JOURNAL'        , E.GetString('E_JOURNAL'));
  F.SetString('BAP_EXERCICE'       , E.GetString('E_EXERCICE'));
  F.SetString('BAP_ETABLISSEMENT'  , E.GetString('E_ETABLISSEMENT'));
  if Assigned(C) then
    F.SetString('BAP_IDGED'        , C.GetString('EC_DOCGUID')); {Attention aux tiers payeurs !!!}
  F.SetDateTime('BAP_DATECOMPTABLE', E.GetDateTime('E_DATECOMPTABLE'));
  F.SetDateTime('BAP_DATEECHEANCE' , E.GetDateTime('E_DATEECHEANCE'));
  F.SetInteger ('BAP_NUMEROPIECE'  , E.GetInteger('E_NUMEROPIECE'));

  GetCircuit(1, F);  {Remplit les champs concernant le circuit de validation}
  {$IFNDEF EAGLSERVER}
  GetTiersPayeur(F, E); {Remplit les champs concernant les Tiers Payeurs}
  {$ENDIF EAGLSERVER}
end;

{---------------------------------------------------------------------------------------}
destructor TBonAPayer.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobBap   ) then FreeAndNil(TobBap   );
  if Assigned(TobCalcul) then FreeAndNil(TobCalcul);
  if Assigned(TobMasse ) then FreeAndNil(TobMasse );
  inherited;
end;

{---------------------------------------------------------------------------------------}
function TBonAPayer.TesteRequete(T : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  n  : Integer;
  Ok : Boolean;
begin
  Result := True;
  {Il faut au moins une ligne TTC et une HT pour générer un bap}
  if T.Detail.Count < 2 then begin
    BapImpossible := True;
    CodeErreur    := ebap_Requete;
    Result        := False;
    Exit;
  end;

  Ok := False;
  {S'il y a plusieurs lignes TTC, on ne génère pas de BAP}
  for n := 0 to T.Detail.Count - 1 do begin
    if (T.Detail[n].GetString('E_TYPEMVT') = 'TTC') and (T.Detail[n].GetInteger('E_NUMECHE') = 1) then begin
      if Ok then begin
        BapImpossible := True;
        CodeErreur    := ebap_MultiTTC;
        Result        := False;
        Exit;
      end
      else
        Ok := True;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TBonAPayer.ChargeTobCalcul(Clef : TClefPiece);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  J : TOB; {06/06/06}
  D : TOB;
  R : TOB;
  W : string;
  Mnt : Double;
  Axe : string;
  Sec : string;
  Ven : Double;
begin
  T := TOB.Create('****', nil, -1);
  J := TOB.Create('****', nil, -1);
  try
    {Récupération des informations nécessaires au calcul}
    with Clef do begin
      W := 'WHERE E_JOURNAL = "' + E_JOURNAL + '" AND ' +
           'E_EXERCICE = "' + E_EXERCICE + '" AND ' +
           'E_DATECOMPTABLE = "' + UsDateTime(E_DATECOMPTABLE) + '" AND ' +
           'E_QUALIFPIECE = "' + E_QUALIFPIECE + '" AND ' +
           'E_NUMEROPIECE = ' + IntToStr(E_NUMEROPIECE) + ' AND ' +
           'NOT (E_TYPEMVT IN ("TVA", "TPF")) ';
      T.LoadDetailFromSQL(REQUETEBAP + W);
    end;

    {Si le résultat de la requête n'est pas cohérent avce la génération des BAP, on sort}
    if not TesteRequete(T) then Exit;

    {JP 15/05/08 : FQ 22606 : pour se positionner sur la ligne de TTC
    D := T.Detail[0];
    {Reprend l'entête de la pièces
    SetEntete(D);}

    Mnt := 0;
    T.Detail.Sort('E_TYPEMVT;E_DATEECHEANCE;');
    D := T.FindFirst(['E_TYPEMVT'], ['TTC'], True);
    {JP 15/05/08 : FQ 22606 : déplacé : Reprend l'entête de la pièces}
    SetEntete(D);
    {Récupération des informations du tiers sur la ligne de première échéance}
    SetTTC(D);
    {Calcul du montant TTC total de toutes les échéances}
    while Assigned(D) do begin
      Mnt := Mnt + D.GetDouble('MONTANT');
      D := T.FindNext(['E_TYPEMVT'], ['TTC'], True);
    end;
    TobCalcul.AddChampSupValeur('MONTANT', Mnt);


    {Récupération des informations du général sur la Ligne de HT au montant le plus important}
    T.Detail.Sort('E_TYPEMVT;');
    Mnt := 0;
    R   := nil;
    D := T.FindFirst(['E_TYPEMVT'], ['HT'], True);
    while Assigned(D) do begin
      if Mnt < Abs(D.GetDouble('MONTANT')) then begin
        Mnt := Abs(D.GetDouble('MONTANT'));
        R := D;
      end;
      D := T.FindNext(['E_TYPEMVT'], ['HT'], True);
    end;
    {Mise à jour de la TobCalcul}
    SetHT(R);

    {06/06/06 : Duplisation de la TOB pour les FindFirst afin d'éviter des boucles infinies}
    J.Dupliquer(T, True, True);
    J.Detail.Sort('E_TYPEMVT;Y_AXE;Y_SECTION;');

    {Récupération des informations des sections sur la Ventilation la plus importante par axe (cumul par section)}
    T.Detail.Sort('E_TYPEMVT;Y_AXE;Y_SECTION;');
    Mnt := 0;
    Ven := 0;
    R   := nil;
    D   := T.FindFirst(['E_TYPEMVT'], ['HT'], True);
    if Assigned(D) then begin
      Axe := D.GetString('Y_AXE');
      Sec := D.GetString('Y_SECTION');
      while Assigned(D) do begin
        {Si la ligne HT n'a pas de ventilation on passe à la suivante}
        if Axe = '' then begin
          D := T.FindNext(['E_TYPEMVT'], ['HT'], True);
          if not Assigned(D) then Break;
          Axe := D.GetString('Y_AXE');
          Sec := D.GetString('Y_SECTION');
          Continue;
        end;

        {On change d'axe}
        if Axe <> D.GetString('Y_AXE') then begin
          {On compare la ventilation sur la dernière section de l'axe ...}
          if Mnt > Ven then
            {... On récupère la première ligne correspondant à cet axe et cette section}
            R := J.FindFirst(['E_TYPEMVT', 'Y_AXE', 'Y_SECTION'], ['HT', Axe, Sec], True);
          {Récupération des tables libres sur la section la "plus ventilée"}
          SetSection(R);

          R   := nil;
          Mnt := 0;
          Ven := 0;
          Axe := D.GetString('Y_AXE');
          Sec := D.GetString('Y_SECTION');
        end;

        {On change de section}
        if Sec <> D.GetString('Y_SECTION') then begin
          {On compare la ventilation sur la précédente section avec la ventilation la plus importante ...}
          if Mnt > Ven then begin
            Ven := Mnt;
            {... On récupère la première ligne correspondant à cet axe et cette section}
            R := J.FindFirst(['E_TYPEMVT', 'Y_AXE', 'Y_SECTION'], ['HT', Axe, Sec], True);
          end;

          Mnt := 0;
          Sec := D.GetString('Y_SECTION');
        end;
        Mnt := Mnt + Abs(D.GetDouble('VENTILATION'));

        D := T.FindNext(['E_TYPEMVT'], ['HT'], True);
      end;

      {Éventuelle récupération des données pour la dernière ventilation}
      if (Axe <> '') then begin
        if Mnt > Ven then
          R := J.FindFirst(['E_TYPEMVT', 'Y_AXE', 'Y_SECTION'], ['HT', Axe, Sec], True);
        SetSection(R);
      end;
    end
    else begin
      BapImpossible := True;
      CodeErreur    := ebap_HT;
    end;
  finally
    if Assigned(T) then FreeAndNil(T);
    if Assigned(J) then FreeAndNil(J);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TBonAPayer.SetEntete(T : TOB);
{---------------------------------------------------------------------------------------}
begin
  if not Assigned(T) then begin
    BapImpossible := True;
    CodeErreur    := ebap_Entete;
    Exit;
  end;

  TobCalcul := TOB.Create('µµµµ', nil, -1);
  try
    TobCalcul.AddChampSupValeur('E_JOURNAL', T.GetString('E_JOURNAL'));
    TobCalcul.AddChampSupValeur('E_EXERCICE', T.GetString('E_EXERCICE'));
    TobCalcul.AddChampSupValeur('E_ETABLISSEMENT', T.GetString('E_ETABLISSEMENT'));
    TobCalcul.AddChampSupValeur('E_DATECOMPTABLE', T.GetDateTime('E_DATECOMPTABLE'));
    TobCalcul.AddChampSupValeur('E_NUMEROPIECE', T.GetInteger('E_NUMEROPIECE'));
    TobCalcul.AddChampSupValeur('E_NATUREPIECE', T.GetString('E_NATUREPIECE'));
    {Théoriquement, on ne traite que les pièces normales ("N")}
    TobCalcul.AddChampSupValeur('E_QUALIFPIECE', T.GetString('E_QUALIFPIECE'));
    TobCalcul.AddChampSupValeur('E_REFINTERNE', T.GetString('E_REFINTERNE'));
    TobCalcul.AddChampSupValeur('E_DEVISE', T.GetString('E_DEVISE'));
    TobCalcul.AddChampSupValeur('E_SOCIETE', T.GetString('E_SOCIETE'));
    TobCalcul.AddChampSupValeur('EC_DOCGUID', T.GetString('EC_DOCGUID'));
    {09/06/06 : FQ 18321 : Les tables libres de l'écritures n'avaient pas été gérées !!}
    TobCalcul.AddChampSupValeur('E_TABLE0', T.GetString('E_TABLE0'));
    TobCalcul.AddChampSupValeur('E_TABLE1', T.GetString('E_TABLE1'));
    TobCalcul.AddChampSupValeur('E_TABLE2', T.GetString('E_TABLE2'));
    TobCalcul.AddChampSupValeur('E_TABLE3', T.GetString('E_TABLE3'));
  except
    on E : Exception do begin
      BapImpossible := True;
      CodeErreur    := ebap_Entete;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TBonAPayer.SetTTC(T : TOB);
{---------------------------------------------------------------------------------------}
begin
  if not Assigned(T) then begin
    BapImpossible := True;
    CodeErreur    := ebap_TTC;
    Exit;
  end;
  try
    TobCalcul.AddChampSupValeur('E_DATEECHEANCE', T.GetDateTime('E_DATEECHEANCE'));
    TobCalcul.AddChampSupValeur('T_AUXILIAIRE',   T.GetString('T_AUXILIAIRE'));
    TobCalcul.AddChampSupValeur('E_PIECETP',      T.GetString('E_PIECETP')); // SBO 22/05/2007 il manquait la référence au TP à priori
    TobCalcul.AddChampSupValeur('E_BLOCNOTE'    , T.GetValue('E_BLOCNOTE')); {JP 28/01/08 : Ajout de E_BLOCNOTE}
    {Pour mémoriser le compte collectif}
    TobCalcul.AddChampSupValeur('T_GENERAL', T.GetString('G_GENERAL'));
    TobCalcul.AddChampSupValeur('T_TABLE0', T.GetString('T_TABLE0'));
    TobCalcul.AddChampSupValeur('T_TABLE1', T.GetString('T_TABLE1'));
    TobCalcul.AddChampSupValeur('T_TABLE2', T.GetString('T_TABLE2'));
    TobCalcul.AddChampSupValeur('T_TABLE3', T.GetString('T_TABLE3'));
    TobCalcul.AddChampSupValeur('T_TABLE4', T.GetString('T_TABLE4'));
    TobCalcul.AddChampSupValeur('T_TABLE5', T.GetString('T_TABLE5'));
    TobCalcul.AddChampSupValeur('T_TABLE6', T.GetString('T_TABLE6'));
    TobCalcul.AddChampSupValeur('T_TABLE7', T.GetString('T_TABLE7'));
    TobCalcul.AddChampSupValeur('T_TABLE8', T.GetString('T_TABLE8'));
    TobCalcul.AddChampSupValeur('T_TABLE9', T.GetString('T_TABLE9'));

  except
    on E : Exception do begin
      BapImpossible := True;
      CodeErreur    := ebap_TTC;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TBonAPayer.SetHT(T : TOB);
{---------------------------------------------------------------------------------------}
begin
  if not Assigned(T) then begin
    BapImpossible := True;
    CodeErreur    := ebap_HT;
    Exit;
  end;
  try
    TobCalcul.AddChampSupValeur('G_TABLE0', T.GetString('G_TABLE0'));
    TobCalcul.AddChampSupValeur('G_TABLE1', T.GetString('G_TABLE1'));
    TobCalcul.AddChampSupValeur('G_TABLE2', T.GetString('G_TABLE2'));
    TobCalcul.AddChampSupValeur('G_TABLE3', T.GetString('G_TABLE3'));
    TobCalcul.AddChampSupValeur('G_TABLE4', T.GetString('G_TABLE4'));
    TobCalcul.AddChampSupValeur('G_TABLE5', T.GetString('G_TABLE5'));
    TobCalcul.AddChampSupValeur('G_TABLE6', T.GetString('G_TABLE6'));
    TobCalcul.AddChampSupValeur('G_TABLE7', T.GetString('G_TABLE7'));
    TobCalcul.AddChampSupValeur('G_TABLE8', T.GetString('G_TABLE8'));
    TobCalcul.AddChampSupValeur('G_TABLE9', T.GetString('G_TABLE9'));
    TobCalcul.AddChampSupValeur('G_GENERAL', T.GetString('G_GENERAL'));
  except
    on E : Exception do begin
      BapImpossible := True;
      CodeErreur    := ebap_HT;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TBonAPayer.SetSection(T : TOB);
{---------------------------------------------------------------------------------------}
begin
  if not Assigned(T) then begin
    BapImpossible := True;
    CodeErreur    := ebap_Section;
    Exit;
  end;

  try
    TobCalcul.AddChampSupValeur('Y_SECTION' + T.GetString('Y_AXE'), T.GetString('Y_SECTION'));
    TobCalcul.AddChampSupValeur('S_TABLE0'  + T.GetString('Y_AXE'), T.GetString('S_TABLE0'));
    TobCalcul.AddChampSupValeur('S_TABLE1'  + T.GetString('Y_AXE'), T.GetString('S_TABLE1'));
    TobCalcul.AddChampSupValeur('S_TABLE2'  + T.GetString('Y_AXE'), T.GetString('S_TABLE2'));
    TobCalcul.AddChampSupValeur('S_TABLE3'  + T.GetString('Y_AXE'), T.GetString('S_TABLE3'));
    TobCalcul.AddChampSupValeur('S_TABLE4'  + T.GetString('Y_AXE'), T.GetString('S_TABLE4'));
    TobCalcul.AddChampSupValeur('S_TABLE5'  + T.GetString('Y_AXE'), T.GetString('S_TABLE5'));
    TobCalcul.AddChampSupValeur('S_TABLE6'  + T.GetString('Y_AXE'), T.GetString('S_TABLE6'));
    TobCalcul.AddChampSupValeur('S_TABLE7'  + T.GetString('Y_AXE'), T.GetString('S_TABLE7'));
    TobCalcul.AddChampSupValeur('S_TABLE8'  + T.GetString('Y_AXE'), T.GetString('S_TABLE8'));
    TobCalcul.AddChampSupValeur('S_TABLE9'  + T.GetString('Y_AXE'), T.GetString('S_TABLE9'));
  except
    on E : Exception do begin
      BapImpossible := True;
      CodeErreur    := ebap_Section;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TBonAPayer.GetTypeVisa;
{---------------------------------------------------------------------------------------}
var
  S : string;
  T : TOB;
  n : Integer;
begin

  S := 'SELECT * FROM CPTYPEVISA WHERE ' + GetClauseWhere;

  if BapImpossible then Exit;

  T := TOB.Create('****', nil, -1);
  try
    T.LoadDetailFromSQL(S);

    if T.Detail.Count = 0 then begin
      BapImpossible := True;
      CodeErreur    := ebap_PasDeVisa;
      Exit;
    end;

    {Cherche s'il est possible de trouver un type de visa correspondant à la pièce
     à partir des filtres sur les zones libres}
    TesteZonesLibres(T);

    {Cherche s'il est possible de trouver un type de visa correspondant à la pièce
     à partir des fourchettes de comptes}
    if not BapImpossible then TesteComptes(T);

    if not BapImpossible then begin
      if T.Detail.Count >= 1 then begin
       if TraitementMasse then
          SetTobBap(T.Detail[0])
        else
          for n := 0 to T.Detail.Count - 1 do
            SetTobBap(T.Detail[n]);
      end;
    end;

    if TraitementMasse then begin
      FreeAndNil(TobCalcul);
    end

    else if not BapImpossible and (TobBap.Detail.Count = 0) then begin
      BapImpossible := True;
      CodeErreur := ebap_GenereBap;
    end;

  finally
    if Assigned(T) then FreeAndNil(T);
  end;
end;

{Cherche s'il est possible de trouver un type de visa correspondant à la pièce
 à partir des fourchettes de comptes
{---------------------------------------------------------------------------------------}
procedure TBonAPayer.TesteComptes(T : TOB);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  D : TOB;
  APrendre : string;
  AExclure : string;
  AP, AE   : string;
  Ok       : Boolean;
begin
  for n := T.Detail.Count - 1 downto 0 do begin
    D := T.Detail[n];
    Ok := False;
    APrendre := D.GetString('CTI_COMPTE');
    AExclure := D.GetString('CTI_EXCLUSION');
    {25/12/05 : Si rien n'est paramétré sur les comptes, on passe au Type de visa suivant}
    if (APrendre = '') and (AExclure = '') then Continue;

    while APrendre <> '' do begin
      AP := ReadTokenSt(APrendre);
      AE := ReadTokenSt(AExclure);
      Ok := CompareGeneral(TobCalcul.GetString('G_GENERAL'), AP, AE);
      if Ok then Break;
    end;
    if not Ok then FreeAndNil(D);
  end;

  {Si la tob a été vidée}
  if T.Detail.Count = 0 then begin
    BapImpossible := True;
    CodeErreur    := ebap_PasDeVisa;
  end;
end;

{Cherche s'il est possible de trouver un type de visa correspondant à la pièce
 à partir des filtres sur les zones libres
{---------------------------------------------------------------------------------------}
procedure TBonAPayer.TesteZonesLibres(T : TOB);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  D : TOB;

    {-------------------------------------------------------------------------}
    function ZoneOk(Ind : Char) : Boolean;
    {-------------------------------------------------------------------------}
    var
      ch : string;
    begin
      Result := True;
      if D.GetString('CTI_TEXTELIBRE' + Ind) <> '' then begin
        if (D.GetString('CTI_TYPELIBRE' + Ind) = tyt_Devise) then
          Result := CompareAvecJoker(D.GetString('CTI_TEXTELIBRE' + Ind), TobCalcul.GetString('E_DEVISE'))

        {Tables libres sur les Généraux, Tiers ou écritures}
        else if D.GetString('CTI_TYPELIBRE' + Ind) = tyt_RefInterne then
          Result := CompareAvecJoker(D.GetString('CTI_TEXTELIBRE' + Ind), TobCalcul.GetString('E_REFINTERNE'))

        {Tables libres sur les sections}
        else if (D.GetString('CTI_TYPELIBRE' + Ind) = tyt_TLSection) then begin
          ch := D.GetString('CTI_CODELIBRE' + Ind);
          if Length(ch) = 3 then
            {08/06/06 : FQ 18321 : Correction de la gestion de l'axe
            Result := (TobCalcul.GetString('Y_AXE') = D.GetString('CTI_AXE' + Ind)) and
                      CompareAvecJoker(TobCalcul.GetString('S_TABLE' + ch[3]), D.GetString('CTI_TEXTELIBRE' + Ind))}
            Result := CompareAvecJoker(TobCalcul.GetString('S_TABLE' + ch[3] + D.GetString('CTI_AXE' + Ind)), D.GetString('CTI_TEXTELIBRE' + Ind))
          else begin
            {Ici on n'interdit pas la génération des BAP, mais on avertit
             d'un éventuel problème de paramétrage}
            CodeErreur := ebap_ParamAna;
            Result     := False;
          end;
        end

        {Tables libres sur les Généraux, Tiers ou écritures}
        else begin
          ch := D.GetString('CTI_CODELIBRE' + Ind);
          if Length(ch) = 3 then
            Result := CompareAvecJoker(TobCalcul.GetString(ch[1] + '_TABLE' + ch[3]), D.GetString('CTI_TEXTELIBRE' + Ind))
          else begin
            {Ici on n'interdit pas la génération des BAP, mais on avertit
             d'un éventuel problème de paramétrage}
            case ch[1] of
              'E' : CodeErreur := ebap_ParamEcr;
              'G' : CodeErreur := ebap_ParamGen;
              'T' : CodeErreur := ebap_ParamAux;
            end;
            Result := False;
          end;
        end;
      end;
    end;
var
  Ok : Boolean;
begin
  for n := T.Detail.Count - 1 downto 0 do begin
    D := T.Detail[n];
    {07/12/06 : FQ 19164 : nouveau champ CTI_AUCHOIX qui permet de dire si l'on fait un et ou bien un ou}
    if D.GetString('CTI_AUCHOIX') = 'X' then
      Ok := ZoneOk('1') or ZoneOk('2') or ZoneOk('3')
    else
      Ok := ZoneOk('1') and ZoneOk('2') and ZoneOk('3');
    if not Ok then FreeAndNil(D);;
  end;

  {Si la tob a été vidée}
  if T.Detail.Count = 0 then begin
    BapImpossible := True;
    CodeErreur    := ebap_PasDeVisa;
  end;
end;

{---------------------------------------------------------------------------------------}
function TBonAPayer.GetClauseWhere : string;
{---------------------------------------------------------------------------------------}
var
  S : string;

      {-------------------------------------------------------------------}
      procedure _AddWhere(ch : string);
      {-------------------------------------------------------------------}
      begin
        if S = '' then S := ch
                  else S := S + ' AND ' + ch;
      end;

begin
  Result := '';

  if not Assigned(TobCalcul) then begin
    BapImpossible := True;
    CodeErreur    := ebap_WhereVisa;
    Exit;
  end;

  try
    {La nature de pièce et l'établissement sont obligatoirement paramétrés dans CPTYPEVISA
     30/05/07 : On rend l'établissement facultatif à la demande de SIC
     20/06/07 : FQ 20778 : Idem pour les natures à la demande de SIC}
    if TobCalcul.GetString('E_NATUREPIECE') <> '' then
      _AddWhere('(CTI_NATUREPIECE = "' + TobCalcul.GetString('E_NATUREPIECE') + '" OR ' +
      {20/06/07 : Ou bien nature de pièce non paramétrée dans le type de visa}
        'CTI_NATUREPIECE = "" OR CTI_NATUREPIECE IS NULL)');
        
    if TobCalcul.GetString('E_ETABLISSEMENT') <> '' then
      _AddWhere('(CTI_ETABLISSEMENT = "' + TobCalcul.GetString('E_ETABLISSEMENT') + '" OR ' +
      {30/05/07 : Ou bien établissement non paramétré dans le type de visa}
        'CTI_ETABLISSEMENT = "" OR CTI_ETABLISSEMENT IS NULL)');

    {Filtre sur les montants}
    _AddWhere('((' + StrFPoint(TobCalcul.GetDouble('MONTANT')) + ' BETWEEN CTI_MONTANTMIN AND CTI_MONTANTMAX) OR (CTI_MONTANTMIN + CTI_MONTANTMAX) = 0)');

  except
    on E : Exception do begin
      BapImpossible := True;
      CodeErreur    := ebap_WhereVisa;
      Exit;
    end;
  end;

  if S = '' then begin
    BapImpossible := True;
    CodeErreur    := ebap_WhereVisa;
  end
  else
    Result := S;
end;

{---------------------------------------------------------------------------------------}
procedure TBonAPayer.SetTobBap(T : TOB);
{---------------------------------------------------------------------------------------}
var
  F : TOB;
begin
  if TraitementMasse then F := TOB.Create('CPBONSAPAYER', TobMasse, -1)
                     else F := TOB.Create('CPBONSAPAYER', TobBap  , -1);
  try
    F.SetString('BAP_CODEVISA'       , T.GetString('CTI_CODEVISA'));
    F.SetString('BAP_CIRCUITBAP'     , T.GetString('CTI_CIRCUITBAP'));
    F.SetDateTime('BAP_DATECREATION' , Now);
    F.SetDateTime('BAP_DATEMODIF'    , Now);
    F.SetString('BAP_STATUTBAP'      , sbap_Encours);
    F.SetString('BAP_CREATEUR'       , V_PGI.User);
    F.SetString('BAP_MODIFICATEUR'   , V_PGI.User);
    F.SetDateTime('BAP_DATEMAIL'     , iDate1900);
    F.SetDateTime('BAP_DATERELANCE1' , iDate1900);
    F.SetDateTime('BAP_DATERELANCE2' , iDate1900);
    F.SetString('BAP_RELANCE1'       , '-'); {Mail de relance au viseur 1}
    F.SetString('BAP_RELANCE2'       , '-'); {Mail de relance au viseur 2}
    F.SetString('BAP_ALERTE1'        , '-'); {Mail d'alerte au viseur 1}
    F.SetString('BAP_ALERTE2'        , '-'); {Mail d'alerte au viseur 2}
    F.SetString('BAP_SOCIETE'        , TobCalcul.GetString('E_SOCIETE'));
    F.SetString('BAP_JOURNAL'        , TobCalcul.GetString('E_JOURNAL'));
    F.SetString('BAP_EXERCICE'       , TobCalcul.GetString('E_EXERCICE'));
    F.SetString('BAP_ETABLISSEMENT'  , TobCalcul.GetString('E_ETABLISSEMENT'));
    F.SetString('BAP_IDGED'          , TobCalcul.GetString('EC_DOCGUID')); {Attention aux tiers payeurs !!!}
    F.SetDateTime('BAP_DATECOMPTABLE', TobCalcul.GetDateTime('E_DATECOMPTABLE'));
    F.SetDateTime('BAP_DATEECHEANCE' , TobCalcul.GetDateTime('E_DATEECHEANCE'));
    F.SetInteger ('BAP_NUMEROPIECE'  , TobCalcul.GetInteger('E_NUMEROPIECE'));
    {JP 28/01/08 : Demande de SIC : Reprise de E_BLOCNOTE}
    F.PutValue   ('BAP_BLOCNOTE'     , TobCalcul.GetValue('E_BLOCNOTE'));

    GetCircuit(1, F);  {Remplit les champs concernant le circuit de validation}
    {$IFNDEF EAGLSERVER}
    GetTiersPayeur(F); {Remplit les champs concernant les Tiers Payeurs}
    {$ENDIF EAGLSERVER}
  except
    on E : Exception do begin
      BapImpossible := True;
      CodeErreur    := ebap_CreateBap;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TBonAPayer.ForceBap(ClePiece : TClefPiece; TypeVisa, Circuit : string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  T := TOB.Create('***', nil, -1);
  T.AddChampSupValeur('CTI_CODEVISA', TypeVisa);
  T.AddChampSupValeur('CTI_CIRCUITBAP', Circuit);
  if TobCalcul.GetInteger('E_NUMEROPIECE') <> ClePiece.E_NUMEROPIECE then
    ChargeTobCalcul(ClePiece);
  SetTobBap(T);
end;

{---------------------------------------------------------------------------------------}
procedure TBonAPayer.GetCircuit(NoOrdre : Integer; var T : TOB);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  try
    Q := OpenSQL('SELECT * FROM CPCIRCUIT WHERE CCI_CIRCUITBAP = "' + T.GetString('BAP_CIRCUITBAP') +
                 '" AND CCI_NUMEROORDRE >= ' + IntToStr(NoOrdre) + ' ORDER BY CCI_NUMEROORDRE', True);
    try
      if Q.EOF then begin
        BapImpossible := True;
        CodeErreur    := ebap_Circuit;
      end
      else
      try
        T.SetInteger ('BAP_NUMEROORDRE' , NoOrdre);
        T.SetString('BAP_VISEUR'        , ''); {Code du viseur qui aura validé le BAP : Viseur 1  ou viseur 2}
        T.SetString('BAP_VISEUR1'       , Q.FindField('CCI_VISEUR1').AsString);{Code du viseur qui doit valider le BAP}
        T.SetString('BAP_VISEUR2'       , Q.FindField('CCI_VISEUR2').AsString);{Code du viseur de substitution}
        {Calcule les délais en fonction de la date d'échéance de la facture et du paramétrage du circuit}
        if not Assigned(TobCalcul) then
          CalculEcheBap(Q, T.GetDateTime('BAP_DATEECHEANCE'), T) {Process Server, ComSx}
        else
          CalculEcheBap(Q, TobCalcul.GetDateTime('E_DATEECHEANCE'), T); {Comptabilité}

      except
        on E : Exception do begin
          BapImpossible := True;
          CodeErreur    := ebap_Circuit;
        end;
      end;
    finally
      Ferme(Q);
    end;
  except
    on E : Exception do begin
      BapImpossible := True;
      CodeErreur    := ebap_Circuit;
      raise E;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TBonAPayer.CalculEcheBap(Q : TQuery; DateFact : TDateTime; var T : TOB);
{---------------------------------------------------------------------------------------}
var
  NbJours : Integer;
  Delais  : Integer;
  Rapport : Double;
begin
  NbJours := 0;
  Delais  := Round(DateFact - Now);
  try
    while not Q.Eof do begin
      NbJours := NbJours + Q.FindField('CCI_NBJOUR').AsInteger;
      Q.Next;
    end;
    {Théoriquement, le premier enregistrement du Query}
    Q.First;

    if Delais > 0 then begin
      Rapport := NbJours / Delais;
      {Le nombre de jours du circuit est supérieur au délais disponible avant l'échéance de la facture ...}
      if Rapport > 1 then begin
        {On arrondit vers 0, la division en le paramétrage du circuit et le rapport}
        T.SetInteger('BAP_NBJOUR'       , Trunc(Q.FindField('CCI_NBJOUR').AsInteger / Rapport));
        T.SetDateTime('BAP_ECHEANCEBAP' , Date + T.GetInteger('BAP_NBJOUR'));
      end

      {Le nombre de jours du circuit est inférieur ou égal au délais disponible avant
       l'échéance de la facture : on reprend ce qui est paramétré dans le circuit}
      else begin
        T.SetInteger ('BAP_NBJOUR'       , Q.FindField('CCI_NBJOUR').AsInteger);
        T.SetDateTime('BAP_ECHEANCEBAP' , Date + Q.FindField('CCI_NBJOUR').AsInteger);
      end;
    end
    else begin
      T.SetInteger('BAP_NBJOUR'       , 0);
      T.SetDateTime('BAP_ECHEANCEBAP' , Date);
    end;
  except
    on E : Exception do begin
      BapImpossible := True;
      CodeErreur    := ebap_CalDelais;
    end;
  end;
end;

{$IFNDEF EAGLSERVER}
{Problèmes : Le EC_DOCGUID est-il celui de la facture ou du tiers payeur
{---------------------------------------------------------------------------------------}
procedure TBonAPayer.GetTiersPayeur(var T : TOB; TobEcr : TOB = nil);
{---------------------------------------------------------------------------------------}
var
  Clef : string;
  RecC : TClefPiece;
begin
  if TobEcr = nil then
    Clef := TobCalcul.GetString('E_PIECETP')
  else
    Clef := TobEcr.GetString('E_PIECETP');

  try
    if Clef <> '' then begin
      RecC := ClefFromPieceTP(Clef);
      Clef := EncodeClefTP(RecC);
      T.SetString('BAP_CLEFFACTURE', Clef);
      T.SetString('BAP_TIERSPAYEUR', 'X');
    end
    else
      T.SetString('BAP_TIERSPAYEUR', '-');
  except
    on E : Exception do begin
      BapImpossible := True;
      CodeErreur    := ebap_TiersPaye;
    end;
  end;
end;
{$ENDIF EAGLSERVER}

{ TObjCircuit }

{---------------------------------------------------------------------------------------}
constructor TObjCircuit.Create(Circuit : string);
{---------------------------------------------------------------------------------------}
begin
  TobCircuit := TOB.Create('%%%', nil, -1);
  TobCircuit.LoadDetailFromSQL('SELECT * FROM CPCIRCUIT WHERE CCI_CIRCUITBAP = "' + Circuit + '"');
end;

{---------------------------------------------------------------------------------------}
destructor TObjCircuit.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobCircuit) then FreeAndNil(TobCircuit);
  inherited;
end;

{---------------------------------------------------------------------------------------}
function TObjCircuit.GetDelais(Num : Integer) : Integer;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  Result := -1;
  T := TobCircuit.FindFirst(['CCI_NUMEROORDRE'], [Num], True);
  if Assigned(T) then
    Result := T.GetInteger('CCI_NBJOUR');
end;

{---------------------------------------------------------------------------------------}
function TObjCircuit.GetViseur1(Num : Integer) : string;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  T := TobCircuit.FindFirst(['CCI_NUMEROORDRE'], [Num], True);
  if Assigned(T) then
    Result := T.GetString('CCI_VISEUR1');
end;

{---------------------------------------------------------------------------------------}
function TObjCircuit.GetViseur2(Num : Integer) : string;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  T := TobCircuit.FindFirst(['CCI_NUMEROORDRE'], [Num], True);
  if Assigned(T) then
    Result := T.GetString('CCI_VISEUR2');
end;

{ TObjMail }

{---------------------------------------------------------------------------------------}
constructor TObjMail.Create(ClePiece : TClefPiece);
{---------------------------------------------------------------------------------------}
begin
  TobMail  := TOB.Create('ùùù', nil, -1);
  TobTypeM := TOB.Create('***', nil, -1);
  TobUser  := TOB.Create('£££', nil, -1);
  TobBap  := TOB.Create('CPBONSAPAYER', nil, -1);

  TobBap.SetString  ('BAP_JOURNAL'      , ClePiece.E_JOURNAL);
  TobBap.SetString  ('BAP_EXERCICE'     , ClePiece.E_EXERCICE);
  TobBap.SetDateTime('BAP_DATECOMPTABLE', ClePiece.E_DATECOMPTABLE);
  TobBap.SetInteger ('BAP_NUMEROPIECE'  , ClePiece.E_NUMEROPIECE);
  TobBap.SetInteger ('BAP_NUMEROORDRE'  , ClePiece.E_NUMLIGNE);
  TobBap.LoadDB;

  TobTypeM.LoadDetailFromSQL('SELECT * FROM CPMAILBAP');
end;

{---------------------------------------------------------------------------------------}
function TObjMail.CreerMail(Niveau : TNiveauMail) : string;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  Result := '';
  if (TobBap.GetString('BAP_VISEUR1') = '') or (TobBap.GetString('BAP_VISEUR1') = '') then begin
    Result := TraduireMemoire('Impossible de charger le bon à payer.');
    Exit;
  end;

  TobUser.LoadDetailFromSQL('SELECT US_EMAIL,US_UTILISATEUR FROM UTILISAT WHERE US_UTILISATEUR IN ' + '("' +
                            TobBap.GetString('BAP_VISEUR1') + '", "' + TobBap.GetString('BAP_VISEUR2') + '")');
  if TobTypeM.Detail.Count < 2 then begin
    Result := TraduireMemoire('Impossible de charger les utilisateurs.');
    Exit;
  end;

  if TobTypeM.Detail.Count = 0 then begin
    Result := TraduireMemoire('Impossible de charger le paramétrage des mails.');
    Exit;
  end;

  s := 'SELECT * FROM CPTACHEBAP WHERE (CTA_DESTINATAIRE ';
  case Niveau of
    nma_Alerte   : s := s + '= "' + TobBap.GetString('BAP_VISEUR1') + '" AND CTA_NATUREMAIL = "' + nam_Alerte + '" ' +
                            'AND CTA_DATEENVOI = "' + UsDateTime(Date) + '")';

    nma_Relance1 : s := s + '= "' + TobBap.GetString('BAP_VISEUR1') + '" AND CTA_NATUREMAIL = "' + nam_Relance + '" ' +
                            'AND CTA_DATEENVOI = "' + UsDateTime(Date) + '") OR (CTA_DESTINATAIRE = "' +
                            TobBap.GetString('BAP_VISEUR2') + '" AND CTA_NATUREMAIL = "' + nam_Suppleant + '" ' +
                            'AND CTA_DATEENVOI = "' + UsDateTime(Date) + '")';

    nma_Relance2 : s := s + 'IN ("' + TobBap.GetString('BAP_VISEUR1') + '", "' + TobBap.GetString('BAP_VISEUR2') +
                            '") AND CTA_NATUREMAIL = "' + nam_Relance + '" ' +
                            'AND CTA_DATEENVOI = "' + UsDateTime(Date) + '")';
    else
      Exit;
  end;
  TobMail.LoadDetailFromSQL(s);
  Result := RemplirTob(Niveau);
  if Result = '' then Result := MajBap(Niveau);
end;

{---------------------------------------------------------------------------------------}
function TObjMail.RemplirTob(Niveau : TNiveauMail) : string;
{---------------------------------------------------------------------------------------}
var
  T1 : TOB;
  T2 : TOB;
  TU : TOB;
  TM : TOB;
begin
  Result := '';
  T2 := nil;

  T1 := TobMail.FindFirst(['CTA_DESTINATAIRE'], [TobBap.GetString('BAP_VISEUR1')], True);
  if Niveau <> nma_Alerte then
    T2 := TobMail.FindFirst(['CTA_DESTINATAIRE'], [TobBap.GetString('BAP_VISEUR2')], True);

  TU := TobUser.FindFirst(['US_UTILISATEUR'], [TobBap.GetString('BAP_VISEUR1')], True);
  if not Assigned(TU) then begin
    Result := TraduireMemoire('Utilisateur introuvable : ' + TobBap.GetString('BAP_VISEUR1'));
    Exit;
  end;

  if TU.GetString('US_EMAIL') = '' then begin
    Result := TraduireMemoire('Pas d''adresse électronique pour l''utilisateur : ' + TobBap.GetString('BAP_VISEUR1'));
    Exit;
  end;

  TM := TobTypeM.FindFirst(['CMA_NATUREMAIL'], [NatMailFromNiveau(Niveau, '1')], True);
  if not Assigned(TM) then begin
    Result := TraduireMemoire('Modèle de mail introuvable : ' + NatMailFromNiveau(Niveau, '1'));
    Exit;
  end;

  if Assigned(T1) then begin
    T1.SetInteger('CTA_NBVISA'     , T1.GetInteger('CTA_NBVISA') + 1);
    T1.UpdateDB;
  end
  else begin
    T1 := TOB.Create('CPTACHEBAP'  , TobMail, -1);
    T1.SetString('CTA_BLOCNOTE'    , TM.GetString('CMA_BLOCNOTE'));
    T1.SetString('CTA_CIRCUITBAP'  , TobBap.GetString('BAP_CIRCUITBAP'));
    T1.SetString('CTA_CODEEVISA'   , TobBap.GetString('BAP_CODEEVISA'));
    T1.SetString('CTA_CODEMAIL'    , TM.GetString('CMA_CODEMAIL'));
    T1.SetDateTime('CTA_DATEENVOI' , Date);
    T1.SetString('CTA_DESTINATAIRE', TobBap.GetString('BAP_VISEUR1'));
    T1.SetString('CTA_EMAIL'       , TU.GetString('US_EMAIL'));
    T1.SetString('CTA_ENVOYE'      , '-');
    T1.SetString('CTA_NATUREMAIL'  , NatMailFromNiveau(Niveau, '1'));
    T1.SetInteger('CTA_NBVISA'     , 1);
    T1.InsertDB(nil);
  end;

  if (Niveau <> nma_Alerte) then begin
    TU := TobUser.FindFirst(['US_UTILISATEUR'], [TobBap.GetString('BAP_VISEUR2')], True);
    if not Assigned(TU) then begin
      Result := TraduireMemoire('Utilisateur introuvable : ' + TobBap.GetString('BAP_VISEUR2'));
      Exit;
    end;

    TM := TobTypeM.FindFirst(['CMA_NATUREMAIL'], [NatMailFromNiveau(Niveau, '2')], True);
    if not Assigned(TM) then begin
      Result := TraduireMemoire('Modèle de mail introuvable : ' + NatMailFromNiveau(Niveau, '2'));
      Exit;
    end;

    if Assigned(T2) then begin
      T2.SetInteger('CTA_NBVISA'     , T2.GetInteger('CTA_NBVISA') + 1);
      T2.UpdateDB;
    end
    else begin
      if TU.GetString('US_EMAIL') = '' then begin
        Result := TraduireMemoire('Pas d''adresse électronique pour l''utilisateur : ' + TobBap.GetString('BAP_VISEUR2'));
        Exit;
      end;
      T2 := TOB.Create('CPTACHEBAP'  , TobMail, -1);
      T2.SetString('CTA_BLOCNOTE'    , TM.GetString('CMA_BLOCNOTE'));
      T2.SetString('CTA_CIRCUITBAP'  , TobBap.GetString('BAP_CIRCUITBAP'));
      T2.SetString('CTA_CODEEVISA'   , TobBap.GetString('BAP_CODEEVISA'));
      T2.SetString('CTA_CODEMAIL'    , TM.GetString('CMA_CODEMAIL'));
      T2.SetDateTime('CTA_DATEENVOI' , Date);
      T2.SetString('CTA_DESTINATAIRE', TobBap.GetString('BAP_VISEUR2'));
      T2.SetString('CTA_EMAIL'       , TU.GetString('US_EMAIL'));
      T2.SetString('CTA_ENVOYE'      , '-');
      T2.SetString('CTA_NATUREMAIL'  , NatMailFromNiveau(Niveau, '2'));
      T2.SetInteger('CTA_NBVISA'     , 1);
      T2.InsertDb(nil);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
destructor TObjMail.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobBap  ) then FreeAndNil(TobBap);
  if Assigned(TobMail ) then FreeAndNil(TobMail);
  if Assigned(TobTypeM) then FreeAndNil(TobTypeM);
  if Assigned(TobUser ) then FreeAndNil(TobUser);
  inherited;
end;

{---------------------------------------------------------------------------------------}
function TObjMail.NatMailFromNiveau(Niveau : TNiveauMail; Viseur : Char): string;
{---------------------------------------------------------------------------------------}
begin
  case Niveau of
    nma_Alerte   : if Viseur = '1' then Result := nam_Alerte
                                   else Result := '';
    nma_Relance1 : if Viseur = '1' then Result := nam_Relance
                                   else Result := nam_Suppleant;
    nma_Relance2 : if Viseur = '1' then Result := nam_Relance
                                   else Result := nam_Relance;
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjMail.MajBap(Niveau : TNiveauMail) : string;
{---------------------------------------------------------------------------------------}
begin
  if Niveau = nma_Alerte then begin
    TobBap.SetString('BAP_ALERTE1', 'X');
    TobBap.SetDateTime('BAP_DATEMAIL', Date);
  end
  else if Niveau = nma_Relance1 then begin
    TobBap.SetString('BAP_RELANCE1', 'X');
    TobBap.SetString('BAP_ALERTE2', 'X');
    TobBap.SetDateTime('BAP_DATERELANCE1', Date);
  end
  else if Niveau = nma_Relance2 then begin
    TobBap.SetString('BAP_RELANCE2', 'X');
    TobBap.SetDateTime('BAP_DATERELANCE2', Date);
  end;
  TobBap.UpdateDB;
end;


{ TObjEnvoiMail }

{---------------------------------------------------------------------------------------}
constructor TObjEnvoiMail.Create(Nat, User : string; DtEnvoi : TDateTime; T : TOB = nil);
{---------------------------------------------------------------------------------------}
begin
  if T = nil then begin
    TobMail := TOB.Create('CPTACHEBAP', nil, -1);
    TobMail.SetString('CTA_NATUREMAIL', Nat);
    TobMail.SetString('CTA_DESTINATAIRE', User);
    TobMail.SetDateTime('CTA_DATEENVOI', DtEnvoi);
    TobMail.LoadDB;
  end
  else
    TobMail := T;
end;

{---------------------------------------------------------------------------------------}
destructor TObjEnvoiMail.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobMail) then FreeAndNil(TobMail);
  inherited;
end;

{---------------------------------------------------------------------------------------}
function TObjEnvoiMail.EnvoieMail : Boolean;
{---------------------------------------------------------------------------------------}
var
  sServer : string;
  sFrom   : string;
  sTitre  : string;
  {$IFDEF EAGLSERVER}
  sBody   : string;
  {$ELSE}
  Corps   : HTStringList;
  {$ENDIF EAGLSERVER}
begin
  Result  := False;
  sServer := '';
  sFrom   := '';

  with TRegistry.Create do begin
    RootKey := HKEY_LOCAL_MACHINE;
    if OpenKey('SOFTWARE\CEGID_RM\PgiService\Journal', False) then begin
      if ValueExists('smtpServer') then sServer := ReadString('smtpServer');
      if ValueExists('smtpFrom')   then sFrom   := ReadString('smtpFrom');
      CloseKey;
    end;
    Free;
  end;

  if sServer = '' then Exit;
  if sFrom   = '' then Exit;
  if TobMail.GetString('CTA_EMAIL') = '' then Exit;

  sTitre := TobMail.GetString('CTA_NBVISA');

       if TobMail.GetString('CTA_NATUREMAIL') = nam_Alerte    then sTitre := sTitre + TraduireMemoire(' bon(s) à payer sont à valider')
  else if TobMail.GetString('CTA_NATUREMAIL') = nam_Bloque    then sTitre := sTitre + TraduireMemoire(' bon(s) à payer ont été bloqués')
  else if TobMail.GetString('CTA_NATUREMAIL') = nam_Refuse    then sTitre := sTitre + TraduireMemoire(' bon(s) à payer ont été refusés')
  else if TobMail.GetString('CTA_NATUREMAIL') = nam_Suppleant then sTitre := sTitre + TraduireMemoire(' bon(s) à payer sont à valider en tant que suppléant')
  else if TobMail.GetString('CTA_NATUREMAIL') = nam_Relance   then sTitre := TraduireMemoire('URGENT : ' + sTitre + ' bons à payer sont à valider');
  {$IFDEF EAGLSERVER}
  sBody := TobMail.GetString('CTA_BLOCNOTE');


  (*
   "C:\Program Files\Cegid\Cegid SIC\app\eCBPortal.exe" /SERVER=SIC-DEV3 /DOSSIER=CEGIDTESTV2008 /USER=PASTERIS /

  <html xmlns="http://www.w3.org/TR/REC-html40">
  <body>
  <div class=Section1>
  <p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial;color:navy'>
  <a href="file:///C:\Program%20Files\CEGID\Cegid%20Sic\app\eCBPortal.exe"></a>
  <a href="%22C:/CWS/eCBPortal.exe%22%20/SERVER=JPASTERIS:8080%20/DOSSIER=NOVATRES%20/USER=PASTERIS%20/">Cliquez ici</a><o:p></o:p></span></p>
  </div>
  </body>
  </html>
  *)

  if SendMailSmtp(sServer, sFrom, TobMail.GetString('CTA_EMAIL'), sTitre, sBody) then begin

    TobMail.SetString('CTA_ENVOYE', 'X');
    Result := True;
    TobMail.UpdateDB;
  end;
  {$ELSE}
  Corps := HTStringList.Create;
  Corps.Add(TobMail.GetString('CTA_BLOCNOTE'));
  try
    if AglSendMail(sTitre, TobMail.GetString('CTA_EMAIL'), '', Corps, '', True, 1) = 0 then begin
      TobMail.SetString('CTA_ENVOYE', 'X');
      TobMail.UpdateDB;
      Result := True;
    end;
  finally
    FreeAndNil(Corps);
  end;
  {$ENDIF EAGLSERVER}

end;

{ TObjTacheBap
 Méthodes publiques}
{---------------------------------------------------------------------------------------}
class function TObjTacheBap.CreateObj(Nat, User : string; DtEnvoi : TDateTime; LeParent : TOB = nil) : TObjTacheBap;
{--------------------------------------------------------------------------------------}
begin
  if Assigned(LeParent) then begin
    {On commence par regarder si un tel enregistrement ne figure pas dans la Tob mère}
    Result := TObjTacheBap(LeParent.FindFirst(['CTA_NATUREMAIL', 'CTA_DESTINATAIRE', 'CTA_DATEENVOI'],
                                              [Nat, User, DtEnvoi], True));
    if Assigned(Result) then Exit;
  end;
  {Sinon, on charge un éventuel enregistrement en base}
  Result := TObjTacheBap.Create('CPTACHEBAP', LeParent, -1);
  Result.Nature := Nat;
  Result.Destinataire := User;
  Result.DateEnvoi := DtEnvoi;
  Result.LoadDB;
end;

{---------------------------------------------------------------------------------------}
procedure TObjTacheBap.PrepareTache(TobInfo, TobMail : TOB);
{---------------------------------------------------------------------------------------}
begin
  {On regarde si la tob contient déjà un enregistrement ...}
  if GetString('CTA_CIRCUITBAP') <> '' then
    {...et dans ce cas on se contente de le mettre à jour ...}
    MajTache
  else
    {... Sinon, on crée l'enregsitrement}
    CreerTache(TobInfo, TobMail);
  {si pas de mail, on détruit la tob}
  if GetString('CTA_EMAIL') = '' then Free;
end;

{---------------------------------------------------------------------------------------}
destructor TObjTacheBap.Destroy;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{Méthodes privées}
{---------------------------------------------------------------------------------------}
procedure TObjTacheBap.MajTache;
{---------------------------------------------------------------------------------------}
begin
  {Ayant mal conçu l'index avant le "blocage" de la SocRef, on peut se retrouvé avec un
   enregistrement dont les mails ont été envoyés}
  if GetString('CTA_ENVOYE') = 'X' then begin
    {Dans ce cas, il faut réinitialiser les zones}
    SetString('CTA_ENVOYE', '-');
    SetInteger('CTA_NBVISA', 1);
  end
  else
    {Sinon, on rajoute un visa au total précédent}
    SetInteger('CTA_NBVISA', GetInteger('CTA_NBVISA') + 1);
end;

{---------------------------------------------------------------------------------------}
procedure TObjTacheBap.CreerTache(TobInfo, TobMail : TOB);
{---------------------------------------------------------------------------------------}
begin
  SetString('CTA_CODEEVISA' , TobInfo.GetString('BAP_CODEVISA'));
  SetString('CTA_CODEMAIL'  , TobMail.GetString('CMA_CODEMAIL'));
  SetString('CTA_CIRCUITBAP', TobInfo.GetString('BAP_CIRCUITBAP'));
  PutValue ('CTA_BLOCNOTE'  , TobMail.GetValue('CMA_BLOCNOTE'));
  SetString('CTA_DESTINATAIRE', Destinataire);
  if ViseurPrincipal then SetString('CTA_EMAIL', TobInfo.GetString('ADDVIS1'))
                     else SetString('CTA_EMAIL', TobInfo.GetString('ADDVIS2'));
  SetString('CTA_ENVOYE', '-');
  SetDateTime('CTA_DATEENVOI', Date);
  SetInteger('CTA_NBVISA', 1);
end;

{---------------------------------------------------------------------------------------}
procedure TObjTacheBap.SetDateEnvoi(Value : TDateTime);
{---------------------------------------------------------------------------------------}
begin
  FDateEnvoi := Value;
  SetDateTime(' CTA_DATEENVOI', Value);
end;

{---------------------------------------------------------------------------------------}
procedure TObjTacheBap.SetDestinataire(Value : string);
{---------------------------------------------------------------------------------------}
begin
  FDestinataire := Value;
  SetString('CTA_DESTINATAIRE', Value);
end;

{---------------------------------------------------------------------------------------}
procedure TObjTacheBap.SetNature(Value : string);
{---------------------------------------------------------------------------------------}
begin
  FNature := Value;
  SetString('CTA_NATUREMAIL', Value);
end;

end.
